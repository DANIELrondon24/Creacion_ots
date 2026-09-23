"""
Server mode: listens to SLACK_INPUT_CHANNEL_ID via Slack Socket Mode and
creates the OTs for every Excel file posted there (by the Slack workflow
"ots_creacion_server").

- Each message is processed once: the bot marks it with a reaction
  (eyes -> processing, white_check_mark -> ok, x -> error). Messages already
  marked by the bot are never reprocessed, not even after a restart.
- On startup, recent messages without a mark are processed (files posted
  while the container was down).
- The resulting Excel (complete or partial) is uploaded to SLACK_CHANNEL_ID.
- Jobs run one at a time in a single worker thread (Playwright sync API).

Required env (besides the Maximo ones used by main.py):
    SLACK_BOT_TOKEN, SLACK_APP_TOKEN, SLACK_INPUT_CHANNEL_ID, SLACK_CHANNEL_ID
Optional:
    DATA_DIR (default ./data), CATCHUP_MESSAGES (default 20)
"""

import os
import queue
import threading
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from slack_sdk.socket_mode import SocketModeClient
from slack_sdk.socket_mode.request import SocketModeRequest
from slack_sdk.socket_mode.response import SocketModeResponse

from main import (
    Config,
    MaximoAutomationError,
    clean_ot,
    logger,
    process_ot_creation_multi_unit,
    send_file_to_slack,
    validate_config,
)

EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
REACTION_PROCESSING = "eyes"
REACTION_OK = "white_check_mark"
REACTION_ERROR = "x"
BOT_REACTIONS = {REACTION_PROCESSING, REACTION_OK, REACTION_ERROR}


class OTServer:
    def __init__(self, config: Config):
        self.config = config
        self.web = WebClient(token=config.slack_bot_token)
        auth = self.web.auth_test()
        self.bot_user_id: str = auth["user_id"]
        self.bot_id: Optional[str] = auth.get("bot_id")
        self.jobs: "queue.Queue[Dict[str, Any]]" = queue.Queue()
        self.seen_ts: set = set()
        self.seen_lock = threading.Lock()
        self.data_dir = Path(config.data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------ utils
    @staticmethod
    def excel_files(message: Dict[str, Any]) -> List[Dict[str, Any]]:
        return [
            f
            for f in message.get("files", [])
            if str(f.get("name", "")).lower().endswith(EXCEL_EXTENSIONS)
        ]

    def is_own_message(self, message: Dict[str, Any]) -> bool:
        return message.get("user") == self.bot_user_id or (
            self.bot_id is not None and message.get("bot_id") == self.bot_id
        )

    def already_marked(self, message: Dict[str, Any]) -> bool:
        return any(
            r.get("name") in BOT_REACTIONS and self.bot_user_id in r.get("users", [])
            for r in message.get("reactions", [])
        )

    def react(self, ts: str, name: str, remove: bool = False) -> None:
        try:
            fn = self.web.reactions_remove if remove else self.web.reactions_add
            fn(channel=self.config.slack_input_channel_id, timestamp=ts, name=name)
        except SlackApiError as e:
            logger.warning(f"Could not {'remove' if remove else 'add'} reaction {name}: {e.response['error']}")

    def notify(self, text: str) -> None:
        try:
            self.web.chat_postMessage(channel=self.config.slack_channel_id, text=text)
        except SlackApiError as e:
            logger.error(f"Could not post message to Slack: {e.response['error']}")

    def download(self, file_info: Dict[str, Any], ts: str) -> Path:
        # Files shared by workflows may arrive partially; fetch full info.
        info = self.web.files_info(file=file_info["id"])["file"]
        url = info.get("url_private_download") or info["url_private"]
        name = Path(info.get("name", "ots.xlsx")).name
        dest = self.data_dir / f"{ts.replace('.', '_')}_{name}"
        req = urllib.request.Request(
            url, headers={"Authorization": f"Bearer {self.config.slack_bot_token}"}
        )
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as out:
            content = resp.read()
            # Without files:read Slack returns the HTML login page instead of the file
            if content[:2] != b"PK" and name.lower().endswith((".xlsx", ".xlsm")):
                raise RuntimeError(
                    "La descarga no es un Excel válido (¿falta el scope files:read?)"
                )
            out.write(content)
        logger.info(f"Downloaded {name} to {dest}")
        return dest

    # ---------------------------------------------------------------- enqueue
    def enqueue(self, message: Dict[str, Any], source: str) -> None:
        ts = message.get("ts")
        if not ts or self.is_own_message(message) or self.already_marked(message):
            return
        files = self.excel_files(message)
        if not files:
            logger.info(f"[{source}] Message {ts} has no Excel attachment, ignored")
            return
        with self.seen_lock:
            if ts in self.seen_ts:
                return
            self.seen_ts.add(ts)
        logger.info(f"[{source}] Queued message {ts} with file {files[0].get('name')}")
        self.jobs.put({"ts": ts, "file": files[0]})

    def catch_up(self) -> None:
        """Queue unprocessed Excel messages posted while the server was down."""
        limit = int(os.getenv("CATCHUP_MESSAGES", "20"))
        try:
            resp = self.web.conversations_history(
                channel=self.config.slack_input_channel_id, limit=limit
            )
        except SlackApiError as e:
            logger.error(f"Catch-up failed: {e.response['error']}")
            return
        for message in reversed(resp.get("messages", [])):  # oldest first
            self.enqueue(message, "catch-up")

    def on_socket_request(self, client: SocketModeClient, req: SocketModeRequest) -> None:
        client.send_socket_mode_response(SocketModeResponse(envelope_id=req.envelope_id))
        if req.type != "events_api":
            return
        event = req.payload.get("event", {})
        if event.get("type") != "message":
            return
        if event.get("channel") != self.config.slack_input_channel_id:
            return
        if event.get("subtype") in ("message_changed", "message_deleted"):
            return
        self.enqueue(event, "event")

    # ----------------------------------------------------------------- worker
    def worker(self) -> None:
        while True:
            job = self.jobs.get()
            try:
                self.process(job["ts"], job["file"])
            except Exception as e:  # never let the worker die
                logger.error(f"Unexpected worker error: {e}", exc_info=True)
            finally:
                self.jobs.task_done()

    def process(self, ts: str, file_info: Dict[str, Any]) -> None:
        self.react(ts, REACTION_PROCESSING)
        excel_path: Optional[Path] = None
        try:
            excel_path = self.download(file_info, ts)
            all_ots = process_ot_creation_multi_unit(self.config, str(excel_path))
            total = sum(len(v) for v in all_ots.values())
            detail = ", ".join(f"{u.value}: {len(v)}" for u, v in all_ots.items())
            logger.info(f"Message {ts} processed: {total} OTs created ({detail})")

            if total > 0:
                send_file_to_slack(
                    self.config,
                    str(excel_path),
                    f"Archivo de creación de Ots\n✅ {total} OTs creadas ({detail})",
                )
            else:
                self.notify("ℹ️ Archivo de creación de Ots procesado: no había OTs nuevas por crear.")
            self.react(ts, REACTION_OK)

        except Exception as e:
            logger.error(f"Processing of message {ts} failed: {e}", exc_info=True)
            created = self.count_ots(excel_path)
            text = (
                f"❌ La creación de OTs se detuvo por un error.\n"
                f"Error: {e}\n"
                f"OTs ya creadas en este archivo: {created}.\n"
                "Corrige el problema y vuelve a subir el Excel adjunto por el workflow: "
                "las filas que ya tienen OT se omiten."
            )
            if excel_path and excel_path.exists() and isinstance(e, MaximoAutomationError):
                if not send_file_to_slack(self.config, str(excel_path), text):
                    self.notify(text)
            else:
                self.notify(text)
            self.react(ts, REACTION_ERROR)
        finally:
            self.react(ts, REACTION_PROCESSING, remove=True)

    @staticmethod
    def count_ots(excel_path: Optional[Path]) -> int:
        if not excel_path or not excel_path.exists():
            return 0
        try:
            df = pd.read_excel(excel_path, dtype={"OT": str})
            if "OT" not in df.columns:
                return 0
            return int(df["OT"].apply(clean_ot).astype(bool).sum())
        except Exception:
            return 0

    # -------------------------------------------------------------------- run
    def run(self) -> None:
        threading.Thread(target=self.worker, name="ot-worker", daemon=True).start()
        socket_client = SocketModeClient(
            app_token=self.config.slack_app_token, web_client=self.web
        )
        socket_client.socket_mode_request_listeners.append(self.on_socket_request)
        socket_client.connect()
        logger.info(
            f"Listening Slack channel {self.config.slack_input_channel_id} "
            f"(results -> {self.config.slack_channel_id})"
        )
        self.catch_up()
        threading.Event().wait()


def main() -> None:
    config = Config.from_env()
    config.headless = True  # no display inside the container
    validate_config(config, server_mode=True)
    logger.info(f"OT server starting at {datetime.now():%Y-%m-%d %H:%M:%S}")
    OTServer(config).run()


if __name__ == "__main__":
    main()
