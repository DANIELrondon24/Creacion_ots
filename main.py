"""
Unified OT Creator Script - Production Ready

This script integrates the logic from CREACION_OT_Z63.py and CREACION_OT_Z67.py,
migrating the automation to Playwright for improved reliability and performance.
It supports processing multiple functional units (Z63 and Z67) from the same Excel file.

Features:
- Playwright-based automation (headless mode)
- Support for multiple units (Z63 and Z67) in the same Excel file
- Automatic navigation back to list view after completing each unit
- Robust error handling with retries
- STOPS execution on error and prompts user to fix data before continuing
- Comprehensive logging
- Configuration management via environment variables
- Input validation and edge case handling
- Modular code structure with type hints
- Production-ready with timeouts and scalability considerations
Requirements:
- Python 3.8+
- pandas
- python-dotenv
- playwright
- tenacity
- slack-sdk

Install dependencies:
    pip install pandas python-dotenv playwright tenacity slack-sdk
    playwright install

Usage:
    Set environment variables or create a .env file with:
    MAXIMO_USERNAME=your_username
    MAXIMO_PASSWORD=your_password
    MAXIMO_URL=http://maximo.example.com/...
    EXCEL_FILE_PATH=path/to/your/excel/file.xlsx
    WORKSHOP_TIME=3
    SUPERVISOR_ID=your_supervisor_id
    SLACK_BOT_TOKEN=xoxb-your-slack-bot-token
    SLACK_CHANNEL_ID=C03RPFUMSR1
    HEADLESS=true  # optional, default true
    TIMEOUT=30000  # optional, default 30000ms

    python main.py            # local: processes EXCEL_FILE_PATH

Server mode (container): see server.py, which listens to the Slack channel
SLACK_INPUT_CHANNEL_ID and processes each Excel uploaded by the workflow.
"""

import os
import logging
import time
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
from dataclasses import dataclass
from enum import Enum
from collections import defaultdict

# Base directory of the project
BASE_DIR = Path(__file__).resolve().parent


# Constants
DEFAULT_TIMEOUT = 30000  # 30 seconds
RETRY_ATTEMPTS = 3
RETRY_WAIT_MIN = 1
RETRY_WAIT_MAX = 10


# Type definitions
class FunctionalUnit(Enum):
    Z63 = "Z63"
    Z67 = "Z67"


@dataclass
class Config:
    username: str
    password: str
    excel_file_path: str
    workshop_time: str
    supervisor_id: str
    maximo_url: str
    slack_bot_token: str
    slack_channel_id: str
    slack_app_token: str = ""
    slack_input_channel_id: str = ""
    data_dir: str = str(BASE_DIR / "data")
    headless: bool = True
    timeout: int = DEFAULT_TIMEOUT

    @classmethod
    def from_env(cls) -> "Config":
        """Create Config from environment variables."""
        excel_path_env = os.getenv("EXCEL_FILE_PATH", "")
        if excel_path_env:
            excel_path = Path(excel_path_env)
            if not excel_path.is_absolute():
                excel_path_env = str(BASE_DIR / excel_path_env)

        return cls(
            username=os.getenv("MAXIMO_USERNAME", ""),
            password=os.getenv("MAXIMO_PASSWORD", ""),
            excel_file_path=excel_path_env,
            workshop_time=os.getenv("WORKSHOP_TIME", "3"),
            supervisor_id=os.getenv("SUPERVISOR_ID", ""),
            maximo_url=os.getenv("MAXIMO_URL", ""),
            slack_bot_token=os.getenv("SLACK_BOT_TOKEN", ""),
            slack_channel_id=os.getenv("SLACK_CHANNEL_ID", ""),
            slack_app_token=os.getenv("SLACK_APP_TOKEN", ""),
            slack_input_channel_id=os.getenv("SLACK_INPUT_CHANNEL_ID", ""),
            data_dir=os.getenv("DATA_DIR", str(BASE_DIR / "data")),
            headless=os.getenv("HEADLESS", "true").lower() == "true",
            timeout=int(os.getenv("TIMEOUT", str(DEFAULT_TIMEOUT))),
        )


# Setup logging
def setup_logging() -> logging.Logger:
    """Setup logging configuration."""
    log_dir = Path(os.getenv("LOG_DIR", str(BASE_DIR)))
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "ot_creator.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(str(log_file), encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    return logging.getLogger(__name__)


logger = setup_logging()

# Import external modules with logging
try:
    import pandas as pd
    # pyrefly: ignore [missing-import]
    from dotenv import load_dotenv
    from playwright.sync_api import (
        sync_playwright,
        Page,
        Browser,
        BrowserContext,
        TimeoutError as PlaywrightTimeoutError,
    )
    from tenacity import (
        retry,
        stop_after_attempt,
        wait_exponential,
        retry_if_exception_type,
    )
    from slack_sdk import WebClient
    from slack_sdk.errors import SlackApiError
    from datetime import datetime

    logger.info("All required modules imported successfully")
except ImportError as e:
    logger.error(f"Failed to import required modules: {e}")
    raise

# Load environment variables
env_path = BASE_DIR / ".env"
load_dotenv(dotenv_path=env_path)

# Word replacement dictionaries
TYPE_REPLACEMENTS = {"HALL": "MCO"}

DESCRIPTION_REPLACEMENTS = {
    "GARANTIA": "GARA",
    "MTTO PREVENTIVO": "SISTEMATICA",
    "DAÑO OPERACIONAL": "DANO OPERACIONAL",
    "FALLA TECNICA": "FALLA TECNICA",
    "HALLAZGO ENTE GESTOR": "HALLAZGO ENTE GESTOR",
}


class MaximoAutomationError(Exception):
    """Custom exception for Maximo automation errors."""

    pass


class RowProcessingError(MaximoAutomationError):
    """A specific Excel row could not be processed; execution must stop."""

    pass


class MaximoAutomator:
    """Handles Playwright-based automation for Maximo OT creation."""

    def __init__(self, config: Config):
        self.config = config
        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def __enter__(self):
        """Context manager entry."""
        self.start_browser()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close_browser()

    def start_browser(self):
        """Start Playwright browser."""
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(
                headless=self.config.headless,
                args=[
                    "--ignore-certificate-errors",
                    "--allow-insecure-localhost",
                    "--disable-dev-shm-usage",
                ],
            )
            self.context = self.browser.new_context(
                viewport={"width": 2100, "height": 1100}
            )
            self.page = self.context.new_page()
            logger.info("Browser started successfully")
        except Exception as e:
            logger.error(f"Failed to start browser: {e}")
            raise MaximoAutomationError(f"Browser startup failed: {e}")

    def close_browser(self):
        """Close browser and cleanup."""
        try:
            if self.page:
                self.page.close()
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
            logger.info("Browser closed")
        except Exception as e:
            logger.warning(f"Error during browser cleanup: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError)),
    )
    def login(self):
        """Perform login to Maximo."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            logger.info("Attempting to login to Maximo...")
            self.page.goto(self.config.maximo_url, timeout=self.config.timeout)
            self.page.fill("#username", self.config.username)
            self.page.fill("#password", self.config.password)
            self.page.click("#loginbutton")
            time.sleep(3)  # Wait for login to complete
            logger.info("Login successful")
        except Exception as e:
            logger.error(f"Login failed: {e}")
            raise MaximoAutomationError(f"Login failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError)),
    )
    def select_site(self, unit: FunctionalUnit):
        """Select the functional unit site."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        selector = {
            FunctionalUnit.Z63: "#toolactions_GRN_CHANGEDEFSITE3-tbb",
            FunctionalUnit.Z67: "#toolactions_GRN_CHANGEDEFSITE5-tbb",
        }.get(unit)

        if not selector:
            raise MaximoAutomationError(f"Unknown functional unit: {unit}")

        try:
            logger.info(f"Selecting site for {unit.value}...")
            self.page.wait_for_selector(selector, timeout=self.config.timeout)
            self.page.click(selector)
            time.sleep(3)  # Wait for site change to complete
            logger.info(f"Site selected: {unit.value}")
        except Exception as e:
            logger.error(f"Site selection failed for {unit.value}: {e}")
            raise MaximoAutomationError(f"Site selection failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError)),
    )
    def return_to_list_view(self):
        """Click the 'Vista de lista' button to return to the list view."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            logger.info("Returning to list view...")
            try:
                self.page.get_by_role("button", name="Vista de lista").click(timeout=5000)
            except Exception:
                try:
                    self.page.click("#toolactions_LISTVIEW-tbb", timeout=5000)
                except Exception:
                    self.page.click("text=Vista de lista", timeout=5000)
            time.sleep(3)  # Wait for the view to load
            logger.info("Successfully returned to list view")
        except Exception as e:
            logger.error(f"Failed to return to list view: {e}")
            raise MaximoAutomationError(f"Return to list view failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError)),
    )
    def create_new_ot(self):
        """Create a new OT."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            self.page.wait_for_selector(
                "#toolactions_INSERT-tbb", timeout=self.config.timeout * 1.5
            )
            self.page.click("#toolactions_INSERT-tbb")
            time.sleep(2)
            logger.info("New OT creation initiated")
        except Exception as e:
            logger.error(f"Failed to create new OT: {e}")
            raise MaximoAutomationError(f"New OT creation failed: {e}")

    def fill_ot_form(self, row: pd.Series, row_index: int) -> str:
        """Fill the OT form with data from Excel row."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        placa = str(row.get("Placa", "")).strip() if pd.notna(row.get("Placa")) else ""
        actividad = str(row.get("Actividad", "")).strip().title() if pd.notna(row.get("Actividad")) else ""

        try:
            logger.info(
                f"Filling OT form for row {row_index + 1} - Plate: {placa}"
            )

            # Activity
            self.page.wait_for_selector("#mad3161b5-tb2", timeout=self.config.timeout)
            self.page.fill("#mad3161b5-tb2", actividad)
            time.sleep(0.5)
            self.page.press("#mad3161b5-tb2", "Tab")

            # License plate
            self.page.wait_for_selector("#m3b6a207f-tb", timeout=self.config.timeout)
            self.page.fill("#m3b6a207f-tb", placa)
            time.sleep(0.5)
            self.page.press("#m3b6a207f-tb", "Tab")
            time.sleep(1)

            # Type
            self.page.wait_for_selector("#me2096203-tb", timeout=self.config.timeout)
            tipo = str(row.get("Tipo", "")).strip() if pd.notna(row.get("Tipo")) else ""
            for old, new in TYPE_REPLACEMENTS.items():
                tipo = tipo.replace(old, new)
            self.page.fill("#me2096203-tb", tipo)
            time.sleep(0.5)
            self.page.press("#me2096203-tb", "Tab")

            # Description Type
            self.page.wait_for_selector("#m78c05445-tb", timeout=self.config.timeout)
            des_tipo = str(row.get("Descripcion_Tipo", "")).strip() if pd.notna(row.get("Descripcion_Tipo")) else ""
            for old, new in DESCRIPTION_REPLACEMENTS.items():
                des_tipo = des_tipo.replace(old, new)
            self.page.fill("#m78c05445-tb", des_tipo)
            time.sleep(0.5)
            self.page.press("#m78c05445-tb", "Tab")

            # Priority
            self.page.wait_for_selector("#m950e5295-tb", timeout=self.config.timeout)
            self.page.fill("#m950e5295-tb", str(self.config.workshop_time))
            time.sleep(0.5)
            self.page.press("#m950e5295-tb", "Tab")

            # Scheduled Start
            self.page.wait_for_selector("#m8b12679c-tb", timeout=self.config.timeout)
            fecha = row.get("Fecha")
            try:
                dt = pd.to_datetime(fecha)
                if pd.notna(dt):
                    fecha_str = dt.strftime("%d/%m/%Y")
                else:
                    fecha_str = str(fecha).strip()
            except Exception:
                fecha_str = str(fecha).strip()
            self.page.fill("#m8b12679c-tb", fecha_str)
            time.sleep(0.5)
            self.page.press("#m8b12679c-tb", "Tab")
            time.sleep(1)

            # Supervisor
            self.page.wait_for_selector("#mb2eb834-tb", timeout=self.config.timeout)
            self.page.fill("#mb2eb834-tb", str(self.config.supervisor_id))
            time.sleep(0.5)
            self.page.press("#mb2eb834-tb", "Tab")
            time.sleep(1)

            # Extract OT number with multiple fallback selectors
            time.sleep(0.5)
            ot_value = None
            ot_selectors = [
                "xpath=/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr[3]/td/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr[1]/td[2]/input[1]",
                "input[id*='wonum']",
                "input[id$='-tb1']",
                "input[aria-label*='Orden de trabajo']",
                "input[aria-label*='Work Order']",
                "#m3b6a207f-tb1",
            ]
            for sel in ot_selectors:
                try:
                    ot_el = self.page.wait_for_selector(sel, timeout=3000)
                    if ot_el:
                        val = ot_el.get_attribute("value")
                        if val and val.strip():
                            ot_value = val.strip()
                            break
                except Exception:
                    continue

            if not ot_value:
                val_js = self.page.evaluate("""() => {
                    const el = document.querySelector('input[id*="wonum"], input[name*="wonum"], input[id$="-tb1"]');
                    return el ? el.value : '';
                }""")
                if val_js and val_js.strip():
                    ot_value = val_js.strip()

            logger.info(f"OT created: {ot_value} for plate: {placa}")

            # Save OT
            try:
                save_button = self.page.wait_for_selector(
                    "#toolactions_SAVE-tbb", timeout=self.config.timeout
                )
                save_button.click()
                time.sleep(2)
            except Exception:
                self.page.evaluate(
                    "document.querySelector('#toolactions_SAVE-tbb') && document.querySelector('#toolactions_SAVE-tbb').click()"
                )
                time.sleep(2)

            logger.info(f"OT {ot_value} saved successfully")
            return ot_value or ""

        except Exception as e:
            logger.error(f"Failed to fill OT form for plate {placa}: {e}")
            raise MaximoAutomationError(f"Form filling failed: {e}")


def clean_ot(value: Any) -> str:
    """Normalize an OT number: empty for NaN/None, without a float '.0' suffix."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in ("nan", "none"):
        return ""
    return text[:-2] if text.endswith(".0") else text


def safe_save_excel(df: pd.DataFrame, file_path: str) -> bool:
    """Safely save DataFrame to Excel, formatting OT numbers to avoid float '.0' suffixes."""
    try:
        if "OT" in df.columns:
            df["OT"] = df["OT"].apply(clean_ot)
        df.to_excel(file_path, index=False)
        return True
    except PermissionError:
        logger.warning(f"Permission denied saving {file_path}. Excel file may be open in another application.")
        print(f"\n⚠️ NO SE PUDO GUARDAR EN EXCEL: El archivo '{file_path}' está abierto en Microsoft Excel.")
        print("   Por favor cierra el archivo para permitir actualizar las OTs.")
        return False
    except Exception as e:
        logger.warning(f"Could not save to Excel: {e}")
        return False


def send_file_to_slack(config: Config, file_path: str, message: str) -> bool:
    """Upload the Excel file to SLACK_CHANNEL_ID (non-interactive).

    Returns:
        bool: True if file was sent successfully, False otherwise
    """
    try:
        client = WebClient(token=config.slack_bot_token)
        fecha_hora = datetime.now().strftime("%d %b %Y %H:%M:%S")

        logger.info(f"Uploading file to Slack channel {config.slack_channel_id}...")
        response = client.files_upload_v2(
            channel=config.slack_channel_id,
            file=file_path,
            initial_comment=f"{message}\n{fecha_hora}",
            title=f"OTs - {fecha_hora}",
        )

        if response["ok"]:
            logger.info("File uploaded to Slack successfully")
            return True
        logger.error(f"Slack upload failed: {response}")
        return False

    except SlackApiError as e:
        logger.error(f"Slack API error: {e.response['error']}")
        return False
    except Exception as e:
        logger.error(f"Failed to send file to Slack: {e}")
        return False


def validate_config(config: Config, server_mode: bool = False) -> None:
    """Validate configuration parameters."""
    errors = []

    if not config.username or not config.password:
        errors.append("Username and password are required")

    if server_mode:
        if not config.slack_app_token:
            errors.append("Slack app token (SLACK_APP_TOKEN) is required")
        if not config.slack_input_channel_id:
            errors.append("Slack input channel (SLACK_INPUT_CHANNEL_ID) is required")
    elif not config.excel_file_path:
        errors.append("Excel file path is required")
    elif not Path(config.excel_file_path).exists():
        errors.append(f"Excel file not found: {config.excel_file_path}")

    if not config.workshop_time:
        errors.append("Workshop time is required")

    if not config.supervisor_id:
        errors.append("Supervisor ID is required")
    
    if not config.maximo_url:
        errors.append("Maximo URL is required")
    
    if not config.slack_bot_token:
        errors.append("Slack bot token is required")
    
    if not config.slack_channel_id:
        errors.append("Slack channel ID is required")

    if errors:
        raise ValueError(
            "Configuration validation failed:\n"
            + "\n".join(f"  - {err}" for err in errors)
        )


def group_by_functional_unit(df: pd.DataFrame) -> Dict[FunctionalUnit, pd.DataFrame]:
    """Group DataFrame rows by functional unit based on mobile identifier prefix."""
    if df.empty:
        raise ValueError("Excel file is empty")

    # Validate required columns
    required_columns = [
        "Movil",
        "Actividad",
        "Placa",
        "Tipo",
        "Descripcion_Tipo",
        "Fecha",
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            f"Missing required columns in Excel file: {', '.join(missing_columns)}"
        )

    groups = defaultdict(list)

    for index, row in df.iterrows():
        mobile = str(row["Movil"]).upper().strip()

        if mobile.startswith("Z63-"):
            groups[FunctionalUnit.Z63].append(index)
        elif mobile.startswith("Z67-"):
            groups[FunctionalUnit.Z67].append(index)
        else:
            logger.warning(
                f"Row {index + 1}: Invalid mobile identifier '{mobile}'. Must start with Z63- or Z67-. Skipping."
            )

    if not groups:
        raise ValueError(
            "No valid mobile identifiers found. All rows must have mobile identifiers starting with Z63- or Z67-"
        )

    # Create DataFrames for each group
    result = {}
    for unit, indices in groups.items():
        result[unit] = df.loc[indices].copy()
        logger.info(f"Found {len(indices)} rows for unit {unit.value}")

    return result


def process_ot_creation_multi_unit(
    config: Config,
    excel_path: str,
) -> Dict[FunctionalUnit, List[Tuple[int, str]]]:
    """Process OT creation for multiple functional units from the same Excel file.

    STOPS execution on the first failing row (raises MaximoAutomationError).
    OTs created so far are already saved in the Excel file, so it can be
    re-uploaded after fixing the error: rows with an OT are skipped.
    """
    # Load and group data safely as strings for OT column
    df = pd.read_excel(excel_path, dtype={"OT": str})
    logger.info(f"Loaded {len(df)} rows from Excel file: {excel_path}")

    # Initialize OT column if it doesn't exist
    if "OT" not in df.columns:
        df["OT"] = ""

    unit_groups = group_by_functional_unit(df)
    all_ots = {}

    with MaximoAutomator(config) as automator:
        automator.login()

        # Process each functional unit
        for unit_index, (unit, unit_df) in enumerate(unit_groups.items()):
            logger.info(f"\n{'='*60}")
            logger.info(
                f"Processing functional unit: {unit.value} ({len(unit_df)} OTs)"
            )
            logger.info(f"{'='*60}")

            ots_created = []

            try:
                # Select the site for this unit
                automator.select_site(unit)

                # Process each row for this unit
                for idx, (original_index, row) in enumerate(unit_df.iterrows()):
                    # Check if OT already exists for this row
                    clean_existing = clean_ot(df.at[original_index, "OT"])
                    if clean_existing:
                        logger.info(
                            f"Skipping row {original_index + 1} - OT already exists: {clean_existing}"
                        )
                        print(
                            f"⏭️  Fila {original_index + 1} omitida - OT ya existe: {clean_existing} (Placa: {row['Placa']})"
                        )
                        continue

                    try:
                        automator.create_new_ot()
                        ot = automator.fill_ot_form(row, original_index)

                        if ot:
                            ots_created.append((original_index, ot))

                            # Save OT immediately to Excel
                            df.at[original_index, "OT"] = ot
                            if safe_save_excel(df, excel_path):
                                logger.info(f"OT {ot} saved to Excel immediately")

                            logger.info(
                                f"Progress: {idx + 1}/{len(unit_df)} OTs created for {unit.value}"
                            )
                            time.sleep(3)  # Wait between OTs
                        else:
                            raise RowProcessingError(
                                f"No se pudo obtener el número de OT para la fila {original_index + 1} "
                                f"(Móvil: {row['Movil']}, Placa: {row['Placa']})"
                            )

                    except RowProcessingError:
                        safe_save_excel(df, excel_path)
                        raise
                    except Exception as e:
                        logger.error(f"Failed to process row {original_index + 1}: {e}")
                        safe_save_excel(df, excel_path)
                        raise RowProcessingError(
                            f"Error en la fila {original_index + 1} "
                            f"(Móvil: {row['Movil']}, Placa: {row['Placa']}): {e}"
                        ) from e

                # After completing all OTs for this unit, return to list view
                # (except if this is the last unit being processed)
                is_last_unit = unit_index == len(unit_groups) - 1

                if not is_last_unit and ots_created:
                    try:
                        logger.info(
                            f"Completed all OTs for {unit.value}. Returning to list view..."
                        )
                        automator.return_to_list_view()
                    except Exception as e:
                        logger.error(
                            f"Failed to return to list view after {unit.value}: {e}"
                        )
                        raise MaximoAutomationError(
                            f"Failed to return to list view: {e}"
                        )

                # Store results
                all_ots[unit] = ots_created
                logger.info(
                    f"✓ Completed {len(ots_created)}/{len(unit_df)} OTs for {unit.value}"
                )

            except MaximoAutomationError:
                # Store partial results and re-raise to stop execution
                all_ots[unit] = ots_created
                raise
            except Exception as e:
                logger.error(f"Unexpected error processing unit {unit.value}: {e}")
                all_ots[unit] = ots_created
                raise MaximoAutomationError(f"Unexpected error: {e}")

    # Final save to ensure all OTs are in the Excel file
    logger.info("\nPerforming final save of Excel file...")
    if safe_save_excel(df, excel_path):
        logger.info(f"Excel file saved successfully: {excel_path}")
    else:
        logger.warning("Final save warning, but OTs were saved incrementally if file was unlocked")

    return all_ots


def main():
    """Main entry point."""
    try:
        logger.info("=" * 80)
        logger.info("UNIFIED OT CREATOR - PRODUCTION VERSION")
        logger.info("=" * 80)

        # Load and validate configuration
        config = Config.from_env()
        validate_config(config)

        logger.info(f"Configuration loaded:")
        logger.info(f"  - Excel file: {config.excel_file_path}")
        logger.info(f"  - Workshop time: {config.workshop_time}")
        logger.info(f"  - Supervisor ID: {config.supervisor_id}")
        logger.info(f"  - Headless mode: {config.headless}")
        logger.info(f"  - Timeout: {config.timeout}ms")

        # Process OT creation
        all_ots = process_ot_creation_multi_unit(config, config.excel_file_path)

        # Summary
        logger.info("\n" + "=" * 80)
        logger.info("EXECUTION SUMMARY")
        logger.info("=" * 80)

        total_ots = 0
        for unit, ots_list in all_ots.items():
            count = len(ots_list)
            total_ots += count
            logger.info(f"  - {unit.value}: {count} OTs created")

        logger.info(f"\nTotal OTs created: {total_ots}")
        logger.info("=" * 80)

        print("\n✓ Las Órdenes de Trabajo han sido creadas exitosamente.")
        print(f"✓ Total de OTs creadas: {total_ots}")
        print("✓ Por favor verifica que estén correctas en el sistema Maximo.")
        print(f"✓ El archivo Excel ha sido actualizado: {config.excel_file_path}")
        
        # Send file to Slack (local mode asks for confirmation)
        if total_ots > 0:
            print(f"\n¿Enviar el archivo al canal de Slack {config.slack_channel_id}?")
            try:
                respuesta = input("Responde 'si' para enviar: ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                respuesta = ""
            if respuesta == "si":
                if send_file_to_slack(config, config.excel_file_path, "Archivo de creación de Ots"):
                    print("✅ Archivo enviado exitosamente a Slack")
                else:
                    print("❌ Error al enviar archivo a Slack (ver log)")
            else:
                print("❌ Envío cancelado por el usuario")

    except MaximoAutomationError as e:
        logger.error(f"Automation error: {e}")
        print(f"\n⚠️ El proceso se detuvo debido a un error: {e}")
        print(f"⚠️ Revisa el archivo 'ot_creator.log' para más detalles.")
    except Exception as e:
        logger.error(f"Script execution failed: {e}", exc_info=True)
        print(f"\n✗ Error: {e}")
        print("✗ Revisa el archivo 'ot_creator.log' para más detalles.")
        raise


if __name__ == "__main__":
    main()