# Creacion_ots

Automatiza la creación de Órdenes de Trabajo (OT) en Maximo a partir de un Excel.

## Modos de ejecución

- **Servidor (`server.py`)**: escucha por Socket Mode el canal `SLACK_INPUT_CHANNEL_ID`.
  Cuando el workflow `ots_creacion_server` publica el Excel ("ots a crear"), lo
  descarga, crea las OTs y sube el Excel con los números de OT a `SLACK_CHANNEL_ID`.
- **Local (`main.py` / `run_ots.bat`)**: procesa el archivo de `EXCEL_FILE_PATH`.

Ante el primer error la ejecución se detiene y se sube a `SLACK_CHANNEL_ID` el
Excel parcial. Se corrige y se vuelve a subir por el workflow: las filas que ya
tienen OT se omiten.

El bot marca cada mensaje del canal de entrada con 👀 (procesando), ✅ (ok) o
❌ (error). Un mensaje marcado nunca se reprocesa; para forzar un reintento,
vuelve a subir el archivo por el workflow.

## Configuración de la app de Slack

1. **Socket Mode**: activarlo y generar un *App-Level Token* (`xapp-...`) con
   `connections:write` → `SLACK_APP_TOKEN`.
2. **Event Subscriptions**: bot events `message.channels` (y `message.groups`
   si el canal es privado).
3. **Bot scopes**: `channels:history` (`groups:history` si es privado),
   `files:read`, `files:write`, `chat:write`, `reactions:write`.
4. Reinstalar la app e invitar el bot a ambos canales.

## Despliegue en Rocky Linux (Podman)

```bash
sudo dnf install -y podman
git clone <repo> ~/crear-ot-src && cd ~/crear-ot-src
podman build -t crear-ot:latest .          # o: docker build -t crear-ot:latest .

mkdir -p ~/crear-ot/{data,logs}
cp .env.example ~/crear-ot/.env             # completar valores
```

Prueba manual:

```bash
podman run --rm -it --env-file ~/crear-ot/.env --shm-size=1g \
  --userns=keep-id:uid=1001,gid=1001 \
  -v ~/crear-ot/data:/app/data:Z -v ~/crear-ot/logs:/app/logs:Z \
  crear-ot:latest
```

Como servicio (arranca con el servidor), usando Quadlet:

```bash
mkdir -p ~/.config/containers/systemd
cp deploy/crear-ot.container ~/.config/containers/systemd/
systemctl --user daemon-reload
systemctl --user start crear-ot
sudo loginctl enable-linger $USER           # que siga corriendo sin sesión abierta
journalctl --user -u crear-ot -f            # logs
```

Para producción basta con cambiar `SLACK_INPUT_CHANNEL_ID` (y `SLACK_CHANNEL_ID`)
en `~/crear-ot/.env` y reiniciar: `systemctl --user restart crear-ot`.

El servidor necesita salida a internet hacia Slack (WebSocket saliente, sin
puertos abiertos) y acceso de red a `MAXIMO_URL`.
