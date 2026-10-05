# Automatizador de correos

Aplicación Flask multiusuario para cargar CSV/Excel, validar destinatarios, revisar la vista previa, adjuntar documentos permitidos y consultar el historial por usuario. El proyecto incluye autenticación, hash de contraseñas, sesiones, CSRF, SQLite y un worker separado para envíos largos.

## Seguridad de envío

La aplicación siempre prepara lotes para envío real: después de revisar la vista previa y confirmar, el lote se pone en cola para `worker.py`. La solicitud web no envía directamente; el worker procesa la cola con SMTP para evitar cortes por tiempos largos.

El worker exige STARTTLS. En producción, la contraseña se entrega mediante `systemd LoadCredential`, nunca desde el formulario ni el código. El antiguo envío desde línea de comandos de `enviar_correos.py` quedó deshabilitado para evitar saltarse estos controles.

## Ejecutar pruebas locales

Windows, desde PowerShell en la carpeta del proyecto:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Las pruebas usan SQLite temporal y SMTP simulado; no envían correos reales.

## Ejecutar en Windows para desarrollo

```powershell
.\.venv\Scripts\python.exe -m flask --app app init-db
.\.venv\Scripts\python.exe -m flask --app app create-user
.\.venv\Scripts\python.exe app.py
```

`create-user` solicita la contraseña de forma oculta en la terminal. Abre `http://127.0.0.1:5000`. El servidor integrado de Flask es solo para desarrollo.

## Ubuntu

No copies `.venv` desde Windows. Crea uno en Ubuntu e instala `requirements.txt`; Gunicorn se instala solo en Linux. Inicializa SQLite y crea usuarios mediante los comandos Flask anteriores, ejecutados dentro del entorno Linux. La ubicación de datos se controla con `APP_DATA_DIR`; por defecto el desarrollo usa `instance/`.

Los ejemplos de Gunicorn/systemd/Nginx están en `deploy/`. Son plantillas, no están instaladas ni activadas. Requieren decidir el usuario de servicio, instalar el proyecto en `/opt/correo-automatizado`, crear configuración bajo `/etc/correo-automatizado` y reemplazar `DOMINIO_REAL` solo después de confirmar el dominio. No publiques el puerto 5000.

## Red y disponibilidad externa

La VM observada usa VirtualBox en modo puente y una IP privada. El MikroTik tiene WAN privada `192.168.18.100`; falta conocer la WAN del equipo `192.168.18.1` para determinar doble NAT/CGNAT. Aún no se deben crear reglas de port forwarding ni activar UFW sin acordar el acceso SSH y completar autenticación/despliegue.

Para HTTPS público faltan dominio real, DNS, certificado Let's Encrypt y confirmar que los routers/proveedor permiten entrada. La plantilla Nginx espera Nginx→socket Unix de Gunicorn; Gunicorn no queda expuesto a Internet. No agregues un registro AAAA hasta verificar conectividad IPv6.

## Archivos de configuración de producción

- `deploy/app.env.example`: plantilla sin credenciales; el archivo real debe permanecer en `/etc` con permisos restrictivos y fuera de Git/ZIP.
- `deploy/correo-web.service.example`: Gunicorn como servicio de usuario limitado, socket local y logs en journal.
- `deploy/correo-worker.service.example`: worker independiente; el secreto SMTP solo se configura con systemd credentials después de autorización.
- `deploy/nginx-site.conf.example`: reverse proxy HTTPS con marcador `DOMINIO_REAL`; no se instala hasta conocer el dominio y obtener el certificado.

SQLite mantiene cuentas y cola transitoria. Mientras un lote está pendiente contiene sus filas y adjuntos para que el worker pueda procesarlo; al concluir borra destinatarios, nombres de adjuntos, payload y bytes de los archivos. Solo conserva conteos y estados anónimos hasta 24 horas; previews abandonadas se purgan automáticamente. Protege y respalda `/var/lib/correo-automatizado` porque la cola pendiente sí contiene datos personales.

`GET /healthz` responde solo `{"status":"ok"}` o HTTP 503 y no requiere login, para comprobaciones de disponibilidad. En producción se recomienda consultarlo localmente desde Nginx/systemd.

El script `scripts/smtp_probe.py` comprueba DNS, TCP, EHLO, STARTTLS y certificado sin ejecutar AUTH ni enviar mensajes. No lo ejecutes hasta que el administrador confirme el host y puerto que deben probarse; aún no se ha ejecutado desde Ubuntu.

`scripts/backup_sqlite.py` crea una copia consistente de SQLite usando la API de backup y permisos restrictivos. La copia de una cola pendiente puede contener destinatarios/adjuntos; el script no cifra la copia. Guarda su destino en almacenamiento cifrado aprobado y define una retención limitada para backups.

El worker y la aplicación escriben eventos operativos sin guardar contraseñas SMTP. En el ejemplo de producción Gunicorn/worker usan el journal de systemd; Nginx utiliza sus logs del sistema. Comprobaciones habituales después de instalar servicios: `systemctl status correo-web correo-worker nginx`, `journalctl -u correo-web -u correo-worker` y `curl -fsS https://DOMINIO_REAL/healthz`.
