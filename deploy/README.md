# Plantillas de despliegue Ubuntu

Estos archivos son ejemplos para una futura instalación, no una configuración activa. No copiar a `/etc`, no habilitar UFW ni crear port forwarding hasta confirmar dominio, WAN/router, SMTP y controles de seguridad.

## Componentes

- `correo-web.service.example`: Gunicorn con dos workers detrás de un socket Unix local. Los logs se envían al journal de systemd.
- `correo-worker.service.example`: procesa la cola fuera de las solicitudes web. SMTP credentials solo se suministran al worker con systemd credentials.
- `app.env.example`: configuración sin contraseña SMTP. Genera `FLASK_SECRET_KEY` localmente en Ubuntu y no la pegues en el chat.
- `nginx-bootstrap.conf.example`: sitio HTTP mínimo para validar DNS y solicitar el primer certificado.
- `nginx-site.conf.example`: configuración final HTTPS. Los marcadores `DOMINIO_REAL` deben reemplazarse por el FQDN real después de confirmar propiedad y DNS.

## Requisitos antes de instalar

1. Confirmar el dominio y su autorización de uso.
2. Comparar la WAN del equipo `192.168.18.1` con la IP pública de salida; actualmente el MikroTik presenta WAN privada `192.168.18.100`. Resolver doble NAT/CGNAT antes de redirigir puertos.
3. Confirmar port-forwarding en cada router, si es viable, hacia la IP estable de Ubuntu. Solo Nginx recibirá HTTP/HTTPS; Gunicorn usa socket Unix y Flask 5000 no se publica.
4. Instalar el proyecto en `/opt/correo-automatizado`, crear usuario de servicio `correoapp` y entorno virtual Linux.
5. Crear `/etc/correo-automatizado/app.env` con permisos restrictivos. Añadir SMTP values solo cuando el administrador confirme host/port/TLS. El archivo real nunca se versiona ni se envía en ZIP.
6. Preparar `/var/lib/correo-automatizado` como directorio de datos de servicio. La app crea SQLite con permisos restringidos.

## Secretos SMTP y envío real

El worker lee la contraseña desde `CREDENTIALS_DIRECTORY/smtp_password`. El archivo fuente del secreto debe pertenecer a root y tener permisos estrictos; activar `LoadCredential` solo después de validar SMTP sin AUTH y de confirmar la autorización para envíos reales.

`scripts/smtp_probe.py` es la prueba previa autorizada: DNS, conexión TCP, EHLO y STARTTLS con validación de certificado, sin AUTH ni correo. Ejecutarla solo después de que el administrador confirme servidor/puerto. `scripts/backup_sqlite.py` usa la API de backup de SQLite; guardar copias en un volumen cifrado y acordar su retención.

## Dominio, certificado y firewall

El bootstrap HTTP se usaría únicamente para el desafío ACME si el dominio apunta a una IP pública de entrada alcanzable. Luego se instala el certificado con Certbot y se cambia al sitio HTTPS final. La renovación automática debe verificarse con `certbot renew --dry-run`.

UFW está actualmente inactivo y no se debe activar a ciegas: primero hay que confirmar la ruta SSH de administración. La configuración final debe permitir solo SSH administrado y HTTP/HTTPS requeridos; nunca abrir TCP 5000 ni el socket Gunicorn.

La dirección WAN privada del MikroTik (`192.168.18.100`) indica otra capa de red. Falta consultar la WAN de `192.168.18.1`; solo después se podrá decidir si hay doble port-forwarding, CGNAT o necesidad de VPS. No copiar reglas del ejemplo al router sin esa verificación.

## Crear usuarios

Las cuentas se crean por administrador desde la VM con el comando Flask `create-user`, que solicita la contraseña de forma oculta en la terminal y guarda solo un hash. No existe registro público, recuperación de contraseña ni roles; todos los usuarios activos pueden preparar lotes y solicitar envíos.

El endpoint `/healthz` devuelve solo estado mínimo. `journalctl -u correo-web -u correo-worker` consulta logs de aplicación/worker; Nginx usa los logs del paquete Ubuntu. `scripts/backup_sqlite.py` debe apuntar a almacenamiento cifrado. La duración de conservación del historial y respaldos queda pendiente de política organizacional antes de producción.
