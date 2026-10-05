import csv
import html
import getpass
import io
import os
import re
import smtplib
import time
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

# --- CONFIGURACIÓN ---
SMTP_SERVER = os.environ.get("SMTP_SERVER", "nombre_smtp")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "25"))
CORREO_ORIGEN = os.environ.get("CORREO_ORIGEN", "correo_de_origen")
SMTP_USUARIO = os.environ.get("SMTP_USUARIO", "correo_de_origen")
NOMBRE_REMITENTE = os.environ.get("NOMBRE_REMITENTE", "Plataforma de Atención Clientes")

# El correo siempre se envia en HTML.
ARCHIVO_CSV = Path("Envio-de-correo.csv")
PAUSA_ENTRE_ENVIOS = 30


def normalizar_empresas(empresa: str) -> str:
    """Separa varias empresas usando guion."""
    if " - " in empresa:
        return empresa
    empresa = re.sub(r"\s*;\s*", " - ", empresa)
    return re.sub(r"\s+y\s+", " - ", empresa, flags=re.IGNORECASE)


def leer_destinatarios() -> list[dict[str, str]]:
    """Lee y valida las columnas obligatorias del archivo CSV."""
    contenido = ARCHIVO_CSV.read_bytes()
    try:
        texto = contenido.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = contenido.decode("cp1252")

    lector = csv.DictReader(io.StringIO(texto, newline=""))
    columnas_requeridas = {"EMPRESA", "DESTINATARIO"}
    columnas = set(lector.fieldnames or [])
    faltantes = columnas_requeridas - columnas
    if faltantes:
        raise ValueError(
            f"El CSV no contiene las columnas obligatorias: {', '.join(sorted(faltantes))}"
        )

    destinatarios = []
    for numero_fila, fila in enumerate(lector, start=2):
        empresa = (fila.get("EMPRESA") or "").strip()
        empresa = normalizar_empresas(empresa)
        archivo = (fila.get("ARCHIVO") or "").strip()
        correo = (fila.get("DESTINATARIO") or "").strip()
        sede = (fila.get("SEDE") or "").strip()
        nc = (fila.get("NC") or "").strip()
        fecha = (fila.get("FECHA") or "").strip()
        motivo = (fila.get("MOTIVO") or "").strip()
        if not empresa or not correo:
            print(f"Aviso: se omite la fila {numero_fila} por datos incompletos.")
            continue
        destinatarios.append(
            {
                "EMPRESA": empresa,
                "ARCHIVO": archivo,
                "DESTINATARIO": correo,
                "SEDE": sede,
                "NC": nc,
                "FECHA": fecha,
                "MOTIVO": motivo or "cambio de una mufla de la infraestructura externa",
                "CC": (fila.get("CC") or "").strip(),
                "FECHA_INICIO": (fila.get("FECHA_INICIO") or "").strip(),
                "HORA_INICIO": (fila.get("HORA_INICIO") or "").strip(),
                "FECHA_TERMINO": (fila.get("FECHA_TERMINO") or "").strip(),
                "HORA_CULMINACION": (fila.get("HORA_CULMINACION") or "").strip(),
                "DURACION": (fila.get("DURACION") or "").strip(),
            }
        )
    return destinatarios


def _direcciones_unicas(valor: str) -> list[str]:
    """Divide direcciones separadas por punto y coma o coma y elimina duplicados."""
    direcciones = []
    vistas = set()
    for bloque in valor.replace(",", ";").split(";"):
        direccion = bloque.strip()
        clave = direccion.casefold()
        if direccion and clave not in vistas:
            direcciones.append(direccion)
            vistas.add(clave)
    return direcciones


def _valores_multilinea(valor: str) -> list[str]:
    """Divide valores separados por punto y coma y elimina elementos vacíos."""
    return [item.strip() for item in valor.split(";") if item.strip()]


def obtener_destinatarios_smtp(mensaje: EmailMessage) -> list[str]:
    """Obtiene los destinatarios reales de To y Cc para el sobre SMTP."""
    return _direcciones_unicas(
        ";".join(filter(None, [mensaje.get("To", ""), mensaje.get("Cc", "")]))
    )


def crear_mensaje(
    correo_destino: str,
    empresa: str,
    archivo: str,
    sede: str,
    nc: str,
    fecha: str,
    motivo: str,
    fecha_inicio: str,
    hora_inicio: str,
    fecha_termino: str,
    hora_culminacion: str,
    duracion: str,
    cc_adicional: str,
) -> EmailMessage:
    """Crea un correo en formato HTML y texto plano."""
    mensaje = EmailMessage()
    mensaje["From"] = f"{NOMBRE_REMITENTE} <{CORREO_ORIGEN}>"
    mensaje["Date"] = formatdate(localtime=True)
    destinatarios = _direcciones_unicas(correo_destino)
    cc = _direcciones_unicas(cc_adicional)
    cc = [direccion for direccion in cc if direccion.casefold() not in {item.casefold() for item in destinatarios}]
    mensaje["To"] = ", ".join(destinatarios)
    if cc:
        mensaje["Cc"] = ", ".join(cc)
    mensaje["Subject"] = f"Trabajos de emergencia | {empresa} | 23 y 24 de Setiembre"

    empresa_html = html.escape(empresa)
    sedes = _valores_multilinea(sede)
    ncs = _valores_multilinea(nc)
    sedes_html = "<br>".join(html.escape(item) for item in sedes)
    ncs_html = "<br>".join(html.escape(item) for item in ncs)
    sedes_texto = "\n".join(sedes)
    ncs_texto = "\n".join(ncs)
    ubicacion_html = ""
    if sedes:
        ubicacion_html += f"<p><strong>SEDE:</strong><br>{sedes_html}</p>"
    if ncs:
        ubicacion_html += f"<p><strong>NC:</strong><br>{ncs_html}</p>"

    ubicacion_texto = ""
    if sedes:
        ubicacion_texto += f"SEDE:\n{sedes_texto}\n\n"
    if ncs:
        ubicacion_texto += f"NC:\n{ncs_texto}\n\n"

    cuerpo_html = f"""\
<!doctype html>
<html lang="es">
<body style="margin:0;font-family:Arial,sans-serif;color:#222;">
    <div style="max-width:900px;padding:12px 0;line-height:1.45;">
        <p>Estimado Cliente:</p>
        <p>Como parte de nuestro compromiso con la continuidad y estabilidad de los servicios, les informamos que se llevarán a cabo trabajos de mantenimiento correctivo de emergencia, correspondientes al cambio de una mufa crítica de la infraestructura externa.</p>
        <p>Los trabajos se encuentran programados de acuerdo con las siguientes ventanas:</p>
        <p>Miercoles 23/09/2026 desde las 10:00 p.m. hasta el Jueves 24/09/2026 a las 07:00 a.m.</p>
        <p>Jueves 24/09/2026 desde las 10:00 p.m. hasta el Viernes 25/09/2026 a las 07:00 a.m.</p>
        <p><strong>Horas de duración:</strong> 9 horas</p>
        {ubicacion_html}
        <p>Durante el tiempo de ejecución de los trabajos podrían presentarse posibles cortes e intermitencias en el servicio dentro del periodo comprendido de la actividad.</p>
        <p>Lamentamos profundamente las molestias ocasionadas, asimismo, le indicamos que nos encontramos tomando las medidas necesarias para asegurar su realización de forma satisfactoria, minimizando el impacto en la continuidad de sus servicios.</p>
        <p>Agradecemos su comprensión.</p>
        <p>Saludos cordiales</p>
        <p><strong><span style="color:#000080;font-size:12px;">Plataforma de Atención Clientes (PAC)</span></strong></p>
        <p style="color:#000080;font-size:12px;">Gtd Perú<br>Av. La Encalada 1257, piso 14, Santiago de Surco, Lima, Perú<br>Tel: (51 1) 743-0911 / 0800-7-0202<br><img src="https://www.gtd.cl/images/079ff2755e4e1362b5e96035e0f75d4e/original.jpg" alt="Gtd Perú" style="max-width:480px;height:auto;"></p>
        <p><strong>Estimado Cliente:</strong> Tener en cuenta que posterior a las 24 horas de haber solicitado su validación, al no tener respuesta por este medio se procederá al cierre administrativo del caso.</p>
    </div>
</body>
</html>
"""
    mensaje.set_content(
            f"Estimado Cliente:\n\n"
            f"Como parte de nuestro compromiso con la continuidad y estabilidad de los servicios, les informamos que se llevarán a cabo trabajos de mantenimiento correctivo de emergencia, correspondientes al cambio de una mufa crítica de la infraestructura externa.\n\n"
            f"Los trabajos se encuentran programados de acuerdo con las siguientes ventanas:\n"
            f"Miercoles 23/09/2026: desde las 10:00 p.m. hasta el Jueves 24/09/2026 a las 07:00 a.m.\n\n"
            f"Jueves 24/09/2026 desde las 10:00 p.m. hasta el Viernes 25/09/2026 a las 07:00 a.m.\n\n"
            f"Horas de duración: 9 horas\n\n"
            f"{ubicacion_texto}"
            f"Durante el tiempo de ejecución de los trabajos podrían presentarse posibles cortes e intermitencias en el servicio dentro del periodo comprendido de la actividad.\n"
            f"Lamentamos profundamente las molestias ocasionadas, asimismo, le indicamos que nos encontramos tomando las medidas necesarias para asegurar su realización de forma satisfactoria, minimizando el impacto en la continuidad de sus servicios.\n\n"
            f"Agradecemos su comprensión.\n\n"
            f"Saludos cordiales\n"
            f"Plataforma de Atención Clientes (PAC)\n"
            f"Gtd Perú\n"
            f"Av. La Encalada 1257, piso 14, Santiago de Surco, Lima, Perú\n"
            f"Tel: (51 1) 743-0911 / 0800-7-0202\n\n"
            f"Estimado Cliente: Tener en cuenta que posterior a las 24 horas de haber solicitado su validación, al no tener respuesta por este medio se procederá al cierre administrativo del caso."
    )
    mensaje.add_alternative(cuerpo_html, subtype="html")

    return mensaje


def main() -> None:
    raise SystemExit("El envío por terminal está deshabilitado; utiliza la aplicación Flask con ENVIO_REAL controlado.")


if __name__ == "__main__":
    main()
