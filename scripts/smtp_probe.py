import socket
import smtplib
import ssl
import sys

from enviar_correos import SMTP_PORT, SMTP_SERVER


def main() -> int:
    print(f"Servidor SMTP: {SMTP_SERVER}:{SMTP_PORT}")
    try:
        addresses = sorted({item[4][0] for item in socket.getaddrinfo(SMTP_SERVER, SMTP_PORT, type=socket.SOCK_STREAM)})
        print("DNS:", ", ".join(addresses))
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            code, response = server.ehlo()
            print(f"EHLO inicial: {code} {response.decode(errors='replace')[:200]}")
            if not server.has_extn("starttls"):
                print("STARTTLS no está anunciado; no se intentó autenticar ni enviar.")
                return 2
            server.starttls(context=ssl.create_default_context())
            code, response = server.ehlo()
            print(f"EHLO después de STARTTLS: {code} {response.decode(errors='replace')[:200]}")
            print("STARTTLS y validación del certificado: correctos")
            print("AUTH anunciado:", "Sí" if server.esmtp_features.get("auth") else "No")
            print("No se ejecutó AUTH y no se envió ningún correo.")
        return 0
    except (OSError, smtplib.SMTPException, ssl.SSLError, socket.gaierror) as error:
        print(f"Falló la comprobación de red/TLS: {type(error).__name__}")
        print("No se ejecutó AUTH y no se envió ningún correo.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
