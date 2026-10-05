import argparse
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


def backup_database(source_path: Path, backup_directory: Path) -> Path:
    if not source_path.is_file():
        raise FileNotFoundError(f"Base SQLite no encontrada: {source_path}")
    backup_directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    if os.name == "posix":
        os.chmod(backup_directory, 0o700)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    destination = backup_directory / f"correo-{timestamp}.sqlite3"
    source = sqlite3.connect(source_path)
    backup = sqlite3.connect(destination)
    try:
        source.backup(backup)
    finally:
        backup.close()
        source.close()
    if os.name == "posix":
        os.chmod(destination, 0o600)
    return destination


def main() -> int:
    parser = argparse.ArgumentParser(description="Crear una copia consistente de la base SQLite de la aplicacion.")
    parser.add_argument("backup_directory", type=Path, help="Directorio de respaldo protegido y, preferiblemente, cifrado")
    parser.add_argument(
        "--database",
        type=Path,
        default=Path(os.environ.get("APP_DATA_DIR", "instance")) / "correo.sqlite3",
        help="Ruta de la base de datos (por defecto APP_DATA_DIR/correo.sqlite3)",
    )
    arguments = parser.parse_args()
    destination = backup_database(arguments.database, arguments.backup_directory)
    print(f"Respaldo creado: {destination}")
    print("El archivo contiene datos personales; guarde el destino en almacenamiento cifrado y con acceso restringido.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
