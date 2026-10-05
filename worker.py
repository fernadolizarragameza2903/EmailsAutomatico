import os
import time
from pathlib import Path

from app import APP_ENV, app, fail_campaign, finish_campaign, get_db, load_campaign, purge_expired_data, send_campaign


POLL_INTERVAL_SECONDS = 5


def read_smtp_password() -> str:
    credentials_directory = os.environ.get("CREDENTIALS_DIRECTORY")
    if credentials_directory:
        credential_path = Path(credentials_directory) / "smtp_password"
        if credential_path.is_file():
            return credential_path.read_text(encoding="utf-8").rstrip("\r\n")
    if APP_ENV == "production":
        return ""
    return os.environ.get("SMTP_PASSWORD", "")


def claim_job():
    database = get_db()
    database.execute("BEGIN IMMEDIATE")
    job = database.execute(
        "SELECT id, user_id FROM jobs WHERE status = 'queued' ORDER BY created_at LIMIT 1"
    ).fetchone()
    if not job:
        database.commit()
        return None
    claimed = database.execute(
        "UPDATE jobs SET status = 'processing' WHERE id = ? AND status = 'queued'",
        (job["id"],),
    ).rowcount
    if claimed:
        database.execute(
            "UPDATE job_rows SET status = 'processing' WHERE job_id = ? AND status = 'queued'",
            (job["id"],),
        )
    database.commit()
    return (job["id"], str(job["user_id"])) if claimed else None


def process_next_job() -> bool:
    with app.app_context():
        claimed = claim_job()
        if not claimed:
            return False
        campaign_id, user_id = claimed
        campaign, _job = load_campaign(campaign_id, user_id)
        password = read_smtp_password()
        if not password:
            fail_campaign(campaign_id, "No hay credencial SMTP configurada para el worker.")
            return True
        try:
            results = send_campaign(campaign, password)
            finish_campaign(campaign_id, results)
            print(f"Lote procesado: {campaign_id}; filas={len(results)}")
        except Exception as error:
            fail_campaign(campaign_id, f"Falló el proceso SMTP: {str(error)[:400]}")
            print(f"Lote fallido: {campaign_id}")
        finally:
            password = ""
        return True


def recover_interrupted_jobs() -> None:
    with app.app_context():
        database = get_db()
        database.execute("BEGIN IMMEDIATE")
        interrupted = database.execute(
            "SELECT id FROM jobs WHERE status = 'processing'"
        ).fetchall()
        for job in interrupted:
            database.execute(
                """UPDATE job_rows SET status = 'needs_review',
                   detail = 'Worker reiniciado durante el lote; verificar entregas antes de repetir.'
                   WHERE job_id = ? AND status = 'processing'""",
                (job["id"],),
            )
            database.execute(
                """UPDATE jobs SET status = 'needs_review', error_count = total_rows,
                   payload_json = '{}', completed_at = CURRENT_TIMESTAMP WHERE id = ?""",
                (job["id"],),
            )
            database.execute("DELETE FROM job_attachments WHERE job_id = ?", (job["id"],))
        database.commit()


def main() -> None:
    recover_interrupted_jobs()
    while True:
        with app.app_context():
            purge_expired_data()
        if not process_next_job():
            time.sleep(POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
