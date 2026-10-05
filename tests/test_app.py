import io
import contextlib
import re
import sqlite3
import tempfile
import unittest
from datetime import date, time
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook

import app
import worker
from scripts import smtp_probe
from scripts.backup_sqlite import backup_database
from werkzeug.security import generate_password_hash


class SpreadsheetTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        app.app.config.update(
            TESTING=True,
            DATABASE=f"{self.temp_dir.name}/test.sqlite3",
            WTF_CSRF_ENABLED=True,
        )
        with app.app.app_context():
            app.init_db()
            app.get_db().execute(
                "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                ("testuser", generate_password_hash("test-password-123")),
            )
            app.get_db().commit()

    def tearDown(self):
        self.temp_dir.cleanup()

    def csrf_token(self, response):
        match = re.search(r'name="csrf_token" value="([^"]+)"', response.get_data(as_text=True))
        self.assertIsNotNone(match)
        return match.group(1)

    def login(self, client):
        return self.login_as(client, "testuser", "test-password-123")

    def login_as(self, client, username, password):
        login_page = client.get("/login")
        response = client.post(
            "/login",
            data={
                "csrf_token": self.csrf_token(login_page),
                "username": username,
                "password": password,
            },
        )
        self.assertEqual(response.status_code, 302)

    def test_reads_comma_csv_and_preserves_semicolon_lists(self):
        content = (
            'EMPRESA,DESTINATARIO,CC,SEDE\n'
            'Acme,"one@example.com;two@example.com",copy@example.com,"Lima;Callao"\n'
        ).encode()
        headers, rows = app.read_spreadsheet("data.csv", content)
        self.assertIn("SEDE", headers)
        self.assertEqual(rows[0]["SEDE"], "Lima;Callao")
        self.assertEqual(app.split_emails(rows[0]["DESTINATARIO"]), ["one@example.com", "two@example.com"])

    def test_reads_semicolon_delimited_csv(self):
        content = (
            'EMPRESA;DESTINATARIO;CC;SEDE\n'
            'Acme;one@example.com;copy@example.com;"Lima;Callao"\n'
        ).encode()
        _, rows = app.read_spreadsheet("data.csv", content)
        self.assertEqual(rows[0]["SEDE"], "Lima;Callao")

    def test_reads_xlsx_dates_and_times(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["EMPRESA", "DESTINATARIO", "CC", "FECHA", "HORA_INICIO"])
        sheet.append(["Acme", "one@example.com", "copy@example.com", date(2026, 10, 2), time(22, 0)])
        stream = io.BytesIO()
        workbook.save(stream)
        _, rows = app.read_spreadsheet("data.xlsx", stream.getvalue())
        self.assertEqual(rows[0]["FECHA"], "02/10/2026")
        self.assertEqual(rows[0]["HORA_INICIO"], "22:00")

    def test_requires_cc_and_rejects_malformed_addresses(self):
        rows = [{
            "_row_number": 2,
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@",
            "CC": "copy@example.com;copy@domain.",
            "ARCHIVO": "",
        }]
        problems = app.validate_rows(rows, False, set())
        self.assertEqual(len(problems), 1)
        self.assertTrue(any("DESTINATARIO" in issue for issue in problems[0]["errors"]))
        self.assertTrue(any("CC" in issue for issue in problems[0]["errors"]))

    def test_missing_attachment_blocks_row(self):
        rows = [{
            "_row_number": 2,
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@example.com",
            "CC": "copy@example.com",
            "ARCHIVO": "notice.pdf",
        }]
        self.assertIn("adjunto", " ".join(app.validate_rows(rows, True, set())[0]["errors"]))
        self.assertEqual(app.validate_rows(rows, True, {"notice.pdf"}), [])

    def test_attachment_extension_and_signature_are_checked(self):
        image = app.validate_attachment("logo.png", b"\x89PNG\r\n\x1a\ncontent")
        self.assertEqual(image["mime_type"], "image/png")
        with self.assertRaises(ValueError):
            app.validate_attachment("script.exe", b"MZ executable")
        with self.assertRaises(ValueError):
            app.validate_attachment("fake.pdf", b"not-a-pdf")

    def test_template_hides_empty_optional_section_and_escapes_values(self):
        template = "{{#SEDE}}<p>{{SEDE}}</p>{{/SEDE}}<p>{{EMPRESA}}</p>"
        rendered = app.render_body(template, {"SEDE": "", "EMPRESA": "A&B <Corp>"})
        self.assertEqual(rendered, "<p>A&amp;B &lt;Corp&gt;</p>")

    def test_formats_two_overnight_windows_in_spanish(self):
        row = {
            "FECHA_INICIO": "9/23/2026",
            "FECHA_TERMINO": "9/25/2026",
            "HORA_INICIO": "22:00",
            "HORA_CULMINACION": "07:00",
        }
        self.assertEqual(
            app.build_windows(row),
            "Miércoles 23/09/2026 desde las 10:00 p.m. hasta el Jueves 24/09/2026 a las 07:00 a.m.\n"
            "Jueves 24/09/2026 desde las 10:00 p.m. hasta el Viernes 25/09/2026 a las 07:00 a.m.",
        )
        self.assertEqual(
            app.render_body("{{#VENTANAS}}{{VENTANAS}}{{/VENTANAS}}", row, False),
            app.build_windows(row),
        )
        self.assertEqual(app.format_time_value("11:00"), "11:00 a.m.")
        self.assertEqual(app.format_duration("9:00"), "9 horas")
        self.assertEqual(app.format_duration("0:30"), "30 minutos")

    def test_preview_renders_two_windows_from_sheet_dates(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        response = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC,FECHA_INICIO,HORA_INICIO,FECHA_TERMINO,HORA_CULMINACION,DURACION\n"
                    b"Acme,person@example.com,copy@example.com,23/09/2026,22:00,25/09/2026,07:00,9:00\n"
                ), "windows.csv"),
                "body_format": "text",
                "body_template": app.DEFAULT_TEXT_BODY,
                "attach_enabled": "no",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(
            "Miércoles 23/09/2026 desde las 10:00 p.m. hasta el Jueves 24/09/2026 a las 07:00 a.m.".encode(),
            response.data,
        )
        self.assertIn(b"Jueves 24/09/2026 desde las 10:00 p.m.", response.data)
        self.assertIn(b"Horas de duraci\xc3\xb3n: 9 horas", response.data)

    def test_date_template_rejects_inverted_window_dates(self):
        rows = [{
            "_row_number": 2,
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@example.com",
            "CC": "copy@example.com",
            "FECHA_INICIO": "9/24/2026",
            "FECHA_TERMINO": "9/9/2026",
            "HORA_INICIO": "22:00",
            "HORA_CULMINACION": "07:00",
        }]
        problems = app.validate_rows(rows, False, set(), "{{VENTANAS}}")
        self.assertTrue(any("posterior" in issue for issue in problems[0]["errors"]))

    def test_message_includes_selected_attachment(self):
        campaign = {
            "subject_prefix": "Aviso",
            "suffix_mode": "fecha",
            "suffix_text": "",
            "body_template": "<p>{{EMPRESA}}</p>",
            "attach_enabled": True,
            "attachments": {"notice.pdf": {"filename": "notice.pdf", "content": b"test-pdf"}},
        }
        row = {
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@example.com",
            "CC": "copy@example.com",
            "ARCHIVO": "notice.pdf",
            "FECHA": "02/10/2026",
        }
        message = app.make_message(row, campaign, "Sender", "sender@example.com")
        parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
        attachments = list(parsed.iter_attachments())
        self.assertEqual(parsed["Subject"], "Aviso | Acme | Viernes 02/10/2026")
        self.assertEqual(len(attachments), 1)
        self.assertEqual(attachments[0].get_filename(), "notice.pdf")

    def test_plain_text_template_keeps_text_and_includes_fixed_html_signature(self):
        campaign = {
            "subject_prefix": "Aviso",
            "suffix_mode": "fixed",
            "suffix_text": "Mantenimiento",
            "body_format": "text",
            "body_template": "Empresa: {{EMPRESA}}\n{{#SEDE}}Sedes:\n{{SEDE}}{{/SEDE}}",
            "attach_enabled": False,
            "attachments": {},
        }
        row = {
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@example.com",
            "CC": "copy@example.com",
            "SEDE": "Lima;Callao",
        }
        message = app.make_message(row, campaign, "Sender", "sender@example.com")
        parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
        self.assertEqual(parsed.get_content_type(), "multipart/alternative")
        plain_part = parsed.get_body(preferencelist=("plain",))
        html_part = parsed.get_body(preferencelist=("html",))
        self.assertIn("Lima\nCallao", plain_part.get_content())
        self.assertNotIn("<html", plain_part.get_content())
        self.assertIn("original.jpg", html_part.get_content())
        self.assertIn("Plataforma de Atención Clientes (PAC)", html_part.get_content())

    def test_preview_route_does_not_send(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        response = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC,FECHA\nAcme,person@example.com,copy@example.com,02/10/2026\n"
                ), "list.csv"),
                "subject_prefix": "Aviso",
                "suffix_mode": "fecha",
                "suffix_text": "",
                "body_format": "text",
                "body_template": "Empresa: {{EMPRESA}}",
                "attach_enabled": "no",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Aviso | Acme | Viernes 02/10/2026", response.data)
        self.assertIn(b"Enviar 1 correos", response.data)
        self.assertNotIn("Prueba sin envío".encode(), response.data)

    def test_anonymous_user_is_redirected_to_login(self):
        response = app.app.test_client().get("/")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/login", response.headers["Location"])

    def test_health_endpoint_returns_only_minimal_status_without_login(self):
        response = app.app.test_client().get("/healthz")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json, {"status": "ok"})

    def test_smtp_probe_never_authenticates_or_sends(self):
        server = MagicMock()
        server.__enter__.return_value = server
        server.ehlo.return_value = (250, b"capabilities")
        server.has_extn.return_value = True
        server.esmtp_features = {"auth": "PLAIN LOGIN"}
        output = io.StringIO()
        with patch("scripts.smtp_probe.socket.getaddrinfo", return_value=[(2, 1, 6, "", ("192.0.2.1", 25))]), \
             patch("scripts.smtp_probe.smtplib.SMTP", return_value=server) as smtp_factory, \
             contextlib.redirect_stdout(output):
            self.assertEqual(smtp_probe.main(), 0)
        smtp_factory.assert_called_once()
        server.starttls.assert_called_once()
        server.login.assert_not_called()
        server.send_message.assert_not_called()

    def test_login_password_is_hashed_and_logout_requires_csrf(self):
        with app.app.app_context():
            stored_hash = app.get_db().execute(
                "SELECT password_hash FROM users WHERE username = ?", ("testuser",)
            ).fetchone()["password_hash"]
        self.assertNotEqual(stored_hash, "test-password-123")
        client = app.app.test_client()
        self.login(client)
        home = client.get("/")
        logout_token = self.csrf_token(home)
        self.assertEqual(client.post("/logout").status_code, 400)
        self.assertEqual(client.post("/logout", data={"csrf_token": logout_token}).status_code, 302)

    def test_login_temporarily_locks_after_repeated_failures(self):
        client = app.app.test_client()
        for _ in range(5):
            login_page = client.get("/login")
            response = client.post(
                "/login",
                data={
                    "csrf_token": self.csrf_token(login_page),
                    "username": "testuser",
                    "password": "wrong-password",
                },
            )
            self.assertEqual(response.status_code, 401)
        login_page = client.get("/login")
        locked = client.post(
            "/login",
            data={
                "csrf_token": self.csrf_token(login_page),
                "username": "testuser",
                "password": "wrong-password",
            },
        )
        self.assertEqual(locked.status_code, 429)

    def test_preview_post_without_csrf_is_rejected(self):
        client = app.app.test_client()
        self.login(client)
        response = client.post(
            "/preview",
            data={"spreadsheet": (io.BytesIO(b""), "empty.csv")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)

    def test_real_send_form_is_available_after_valid_preview(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        preview = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC\nAcme,person@example.com,copy@example.com\n"
                ), "ready-to-send.csv"),
                "body_template": "Empresa: {{EMPRESA}}",
                "attach_enabled": "no",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        send_path = re.search(r'action="([^"]*/send/[^"]+)"', preview.get_data(as_text=True))
        self.assertIsNotNone(send_path)
        self.assertIn(b"Enviar 1 correos", preview.data)
        self.assertIn(b'name="smtp_password"', preview.data)
        self.assertNotIn(b"Prueba sin env", preview.data)
        with patch("app.smtplib.SMTP") as smtp_factory, patch("app.send_campaign") as send_campaign:
            result = client.post(
                send_path.group(1),
                data={"csrf_token": self.csrf_token(preview), "confirm_send": "yes"},
                follow_redirects=True,
            )
        self.assertEqual(result.status_code, 200)
        self.assertIn("Ingresa la clave SMTP".encode(), result.data)
        smtp_factory.assert_not_called()
        send_campaign.assert_not_called()
        with app.app.app_context():
            stored_job = app.get_db().execute(
                "SELECT status, dry_run FROM jobs WHERE source_name = ?",
                ("ready-to-send.csv",),
            ).fetchone()
        self.assertEqual(stored_job["status"], "ready")
        self.assertEqual(stored_job["dry_run"], 0)

    def test_real_mode_web_request_sends_with_entered_password(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        preview = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC\nAcme,person@example.com,copy@example.com\n"
                ), "queued.csv"),
                "body_template": "Empresa: {{EMPRESA}}",
                "attach_enabled": "no",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        send_path = re.search(r'action="([^"]*/send/[^"]+)"', preview.get_data(as_text=True))
        self.assertIsNotNone(send_path)
        with patch("app.send_campaign") as send_campaign:
            result = client.post(
                send_path.group(1),
                data={"csrf_token": self.csrf_token(preview), "confirm_send": "yes"},
                follow_redirects=True,
            )
        self.assertEqual(result.status_code, 200)
        self.assertIn(b"En cola", result.data)
        self.assertNotIn(b"smtp-password", result.data)
        send_campaign.assert_not_called()
        with app.app.app_context():
            job_id = app.get_db().execute(
                "SELECT id FROM jobs WHERE source_name = ?", ("queued.csv",)
            ).fetchone()["id"]
            status = app.get_db().execute(
                "SELECT status FROM jobs WHERE source_name = ?", ("queued.csv",)
            ).fetchone()["status"]
        self.assertEqual(status, "queued")
        with patch("worker.read_smtp_password", return_value="test-only-secret"), patch(
            "worker.send_campaign",
            return_value=[{
                "row": 2,
                "recipient": "person@example.com",
                "attachment": "",
                "ok": True,
                "detail": "Enviado",
            }],
        ) as worker_send:
            self.assertTrue(worker.process_next_job())
            worker_send.assert_called_once()
        with app.app.app_context():
            status = app.get_db().execute(
                "SELECT status FROM jobs WHERE id = ?", (job_id,)
            ).fetchone()["status"]
            row_data = app.get_db().execute(
                """SELECT recipient, attachment_name FROM job_rows
                   WHERE job_id = ?""",
                (job_id,),
            ).fetchone()
        self.assertEqual(status, "sent")
        self.assertEqual(row_data["recipient"], "")
        self.assertEqual(row_data["attachment_name"], "")

    def test_attachment_and_recipient_data_are_removed_after_real_worker_completion(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        preview = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC,ARCHIVO\nAcme,person@example.com,copy@example.com,notice.pdf\n"
                ), "attachment.csv"),
                "attach_enabled": "yes",
                "attachments": (io.BytesIO(b"%PDF-1.7\ntest attachment"), "notice.pdf"),
                "body_template": "Empresa: {{EMPRESA}}",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        send_path = re.search(r'action="([^"]*/send/[^"]+)"', preview.get_data(as_text=True))
        self.assertIsNotNone(send_path)
        with app.app.app_context():
            job_id = app.get_db().execute(
                "SELECT id FROM jobs WHERE source_name = ?", ("attachment.csv",)
            ).fetchone()["id"]
            stored = app.get_db().execute(
                "SELECT filename, mime_type, length(content) AS size FROM job_attachments WHERE job_id = ?",
                (job_id,),
            ).fetchone()
        self.assertEqual(stored["filename"], "notice.pdf")
        self.assertEqual(stored["mime_type"], "application/pdf")
        self.assertGreater(stored["size"], 0)
        client.post(
            send_path.group(1),
            data={"csrf_token": self.csrf_token(preview), "confirm_send": "yes"},
        )
        with patch("worker.read_smtp_password", return_value="test-only-secret"), patch(
            "worker.send_campaign",
            return_value=[{
                "row": 2,
                "recipient": "person@example.com",
                "attachment": "notice.pdf",
                "ok": True,
                "detail": "Enviado",
            }],
        ):
            self.assertTrue(worker.process_next_job())
        with app.app.app_context():
            remaining = app.get_db().execute(
                "SELECT COUNT(*) AS total FROM job_attachments WHERE job_id = ?", (job_id,)
            ).fetchone()["total"]
            history_attachment = app.get_db().execute(
                "SELECT attachment_name FROM job_rows WHERE job_id = ?", (job_id,)
            ).fetchone()["attachment_name"]
        self.assertEqual(remaining, 0)
        self.assertEqual(history_attachment, "")

    def test_campaign_is_not_visible_to_another_user(self):
        client = app.app.test_client()
        self.login(client)
        form_page = client.get("/")
        preview = client.post(
            "/preview",
            data={
                "csrf_token": self.csrf_token(form_page),
                "spreadsheet": (io.BytesIO(
                    b"EMPRESA,DESTINATARIO,CC\nAcme,person@example.com,copy@example.com\n"
                ), "private.csv"),
                "body_template": "Empresa: {{EMPRESA}}",
                "attach_enabled": "no",
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        match = re.search(r"/preview/([^\"?]+)", preview.request.path)
        self.assertIsNotNone(match)
        with app.app.app_context():
            app.get_db().execute(
                "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                ("otheruser", generate_password_hash("other-password-123")),
            )
            app.get_db().commit()
        other_client = app.app.test_client()
        self.login_as(other_client, "otheruser", "other-password-123")
        self.assertEqual(other_client.get(f"/preview/{match.group(1)}").status_code, 404)
        self.assertEqual(other_client.get(f"/history/{match.group(1)}").status_code, 404)

    def test_interrupted_worker_jobs_require_manual_review(self):
        with app.app.app_context():
            user_id = app.get_db().execute(
                "SELECT id FROM users WHERE username = ?", ("testuser",)
            ).fetchone()["id"]
            app.get_db().execute(
                """INSERT INTO jobs (id, user_id, source_name, status, total_rows, payload_json)
                   VALUES ('interrupted-job', ?, 'old.csv', 'processing', 1, '{}')""",
                (user_id,),
            )
            app.get_db().execute(
                """INSERT INTO job_rows (job_id, row_number, recipient, status)
                   VALUES ('interrupted-job', 2, 'person@example.com', 'processing')"""
            )
            app.get_db().commit()
        worker.recover_interrupted_jobs()
        with app.app.app_context():
            status = app.get_db().execute(
                "SELECT status FROM jobs WHERE id = 'interrupted-job'"
            ).fetchone()["status"]
            row_status = app.get_db().execute(
                "SELECT status FROM job_rows WHERE job_id = 'interrupted-job'"
            ).fetchone()["status"]
        self.assertEqual(status, "needs_review")
        self.assertEqual(row_status, "needs_review")

    def test_expired_job_summaries_are_purged(self):
        with app.app.app_context():
            user_id = app.get_db().execute(
                "SELECT id FROM users WHERE username = ?", ("testuser",)
            ).fetchone()["id"]
            app.get_db().execute(
                """INSERT INTO jobs (id, user_id, source_name, status, total_rows, payload_json, created_at)
                   VALUES ('expired-job', ?, 'Lote', 'dry-run', 1, '{}', datetime('now', '-2 days'))""",
                (user_id,),
            )
            app.get_db().commit()
            app.purge_expired_data()
            remaining = app.get_db().execute(
                "SELECT COUNT(*) AS total FROM jobs WHERE id = 'expired-job'"
            ).fetchone()["total"]
        self.assertEqual(remaining, 0)

    def test_sqlite_backup_uses_consistent_database_copy(self):
        source = f"{self.temp_dir.name}/source.sqlite3"
        backup_dir = f"{self.temp_dir.name}/backups"
        database = sqlite3.connect(source)
        try:
            database.execute("CREATE TABLE sample (value TEXT)")
            database.execute("INSERT INTO sample VALUES ('test-data')")
            database.commit()
        finally:
            database.close()
        backup = backup_database(Path(source), Path(backup_dir))
        database = sqlite3.connect(backup)
        try:
            value = database.execute("SELECT value FROM sample").fetchone()[0]
        finally:
            database.close()
        self.assertEqual(value, "test-data")

    def test_existing_sqlite_schema_is_migrated_without_losing_attachment(self):
        database_path = Path(self.temp_dir.name) / "legacy.sqlite3"
        legacy_db = sqlite3.connect(database_path)
        try:
            legacy_db.execute(
                """CREATE TABLE login_attempts (
                    username TEXT PRIMARY KEY COLLATE NOCASE,
                    failures INTEGER NOT NULL,
                    window_started REAL NOT NULL,
                    locked_until REAL NOT NULL DEFAULT 0
                )"""
            )
            legacy_db.execute(
                "INSERT INTO login_attempts VALUES ('old-user', 3, 1, 0)"
            )
            legacy_db.execute(
                """CREATE TABLE job_attachments (
                    job_id TEXT NOT NULL,
                    lookup_name TEXT NOT NULL,
                    filename TEXT NOT NULL,
                    content BLOB NOT NULL,
                    PRIMARY KEY (job_id, lookup_name)
                )"""
            )
            legacy_db.execute(
                "INSERT INTO job_attachments VALUES ('old-job', 'a.pdf', 'a.pdf', ?)" ,
                (b"%PDF-legacy",),
            )
            legacy_db.commit()
        finally:
            legacy_db.close()
        original_path = app.app.config["DATABASE"]
        app.app.config["DATABASE"] = str(database_path)
        try:
            with app.app.app_context():
                app.init_db()
                migrated = app.get_db().execute(
                    "SELECT mime_type, content FROM job_attachments WHERE job_id = 'old-job'"
                ).fetchone()
                columns = {
                    row["name"] for row in app.get_db().execute("PRAGMA table_info(login_attempts)")
                }
                attempts = app.get_db().execute("SELECT COUNT(*) AS total FROM login_attempts").fetchone()["total"]
            self.assertEqual(migrated["mime_type"], "application/octet-stream")
            self.assertEqual(migrated["content"], b"%PDF-legacy")
            self.assertIn("remote_addr", columns)
            self.assertEqual(attempts, 0)
        finally:
            app.app.config["DATABASE"] = original_path

    def test_starttls_failure_prevents_smtp_authentication(self):
        rows = [{
            "_row_number": 2,
            "EMPRESA": "Acme",
            "DESTINATARIO": "person@example.com",
            "CC": "copy@example.com",
        }]
        campaign = {"rows": rows}
        with patch("app.smtplib.SMTP") as smtp_factory:
            server = smtp_factory.return_value
            server.starttls.side_effect = app.smtplib.SMTPNotSupportedError("STARTTLS unavailable")
            with self.assertRaises(app.smtplib.SMTPNotSupportedError):
                app.send_campaign(campaign, "not-a-real-password")
            server.login.assert_not_called()
            server.send_message.assert_not_called()

    def test_form_shows_fixed_logo_and_no_smtp_security_choices(self):
        client = app.app.test_client()
        self.login(client)
        response = client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"original.jpg", response.data)
        self.assertNotIn(b"Sin cifrado", response.data)
        self.assertNotIn(b"Seguridad SMTP", response.data)


if __name__ == "__main__":
    unittest.main()
