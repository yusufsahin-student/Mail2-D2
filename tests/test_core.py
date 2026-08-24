from __future__ import annotations

import csv
import base64
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook

from mailbot_core import (
    MailBotError,
    Record,
    SMTPConfig,
    SMTPMailer,
    build_message,
    canonical_name,
    create_previews,
    load_records,
    random_delay_seconds,
    render_template,
)


class TemplateTests(unittest.TestCase):
    def test_turkish_header_matching_and_literal_parentheses(self) -> None:
        result = render_template(
            "Merhaba (ISIM), (ŞIRKET) ((not))",
            {"İsim": "Ayşe", "şirket": "Robotik A.Ş."},
        )
        self.assertEqual(result.text, "Merhaba Ayşe, Robotik A.Ş. (not)")
        self.assertEqual(result.missing_headers, ())
        self.assertEqual(result.empty_fields, ())

    def test_missing_and_empty_fields_are_reported(self) -> None:
        result = render_template("(isim) / (soyisim) / (şehir)", {"isim": "", "soyisim": "Yılmaz"})
        self.assertEqual(result.empty_fields, ("isim",))
        self.assertEqual(result.missing_headers, ("şehir",))

    def test_canonical_names(self) -> None:
        self.assertEqual(canonical_name("  ŞİRKET  ADI "), canonical_name("sirket adi"))


class PreviewTests(unittest.TestCase):
    def test_invalid_empty_and_duplicate_rows_are_skipped(self) -> None:
        records = [
            Record(2, {"email": "a@example.com", "isim": "Ada"}),
            Record(3, {"email": "A@example.com", "isim": "Ali"}),
            Record(4, {"email": "bozuk", "isim": ""}),
        ]
        previews = create_previews(records, "email", "Merhaba (isim)", "Sayın (isim)")
        self.assertTrue(previews[0].valid)
        self.assertFalse(previews[1].valid)
        self.assertIn("Tekrarlanan e-posta", previews[1].problems)
        self.assertFalse(previews[2].valid)
        self.assertTrue(any(problem.startswith("Boş alan") for problem in previews[2].problems))


class DelayTests(unittest.TestCase):
    def test_random_delay_stays_inside_requested_range(self) -> None:
        values = [random_delay_seconds(30, 60) for _ in range(500)]
        self.assertTrue(all(30 <= value <= 60 for value in values))
        self.assertGreater(max(values), min(values))

    def test_equal_limits_produce_fixed_delay(self) -> None:
        self.assertEqual(random_delay_seconds(45, 45), 45)


class CsvTests(unittest.TestCase):
    def test_semicolon_csv_is_loaded(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "alicilar.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(["isim", "email"])
                writer.writerow(["Deniz", "deniz@example.com"])
            headers, records = load_records(path)
            self.assertEqual(headers, ["isim", "email"])
            self.assertEqual(records[0].values["isim"], "Deniz")
            self.assertEqual(records[0].row_number, 2)


class ExcelTests(unittest.TestCase):
    def test_xlsx_is_loaded_with_selected_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "alicilar.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sponsorlar"
            sheet.append(["İsim", "Şirket", "E-posta"])
            sheet.append(["Ece", "Robotik Ltd.", "ece@example.com"])
            workbook.save(path)
            workbook.close()

            headers, records = load_records(path, "Sponsorlar")
            self.assertEqual(headers, ["İsim", "Şirket", "E-posta"])
            self.assertEqual(records[0].values["Şirket"], "Robotik Ltd.")
            self.assertEqual(records[0].row_number, 2)


class MessageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.config = SMTPConfig(
            host="smtp.example.com",
            port=587,
            security="STARTTLS",
            username="team@example.com",
            password="secret",
            sender_address="team@example.com",
            sender_name="FRC Takımı",
        )

    def test_message_headers_and_body(self) -> None:
        message = build_message(self.config, "sponsor@example.com", "Konu", "Metin")
        self.assertEqual(message["To"], "sponsor@example.com")
        self.assertEqual(message["Subject"], "Konu")
        self.assertIn("Metin", message.get_content())

    def test_subject_header_injection_is_rejected(self) -> None:
        with self.assertRaises(MailBotError):
            build_message(self.config, "sponsor@example.com", "Konu\nBcc: x@example.com", "Metin")

    @patch("mailbot_core.smtplib.SMTP")
    def test_starttls_login_and_send_flow(self, smtp_class: MagicMock) -> None:
        connection = smtp_class.return_value
        with SMTPMailer(self.config) as mailer:
            mailer.send("sponsor@example.com", "Konu", "Metin")
        connection.starttls.assert_called_once()
        connection.login.assert_called_once_with("team@example.com", "secret")
        connection.send_message.assert_called_once()
        connection.quit.assert_called_once()

    @patch("mailbot_core.smtplib.SMTP")
    def test_unicode_password_uses_utf8_sasl_plain(self, smtp_class: MagicMock) -> None:
        connection = smtp_class.return_value
        connection.esmtp_features = {"auth": "PLAIN LOGIN"}
        connection.docmd.return_value = (235, b"Authentication successful")
        unicode_config = SMTPConfig(
            host="smtp.example.com",
            port=587,
            security="STARTTLS",
            username="takim@example.com",
            password="Şifre-123",
            sender_address="takim@example.com",
        )

        with SMTPMailer(unicode_config):
            pass

        connection.login.assert_not_called()
        command, argument = connection.docmd.call_args_list[0].args
        self.assertEqual(command, "AUTH")
        self.assertTrue(argument.startswith("PLAIN "))
        decoded = base64.b64decode(argument.split(" ", 1)[1])
        self.assertEqual(decoded, "\0takim@example.com\0Şifre-123".encode("utf-8"))


if __name__ == "__main__":
    unittest.main()
