import base64
import contextlib
import datetime as dt
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from scripts import daily_receipt_checker as checker


class GmailDownloadTests(unittest.TestCase):
    def test_query_covers_all_attachments_on_selected_day(self):
        query = checker.build_query(
            dt.date(2026, 5, 8),
            {"known_sender_query": "from:guide@example.com", "gmail_extra_query": "-in:trash"},
        )
        self.assertEqual(
            query,
            "from:guide@example.com after:2026/5/8 before:2026/5/9 "
            "has:attachment -in:trash",
        )

    def test_downloads_inline_and_separate_attachment_data(self):
        inline_data = base64.urlsafe_b64encode(b"inline image").decode("ascii")
        separate_data = base64.urlsafe_b64encode(b"separate invoice").decode("ascii")
        service = MagicMock()
        service.users.return_value.messages.return_value.get.return_value.execute.return_value = {
            "payload": {
                "headers": [
                    {"name": "From", "value": "Guide <guide@example.com>"},
                    {"name": "Subject", "value": "TKE1234567 receipt"},
                ],
                "parts": [
                    {"filename": "0508_receipt.jpg", "body": {"data": inline_data}},
                    {"filename": "0508_invoice.xlsx", "body": {"attachmentId": "att-1"}},
                ],
            }
        }
        attachment_get = service.users.return_value.messages.return_value.attachments.return_value.get
        attachment_get.return_value.execute.return_value = {"data": separate_data}

        with tempfile.TemporaryDirectory() as temporary, \
                patch.object(checker, "get_gmail_service", return_value=service), \
                patch.object(checker, "list_message_ids", return_value=["msg-1"]), \
                contextlib.redirect_stdout(io.StringIO()):
            day_folder = checker.download_gmail(
                dt.date(2026, 5, 8), Path(temporary), {}
            )
            files = {path.name: path.read_bytes() for path in day_folder.rglob("*") if path.is_file()}

        self.assertEqual(files["0508_receipt.jpg"], b"inline image")
        self.assertEqual(files["0508_invoice.xlsx"], b"separate invoice")
        attachment_get.assert_called_once_with(userId="me", messageId="msg-1", id="att-1")


if __name__ == "__main__":
    unittest.main()
