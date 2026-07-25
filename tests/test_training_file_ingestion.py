"""Tests cho safe file ingestion: MIME, signature, zip-bomb, max-bytes, DOCX."""
import os
import struct
import zipfile

import pytest

from services.training_file_ingestion import (
    FileValidationError, detect_file_type, validate_txt_file,
    validate_paste_text, validate_docx_file, extract_docx_text,
    MAX_TEXT_BYTES,
)


class TestDetectFileType:
    def test_txt(self, tmp_path):
        p = str(tmp_path / "a.txt")
        with open(p, "w") as f:
            f.write("hello")
        assert detect_file_type(p) == "txt"

    def test_docx(self, tmp_path):
        p = str(tmp_path / "a.docx")
        with open(p, "wb") as f:
            f.write(b"PK\x03\x04")
        assert detect_file_type(p) == "docx"

    def test_unsupported(self, tmp_path):
        p = str(tmp_path / "a.pdf")
        with open(p, "wb") as f:
            f.write(b"%PDF")
        with pytest.raises(FileValidationError, match="không hỗ trợ"):
            detect_file_type(p)


class TestValidateTxtFile:
    def test_valid_utf8(self, tmp_path):
        p = str(tmp_path / "ok.txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write("Nội dung tiếng Việt OK.")
        content = validate_txt_file(p)
        assert "Nội dung" in content

    def test_empty_file(self, tmp_path):
        p = str(tmp_path / "empty.txt")
        with open(p, "w") as f:
            f.write("")
        with pytest.raises(FileValidationError, match="rỗng"):
            validate_txt_file(p)

    def test_whitespace_only(self, tmp_path):
        p = str(tmp_path / "ws.txt")
        with open(p, "w") as f:
            f.write("   \n\n  ")
        with pytest.raises(FileValidationError, match="rỗng"):
            validate_txt_file(p)

    def test_too_large(self, tmp_path):
        p = str(tmp_path / "big.txt")
        with open(p, "w") as f:
            f.write("x" * (MAX_TEXT_BYTES + 10))
        with pytest.raises(FileValidationError, match="lớn"):
            validate_txt_file(p)

    def test_mime_mismatch_zip_sig(self, tmp_path):
        p = str(tmp_path / "fake.txt")
        with open(p, "wb") as f:
            f.write(b"PK\x03\x04 some zip")
        with pytest.raises(FileValidationError, match="MIME|sai đuôi"):
            validate_txt_file(p)

    def test_not_found(self):
        with pytest.raises(FileValidationError, match="không tồn tại"):
            validate_txt_file("/nonexistent/file.txt")


class TestValidatePasteText:
    def test_valid(self):
        assert validate_paste_text("Hello content") == "Hello content"

    def test_empty(self):
        with pytest.raises(FileValidationError, match="rỗng"):
            validate_paste_text("")

    def test_whitespace_only(self):
        with pytest.raises(FileValidationError, match="rỗng"):
            validate_paste_text("  \n  ")

    def test_too_large(self):
        big = "x" * (MAX_TEXT_BYTES + 10)
        with pytest.raises(FileValidationError, match="giới hạn"):
            validate_paste_text(big)


class TestValidateDocxFile:
    def _make_minimal_docx(self, path):
        """Tạo DOCX hợp lệ tối thiểu (zip với [Content_Types].xml)."""
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",
                        '<?xml version="1.0"?>\n<Types xmlns="..."/>')
            zf.writestr("word/document.xml",
                        '<?xml version="1.0"?>\n<w:document/>')

    def test_valid_docx(self, tmp_path):
        p = str(tmp_path / "ok.docx")
        self._make_minimal_docx(p)
        assert validate_docx_file(p) is True

    def test_bad_signature(self, tmp_path):
        p = str(tmp_path / "bad.docx")
        with open(p, "wb") as f:
            f.write(b"Not a ZIP file content")
        with pytest.raises(FileValidationError, match="chữ ký"):
            validate_docx_file(p)

    def test_not_zip(self, tmp_path):
        p = str(tmp_path / "fake.docx")
        with open(p, "wb") as f:
            f.write(b"PK\x03\x04corrupt")
        with pytest.raises(FileValidationError):
            validate_docx_file(p)

    def test_missing_content_types(self, tmp_path):
        p = str(tmp_path / "noct.docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", "<doc/>")
        with pytest.raises(FileValidationError, match="Content_Types"):
            validate_docx_file(p)

    def test_too_large(self, tmp_path):
        p = str(tmp_path / "big.docx")
        with open(p, "wb") as f:
            f.write(b"PK\x03\x04")
            f.write(b"\x00" * 100)
        with pytest.raises(FileValidationError, match="lớn"):
            validate_docx_file(p, max_bytes=50)

    def test_not_found(self):
        with pytest.raises(FileValidationError, match="không tồn tại"):
            validate_docx_file("/nonexistent/file.docx")


class TestZipBombProtection:
    def test_high_ratio_entry_rejected(self, tmp_path):
        p = str(tmp_path / "bomb.docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", "<t/>")
            huge = "\x00" * 500000
            zf.writestr("word/huge.xml", huge)
        with pytest.raises(FileValidationError, match="zip bomb"):
            validate_docx_file(p)

    def test_too_many_entries_rejected(self, tmp_path):
        p = str(tmp_path / "many.docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", "<t/>")
            for i in range(10):
                zf.writestr(f"word/{i}.xml", "<x/>")
        from services.training_file_ingestion import _MAX_ZIP_ENTRIES
        with pytest.raises(FileValidationError, match="zip bomb|entry"):
            validate_docx_file(p, max_entries=5)
