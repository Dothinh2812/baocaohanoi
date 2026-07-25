"""Safe file ingestion for knowledge import.

Kiểm tra suffix, signature/MIME, max bytes, DOCX/ZIP safety trước parser.
Không parse file chưa qua validate.
"""

import os
import struct
import zipfile

MAX_TEXT_BYTES = int(os.getenv("DASHV4_TRAINING_MAX_TEXT_BYTES", "200000"))
MAX_DOCX_BYTES = int(os.getenv("DASHV4_TRAINING_MAX_DOCX_BYTES", str(MAX_TEXT_BYTES * 2)))

_DOCX_SIGNATURE = b"PK\x03\x04"
_TXT_SIGNATURES = (b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff", b"\x00\x00\xfe\xff")
_MAX_ZIP_RATIO = 100
_MAX_ZIP_ENTRIES = 5000


class FileValidationError(ValueError):
    """File validation failed. code attribute matches ErrorCode.VALIDATION_ERROR scope."""

    def __init__(self, message, *, code="FILE_INVALID"):
        super().__init__(message)
        self.code = code


def _check_size(path, max_bytes):
    size = os.path.getsize(path)
    if size == 0:
        raise FileValidationError("File rỗng", code="FILE_EMPTY")
    if size > max_bytes:
        raise FileValidationError(
            f"File quá lớn: {size} bytes > {max_bytes} bytes",
            code="FILE_TOO_LARGE",
        )
    return size


def _read_signature(path, num_bytes=8):
    with open(path, "rb") as f:
        return f.read(num_bytes)


def validate_txt_file(path, *, max_bytes=None):
    """Validate TXT file: exists, non-empty, within size, readable as UTF-8."""
    max_bytes = max_bytes or MAX_TEXT_BYTES
    if not os.path.exists(path):
        raise FileValidationError(f"File không tồn tại: {path}", code="FILE_NOT_FOUND")
    _check_size(path, max_bytes)
    sig = _read_signature(path, 4)
    if sig[:4] == _DOCX_SIGNATURE:
        raise FileValidationError(
            "File có chữ ký ZIP/DOCX nhưng đuôi .txt — có thể sai đuôi",
            code="FILE_MIME_MISMATCH",
        )
    try:
        with open(path, "r", encoding="utf-8", errors="strict") as f:
            content = f.read(max_bytes + 1)
    except UnicodeDecodeError:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(max_bytes + 1)
    if len(content) > max_bytes:
        raise FileValidationError(
            f"Nội dung text vượt giới hạn {max_bytes} bytes",
            code="FILE_TOO_LARGE",
        )
    if not content.strip():
        raise FileValidationError("Nội dung text rỗng", code="FILE_EMPTY")
    return content


def validate_paste_text(text, *, max_bytes=None):
    """Validate pasted text string."""
    max_bytes = max_bytes or MAX_TEXT_BYTES
    if not text or not text.strip():
        raise FileValidationError("Nội dung paste rỗng", code="FILE_EMPTY")
    encoded = text.encode("utf-8")
    if len(encoded) > max_bytes:
        raise FileValidationError(
            f"Nội dung paste vượt giới hạn {max_bytes} bytes",
            code="FILE_TOO_LARGE",
        )
    return text


def _check_zip_safety(path, *, max_uncompressed_ratio=_MAX_ZIP_RATIO,
                      max_entries=_MAX_ZIP_ENTRIES):
    """Kiểm tra zip bomb và cấu trúc DOCX không hợp lệ.

    DOCX là ZIP archive chứa XML parts. Kiểm tra:
    - Số entry hợp lý (tránh zip bomb)
    - Tỷ lệ nén không quá cao
    - Có [Content_Types].xml (đặc trưng DOCX/OOXML)
    """
    if not zipfile.is_zipfile(path):
        raise FileValidationError(
            "File .docx không phải là ZIP archive hợp lệ",
            code="DOCX_NOT_ZIP",
        )
    try:
        with zipfile.ZipFile(path, "r") as zf:
            entries = zf.infolist()
            if len(entries) > max_entries:
                raise FileValidationError(
                    f"DOCX có quá nhiều entry: {len(entries)} > {max_entries}",
                    code="DOCX_ZIP_BOMB",
                )
            total_compressed = 0
            total_uncompressed = 0
            names = set()
            for entry in entries:
                total_compressed += entry.compress_size
                total_uncompressed += entry.file_size
                names.add(entry.filename)
                if entry.file_size > 0:
                    ratio = entry.file_size / max(entry.compress_size, 1)
                    if ratio > max_uncompressed_ratio:
                        raise FileValidationError(
                            f"Entry {entry.filename!r} có tỷ lệ nén {ratio:.0f}x "
                            f"(giới hạn {max_uncompressed_ratio}x) — nghi zip bomb",
                            code="DOCX_ZIP_BOMB",
                        )
            if total_compressed > 0:
                overall_ratio = total_uncompressed / total_compressed
                if overall_ratio > max_uncompressed_ratio:
                    raise FileValidationError(
                        f"Tỷ lệ nén tổng thể {overall_ratio:.0f}x — nghi zip bomb",
                        code="DOCX_ZIP_BOMB",
                    )
            if "[Content_Types].xml" not in names:
                raise FileValidationError(
                    "File ZIP không chứa [Content_Types].xml — không phải DOCX/OOXML hợp lệ",
                    code="DOCX_INVALID_STRUCTURE",
                )
    except zipfile.BadZipFile as exc:
        raise FileValidationError(
            f"DOCX hỏng: {exc}", code="DOCX_CORRUPT"
        ) from exc


def validate_docx_file(path, *, max_bytes=None, max_uncompressed_ratio=_MAX_ZIP_RATIO,
                      max_entries=_MAX_ZIP_ENTRIES):
    """Validate DOCX file: exists, size, ZIP signature, zip-bomb, structure.

    Trả (content_text, safe_to_parse=True). Không parse ở đây để caller kiểm tra
    python-docx availability trước.
    """
    max_bytes = max_bytes or MAX_DOCX_BYTES
    if not os.path.exists(path):
        raise FileValidationError(f"File không tồn tại: {path}", code="FILE_NOT_FOUND")
    _check_size(path, max_bytes)
    sig = _read_signature(path, 4)
    if sig[:4] != _DOCX_SIGNATURE:
        raise FileValidationError(
            "File không có chữ ký ZIP (PK\\x03\\x04) — không phải DOCX hợp lệ",
            code="DOCX_BAD_SIGNATURE",
        )
    _check_zip_safety(path, max_uncompressed_ratio=max_uncompressed_ratio,
                      max_entries=max_entries)
    return True


def extract_docx_text(path):
    """Extract text từ DOCX đã validated. Lazy import python-docx."""
    try:
        from docx import Document
    except ImportError:
        raise FileValidationError(
            "python-docx chưa cài đặt; chạy: pip install python-docx",
            code="DOCX_DEPENDENCY_MISSING",
        )
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n\n".join(paragraphs)


def detect_file_type(path):
    """Phát hiện loại file qua suffix + signature. Trả 'txt', 'docx', hoặc raise."""
    lower = path.lower()
    if lower.endswith(".txt"):
        sig = _read_signature(path, 4)
        if sig[:4] == _DOCX_SIGNATURE:
            raise FileValidationError(
                "File .txt có chữ ký ZIP — sai đuôi, đổi thành .docx",
                code="FILE_MIME_MISMATCH",
            )
        return "txt"
    if lower.endswith(".docx"):
        return "docx"
    raise FileValidationError(
        f"Đuôi file không hỗ trợ. Chỉ chấp nhận .txt hoặc .docx",
        code="FILE_TYPE_UNSUPPORTED",
    )
