import json
import unicodedata
from pathlib import Path
from typing import Any


def normalize_text(value: Any, *, newline: str = "\n") -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if newline != "\n":
        text = text.replace("\n", newline)
    return unicodedata.normalize("NFC", text)


def read_text_utf8(path: str | Path, *, errors: str = "strict") -> str:
    return Path(path).read_text(encoding="utf-8", errors=errors)


def write_text_utf8(path: str | Path, content: Any, *, newline: str = "\n") -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(normalize_text(content, newline=newline), encoding="utf-8", newline=newline)
    return target


def read_json_utf8(path: str | Path, *, default: Any = None) -> Any:
    target = Path(path)
    if not target.exists():
        return default
    with target.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def write_json_utf8(
    path: str | Path,
    data: Any,
    *,
    indent: int = 2,
    ensure_ascii: bool = False,
) -> Path:
    payload = json.dumps(data, indent=indent, ensure_ascii=ensure_ascii)
    return write_text_utf8(path, payload + "\n")
