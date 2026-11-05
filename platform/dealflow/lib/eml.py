# -*- coding: utf-8 -*-
"""보내기 전 '임시 메일'(.eml) — 더블클릭하면 아웃룩에서 편집 가능한 새 메일로 열린다(X-Unsent: 1)."""
from __future__ import annotations

import mimetypes
import pathlib
from email.message import EmailMessage
from email.utils import formataddr


def draft(path: pathlib.Path, *, to: list[str], subject: str, body: str, cc: list[str] | None = None,
          attachments: list[pathlib.Path] = (), sender: str = "") -> pathlib.Path:
    m = EmailMessage()
    m["X-Unsent"] = "1"
    if sender:
        m["From"] = sender
    m["To"] = ", ".join(x for x in to if x)
    if cc:
        m["Cc"] = ", ".join(x for x in cc if x)
    m["Subject"] = subject
    m.set_content(body)
    for a in attachments:
        a = pathlib.Path(a)
        ctype = mimetypes.guess_type(a.name)[0] or "application/octet-stream"
        main, sub = ctype.split("/", 1)
        m.add_attachment(a.read_bytes(), maintype=main, subtype=sub, filename=a.name)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(m))
    return path


def addr(name: str, email: str) -> str:
    return formataddr((name, email)) if name else email
