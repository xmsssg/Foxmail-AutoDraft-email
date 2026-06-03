from __future__ import annotations

from email.message import EmailMessage
from pathlib import Path


def build_eml_backup(
    *,
    subject: str,
    body: str,
    to_recipients: list[str],
    cc_recipients: list[str],
    attachment_path: Path,
    output_dir: Path,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    msg = EmailMessage()
    msg["From"] = ""
    msg["To"] = ", ".join(to_recipients)
    if cc_recipients:
        msg["Cc"] = ", ".join(cc_recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    data = attachment_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_path.name,
    )

    eml_path = output_dir / f"{subject}.eml"
    if eml_path.exists():
        eml_path = output_dir / f"{subject}_{int(attachment_path.stat().st_mtime)}.eml"
    eml_path.write_bytes(msg.as_bytes())
    return eml_path
