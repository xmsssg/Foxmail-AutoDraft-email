from __future__ import annotations

import shutil
from pathlib import Path
from xml.sax.saxutils import escape

from app.core.config_loader import expand_path


def cdata(value: str) -> str:
    return "<![CDATA[" + value.replace("]]>", "]]]]><![CDATA[>") + "]]>"


def build_mapi_xml(
    *,
    subject: str,
    body: str,
    to_recipients: list[str],
    cc_recipients: list[str],
    attachment_path: Path,
    xml_path: str,
    attachment_dir: str,
) -> tuple[Path, Path]:
    xml_target = Path(expand_path(xml_path))
    attach_dir = Path(expand_path(attachment_dir))
    xml_target.parent.mkdir(parents=True, exist_ok=True)
    attach_dir.mkdir(parents=True, exist_ok=True)

    copied_attachment = attach_dir / attachment_path.name
    shutil.copy2(attachment_path, copied_attachment)

    receiver_lines: list[str] = []
    for email in to_recipients:
        receiver_lines.append(f'    <Receiver email="{escape(email)}" type="1"/>')
    for email in cc_recipients:
        receiver_lines.append(f'    <Receiver email="{escape(email)}" type="2"/>')

    xml = "\n".join(
        [
            '<FMMAPI version="1.0">',
            f"  <Subject>{cdata(subject)}</Subject>",
            "  <Receivers>",
            *receiver_lines,
            "  </Receivers>",
            "  <Attachments>",
            "    <attachment",
            f'      name="{escape(copied_attachment.name)}"',
            f'      path="{escape(str(copied_attachment))}"/>',
            "  </Attachments>",
            f"  <Content>{cdata(body)}</Content>",
            "</FMMAPI>",
            "",
        ]
    )
    xml_target.write_text(xml, encoding="utf-8")
    return xml_target, copied_attachment
