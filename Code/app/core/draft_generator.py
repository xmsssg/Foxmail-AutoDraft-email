from __future__ import annotations

import json
from pathlib import Path

from app.core.config_loader import expand_path
from app.core.duplicate_guard import sha256_file
from app.core.eml_builder import build_eml_backup
from app.core.file_stability import is_file_stable
from app.core.foxmail_mapi_importer import FoxmailMapiDraftImporter
from app.core.mapi_xml_builder import build_mapi_xml
from app.core.scanner import scan_customer_files
from app.storage.repository import DraftRepository


def subject_from_file(path: Path) -> str:
    return path.stem


class DraftGenerator:
    def __init__(self, settings: dict, customers: list[dict], repository: DraftRepository, logger) -> None:
        self.settings = settings
        self.customers = customers
        self.repository = repository
        self.logger = logger
        self.importer = FoxmailMapiDraftImporter(settings["foxmail"])

    def scan_once(self) -> None:
        scanner_settings = self.settings.get("scanner", {})
        stable_checks = int(scanner_settings.get("stable_checks", 3))
        stable_interval = float(scanner_settings.get("stable_interval_seconds", 1))

        for customer in self.customers:
            for file_path in scan_customer_files(customer):
                self.process_file(customer, file_path, stable_checks, stable_interval)

    def process_file(
        self,
        customer: dict,
        file_path: Path,
        stable_checks: int,
        stable_interval: float,
    ) -> None:
        self.logger.info("发现候选文件 customer=%s file=%s", customer["customer_name"], file_path)
        if not is_file_stable(file_path, stable_checks, stable_interval):
            self.logger.warning("文件未稳定，跳过 file=%s", file_path)
            return

        file_hash = sha256_file(file_path)
        existing = self.repository.find_by_file_hash(str(file_path), file_hash)
        if existing and existing["status"] == "imported":
            self.logger.info("文件已导入，跳过 file=%s", file_path)
            return

        stat = file_path.stat()
        if existing:
            record_id = int(existing["id"])
            self.repository.update_record(
                record_id,
                file_size=stat.st_size,
                file_mtime=str(stat.st_mtime),
                import_status="pending",
                status="pending",
                error_message=None,
            )
        else:
            record_id = self.repository.create_pending(
                customer_id=customer["customer_id"],
                customer_name=customer["customer_name"],
                file_path=str(file_path),
                file_name=file_path.name,
                file_size=stat.st_size,
                file_mtime=str(stat.st_mtime),
                file_hash=file_hash,
            )

        subject = subject_from_file(file_path)
        body_template = customer.get("body", "您好，附件为{subject}，请查收。")
        body = body_template.format(subject=subject, file_name=file_path.name)
        to_recipients = list(customer.get("to", []))
        cc_recipients = list(customer.get("cc", []))

        try:
            output_settings = self.settings["output"]
            eml_path = build_eml_backup(
                subject=subject,
                body=body,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                attachment_path=file_path,
                output_dir=Path(expand_path(output_settings["eml_dir"])),
            )
            xml_path, _copied_attachment = build_mapi_xml(
                subject=subject,
                body=body,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                attachment_path=file_path,
                xml_path=self.settings["foxmail"]["mapi_xml_path"],
                attachment_dir=self.settings["foxmail"]["mapi_attachment_dir"],
            )
            archive_xml_path = Path(expand_path(output_settings["mapi_xml_dir"])) / f"{subject}.xml"
            archive_xml_path.parent.mkdir(parents=True, exist_ok=True)
            archive_xml_path.write_text(xml_path.read_text(encoding="utf-8"), encoding="utf-8")

            self.repository.update_record(
                record_id,
                subject=subject,
                to_recipients=json.dumps(to_recipients, ensure_ascii=False),
                cc_recipients=json.dumps(cc_recipients, ensure_ascii=False),
                body=body,
                eml_path=str(eml_path),
                mapi_xml_path=str(archive_xml_path),
                import_status="generated",
                status="generated",
            )

            self.logger.info("草稿描述生成成功 eml=%s xml=%s", eml_path, xml_path)
            self.repository.update_record(record_id, import_status="importing", status="importing")
            result = self.importer.import_xml(xml_path, subject, body)

            if result.success:
                self.repository.update_record(
                    record_id,
                    import_status="success",
                    import_message=result.message,
                    imported_at=result.imported_at,
                    foxmail_msg_id=result.foxmail_msg_id,
                    foxmail_mail_path=result.foxmail_mail_path,
                    status="imported",
                    error_message=None,
                )
                self.logger.info(
                    "Foxmail 导入成功 file=%s msg_id=%s",
                    file_path,
                    result.foxmail_msg_id,
                )
            else:
                self.repository.update_record(
                    record_id,
                    import_status="failed",
                    import_message=result.message,
                    status="failed",
                    error_message=result.message,
                )
                self.logger.error("Foxmail 导入失败 file=%s error=%s", file_path, result.message)
        except Exception as exc:
            self.repository.update_record(
                record_id,
                import_status="failed",
                import_message=str(exc),
                status="failed",
                error_message=str(exc),
            )
            self.logger.exception("处理失败 file=%s", file_path)
