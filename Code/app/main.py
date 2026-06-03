from __future__ import annotations

import argparse

from app.core.config_loader import load_customers, load_settings
from app.desktop_app import run_desktop_app
from app.core.draft_generator import DraftGenerator
from app.core.logger import setup_logger
from app.storage.repository import DraftRepository


def main() -> None:
    parser = argparse.ArgumentParser(description="发货单邮件草稿自动生成工具")
    parser.add_argument("--scan-once", action="store_true", help="扫描一次客户目录并处理新增文件")
    parser.add_argument("--gui", action="store_true", help="打开桌面操作界面")
    args = parser.parse_args()

    if args.gui or not args.scan_once:
        run_desktop_app()
        return

    settings = load_settings()
    customers = load_customers()
    logger = setup_logger(settings)

    repository = DraftRepository(settings["database"]["path"])
    repository.init_schema()

    generator = DraftGenerator(settings, customers, repository, logger)
    generator.scan_once()


if __name__ == "__main__":
    main()
