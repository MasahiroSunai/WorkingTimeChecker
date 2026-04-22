# -*- coding: utf-8 -*-
import logging
from pathlib import Path
from datetime import datetime, timedelta

def setup_logger(name: str):
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / f"{name}_{datetime.now():%Y%m%d}.log"

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger  # 二重登録防止

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    )

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(formatter)

    sh = logging.StreamHandler()
    sh.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(sh)

    return logger


def cleanup_logs(log_dir="logs", keep_days=7):
    limit = datetime.now() - timedelta(days=keep_days)
    for f in Path(log_dir).glob("*.log"):
        if datetime.fromtimestamp(f.stat().st_mtime) < limit:
            f.unlink()
