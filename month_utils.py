# -*- coding: utf-8 -*-
from datetime import date

def resolve_target_month(config: dict) -> dict:
    """
    return: dict { YYYY, YY, MM }
    """
    target = config["system"].get("target_month", "auto")

    if target == "auto":
        today = date.today()
        year = today.year
        month = today.month
    else:
        # "2026-04" 形式
        year, month = map(int, target.split("-"))

    return {
        "YYYY": f"{year}",
        "YY": f"{year % 100:02d}",
        "MM": f"{month:02d}",
    }