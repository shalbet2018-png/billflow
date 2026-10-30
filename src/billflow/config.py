import calendar
import re
from pathlib import Path

import yaml


class ConfigError(Exception):
    pass


def load_config(path):
    path = Path(path).resolve()
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")
    cfg = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    for key in ("company", "clients", "charges", "outputs"):
        if key not in cfg:
            raise ConfigError(f"Config is missing required section '{key}'")
    if not cfg["charges"]:
        raise ConfigError("Config needs at least one charge")
    cfg["_dir"] = path.parent
    return cfg


def parse_period(period):
    m = re.fullmatch(r"(\d{4})-(\d{2})", period or "")
    if not m or not 1 <= int(m.group(2)) <= 12:
        raise ConfigError("Period must look like 2026-11")
    return int(m.group(1)), int(m.group(2))


def period_vars(period):
    year, month = parse_period(period)
    return {"period": period, "year": year, "month": f"{month:02d}", "month_name": calendar.month_name[month]}


def safe_name(text):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(text)).strip() or "unnamed"


def resolve_path(cfg, template, **vars):
    return (cfg["_dir"] / str(template).format(**vars)).resolve() if not Path(str(template)).is_absolute() \
        else Path(str(template).format(**vars))
