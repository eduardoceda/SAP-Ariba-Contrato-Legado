from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    app_name: str = "Ariba Legacy Contracts Validator"
    api_prefix: str = "/api"


BASE_DIR = Path(__file__).resolve().parent.parent
