from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from app.config import BASE_DIR
from app.models import (
    EXPECTED_COLUMNS,
    ClientProfile,
    SaveClientProfileRequest,
)

PROFILE_FILE = BASE_DIR / "data" / "client_profiles.json"


def _read_json_file(path: Path) -> list[dict[str, object]]:
    if not path.exists():
        return []

    try:
        content = path.read_text(encoding="utf-8")
    except OSError:
        return []

    if not content.strip():
        return []

    try:
        payload = json.loads(content)
    except json.JSONDecodeError:
        return []

    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]

    return []


def _write_json_file(path: Path, payload: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def _normalize_import_parameters(raw_parameters: dict[str, str]) -> dict[str, str]:
    valid_columns = set(EXPECTED_COLUMNS["import_projects_parameters"])
    normalized: dict[str, str] = {}
    for key, value in raw_parameters.items():
        if key in valid_columns and value is not None:
            normalized[key] = str(value).strip()
    return normalized


def list_profiles() -> list[ClientProfile]:
    raw_profiles = _read_json_file(PROFILE_FILE)
    profiles: list[ClientProfile] = []
    for item in raw_profiles:
        try:
            profiles.append(ClientProfile.model_validate(item))
        except Exception:
            continue

    return sorted(profiles, key=lambda item: item.name.lower())


def save_profile(payload: SaveClientProfileRequest) -> ClientProfile:
    profile_name = payload.name.strip()
    if not profile_name:
        raise ValueError("Nome do perfil é obrigatório.")

    profiles = list_profiles()
    now_iso = datetime.now(tz=timezone.utc).isoformat()

    normalized_profile = ClientProfile(
        name=profile_name,
        validation_rules=payload.validation_rules,
        import_parameters=_normalize_import_parameters(payload.import_parameters),
        updated_at=now_iso,
    )

    by_name = {profile.name.lower(): profile for profile in profiles}
    by_name[profile_name.lower()] = normalized_profile

    serialized = [profile.model_dump() for profile in sorted(by_name.values(), key=lambda item: item.name.lower())]
    _write_json_file(PROFILE_FILE, serialized)
    return normalized_profile


def delete_profile(profile_name: str) -> bool:
    normalized_name = profile_name.strip().lower()
    if not normalized_name:
        return False

    profiles = list_profiles()
    filtered = [profile for profile in profiles if profile.name.lower() != normalized_name]
    removed = len(filtered) != len(profiles)
    if not removed:
        return False

    serialized = [profile.model_dump() for profile in filtered]
    _write_json_file(PROFILE_FILE, serialized)
    return True

