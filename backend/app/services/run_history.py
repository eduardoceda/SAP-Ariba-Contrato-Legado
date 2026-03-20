from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal
from uuid import uuid4

from app.config import BASE_DIR
from app.models import (
    ExecutionRun,
    ExecutiveSummary,
    RunArtifacts,
    ValidationReport,
)

HISTORY_FILE = BASE_DIR / "data" / "execution_history.json"
MAX_RUNS = 300
ArtifactKey = Literal["package_zip", "package_with_attachments_zip", "report_xlsx"]


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


def _load_runs() -> list[ExecutionRun]:
    raw_runs = _read_json_file(HISTORY_FILE)
    runs: list[ExecutionRun] = []
    for item in raw_runs:
        try:
            runs.append(ExecutionRun.model_validate(item))
        except Exception:
            continue
    return runs


def _save_runs(runs: list[ExecutionRun]) -> None:
    serialized = [run.model_dump() for run in runs[:MAX_RUNS]]
    _write_json_file(HISTORY_FILE, serialized)


def list_runs(limit: int = 20) -> list[ExecutionRun]:
    limit = max(1, min(limit, 200))
    runs = _load_runs()
    runs.sort(key=lambda item: item.created_at, reverse=True)
    return runs[:limit]


def create_run(
    *,
    source: Literal["zip", "unified", "manual", "json", "unknown"],
    profile_name: str | None,
    report: ValidationReport,
    executive_summary: ExecutiveSummary | None,
) -> ExecutionRun:
    now_iso = datetime.now(tz=timezone.utc).isoformat()
    readiness_percent = executive_summary.readiness_percent if executive_summary else 0.0
    contracts_total = report.record_counts.get("contracts", 0)

    run = ExecutionRun(
        run_id=uuid4().hex,
        created_at=now_iso,
        source=source,
        profile_name=profile_name.strip() if profile_name and profile_name.strip() else None,
        summary=report.summary,
        record_counts=report.record_counts,
        readiness_percent=readiness_percent,
        contracts_total=contracts_total,
        artifacts=RunArtifacts(),
    )

    runs = _load_runs()
    runs = [run, *[item for item in runs if item.run_id != run.run_id]]
    _save_runs(runs)
    return run


def mark_run_artifact(run_id: str | None, artifact: ArtifactKey) -> bool:
    normalized = (run_id or "").strip()
    if not normalized:
        return False

    runs = _load_runs()
    changed = False
    for index, run in enumerate(runs):
        if run.run_id != normalized:
            continue

        artifacts = run.artifacts.model_copy(deep=True)
        setattr(artifacts, artifact, True)
        runs[index] = run.model_copy(update={"artifacts": artifacts})
        changed = True
        break

    if changed:
        _save_runs(runs)
    return changed

