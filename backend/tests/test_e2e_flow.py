from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from fastapi.testclient import TestClient

from app.main import app
import app.services.run_history as run_history

SAMPLE_DIR = Path(__file__).resolve().parent / "fixtures" / "sample_zip"


def _build_zip_from_paths(base_dir: Path, entries: list[str]) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as archive:
        for entry in entries:
            source = base_dir / entry
            if source.is_dir():
                for child in source.rglob("*"):
                    if child.is_file():
                        archive.write(child, child.relative_to(base_dir).as_posix())
            elif source.is_file():
                archive.write(source, source.relative_to(base_dir).as_posix())
    return buffer.getvalue()


def _build_full_sample_zip(base_dir: Path) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as archive:
        for child in base_dir.rglob("*"):
            if child.is_file():
                archive.write(child, child.relative_to(base_dir).as_posix())
    return buffer.getvalue()


def test_e2e_critical_flow(tmp_path, monkeypatch) -> None:
    assert SAMPLE_DIR.exists()
    monkeypatch.setattr(run_history, "HISTORY_FILE", tmp_path / "execution_history.json")

    client = TestClient(app)

    health = client.get("/api/health")
    assert health.status_code == 200
    assert health.json().get("status") == "ok"

    sample_zip = _build_full_sample_zip(SAMPLE_DIR)
    analyze_response = client.post(
        "/api/analyze/upload",
        files={"file": ("sample.zip", sample_zip, "application/zip")},
        data={"profile_name": "Cliente E2E"},
    )
    assert analyze_response.status_code == 200
    analysis_payload = analyze_response.json()
    assert analysis_payload.get("run_id")
    run_id = analysis_payload["run_id"]
    assert analysis_payload["report"]["summary"]["errors"] >= 0

    dataset_json = json.dumps(analysis_payload["dataset"], ensure_ascii=False)
    report_json = json.dumps(analysis_payload["report"], ensure_ascii=False)

    attachments_zip = _build_zip_from_paths(SAMPLE_DIR, ["Documentos contratos", "Documentos CLID"])
    attachments_response = client.post(
        "/api/attachments/validate",
        data={"dataset_json": dataset_json},
        files={"attachments_zip": ("attachments.zip", attachments_zip, "application/zip")},
    )
    assert attachments_response.status_code == 200
    assert "summary" in attachments_response.json()

    export_package_response = client.post(
        "/api/export/package",
        json={
            "dataset": analysis_payload["dataset"],
            "report": analysis_payload["report"],
            "include_report_json": True,
            "run_id": run_id,
        },
    )
    assert export_package_response.status_code == 200
    assert len(export_package_response.content) > 500

    export_with_attachments_response = client.post(
        "/api/export/package-with-attachments",
        data={
            "dataset_json": dataset_json,
            "report_json": report_json,
            "include_report_json": "true",
            "run_id": run_id,
        },
        files={"attachments_zip": ("attachments.zip", attachments_zip, "application/zip")},
    )
    assert export_with_attachments_response.status_code == 200
    assert len(export_with_attachments_response.content) > 500

    export_report_response = client.post(
        "/api/export/report.xlsx",
        json={
            "report": analysis_payload["report"],
            "executive_summary": analysis_payload.get("executive_summary"),
            "run_id": run_id,
        },
    )
    assert export_report_response.status_code == 200
    assert (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        in export_report_response.headers.get("content-type", "")
    )

    runs_response = client.get("/api/runs?limit=10")
    assert runs_response.status_code == 200
    runs = runs_response.json().get("runs", [])
    current_run = next((item for item in runs if item.get("run_id") == run_id), None)
    assert current_run is not None
    assert current_run["profile_name"] == "Cliente E2E"
    assert current_run["artifacts"]["package_zip"] is True
    assert current_run["artifacts"]["package_with_attachments_zip"] is True
    assert current_run["artifacts"]["report_xlsx"] is True
