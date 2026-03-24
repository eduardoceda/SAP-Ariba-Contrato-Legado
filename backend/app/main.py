from __future__ import annotations

import csv
import json
from io import BytesIO, StringIO
from zipfile import ZipFile

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from app.config import Settings
from app.models import (
    AttachmentValidationResponse,
    AnalysisResponse,
    AnalyzeJsonRequest,
    AribaDataset,
    ClientProfilesResponse,
    DEFAULT_IMPORT_PARAMETERS,
    ExecutionRunsResponse,
    EXPECTED_COLUMNS,
    ExportRequest,
    ReportExportRequest,
    SaveClientProfileRequest,
    UNIFIED_COLUMNS,
    UnifiedManualRequest,
    ValidationReport,
    ValidationRules,
    default_validation_rules,
)
from app.services.csv_io import load_ariba_dataset_from_zip_bytes, load_unified_rows_from_file
from app.services.exporter import build_package_zip, build_package_zip_with_attachments, build_report_xlsx
from app.services.profile_store import delete_profile, list_profiles, save_profile
from app.services.run_history import create_run, list_runs, mark_run_artifact
from app.services.unified_mapper import map_unified_rows_to_dataset
from app.services.validator import build_executive_summary, validate_attachment_bundle, validate_dataset

settings = Settings()
app = FastAPI(title=settings.app_name, version="1.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _parse_rules(raw_rules: str | None) -> ValidationRules:
    if not raw_rules:
        return default_validation_rules()

    try:
        return ValidationRules.model_validate_json(raw_rules)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"validation_rules inválido: {exc}") from exc


def _parse_import_parameters_override(raw_override: str | None) -> dict[str, str] | None:
    if not raw_override:
        return None

    try:
        payload = json.loads(raw_override)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"import_parameters_override inválido: {exc}") from exc

    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="import_parameters_override deve ser um objeto JSON")

    valid_columns = set(EXPECTED_COLUMNS["import_projects_parameters"])
    normalized: dict[str, str] = {}
    for key, value in payload.items():
        if key in valid_columns and value is not None:
            normalized[key] = str(value).strip()

    return normalized


def _apply_import_parameters_override(
    dataset: AribaDataset,
    override: dict[str, str] | None,
) -> AribaDataset:
    if not override:
        return dataset

    base = DEFAULT_IMPORT_PARAMETERS.copy()
    if dataset.import_projects_parameters:
        for key, value in dataset.import_projects_parameters[0].items():
            if value is not None:
                base[key] = str(value).strip()

    for key, value in override.items():
        base[key] = value

    dataset.import_projects_parameters = [base]
    return dataset


def _list_files_in_zip(data: bytes) -> set[str]:
    with ZipFile(BytesIO(data)) as zip_file:
        return {name.replace("\\", "/") for name in zip_file.namelist() if not name.endswith("/")}


def _build_analysis_response(
    dataset: AribaDataset,
    report: ValidationReport,
    rules: ValidationRules,
    run_id: str | None = None,
) -> AnalysisResponse:
    executive_summary = build_executive_summary(dataset, report)
    return AnalysisResponse(
        dataset=dataset,
        report=report,
        applied_rules=rules,
        executive_summary=executive_summary,
        run_id=run_id,
    )


def _track_run(
    *,
    response: AnalysisResponse,
    source: str,
    profile_name: str | None = None,
) -> AnalysisResponse:
    normalized_source = source if source in {"zip", "unified", "manual", "json"} else "unknown"
    run = create_run(
        source=normalized_source,  # type: ignore[arg-type]
        profile_name=profile_name,
        report=response.report,
        executive_summary=response.executive_summary,
    )
    response.run_id = run.run_id
    return response


@app.get(f"{settings.api_prefix}/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get(f"{settings.api_prefix}/rules/default")
def get_default_rules() -> ValidationRules:
    return default_validation_rules()


@app.get(f"{settings.api_prefix}/profiles", response_model=ClientProfilesResponse)
def get_client_profiles() -> ClientProfilesResponse:
    return ClientProfilesResponse(profiles=list_profiles())


@app.post(f"{settings.api_prefix}/profiles", response_model=ClientProfilesResponse)
def upsert_client_profile(payload: SaveClientProfileRequest) -> ClientProfilesResponse:
    try:
        save_profile(payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return ClientProfilesResponse(profiles=list_profiles())


@app.delete(f"{settings.api_prefix}/profiles/{{profile_name}}", response_model=ClientProfilesResponse)
def remove_client_profile(profile_name: str) -> ClientProfilesResponse:
    removed = delete_profile(profile_name)
    if not removed:
        raise HTTPException(status_code=404, detail=f"Perfil não encontrado: {profile_name}")

    return ClientProfilesResponse(profiles=list_profiles())


@app.get(f"{settings.api_prefix}/runs", response_model=ExecutionRunsResponse)
def get_execution_runs(limit: int = 20) -> ExecutionRunsResponse:
    return ExecutionRunsResponse(runs=list_runs(limit))


@app.get(f"{settings.api_prefix}/unified/schema")
def unified_schema() -> dict[str, object]:
    return {
        "columns": UNIFIED_COLUMNS,
        "example": {
            "ContractId": "LCW4700001278",
            "Title": "LCW4700001278",
            "Owner": "G571174",
            "BaseLanguage": "BrazilianPortuguese",
            "Description": "Contrato legado importado do sistema origem",
            "Supplier": "sap:0000381965",
            "AffectedParties": "sap:0000381965",
            "HierarchicalType": "MasterAgreement",
            "ParentAgreement": "",
            "ProposedAmount": "150000.00",
            "Amount": "150000.00",
            "Commodity": "Serviços de TI",
            "Region": "BRA",
            "Client": "Cliente Exemplo",
            "AgreementDate": "2023-03-01",
            "EffectiveDate": "2023-03-01",
            "ExpirationDate": "2026-02-28",
            "ContractStatus": "Published",
            "RelatedId": "4700001278",
            "TeamProjectGroup": "Comprador",
            "TeamMember": "G571174",
            "DocumentFile": "Documentos contratos/Contrato 4700001278.pdf",
            "DocumentTitle": "Contrato 4700001278",
            "DocumentFolder": "",
            "DocumentOwner": "G571174",
            "DocumentStatus": "",
            "ClidFile": "Documentos CLID/CLID_4700001278.xlsx",
            "ClidTitle": "CLID_4700001278.xlsx",
        },
    }


@app.post(f"{settings.api_prefix}/analyze/upload", response_model=AnalysisResponse)
async def analyze_ariba_package(
    file: UploadFile = File(...),
    validation_rules: str | None = Form(default=None),
    import_parameters_override: str | None = Form(default=None),
    profile_name: str | None = Form(default=None),
) -> AnalysisResponse:
    filename = file.filename or ""
    if not filename.lower().endswith(".zip"):
        raise HTTPException(status_code=400, detail="Envie um arquivo .zip")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Arquivo vazio")

    rules = _parse_rules(validation_rules)
    import_override = _parse_import_parameters_override(import_parameters_override)

    try:
        dataset, headers_by_key, missing_files, available_paths = load_ariba_dataset_from_zip_bytes(content)
    except Exception as exc:  # pragma: no cover
        raise HTTPException(status_code=400, detail=f"Falha ao ler ZIP: {exc}") from exc

    dataset = _apply_import_parameters_override(dataset, import_override)

    report = validate_dataset(
        dataset=dataset,
        headers_by_key=headers_by_key,
        missing_files=missing_files,
        available_paths=available_paths,
        rules=rules,
    )

    response = _build_analysis_response(dataset=dataset, report=report, rules=rules)
    return _track_run(response=response, source="zip", profile_name=profile_name)


@app.post(f"{settings.api_prefix}/analyze/json", response_model=AnalysisResponse)
def analyze_json_dataset(payload: AnalyzeJsonRequest) -> AnalysisResponse:
    rules = payload.validation_rules or default_validation_rules()
    report = validate_dataset(dataset=payload.dataset, rules=rules)
    response = _build_analysis_response(dataset=payload.dataset, report=report, rules=rules)
    return _track_run(response=response, source="json", profile_name=payload.profile_name)


@app.post(f"{settings.api_prefix}/unified/upload", response_model=AnalysisResponse)
async def ingest_unified_file(
    file: UploadFile = File(...),
    validation_rules: str | None = Form(default=None),
    import_parameters_override: str | None = Form(default=None),
    profile_name: str | None = Form(default=None),
) -> AnalysisResponse:
    filename = file.filename or ""
    if not filename:
        raise HTTPException(status_code=400, detail="Nome do arquivo não informado")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Arquivo vazio")

    rules = _parse_rules(validation_rules)
    import_override = _parse_import_parameters_override(import_parameters_override)

    try:
        _, rows = load_unified_rows_from_file(filename, content)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # pragma: no cover
        raise HTTPException(status_code=400, detail=f"Falha ao ler arquivo: {exc}") from exc

    dataset = map_unified_rows_to_dataset(rows)
    dataset = _apply_import_parameters_override(dataset, import_override)
    report = validate_dataset(dataset=dataset, rules=rules)
    response = _build_analysis_response(dataset=dataset, report=report, rules=rules)
    return _track_run(response=response, source="unified", profile_name=profile_name)


@app.post(f"{settings.api_prefix}/unified/manual", response_model=AnalysisResponse)
def ingest_unified_manual(payload: UnifiedManualRequest) -> AnalysisResponse:
    rules = payload.validation_rules or default_validation_rules()
    dataset = map_unified_rows_to_dataset(payload.rows, payload.import_parameters_override)
    report = validate_dataset(dataset=dataset, rules=rules)
    response = _build_analysis_response(dataset=dataset, report=report, rules=rules)
    return _track_run(response=response, source="manual", profile_name=payload.profile_name)


@app.post(f"{settings.api_prefix}/attachments/validate", response_model=AttachmentValidationResponse)
async def validate_attachments_zip(
    dataset_json: str = Form(...),
    attachments_zip: UploadFile = File(...),
) -> AttachmentValidationResponse:
    if attachments_zip.filename and not attachments_zip.filename.lower().endswith(".zip"):
        raise HTTPException(status_code=400, detail="attachments_zip deve ser um arquivo .zip")

    try:
        dataset = AribaDataset.model_validate_json(dataset_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"dataset_json inválido: {exc}") from exc

    attachments_bytes = await attachments_zip.read()
    if not attachments_bytes:
        raise HTTPException(status_code=400, detail="attachments_zip vazio")

    try:
        available_paths = _list_files_in_zip(attachments_bytes)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Falha ao ler attachments_zip: {exc}") from exc

    summary, issues, stats = validate_attachment_bundle(dataset, available_paths)
    return AttachmentValidationResponse(summary=summary, issues=issues, stats=stats)


@app.post(f"{settings.api_prefix}/export/package")
def export_package(payload: ExportRequest) -> StreamingResponse:
    zip_content = build_package_zip(
        dataset=payload.dataset,
        report=payload.report,
        include_report_json=payload.include_report_json,
    )

    mark_run_artifact(payload.run_id, "package_zip")

    return StreamingResponse(
        BytesIO(zip_content),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=ariba-package.zip"},
    )


@app.post(f"{settings.api_prefix}/export/package-with-attachments")
async def export_package_with_attachments(
    dataset_json: str = Form(...),
    include_report_json: bool = Form(default=True),
    report_json: str | None = Form(default=None),
    attachments_zip: UploadFile | None = File(default=None),
    run_id: str | None = Form(default=None),
) -> StreamingResponse:
    try:
        dataset = AribaDataset.model_validate_json(dataset_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"dataset_json inválido: {exc}") from exc

    report: ValidationReport | None = None
    if report_json:
        try:
            report = ValidationReport.model_validate_json(report_json)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"report_json inválido: {exc}") from exc

    attachments_bytes: bytes | None = None
    if attachments_zip is not None:
        if attachments_zip.filename and not attachments_zip.filename.lower().endswith(".zip"):
            raise HTTPException(status_code=400, detail="attachments_zip deve ser um arquivo .zip")
        attachments_bytes = await attachments_zip.read()

    zip_content = build_package_zip_with_attachments(
        dataset=dataset,
        attachments_zip_bytes=attachments_bytes,
        report=report,
        include_report_json=include_report_json,
    )

    mark_run_artifact(run_id, "package_with_attachments_zip")

    return StreamingResponse(
        BytesIO(zip_content),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=ariba-package-importable.zip"},
    )


@app.post(f"{settings.api_prefix}/export/report.xlsx")
def export_report_xlsx(payload: ReportExportRequest) -> StreamingResponse:
    xlsx_content = build_report_xlsx(payload.report, payload.executive_summary)
    mark_run_artifact(payload.run_id, "report_xlsx")
    return StreamingResponse(
        BytesIO(xlsx_content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=validation-report.xlsx"},
    )


@app.get(f"{settings.api_prefix}/unified/template")
def download_unified_template() -> StreamingResponse:
    format_row = {
        "ContractId": "LCW + números",
        "Title": "Texto livre",
        "Owner": "Código usuário Ariba",
        "BaseLanguage": "BrazilianPortuguese",
        "Description": "Texto livre",
        "Supplier": "sap: + números",
        "AffectedParties": "sap: + números",
        "HierarchicalType": "MasterAgreement | SubAgreement",
        "ParentAgreement": "LCW... (se SubAgreement)",
        "ProposedAmount": "Número decimal (ex.: 150000.00)",
        "Amount": "Número decimal (ex.: 150000.00)",
        "Commodity": "Texto livre",
        "Region": "BRA",
        "Client": "Texto livre",
        "AgreementDate": "YYYY-MM-DD",
        "EffectiveDate": "YYYY-MM-DD",
        "ExpirationDate": "YYYY-MM-DD",
        "ContractStatus": "Published",
        "RelatedId": "Somente números",
        "TeamProjectGroup": "Grupo de projeto Ariba",
        "TeamMember": "Código usuário Ariba",
        "DocumentFile": "Documentos contratos/arquivo.pdf",
        "DocumentTitle": "Título do documento",
        "DocumentFolder": "Opcional",
        "DocumentOwner": "Código usuário Ariba",
        "DocumentStatus": "Opcional",
        "ClidFile": "Documentos CLID/arquivo.xlsx",
        "ClidTitle": "Título do CLID",
    }
    example_row = {
        "ContractId": "LCW4700001278",
        "Title": "Contrato 4700001278",
        "Owner": "G571174",
        "BaseLanguage": "BrazilianPortuguese",
        "Description": "Contrato legado importado do sistema origem",
        "Supplier": "sap:0000381965",
        "AffectedParties": "sap:0000381965",
        "HierarchicalType": "MasterAgreement",
        "ParentAgreement": "",
        "ProposedAmount": "150000.00",
        "Amount": "150000.00",
        "Commodity": "Serviços de TI",
        "Region": "BRA",
        "Client": "Cliente Exemplo",
        "AgreementDate": "2023-03-01",
        "EffectiveDate": "2023-03-01",
        "ExpirationDate": "2026-02-28",
        "ContractStatus": "Published",
        "RelatedId": "4700001278",
        "TeamProjectGroup": "Comprador",
        "TeamMember": "G571174",
        "DocumentFile": "Documentos contratos/Contrato 4700001278.pdf",
        "DocumentTitle": "Contrato 4700001278",
        "DocumentFolder": "",
        "DocumentOwner": "G571174",
        "DocumentStatus": "",
        "ClidFile": "Documentos CLID/CLID_4700001278.xlsx",
        "ClidTitle": "CLID_4700001278.xlsx",
    }

    output = StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(UNIFIED_COLUMNS)
    writer.writerow([format_row.get(column, "") for column in UNIFIED_COLUMNS])
    writer.writerow([example_row.get(column, "") for column in UNIFIED_COLUMNS])
    content = output.getvalue()

    return StreamingResponse(
        BytesIO(content.encode("utf-8")),
        media_type="text/csv",
        headers={
            "Content-Disposition": "attachment; filename=base-unica.csv; filename*=UTF-8''base%20%C3%BAnica.csv"
        },
    )


@app.get(f"{settings.api_prefix}/attachments/template")
def download_attachments_template() -> StreamingResponse:
    buffer = BytesIO()
    with ZipFile(buffer, "w") as zip_file:
        zip_file.writestr("Documentos contratos/", "")
        zip_file.writestr("Documentos CLID/", "")

    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=modelo-anexos-clid.zip"},
    )


@app.exception_handler(HTTPException)
async def http_exception_handler(_, exc: HTTPException) -> JSONResponse:
    return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})

