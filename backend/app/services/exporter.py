from __future__ import annotations

import json
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from app.models import AribaDataset, ExecutiveSummary, ValidationReport
from app.services.csv_io import dataset_to_csv_string

RESERVED_EXPORT_FILES = {
    "contracts.csv",
    "contractdocuments.csv",
    "contractcontentdocuments.csv",
    "contractteams.csv",
    "importprojectsparameters.csv",
    "validation-report.json",
}


def build_package_zip(dataset: AribaDataset, report: ValidationReport | None = None, include_report_json: bool = True) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as zip_file:
        zip_file.writestr("Contracts.csv", dataset_to_csv_string("contracts", dataset.contracts))
        zip_file.writestr("ContractDocuments.csv", dataset_to_csv_string("contract_documents", dataset.contract_documents))
        zip_file.writestr(
            "ContractContentDocuments.csv",
            dataset_to_csv_string("contract_content_documents", dataset.contract_content_documents),
        )
        zip_file.writestr("ContractTeams.csv", dataset_to_csv_string("contract_teams", dataset.contract_teams))
        zip_file.writestr(
            "ImportProjectsParameters.csv",
            dataset_to_csv_string("import_projects_parameters", dataset.import_projects_parameters),
        )

        if include_report_json and report is not None:
            zip_file.writestr("validation-report.json", json.dumps(report.model_dump(), indent=2, ensure_ascii=False))

    return buffer.getvalue()


def build_package_zip_with_attachments(
    dataset: AribaDataset,
    attachments_zip_bytes: bytes | None = None,
    report: ValidationReport | None = None,
    include_report_json: bool = True,
) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as zip_file:
        zip_file.writestr("Contracts.csv", dataset_to_csv_string("contracts", dataset.contracts))
        zip_file.writestr("ContractDocuments.csv", dataset_to_csv_string("contract_documents", dataset.contract_documents))
        zip_file.writestr(
            "ContractContentDocuments.csv",
            dataset_to_csv_string("contract_content_documents", dataset.contract_content_documents),
        )
        zip_file.writestr("ContractTeams.csv", dataset_to_csv_string("contract_teams", dataset.contract_teams))
        zip_file.writestr(
            "ImportProjectsParameters.csv",
            dataset_to_csv_string("import_projects_parameters", dataset.import_projects_parameters),
        )

        if attachments_zip_bytes:
            with ZipFile(BytesIO(attachments_zip_bytes), "r") as source_zip:
                for entry in source_zip.infolist():
                    name = entry.filename.replace("\\", "/")
                    if name.endswith("/"):
                        continue

                    base_name = name.split("/")[-1].lower()
                    if base_name in RESERVED_EXPORT_FILES:
                        continue

                    zip_file.writestr(name, source_zip.read(entry.filename))

        if include_report_json and report is not None:
            zip_file.writestr("validation-report.json", json.dumps(report.model_dump(), indent=2, ensure_ascii=False))

    return buffer.getvalue()


def build_report_xlsx(report: ValidationReport, executive_summary: ExecutiveSummary | None = None) -> bytes:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Resumo"

    summary_sheet.append(["Indicador", "Valor"])
    summary_sheet.append(["Apto para carga", "Sim" if report.summary.is_valid else "Nao"])
    summary_sheet.append(["Erros", report.summary.errors])
    summary_sheet.append(["Avisos", report.summary.warnings])
    summary_sheet.append(["Informativos", report.summary.infos])
    summary_sheet.append(["Total de inconsistencias", report.summary.total_issues])
    summary_sheet.append([])
    summary_sheet.append(["Arquivo", "Quantidade de linhas"])

    for key, value in report.record_counts.items():
        summary_sheet.append([key, value])

    if executive_summary is not None:
        executive_sheet = workbook.create_sheet("Executivo")
        executive_sheet.append(["Indicador", "Valor"])
        executive_sheet.append(["Total de contratos", executive_summary.total_contracts])
        executive_sheet.append(["Contratos prontos para importar", executive_summary.contracts_ready_for_import])
        executive_sheet.append(["Contratos com erros", executive_summary.contracts_with_errors])
        executive_sheet.append(["Contratos com avisos", executive_summary.contracts_with_warnings])
        executive_sheet.append(["Contratos com infos", executive_summary.contracts_with_infos])
        executive_sheet.append(["Anexos de contrato mapeados", executive_summary.mapped_contract_documents])
        executive_sheet.append(["Arquivos CLID mapeados", executive_summary.mapped_clid_documents])
        executive_sheet.append(["Membros de time", executive_summary.team_assignments])
        executive_sheet.append(["Percentual de prontidão (%)", executive_summary.readiness_percent])
        executive_sheet.append(["Recomendação", executive_summary.recommendation])

    issues_sheet = workbook.create_sheet("Inconsistencias")
    issues_sheet.append(["Severidade", "Codigo", "Mensagem", "Arquivo", "Linha", "Campo", "Contrato"])
    for issue in report.issues:
        issues_sheet.append(
            [
                issue.severity,
                issue.code,
                issue.message,
                issue.source_file or "",
                issue.row or "",
                issue.field or "",
                issue.contract_id or "",
            ]
        )

    contract_sheet = workbook.create_sheet("PorContrato")
    contract_sheet.append(["Contrato", "Erros", "Avisos", "Infos", "Total"])

    by_contract: dict[str, dict[str, int]] = {}
    for issue in report.issues:
        contract_id = issue.contract_id or "(sem contrato)"
        if contract_id not in by_contract:
            by_contract[contract_id] = {"error": 0, "warning": 0, "info": 0, "total": 0}
        by_contract[contract_id][issue.severity] += 1
        by_contract[contract_id]["total"] += 1

    for contract_id in sorted(by_contract.keys()):
        row = by_contract[contract_id]
        contract_sheet.append([contract_id, row["error"], row["warning"], row["info"], row["total"]])

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
