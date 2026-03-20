from __future__ import annotations

import re
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.models import (
    AribaDataset,
    EXPECTED_COLUMNS,
    ExecutiveSummary,
    ValidationIssue,
    ValidationReport,
    ValidationRules,
    ValidationSummary,
    default_validation_rules,
)

DATE_FORMAT = "%Y-%m-%d"
DOC_FOLDER = "Documentos contratos/"
CLID_FOLDER = "Documentos CLID/"
ALLOWED_CONTRACT_EXTENSIONS = {
    ".pdf",
    ".doc",
    ".docx",
    ".rtf",
    ".txt",
    ".eml",
    ".msg",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".csv",
    ".zip",
    ".png",
    ".jpg",
    ".jpeg",
}
ALLOWED_CLID_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".csv"}


def _date_is_valid(value: str) -> bool:
    if not value:
        return True
    try:
        datetime.strptime(value, DATE_FORMAT)
        return True
    except ValueError:
        return False


def _decimal_is_valid(value: str) -> bool:
    if not value:
        return True
    try:
        Decimal(value)
        return True
    except (InvalidOperation, ValueError):
        return False


def _path_exists(available_paths: set[str] | None, relative_path: str) -> bool:
    if not available_paths:
        return True

    normalized = relative_path.replace("\\", "/").lstrip("./")
    if normalized in available_paths:
        return True

    suffix = f"/{normalized}"
    return any(path.endswith(suffix) for path in available_paths)


def _normalize_path(path: str) -> str:
    return path.replace("\\", "/").strip().lstrip("./")


def _extract_attachment_relative(path: str) -> str:
    normalized = _normalize_path(path)
    lowered = normalized.lower()

    contract_index = lowered.find(DOC_FOLDER.lower())
    if contract_index >= 0:
        return normalized[contract_index:]

    clid_index = lowered.find(CLID_FOLDER.lower())
    if clid_index >= 0:
        return normalized[clid_index:]

    return normalized


def _is_attachment_path(path: str) -> bool:
    lowered = _extract_attachment_relative(path).lower()
    return lowered.startswith(DOC_FOLDER.lower()) or lowered.startswith(CLID_FOLDER.lower())


def _build_summary(issues: list[ValidationIssue]) -> ValidationSummary:
    counter = Counter(issue.severity for issue in issues)
    errors = counter.get("error", 0)
    warnings = counter.get("warning", 0)
    infos = counter.get("info", 0)
    return ValidationSummary(
        errors=errors,
        warnings=warnings,
        infos=infos,
        total_issues=len(issues),
        is_valid=errors == 0,
    )


def _compile_regex(pattern: str, fallback: str, issues: list[ValidationIssue], field_name: str) -> re.Pattern[str]:
    try:
        return re.compile(pattern)
    except re.error:
        issues.append(
            ValidationIssue(
                severity="warning",
                code="INVALID_RULE_REGEX",
                message=f"Regex inválida em {field_name}. Usando padrão fallback.",
                source_file="validation_rules",
                field=field_name,
            )
        )
        return re.compile(fallback)


def _collect_attachment_issues(
    dataset: AribaDataset,
    available_paths: set[str] | None,
    *,
    include_missing: bool,
) -> tuple[list[ValidationIssue], dict[str, int]]:
    issues: list[ValidationIssue] = []

    referenced_files: list[tuple[str, str, str]] = []
    for row in dataset.contract_documents:
        contract_id = row.get("ContractId", "")
        file_path = row.get("File", "")
        if file_path:
            referenced_files.append((contract_id, _extract_attachment_relative(file_path), "ContractDocuments.csv"))

    for row in dataset.contract_content_documents:
        contract_id = row.get("Workspace", "")
        file_path = row.get("File", "")
        if file_path:
            referenced_files.append((contract_id, _extract_attachment_relative(file_path), "ContractContentDocuments.csv"))

    referenced_set = {relative_path for _, relative_path, _ in referenced_files if relative_path}
    available_relative_paths = {
        _extract_attachment_relative(path)
        for path in (available_paths or set())
        if path and not path.endswith("/")
    }
    attachment_files_in_zip = {path for path in available_relative_paths if _is_attachment_path(path)}

    if include_missing and available_paths is not None:
        for contract_id, relative_path, source_file in referenced_files:
            if relative_path and not _path_exists(available_paths, relative_path):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        code="MISSING_ATTACHMENT",
                        message=f"Arquivo não encontrado no pacote: {relative_path}",
                        source_file=source_file,
                        field="File",
                        contract_id=contract_id or None,
                    )
                )

    filename_counter = Counter()
    for contract_id, relative_path, source_file in referenced_files:
        lowered_path = relative_path.lower()
        if source_file == "ContractDocuments.csv" and not lowered_path.startswith(DOC_FOLDER.lower()):
            issues.append(
                ValidationIssue(
                    severity="warning",
                    code="INVALID_ATTACHMENT_FOLDER",
                    message=f"Arquivo de contrato fora da pasta esperada ({DOC_FOLDER}): {relative_path}",
                    source_file=source_file,
                    field="File",
                    contract_id=contract_id or None,
                )
            )

        if source_file == "ContractContentDocuments.csv" and not lowered_path.startswith(CLID_FOLDER.lower()):
            issues.append(
                ValidationIssue(
                    severity="warning",
                    code="INVALID_ATTACHMENT_FOLDER",
                    message=f"Arquivo CLID fora da pasta esperada ({CLID_FOLDER}): {relative_path}",
                    source_file=source_file,
                    field="File",
                    contract_id=contract_id or None,
                )
            )

        file_name = Path(relative_path).name.lower()
        if file_name:
            filename_counter[file_name] += 1

        extension = Path(relative_path).suffix.lower()
        if source_file == "ContractDocuments.csv":
            if extension and extension not in ALLOWED_CONTRACT_EXTENSIONS:
                issues.append(
                    ValidationIssue(
                        severity="warning",
                        code="INVALID_ATTACHMENT_EXTENSION",
                        message=f"Extensão não usual para documento de contrato: {extension}",
                        source_file=source_file,
                        field="File",
                        contract_id=contract_id or None,
                    )
                )
        elif source_file == "ContractContentDocuments.csv":
            if extension and extension not in ALLOWED_CLID_EXTENSIONS:
                issues.append(
                    ValidationIssue(
                        severity="warning",
                        code="INVALID_ATTACHMENT_EXTENSION",
                        message=f"Extensão inválida para CLID (esperado Excel/CSV): {extension}",
                        source_file=source_file,
                        field="File",
                        contract_id=contract_id or None,
                    )
                )

    for file_name, count in filename_counter.items():
        if count > 1:
            issues.append(
                ValidationIssue(
                    severity="warning",
                    code="DUPLICATE_ATTACHMENT_FILENAME",
                    message=f"Mesmo nome de arquivo aparece em mais de uma referência: {file_name}",
                    source_file="attachments",
                )
            )

    extra_files = sorted(path for path in attachment_files_in_zip if path not in referenced_set)
    for extra_path in extra_files[:50]:
        issues.append(
            ValidationIssue(
                severity="info",
                code="UNREFERENCED_ATTACHMENT",
                message=f"Arquivo no ZIP sem referência nos CSVs: {extra_path}",
                source_file="attachments",
                field="File",
            )
        )
    if len(extra_files) > 50:
        issues.append(
            ValidationIssue(
                severity="info",
                code="UNREFERENCED_ATTACHMENT_TRUNCATED",
                message=f"Foram encontrados {len(extra_files)} arquivos extras. Exibindo apenas os 50 primeiros.",
                source_file="attachments",
            )
        )

    stats = {
        "referenced_files": len(referenced_set),
        "files_in_zip": len(attachment_files_in_zip),
        "missing_files": len(
            [
                1
                for _, relative_path, _ in referenced_files
                if include_missing and available_paths is not None and not _path_exists(available_paths, relative_path)
            ]
        ),
        "extra_files": len(extra_files),
    }
    return issues, stats


def validate_attachment_bundle(
    dataset: AribaDataset,
    available_paths: set[str] | None,
) -> tuple[ValidationSummary, list[ValidationIssue], dict[str, int]]:
    issues, stats = _collect_attachment_issues(dataset, available_paths, include_missing=True)
    summary = _build_summary(issues)
    return summary, issues, stats


def validate_dataset(
    dataset: AribaDataset,
    headers_by_key: dict[str, list[str]] | None = None,
    missing_files: list[str] | None = None,
    available_paths: set[str] | None = None,
    rules: ValidationRules | None = None,
) -> ValidationReport:
    headers_by_key = headers_by_key or {}
    missing_files = missing_files or []
    rules = rules or default_validation_rules()

    issues: list[ValidationIssue] = []

    contract_id_pattern = _compile_regex(
        pattern=rules.contract_id_regex,
        fallback=r"^LCW\d+$",
        issues=issues,
        field_name="contract_id_regex",
    )
    sap_party_pattern = _compile_regex(
        pattern=rules.sap_party_regex,
        fallback=r"^sap:\d+$",
        issues=issues,
        field_name="sap_party_regex",
    )

    allowed_hierarchical_types = set(rules.allowed_hierarchical_types)
    allowed_contract_statuses = set(rules.allowed_contract_statuses)
    allowed_base_languages = set(rules.allowed_base_languages)

    def add_issue(
        severity: str,
        code: str,
        message: str,
        source_file: str | None = None,
        row: int | None = None,
        field: str | None = None,
        contract_id: str | None = None,
    ) -> None:
        issues.append(
            ValidationIssue(
                severity=severity,  # type: ignore[arg-type]
                code=code,
                message=message,
                source_file=source_file,
                row=row,
                field=field,
                contract_id=contract_id,
            )
        )

    for missing_file in missing_files:
        add_issue("error", "MISSING_FILE", f"Arquivo obrigatório ausente: {missing_file}", source_file=missing_file)

    file_map = {
        "contracts": "Contracts.csv",
        "contract_documents": "ContractDocuments.csv",
        "contract_content_documents": "ContractContentDocuments.csv",
        "contract_teams": "ContractTeams.csv",
        "import_projects_parameters": "ImportProjectsParameters.csv",
    }

    for key, expected_columns in EXPECTED_COLUMNS.items():
        actual_headers = headers_by_key.get(key, [])
        source_file = file_map.get(key)
        if actual_headers:
            missing_columns = [column for column in expected_columns if column not in actual_headers]
            unexpected_columns = [column for column in actual_headers if column not in expected_columns]

            if missing_columns:
                add_issue(
                    "error",
                    "MISSING_COLUMN",
                    f"Colunas ausentes em {source_file}: {', '.join(missing_columns)}",
                    source_file=source_file,
                )

            if unexpected_columns:
                add_issue(
                    "warning",
                    "UNEXPECTED_COLUMN",
                    f"Colunas não esperadas em {source_file}: {', '.join(unexpected_columns)}",
                    source_file=source_file,
                )

    contracts_by_id: dict[str, dict[str, str]] = {}
    contract_ids: list[str] = []

    required_contract_fields = rules.required_contract_fields
    for index, row in enumerate(dataset.contracts, start=2):
        contract_id = row.get("ContractId", "")
        contract_ids.append(contract_id)

        for field in required_contract_fields:
            if not row.get(field, ""):
                add_issue(
                    "error",
                    "REQUIRED_FIELD",
                    f"Campo obrigatório vazio: {field}",
                    source_file="Contracts.csv",
                    row=index,
                    field=field,
                    contract_id=contract_id or None,
                )

        if contract_id and not contract_id_pattern.match(contract_id):
            add_issue(
                "error",
                "INVALID_ID_FORMAT",
                "ContractId inválido. Formato esperado pelas regras do cliente.",
                source_file="Contracts.csv",
                row=index,
                field="ContractId",
                contract_id=contract_id,
            )

        hierarchical_type = row.get("HierarchicalType", "")
        if hierarchical_type and allowed_hierarchical_types and hierarchical_type not in allowed_hierarchical_types:
            add_issue(
                "error",
                "INVALID_HIERARCHICAL_TYPE",
                f"HierarchicalType inválido. Valores permitidos: {', '.join(sorted(allowed_hierarchical_types))}",
                source_file="Contracts.csv",
                row=index,
                field="HierarchicalType",
                contract_id=contract_id or None,
            )

        contract_status = row.get("ContractStatus", "")
        if contract_status and allowed_contract_statuses and contract_status not in allowed_contract_statuses:
            add_issue(
                "error",
                "INVALID_CONTRACT_STATUS",
                f"ContractStatus fora das regras permitidas: {', '.join(sorted(allowed_contract_statuses))}",
                source_file="Contracts.csv",
                row=index,
                field="ContractStatus",
                contract_id=contract_id or None,
            )

        base_language = row.get("BaseLanguage", "")
        if base_language and allowed_base_languages and base_language not in allowed_base_languages:
            add_issue(
                "warning",
                "INVALID_BASE_LANGUAGE",
                f"BaseLanguage fora das regras permitidas: {', '.join(sorted(allowed_base_languages))}",
                source_file="Contracts.csv",
                row=index,
                field="BaseLanguage",
                contract_id=contract_id or None,
            )

        parent_agreement = row.get("ParentAgreement", "")
        if hierarchical_type == "SubAgreement" and not parent_agreement:
            add_issue(
                "error",
                "MISSING_PARENT",
                "SubAgreement precisa informar ParentAgreement.",
                source_file="Contracts.csv",
                row=index,
                field="ParentAgreement",
                contract_id=contract_id or None,
            )

        if parent_agreement and not contract_id_pattern.match(parent_agreement):
            add_issue(
                "warning",
                "INVALID_PARENT_FORMAT",
                "ParentAgreement fora do formato esperado nas regras do cliente.",
                source_file="Contracts.csv",
                row=index,
                field="ParentAgreement",
                contract_id=contract_id or None,
            )

        for date_field in ["AgreementDate", "EffectiveDate", "ExpirationDate"]:
            value = row.get(date_field, "")
            if value and not _date_is_valid(value):
                add_issue(
                    "error",
                    "INVALID_DATE",
                    f"Data inválida em {date_field}. Formato esperado: YYYY-MM-DD.",
                    source_file="Contracts.csv",
                    row=index,
                    field=date_field,
                    contract_id=contract_id or None,
                )

        for numeric_field in ["ProposedAmount", "Amount"]:
            value = row.get(numeric_field, "")
            if value and not _decimal_is_valid(value):
                add_issue(
                    "error",
                    "INVALID_NUMBER",
                    f"Número inválido em {numeric_field}.",
                    source_file="Contracts.csv",
                    row=index,
                    field=numeric_field,
                    contract_id=contract_id or None,
                )

        for party_field in ["Supplier", "AffectedParties"]:
            value = row.get(party_field, "")
            if value and not sap_party_pattern.match(value):
                add_issue(
                    "warning",
                    "INVALID_PARTY_FORMAT",
                    f"{party_field} fora do padrão esperado pelas regras do cliente.",
                    source_file="Contracts.csv",
                    row=index,
                    field=party_field,
                    contract_id=contract_id or None,
                )

        related_id = row.get("RelatedId", "")
        if rules.enforce_related_id_numeric and related_id and not related_id.isdigit():
            add_issue(
                "warning",
                "INVALID_RELATED_ID",
                "RelatedId deve conter apenas números.",
                source_file="Contracts.csv",
                row=index,
                field="RelatedId",
                contract_id=contract_id or None,
            )

        if contract_id and contract_id not in contracts_by_id:
            contracts_by_id[contract_id] = row

    duplicates_contract = [key for key, count in Counter(contract_ids).items() if key and count > 1]
    for duplicate in duplicates_contract:
        add_issue(
            rules.duplicate_contract_severity,
            "DUPLICATE_CONTRACT_ID",
            f"ContractId duplicado encontrado: {duplicate}",
            source_file="Contracts.csv",
            contract_id=duplicate,
        )

    for index, row in enumerate(dataset.contracts, start=2):
        contract_id = row.get("ContractId", "")
        parent_agreement = row.get("ParentAgreement", "")
        if parent_agreement and parent_agreement not in contracts_by_id:
            add_issue(
                "error",
                "PARENT_NOT_FOUND",
                f"ParentAgreement não encontrado em Contracts.csv: {parent_agreement}",
                source_file="Contracts.csv",
                row=index,
                field="ParentAgreement",
                contract_id=contract_id or None,
            )
            continue

        if parent_agreement:
            parent_type = contracts_by_id.get(parent_agreement, {}).get("HierarchicalType", "")
            if parent_type and parent_type != "MasterAgreement":
                add_issue(
                    "warning",
                    "PARENT_NOT_MASTER",
                    f"ParentAgreement {parent_agreement} não está como MasterAgreement.",
                    source_file="Contracts.csv",
                    row=index,
                    field="ParentAgreement",
                    contract_id=contract_id or None,
                )

    contract_id_set = {contract_id for contract_id in contract_ids if contract_id}

    required_doc_fields = ["ContractId", "File", "Title", "Owner"]
    doc_keys: list[tuple[str, str]] = []
    for index, row in enumerate(dataset.contract_documents, start=2):
        contract_id = row.get("ContractId", "")
        file_path = row.get("File", "")
        doc_keys.append((contract_id, file_path))

        for field in required_doc_fields:
            if not row.get(field, ""):
                add_issue(
                    "error",
                    "REQUIRED_FIELD",
                    f"Campo obrigatório vazio: {field}",
                    source_file="ContractDocuments.csv",
                    row=index,
                    field=field,
                    contract_id=contract_id or None,
                )

        if contract_id and contract_id not in contract_id_set:
            add_issue(
                "error",
                "MISSING_CONTRACT_REFERENCE",
                "ContractId não encontrado em Contracts.csv.",
                source_file="ContractDocuments.csv",
                row=index,
                field="ContractId",
                contract_id=contract_id,
            )

        expected_prefix = rules.expected_contract_documents_prefix
        if file_path and expected_prefix and not file_path.replace("\\", "/").startswith(expected_prefix):
            add_issue(
                "warning",
                "UNEXPECTED_FILE_PATH",
                f"Arquivo deveria estar na pasta {expected_prefix}.",
                source_file="ContractDocuments.csv",
                row=index,
                field="File",
                contract_id=contract_id or None,
            )

        if file_path and not _path_exists(available_paths, file_path):
            add_issue(
                "error",
                "MISSING_ATTACHMENT",
                f"Arquivo não encontrado no pacote: {file_path}",
                source_file="ContractDocuments.csv",
                row=index,
                field="File",
                contract_id=contract_id or None,
            )

    for (contract_id, file_path), count in Counter(doc_keys).items():
        if contract_id and file_path and count > 1:
            add_issue(
                rules.duplicate_document_severity,
                "DUPLICATE_DOCUMENT",
                f"Documento duplicado para o contrato: {contract_id} -> {file_path}",
                source_file="ContractDocuments.csv",
                contract_id=contract_id,
            )

    required_content_fields = ["Workspace", "File", "title"]
    content_keys: list[tuple[str, str]] = []
    for index, row in enumerate(dataset.contract_content_documents, start=2):
        workspace = row.get("Workspace", "")
        file_path = row.get("File", "")
        content_keys.append((workspace, file_path))

        for field in required_content_fields:
            if not row.get(field, ""):
                add_issue(
                    "error",
                    "REQUIRED_FIELD",
                    f"Campo obrigatório vazio: {field}",
                    source_file="ContractContentDocuments.csv",
                    row=index,
                    field=field,
                    contract_id=workspace or None,
                )

        if workspace and workspace not in contract_id_set:
            add_issue(
                "error",
                "MISSING_CONTRACT_REFERENCE",
                "Workspace não encontrado em Contracts.csv.",
                source_file="ContractContentDocuments.csv",
                row=index,
                field="Workspace",
                contract_id=workspace,
            )

        expected_prefix = rules.expected_clid_documents_prefix
        if file_path and expected_prefix and not file_path.replace("\\", "/").startswith(expected_prefix):
            add_issue(
                "warning",
                "UNEXPECTED_FILE_PATH",
                f"Arquivo deveria estar na pasta {expected_prefix}.",
                source_file="ContractContentDocuments.csv",
                row=index,
                field="File",
                contract_id=workspace or None,
            )

        if file_path and not _path_exists(available_paths, file_path):
            add_issue(
                "error",
                "MISSING_ATTACHMENT",
                f"Arquivo não encontrado no pacote: {file_path}",
                source_file="ContractContentDocuments.csv",
                row=index,
                field="File",
                contract_id=workspace or None,
            )

    for (workspace, file_path), count in Counter(content_keys).items():
        if workspace and file_path and count > 1:
            add_issue(
                rules.duplicate_content_document_severity,
                "DUPLICATE_CONTENT_DOCUMENT",
                f"CLID duplicado para o contrato: {workspace} -> {file_path}",
                source_file="ContractContentDocuments.csv",
                contract_id=workspace,
            )

    required_team_fields = rules.required_team_fields
    team_keys: list[tuple[str, str, str]] = []
    teams_by_contract: dict[str, int] = defaultdict(int)
    group_by_contract: dict[str, set[str]] = defaultdict(set)

    for index, row in enumerate(dataset.contract_teams, start=2):
        workspace = row.get("Workspace", "")
        project_group = row.get("ProjectGroup", "")
        member = row.get("Member", "")
        team_keys.append((workspace, project_group, member))

        for field in required_team_fields:
            if not row.get(field, ""):
                add_issue(
                    "error",
                    "REQUIRED_FIELD",
                    f"Campo obrigatório vazio: {field}",
                    source_file="ContractTeams.csv",
                    row=index,
                    field=field,
                    contract_id=workspace or None,
                )

        if workspace and workspace not in contract_id_set:
            add_issue(
                "error",
                "MISSING_CONTRACT_REFERENCE",
                "Workspace não encontrado em Contracts.csv.",
                source_file="ContractTeams.csv",
                row=index,
                field="Workspace",
                contract_id=workspace,
            )

        if workspace:
            teams_by_contract[workspace] += 1
            if project_group:
                group_by_contract[workspace].add(project_group)

    for (workspace, project_group, member), count in Counter(team_keys).items():
        if workspace and project_group and member and count > 1:
            add_issue(
                "info",
                "DUPLICATE_TEAM_MEMBER",
                f"Membro repetido no mesmo grupo: {workspace} / {project_group} / {member}",
                source_file="ContractTeams.csv",
                contract_id=workspace,
            )

    for contract_id in sorted(contract_id_set):
        if teams_by_contract.get(contract_id, 0) == 0:
            add_issue(
                rules.missing_team_severity,
                "MISSING_TEAM",
                f"Contrato sem time em ContractTeams.csv: {contract_id}",
                source_file="ContractTeams.csv",
                contract_id=contract_id,
            )

        required_groups = rules.required_team_project_groups
        if required_groups:
            existing_groups = group_by_contract.get(contract_id, set())
            for required_group in required_groups:
                if required_group not in existing_groups:
                    add_issue(
                        "warning",
                        "MISSING_REQUIRED_GROUP",
                        f"Contrato sem grupo obrigatório '{required_group}' em ContractTeams.csv.",
                        source_file="ContractTeams.csv",
                        contract_id=contract_id,
                    )

    params_rows = dataset.import_projects_parameters
    if len(params_rows) != 1:
        add_issue(
            "error",
            "IMPORT_PARAMS_ROW_COUNT",
            "ImportProjectsParameters.csv deve conter exatamente 1 linha.",
            source_file="ImportProjectsParameters.csv",
        )

    if params_rows:
        params = params_rows[0]
        required_params_fields = [
            "WorkspaceLookupKey",
            "TemplateName",
            "AttributesFileLocation",
            "DocumentsFileLocation",
            "TeamsFileLocation",
            "ContractContentDocumentsFileLocation",
            "FolderFieldName",
            "FolderFieldPattern",
            "FolderFormat",
        ]
        for field in required_params_fields:
            if not params.get(field, ""):
                add_issue(
                    "error",
                    "REQUIRED_FIELD",
                    f"Campo obrigatório vazio: {field}",
                    source_file="ImportProjectsParameters.csv",
                    row=2,
                    field=field,
                )

        expected_file_locations = {
            "AttributesFileLocation": "Contracts.csv",
            "DocumentsFileLocation": "ContractDocuments.csv",
            "TeamsFileLocation": "ContractTeams.csv",
            "ContractContentDocumentsFileLocation": "ContractContentDocuments.csv",
        }
        for field, expected_value in expected_file_locations.items():
            actual_value = params.get(field, "")
            if actual_value and actual_value != expected_value:
                add_issue(
                    "warning",
                    "UNEXPECTED_IMPORT_PARAM",
                    f"{field} diferente do padrão ({expected_value}). Valor atual: {actual_value}",
                    source_file="ImportProjectsParameters.csv",
                    row=2,
                    field=field,
                )

    if len(dataset.contract_documents) == 0:
        add_issue(
            "info",
            "EMPTY_DOCUMENTS",
            "ContractDocuments.csv está vazio.",
            source_file="ContractDocuments.csv",
        )

    if len(dataset.contract_content_documents) == 0:
        add_issue(
            "info",
            "EMPTY_CLID_DOCUMENTS",
            "ContractContentDocuments.csv está vazio.",
            source_file="ContractContentDocuments.csv",
        )

    if len(dataset.contracts) == 0:
        add_issue(
            "error",
            "EMPTY_CONTRACTS",
            "Contracts.csv está vazio.",
            source_file="Contracts.csv",
        )

    if available_paths is not None:
        attachment_issues, _ = _collect_attachment_issues(dataset, available_paths, include_missing=False)
        issues.extend(attachment_issues)

    summary = _build_summary(issues)
    record_counts = {
        "contracts": len(dataset.contracts),
        "contract_documents": len(dataset.contract_documents),
        "contract_content_documents": len(dataset.contract_content_documents),
        "contract_teams": len(dataset.contract_teams),
        "import_projects_parameters": len(dataset.import_projects_parameters),
    }

    return ValidationReport(summary=summary, issues=issues, record_counts=record_counts)


def build_executive_summary(dataset: AribaDataset, report: ValidationReport) -> ExecutiveSummary:
    unique_contract_ids: list[str] = []
    seen_contracts: set[str] = set()
    for row in dataset.contracts:
        contract_id = row.get("ContractId", "").strip()
        if contract_id and contract_id not in seen_contracts:
            seen_contracts.add(contract_id)
            unique_contract_ids.append(contract_id)

    issues_by_contract: dict[str, Counter[str]] = defaultdict(Counter)
    for issue in report.issues:
        if issue.contract_id:
            issues_by_contract[issue.contract_id][issue.severity] += 1

    contracts_with_errors = 0
    contracts_with_warnings = 0
    contracts_with_infos = 0
    contracts_ready_for_import = 0

    for contract_id in unique_contract_ids:
        counts = issues_by_contract.get(contract_id, Counter())
        has_error = counts.get("error", 0) > 0
        has_warning = counts.get("warning", 0) > 0
        has_info = counts.get("info", 0) > 0

        if has_error:
            contracts_with_errors += 1
        if has_warning:
            contracts_with_warnings += 1
        if has_info:
            contracts_with_infos += 1
        if not has_error:
            contracts_ready_for_import += 1

    total_contracts = len(unique_contract_ids)
    readiness_percent = round((contracts_ready_for_import / total_contracts) * 100, 2) if total_contracts else 0.0

    recommendation = "Pacote pronto para importação."
    if total_contracts == 0:
        recommendation = "Inclua ao menos 1 contrato antes de gerar o pacote."
    elif report.summary.errors > 0:
        recommendation = "Corrija os erros críticos antes de exportar para o Ariba."
    elif report.summary.warnings > 0:
        recommendation = "Revise os avisos para reduzir risco de rejeição na importação."

    return ExecutiveSummary(
        total_contracts=total_contracts,
        contracts_ready_for_import=contracts_ready_for_import,
        contracts_with_errors=contracts_with_errors,
        contracts_with_warnings=contracts_with_warnings,
        contracts_with_infos=contracts_with_infos,
        mapped_contract_documents=len(dataset.contract_documents),
        mapped_clid_documents=len(dataset.contract_content_documents),
        team_assignments=len(dataset.contract_teams),
        readiness_percent=readiness_percent,
        recommendation=recommendation,
    )

