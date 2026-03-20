from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


Severity = Literal["error", "warning", "info"]

REQUIRED_FILES = {
    "Contracts.csv": "contracts",
    "ContractDocuments.csv": "contract_documents",
    "ContractContentDocuments.csv": "contract_content_documents",
    "ContractTeams.csv": "contract_teams",
    "ImportProjectsParameters.csv": "import_projects_parameters",
}

EXPECTED_COLUMNS = {
    "contracts": [
        "Owner",
        "Title",
        "ContractId",
        "BaseLanguage",
        "Description",
        "Supplier",
        "AffectedParties",
        "HierarchicalType",
        "ParentAgreement",
        "ProposedAmount",
        "Amount",
        "Commodity",
        "Region",
        "Client",
        "ExpirationTermType",
        "AutoRenewalInterval",
        "MaxAutoRenewalsAllowed",
        "AgreementDate",
        "EffectiveDate",
        "ExpirationDate",
        "NoticePeriod",
        "NoticeEmailRecipients",
        "ContractStatus",
        "RelatedId",
    ],
    "contract_documents": ["ContractId", "File", "Title", "Folder", "Owner", "Status"],
    "contract_content_documents": ["Workspace", "File", "title"],
    "contract_teams": ["Workspace", "ProjectGroup", "Member"],
    "import_projects_parameters": [
        "WorkspaceLookupKey",
        "TemplateName",
        "AttributesFileLocation",
        "DocumentsFileLocation",
        "TeamsFileLocation",
        "ContractContentDocumentsFileLocation",
        "RootParentId",
        "TopFolderName",
        "FolderFieldName",
        "FolderFieldPattern",
        "FolderFormat",
    ],
}

UNIFIED_COLUMNS = [
    "ContractId",
    "Title",
    "Owner",
    "BaseLanguage",
    "Description",
    "Supplier",
    "AffectedParties",
    "HierarchicalType",
    "ParentAgreement",
    "ProposedAmount",
    "Amount",
    "Commodity",
    "Region",
    "Client",
    "AgreementDate",
    "EffectiveDate",
    "ExpirationDate",
    "ContractStatus",
    "RelatedId",
    "TeamProjectGroup",
    "TeamMember",
    "DocumentFile",
    "DocumentTitle",
    "DocumentFolder",
    "DocumentOwner",
    "DocumentStatus",
    "ClidFile",
    "ClidTitle",
]

DEFAULT_IMPORT_PARAMETERS = {
    "WorkspaceLookupKey": "ContractId",
    "TemplateName": "Contrato - Legado",
    "AttributesFileLocation": "Contracts.csv",
    "DocumentsFileLocation": "ContractDocuments.csv",
    "TeamsFileLocation": "ContractTeams.csv",
    "ContractContentDocumentsFileLocation": "ContractContentDocuments.csv",
    "RootParentId": "",
    "TopFolderName": "",
    "FolderFieldName": "ContractId",
    "FolderFieldPattern": "([1-9][0-9]*)",
    "FolderFormat": "{0} to {1}",
}


class ValidationRules(BaseModel):
    allowed_hierarchical_types: list[str] = Field(
        default_factory=lambda: ["MasterAgreement", "SubAgreement"]
    )
    allowed_contract_statuses: list[str] = Field(default_factory=lambda: ["Published"])
    allowed_base_languages: list[str] = Field(default_factory=lambda: ["BrazilianPortuguese"])

    required_contract_fields: list[str] = Field(
        default_factory=lambda: [
            "Owner",
            "Title",
            "ContractId",
            "BaseLanguage",
            "HierarchicalType",
            "ContractStatus",
        ]
    )
    required_team_fields: list[str] = Field(default_factory=lambda: ["Workspace", "ProjectGroup", "Member"])
    required_team_project_groups: list[str] = Field(default_factory=list)

    contract_id_regex: str = r"^LCW\d+$"
    sap_party_regex: str = r"^sap:\d+$"

    enforce_related_id_numeric: bool = True
    expected_contract_documents_prefix: str = "Documentos contratos/"
    expected_clid_documents_prefix: str = "Documentos CLID/"

    missing_team_severity: Severity = "info"
    duplicate_contract_severity: Severity = "warning"
    duplicate_document_severity: Severity = "warning"
    duplicate_content_document_severity: Severity = "warning"


class ValidationIssue(BaseModel):
    severity: Severity
    code: str
    message: str
    source_file: str | None = None
    row: int | None = None
    field: str | None = None
    contract_id: str | None = None


class ValidationSummary(BaseModel):
    errors: int
    warnings: int
    infos: int
    total_issues: int
    is_valid: bool


class ValidationReport(BaseModel):
    summary: ValidationSummary
    issues: list[ValidationIssue]
    record_counts: dict[str, int] = Field(default_factory=dict)


class AribaDataset(BaseModel):
    contracts: list[dict[str, str]] = Field(default_factory=list)
    contract_documents: list[dict[str, str]] = Field(default_factory=list)
    contract_content_documents: list[dict[str, str]] = Field(default_factory=list)
    contract_teams: list[dict[str, str]] = Field(default_factory=list)
    import_projects_parameters: list[dict[str, str]] = Field(default_factory=list)


class ExecutiveSummary(BaseModel):
    total_contracts: int = 0
    contracts_ready_for_import: int = 0
    contracts_with_errors: int = 0
    contracts_with_warnings: int = 0
    contracts_with_infos: int = 0
    mapped_contract_documents: int = 0
    mapped_clid_documents: int = 0
    team_assignments: int = 0
    readiness_percent: float = 0.0
    recommendation: str = ""


class AttachmentValidationResponse(BaseModel):
    summary: ValidationSummary
    issues: list[ValidationIssue] = Field(default_factory=list)
    stats: dict[str, int] = Field(default_factory=dict)


class AnalysisResponse(BaseModel):
    dataset: AribaDataset
    report: ValidationReport
    applied_rules: ValidationRules
    executive_summary: ExecutiveSummary
    run_id: str | None = None


class ClientProfile(BaseModel):
    name: str
    validation_rules: ValidationRules
    import_parameters: dict[str, str] = Field(default_factory=dict)
    updated_at: str


class ClientProfilesResponse(BaseModel):
    profiles: list[ClientProfile] = Field(default_factory=list)


class SaveClientProfileRequest(BaseModel):
    name: str
    validation_rules: ValidationRules
    import_parameters: dict[str, str] = Field(default_factory=dict)


class UnifiedManualRequest(BaseModel):
    rows: list[dict[str, Any]] = Field(default_factory=list)
    import_parameters_override: dict[str, str] | None = None
    validation_rules: ValidationRules | None = None
    profile_name: str | None = None


class AnalyzeJsonRequest(BaseModel):
    dataset: AribaDataset
    validation_rules: ValidationRules | None = None
    profile_name: str | None = None


class ExportRequest(BaseModel):
    dataset: AribaDataset
    include_report_json: bool = True
    report: ValidationReport | None = None
    run_id: str | None = None


class ReportExportRequest(BaseModel):
    report: ValidationReport
    executive_summary: ExecutiveSummary | None = None
    run_id: str | None = None


class RunArtifacts(BaseModel):
    package_zip: bool = False
    package_with_attachments_zip: bool = False
    report_xlsx: bool = False


class ExecutionRun(BaseModel):
    run_id: str
    created_at: str
    source: Literal["zip", "unified", "manual", "json", "unknown"]
    profile_name: str | None = None
    summary: ValidationSummary
    record_counts: dict[str, int] = Field(default_factory=dict)
    readiness_percent: float = 0.0
    contracts_total: int = 0
    artifacts: RunArtifacts = Field(default_factory=RunArtifacts)


class ExecutionRunsResponse(BaseModel):
    runs: list[ExecutionRun] = Field(default_factory=list)


def default_validation_rules() -> ValidationRules:
    return ValidationRules()
