from __future__ import annotations

from pathlib import Path

from app.models import AribaDataset, DEFAULT_IMPORT_PARAMETERS, EXPECTED_COLUMNS


def _value(row: dict[str, object], *keys: str) -> str:
    for key in keys:
        if key in row:
            raw = row.get(key)
            if raw is None:
                continue
            value = str(raw).strip()
            if value:
                return value
    return ""


def _empty_contract() -> dict[str, str]:
    return {column: "" for column in EXPECTED_COLUMNS["contracts"]}


def _normalize_path(value: str) -> str:
    return value.replace("\\", "/").strip()


def _basename_title(file_path: str) -> str:
    if not file_path:
        return ""
    return Path(file_path).stem


def map_unified_rows_to_dataset(
    rows: list[dict[str, object]],
    import_parameters_override: dict[str, str] | None = None,
) -> AribaDataset:
    contracts_by_id: dict[str, dict[str, str]] = {}
    contracts_order: list[str] = []
    contract_documents: list[dict[str, str]] = []
    content_documents: list[dict[str, str]] = []
    contract_teams: list[dict[str, str]] = []

    for raw_row in rows:
        contract_id = _value(raw_row, "ContractId", "contract_id", "workspace")
        if not contract_id:
            continue

        if contract_id not in contracts_by_id:
            contracts_by_id[contract_id] = _empty_contract()
            contracts_order.append(contract_id)

        contract_row = contracts_by_id[contract_id]
        contract_field_map = {
            "Owner": _value(raw_row, "Owner", "owner"),
            "Title": _value(raw_row, "Title", "title"),
            "ContractId": contract_id,
            "BaseLanguage": _value(raw_row, "BaseLanguage", "base_language") or "BrazilianPortuguese",
            "Description": _value(raw_row, "Description", "description"),
            "Supplier": _value(raw_row, "Supplier", "supplier"),
            "AffectedParties": _value(raw_row, "AffectedParties", "affected_parties"),
            "HierarchicalType": _value(raw_row, "HierarchicalType", "hierarchical_type") or "MasterAgreement",
            "ParentAgreement": _value(raw_row, "ParentAgreement", "parent_agreement"),
            "ProposedAmount": _value(raw_row, "ProposedAmount", "proposed_amount"),
            "Amount": _value(raw_row, "Amount", "amount"),
            "Commodity": _value(raw_row, "Commodity", "commodity"),
            "Region": _value(raw_row, "Region", "region") or "BRA",
            "Client": _value(raw_row, "Client", "client"),
            "ExpirationTermType": _value(raw_row, "ExpirationTermType", "expiration_term_type"),
            "AutoRenewalInterval": _value(raw_row, "AutoRenewalInterval", "auto_renewal_interval"),
            "MaxAutoRenewalsAllowed": _value(raw_row, "MaxAutoRenewalsAllowed", "max_auto_renewals_allowed"),
            "AgreementDate": _value(raw_row, "AgreementDate", "agreement_date"),
            "EffectiveDate": _value(raw_row, "EffectiveDate", "effective_date"),
            "ExpirationDate": _value(raw_row, "ExpirationDate", "expiration_date"),
            "NoticePeriod": _value(raw_row, "NoticePeriod", "notice_period"),
            "NoticeEmailRecipients": _value(raw_row, "NoticeEmailRecipients", "notice_email_recipients"),
            "ContractStatus": _value(raw_row, "ContractStatus", "contract_status") or "Published",
            "RelatedId": _value(raw_row, "RelatedId", "related_id"),
        }

        for field, value in contract_field_map.items():
            if value and not contract_row.get(field, ""):
                contract_row[field] = value

        document_file = _value(raw_row, "DocumentFile", "document_file", "ContractDocumentFile")
        if document_file:
            file_path = _normalize_path(document_file)
            if not file_path.startswith("Documentos contratos/"):
                file_path = f"Documentos contratos/{file_path.lstrip('/')}"

            contract_documents.append(
                {
                    "ContractId": contract_id,
                    "File": file_path,
                    "Title": _value(raw_row, "DocumentTitle", "document_title") or _basename_title(file_path),
                    "Folder": _value(raw_row, "DocumentFolder", "document_folder"),
                    "Owner": _value(raw_row, "DocumentOwner", "document_owner") or contract_row.get("Owner", ""),
                    "Status": _value(raw_row, "DocumentStatus", "document_status"),
                }
            )

        clid_file = _value(raw_row, "ClidFile", "clid_file", "ContractContentFile")
        if clid_file:
            file_path = _normalize_path(clid_file)
            if not file_path.startswith("Documentos CLID/"):
                file_path = f"Documentos CLID/{file_path.lstrip('/')}"

            content_documents.append(
                {
                    "Workspace": contract_id,
                    "File": file_path,
                    "title": _value(raw_row, "ClidTitle", "clid_title") or Path(file_path).name,
                }
            )

        team_group = _value(raw_row, "TeamProjectGroup", "team_project_group", "ProjectGroup")
        team_member = _value(raw_row, "TeamMember", "team_member", "Member")
        if team_group and team_member:
            contract_teams.append(
                {
                    "Workspace": contract_id,
                    "ProjectGroup": team_group,
                    "Member": team_member,
                }
            )

    contracts = [contracts_by_id[contract_id] for contract_id in contracts_order]

    import_parameters = DEFAULT_IMPORT_PARAMETERS.copy()
    if import_parameters_override:
        for key, value in import_parameters_override.items():
            if key in import_parameters and value is not None:
                import_parameters[key] = str(value).strip()

    return AribaDataset(
        contracts=contracts,
        contract_documents=contract_documents,
        contract_content_documents=content_documents,
        contract_teams=contract_teams,
        import_projects_parameters=[import_parameters],
    )
