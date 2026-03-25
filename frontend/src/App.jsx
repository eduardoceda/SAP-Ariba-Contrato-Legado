import { useEffect, useMemo, useRef, useState } from "react";
import stratesysLogo from "./assets/stratesys_logo_small.png";

const API_PREFIX = import.meta.env.VITE_API_BASE_URL || "/api";

const MANUAL_COLUMNS = [
  "ContractId",
  "Title",
  "Owner",
  "BaseLanguage",
  "HierarchicalType",
  "ParentAgreement",
  "AgreementDate",
  "EffectiveDate",
  "ExpirationDate",
  "ContractStatus",
  "TeamProjectGroup",
  "TeamMember",
  "DocumentFile",
  "DocumentTitle",
  "ClidFile",
  "ClidTitle",
];

const EXPECTED_COLUMNS = {
  contracts: [
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
  contract_documents: ["ContractId", "File", "Title", "Folder", "Owner", "Status"],
  contract_content_documents: ["Workspace", "File", "title"],
  contract_teams: ["Workspace", "ProjectGroup", "Member"],
  import_projects_parameters: [
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
};

const DEFAULT_RULES = {
  allowed_hierarchical_types: ["MasterAgreement", "SubAgreement"],
  allowed_contract_statuses: ["Published"],
  allowed_base_languages: ["BrazilianPortuguese"],
  required_contract_fields: [
    "Owner",
    "Title",
    "ContractId",
    "BaseLanguage",
    "HierarchicalType",
    "ContractStatus",
  ],
  required_team_fields: ["Workspace", "ProjectGroup", "Member"],
  required_team_project_groups: [],
  contract_id_regex: "^LCW\\d+$",
  sap_party_regex: "^sap:\\d+$",
  enforce_related_id_numeric: true,
  expected_contract_documents_prefix: "Documentos contratos/",
  expected_clid_documents_prefix: "Documentos CLID/",
  missing_team_severity: "info",
  duplicate_contract_severity: "warning",
  duplicate_document_severity: "warning",
  duplicate_content_document_severity: "warning",
};

const DEFAULT_IMPORT_PARAMETERS = {
  WorkspaceLookupKey: "ContractId",
  TemplateName: "Contrato - Legado",
  AttributesFileLocation: "Contracts.csv",
  DocumentsFileLocation: "ContractDocuments.csv",
  TeamsFileLocation: "ContractTeams.csv",
  ContractContentDocumentsFileLocation: "ContractContentDocuments.csv",
  RootParentId: "",
  TopFolderName: "",
  FolderFieldName: "ContractId",
  FolderFieldPattern: "([1-9][0-9]*)",
  FolderFormat: "{0} to {1}",
};

const IMPORT_PARAMETERS_FIELDS = [
  { key: "WorkspaceLookupKey", label: "WorkspaceLookupKey" },
  { key: "TemplateName", label: "TemplateName" },
  { key: "AttributesFileLocation", label: "AttributesFileLocation" },
  { key: "DocumentsFileLocation", label: "DocumentsFileLocation" },
  { key: "TeamsFileLocation", label: "TeamsFileLocation" },
  {
    key: "ContractContentDocumentsFileLocation",
    label: "ContractContentDocumentsFileLocation",
  },
  { key: "RootParentId", label: "RootParentId" },
  { key: "TopFolderName", label: "TopFolderName" },
  { key: "FolderFieldName", label: "FolderFieldName" },
  { key: "FolderFieldPattern", label: "FolderFieldPattern" },
  { key: "FolderFormat", label: "FolderFormat" },
];

const FIELD_DICTIONARY = [
  {
    field: "ContractId",
    description: "Identificador único do contrato.",
    example: "LCW4700001278",
  },
  {
    field: "Supplier / AffectedParties",
    description: "Parceiro SAP no formato prefixo + número.",
    example: "sap:0000381965",
  },
  {
    field: "TeamProjectGroup",
    description: "Grupo de projeto existente no Ariba.",
    example: "Comprador",
  },
  {
    field: "DocumentFile",
    description: "Caminho do anexo do contrato no pacote ZIP.",
    example: "Documentos contratos/Contrato_4700001278.pdf",
  },
  {
    field: "ClidFile",
    description: "Caminho do arquivo de itens de linha (CLID).",
    example: "Documentos CLID/CLID_4700001278.xlsx",
  },
  {
    field: "TemplateName",
    description: "Template do Ariba usado na importação.",
    example: "Contrato - Legado",
  },
];

const ISSUE_GUIDANCE = {
  MISSING_FILE: "Inclua o arquivo CSV obrigatório no pacote base.",
  MISSING_COLUMN: "Ajuste o cabeçalho para conter todas as colunas esperadas.",
  REQUIRED_FIELD: "Preencha o campo obrigatório na linha indicada.",
  INVALID_ID_FORMAT: "Padronize o ContractId conforme a regra definida no cliente.",
  INVALID_DATE: "Use formato de data YYYY-MM-DD.",
  INVALID_NUMBER: "Use número com ponto decimal, sem texto.",
  INVALID_PARTY_FORMAT: "Use sap: + números (ex.: sap:0000381965).",
  MISSING_CONTRACT_REFERENCE: "Garanta que o contrato exista em Contracts.csv.",
  MISSING_ATTACHMENT: "Confirme se o arquivo está no ZIP e no caminho correto.",
  UNEXPECTED_FILE_PATH: "Ajuste o caminho para a pasta esperada.",
  INVALID_ATTACHMENT_FOLDER: "Mova o arquivo para a pasta padrão correspondente.",
  INVALID_ATTACHMENT_EXTENSION: "Use extensões adequadas para anexo e CLID.",
  DUPLICATE_CONTRACT_ID: "Remova duplicidade de ContractId.",
  DUPLICATE_DOCUMENT: "Remova linhas duplicadas de documentos.",
  DUPLICATE_CONTENT_DOCUMENT: "Remova linhas CLID duplicadas.",
  DUPLICATE_TEAM_MEMBER: "Mantenha uma linha por contrato/grupo/membro.",
  IMPORT_PARAMS_ROW_COUNT: "Mantenha exatamente 1 linha em ImportProjectsParameters.",
  UNEXPECTED_IMPORT_PARAM: "Confirme o valor com o template oficial do cliente.",
  MISSING_TEAM: "Inclua ao menos 1 membro em ContractTeams.csv para o contrato.",
  MISSING_REQUIRED_GROUP: "Adicione o grupo obrigatório configurado.",
  EMPTY_CONTRACTS: "Inclua ao menos um contrato antes de gerar pacote.",
  UNREFERENCED_ATTACHMENT: "Arquivo extra no ZIP. Pode ser removido.",
  INVALID_RULE_REGEX: "Corrija a expressão regular na configuração avançada.",
};

const AUTO_FIXABLE_CODES = new Set([
  "UNEXPECTED_FILE_PATH",
  "REQUIRED_FIELD",
  "INVALID_PARTY_FORMAT",
  "INVALID_ID_FORMAT",
  "DUPLICATE_DOCUMENT",
  "DUPLICATE_CONTENT_DOCUMENT",
  "DUPLICATE_TEAM_MEMBER",
  "IMPORT_PARAMS_ROW_COUNT",
  "DUPLICATE_ATTACHMENT_FILENAME",
  "UNREFERENCED_ATTACHMENT",
]);

const SESSION_STORAGE_KEY = "ariba_legacy_session_v2";

const MANUAL_ROW_DEFAULTS = Object.freeze({
  ContractId: "",
  Title: "",
  Owner: "",
  BaseLanguage: "BrazilianPortuguese",
  HierarchicalType: "MasterAgreement",
  ParentAgreement: "",
  AgreementDate: "",
  EffectiveDate: "",
  ExpirationDate: "",
  ContractStatus: "Published",
  TeamProjectGroup: "Comprador",
  TeamMember: "",
  DocumentFile: "",
  DocumentTitle: "",
  ClidFile: "",
  ClidTitle: "",
});

function createManualRow() {
  return { ...MANUAL_ROW_DEFAULTS };
}

function hasManualRowInput(row) {
  return Object.entries(MANUAL_ROW_DEFAULTS).some(
    ([field, defaultValue]) => String(row?.[field] ?? "").trim() !== String(defaultValue ?? "").trim()
  );
}

function isManualRowComplete(row) {
  return MANUAL_COLUMNS.every((field) => String(row?.[field] ?? "").trim() !== "");
}

function buildAttachmentUploadEntries(files, rootFolder) {
  return (files || []).map((file) => {
    const normalizedRelativePath = String(file.webkitRelativePath || "")
      .replace(/\\/g, "/")
      .replace(/^\/+/, "");
    const relativeParts = normalizedRelativePath.split("/").filter(Boolean);
    const relativeName =
      relativeParts.length > 1 ? relativeParts.slice(1).join("/") : file.name || relativeParts[0] || "arquivo";

    return {
      file,
      path: `${rootFolder}/${relativeName}`,
    };
  });
}

function parseCsvList(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function toCsvList(items) {
  return Array.isArray(items) ? items.join(", ") : "";
}

function uniqueNonEmptyValues(values) {
  return Array.from(
    new Set(
      (values || [])
        .map((value) => String(value || "").trim())
        .filter(Boolean)
    )
  );
}

function pickSingleCandidate(values) {
  const candidates = uniqueNonEmptyValues(values);
  return candidates.length === 1 ? candidates[0] : "";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  const rounded = value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1);
  return `${rounded} ${units[unitIndex]}`;
}

async function readResponseBlobWithProgress(response, onProgress) {
  const rawContentLength = response.headers.get("Content-Length");
  const parsedContentLength = rawContentLength ? Number.parseInt(rawContentLength, 10) : Number.NaN;
  const totalBytes = Number.isFinite(parsedContentLength) && parsedContentLength > 0 ? parsedContentLength : null;
  const reader = response.body?.getReader?.();

  if (!reader) {
    const blob = await response.blob();
    if (onProgress) {
      onProgress({ downloadedBytes: blob.size, totalBytes: blob.size || totalBytes });
    }
    return blob;
  }

  const chunks = [];
  let downloadedBytes = 0;
  if (onProgress) {
    onProgress({ downloadedBytes: 0, totalBytes });
  }

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    if (!value) {
      continue;
    }

    chunks.push(value);
    downloadedBytes += value.byteLength;
    if (onProgress) {
      onProgress({ downloadedBytes, totalBytes });
    }
  }

  return new Blob(chunks, {
    type: response.headers.get("Content-Type") || "application/octet-stream",
  });
}

function formatDateSuffix() {
  return new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
}

function formatSeverityLabel(severity) {
  if (severity === "error") {
    return "Erro";
  }
  if (severity === "warning") {
    return "Aviso";
  }
  return "Informativo";
}

function getIssueLineNumber(issue) {
  const raw = issue?.row;
  if (typeof raw === "number" && Number.isFinite(raw)) {
    return raw;
  }
  const normalized = String(raw ?? "").trim();
  if (!normalized) {
    return Number.POSITIVE_INFINITY;
  }
  const match = normalized.match(/-?\d+/);
  if (!match) {
    return Number.POSITIVE_INFINITY;
  }
  const parsed = Number.parseInt(match[0], 10);
  return Number.isFinite(parsed) ? parsed : Number.POSITIVE_INFINITY;
}

function normalizeUiError(message) {
  const normalized = String(message || "").trim();
  if (!normalized) {
    return "Falha inesperada.";
  }
  if (normalized.toLowerCase().includes("failed to fetch")) {
    return "Falha de conexão com a API. Verifique se o backend está ativo.";
  }
  return normalized;
}

function getIssueGuidance(issue) {
  return ISSUE_GUIDANCE[issue.code] || "Revise a linha/campo indicado e ajuste conforme o template.";
}

async function parseError(response) {
  try {
    const payload = await response.json();
    if (typeof payload?.detail === "string") {
      return payload.detail;
    }
    return "Erro na requisição";
  } catch {
    return `Erro HTTP ${response.status}`;
  }
}

function normalizePath(path) {
  return String(path || "").replaceAll("\\", "/").trim();
}

function ensurePrefix(path, prefix) {
  const normalized = normalizePath(path).replace(/^\/+/, "");
  if (!normalized) {
    return "";
  }
  if (normalized.startsWith(prefix)) {
    return normalized;
  }
  return `${prefix}${normalized}`;
}

function stemFromPath(path) {
  const normalized = normalizePath(path);
  if (!normalized) {
    return "";
  }
  const name = normalized.split("/").pop() || "";
  return name.replace(/\.[^.]+$/, "");
}

function normalizeSapParty(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  if (/^sap:\d+$/i.test(raw)) {
    const [, digits] = raw.split(":");
    return `sap:${digits}`;
  }
  if (/^\d+$/.test(raw)) {
    return `sap:${raw}`;
  }
  return raw;
}

function dedupeRows(rows) {
  const seen = new Set();
  const normalizedRows = [];
  let removed = 0;
  for (const row of rows || []) {
    const key = JSON.stringify(row || {});
    if (seen.has(key)) {
      removed += 1;
      continue;
    }
    seen.add(key);
    normalizedRows.push(row);
  }
  return { rows: normalizedRows, removed };
}

function buildSafeAutoFixDataset(sourceDataset) {
  const dataset = JSON.parse(JSON.stringify(sourceDataset || {}));
  dataset.contracts = Array.isArray(dataset.contracts) ? dataset.contracts : [];
  dataset.contract_documents = Array.isArray(dataset.contract_documents) ? dataset.contract_documents : [];
  dataset.contract_content_documents = Array.isArray(dataset.contract_content_documents)
    ? dataset.contract_content_documents
    : [];
  dataset.contract_teams = Array.isArray(dataset.contract_teams) ? dataset.contract_teams : [];
  dataset.import_projects_parameters = Array.isArray(dataset.import_projects_parameters)
    ? dataset.import_projects_parameters
    : [];

  const changes = [];
  const ownerByContract = new Map();

  dataset.contracts = dataset.contracts.map((row) => {
    const next = { ...row };
    const originalId = next.ContractId || "";
    next.ContractId = String(next.ContractId || "").trim();
    next.ParentAgreement = String(next.ParentAgreement || "").trim();
    next.Supplier = normalizeSapParty(next.Supplier);
    next.AffectedParties = normalizeSapParty(next.AffectedParties);
    if (next.ContractId && !next.Title) {
      next.Title = next.ContractId;
      changes.push(`Título preenchido com ContractId (${next.ContractId}).`);
    }
    if (originalId !== next.ContractId) {
      changes.push(`ContractId normalizado: ${originalId} -> ${next.ContractId}.`);
    }
    if (next.ContractId) {
      ownerByContract.set(next.ContractId, String(next.Owner || "").trim());
    }
    return next;
  });

  const dedupedContractDocuments = dedupeRows(
    dataset.contract_documents.map((row) => {
      const next = { ...row };
      const originalPath = next.File || "";
      next.ContractId = String(next.ContractId || "").trim();
      next.File = ensurePrefix(next.File, "Documentos contratos/");
      if (!next.Title) {
        next.Title = stemFromPath(next.File);
      }
      if (!next.Owner && next.ContractId) {
        next.Owner = ownerByContract.get(next.ContractId) || "";
      }
      if (originalPath !== next.File) {
        changes.push(`Caminho de anexo normalizado: ${next.File}`);
      }
      return next;
    })
  );
  dataset.contract_documents = dedupedContractDocuments.rows;
  if (dedupedContractDocuments.removed > 0) {
    changes.push(`${dedupedContractDocuments.removed} linha(s) duplicada(s) removida(s) em ContractDocuments.`);
  }

  const dedupedContentDocuments = dedupeRows(
    dataset.contract_content_documents.map((row) => {
      const next = { ...row };
      const originalPath = next.File || "";
      next.Workspace = String(next.Workspace || "").trim();
      next.File = ensurePrefix(next.File, "Documentos CLID/");
      if (!next.title) {
        next.title = normalizePath(next.File).split("/").pop() || "";
      }
      if (originalPath !== next.File) {
        changes.push(`Caminho CLID normalizado: ${next.File}`);
      }
      return next;
    })
  );
  dataset.contract_content_documents = dedupedContentDocuments.rows;
  if (dedupedContentDocuments.removed > 0) {
    changes.push(`${dedupedContentDocuments.removed} linha(s) duplicada(s) removida(s) em ContractContentDocuments.`);
  }

  const dedupedTeams = dedupeRows(
    dataset.contract_teams.map((row) => ({
      ...row,
      Workspace: String(row.Workspace || "").trim(),
      ProjectGroup: String(row.ProjectGroup || "").trim(),
      Member: String(row.Member || "").trim(),
    }))
  );
  dataset.contract_teams = dedupedTeams.rows;
  if (dedupedTeams.removed > 0) {
    changes.push(`${dedupedTeams.removed} linha(s) duplicada(s) removida(s) em ContractTeams.`);
  }

  if (dataset.import_projects_parameters.length === 0) {
    dataset.import_projects_parameters = [{ ...DEFAULT_IMPORT_PARAMETERS }];
    changes.push("Linha padrão criada em ImportProjectsParameters.");
  }
  if (dataset.import_projects_parameters.length > 1) {
    dataset.import_projects_parameters = [dataset.import_projects_parameters[0]];
    changes.push("ImportProjectsParameters ajustado para conter apenas 1 linha.");
  }

  dataset.import_projects_parameters = dataset.import_projects_parameters.map((row) => {
    const next = { ...DEFAULT_IMPORT_PARAMETERS, ...row };
    for (const field of IMPORT_PARAMETERS_FIELDS) {
      if (!String(next[field.key] || "").trim() && DEFAULT_IMPORT_PARAMETERS[field.key] !== undefined) {
        next[field.key] = DEFAULT_IMPORT_PARAMETERS[field.key];
      }
    }
    return next;
  });

  return { dataset, changes: Array.from(new Set(changes)) };
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function datasetsAreEqual(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function extractContractDigits(contractId) {
  const match = String(contractId || "").match(/(\d{10})/);
  return match ? match[1] : "";
}

function extractDigitsFromPath(path) {
  const baseName = normalizePath(path).split("/").pop() || "";
  const match = baseName.match(/(\d{10})/);
  return match ? match[1] : "";
}

function normalizeFileName(path) {
  return (normalizePath(path).split("/").pop() || "").toLowerCase();
}

function toSourceKey(sourceFile) {
  if (sourceFile === "Contracts.csv") {
    return "contracts";
  }
  if (sourceFile === "ContractDocuments.csv") {
    return "contract_documents";
  }
  if (sourceFile === "ContractContentDocuments.csv") {
    return "contract_content_documents";
  }
  if (sourceFile === "ContractTeams.csv") {
    return "contract_teams";
  }
  if (sourceFile === "ImportProjectsParameters.csv") {
    return "import_projects_parameters";
  }
  return "";
}

function parseAttachmentPathFromIssueMessage(message) {
  const marker = "Arquivo no ZIP sem referência nos CSVs:";
  const rawMessage = String(message || "");
  const markerIndex = rawMessage.indexOf(marker);
  if (markerIndex < 0) {
    return "";
  }
  return normalizePath(rawMessage.slice(markerIndex + marker.length));
}

function deriveDocumentTitleFromPath(path) {
  const stem = stemFromPath(path);
  if (!stem) {
    return "";
  }
  return stem.replace(/^Contrato\s+\d+\s*-\s*/i, "").trim() || stem;
}

function applyContractIdRenameAcrossDataset(dataset, oldId, newId) {
  if (!oldId || !newId || oldId === newId) {
    return;
  }
  for (const row of dataset.contracts || []) {
    if (String(row.ContractId || "").trim() === oldId) {
      row.ContractId = newId;
    }
    if (String(row.ParentAgreement || "").trim() === oldId) {
      row.ParentAgreement = newId;
    }
  }
  for (const row of dataset.contract_documents || []) {
    if (String(row.ContractId || "").trim() === oldId) {
      row.ContractId = newId;
    }
  }
  for (const row of dataset.contract_content_documents || []) {
    if (String(row.Workspace || "").trim() === oldId) {
      row.Workspace = newId;
    }
  }
  for (const row of dataset.contract_teams || []) {
    if (String(row.Workspace || "").trim() === oldId) {
      row.Workspace = newId;
    }
  }
}

function buildIssueAutoFixDataset(sourceDataset, issue, allIssues = []) {
  const dataset = deepClone(sourceDataset || {});
  dataset.contracts = Array.isArray(dataset.contracts) ? dataset.contracts : [];
  dataset.contract_documents = Array.isArray(dataset.contract_documents) ? dataset.contract_documents : [];
  dataset.contract_content_documents = Array.isArray(dataset.contract_content_documents)
    ? dataset.contract_content_documents
    : [];
  dataset.contract_teams = Array.isArray(dataset.contract_teams) ? dataset.contract_teams : [];
  dataset.import_projects_parameters = Array.isArray(dataset.import_projects_parameters)
    ? dataset.import_projects_parameters
    : [];

  const changes = [];
  const sourceKey = toSourceKey(issue?.source_file);
  const sourceRows = sourceKey ? dataset[sourceKey] || [] : [];
  const rowIndex = Number.isInteger(issue?.row) ? issue.row - 2 : -1;
  const targetRow = rowIndex >= 0 && rowIndex < sourceRows.length ? sourceRows[rowIndex] : null;
  const targetContract = String(issue?.contract_id || "").trim();
  const fieldName = String(issue?.field || "").trim();

  const resolveContractOwner = (contractId) => {
    const normalizedContractId = String(contractId || "").trim();
    if (!normalizedContractId) {
      return "";
    }

    const found = (dataset.contracts || []).find(
      (row) => String(row.ContractId || "").trim() === normalizedContractId
    );
    const contractOwner = String(found?.Owner || "").trim();
    if (contractOwner) {
      return contractOwner;
    }

    return pickSingleCandidate([
      ...(dataset.contract_teams || [])
        .filter((row) => String(row.Workspace || "").trim() === normalizedContractId)
        .map((row) => row.Member),
      ...(dataset.contract_documents || [])
        .filter((row) => String(row.ContractId || "").trim() === normalizedContractId)
        .map((row) => row.Owner),
    ]);
  };

  const resolveTeamMember = (contractId) => {
    const normalizedContractId = String(contractId || "").trim();
    if (!normalizedContractId) {
      return "";
    }

    const ownerCandidate = resolveContractOwner(normalizedContractId);
    if (ownerCandidate) {
      return ownerCandidate;
    }

    return pickSingleCandidate(
      (dataset.contract_documents || [])
        .filter((row) => String(row.ContractId || "").trim() === normalizedContractId)
        .map((row) => row.Owner)
    );
  };

  if (issue?.code === "REQUIRED_FIELD") {
    if (sourceKey === "contracts") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contracts || []).filter((row) => String(row.ContractId || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        if (fieldName === "Title" && !row.Title && row.ContractId) {
          row.Title = String(row.ContractId);
          changes.push(`Título preenchido com ContractId (${row.ContractId}).`);
        }
        if (fieldName === "BaseLanguage" && !row.BaseLanguage) {
          row.BaseLanguage = "BrazilianPortuguese";
          changes.push("BaseLanguage preenchido com BrazilianPortuguese.");
        }
        if (fieldName === "HierarchicalType" && !row.HierarchicalType) {
          row.HierarchicalType = "MasterAgreement";
          changes.push("HierarchicalType preenchido com MasterAgreement.");
        }
        if (fieldName === "ContractStatus" && !row.ContractStatus) {
          row.ContractStatus = "Published";
          changes.push("ContractStatus preenchido com Published.");
        }
        if (fieldName === "Owner" && !row.Owner) {
          const ownerCandidate = resolveContractOwner(row.ContractId || targetContract);
          if (ownerCandidate) {
            row.Owner = ownerCandidate;
            changes.push(`Owner preenchido automaticamente para ${row.ContractId || targetContract}.`);
          }
        }
      }
    } else if (sourceKey === "contract_documents") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contract_documents || []).filter((row) => String(row.ContractId || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        if (fieldName === "File" && !row.File) {
          row.File = `Documentos contratos/Contrato ${extractContractDigits(row.ContractId)}.pdf`;
          changes.push(`Arquivo de contrato preenchido para ${row.ContractId}.`);
        }
        if (fieldName === "Title" && !row.Title) {
          row.Title = deriveDocumentTitleFromPath(row.File || "") || String(row.ContractId || "").trim();
          changes.push(`Título de documento preenchido para ${row.ContractId || "linha alvo"}.`);
        }
        if (fieldName === "Owner" && !row.Owner) {
          row.Owner = resolveContractOwner(row.ContractId);
          if (row.Owner) {
            changes.push(`Owner de documento preenchido a partir do contrato ${row.ContractId}.`);
          }
        }
        if (fieldName === "ContractId" && !row.ContractId && targetContract) {
          row.ContractId = targetContract;
          changes.push(`ContractId de documento preenchido com ${targetContract}.`);
        }
      }
    } else if (sourceKey === "contract_content_documents") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contract_content_documents || []).filter((row) => String(row.Workspace || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        if (fieldName === "Workspace" && !row.Workspace && targetContract) {
          row.Workspace = targetContract;
          changes.push(`Workspace CLID preenchido com ${targetContract}.`);
        }
        if (fieldName === "File" && !row.File) {
          row.File = `Documentos CLID/CLID_${extractContractDigits(row.Workspace || targetContract)}.xlsx`;
          changes.push(`Arquivo CLID preenchido para ${row.Workspace || targetContract}.`);
        }
        if ((fieldName === "title" || fieldName === "Title") && !row.title) {
          row.title = normalizePath(row.File).split("/").pop() || "";
          changes.push(`Título CLID preenchido para ${row.Workspace || targetContract}.`);
        }
      }
    } else if (sourceKey === "contract_teams") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contract_teams || []).filter((row) => String(row.Workspace || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        if (fieldName === "Workspace" && !row.Workspace && targetContract) {
          row.Workspace = targetContract;
          changes.push(`Workspace de time preenchido com ${targetContract}.`);
        }
        if (fieldName === "ProjectGroup" && !row.ProjectGroup) {
          row.ProjectGroup = "Comprador";
          changes.push("ProjectGroup preenchido com Comprador.");
        }
        if (fieldName === "Member" && !row.Member) {
          const memberCandidate = resolveTeamMember(row.Workspace || targetContract);
          if (memberCandidate) {
            row.Member = memberCandidate;
            changes.push(`Member preenchido automaticamente para ${row.Workspace || targetContract}.`);
          }
        }
      }
    } else if (sourceKey === "import_projects_parameters" && fieldName) {
      const row = targetRow || dataset.import_projects_parameters[0] || {};
      const fallback = DEFAULT_IMPORT_PARAMETERS[fieldName];
      if (!String(row[fieldName] || "").trim() && fallback !== undefined) {
        row[fieldName] = fallback;
        if (!dataset.import_projects_parameters.length) {
          dataset.import_projects_parameters = [row];
        }
        changes.push(`Parâmetro ${fieldName} restaurado com valor padrão.`);
      }
    }
  }

  if (issue?.code === "UNEXPECTED_FILE_PATH") {
    if (sourceKey === "contract_documents") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contract_documents || []).filter((row) => String(row.ContractId || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        const next = ensurePrefix(row.File, "Documentos contratos/");
        if (next && next !== row.File) {
          row.File = next;
          changes.push(`Caminho de documento ajustado para ${next}.`);
        }
      }
    } else if (sourceKey === "contract_content_documents") {
      const rowsToUpdate = targetRow
        ? [targetRow]
        : (dataset.contract_content_documents || []).filter((row) => String(row.Workspace || "").trim() === targetContract);
      for (const row of rowsToUpdate) {
        const next = ensurePrefix(row.File, "Documentos CLID/");
        if (next && next !== row.File) {
          row.File = next;
          changes.push(`Caminho CLID ajustado para ${next}.`);
        }
      }
    }
  }

  if (issue?.code === "INVALID_PARTY_FORMAT" && sourceKey === "contracts") {
    const rowsToUpdate = targetRow
      ? [targetRow]
      : (dataset.contracts || []).filter((row) => String(row.ContractId || "").trim() === targetContract);
    for (const row of rowsToUpdate) {
      if (fieldName === "Supplier") {
        const next = normalizeSapParty(row.Supplier);
        if (next !== row.Supplier) {
          row.Supplier = next;
          changes.push(`Supplier normalizado para ${next}.`);
        }
      }
      if (fieldName === "AffectedParties") {
        const next = normalizeSapParty(row.AffectedParties);
        if (next !== row.AffectedParties) {
          row.AffectedParties = next;
          changes.push(`AffectedParties normalizado para ${next}.`);
        }
      }
    }
  }

  if (issue?.code === "INVALID_ID_FORMAT" && sourceKey === "contracts") {
    const rowsToUpdate = targetRow
      ? [targetRow]
      : (dataset.contracts || []).filter((row) => String(row.ContractId || "").trim() === targetContract);
    for (const row of rowsToUpdate) {
      const oldId = String(row.ContractId || "").trim();
      const digitsMatch = oldId.match(/(\d{6,})/);
      if (!digitsMatch) {
        continue;
      }
      const nextId = `LCW${digitsMatch[1]}`;
      if (oldId !== nextId) {
        applyContractIdRenameAcrossDataset(dataset, oldId, nextId);
        changes.push(`ContractId normalizado: ${oldId} -> ${nextId}.`);
      }
    }
  }

  if (issue?.code === "DUPLICATE_DOCUMENT") {
    const match = String(issue.message || "").match(/contrato:\s*([^\s]+)\s*->\s*(.+)$/i);
    const contractId = match ? match[1].trim() : targetContract;
    const filePath = match ? normalizePath(match[2]) : "";
    const seen = new Set();
    const nextRows = [];
    for (const row of dataset.contract_documents || []) {
      const key = `${String(row.ContractId || "").trim()}|${normalizePath(row.File)}`;
      const isTarget = !contractId || String(row.ContractId || "").trim() === contractId;
      const isTargetFile = !filePath || normalizePath(row.File) === filePath;
      if (isTarget && isTargetFile && seen.has(key)) {
        changes.push(`Linha duplicada removida em ContractDocuments: ${key}.`);
        continue;
      }
      seen.add(key);
      nextRows.push(row);
    }
    dataset.contract_documents = nextRows;
  }

  if (issue?.code === "DUPLICATE_CONTENT_DOCUMENT") {
    const match = String(issue.message || "").match(/contrato:\s*([^\s]+)\s*->\s*(.+)$/i);
    const workspace = match ? match[1].trim() : targetContract;
    const filePath = match ? normalizePath(match[2]) : "";
    const seen = new Set();
    const nextRows = [];
    for (const row of dataset.contract_content_documents || []) {
      const key = `${String(row.Workspace || "").trim()}|${normalizePath(row.File)}`;
      const isTarget = !workspace || String(row.Workspace || "").trim() === workspace;
      const isTargetFile = !filePath || normalizePath(row.File) === filePath;
      if (isTarget && isTargetFile && seen.has(key)) {
        changes.push(`Linha duplicada removida em ContractContentDocuments: ${key}.`);
        continue;
      }
      seen.add(key);
      nextRows.push(row);
    }
    dataset.contract_content_documents = nextRows;
  }

  if (issue?.code === "DUPLICATE_TEAM_MEMBER") {
    const match = String(issue.message || "").match(/:\s*(.+?)\s*\/\s*(.+?)\s*\/\s*(.+)$/);
    const workspace = match ? match[1].trim() : targetContract;
    const projectGroup = match ? match[2].trim() : "";
    const member = match ? match[3].trim() : "";
    const seen = new Set();
    const nextRows = [];
    for (const row of dataset.contract_teams || []) {
      const key = `${String(row.Workspace || "").trim()}|${String(row.ProjectGroup || "").trim()}|${String(
        row.Member || ""
      ).trim()}`;
      const isTargetWorkspace = !workspace || String(row.Workspace || "").trim() === workspace;
      const isTargetGroup = !projectGroup || String(row.ProjectGroup || "").trim() === projectGroup;
      const isTargetMember = !member || String(row.Member || "").trim() === member;
      if (isTargetWorkspace && isTargetGroup && isTargetMember && seen.has(key)) {
        changes.push(`Linha duplicada removida em ContractTeams: ${key}.`);
        continue;
      }
      seen.add(key);
      nextRows.push(row);
    }
    dataset.contract_teams = nextRows;
  }

  if (issue?.code === "IMPORT_PARAMS_ROW_COUNT") {
    if ((dataset.import_projects_parameters || []).length === 0) {
      dataset.import_projects_parameters = [{ ...DEFAULT_IMPORT_PARAMETERS }];
      changes.push("ImportProjectsParameters recriado com linha padrão.");
    } else if (dataset.import_projects_parameters.length > 1) {
      dataset.import_projects_parameters = [dataset.import_projects_parameters[0]];
      changes.push("ImportProjectsParameters reduzido para 1 linha.");
    }
  }

  if (issue?.code === "DUPLICATE_ATTACHMENT_FILENAME") {
    const fileName = String(issue.message || "").split(":").pop()?.trim().toLowerCase() || "";
    if (fileName) {
      const extras = (allIssues || [])
        .filter((item) => item.code === "UNREFERENCED_ATTACHMENT")
        .map((item) => parseAttachmentPathFromIssueMessage(item.message))
        .filter(Boolean);

      const referencedDocPaths = new Set(
        (dataset.contract_documents || []).map((row) => normalizePath(row.File).toLowerCase())
      );

      for (const row of dataset.contract_documents || []) {
        if (normalizeFileName(row.File) !== fileName) {
          continue;
        }
        const contractDigits = extractContractDigits(row.ContractId);
        const fileDigits = extractDigitsFromPath(row.File);
        if (contractDigits && fileDigits && contractDigits !== fileDigits) {
          const nextPath = normalizePath(row.File).replace(fileDigits, contractDigits);
          if (nextPath !== row.File) {
            row.File = nextPath;
            referencedDocPaths.add(nextPath.toLowerCase());
            changes.push(`Anexo alinhado ao ContractId ${row.ContractId}: ${nextPath}.`);
          }
        }
      }

      const groupedDocRows = new Map();
      (dataset.contract_documents || []).forEach((row, index) => {
        const key = `${String(row.ContractId || "").trim()}|${normalizePath(row.File).toLowerCase()}`;
        if (!groupedDocRows.has(key)) {
          groupedDocRows.set(key, []);
        }
        groupedDocRows.get(key).push(index);
      });

      for (const [key, indexes] of groupedDocRows.entries()) {
        if (indexes.length < 2) {
          continue;
        }
        const [contractId] = key.split("|");
        const contractDigits = extractContractDigits(contractId);
        const contractExtras = extras.filter(
          (path) =>
            normalizePath(path).toLowerCase().startsWith("documentos contratos/") &&
            (!contractDigits || path.includes(contractDigits)) &&
            !referencedDocPaths.has(normalizePath(path).toLowerCase())
        );
        for (let i = 1; i < indexes.length; i += 1) {
          const row = dataset.contract_documents[indexes[i]];
          const titleWords = String(row.Title || "")
            .toLowerCase()
            .split(/[^a-z0-9]+/)
            .filter((token) => token.length > 2);
          const bestExtraIndex = contractExtras.findIndex((path) => {
            const pathLower = normalizePath(path).toLowerCase();
            return titleWords.length === 0 || titleWords.some((token) => pathLower.includes(token));
          });
          if (bestExtraIndex < 0) {
            continue;
          }
          const nextPath = normalizePath(contractExtras[bestExtraIndex]);
          contractExtras.splice(bestExtraIndex, 1);
          row.File = nextPath;
          referencedDocPaths.add(nextPath.toLowerCase());
          changes.push(`Anexo duplicado redistribuído para caminho sugerido: ${nextPath}.`);
        }
      }

      const dedupedClid = dedupeRows(dataset.contract_content_documents || []);
      if (dedupedClid.removed > 0) {
        dataset.contract_content_documents = dedupedClid.rows;
        changes.push(`${dedupedClid.removed} linha(s) duplicada(s) removida(s) em CLID.`);
      }
    }
  }

  if (issue?.code === "UNREFERENCED_ATTACHMENT") {
    const attachmentPath = parseAttachmentPathFromIssueMessage(issue.message);
    if (attachmentPath.toLowerCase().startsWith("documentos contratos/")) {
      const digits = extractDigitsFromPath(attachmentPath);
      const contractId = digits ? `LCW${digits}` : "";
      const contractExists = (dataset.contracts || []).some((row) => String(row.ContractId || "").trim() === contractId);
      if (contractExists) {
        const alreadyMapped = (dataset.contract_documents || []).some(
          (row) => normalizePath(row.File).toLowerCase() === attachmentPath.toLowerCase()
        );
        if (!alreadyMapped) {
          dataset.contract_documents.push({
            ContractId: contractId,
            File: attachmentPath,
            Title: deriveDocumentTitleFromPath(attachmentPath),
            Folder: "",
            Owner: resolveContractOwner(contractId),
            Status: "",
          });
          changes.push(`Anexo extra referenciado em ContractDocuments: ${attachmentPath}.`);
        }
      }
    }

    if (attachmentPath.toLowerCase().startsWith("documentos clid/")) {
      const digits = extractDigitsFromPath(attachmentPath);
      const workspace = digits ? `LCW${digits}` : "";
      const contractExists = (dataset.contracts || []).some((row) => String(row.ContractId || "").trim() === workspace);
      if (contractExists) {
        const alreadyMapped = (dataset.contract_content_documents || []).some(
          (row) => normalizePath(row.File).toLowerCase() === attachmentPath.toLowerCase()
        );
        if (!alreadyMapped) {
          dataset.contract_content_documents.push({
            Workspace: workspace,
            File: attachmentPath,
            title: normalizePath(attachmentPath).split("/").pop() || "",
          });
          changes.push(`Anexo CLID extra referenciado em ContractContentDocuments: ${attachmentPath}.`);
        }
      }
    }
  }

  return { dataset, changes: Array.from(new Set(changes)) };
}

function getIssueAutoFixPriority(issue) {
  if (issue?.code === "INVALID_ID_FORMAT") {
    return 0;
  }
  if (issue?.code === "REQUIRED_FIELD" && issue?.source_file === "Contracts.csv") {
    return 1;
  }
  if (issue?.code === "REQUIRED_FIELD" && issue?.source_file === "ContractDocuments.csv") {
    return 2;
  }
  if (issue?.code === "REQUIRED_FIELD" && issue?.source_file === "ContractTeams.csv") {
    return 3;
  }
  if (issue?.code === "REQUIRED_FIELD") {
    return 4;
  }
  if (issue?.code === "UNEXPECTED_FILE_PATH" || issue?.code === "INVALID_PARTY_FORMAT") {
    return 5;
  }
  if (
    issue?.code === "DUPLICATE_DOCUMENT" ||
    issue?.code === "DUPLICATE_CONTENT_DOCUMENT" ||
    issue?.code === "DUPLICATE_TEAM_MEMBER"
  ) {
    return 6;
  }
  if (issue?.code === "IMPORT_PARAMS_ROW_COUNT") {
    return 7;
  }
  if (issue?.code === "DUPLICATE_ATTACHMENT_FILENAME" || issue?.code === "UNREFERENCED_ATTACHMENT") {
    return 8;
  }
  return 99;
}

function buildBatchAutoFixDataset(sourceDataset, allIssues = []) {
  let { dataset, changes } = buildSafeAutoFixDataset(sourceDataset);

  const fixableIssues = (allIssues || [])
    .filter((issue) => AUTO_FIXABLE_CODES.has(issue.code))
    .slice()
    .sort((left, right) => getIssueAutoFixPriority(left) - getIssueAutoFixPriority(right));

  for (const issue of fixableIssues) {
    const issueFix = buildIssueAutoFixDataset(dataset, issue, allIssues);
    dataset = issueFix.dataset;
    changes = [...changes, ...(issueFix.changes || [])];
  }

  return { dataset, changes: Array.from(new Set(changes)) };
}

function formatDateTimeLabel(isoDate) {
  if (!isoDate) {
    return "-";
  }
  const parsed = new Date(isoDate);
  if (Number.isNaN(parsed.getTime())) {
    return isoDate;
  }
  return parsed.toLocaleString("pt-BR");
}

export default function App() {
  const sourceRef = useRef(null);
  const attachmentsRef = useRef(null);
  const rulesRef = useRef(null);
  const analyzeRef = useRef(null);
  const reviewRef = useRef(null);
  const contractAttachmentsInputRef = useRef(null);
  const clidAttachmentsInputRef = useRef(null);
  const wizardModeInitializedRef = useRef(false);
  const skipNextWizardResetRef = useRef(false);
  const sessionHydratedRef = useRef(false);

  const [wizardMode, setWizardMode] = useState("zip");
  const [zipFile, setZipFile] = useState(null);
  const [unifiedFile, setUnifiedFile] = useState(null);
  const [contractAttachmentFiles, setContractAttachmentFiles] = useState([]);
  const [clidAttachmentFiles, setClidAttachmentFiles] = useState([]);

  const [manualRows, setManualRows] = useState([createManualRow()]);
  const [analysis, setAnalysis] = useState(null);
  const [analysisSource, setAnalysisSource] = useState("");
  const [sourceZipForExport, setSourceZipForExport] = useState(null);

  const [loading, setLoading] = useState(false);
  const [packageExportStatus, setPackageExportStatus] = useState("idle");
  const [packageExportDownloadedBytes, setPackageExportDownloadedBytes] = useState(0);
  const [packageExportTotalBytes, setPackageExportTotalBytes] = useState(null);
  const [error, setError] = useState("");
  const [successMessage, setSuccessMessage] = useState("");

  const [validationRules, setValidationRules] = useState(DEFAULT_RULES);
  const [importParameters, setImportParameters] = useState(DEFAULT_IMPORT_PARAMETERS);
  const [severityFilter, setSeverityFilter] = useState("all");
  const [contractFilter, setContractFilter] = useState("");
  const [previewFileKey, setPreviewFileKey] = useState("contracts");
  const [showOnlyCriticalPending, setShowOnlyCriticalPending] = useState(true);

  const [profiles, setProfiles] = useState([]);
  const [profilesLoading, setProfilesLoading] = useState(false);
  const [profileNameInput, setProfileNameInput] = useState("");
  const [selectedProfileName, setSelectedProfileName] = useState("");

  const [attachmentValidation, setAttachmentValidation] = useState(null);
  const [attachmentValidationLoading, setAttachmentValidationLoading] = useState(false);
  const [currentStepId, setCurrentStepId] = useState("source");
  const [reviewTab, setReviewTab] = useState("summary");
  const [pendingFixPlan, setPendingFixPlan] = useState(null);
  const [lastAppliedSnapshot, setLastAppliedSnapshot] = useState(null);
  const [executionRuns, setExecutionRuns] = useState([]);
  const [runsLoading, setRunsLoading] = useState(false);

  const manualRowsStarted = useMemo(
    () => manualRows.filter((row) => hasManualRowInput(row)).length,
    [manualRows]
  );
  const manualRowsCompleted = useMemo(
    () => manualRows.filter((row) => hasManualRowInput(row) && isManualRowComplete(row)).length,
    [manualRows]
  );
  const manualRowsIncomplete = useMemo(
    () => manualRows.filter((row) => hasManualRowInput(row) && !isManualRowComplete(row)).length,
    [manualRows]
  );
  const attachmentUploadEntries = useMemo(
    () => [
      ...buildAttachmentUploadEntries(contractAttachmentFiles, "Documentos contratos"),
      ...buildAttachmentUploadEntries(clidAttachmentFiles, "Documentos CLID"),
    ],
    [contractAttachmentFiles, clidAttachmentFiles]
  );
  const selectedAttachmentFilesCount = attachmentUploadEntries.length;
  const uploadedAttachmentsSource = selectedAttachmentFilesCount > 0 ? attachmentUploadEntries : null;

  useEffect(() => {
    [contractAttachmentsInputRef.current, clidAttachmentsInputRef.current].forEach((input) => {
      if (!input) {
        return;
      }
      input.setAttribute("webkitdirectory", "");
      input.setAttribute("directory", "");
    });
  }, []);

  useEffect(() => {
    let active = true;
    fetch(`${API_PREFIX}/rules/default`)
      .then((response) => {
        if (!response.ok) {
          throw new Error("Não foi possível carregar regras padrão.");
        }
        return response.json();
      })
      .then((payload) => {
        if (active) {
          setValidationRules(payload);
        }
      })
      .catch(() => {
        if (active) {
          setValidationRules(DEFAULT_RULES);
        }
      });

    return () => {
      active = false;
    };
  }, []);

  async function loadProfiles() {
    setProfilesLoading(true);
    try {
      const response = await fetch(`${API_PREFIX}/profiles`);
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const payload = await response.json();
      const nextProfiles = payload.profiles || [];
      setProfiles(nextProfiles);
      if (!selectedProfileName && nextProfiles.length > 0) {
        setSelectedProfileName(nextProfiles[0].name);
      }
    } catch {
      setProfiles([]);
    } finally {
      setProfilesLoading(false);
    }
  }

  async function loadExecutionRuns(limit = 20) {
    setRunsLoading(true);
    try {
      const response = await fetch(`${API_PREFIX}/runs?limit=${limit}`);
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const payload = await response.json();
      setExecutionRuns(payload.runs || []);
    } catch {
      setExecutionRuns([]);
    } finally {
      setRunsLoading(false);
    }
  }

  useEffect(() => {
    loadProfiles();
    loadExecutionRuns();
  }, []);

  useEffect(() => {
    let payload = null;
    try {
      const raw = localStorage.getItem(SESSION_STORAGE_KEY);
      payload = raw ? JSON.parse(raw) : null;
    } catch {
      payload = null;
    }

    if (payload && typeof payload === "object") {
      const nextWizardMode = String(payload.wizardMode || "").trim();
      if (["zip", "unified", "manual"].includes(nextWizardMode) && nextWizardMode !== wizardMode) {
        skipNextWizardResetRef.current = true;
        setWizardMode(nextWizardMode);
      }

      if (payload.validationRules && typeof payload.validationRules === "object") {
        setValidationRules(payload.validationRules);
      }
      if (payload.importParameters && typeof payload.importParameters === "object") {
        setImportParameters({ ...DEFAULT_IMPORT_PARAMETERS, ...payload.importParameters });
      }
      if (typeof payload.severityFilter === "string") {
        setSeverityFilter(payload.severityFilter);
      }
      if (typeof payload.contractFilter === "string") {
        setContractFilter(payload.contractFilter);
      }
      if (typeof payload.showOnlyCriticalPending === "boolean") {
        setShowOnlyCriticalPending(payload.showOnlyCriticalPending);
      }
      if (payload.reviewTab === "summary" || payload.reviewTab === "issues" || payload.reviewTab === "attachments" || payload.reviewTab === "history") {
        setReviewTab(payload.reviewTab);
      }
      if (typeof payload.profileNameInput === "string") {
        setProfileNameInput(payload.profileNameInput);
      }
      if (typeof payload.selectedProfileName === "string") {
        setSelectedProfileName(payload.selectedProfileName);
      }
      if (Array.isArray(payload.manualRows) && payload.manualRows.length > 0) {
        setManualRows(payload.manualRows);
      }
      if (payload.analysis && typeof payload.analysis === "object") {
        setAnalysis(payload.analysis);
        setAnalysisSource(String(payload.analysisSource || ""));
        if (payload.attachmentValidation && typeof payload.attachmentValidation === "object") {
          setAttachmentValidation(payload.attachmentValidation);
        }
        setCurrentStepId(String(payload.currentStepId || "review"));
      }
      if (typeof payload.previewFileKey === "string") {
        setPreviewFileKey(payload.previewFileKey);
      }
    }

    sessionHydratedRef.current = true;
  }, []);

  useEffect(() => {
    if (!sessionHydratedRef.current) {
      return;
    }
    let analysisSnapshot = analysis;
    if (analysisSnapshot) {
      try {
        const serialized = JSON.stringify(analysisSnapshot);
        if (serialized.length > 2_500_000) {
          analysisSnapshot = null;
        }
      } catch {
        analysisSnapshot = null;
      }
    }

    const payload = {
      wizardMode,
      validationRules,
      importParameters,
      severityFilter,
      contractFilter,
      showOnlyCriticalPending,
      profileNameInput,
      selectedProfileName,
      manualRows,
      analysis: analysisSnapshot,
      analysisSource,
      attachmentValidation,
      currentStepId,
      reviewTab,
      previewFileKey,
      savedAt: new Date().toISOString(),
    };

    try {
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Ignore storage quota and serialization failures.
    }
  }, [
    wizardMode,
    validationRules,
    importParameters,
    severityFilter,
    contractFilter,
    showOnlyCriticalPending,
    profileNameInput,
    selectedProfileName,
    manualRows,
    analysis,
    analysisSource,
    attachmentValidation,
    currentStepId,
    reviewTab,
    previewFileKey,
  ]);

  useEffect(() => {
    if (!wizardModeInitializedRef.current) {
      wizardModeInitializedRef.current = true;
      return;
    }
    if (skipNextWizardResetRef.current) {
      skipNextWizardResetRef.current = false;
      return;
    }
    setCurrentStepId("source");
    setAnalysis(null);
    setAnalysisSource("");
    setSourceZipForExport(null);
    setAttachmentValidation(null);
    setPendingFixPlan(null);
    setLastAppliedSnapshot(null);
    setReviewTab("summary");
    setError("");
    setSuccessMessage("");
  }, [wizardMode]);

  const summary = analysis?.report?.summary;
  const recordCounts = analysis?.report?.record_counts || {};
  const issues = analysis?.report?.issues || [];
  const executiveSummary = analysis?.executive_summary || null;
  const activeProfileName = (selectedProfileName || profileNameInput || "").trim();

  const groupedIssues = useMemo(
    () => ({
      error: issues.filter((item) => item.severity === "error"),
      warning: issues.filter((item) => item.severity === "warning"),
      info: issues.filter((item) => item.severity === "info"),
    }),
    [issues]
  );

  const filteredIssues = useMemo(() => {
    const contractTerm = contractFilter.trim().toLowerCase();
    const nextIssues = issues.filter((item) => {
      const severityMatches = severityFilter === "all" ? true : item.severity === severityFilter;
      const contractMatches =
        !contractTerm || (item.contract_id || "").toLowerCase().includes(contractTerm);
      return severityMatches && contractMatches;
    });
    nextIssues.sort((left, right) => {
      const lineDiff = getIssueLineNumber(left) - getIssueLineNumber(right);
      if (lineDiff !== 0) {
        return lineDiff;
      }
      const fileDiff = String(left.source_file || "").localeCompare(String(right.source_file || ""), "pt-BR");
      if (fileDiff !== 0) {
        return fileDiff;
      }
      const contractDiff = String(left.contract_id || "").localeCompare(String(right.contract_id || ""), "pt-BR");
      if (contractDiff !== 0) {
        return contractDiff;
      }
      return String(left.code || "").localeCompare(String(right.code || ""), "pt-BR");
    });
    return nextIssues;
  }, [issues, severityFilter, contractFilter]);

  const attachmentFixSuggestions = useMemo(() => {
    if (!analysis?.dataset) {
      return [];
    }
    const attachmentIssues = issues.filter(
      (item) => item.code === "DUPLICATE_ATTACHMENT_FILENAME" || item.code === "UNREFERENCED_ATTACHMENT"
    );
    const suggestions = [];
    const seen = new Set();

    attachmentIssues.forEach((issue, index) => {
      const result = buildIssueAutoFixDataset(analysis.dataset, issue, issues);
      if (datasetsAreEqual(result.dataset, analysis.dataset)) {
        return;
      }
      const signature = `${issue.code}|${result.changes.join("|")}`;
      if (seen.has(signature)) {
        return;
      }
      seen.add(signature);
      suggestions.push({
        id: `attachment-fix-${index}`,
        issue,
        changes: result.changes,
        dataset: result.dataset,
      });
    });

    return suggestions;
  }, [analysis, issues]);

  const issuesByContract = useMemo(() => {
    const map = new Map();
    for (const issue of issues) {
      const key = issue.contract_id || "(sem contrato)";
      if (!map.has(key)) {
        map.set(key, { contract: key, error: 0, warning: 0, info: 0, total: 0 });
      }
      const row = map.get(key);
      row[issue.severity] += 1;
      row.total += 1;
    }
    return Array.from(map.values()).sort((a, b) => b.total - a.total);
  }, [issues]);

  const sourceReady = useMemo(() => {
    if (wizardMode === "zip") {
      return Boolean(zipFile);
    }
    if (wizardMode === "unified") {
      return Boolean(unifiedFile);
    }
    return manualRowsCompleted > 0 && manualRowsIncomplete === 0;
  }, [wizardMode, zipFile, unifiedFile, manualRowsCompleted, manualRowsIncomplete]);

  const attachmentsReady = wizardMode === "zip" ? true : selectedAttachmentFilesCount > 0;
  const analysisReady = Boolean(analysis);

  const steps = useMemo(() => {
    const base = [{ id: "source", label: "Fonte dos dados", ready: sourceReady }];
    if (wizardMode !== "zip") {
      base.push({ id: "attachments", label: "Anexos e CLID", ready: attachmentsReady });
    }
    base.push(
      { id: "rules", label: "Regras e parâmetros", ready: true },
      { id: "analyze", label: "Pré-análise", ready: analysisReady },
      { id: "review", label: "Laudo final", ready: analysisReady }
    );
    return base;
  }, [wizardMode, sourceReady, attachmentsReady, analysisReady]);

  const stepIndexById = useMemo(() => {
    const map = new Map();
    steps.forEach((step, index) => map.set(step.id, index + 1));
    return map;
  }, [steps]);

  const nextPendingStep = useMemo(() => {
    return steps.find((step) => !step.ready) || steps[steps.length - 1];
  }, [steps]);

  const progressPercent = useMemo(() => {
    if (steps.length === 0) {
      return 0;
    }
    const completed = steps.filter((step) => step.ready).length;
    return Math.round((completed / steps.length) * 100);
  }, [steps]);

  const contractsTotal = analysis?.dataset?.contracts?.length || 0;
  const contractDocumentsTotal = analysis?.dataset?.contract_documents?.length || 0;
  const clidDocumentsTotal = analysis?.dataset?.contract_content_documents?.length || 0;
  const importParamsTotal = analysis?.dataset?.import_projects_parameters?.length || 0;
  const hasMappedDocuments = contractDocumentsTotal > 0 || clidDocumentsTotal > 0;
  const attachmentSourceAvailable =
    analysisSource === "zip" ? Boolean(sourceZipForExport) : selectedAttachmentFilesCount > 0;

  const attachmentCheckRequired = hasMappedDocuments && analysisSource !== "zip";
  const attachmentCheckExecuted = !attachmentCheckRequired || Boolean(attachmentValidation);
  const attachmentCheckErrors = attachmentValidation?.summary?.errors || 0;
  const attachmentCheckWarnings = attachmentValidation?.summary?.warnings || 0;
  const attachmentMissingReferences =
    attachmentValidation?.stats?.missing_files ?? attachmentCheckErrors;
  const mappedAttachmentParts = [];
  if (contractDocumentsTotal > 0) {
    mappedAttachmentParts.push(`${contractDocumentsTotal} anexo(s) de contrato`);
  }
  if (clidDocumentsTotal > 0) {
    mappedAttachmentParts.push(`${clidDocumentsTotal} documento(s) CLID`);
  }
  const mappedAttachmentLabel = mappedAttachmentParts.join(" e ");
  const attachmentSourceLabel = attachmentCheckRequired
    ? attachmentSourceAvailable
      ? `Arquivos das pastas de anexos/CLID informados para validar ${mappedAttachmentLabel}.`
      : `Selecione as pastas de anexos/CLID para validar ${mappedAttachmentLabel} antes da exportação.`
    : "Sem anexos/CLID mapeados: seleção adicional de pastas não é obrigatória.";
  const attachmentValidationExecutionRequired =
    attachmentCheckRequired && attachmentSourceAvailable;
  const attachmentValidationExecutionLabel = attachmentValidationExecutionRequired
    ? attachmentCheckExecuted
      ? "Validação inteligente de anexos executada."
      : "Clique em 'Validar anexos selecionados' para conferir as referências antes de exportar."
    : "Validação inteligente de anexos será executada após selecionar as pastas de anexos/CLID.";
  const attachmentReferenceCheckRequired =
    attachmentCheckRequired && attachmentSourceAvailable && attachmentCheckExecuted;
  const attachmentReferenceLabel = attachmentReferenceCheckRequired
    ? attachmentMissingReferences === 0
      ? "Todas as referências de anexos/CLID foram encontradas nas pastas selecionadas."
      : `${attachmentMissingReferences} referência(s) de anexos/CLID não foram encontradas nas pastas selecionadas. Revise a coluna File nos CSVs ou inclua os arquivos faltantes.`
    : "Referências de anexos/CLID serão verificadas após validar as pastas selecionadas.";
  const attachmentWarningsLabel = attachmentReferenceCheckRequired
    ? attachmentCheckWarnings === 0
      ? "Sem avisos de anexos/CLID."
      : `${attachmentCheckWarnings} aviso(s) de anexos/CLID detectados.`
    : "Avisos de anexos/CLID serão exibidos após a validação do ZIP.";
  const attachmentPendingActionKey =
    !attachmentSourceAvailable && wizardMode !== "zip" ? "goToAttachmentsStep" : "openAttachmentsTab";

  const readinessItems = [
    { key: "analysis-no-errors", label: "Pré-análise sem erros", ok: (summary?.errors ?? 0) === 0, critical: true },
    { key: "contracts-has-rows", label: "Contracts.csv com registros", ok: contractsTotal > 0, critical: true },
    {
      key: "import-params-present",
      label: "ImportProjectsParameters.csv presente",
      ok: importParamsTotal === 1,
      critical: true,
    },
    {
      key: "attachments-source",
      label: attachmentSourceLabel,
      ok: !attachmentCheckRequired || attachmentSourceAvailable,
      critical: attachmentCheckRequired,
      actionKey: attachmentPendingActionKey,
    },
    {
      key: "attachments-validation-ran",
      label: attachmentValidationExecutionLabel,
      ok: !attachmentValidationExecutionRequired || attachmentCheckExecuted,
      critical: attachmentValidationExecutionRequired,
      actionKey: "openAttachmentsTab",
    },
    {
      key: "attachments-reference-check",
      label: attachmentReferenceLabel,
      ok: !attachmentReferenceCheckRequired || attachmentMissingReferences === 0,
      critical: attachmentReferenceCheckRequired,
      actionKey: "openAttachmentsTab",
    },
    {
      key: "attachments-warnings",
      label: attachmentWarningsLabel,
      ok: !attachmentReferenceCheckRequired || attachmentCheckWarnings === 0,
      critical: false,
      actionKey: "openAttachmentsTab",
    },
  ];

  const blockingReadinessItems = readinessItems.filter((item) => item.critical && !item.ok);
  const readyToImport = blockingReadinessItems.length === 0;
  const visibleReadinessItems = showOnlyCriticalPending
    ? readinessItems.filter((item) => item.critical && !item.ok)
    : readinessItems;

  const previewOptions = [
    { key: "contracts", label: "Contracts.csv", rows: analysis?.dataset?.contracts || [] },
    {
      key: "contract_documents",
      label: "ContractDocuments.csv",
      rows: analysis?.dataset?.contract_documents || [],
    },
    {
      key: "contract_content_documents",
      label: "ContractContentDocuments.csv",
      rows: analysis?.dataset?.contract_content_documents || [],
    },
    { key: "contract_teams", label: "ContractTeams.csv", rows: analysis?.dataset?.contract_teams || [] },
    {
      key: "import_projects_parameters",
      label: "ImportProjectsParameters.csv",
      rows: analysis?.dataset?.import_projects_parameters || [],
    },
  ];

  const selectedPreview = previewOptions.find((option) => option.key === previewFileKey) || previewOptions[0];
  const selectedPreviewHeaders = EXPECTED_COLUMNS[selectedPreview.key] || [];
  const selectedPreviewRows = selectedPreview.rows.slice(0, 5);

  function scrollToStep(stepId) {
    setCurrentStepId(stepId);
    const refs = {
      source: sourceRef,
      attachments: attachmentsRef,
      rules: rulesRef,
      analyze: analyzeRef,
      review: reviewRef,
    };
    const target = refs[stepId]?.current;
    if (target) {
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function updateRuleField(field, value) {
    setValidationRules((prev) => ({ ...prev, [field]: value }));
  }

  function updateImportParameterField(field, value) {
    setImportParameters((prev) => ({ ...prev, [field]: value }));
  }

  function restoreDefaultImportParameters() {
    setImportParameters({ ...DEFAULT_IMPORT_PARAMETERS });
  }

  function handleManualChange(index, field, value) {
    setManualRows((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], [field]: value };
      return copy;
    });
  }

  function addManualRow() {
    setManualRows((prev) => [...prev, createManualRow()]);
  }

  function removeManualRow(index) {
    setManualRows((prev) => {
      if (prev.length === 1) {
        return [createManualRow()];
      }
      return prev.filter((_, rowIndex) => rowIndex !== index);
    });
  }

  function applyProfileFromSelection() {
    const profile = profiles.find((item) => item.name === selectedProfileName);
    if (!profile) {
      setError("Selecione um perfil existente para carregar.");
      return;
    }
    setValidationRules(profile.validation_rules || DEFAULT_RULES);
    setImportParameters({ ...DEFAULT_IMPORT_PARAMETERS, ...(profile.import_parameters || {}) });
    setProfileNameInput(profile.name);
    setError("");
    setSuccessMessage(`Perfil '${profile.name}' carregado com sucesso.`);
  }

  async function saveCurrentProfile() {
    const profileName = profileNameInput.trim();
    if (!profileName) {
      setError("Informe o nome do perfil antes de salvar.");
      return;
    }

    setLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_PREFIX}/profiles`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: profileName,
          validation_rules: validationRules,
          import_parameters: importParameters,
        }),
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const payload = await response.json();
      const nextProfiles = payload.profiles || [];
      setProfiles(nextProfiles);
      setSelectedProfileName(profileName);
      setSuccessMessage(`Perfil '${profileName}' salvo com sucesso.`);
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    } finally {
      setLoading(false);
    }
  }

  async function deleteSelectedProfile() {
    const profileName = selectedProfileName.trim();
    if (!profileName) {
      setError("Selecione um perfil para excluir.");
      return;
    }

    setLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_PREFIX}/profiles/${encodeURIComponent(profileName)}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const payload = await response.json();
      const nextProfiles = payload.profiles || [];
      setProfiles(nextProfiles);
      setSelectedProfileName(nextProfiles[0]?.name || "");
      setSuccessMessage(`Perfil '${profileName}' removido.`);
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    } finally {
      setLoading(false);
    }
  }

  async function runRequest(requestFn, source, sourceZipFile = null) {
    setLoading(true);
    setError("");
    try {
      const result = await requestFn();
      setAnalysis(result);
      setAnalysisSource(source);
      setSourceZipForExport(sourceZipFile);
      setAttachmentValidation(null);
      setPendingFixPlan(null);
      setLastAppliedSnapshot(null);
      setPreviewFileKey("contracts");
      const loadedParams = result?.dataset?.import_projects_parameters?.[0];
      if (loadedParams && typeof loadedParams === "object") {
        setImportParameters((prev) => ({ ...prev, ...loadedParams }));
      }
      if (result?.run_id) {
        loadExecutionRuns();
      }
      return result;
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
      return null;
    } finally {
      setLoading(false);
    }
  }

  async function analyzeZipMode() {
    if (!zipFile) {
      setError("Passo 1: selecione o pacote Ariba (.zip).");
      return null;
    }
    return runRequest(async () => {
      const form = new FormData();
      form.append("file", zipFile);
      form.append("validation_rules", JSON.stringify(validationRules));
      form.append("import_parameters_override", JSON.stringify(importParameters));
      if (activeProfileName) {
        form.append("profile_name", activeProfileName);
      }
      const response = await fetch(`${API_PREFIX}/analyze/upload`, { method: "POST", body: form });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      return response.json();
    }, "zip", zipFile);
  }

  async function analyzeUnifiedMode() {
    if (!unifiedFile) {
      setError("Passo 1: selecione o arquivo de base única (.csv/.xlsx).");
      return null;
    }
    return runRequest(async () => {
      const form = new FormData();
      form.append("file", unifiedFile);
      form.append("validation_rules", JSON.stringify(validationRules));
      form.append("import_parameters_override", JSON.stringify(importParameters));
      if (activeProfileName) {
        form.append("profile_name", activeProfileName);
      }
      const response = await fetch(`${API_PREFIX}/unified/upload`, { method: "POST", body: form });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      return response.json();
    }, "unified", null);
  }

  async function analyzeManualMode() {
    const startedRows = manualRows.filter((row) => hasManualRowInput(row));
    if (startedRows.length === 0) {
      setError("Passo 1: preencha ao menos uma linha manual.");
      return null;
    }
    const incompleteRows = startedRows.filter((row) => !isManualRowComplete(row));
    if (incompleteRows.length > 0) {
      setError("Passo 1: preencha todos os campos das linhas manuais antes de continuar.");
      return null;
    }
    return runRequest(async () => {
      const response = await fetch(`${API_PREFIX}/unified/manual`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: startedRows,
          validation_rules: validationRules,
          import_parameters_override: importParameters,
          profile_name: activeProfileName || null,
        }),
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      return response.json();
    }, "manual", null);
  }

  function handleContractAttachmentsChange(event) {
    setContractAttachmentFiles(Array.from(event.target.files || []));
    setAttachmentValidation(null);
  }

  function handleClidAttachmentsChange(event) {
    setClidAttachmentFiles(Array.from(event.target.files || []));
    setAttachmentValidation(null);
  }

  function appendAttachmentsPayload(form, attachmentSource) {
    if (!attachmentSource) {
      return;
    }

    if (attachmentSource instanceof File) {
      form.append("attachments_zip", attachmentSource);
      return;
    }

    attachmentSource.forEach(({ file, path }) => {
      const fieldName = path.startsWith("Documentos CLID/") ? "clid_attachments" : "contract_attachments";
      form.append(fieldName, file, file.name);
    });
  }

  async function validateAttachmentsAgainstDataset(dataset, attachmentSource) {
    if (!dataset || !attachmentSource) {
      return null;
    }
    setAttachmentValidationLoading(true);
    setError("");
    try {
      const form = new FormData();
      form.append("dataset_json", JSON.stringify(dataset));
      appendAttachmentsPayload(form, attachmentSource);
      const response = await fetch(`${API_PREFIX}/attachments/validate`, {
        method: "POST",
        body: form,
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const payload = await response.json();
      setAttachmentValidation(payload);
      return payload;
    } catch (requestError) {
      setAttachmentValidation(null);
      setError(normalizeUiError(requestError.message));
      return null;
    } finally {
      setAttachmentValidationLoading(false);
    }
  }

  async function runAttachmentValidation() {
    if (!analysis?.dataset) {
      setError("Execute a pré-análise antes de validar anexos.");
      return;
    }
    const source = analysisSource === "zip" ? sourceZipForExport || zipFile : uploadedAttachmentsSource;
    if (!source) {
      setError("Selecione as pastas de anexos/CLID para validar.");
      return;
    }
    const result = await validateAttachmentsAgainstDataset(analysis.dataset, source);
    if (result) {
      setSuccessMessage("Validação inteligente de anexos concluída.");
    }
  }

  async function executeAnalysis() {
    let result = null;
    if (wizardMode === "zip") {
      result = await analyzeZipMode();
    } else if (wizardMode === "unified") {
      result = await analyzeUnifiedMode();
    } else {
      result = await analyzeManualMode();
    }
    if (!result) {
      return;
    }

    const hasDocuments =
      (result.dataset?.contract_documents?.length || 0) > 0 ||
      (result.dataset?.contract_content_documents?.length || 0) > 0;
    if (hasDocuments) {
      const source = wizardMode === "zip" ? zipFile : uploadedAttachmentsSource;
      if (source) {
        await validateAttachmentsAgainstDataset(result.dataset, source);
      }
    }

    setSuccessMessage("Pré-análise executada com sucesso.");
    scrollToStep("review");
  }

  async function runAnalyzeJson(dataset, source = analysisSource, sourceZipFile = sourceZipForExport) {
    return runRequest(async () => {
      const response = await fetch(`${API_PREFIX}/analyze/json`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ dataset, validation_rules: validationRules, profile_name: activeProfileName || null }),
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      return response.json();
    }, source, sourceZipFile);
  }

  async function executePreparedFixPlan(plan) {
    if (!plan?.dataset || !analysis?.dataset) {
      setError("Execute a pré-análise antes de aplicar correções automáticas.");
      return;
    }

    if (datasetsAreEqual(plan.dataset, analysis.dataset)) {
      setPendingFixPlan(null);
      setError("");
      setSuccessMessage("Nenhuma correção automática segura foi necessária.");
      setReviewTab("issues");
      scrollToStep("review");
      return;
    }

    const snapshot = {
      analysis,
      analysisSource,
      sourceZipForExport,
      attachmentValidation,
      previewFileKey,
    };
    const result = await runAnalyzeJson(plan.dataset);
    if (!result) {
      return;
    }

    setLastAppliedSnapshot(snapshot);
    setPendingFixPlan(null);

    const attachmentsSource =
      snapshot.analysisSource === "zip"
        ? snapshot.sourceZipForExport
        : uploadedAttachmentsSource;
    if (attachmentsSource) {
      await validateAttachmentsAgainstDataset(result.dataset, attachmentsSource);
    }

    const changeCount = plan.changes?.length || 1;
    setError("");
    setSuccessMessage(`${plan.label}: ${changeCount} ajuste(s) aplicado(s).`);
    setReviewTab("issues");
    scrollToStep("review");
  }

  function queueFixPlan({ label, dataset: nextDataset, changes }) {
    if (!analysis?.dataset) {
      setError("Execute a pré-análise antes de preparar correções.");
      return;
    }
    if (datasetsAreEqual(nextDataset, analysis.dataset)) {
      setSuccessMessage("Nenhuma correção automática segura foi necessária.");
      scrollToStep("review");
      return;
    }
    setPendingFixPlan({
      label,
      dataset: nextDataset,
      changes: Array.from(new Set(changes || [])),
      created_at: new Date().toISOString(),
    });
    setError("");
    setSuccessMessage("Prévia de correções pronta. Revise e clique em aplicar.");
    scrollToStep("review");
  }

  async function applyPendingFixPlan() {
    if (!pendingFixPlan?.dataset || !analysis?.dataset) {
      return;
    }
    await executePreparedFixPlan(pendingFixPlan);
  }

  function cancelPendingFixPlan() {
    setPendingFixPlan(null);
    setSuccessMessage("Prévia de correções cancelada.");
  }

  function undoLastFixPlan() {
    if (!lastAppliedSnapshot) {
      setError("Nenhuma correção aplicada para desfazer.");
      return;
    }
    setAnalysis(lastAppliedSnapshot.analysis || null);
    setAnalysisSource(lastAppliedSnapshot.analysisSource || "");
    setSourceZipForExport(lastAppliedSnapshot.sourceZipForExport || null);
    setAttachmentValidation(lastAppliedSnapshot.attachmentValidation || null);
    setPreviewFileKey(lastAppliedSnapshot.previewFileKey || "contracts");
    setPendingFixPlan(null);
    setLastAppliedSnapshot(null);
    setError("");
    setSuccessMessage("Última correção desfeita.");
    scrollToStep("review");
  }

  async function applySafeAutoFixes(issue = null) {
    if (!analysis?.dataset) {
      setError("Execute a pré-análise antes de aplicar correções automáticas.");
      return;
    }

    if (issue) {
      const issueFix = buildIssueAutoFixDataset(analysis.dataset, issue, issues);
      await executePreparedFixPlan({
        label: `Correção automática para ${issue.code}`,
        dataset: issueFix.dataset,
        changes: issueFix.changes,
      });
      return;
    }

    const batchFix = buildBatchAutoFixDataset(analysis.dataset, issues);
    await executePreparedFixPlan({
      label: "Correções automáticas seguras",
      dataset: batchFix.dataset,
      changes: batchFix.changes,
    });
  }

  async function applyAttachmentSuggestion(suggestion) {
    if (!suggestion?.dataset) {
      return;
    }
    await executePreparedFixPlan({
      label: `Sugestão de anexo ${suggestion.issue?.code || "ajuste"}`,
      dataset: suggestion.dataset,
      changes: suggestion.changes,
    });
  }

  async function downloadTemplate() {
    setError("");
    try {
      const response = await fetch(`${API_PREFIX}/unified/template`);
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const blob = await response.blob();
      downloadBlob(blob, "base única.csv");
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    }
  }

  async function downloadAttachmentsTemplate() {
    setError("");
    try {
      const response = await fetch(`${API_PREFIX}/attachments/template`);
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const blob = await response.blob();
      downloadBlob(blob, "modelo-anexos-clid.zip");
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    }
  }

  async function exportPackage() {
    if (!analysis?.dataset) {
      setError("Execute a pré-análise antes de exportar.");
      return;
    }
    if (blockingReadinessItems.length > 0) {
      const pending = blockingReadinessItems.map((item) => item.label).join("; ");
      setError(`Pendências críticas antes da exportação: ${pending}`);
      return;
    }

    const attachmentsSource = analysisSource === "zip" ? sourceZipForExport : uploadedAttachmentsSource;
    const hasDocuments =
      (analysis.dataset.contract_documents?.length || 0) > 0 ||
      (analysis.dataset.contract_content_documents?.length || 0) > 0;

    if (analysisSource !== "zip" && hasDocuments && !attachmentsSource) {
      setError(
        "Passo 2: selecione as pastas Documentos contratos/ e/ou Documentos CLID/ para gerar pacote importável no Ariba."
      );
      return;
    }

    setLoading(true);
    setPackageExportStatus("generating");
    setPackageExportDownloadedBytes(0);
    setPackageExportTotalBytes(null);
    setError("");
    try {
      const form = new FormData();
      form.append("dataset_json", JSON.stringify(analysis.dataset));
      form.append("include_report_json", "true");
      form.append("report_json", JSON.stringify(analysis.report));
      if (analysis.run_id) {
        form.append("run_id", analysis.run_id);
      }
      appendAttachmentsPayload(form, attachmentsSource);

      const response = await fetch(`${API_PREFIX}/export/package-with-attachments`, {
        method: "POST",
        body: form,
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      setPackageExportStatus("downloading");
      const blob = await readResponseBlobWithProgress(response, ({ downloadedBytes, totalBytes }) => {
        setPackageExportDownloadedBytes(downloadedBytes);
        setPackageExportTotalBytes(totalBytes);
      });
      setPackageExportStatus("finalizing");
      downloadBlob(blob, `ariba-package-importable-${formatDateSuffix()}.zip`);
      setSuccessMessage("Pacote Ariba gerado com sucesso.");
      loadExecutionRuns();
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    } finally {
      setPackageExportStatus("idle");
      setPackageExportDownloadedBytes(0);
      setPackageExportTotalBytes(null);
      setLoading(false);
    }
  }

  function exportReportJson() {
    if (!analysis?.report) {
      setError("Nenhum relatório para baixar.");
      return;
    }
    const payload = JSON.stringify(analysis.report, null, 2);
    const blob = new Blob([payload], { type: "application/json" });
    downloadBlob(blob, `validation-report-${formatDateSuffix()}.json`);
  }

  async function exportReportXlsx() {
    if (!analysis?.report) {
      setError("Nenhum relatório para exportar em Excel.");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_PREFIX}/export/report.xlsx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          report: analysis.report,
          executive_summary: analysis.executive_summary || null,
          run_id: analysis.run_id || null,
        }),
      });
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      const blob = await response.blob();
      downloadBlob(blob, `validation-report-${formatDateSuffix()}.xlsx`);
      loadExecutionRuns();
    } catch (requestError) {
      setError(normalizeUiError(requestError.message));
    } finally {
      setLoading(false);
    }
  }

  const sourceStepNumber = stepIndexById.get("source") || 1;
  const attachmentsStepNumber = stepIndexById.get("attachments") || 2;
  const rulesStepNumber = stepIndexById.get("rules") || 3;
  const analyzeStepNumber = stepIndexById.get("analyze") || 4;
  const reviewStepNumber = stepIndexById.get("review") || 5;
  const reviewTabs = [
    { id: "summary", label: "Resumo" },
    { id: "issues", label: `Inconsistências (${issues.length})` },
    { id: "attachments", label: "Anexos" },
    { id: "history", label: "Histórico" },
  ];
  const packageExportPercent =
    packageExportTotalBytes && packageExportTotalBytes > 0
      ? Math.min(100, Math.round((packageExportDownloadedBytes / packageExportTotalBytes) * 100))
      : null;
  const packageExportButtonLabel =
    packageExportStatus === "generating"
      ? "Gerando pacote..."
      : packageExportStatus === "downloading"
        ? packageExportPercent !== null
          ? `Baixando pacote... ${packageExportPercent}%`
          : "Baixando pacote..."
        : packageExportStatus === "finalizing"
          ? "Finalizando download..."
          : "Gerar pacote Ariba ZIP";
  const packageExportProgressLabel =
    packageExportStatus === "downloading"
      ? packageExportPercent !== null
        ? `${packageExportPercent}% (${formatBytes(packageExportDownloadedBytes)} de ${formatBytes(packageExportTotalBytes)})`
        : `${formatBytes(packageExportDownloadedBytes)} recebidos`
      : packageExportStatus === "generating"
        ? "Preparando arquivo para download..."
        : packageExportStatus === "finalizing"
          ? "Finalizando arquivo..."
          : "";

  return (
    <div className="page density-compact">
      <header className="hero">
        <div className="heroTop">
          <img src={stratesysLogo} alt="Stratesys" className="brandLogo" />
        </div>
        <div>
          <h1>Assistente de Carga - Contratos Legados SAP Ariba</h1>
          <p>
            Fluxo guiado para receber dados + anexos, validar amarrações, reduzir erros e gerar o
            pacote final de importação no Ariba.
          </p>
        </div>
      </header>

      <section className="panel wizardProgress">
        <div className="progressTop">
          <h2>Progresso do processo</h2>
          <strong>{progressPercent}% concluído</strong>
        </div>
        <div className="progressTrack">
          <div className="progressFill" style={{ width: `${progressPercent}%` }} />
        </div>
        <div className="progressSteps">
          {steps.map((step, index) => (
            <button
              key={step.id}
              className={[
                "progressStepButton",
                step.ready ? "done" : "pending",
                currentStepId === step.id ? "current" : "",
              ].join(" ")}
              type="button"
              onClick={() => scrollToStep(step.id)}
            >
              <span>{index + 1}</span>
              {step.label}
            </button>
          ))}
        </div>
        <div className="nextActionBox">
          <p>
            Próxima ação recomendada: <strong>{nextPendingStep?.label}</strong>
          </p>
          <button className="secondary" type="button" onClick={() => scrollToStep(nextPendingStep.id)}>
            Ir para próxima etapa
          </button>
        </div>
      </section>

      <section ref={sourceRef} className="panel stepPanel">
        <div className="stepHeader">
          <span className="stepNumber">Passo {sourceStepNumber}</span>
          <h2>Escolher fonte dos dados</h2>
          <span className={sourceReady ? "stepStatus done" : "stepStatus pending"}>
            {sourceReady ? "Concluído" : "Pendente"}
          </span>
        </div>

        <div className="wizardModes">
          <button
            className={wizardMode === "zip" ? "modeButton active" : "modeButton"}
            onClick={() => setWizardMode("zip")}
            type="button"
          >
            Já tenho pacote Ariba (.zip)
          </button>
          <button
            className={wizardMode === "unified" ? "modeButton active" : "modeButton"}
            onClick={() => setWizardMode("unified")}
            type="button"
          >
            Base única (CSV/XLSX)
          </button>
          <button
            className={wizardMode === "manual" ? "modeButton active" : "modeButton"}
            onClick={() => setWizardMode("manual")}
            type="button"
          >
            Preenchimento manual
          </button>
        </div>

        {wizardMode === "zip" && (
          <div className="inputBlock">
            <label htmlFor="zipUpload">Pacote Ariba (.zip)</label>
            <input
              id="zipUpload"
              type="file"
              accept=".zip"
              onChange={(event) => setZipFile(event.target.files?.[0] || null)}
            />
            <p className="hint">Ideal para auditar e reprocessar um pacote que já segue o layout Ariba.</p>
          </div>
        )}

        {wizardMode === "unified" && (
          <div className="inputBlock">
            <label htmlFor="unifiedUpload">Arquivo de base única (.csv ou .xlsx)</label>
            <input
              id="unifiedUpload"
              type="file"
              accept=".csv,.xlsx,.xlsm"
              onChange={(event) => setUnifiedFile(event.target.files?.[0] || null)}
            />
            <p className="hint">O sistema converte automaticamente para os templates do Ariba.</p>
            <div className="inlineActions">
              <button className="secondary" onClick={downloadTemplate} type="button">
                Baixar template base única
              </button>
            </div>
          </div>
        )}

        {wizardMode === "manual" && (
          <div className="inputBlock">
            <p className="hint">
              Preencha linhas da base única. Cada linha pode representar contrato + time + documento +
              CLID.
            </p>
            <div className="tableWrap">
              <table>
                <thead>
                  <tr>
                    {MANUAL_COLUMNS.map((column) => (
                      <th key={column}>{column}</th>
                    ))}
                    <th>Ações</th>
                  </tr>
                </thead>
                <tbody>
                  {manualRows.map((row, index) => (
                    <tr key={`manual-row-${index}`}>
                      {MANUAL_COLUMNS.map((column) => (
                        <td key={`${column}-${index}`}>
                          <input
                            value={row[column] || ""}
                            onChange={(event) =>
                              handleManualChange(index, column, event.target.value)
                            }
                          />
                        </td>
                      ))}
                      <td>
                        <button className="danger" onClick={() => removeManualRow(index)} type="button">
                          Remover
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="manualActions">
              <button className="secondary" onClick={addManualRow} type="button">
                Adicionar linha
              </button>
              <span className="hint">Linhas completas: {manualRowsCompleted}</span>
              {manualRowsStarted > 0 && manualRowsIncomplete > 0 && (
                <span className="hint">Linhas incompletas: {manualRowsIncomplete}</span>
              )}
            </div>
          </div>
        )}
      </section>

      {wizardMode !== "zip" && (
        <section ref={attachmentsRef} className="panel stepPanel">
          <div className="stepHeader">
            <span className="stepNumber">Passo {attachmentsStepNumber}</span>
            <h2>Selecionar anexos e CLID</h2>
            <span className={attachmentsReady ? "stepStatus done" : "stepStatus pending"}>
              {attachmentsReady ? "Concluído" : "Pendente"}
            </span>
          </div>
          <div className="inputBlock attachmentsBlock">
            <p className="hint">
              Selecione as pastas <code>Documentos contratos/</code> e <code>Documentos CLID/</code>.
              O sistema reempacota o ZIP final do Ariba com essas duas pastas na raiz, sem estruturas
              extras.
            </p>
            <label htmlFor="contractAttachmentsUpload">Pasta Documentos contratos/</label>
            <input
              id="contractAttachmentsUpload"
              ref={contractAttachmentsInputRef}
              type="file"
              multiple
              onChange={handleContractAttachmentsChange}
            />
            <span className="hint">{contractAttachmentFiles.length} arquivo(s) selecionado(s).</span>
            <label htmlFor="clidAttachmentsUpload">Pasta Documentos CLID/</label>
            <input
              id="clidAttachmentsUpload"
              ref={clidAttachmentsInputRef}
              type="file"
              multiple
              accept=".xls,.xlsx,.xlsm,.csv"
              onChange={handleClidAttachmentsChange}
            />
            <span className="hint">{clidAttachmentFiles.length} arquivo(s) selecionado(s).</span>
            <div className="inlineActions">
              <button className="secondary" type="button" onClick={downloadAttachmentsTemplate}>
                Baixar modelo das pastas
              </button>
              <button
                className="secondary"
                type="button"
                onClick={runAttachmentValidation}
                disabled={!analysis || !uploadedAttachmentsSource || attachmentValidationLoading}
              >
                {attachmentValidationLoading ? "Validando anexos..." : "Validar anexos selecionados"}
              </button>
            </div>
          </div>
        </section>
      )}

      <section ref={rulesRef} className="panel stepPanel rulesPanel">
        <div className="stepHeader">
          <span className="stepNumber">Passo {rulesStepNumber}</span>
          <h2>Configurar regras de validação e parâmetros do cliente</h2>
          <span className="stepStatus done">Concluído</span>
        </div>

        <details className="profilePanel">
          <summary>Perfis por cliente</summary>
          <div className="profilePanelBody">
            <p className="hint">Salve combinações de regras + parâmetros para reutilizar.</p>
            <div className="profileGrid">
              <label>
                Nome do perfil
                <input
                  value={profileNameInput}
                  onChange={(event) => setProfileNameInput(event.target.value)}
                  placeholder="Ex.: Nome do Cliente"
                />
              </label>
              <label>
                Perfis salvos
                <select
                  value={selectedProfileName}
                  onChange={(event) => setSelectedProfileName(event.target.value)}
                  disabled={profilesLoading || profiles.length === 0}
                >
                  <option value="">Selecione</option>
                  {profiles.map((profile) => (
                    <option key={profile.name} value={profile.name}>
                      {profile.name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <div className="inlineActions">
              <button onClick={saveCurrentProfile} type="button" disabled={loading}>
                Salvar perfil
              </button>
              <details className="actionsMenu">
                <summary>Mais ações</summary>
                <div className="actionsMenuList">
                  <button className="secondary" onClick={applyProfileFromSelection} type="button">
                    Carregar perfil
                  </button>
                  <button className="secondary" onClick={deleteSelectedProfile} type="button" disabled={loading}>
                    Excluir perfil
                  </button>
                  <button className="secondary" onClick={loadProfiles} type="button" disabled={profilesLoading}>
                    Atualizar lista
                  </button>
                </div>
              </details>
            </div>
          </div>
        </details>

        <div className="rulesGrid">
          <label title="Lista de status válidos para ContractStatus.">
            Status permitidos (CSV)
            <input
              value={toCsvList(validationRules.allowed_contract_statuses)}
              onChange={(event) =>
                updateRuleField("allowed_contract_statuses", parseCsvList(event.target.value))
              }
            />
          </label>
          <label title="Idiomas válidos para BaseLanguage no cliente.">
            Idiomas permitidos (CSV)
            <input
              value={toCsvList(validationRules.allowed_base_languages)}
              onChange={(event) =>
                updateRuleField("allowed_base_languages", parseCsvList(event.target.value))
              }
            />
          </label>
          <label title="Grupos obrigatórios em ContractTeams.csv por contrato.">
            Grupos obrigatórios no time (CSV)
            <input
              value={toCsvList(validationRules.required_team_project_groups)}
              onChange={(event) =>
                updateRuleField("required_team_project_groups", parseCsvList(event.target.value))
              }
            />
          </label>
          <label title="Como tratar contratos sem time durante a validação.">
            Severidade sem time
            <select
              value={validationRules.missing_team_severity}
              onChange={(event) => updateRuleField("missing_team_severity", event.target.value)}
            >
              <option value="error">Erro</option>
              <option value="warning">Aviso</option>
              <option value="info">Informativo</option>
            </select>
          </label>
        </div>

        <div className="ruleHint">
          <p>
            Formato esperado para <strong>Supplier/AffectedParties</strong>: prefixo <code>sap:</code>{" "}
            + números. Exemplo: <code>sap:0000381965</code>.
          </p>
        </div>

        <details className="advancedRules">
          <summary>Configuração avançada (opcional)</summary>
          <div className="rulesGrid advancedGrid">
            <label>
              Expressão regular ContractId
              <input
                value={validationRules.contract_id_regex || ""}
                onChange={(event) => updateRuleField("contract_id_regex", event.target.value)}
              />
            </label>
            <label>
              Expressão regular Supplier/AffectedParties
              <input
                value={validationRules.sap_party_regex || ""}
                onChange={(event) => updateRuleField("sap_party_regex", event.target.value)}
              />
            </label>
          </div>
          <div className="importParamsPanel">
            <div className="importParamsHeader">
              <h3>Parâmetros do arquivo ImportProjectsParameters.csv</h3>
              <button className="secondary" onClick={restoreDefaultImportParameters} type="button">
                Restaurar padrão
              </button>
            </div>
            <p className="hint">
              Estes campos serão usados para montar o arquivo <code>ImportProjectsParameters.csv</code>.
            </p>
            <div className="importParamsGrid">
              {IMPORT_PARAMETERS_FIELDS.map((field) => (
                <label key={field.key}>
                  {field.label}
                  <input
                    value={importParameters[field.key] || ""}
                    onChange={(event) => updateImportParameterField(field.key, event.target.value)}
                  />
                </label>
              ))}
            </div>
          </div>
        </details>

        <details className="dictionaryPanel">
          <summary>Dicionário de campos (ajuda rápida)</summary>
          <div className="dictionaryGrid">
            {FIELD_DICTIONARY.map((item) => (
              <article key={item.field} className="dictionaryCard">
                <h4>{item.field}</h4>
                <p>{item.description}</p>
                <code>{item.example}</code>
              </article>
            ))}
          </div>
        </details>
      </section>

      <section ref={analyzeRef} className="panel stepPanel">
        <div className="stepHeader">
          <span className="stepNumber">Passo {analyzeStepNumber}</span>
          <h2>Executar pré-análise</h2>
          <span className={analysisReady ? "stepStatus done" : "stepStatus pending"}>
            {analysisReady ? "Concluído" : "Pendente"}
          </span>
        </div>
        <div className="inputBlock">
          <button onClick={executeAnalysis} disabled={loading || !sourceReady} type="button">
            {loading ? "Executando análise..." : "Executar pré-análise"}
          </button>
        </div>
        {error && <div className="errorBox">{error}</div>}
        {successMessage && <div className="successBox">{successMessage}</div>}
      </section>

      {analysis && (
        <section ref={reviewRef} className="panel report stepPanel">
          <div className="stepHeader">
            <span className="stepNumber">Passo {reviewStepNumber}</span>
            <h2>Revisar laudo e gerar pacote final Ariba</h2>
            <span className={readyToImport ? "stepStatus done" : "stepStatus pending"}>
              {readyToImport ? "Pronto para exportar" : "Com pendências"}
            </span>
          </div>
          <p className="hint stepFlowHint">
            {wizardMode === "zip"
              ? "No modo 'Já tenho pacote Ariba (.zip)', esta etapa aparece como Passo 4 porque o passo de anexos já está embutido no pacote."
              : "No modo 'Base única' ou 'Preenchimento manual', esta etapa aparece como Passo 5 porque existe um passo dedicado para anexos."}
          </p>

          <div className="reviewTabs">
            {reviewTabs.map((tab) => (
              <button
                key={tab.id}
                type="button"
                className={reviewTab === tab.id ? "reviewTabButton active" : "reviewTabButton"}
                onClick={() => setReviewTab(tab.id)}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {reviewTab === "summary" && (
            <>
              <section className="finalReadiness">
            <div className={readyToImport ? "readyBadge ok" : "readyBadge pending"}>
              {readyToImport ? "Pronto para importar no Ariba" : "Pendências antes da importação"}
            </div>
            <div className="finalCounters">
              <article className="summaryCard compact">
                <h3>Contratos</h3>
                <p>{contractsTotal}</p>
              </article>
              <article className="summaryCard compact">
                <h3>Anexos de contrato</h3>
                <p>{contractDocumentsTotal}</p>
              </article>
              <article className="summaryCard compact">
                <h3>Documentos CLID</h3>
                <p>{clidDocumentsTotal}</p>
              </article>
            </div>
            <div className="checklistToolbar">
              <button
                className="secondary"
                type="button"
                onClick={() => setShowOnlyCriticalPending((prev) => !prev)}
              >
                {showOnlyCriticalPending ? "Mostrar tudo" : "Mostrar só pendências críticas"}
              </button>
            </div>
            <ul className="checklist">
              {visibleReadinessItems.length === 0 && (
                <li className="ok">
                  <strong>OK:</strong> Nenhuma pendência crítica no momento.
                </li>
              )}
              {visibleReadinessItems.map((item) => (
                <li
                  key={item.key || item.label}
                  className={[
                    item.ok ? "ok" : "pending",
                    !item.ok && item.critical ? "criticalPending" : "",
                    !item.ok && !item.critical ? "nonCriticalPending" : "",
                  ]
                    .join(" ")
                    .trim()}
                >
                  <strong>{item.ok ? "OK" : item.critical ? "Pendente crítico" : "Pendente"}:</strong>{" "}
                  {item.label}
                  {!item.ok && item.actionKey && (
                    <button
                      type="button"
                      className="inlineLinkAction"
                      onClick={() => {
                        if (item.actionKey === "goToAttachmentsStep") {
                          scrollToStep("attachments");
                          return;
                        }
                        setReviewTab("attachments");
                        scrollToStep("review");
                      }}
                    >
                      {item.actionKey === "goToAttachmentsStep"
                        ? "Ir para Passo 2 (Selecionar anexos e CLID)"
                        : "Abrir aba Anexos"}
                    </button>
                  )}
                </li>
              ))}
            </ul>
              </section>

              {executiveSummary && (
                <section className="executivePanel">
              <h3>Resumo executivo</h3>
              <div className="summaryGrid">
                <article className="summaryCard compact">
                  <h3>Prontos para importar</h3>
                  <p>{executiveSummary.contracts_ready_for_import}</p>
                </article>
                <article className="summaryCard compact">
                  <h3>Com erros</h3>
                  <p>{executiveSummary.contracts_with_errors}</p>
                </article>
                <article className="summaryCard compact">
                  <h3>Com avisos</h3>
                  <p>{executiveSummary.contracts_with_warnings}</p>
                </article>
                <article className="summaryCard compact">
                  <h3>Prontidão (%)</h3>
                  <p>{executiveSummary.readiness_percent}</p>
                </article>
              </div>
              <p className="hint">
                <strong>Recomendação:</strong> {executiveSummary.recommendation}
              </p>
                </section>
              )}

              <section className="previewPanel">
            <h3>Pré-visualização do pacote final</h3>
            <p className="hint">Revise os arquivos que serão gerados antes de exportar. Até 5 linhas.</p>
            <div className="previewTree">
              {previewOptions.map((option) => (
                <button
                  key={option.key}
                  type="button"
                  className={previewFileKey === option.key ? "previewTab active" : "previewTab"}
                  onClick={() => setPreviewFileKey(option.key)}
                >
                  {option.label} ({option.rows.length})
                </button>
              ))}
            </div>
            <div className="tableWrap compactTable">
              <table>
                <thead>
                  <tr>
                    {selectedPreviewHeaders.map((header) => (
                      <th key={header}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {selectedPreviewRows.length === 0 && (
                    <tr>
                      <td colSpan={selectedPreviewHeaders.length || 1} className="centered">
                        Sem registros neste arquivo.
                      </td>
                    </tr>
                  )}
                  {selectedPreviewRows.map((row, index) => (
                    <tr key={`${selectedPreview.key}-${index}`}>
                      {selectedPreviewHeaders.map((header) => (
                        <td key={`${header}-${index}`}>{row?.[header] || ""}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
              </section>

              <div className="reportHeader">
            <div>
              <p className="hint">
                Exportação:{" "}
                {analysisSource === "zip"
                  ? "anexos serão preservados do pacote Ariba original."
                  : "selecione as pastas de anexos/CLID para gerar pacote 100% importável no Ariba."}
              </p>
            </div>
            <div className="actions">
              <div className="packageExportAction">
                <button onClick={exportPackage} disabled={loading} type="button">
                  {packageExportButtonLabel}
                </button>
                {packageExportStatus !== "idle" && (
                  <p className="hint downloadProgressHint">{packageExportProgressLabel}</p>
                )}
              </div>
              <details className="actionsMenu">
                <summary>Mais ações</summary>
                <div className="actionsMenuList">
                  <button className="secondary" onClick={exportReportJson} type="button">
                    Baixar relatório JSON
                  </button>
                  <button className="secondary" onClick={exportReportXlsx} type="button">
                    Baixar relatório Excel
                  </button>
                </div>
              </details>
            </div>
              </div>

              <div className="summaryGrid">
            <article className={summary?.is_valid ? "summaryCard valid" : "summaryCard invalid"}>
              <h3>Status</h3>
              <p>{summary?.is_valid ? "Apto para carga" : "Requer ajustes"}</p>
            </article>
            <article className="summaryCard">
              <h3>Erros</h3>
              <p>{summary?.errors ?? 0}</p>
            </article>
            <article className="summaryCard">
              <h3>Avisos</h3>
              <p>{summary?.warnings ?? 0}</p>
            </article>
            <article className="summaryCard">
              <h3>Informativos</h3>
              <p>{summary?.infos ?? 0}</p>
            </article>
              </div>

              <div className="summaryGrid counts">
            <article className="summaryCard compact">
              <h3>Registros em Contracts.csv</h3>
              <p>{recordCounts.contracts || 0}</p>
            </article>
            <article className="summaryCard compact">
              <h3>Registros em ContractDocuments.csv</h3>
              <p>{recordCounts.contract_documents || 0}</p>
            </article>
            <article className="summaryCard compact">
              <h3>Registros em ContractContentDocuments.csv</h3>
              <p>{recordCounts.contract_content_documents || 0}</p>
            </article>
            <article className="summaryCard compact">
              <h3>Registros em ContractTeams.csv</h3>
              <p>{recordCounts.contract_teams || 0}</p>
            </article>
            <article className="summaryCard compact">
              <h3>Registros em ImportProjectsParameters.csv</h3>
              <p>{recordCounts.import_projects_parameters || 0}</p>
            </article>
              </div>
            </>
          )}

          {reviewTab === "issues" && (
            <>
              <div className="issuesHeader">
                <h3>Inconsistências ({issues.length})</h3>
                <p>
                  {groupedIssues.error.length} erros | {groupedIssues.warning.length} avisos |{" "}
                  {groupedIssues.info.length} informativos
                </p>
              </div>

              <div className="inlineActions">
                <button className="secondary" type="button" onClick={() => applySafeAutoFixes()} disabled={loading}>
                  Aplicar correções automáticas (seguras)
                </button>
                <button className="secondary" type="button" onClick={exportReportXlsx} disabled={loading}>
                  Baixar relatório Excel
                </button>
                {lastAppliedSnapshot && (
                  <button className="secondary" type="button" onClick={undoLastFixPlan} disabled={loading}>
                    Desfazer última correção
                  </button>
                )}
              </div>

              {pendingFixPlan && (
                <section className="fixPreviewPanel">
                  <h3>Prévia antes de aplicar</h3>
                  <p className="hint">{pendingFixPlan.label}</p>
                  <ul className="checklist">
                    {(pendingFixPlan.changes || []).slice(0, 12).map((change, index) => (
                      <li key={`pending-change-${index}`} className="ok">
                        {change}
                      </li>
                    ))}
                    {(pendingFixPlan.changes || []).length > 12 && (
                      <li className="pending">
                        + {(pendingFixPlan.changes || []).length - 12} ajuste(s) adicional(is)
                      </li>
                    )}
                  </ul>
                  <div className="inlineActions">
                    <button type="button" onClick={applyPendingFixPlan} disabled={loading}>
                      Aplicar correções da prévia
                    </button>
                    <button className="secondary" type="button" onClick={cancelPendingFixPlan} disabled={loading}>
                      Cancelar
                    </button>
                  </div>
                </section>
              )}

              <div className="filtersRow">
                <label>
                  Severidade
                  <select value={severityFilter} onChange={(event) => setSeverityFilter(event.target.value)}>
                    <option value="all">Todos</option>
                    <option value="error">Erro</option>
                    <option value="warning">Aviso</option>
                    <option value="info">Informativo</option>
                  </select>
                </label>
                <label>
                  Contrato
                  <input
                    placeholder="Filtrar por ContractId"
                    value={contractFilter}
                    onChange={(event) => setContractFilter(event.target.value)}
                  />
                </label>
              </div>

              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Severidade</th>
                      <th>Código</th>
                      <th>Mensagem</th>
                      <th>Como corrigir</th>
                      <th>Arquivo</th>
                      <th>Linha</th>
                      <th>Campo</th>
                      <th>Contrato</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredIssues.length === 0 && (
                      <tr>
                        <td colSpan={8} className="centered">
                          Nenhuma inconsistência encontrada para o filtro selecionado.
                        </td>
                      </tr>
                    )}
                    {filteredIssues.map((issue, index) => (
                      <tr key={`issue-${index}`}>
                        <td>
                          <span className={`badge ${issue.severity}`}>{formatSeverityLabel(issue.severity)}</span>
                        </td>
                        <td>{issue.code}</td>
                        <td>{issue.message}</td>
                        <td>
                          {getIssueGuidance(issue)}
                          {AUTO_FIXABLE_CODES.has(issue.code) && (
                            <div>
                              <button
                                className="tinyAction"
                                type="button"
                                onClick={() => applySafeAutoFixes(issue)}
                                disabled={loading}
                              >
                                Corrigir automaticamente
                              </button>
                            </div>
                          )}
                        </td>
                        <td>{issue.source_file || "-"}</td>
                        <td>{issue.row || "-"}</td>
                        <td>{issue.field || "-"}</td>
                        <td>{issue.contract_id || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <h3 className="contractTitle">Principais inconsistências por contrato</h3>
              <div className="tableWrap compactTable">
                <table>
                  <thead>
                    <tr>
                      <th>Contrato</th>
                      <th>Erros</th>
                      <th>Avisos</th>
                      <th>Informativos</th>
                      <th>Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {issuesByContract.length === 0 && (
                      <tr>
                        <td colSpan={5} className="centered">
                          Sem inconsistências por contrato.
                        </td>
                      </tr>
                    )}
                    {issuesByContract.map((item) => (
                      <tr key={`contract-row-${item.contract}`}>
                        <td>{item.contract}</td>
                        <td>{item.error}</td>
                        <td>{item.warning}</td>
                        <td>{item.info}</td>
                        <td>{item.total}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {reviewTab === "attachments" && (
            <>
              <section className="attachmentsValidationPanel">
                <div className="sectionTitleRow">
                  <h3>Validação inteligente de anexos</h3>
                  <button
                    className="secondary"
                    type="button"
                    onClick={runAttachmentValidation}
                    disabled={
                      attachmentValidationLoading ||
                      (analysisSource === "zip" ? !(sourceZipForExport || zipFile) : !uploadedAttachmentsSource)
                    }
                  >
                    {attachmentValidationLoading ? "Validando..." : "Revalidar anexos"}
                  </button>
                </div>
                {!attachmentValidation && (
                  <p className="hint">
                    Execute a validação para detectar arquivos faltantes, extras e extensões fora do padrão.
                  </p>
                )}
                {attachmentValidation && (
                  <div className="summaryGrid">
                    <article className="summaryCard compact">
                      <h3>Erros</h3>
                      <p>{attachmentValidation.summary?.errors || 0}</p>
                    </article>
                    <article className="summaryCard compact">
                      <h3>Avisos</h3>
                      <p>{attachmentValidation.summary?.warnings || 0}</p>
                    </article>
                    <article className="summaryCard compact">
                      <h3>Referenciados</h3>
                      <p>{attachmentValidation.stats?.referenced_files || 0}</p>
                    </article>
                    <article className="summaryCard compact">
                      <h3>No ZIP</h3>
                      <p>{attachmentValidation.stats?.files_in_zip || 0}</p>
                    </article>
                  </div>
                )}
              </section>

              {attachmentFixSuggestions.length > 0 && (
                <section className="attachmentsSuggestionsPanel">
                  <h3>Sugestões automáticas para anexos</h3>
                  <p className="hint">
                    Correções sugeridas com base em ContractId e arquivos extras detectados no pacote.
                  </p>
                  <div className="suggestionList">
                    {attachmentFixSuggestions.slice(0, 8).map((suggestion) => (
                      <article className="suggestionCard" key={suggestion.id}>
                        <p>
                          <strong>{suggestion.issue.code}</strong>: {suggestion.issue.message}
                        </p>
                        <p className="hint">{(suggestion.changes || []).slice(0, 2).join(" | ")}</p>
                        <button
                          className="secondary"
                          type="button"
                          onClick={() => applyAttachmentSuggestion(suggestion)}
                          disabled={loading}
                        >
                          Aplicar sugestão
                        </button>
                      </article>
                    ))}
                  </div>
                </section>
              )}
            </>
          )}

          {reviewTab === "history" && (
            <section className="historyPanel">
              <div className="sectionTitleRow">
                <h3>Histórico de execuções</h3>
                <button className="secondary" type="button" onClick={() => loadExecutionRuns()} disabled={runsLoading}>
                  {runsLoading ? "Atualizando..." : "Atualizar histórico"}
                </button>
              </div>
              <p className="hint">Últimas análises executadas com status e artefatos gerados.</p>
              <div className="tableWrap compactTable">
                <table>
                  <thead>
                    <tr>
                      <th>Data</th>
                      <th>Origem</th>
                      <th>Perfil</th>
                      <th>Erros</th>
                      <th>Avisos</th>
                      <th>Infos</th>
                      <th>Prontidão</th>
                      <th>Artefatos</th>
                    </tr>
                  </thead>
                  <tbody>
                    {executionRuns.length === 0 && (
                      <tr>
                        <td colSpan={8} className="centered">
                          Sem execuções registradas.
                        </td>
                      </tr>
                    )}
                    {executionRuns.map((run) => (
                      <tr key={run.run_id}>
                        <td>{formatDateTimeLabel(run.created_at)}</td>
                        <td>{run.source || "-"}</td>
                        <td>{run.profile_name || "-"}</td>
                        <td>{run.summary?.errors ?? 0}</td>
                        <td>{run.summary?.warnings ?? 0}</td>
                        <td>{run.summary?.infos ?? 0}</td>
                        <td>{run.readiness_percent ?? 0}%</td>
                        <td>
                          {run.artifacts?.package_with_attachments_zip ? "ZIP" : "-"} /{" "}
                          {run.artifacts?.report_xlsx ? "XLSX" : "-"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </section>
      )}
    </div>
  );
}
