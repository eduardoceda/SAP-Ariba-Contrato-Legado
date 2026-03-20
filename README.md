# SAP Ariba - Contratos Legados (Pre-Analise + Geracao de Carga)

Sistema full stack para validar e preparar cargas de contratos legados no SAP Ariba Contracts.

## O que o sistema faz

- Analisa pacote Ariba em massa (`.zip`) com:
  - `Contracts.csv`
  - `ContractDocuments.csv`
  - `ContractContentDocuments.csv`
  - `ContractTeams.csv`
  - `ImportProjectsParameters.csv`
- Valida amarracoes e formatos antes da importacao:
  - presenca dos arquivos obrigatorios
  - colunas esperadas
  - campos obrigatorios
  - formato de IDs, datas e numeros
  - referencias cruzadas entre CSVs
  - existencia de anexos nas pastas (`Documentos contratos/` e `Documentos CLID/`)
  - duplicidades
- Permite entrada por base unica:
  - importacao de CSV/XLSX unificado
  - alimentacao manual em tela
- Permite regras configuraveis por cliente (status, idiomas, grupos obrigatorios, regex de IDs).
- Permite salvar perfis por cliente (regras + parametros de importacao) para reaproveitar em novas cargas.
- Exibe fluxo guiado com progresso de etapas e checklist bloqueante antes da exportacao final.
- Gera orientacoes de correcao por inconsistencia e oferece correcao automatica segura para casos comuns.
- Permite correcao automatica por item de inconsistência, com pre-visualizacao antes/depois e botao de desfazer.
- Valida zip de anexos de forma inteligente (faltantes, extras, pasta incorreta e extensoes suspeitas).
- Oferece sugestoes de auto-fix para anexos duplicados/extras com aplicacao em 1 clique.
- Exibe resumo executivo (prontidao por contrato, riscos e recomendacao) junto do laudo tecnico.
- Mostra pre-visualizacao dos arquivos finais do pacote Ariba antes de gerar o ZIP.
- Persiste a sessao local do assistente para retomar trabalho apos recarregar a pagina.
- Mantem historico de execucoes (fonte, perfil, severidades e artefatos exportados).
- Inclui dicionario de campos na tela para facilitar o uso por usuarios sem contexto tecnico.
- Gera pacote final Ariba (`.zip`) com os 5 CSVs e relatorio JSON.
- Exporta laudo em Excel (`.xlsx`) com abas de resumo, inconsistencias e visao por contrato.

## Arquitetura

- Frontend: React + Vite
- Backend: FastAPI (Python)

## Estrutura

- `backend/`: API e motor de validacao
- `frontend/`: interface web unica
- `backend/tests/fixtures/sample_zip/`: amostra sintetica usada nos testes automatizados
- `_sample_zip/`: conteudo local opcional para referencia manual (ignorado no Git)

## Como executar

### Opcao simples (1 comando na pasta raiz)

```powershell
cd "C:\Users\eduar\OneDrive\Documentos\CONSULTOR JEC\PESSOA JURÍDICA\Sistemas\SAP Ariba - Contrato Legado"
npm install
npm run dev
```

Isso sobe backend (`:8001`) e frontend (`:5173`) juntos no mesmo terminal.

### 1) Backend (FastAPI)

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

API: `http://localhost:8000`

Se a porta `8000` estiver ocupada por outro sistema, rode em outra porta (ex.: `8001`):

```powershell
uvicorn app.main:app --reload --host 0.0.0.0 --port 8001
```

### 2) Frontend (React)

```powershell
cd frontend
npm install
npm run dev
```

UI: `http://localhost:5173`

Se backend estiver em porta diferente de `8000`, defina a URL da API antes de subir o front:

```powershell
$env:VITE_API_BASE_URL="http://localhost:8001/api"
npm run dev
```

## Fluxo recomendado

1. Abra a UI no modo **Assistente de Carga**.
2. **Passo 1**: escolha a fonte dos dados (`ZIP Ariba`, `Base unica` ou `Manual`).
3. **Passo 2** (quando aplicavel): envie o zip de anexos com `Documentos contratos/` e `Documentos CLID/`.
4. **Passo 3**: ajuste regras do cliente (status, regex, grupos obrigatorios etc.).
5. **Passo 4**: execute a pre-analise.
6. **Passo 5**: revise o laudo e gere o pacote final importavel no Ariba.

## Endpoints principais

- `GET /api/health`
- `GET /api/rules/default` (regras padrao)
- `POST /api/analyze/upload` (zip Ariba)
- `POST /api/analyze/json` (dataset JSON)
- `POST /api/unified/upload` (CSV/XLSX unificado)
- `POST /api/unified/manual` (linhas manuais)
- `GET /api/profiles` (listar perfis de cliente)
- `POST /api/profiles` (salvar/atualizar perfil de cliente)
- `DELETE /api/profiles/{profile_name}` (excluir perfil de cliente)
- `POST /api/attachments/validate` (validacao inteligente do ZIP de anexos)
- `GET /api/unified/template` (template base unica)
- `POST /api/export/package` (gera zip final)
- `POST /api/export/package-with-attachments` (gera zip importavel no Ariba com anexos)
- `POST /api/export/report.xlsx` (gera laudo Excel)
- `GET /api/runs` (historico de execucoes)

## Testes automatizados

```powershell
cd backend
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
python -m pytest -q
```

## Observacoes

- O relatorio classifica inconsistencias em `error`, `warning` e `info`.
- O pacote pode ser gerado mesmo com avisos, mas erros devem ser tratados antes da carga no Ariba.
- Para endpoints de upload (`/analyze/upload` e `/unified/upload`), envie `validation_rules` no `FormData` como JSON string quando quiser regras customizadas.
- Para gerar pacote realmente importavel no Ariba, use `export/package-with-attachments`:
  - em carga via ZIP, reutilize o pacote original como `attachments_zip`;
  - em carga manual/base unica, envie um zip com as pastas de anexos.
