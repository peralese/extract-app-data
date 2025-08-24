# Excel → Markdown App Extractor (Single or Batch)

Generate clean **Application Summary** Markdown docs from **multiple Excel workbooks** using a JSON config that maps **what to show** ↔ **where to find it**.

- Supports **any number of sources** (A–D…).
- Handles **single values**, **multi-value cells**, **multi-row per app**, **grouped sections**, and **upstream/downstream dependencies**.
- Includes **server / OS inventory summary** and **full inventory tables**.
- **New (2025‑08‑23): Batch mode** — pass a file of App IDs to process many apps in one run.

---

## Contents
- [Prereqs](#prereqs)
- [Files](#files)
- [Quick Start](#quick-start)
- [CLI Options](#cli-options)
- [Config Overview](#config-overview)
- [Aggregators](#aggregators)
- [Examples](#examples)
- [Output Structure](#output-structure)
- [Troubleshooting](#troubleshooting)
- [Convert to DOCX](#convert-to-docx)
- [Changelog](#changelog)
- [License](#license)
- [Author](#-author)

---

## Prereqs

- Python 3.9+ recommended
- Packages:
  ```bash
  pip install pandas openpyxl
  ```
- PowerShell users: you can use backticks (`` ` ``) for line continuations.

---

## Files

- `extract_app_data.py` — the extractor (supports single ID and batch)
- `config.json` — example config (edit this)
- Output: per‑app subfolders under `./output/`

---

## Quick Start

### Single App ID (original behavior)
```powershell
python extract_app_data.py `
  --config "C:/path/to/config.json" `
  --app-id 11334
```

### Batch Mode (CSV/XLSX/TXT)
```powershell
# Auto-detect the column (prefers app_id/application_id/appid; else first column)
python extract_app_data.py `
  --config "C:/path/to/config.json" `
  --ids-file "C:/path/to/app_ids.csv"
```

Specify a column explicitly (CSV/XLSX only):
```powershell
python extract_app_data.py `
  --config "C:/path/to/config.json" `
  --ids-file "C:/path/to/app_ids.xlsx" `
  --ids-col "AppNumber"
```

TXT format: one ID per line **or** comma/semicolon/space‑separated.

### Override Source Paths on the CLI
```powershell
python extract_app_data.py `
  --config "C:/path/to/config.json" `
  --app-id 11334 `
  --source B="C:/data/servers.xlsx" `
  --source C="C:/data/databases.xlsx"
```

---

## CLI Options

```
--app-id <ID>                 Process a single Application ID.
--ids-file <path>             Process multiple IDs from a CSV/XLSX/TXT file.
--ids-col <name>              (Optional) Column name for IDs when using CSV/XLSX.
--config <path>               JSON config defining sources and fields. (Required)
--source ALIAS=path           Override a source path defined in the config. Repeatable.
```

Notes:
- `--app-id` and `--ids-file` are **mutually exclusive** (pick one).
- Batch mode **dedupes** IDs while preserving order.

---

## Config Overview

Your JSON config has three main sections:

- **sources**: define Excel workbooks and defaults
- **fields**: what to extract and how to render it
- **extra_files**: optional additional outputs (e.g., servers.md)

Minimal example:

```json
{
  "app_name_field_label": "Application Name",
  "doc_title_template": "Application Summary — {app_id}",

  "sources": {
    "A": {
      "path": "C:/data/apps.xlsx",
      "sheet_name_default": "Applications",
      "id_column_default": "ApplicationID"
    },
    "B": {
      "path": "C:/data/servers.xlsx",
      "sheet_name_default": "Servers",
      "id_column_default": "ApplicationID"
    }
  },

  "fields": [
    {"label": "Application Name", "source": "A", "column": "AppName"},
    {"label": "Owner", "source": "A", "column": "OwnerName"},
    {
      "label": "Environment / Server / OS Summary",
      "source": "B",
      "aggregate": "inventory_summary",
      "env_column": "ENVIRONMENT",
      "server_column": "SERVER",
      "os_name_column": "OS_NAME",
      "os_version_column": "OS_VERSION"
    },
    {
      "label": "Server Inventory",
      "source": "B",
      "aggregate": "inventory_table",
      "emit_file": "servers",
      "columns": ["SERVER", "ENVIRONMENT", "OS_NAME", "OS_VERSION"],
      "headers": {
        "SERVER": "Server",
        "ENVIRONMENT": "Environment",
        "OS_NAME": "OS Name",
        "OS_VERSION": "OS Version"
      },
      "sort_by": ["ENVIRONMENT", "SERVER"],
      "env_column": "ENVIRONMENT"
    }
  ],

  "extra_files": {
    "servers": {
      "title_template": "{app} — Server Inventory",
      "filename_template": "{app}-{app_id}-servers.md"
    }
  }
}
```

---

## Aggregators

- `unique_join` — flatten values into a comma‑separated list
- `group_by` — group one column by another (inline or bulleted)
- `dependencies` — upstream/downstream lookups via match/return columns
- `inventory_summary` — compact summary of Environments, Servers, OS, OS Versions  
  - Keys: `env_column`, `server_column`, `os_name_column`, `os_version_column`
- `inventory_table` — full Markdown table for a chosen set of columns  
  - Keys: `columns`, `headers`, `sort_by`, `env_column`

---

## Examples

### Inventory Summary
```json
{
  "label": "Environment / Server / OS Summary",
  "source": "B",
  "aggregate": "inventory_summary",
  "env_column": "ENVIRONMENT",
  "server_column": "SERVER",
  "os_name_column": "OS_NAME",
  "os_version_column": "OS_VERSION"
}
```

### Database Servers by Environment (bulleted)
```json
{
  "label": "Database Servers by Environment",
  "source": "C",
  "aggregate": "group_by",
  "group_by_column": "PHASE",
  "value_column": "SERVER",
  "style": "bulleted",
  "key_order": ["Production", "Test", "Dev"],
  "unique": true
}
```

### Server Inventory Table (separate file)
```json
{
  "label": "Server Inventory",
  "source": "B",
  "aggregate": "inventory_table",
  "emit_file": "servers",
  "columns": ["SERVER", "ENVIRONMENT", "OS_NAME", "OS_VERSION"],
  "headers": {
    "SERVER": "Server",
    "ENVIRONMENT": "Environment",
    "OS_NAME": "OS Name",
    "OS_VERSION": "OS Version"
  },
  "sort_by": ["ENVIRONMENT", "SERVER"],
  "env_column": "ENVIRONMENT"
}
```

---

## Output Structure

Each app processed creates its own timestamped folder:

```
output/AppName-11334-20250823-213015/
  AppName-11334.md            # main summary
  AppName-11334-servers.md    # (optional) extra files defined in config.extra_files
```

When running **batch mode**, you’ll see one such folder per App ID.

---

## Troubleshooting

- **Column not found**: column headers are case‑sensitive; check config.
- **_(not found)_** in output: no matching data for that field/App ID.
- **Invalid path**: in JSON, escape backslashes (`\"`) or use forward slashes (`/`).
- **Multiple folders unexpectedly**: ensure only `extract_fields()` writes files (main + extras).
- **Excel engine error**: `openpyxl` must be installed; reinstall with `pip install --upgrade openpyxl`.

---

## Convert to DOCX

```bash
pandoc output/AppName-11334-20250823-213015/AppName-11334.md -o summary.docx
```

---

## Changelog

- **2025‑08‑23**
  - Added **batch mode** via `--ids-file` (CSV/XLSX/TXT) and optional `--ids-col`.
  - Kept **single‑ID** mode unchanged.
  - Preserved per‑app output subfolder naming: `./output/<AppName>-<ID>-<timestamp>/`.
  - Minor docs polish and config examples.

---

## License

MIT (or your team’s standard). Use freely in internal tooling.

---

## 👤 Author

Erick Perales — Cloud Migration IT Architect, Cloud Migration Specialist  
GitHub: [peralese](https://github.com/peralese)
