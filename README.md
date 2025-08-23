# Excel → Markdown App Extractor (Multi-Source)

Generate a clean **Application Summary** Markdown doc from **multiple Excel workbooks** using a JSON config that maps **what to show** ↔ **where to find it**.

- Supports **any number of sources** (A–D…).
- Handles **single values**, **multi-value cells**, **multi-row per app**, **grouped sections**, and **upstream/downstream dependencies**.
- New: **server + database inventories** as separate Markdown tables.
- Output is organized in a **timestamped folder** per run.

---

## Contents
- [Prereqs](#prereqs)
- [Files](#files)
- [Quick Start](#quick-start)
- [Config Overview](#config-overview)
- [Aggregators](#aggregators)
- [Examples](#examples)
- [Output Structure](#output-structure)
- [Troubleshooting](#troubleshooting)
- [Convert to DOCX](#convert-to-docx)
- [License](#license)

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

- `extract_app_data_multi_v2.py` — the extractor
- `config.json` — example config (edit this)
- Output: per-run subfolder under `./output/`

---

## Quick Start

1. **Edit the config** to point to your Excel files and column names.
2. **Run the extractor**:
   ```powershell
   python extract_app_data_multi_v2.py `
     --config "C:/path/to/config.json" `
     --app-id 11334
   ```
3. **(Optional) Override source paths** on the CLI:
   ```powershell
   python extract_app_data_multi_v2.py `
     --config "C:/path/to/config.json" `
     --app-id 11334 `
     --source B="C:/data/servers.xlsx" `
     --source C="C:/data/databases.xlsx"
   ```

---

## Config Overview

- **sources**: define your Excel workbooks (path, sheet, id column)
- **fields**: each item describes what to extract and how to render it
- **extra_files**: defines additional outputs (e.g. servers.md, databases.md)

---

## Aggregators

- `unique_join` — flatten values into a comma-separated list
- `group_by` — group one column by another, inline or bulleted
- `dependencies` — find upstream/downstream links by id
- `inventory_summary` — compact summary of Environments, Servers, OS, OS Versions  
  - Configurable keys: `env_column`, `server_column`, `os_name_column`, `os_version_column`  
  - Toggle DB detection with `"show_db_hosts": false`
- `inventory_table` — render a full Markdown table  
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
  "os_version_column": "OS_VERSION",
  "show_db_hosts": false
}
```

### Database Servers by Environment
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

Each run creates a timestamped folder:

```
output/AppName-11334-20250822-213015/
  AppName-11334.md            # main summary
  AppName-11334-servers.md    # server inventory table
  AppName-11334-databases.md  # database inventory table
```

---

## Troubleshooting

- **Column not found**: column headers are case-sensitive; check config
- **Duplicate DB sections**: set `"show_db_hosts": false` in the inventory_summary field
- **_(not found)_**: no data matched that field for this app id
- **Invalid path**: escape backslashes in JSON or use forward slashes
- **Multiple folders created**: ensure only `extract_fields()` writes files (main + extras)

---

## Convert to DOCX

```bash
pandoc output/AppName-11334-20250822-213015/AppName-11334.md -o summary.docx
```

---

## License

MIT (or your team’s standard). Use freely in internal tooling.

---

## 👤 Author
Erick Perales — Cloud Migration IT Architect, Cloud Migration Specialist  
GitHub: [peralese](https://github.com/peralese)
