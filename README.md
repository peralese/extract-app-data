# Excel → Markdown App Extractor (Multi-Source)

Generate a clean **Application Summary** Markdown doc from **multiple Excel workbooks** using a JSON config that maps **what to show** ↔ **where to find it**.

- Supports **any number of sources** (A–D…).
- Handles **single values**, **multi-value cells**, **multi-row per app**, **grouped sections**, and **upstream/downstream dependencies**.
- Output is **Markdown** (easy to diff/convert to DOCX).

---

## Contents
- [Prereqs](#prereqs)
- [Files](#files)
- [Quick Start](#quick-start)
- [Config Overview](#config-overview)
- [Field Recipes](#field-recipes)
- [Dependencies (Upstream/Downstream)](#dependencies-upstreamdownstream)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)
- [Tips](#tips)
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
- `config_from_user_fixed.json` — example config (edit this)
- Output: `app_<APP_ID>.md`

---

## Quick Start

1) **Edit the config** to point to your Excel files and column names.
2) **Run the extractor**:
   ```powershell
   python extract_app_data_multi_v2.py `
     --config "C:/path/to/config_from_user_fixed.json" `
     --app-id 6847 `
     --out "app_6847.md"
   ```
3) **(Optional) Override source paths on the CLI** without editing JSON:
   ```powershell
   python extract_app_data_multi_v2.py `
     --config "C:/path/to/config.json" `
     --app-id 6847 `
     --out "app_6847.md" `
     --source A="C:/data/A.xlsx" `
     --source B="C:/data/B.xlsx"
   ```

---

## Config Overview

- **sources**: define your Excel workbooks (path, sheet_name_default, id_column_default)
- **fields**: each item describes what to extract and how to render it

Aggregators supported:
- `unique_join`
- `group_by`
- `dependencies`

---

## Examples

### Simple
```json
{ "label": "Application Name", "source": "A", "column": "BA_Name" }
```

### Unique Join
```json
{
  "label": "Servers",
  "source": "B",
  "column": "Server",
  "aggregate": "unique_join",
  "join": ", "
}
```

### Group By
```json
{
  "label": "Servers by Environment",
  "source": "B",
  "aggregate": "group_by",
  "group_by_column": "Environment",
  "value_column": "Server",
  "style": "bulleted",
  "key_order": ["Prod", "Test", "Dev"],
  "unique": true,
  "join": ", "
}
```

### Dependencies
```json
{
  "label": "Downstream Dependencies",
  "source": "D",
  "aggregate": "dependencies",
  "match_column": "SEND_ESATS_ID",
  "return_column": "REC_ESATS_ID",
  "join": ", "
}
```

---

## Troubleshooting

- **Invalid \escape**: use forward slashes or double backslashes in JSON paths
- **Unrecognized arguments**: wrap paths with spaces in quotes
- **Column not found**: column headers are case-sensitive
- **Source not found**: check the `sources` path or override with `--source`
- **_(not found)_**: no data matched that field

---

## Convert to DOCX

```bash
pandoc app_6847.md -o app_6847.docx
```

---

## License

MIT (or your team’s standard). Use freely in internal tooling.

## 👤 Author
Erick Perales — Cloud Migration IT Architect, Cloud Migration Specialist  
GitHub: https://github.com/peralese