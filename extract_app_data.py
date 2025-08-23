import argparse, json, sys, os, datetime, re
import pandas as pd
from collections import defaultdict
import re

# --- NEW: normalization + DB detection helpers ---
ENV_NORMALIZE = {
    "prod": "Production",
    "production": "Production",
    "prd": "Production",
    "test": "Test",
    "tst": "Test",
    "uat": "UAT",
    "dev": "Dev",
    "development": "Dev",
}
DB_NAME_PATTERN = re.compile(r'(^|[^a-z])(db|sql|ora)(\d+)?($|[^a-z])', re.IGNORECASE)

def normalize_env(env: str) -> str:
    if pd.isna(env) or not str(env).strip():
        return "Unknown"
    key = str(env).strip().lower()
    return ENV_NORMALIZE.get(key, str(env).strip().title())

def is_db_host(server_name: str) -> bool:
    if pd.isna(server_name) or not str(server_name).strip():
        return False
    return bool(DB_NAME_PATTERN.search(str(server_name)))

# -------------------------------------------------

def load_sheet(path, sheet_name=None):
    xls = pd.ExcelFile(path, engine="openpyxl")
    use_sheet = sheet_name if sheet_name in xls.sheet_names else (sheet_name or xls.sheet_names[0])
    return pd.read_excel(path, sheet_name=use_sheet, engine="openpyxl")

def transform_value(val, transform=None, split=None, join=None):
    if pd.isna(val):
        return ""
    out = str(val)
    if split:
        parts = [p.strip() for p in out.split(split) if p is not None]
        out = (join or ", ").join([p for p in parts if p])
    if transform == "strip":
        out = out.strip()
    elif transform == "upper":
        out = out.upper()
    elif transform == "lower":
        out = out.lower()
    return out

def stable_unique(seq):
    seen = set()
    out = []
    for x in seq:
        if x not in seen and x != "":
            seen.add(x)
            out.append(x)
    return out

def parse_source_overrides(pairs):
    mapping = {}
    if not pairs:
        return mapping
    for p in pairs:
        if "=" not in p:
            raise ValueError(f"Invalid --source override '{p}'. Use ALIAS=/path.xlsx")
        alias, path = p.split("=", 1)
        alias = alias.strip()
        path = path.strip().strip('"').strip("'")
        if not alias:
            raise ValueError(f"Invalid alias in --source '{p}'")
        mapping[alias] = path
    return mapping

def sanitize_filename(name: str) -> str:
    # Replace invalid filename chars with underscore
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()

def extract_fields(cfg, app_id, source_overrides):
    sources_cfg = cfg.get("sources", {})
    if not sources_cfg:
        raise ValueError("Config 'sources' is empty. Define at least one source with path/id default.")

    # Apply CLI overrides
    for alias, override_path in source_overrides.items():
        if alias not in sources_cfg:
            raise ValueError(f"Override provided for unknown source alias '{alias}'. Add it to config 'sources'.")
        sources_cfg[alias]["path"] = override_path

    # Cache DataFrames per (alias, sheet_name)
    cache = {}
    def get_df(alias, sheet_name=None):
        src = sources_cfg.get(alias)
        if not src:
            raise ValueError(f"Unknown source alias '{alias}' in fields.")
        path = src.get("path")
        if not path or not os.path.exists(path):
            raise ValueError(f"Source '{alias}' path not found: {path}")
        default_sheet = src.get("sheet_name_default")
        key = (alias, sheet_name or default_sheet)
        if key in cache:
            return cache[key]
        df = load_sheet(path, key[1])
        cache[key] = df
        return df

    # Get ALL matching rows for an app_id
    def find_rows(alias, sheet_name, id_column, app_id):
        src = sources_cfg[alias]
        the_id_col = id_column or src.get("id_column_default") or "ApplicationID"
        df = get_df(alias, sheet_name or src.get("sheet_name_default"))
        if the_id_col not in df.columns:
            raise ValueError(f"ID column '{the_id_col}' not in source {alias}. Available: {list(df.columns)}")
        matches = df[df[the_id_col].astype(str).str.strip().str.upper() == str(app_id).strip().upper()]
        return matches  # DataFrame (possibly empty)

    # --- NEW: support multiple output files ---
    outputs = defaultdict(list)  # key -> list of lines
    def out_lines(key):          # get writer (defaults to "main")
        return outputs[key or "main"]

    app_name = None

    for fld in cfg.get("fields", []):
        label = fld["label"]
        alias = fld["source"]
        sheet_name = fld.get("sheet_name")
        id_column = fld.get("id_column")
        aggregate = fld.get("aggregate")
        joiner = fld.get("join", ", ")
        emit_key = fld.get("emit_file", "main")  # <-- NEW: which file to write to

        rows = find_rows(alias, sheet_name, id_column, app_id)
        lines = out_lines(emit_key)  # writer for this field

        # Default: simple single-column on first match
        def render_simple():
            if len(rows) == 0:
                return ""
            col = fld["column"]
            if col not in rows.columns:
                raise ValueError(f"Column '{col}' missing for label '{label}' in source {alias}. Available: {list(rows.columns)}")
            raw = rows.iloc[0][col]
            return transform_value(raw, transform=fld.get("transform"), split=fld.get("split"), join=fld.get("join"))

        rendered = ""

        if aggregate == "unique_join":
            col = fld["column"]
            if len(rows) > 0 and col not in rows.columns:
                raise ValueError(f"Column '{col}' missing for label '{label}' in source {alias}. Available: {list(rows.columns)}")
            values = []
            for _, r in rows.iterrows():
                v = r[col] if col in r.index else ""
                v = transform_value(v, transform=fld.get("transform"))
                if v and fld.get("split"):
                    parts = [p.strip() for p in v.split(fld["split"]) if p is not None]
                    values.extend(parts)
                elif v:
                    values.append(str(v).strip())
            values = stable_unique(values)
            rendered = joiner.join(values)

        elif aggregate == "group_by":
            key_col = fld["group_by_column"]
            val_col = fld["value_column"]
            style = fld.get("style", "inline")
            want_unique = fld.get("unique", True)

            for c in [key_col, val_col]:
                if len(rows) > 0 and c not in rows.columns:
                    raise ValueError(f"Column '{c}' missing for label '{label}' in source {alias}. Available: {list(rows.columns)}")

            grouped = {}
            for _, r in rows.iterrows():
                k = r[key_col] if key_col in r.index else ""
                v = r[val_col] if val_col in r.index else ""
                k = "" if pd.isna(k) else str(k).strip()
                v = "" if pd.isna(v) else str(v).strip()
                if k == "" or v == "":
                    continue
                grouped.setdefault(k, []).append(v)

            key_order = fld.get("key_order")
            keys = key_order if key_order else list(grouped.keys())

            if style == "bulleted":
                segs = [f"**{label}:**"]
                for k in keys:
                    vals = grouped.get(k, [])
                    if want_unique:
                        vals = stable_unique(vals)
                    if not vals:
                        continue
                    segs.append(f"- {k}: {joiner.join(vals)}")
                rendered = "\n".join(segs)
            else:
                parts = []
                for k in keys:
                    vals = grouped.get(k, [])
                    if want_unique:
                        vals = stable_unique(vals)
                    if not vals:
                        continue
                    parts.append(f"{k}: {joiner.join(vals)}")
                rendered = f"**{label}:** " + "; ".join(parts) if parts else ""

        elif aggregate == "dependencies":
            match_col = fld["match_column"]
            ret_col = fld["return_column"]
            src = cfg["sources"][alias]
            df = load_sheet(src["path"], src.get("sheet_name_default"))
            if match_col not in df.columns or ret_col not in df.columns:
                raise ValueError(f"Dependency columns '{match_col}'/'{ret_col}' not found in source {alias}. Available: {list(df.columns)}")
            mm = df[df[match_col].astype(str).str.strip().str.upper() == str(app_id).strip().upper()]
            values = []
            for _, r in mm.iterrows():
                v = r[ret_col]
                if pd.isna(v): 
                    continue
                values.append(str(v).strip())
            values = stable_unique(values)
            rendered = joiner.join(values)

        #---- inventory summary table ----
        elif aggregate == "inventory_summary":
        # Column names (override in config if needed)
            env_col = fld.get("env_column", "ENVIRONMENT")
            server_col = fld.get("server_column", "SERVER")
            os_name_col = fld.get("os_name_column", "OS_NAME")
            os_ver_col = fld.get("os_version_column", "OS_VERSION")

            for c in [env_col, server_col, os_name_col, os_ver_col]:
                if len(rows) > 0 and c not in rows.columns:
                    raise ValueError(f"Column '{c}' missing for inventory_summary in source {alias}. Available: {list(rows.columns)}")

            envs = set()
            servers_by_env = defaultdict(set)
            os_names = set()
            os_versions = set()
            db_by_env = defaultdict(set)

            for _, r in rows.iterrows():
                env = normalize_env(r.get(env_col, ""))
                server = "" if pd.isna(r.get(server_col)) else str(r.get(server_col)).strip()
                osn = "" if pd.isna(r.get(os_name_col)) else str(r.get(os_name_col)).strip()
                osv = "" if pd.isna(r.get(os_ver_col)) else str(r.get(os_ver_col)).strip()

                if env != "Unknown":
                    envs.add(env)
                if server:
                    servers_by_env[env].add(server)
                    if is_db_host(server):
                        db_by_env[env].add(server)
                if osn:
                    os_names.add(osn)
                if osv:
                    os_versions.add(osv)

            lines.append(f"**Environment(s):** {', '.join(sorted(envs)) if envs else '_(not found)_'}")

            lines.append("**Servers by Environment:**")
            if servers_by_env:
                for e, svrs in sorted(servers_by_env.items()):
                    if svrs:
                        lines.append(f"- {e}: {', '.join(sorted(svrs))}")
            else:
                lines.append("- _(not found)_")

            lines.append(f"**Operating System:** {', '.join(sorted(os_names)) if os_names else '_(not found)_'}")
            lines.append(f"**OS Version:** {', '.join(sorted(os_versions)) if os_versions else '_(not found)_'}")

            rendered = ""  # we've already appended to output

        # --- NEW: inventory_table (Markdown table) ---
        elif aggregate == "inventory_table":
            cols = fld.get("columns")
            if not cols:
                raise ValueError(f"'inventory_table' for '{label}' requires a 'columns' array in config.")
            headers_map = fld.get("headers", {})
            sort_by = fld.get("sort_by", ["ENVIRONMENT", "SERVER"])
            env_col = fld.get("env_column", "ENVIRONMENT")

            for c in cols:
                if len(rows) > 0 and c not in rows.columns:
                    raise ValueError(f"Column '{c}' missing for inventory_table in source {alias}. Available: {list(rows.columns)}")

            df = rows.copy()
            if env_col in df.columns:
                df[env_col] = df[env_col].apply(normalize_env)

            # Filter out fully-empty rows in requested cols (optional, tidy)
            df = df.dropna(how="all", subset=cols)

            # Safe sort (only keep sort columns that exist)
            sort_by = [c for c in sort_by if c in df.columns]
            if sort_by:
                df = df.sort_values(by=sort_by, kind="stable")

            out_df = df[cols].copy()
            final_headers = [headers_map.get(c, c.title() if c.isupper() else c) for c in cols]

            def to_md_table(df, headers):
                lines_tbl = [
                    "| " + " | ".join(headers) + " |",
                    "| " + " | ".join("---" for _ in headers) + " |"
                ]
                for _, r in df.iterrows():
                    cells = []
                    for c in df.columns:
                        v = r[c]
                        if pd.isna(v): v = ""
                        v = str(v).replace("\r", " ").replace("\n", " ").strip()
                        cells.append(v)
                    lines_tbl.append("| " + " | ".join(cells) + " |")
                return "\n".join(lines_tbl)

            lines.append(f"**{label}:**")
            lines.append(to_md_table(out_df, final_headers))
            lines.append("")
            rendered = ""  # we've appended directly

        else:
            rendered = render_simple()

        # Capture app name for output filename
        if label == cfg.get("app_name_field_label", "Application Name") and rendered:
            app_name = rendered

        # Write to the proper output file
        if aggregate == "group_by" and rendered.startswith("**"):
            lines.append(rendered if rendered else f"**{label}:** _(not found)_")
        elif aggregate not in ("inventory_table", "inventory_summary"):
            lines.append(f"**{label}:** {rendered if rendered else '_(not found)_'}")


    # --- finalize & write all files ---
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    main_title = cfg.get("doc_title_template", "Application Summary — {app_id}").format(app_id=app_id)
    if app_name:
        main_title = f"{app_name} — {app_id}"

    # Prepend headers per file
    for key, lines in outputs.items():
        title_tpl = main_title if key == "main" else cfg.get("extra_files", {}).get(key, {}).get("title_template", f"{key} — {{app_id}}")
        title = title_tpl.format(app_id=app_id, app=(app_name or "App"))
        lines[:0] = [f"# {title}", "", f"_Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}_", ""]
        outputs[key] = lines + [""]

    # Determine filenames
    files_written = []
    # main file (keep your existing behavior)
    safe_app = sanitize_filename(app_name if app_name else "App")
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")

    # Create a dedicated run folder: ./output/<app>-<id>-<ts>/
    run_folder = os.path.join(".", "output", f"{safe_app}-{app_id}-{ts}")
    os.makedirs(run_folder, exist_ok=True)

    # Main file goes inside run_folder
    if cfg.get("main_filename_template"):
        # allow template override, but root it inside run_folder
        fname = cfg["main_filename_template"].format(app_id=app_id, app=safe_app, ts=ts)
        main_out = os.path.join(run_folder, os.path.basename(fname))
    else:
        main_out = os.path.join(run_folder, f"{safe_app}-{app_id}.md")



    # Write main
    main_text = "\n".join(outputs["main"])
    with open(main_out, "w", encoding="utf-8") as f:
        f.write(main_text)
    files_written.append(main_out)

    # Write extras
    for key, meta in cfg.get("extra_files", {}).items():
        if key == "main": 
            continue
        if key not in outputs:
            continue  # nothing emitted
        tpl = meta.get("filename_template", f"{safe_app}-{app_id}-{key}.md")
        fname = tpl.format(app_id=app_id, app=safe_app, ts=ts)
        out_path = os.path.join(run_folder, os.path.basename(fname))
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(outputs[key]))

            files_written.append(out_path)

    # Return main text + app_name (for CLI print)
    return main_text, app_name or ""

def main():
    p = argparse.ArgumentParser(description="Extract app data from multiple Excel workbooks into a Markdown summary.")
    p.add_argument("--config", required=True, help="Path to JSON config defining sources and fields.")
    p.add_argument("--app-id", required=True, help="Application ID to look up.")
    p.add_argument("--out", required=False, help="Output markdown file path. If omitted, auto-generates to ./output/AppName-AppID-YYYYMMDD-HHMMSS.md")
    p.add_argument("--source", action="append", help="Override a source path like A=/path/to/file.xlsx (can repeat).")
    args = p.parse_args()

    try:
        with open(args.config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        overrides = parse_source_overrides(args.source)
        md, app_name = extract_fields(cfg, args.app_id, overrides)

        print("Done.")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
