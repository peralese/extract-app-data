import argparse, json, sys, os, datetime
import pandas as pd
from collections import OrderedDict

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

    # Build doc
    lines = []
    app_name = None

    for fld in cfg.get("fields", []):
        label = fld["label"]
        alias = fld["source"]
        sheet_name = fld.get("sheet_name")
        id_column = fld.get("id_column")
        aggregate = fld.get("aggregate")
        joiner = fld.get("join", ", ")

        rows = find_rows(alias, sheet_name, id_column, app_id)

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
            # Collect values across all rows from one column (with optional split/join/transform)
            col = fld["column"]
            if len(rows) > 0 and col not in rows.columns:
                raise ValueError(f"Column '{col}' missing for label '{label}' in source {alias}. Available: {list(rows.columns)}")
            values = []
            for _, r in rows.iterrows():
                v = r[col] if col in r.index else ""
                v = transform_value(v, transform=fld.get("transform"))  # transform before splitting
                if v and fld.get("split"):
                    parts = [p.strip() for p in v.split(fld["split"]) if p is not None]
                    values.extend(parts)
                elif v:
                    values.append(v.strip())
            values = stable_unique(values)
            rendered = joiner.join(values)

        elif aggregate == "group_by":
            # Group rows by one column and list values from another column
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

            # Optional explicit key order
            key_order = fld.get("key_order")
            keys = key_order if key_order else list(grouped.keys())

            # Format
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
                # inline "Env: a, b; Env2: c"
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
            # special: search by match_column==app_id, return values from return_column
            match_col = fld["match_column"]
            ret_col = fld["return_column"]

            # Look in ALL rows of this source's sheet, not just rows matching by default id_column
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

        else:
            # simple
            rendered = render_simple()

        # Title capture
        if label == cfg.get("app_name_field_label", "Application Name") and rendered:
            app_name = rendered

        # Write to doc
        if aggregate == "group_by" and rendered.startswith("**"):
            lines.append(rendered if rendered else f"**{label}:** _(not found)_")
        else:
            lines.append(f"**{label}:** {rendered if rendered else '_(not found)_'}")

    # Title
    title = cfg.get("doc_title_template", "Application Summary — {app_id}").format(app_id=app_id)
    if app_name:
        title = f"{app_name} — {app_id}"

    # Prepend title + timestamp
    lines = [f"# {title}", "", f"_Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}_", ""] + lines + [""]
    return "\n".join(lines)

def main():
    p = argparse.ArgumentParser(description="Extract app data from multiple Excel workbooks into a Markdown summary.")
    p.add_argument("--config", required=True, help="Path to JSON config defining sources and fields.")
    p.add_argument("--app-id", required=True, help="Application ID to look up.")
    p.add_argument("--out", required=True, help="Output markdown file path.")
    p.add_argument("--source", action="append", help="Override a source path like A=/path/to/file.xlsx (can repeat).")
    args = p.parse_args()

    try:
        with open(args.config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        overrides = parse_source_overrides(args.source)
        md = extract_fields(cfg, args.app_id, overrides)
        out_dir = os.path.dirname(args.out)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"Wrote: {args.out}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
