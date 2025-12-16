import os
import json
import re
import pandas as pd

# Utilities
def get_nested(data, keys, default=None):
    """Walk a nested dict using an ordered list of keys.
       Returns default if any step is missing.
       Case-insensitive for dict key matching.
    """
    for key in keys:
        if not isinstance(data, dict):
            return default

        # exact match first
        if key in data:
            data = data[key]
            continue

        # case-insensitive match
        if isinstance(key, str):
            target = key.lower()
            found = False
            for kk in data.keys():
                if isinstance(kk, str) and kk.lower() == target:
                    data = data[kk]
                    found = True
                    break
            if found:
                continue

        return default

    return data

def walk(obj):
    """Depth-first walk that returns all dicts and lists inside obj."""
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from walk(v)
    elif isinstance(obj, list):
        for item in obj:
            yield from walk(item)

def dedupe_preserve_order(seq):
    """Remove duplicates but keep original order."""
    out, seen = [], set()
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def _get_any(d: dict, keys: list):
    """Get the first matching key from a dict, case-insensitive."""
    if not isinstance(d, dict):
        return None

    # exact match first
    for k in keys:
        if k in d:
            return d[k]

    # case-insensitive match
    for desired in keys:
        if not isinstance(desired, str):
            continue
        dl = desired.lower()
        for kk in d.keys():
            if isinstance(kk, str) and kk.lower() == dl:
                return d[kk]

    return None

def _has_any(d: dict, keys: list) -> bool:
    """True if any matching key exists in the dict (case-insensitive), regardless of the value."""
    if not isinstance(d, dict):
        return False

    for k in keys:
        if k in d:
            return True

    for desired in keys:
        if not isinstance(desired, str):
            continue
        dl = desired.lower()
        for kk in d.keys():
            if isinstance(kk, str) and kk.lower() == dl:
                return True

    return False

def _norm_text(x) -> str:
    """Normalize text for comparisons only (not output):
       - case-insensitive
       - remove spaces, underscores, hyphens
       So: 'Page Navigation', 'page_navigation', 'page-navigation' -> 'pagenavigation'
    """
    if x is None:
        return ""
    s = str(x).strip().casefold()
    s = re.sub(r"[\s_-]+", "", s)
    return s

def _eq_ci(a, b) -> bool:
    """Case-insensitive + space/underscore/hyphen-insensitive equality for string-like values."""
    return _norm_text(a) == _norm_text(b)

# Bookmark and tooltip helpers
def find_first_bookmark(obj):
    """Return the first bookmark name if a bookmark action is found."""
    if isinstance(obj, dict):
        if _has_any(obj, ["bookmark"]):
            val = get_nested(obj, ["bookmark", "expr", "Literal", "Value"])
            if val:
                return str(val).strip("'")
        for v in obj.values():
            found = find_first_bookmark(v)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_first_bookmark(item)
            if found:
                return found
    return None

def find_first_tooltip_value(obj):
    """
    Find the first tooltip text under 'visualTooltip' or 'tooltip'.
    Some action buttons store tooltip under visualLink.properties.tooltip.
    """
    if isinstance(obj, dict):
        for key in ["visualTooltip", "tooltip"]:
            if _has_any(obj, [key]):
                tv = _get_any(obj, [key])

                if isinstance(tv, list):
                    for item in tv:
                        val = get_nested(item, ["properties", "section", "expr", "Literal", "Value"])
                        if val:
                            return str(val).strip("'")
                        val2 = get_nested(item, ["properties", "section", "value"])
                        if val2:
                            return str(val2).strip("'")
                        if isinstance(item, str) and item:
                            return item.strip("'")

                elif isinstance(tv, dict):
                    # fallback: look for expr->Literal->Value directly
                    val = get_nested(tv, ["expr", "Literal", "Value"])
                    if val:
                        return str(val).strip("'")
                    val3 = _get_any(tv, ["value"])
                    if val3:
                        return str(val3).strip("'")

        for v in obj.values():
            found = find_first_tooltip_value(v)
            if found:
                return found

    elif isinstance(obj, list):
        for item in obj:
            found = find_first_tooltip_value(item)
            if found:
                return found
    return None

# Page Navigation / Bookmark / Tooltip detector for action buttons
def _literal_or_string(value_node):
    """
    Extract a concrete value from these shapes:
    - dict: { 'expr': { 'Literal': { 'Value': '...' } } }
    - dict: { 'Literal': { 'Value': '...' } }
    - direct string/int/float
    """
    if isinstance(value_node, dict):
        v = get_nested(value_node, ["expr", "Literal", "Value"])
        if v is not None:
            return str(v).strip("'\"")

        v = get_nested(value_node, ["Literal", "Value"])
        if v is not None:
            return str(v).strip("'\"")

        # handle Value/value keys directly
        raw = _get_any(value_node, ["Value"])
        if raw is not None and not isinstance(raw, (dict, list)):
            return str(raw).strip("'\"")

    elif isinstance(value_node, (str, int, float)):
        return str(value_node).strip("'\"")

    return None

def find_action_button_actions(visual_data):
    """
    Walk visualContainerObjects.visualLink entries (list or dict) and return
    a list of action strings like:
    - 'Page Navigation: <id>'
    - 'Bookmark: <id>' (if present)
    - 'Tooltip: <text>' (from visualLink properties)
    """
    actions = []

    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        visual_link = _get_any(node, ["visualLink"])
        if not visual_link:
            continue

        items = visual_link if isinstance(visual_link, list) else [visual_link]
        for link_item in items:
            if not isinstance(link_item, dict):
                continue

            props = _get_any(link_item, ["properties"]) or link_item

            # type: PageNavigation / Bookmark / etc.
            type_val = _literal_or_string(_get_any(props, ["type"]))
            type_val_norm = _norm_text(type_val)

            # Tooltip text (stored directly under visualLink.properties.tooltip)
            tooltip_text = _literal_or_string(_get_any(props, ["tooltip"]))
            if tooltip_text:
                actions.append(f"Tooltip: {tooltip_text}")

            # PageNavigation target
            if type_val_norm == _norm_text("pagenavigation"):
                nav_id = _literal_or_string(_get_any(props, ["navigationSection"]))
                if nav_id:
                    actions.append(f"Page Navigation: {nav_id}")
                else:
                    actions.append("Page Navigation")

            # Bookmark action (may or may not have an explicit 'bookmark' id)
            elif type_val_norm == _norm_text("bookmark"):
                bookmark_id = _literal_or_string(_get_any(props, ["bookmark"]))
                if bookmark_id:
                    actions.append(f"Bookmark: {bookmark_id}")
                else:
                    actions.append("Bookmark")

    return dedupe_preserve_order(actions)

# Alias resolution and fields to help output
def format_entity_with_spaces(entity_name: str) -> str:
    """Entity (table) name uses spaces; convert dotted paths to spaces."""
    if not entity_name:
        return ""
    s = str(entity_name)
    if "." in s and " " not in s:
        return " ".join(part for part in s.split(".") if part)
    return s

def build_alias_map(visual_data):
    """Build alias map (e.g., g -> table name) from any 'From' blocks."""
    alias = {}

    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        frm = _get_any(node, ["From"])
        if isinstance(frm, list):
            for item in frm:
                if isinstance(item, dict):
                    name = _get_any(item, ["Name"])
                    entity = _get_any(item, ["Entity"])
                    if name and entity:
                        alias[str(name)] = format_entity_with_spaces(entity)

        elif isinstance(frm, dict):
            name = _get_any(frm, ["Name"])
            entity = _get_any(frm, ["Entity"])
            if name and entity:
                alias[str(name)] = format_entity_with_spaces(entity)

    return alias

def unwrap_field_like(field_like: dict) -> dict:
    """Unwrap 'Column'/'Measure'/'Aggregation'/'Field' to reach Expression/Property."""
    if not isinstance(field_like, dict):
        return {}

    for wrapper in ("Column", "Measure", "Aggregation", "Field"):
        inner = _get_any(field_like, [wrapper])
        if isinstance(inner, dict):
            return inner

    return field_like

def to_table_column_from_fieldlike(field_like: dict, alias_map=None):
    """
    Produce 'Table Name.Column' from any field-like dict.
    Treat '_DAX' like a normal table; only ignore numeric-only properties (sort hints).
    """
    fld = unwrap_field_like(field_like)

    expr = _get_any(fld, ["Expression"]) or {}
    src = _get_any(expr, ["SourceRef"]) or {}

    entity = _get_any(src, ["Entity"])
    source = _get_any(src, ["Source"])

    if not entity and source and isinstance(alias_map, dict):
        entity = alias_map.get(str(source))

    if not entity:
        entity = _get_any(fld, ["Entity"])

    prop = _get_any(fld, ["Property"])
    if prop is not None and re.fullmatch(r"\d+", str(prop)):
        return None

    if entity and prop:
        entity_spaces = format_entity_with_spaces(str(entity))
        return f"{entity_spaces}.{prop}"

    if prop:
        return str(prop)

    return None

# Literal extraction and value cleanup
def extract_literals(values_node):
    """Collect literal values across nested structures."""
    vals = []

    def visit(n):
        if isinstance(n, dict):
            lit = _get_any(n, ["Literal"])
            if isinstance(lit, dict):
                v = _get_any(lit, ["Value"])
                if v is not None:
                    vals.append(v)
            for vv in n.values():
                visit(vv)

        elif isinstance(n, list):
            for item in n:
                visit(item)

        elif isinstance(n, (str, int, float)) and n is not None:
            vals.append(n)

    visit(values_node)

    uniq, seen = [], set()
    for v in vals:
        key = repr(v)
        if key not in seen:
            seen.add(key)
            uniq.append(v)

    return uniq

def is_valid_filter_value(v):
    """
    Keep meaningful values only:
    - Drop None or 'null' (any casing).
    - Drop single-letter alphabetic tokens (e.g., 'a', 'C').
    """
    if v is None:
        return False

    s = str(v).strip()
    s_clean = s.strip("'").strip('"')

    # 'null' check is case-insensitive and also normalizes spaces/underscores/hyphens (comparison only)
    if _eq_ci(s_clean, "null"):
        return False

    s_unquoted = s
    if (s_unquoted.startswith("'") and s_unquoted.endswith("'")) or \
       (s_unquoted.startswith('"') and s_unquoted.endswith('"')):
        s_unquoted = s_unquoted[1:-1]

    if re.fullmatch(r"[A-Za-z]", s_unquoted):
        return False

    return True

def stringify_value(v):
    """Render values as single-quoted, doubling any embedded single quotes."""
    s = str(v)
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1]
    s = s.replace("'", "''")
    return f"'{s}'"

# Column extraction
def extract_visual_columns(visual_data, alias_map):
    """
    Gather all columns used in the visual, formatted as 'Table Name.Column'.
    Finds them in fields, aggregations, and generic Expression/Property patterns.
    """
    cols, seen = [], set()

    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        for key in ("field", "Aggregation"):
            sub = _get_any(node, [key])
            if isinstance(sub, dict):
                tc = to_table_column_from_fieldlike(sub, alias_map)
                if tc and tc not in seen:
                    seen.add(tc)
                    cols.append(tc)

        fields_list = _get_any(node, ["fields"])
        if isinstance(fields_list, list):
            for f in fields_list:
                if isinstance(f, dict):
                    tc = to_table_column_from_fieldlike(f, alias_map)
                    if tc and tc not in seen:
                        seen.add(tc)
                        cols.append(tc)

        expr = _get_any(node, ["Expression"])
        src = _get_any(expr or {}, ["SourceRef"])

        has_entity_or_source = (
            _get_any(src or {}, ["Entity", "Source"]) is not None
            or _get_any(node, ["Entity"]) is not None
        )
        has_prop = _get_any(node, ["Property"]) is not None

        if has_entity_or_source and has_prop:
            tc = to_table_column_from_fieldlike(node, alias_map)
            if tc and tc not in seen:
                seen.add(tc)
                cols.append(tc)

    return dedupe_preserve_order(cols)

# Filter extraction (supports NOT, list/dict Where, and alias resolution)
def find_table_column_in_condition(condition_node, alias_map):
    """Find the first column reference inside a condition and format it."""
    for node in walk(condition_node):
        if isinstance(node, dict):
            col = _get_any(node, ["Column"])
            if isinstance(col, dict):
                tc = to_table_column_from_fieldlike(col, alias_map)
                if tc:
                    return tc
    return None

def parse_operator_and_values(cond_node):
    """
    Parse operator and values from a condition:
    IN / = / BETWEEN / NOT IN / NOT BETWEEN,
    and handle NOT {Expression} wrapping the inner operator.
    """
    if not isinstance(cond_node, dict):
        return None, []

    in_node = _get_any(cond_node, ["In"])
    eq_node = _get_any(cond_node, ["Equals"])
    bt_node = _get_any(cond_node, ["Between"])
    ni_node = _get_any(cond_node, ["NotIn"])

    if in_node is not None:
        values = _get_any(in_node, ["Values"]) or in_node
        return "IN", extract_literals(values)

    if eq_node is not None:
        values = _get_any(eq_node, ["Values"]) or eq_node
        return "=", extract_literals(values)

    if bt_node is not None:
        values = _get_any(bt_node, ["Values"]) or bt_node
        return "BETWEEN", extract_literals(values)

    if ni_node is not None:
        values = _get_any(ni_node, ["Values"]) or ni_node
        return "NOT IN", extract_literals(values)

    not_node = _get_any(cond_node, ["Not"])
    if isinstance(not_node, dict):
        inner = _get_any(not_node, ["Expression"]) or not_node
        inner_op, inner_vals = parse_operator_and_values(inner)
        if inner_op:
            if inner_op == "IN":
                return "NOT IN", inner_vals
            if inner_op == "=":
                return "NOT IN", inner_vals
            if inner_op == "BETWEEN":
                return "NOT BETWEEN", inner_vals
            return "NOT " + inner_op, inner_vals

    expr_node = _get_any(cond_node, ["Expression"])
    if isinstance(expr_node, dict):
        return parse_operator_and_values(expr_node)

    return None, []

def extract_visual_filters(visual_data, alias_map):
    """
    Collect filters from both 'filterConfig.filters' entries and standalone 'filter' objects.
    Supports 'Where' as list or dict, NOT-wrapped operators, and cleans values.
    """
    results = []

    def handle_filter_like(filter_like, field_hint=None):
        where = _get_any(filter_like, ["Where"])
        if isinstance(where, list):
            where_items = where
        elif isinstance(where, dict):
            where_items = [where]
        else:
            return

        for where_item in where_items:
            if not isinstance(where_item, dict):
                continue

            cond = _get_any(where_item, ["Condition"]) or {}
            if not isinstance(cond, dict):
                continue

            table_col = field_hint or find_table_column_in_condition(cond, alias_map)
            if not table_col:
                continue

            op, vals = parse_operator_and_values(cond)
            if not op:
                continue

            vals = [v for v in vals if is_valid_filter_value(v)]
            if not vals and op not in ("BETWEEN", "NOT BETWEEN"):
                continue

            formatted = [stringify_value(v) for v in vals]

            if op in ("=", "IN") and len(formatted) == 1:
                predicate = f"{table_col} = {formatted[0]}"
            elif op in ("IN", "NOT IN"):
                predicate = f"{table_col} {op} [{','.join(formatted)}]"
            elif op in ("BETWEEN", "NOT BETWEEN") and len(formatted) >= 2:
                predicate = f"{table_col} {op} {formatted[0]} AND {formatted[1]}"
            else:
                predicate = f"{table_col} {op} [{','.join(formatted)}]"

            results.append(predicate)

    # Filters declared in filterConfig.filters
    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        filters_node = _get_any(node, ["filters"])
        if isinstance(filters_node, list):
            for f in filters_node:
                if not isinstance(f, dict):
                    continue

                field_dict = _get_any(f, ["field"]) or f
                field_tc = to_table_column_from_fieldlike(field_dict, alias_map)

                filter_container = _get_any(f, ["filter"]) or f
                handle_filter_like(filter_container, field_hint=field_tc)

    # Standalone filter dicts (e.g., slicer objects.general.properties.filter)
    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        filter_like = _get_any(node, ["filter"])
        if isinstance(filter_like, dict):
            field_tc = None
            field_dict = _get_any(node, ["field"])
            if isinstance(field_dict, dict):
                field_tc = to_table_column_from_fieldlike(field_dict, alias_map)
            handle_filter_like(filter_like, field_hint=field_tc)

    return dedupe_preserve_order(results)

# High-level extractors (per visual and per page)
def extract_visual_info(visual_json_path):
    with open(visual_json_path, "r", encoding="utf-8") as f:
        visual_data = json.load(f)

    alias_map = build_alias_map(visual_data)

    visual_id = _get_any(visual_data, ["name"]) or ""
    visual_type = get_nested(visual_data, ["visual", "visualType"], "") or ""

    # Action Type: combine legacy + visualLink-based actions
    legacy_bookmark = find_first_bookmark(visual_data)
    legacy_tooltip = find_first_tooltip_value(visual_data)
    link_actions = find_action_button_actions(visual_data)

    action_type_parts = []
    if legacy_bookmark:
        action_type_parts.append(f"Bookmark: {legacy_bookmark}")
    if legacy_tooltip:
        action_type_parts.append(f"Tooltip: {legacy_tooltip}")

    # Append visualLink-derived actions only for action buttons
    # Comparison is case + space/underscore/hyphen-insensitive
    if _eq_ci(visual_type, "actionbutton"):
        action_type_parts.extend(link_actions)

    action_type = "; ".join(dedupe_preserve_order(action_type_parts)) if action_type_parts else ""

    visual_columns = extract_visual_columns(visual_data, alias_map)
    visual_filters = extract_visual_filters(visual_data, alias_map)

    return (
        visual_id,
        visual_type,
        action_type,
        "; ".join(visual_columns),
        "; ".join(visual_filters),
    )

def extract_page_info(page_json_path):
    with open(page_json_path, "r", encoding="utf-8") as f:
        page_data = json.load(f)

    page_id = _get_any(page_data, ["name"]) or ""
    page_name = _get_any(page_data, ["displayName"]) or ""
    return page_id, page_name

def parse_pages_folder(pages_folder):
    rows = []

    for page_folder in os.listdir(pages_folder):
        page_path = os.path.join(pages_folder, page_folder)
        if not os.path.isdir(page_path):
            continue

        page_json_path = os.path.join(page_path, "page.json")
        if not os.path.exists(page_json_path):
            continue

        page_id, page_name = extract_page_info(page_json_path)

        visuals_folder = os.path.join(page_path, "visuals")
        if os.path.exists(visuals_folder):
            for visual_folder in os.listdir(visuals_folder):
                visual_path = os.path.join(visuals_folder, visual_folder)
                if not os.path.isdir(visual_path):
                    continue

                visual_json_path = os.path.join(visual_path, "visual.json")
                if os.path.exists(visual_json_path):
                    v_id, v_type, action_type, v_cols, v_filters = extract_visual_info(visual_json_path)
                    rows.append({
                        "Page ID": page_id,
                        "Page Name": page_name,
                        "Visual ID": v_id,
                        "Visual Type": v_type,
                        "Action Type": action_type,
                        "Visual Columns": v_cols,
                        "Visual Filters": v_filters
                    })

    return rows

# Main
if __name__ == "__main__":
    # Update folder path
    pages_folder = r"C:\Path\To\Your\pages"
    output_excel = "pages log.xlsx"

    rows = parse_pages_folder(pages_folder)
    df = pd.DataFrame(rows, columns=[
        "Page ID", "Page Name", "Visual ID", "Visual Type", "Action Type",
        "Visual Columns", "Visual Filters"
    ])

    # Write to Excel
    df.to_excel(output_excel, index=False, engine="openpyxl")
    print(f"Extraction complete Rows: {len(df)} in {output_excel}")
