import os
import json
import re
import pandas as pd
from typing import Any, Dict, List, Tuple

# Shared Utilities (case-insensitive keys + normalized value comparisons)

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

    for k in keys:
        if k in d:
            return d[k]

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
    """Case-insensitive + space/underscore/hyphen-insensitive equality."""
    return _norm_text(a) == _norm_text(b)

# Leaf-walk used by Bookmark reader helpers
def walk_values(obj: Any, path: Tuple[str, ...] = ()):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from walk_values(v, path + (str(k),))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from walk_values(v, path + (f"[{i}]",))
    else:
        yield (path, obj)

def find_first_value(obj: Any, keys: List[str]) -> Any:
    """Find first value where leaf path ends with any of keys (case-insensitive)."""
    keys_lower = set(k.lower() for k in keys)
    for path, val in walk_values(obj):
        if path and path[-1].lower() in keys_lower:
            return val
    return None

# Pages Reader 

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

def _literal_or_string(value_node):
    """Extract a concrete value from common literal shapes or direct primitives."""
    if isinstance(value_node, dict):
        v = get_nested(value_node, ["expr", "Literal", "Value"])
        if v is not None:
            return str(v).strip("'\"")
        v = get_nested(value_node, ["Literal", "Value"])
        if v is not None:
            return str(v).strip("'\"")
        raw = _get_any(value_node, ["Value"])
        if raw is not None and not isinstance(raw, (dict, list)):
            return str(raw).strip("'\"")
    elif isinstance(value_node, (str, int, float)):
        return str(value_node).strip("'\"")
    return None

def find_action_button_actions(visual_data):
    """
    Return a list of action strings:
    - 'Page Navigation: <id>'
    - 'Bookmark: <id>'
    - 'Tooltip: <id>'
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

            type_val = _literal_or_string(_get_any(props, ["type"]))
            type_val_norm = _norm_text(type_val)

            tooltip_text = _literal_or_string(_get_any(props, ["tooltip"]))
            if tooltip_text:
                actions.append(f"Tooltip: {tooltip_text}")

            if type_val_norm == _norm_text("pagenavigation"):
                nav_id = _literal_or_string(_get_any(props, ["navigationSection"]))
                if nav_id:
                    actions.append(f"Page Navigation: {nav_id}")
                else:
                    actions.append("Page Navigation")

            elif type_val_norm == _norm_text("bookmark"):
                bookmark_id = _literal_or_string(_get_any(props, ["bookmark"]))
                if bookmark_id:
                    actions.append(f"Bookmark: {bookmark_id}")
                else:
                    actions.append("Bookmark")

    return dedupe_preserve_order(actions)

def format_entity_with_spaces(entity_name: str) -> str:
    if not entity_name:
        return ""
    s = str(entity_name)
    if "." in s and " " not in s:
        return " ".join(part for part in s.split(".") if part)
    return s

def build_alias_map(visual_data):
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
    if not isinstance(field_like, dict):
        return {}
    for wrapper in ("Column", "Measure", "Aggregation", "Field"):
        inner = _get_any(field_like, [wrapper])
        if isinstance(inner, dict):
            return inner
    return field_like

def to_table_column_from_fieldlike(field_like: dict, alias_map=None):
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

def extract_literals(values_node):
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
    if v is None:
        return False
    s = str(v).strip()
    s_clean = s.strip("'").strip('"')
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
    s = str(v)
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1]
    s = s.replace("'", "''")
    return f"'{s}'"

def extract_visual_columns(visual_data, alias_map):
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

def find_table_column_in_condition(condition_node, alias_map):
    for node in walk(condition_node):
        if isinstance(node, dict):
            col = _get_any(node, ["Column"])
            if isinstance(col, dict):
                tc = to_table_column_from_fieldlike(col, alias_map)
                if tc:
                    return tc
    return None

def parse_operator_and_values(cond_node):
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

def extract_visual_info(visual_json_path):
    with open(visual_json_path, "r", encoding="utf-8") as f:
        visual_data = json.load(f)

    alias_map = build_alias_map(visual_data)
    visual_id = _get_any(visual_data, ["name"]) or ""
    visual_type = get_nested(visual_data, ["visual", "visualType"], "") or ""

    legacy_bookmark = find_first_bookmark(visual_data)
    legacy_tooltip = find_first_tooltip_value(visual_data)
    link_actions = find_action_button_actions(visual_data)

    action_type_parts = []
    if legacy_bookmark:
        action_type_parts.append(f"Bookmark: {legacy_bookmark}")
    if legacy_tooltip:
        action_type_parts.append(f"Tooltip: {legacy_tooltip}")

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

# Bookmark Reader 

def stem_bookmark_id(filename: str) -> str:
    return os.path.splitext(os.path.splitext(filename)[0])[0]

def normalize_entity(entity: str) -> str:
    if not isinstance(entity, str):
        return str(entity) if entity else ""
    return ".".join(part for part in entity.split() if part)

def stringify_list(values: List[Any]) -> str:
    return "[" + ",".join(f"'{v}'" if v is not None else "null" for v in values) + "]"

def extract_entity_property(filter_obj: Dict[str, Any]) -> Tuple[str, str]:
    entity = find_first_value(filter_obj, ["Entity", "Source", "Table"]) or ""
    entity = normalize_entity(entity)
    prop = find_first_value(filter_obj, ["Property", "Column", "Field"]) or ""
    return entity, prop

def extract_values(filter_obj: Dict[str, Any]) -> List[Any]:
    out = []
    expr = find_first_value(filter_obj, ["expression"]) or {}
    values_blocks = []

    if isinstance(expr, dict):
        vlist = _get_any(expr, ["Values"])
        if isinstance(vlist, list):
            values_blocks.extend(vlist)

    top_values = _get_any(filter_obj, ["Values"])
    if isinstance(top_values, list):
        values_blocks.extend(top_values)

    accept_literals = []
    for path, val in walk_values(filter_obj):
        if path and path[-1].lower() == "value" and any(p.lower() == "literal" for p in path):
            accept_literals.append(val)

    def flatten(v):
        if isinstance(v, dict):
            lit = _get_any(v, ["Literal"])
            if isinstance(lit, dict):
                if _has_any(lit, ["Value"]):
                    return [_get_any(lit, ["Value"])]
            if _has_any(v, ["Value"]):
                return [_get_any(v, ["Value"])]
            return []
        elif isinstance(v, list):
            flat = []
            for item in v:
                flat.extend(flatten(item))
            return flat
        else:
            return [v]

    for v in values_blocks:
        out.extend(flatten(v))

    if not out and accept_literals:
        out.extend(accept_literals)

    cleaned = []
    for v in out:
        if v is None:
            cleaned.append(None)
        else:
            cleaned.append(None if _eq_ci(str(v), "null") else v)
    return cleaned

def detect_operator(filter_obj: Dict[str, Any], values: List[Any]) -> Tuple[str, bool]:
    mode = find_first_value(filter_obj, ["mode"]) or ""
    where_texts = []

    for k in ["Where", "Condition", "where", "condition"]:
        val = find_first_value(filter_obj, [k])
        if isinstance(val, str):
            where_texts.append(val)
        elif isinstance(val, dict):
            for _, v in walk_values(val):
                if isinstance(v, str):
                    where_texts.append(v)

    blob = " ".join(where_texts).lower()
    is_negative = "not" in blob or any("not" in ".".join(path).lower() for path, _ in walk_values(filter_obj))

    if str(mode).lower() == "between" and len(values) == 2:
        return "BETWEEN", is_negative
    if " between " in blob and len(values) == 2:
        return "BETWEEN", is_negative
    if " in " in blob or len(values) > 1:
        return "IN", is_negative
    if len(values) == 1:
        return "=", is_negative

    return "", is_negative

def render_condition(entity: str, prop: str, operator: str, values: List[Any], is_negative: bool) -> str:
    lhs = f"{entity}.{prop}" if entity and prop else (entity or prop)
    if operator == "BETWEEN" and len(values) == 2:
        return f"{lhs} BETWEEN '{values[0]}' AND '{values[1]}'"
    if operator == "IN" and values:
        return f"{lhs} {'NOT IN' if is_negative else 'IN'} {stringify_list(values)}"
    if operator == "=" and values:
        # no special characters: use <> instead of ≠
        return f"{lhs} {'<>' if is_negative else '='} '{values[0]}'"
    return lhs

def summarize_filter(filter_obj: Dict[str, Any]) -> str:
    entity, prop = extract_entity_property(filter_obj)
    values = extract_values(filter_obj)
    operator, is_negative = detect_operator(filter_obj, values)
    return render_condition(entity, prop, operator, values, is_negative)

def summarize_filters(filters: List[Dict[str, Any]]) -> str:
    parts = []
    for f in filters:
        try:
            parts.append(summarize_filter(f))
        except Exception:
            parts.append(json.dumps(f))
    return "; ".join(parts) if parts else "None"

def flatten_values(values_block):
    out = []
    if isinstance(values_block, list):
        for item in values_block:
            out.extend(flatten_values(item))
    elif isinstance(values_block, dict):
        lit = _get_any(values_block, ["Literal"])
        if isinstance(lit, dict) and _has_any(lit, ["Value"]):
            out.append(str(_get_any(lit, ["Value"])).strip("'"))
    return out

def extract_slicer_selections(vdata):
    merge = (get_nested(vdata, ["singleVisual", "objects", "merge"], {}) or {})
    general_list = _get_any(merge, ["general"]) or []
    if isinstance(general_list, dict):
        general_list = [general_list]

    selections = []
    for general in general_list:
        props = _get_any(general, ["properties"]) or {}
        fil = _get_any(props, ["filter"]) or {}
        deep = _get_any(fil, ["filter"]) or {}
        if not deep:
            continue

        where_list = _get_any(deep, ["Where"]) or []
        if isinstance(where_list, dict):
            where_list = [where_list]

        for where_item in where_list:
            cond = _get_any(where_item, ["Condition"]) or {}
            in_block = _get_any(cond, ["In"]) or {}
            if not in_block:
                continue

            expressions = _get_any(in_block, ["Expressions"]) or []
            if isinstance(expressions, dict):
                expressions = [expressions]

            prop_name = ""
            if expressions:
                col = _get_any(expressions[0], ["Column"]) or {}
                prop_name = _get_any(col, ["Property"]) or ""

            values = flatten_values(_get_any(in_block, ["Values"]) or [])
            if prop_name:
                if len(values) == 1:
                    selections.append(f"{prop_name} = '{values[0]}'")
                elif len(values) > 1:
                    vals = ",".join(f"'{v}'" for v in values)
                    selections.append(f"{prop_name} IN [{vals}]")
                else:
                    selections.append(prop_name)

    return "; ".join(selections) if selections else "None"

def extract_visual_mode(vdata: Dict[str, Any]) -> str:
    mode = get_nested(vdata, ["singleVisual", "objects", "display", "mode"])
    if isinstance(mode, dict):
        literal = get_nested(mode, ["expr", "Literal", "Value"])
        if literal is not None:
            return str(literal).strip("'")
    elif isinstance(mode, str):
        return mode

    found = find_first_value(vdata, ["mode"])
    if isinstance(found, dict):
        literal = get_nested(found, ["expr", "Literal", "Value"])
        if literal is not None:
            return str(literal).strip("'")
    elif isinstance(found, str):
        return found

    return ""

def extract_visual_rows(bookmark, filename):
    bookmark_id = stem_bookmark_id(filename)
    bookmark_name = _get_any(bookmark, ["displayName"]) or ""

    options = _get_any(bookmark, ["options"]) or {}
    expl_state = _get_any(bookmark, ["explorationState"]) or {}
    expl_options = _get_any(expl_state, ["options"]) or {}

    selected_visuals = set(
        (_get_any(options, ["targetVisualNames"]) or []) +
        (_get_any(bookmark, ["targetVisualNames"]) or []) +
        (_get_any(expl_options, ["targetVisualNames"]) or [])
    )

    sections = _get_any(expl_state, ["sections"]) or {}
    active = _get_any(expl_state, ["activeSection"]) or ""

    active_block = sections.get(active, {}) if isinstance(sections, dict) else {}
    visuals = _get_any(active_block, ["visualContainers"]) or {}

    rows = []
    if isinstance(visuals, dict):
        for vid, vdata in visuals.items():
            vtype = (
                get_nested(vdata, ["singleVisual", "visualType"]) or
                _get_any(vdata, ["visualType"]) or
                _get_any(vdata, ["type"]) or
                ""
            )

            filt_block = _get_any(vdata, ["filters"]) or {}
            vfilters = _get_any(filt_block, ["byExpr"]) or []
            applied_filters_str = summarize_filters([x for x in vfilters if isinstance(x, dict)])

            slicer_selections = extract_slicer_selections(vdata) if _eq_ci(vtype, "slicer") else "None"
            selected_flag = "Yes" if str(vid) in selected_visuals else "No"
            mode_value = extract_visual_mode(vdata)

            rows.append({
                "Bookmark ID": bookmark_id,
                "Bookmark Name": bookmark_name,
                "Visual ID": vid,
                "Visual Type": vtype,
                "Selected Visual": selected_flag,
                "Mode": mode_value,
                "Applied Filters": applied_filters_str,
                "Slicer Selections": slicer_selections
            })

    return rows

def parse_bookmarks_folder(folder_path):
    all_rows = []
    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.lower().endswith(".json"):
                try:
                    with open(os.path.join(root, filename), "r", encoding="utf-8") as f:
                        bookmark = json.load(f)
                    all_rows.extend(extract_visual_rows(bookmark, filename))
                except Exception as e:
                    print(f"Error processing {filename}: {e}")
    return all_rows

# Integration: mapping + column additions in bookmarks and pages

def build_page_maps(pages_rows):
    """Build:
       - page_id_to_name
       - visual_id_to_page_name
    """
    page_id_to_name = {}
    visual_id_to_page_name = {}
    for r in pages_rows:
        pid = r.get("Page ID", "")
        pname = r.get("Page Name", "")
        vid = r.get("Visual ID", "")
        if pid and pid not in page_id_to_name:
            page_id_to_name[pid] = pname
        if vid and vid not in visual_id_to_page_name:
            visual_id_to_page_name[vid] = pname
    return page_id_to_name, visual_id_to_page_name

def build_bookmark_map(bookmark_rows):
    """bookmark_id_to_name"""
    bm = {}
    for r in bookmark_rows:
        bid = r.get("Bookmark ID", "")
        bname = r.get("Bookmark Name", "")
        if bid and bid not in bm:
            bm[bid] = bname
    return bm

def add_page_names_to_bookmarks(bookmark_rows, visual_id_to_page_name):
    """Insert Page Name after Visual ID in each bookmark row."""
    out = []
    for r in bookmark_rows:
        vid = r.get("Visual ID", "")
        page_name = visual_id_to_page_name.get(vid, "")
        # Create new dict in desired order
        new_r = {
            "Bookmark ID": r.get("Bookmark ID", ""),
            "Bookmark Name": r.get("Bookmark Name", ""),
            "Visual ID": vid,
            "Page Name": page_name,
            "Visual Type": r.get("Visual Type", ""),
            "Selected Visual": r.get("Selected Visual", ""),
            "Mode": r.get("Mode", ""),
            "Applied Filters": r.get("Applied Filters", ""),
            "Slicer Selections": r.get("Slicer Selections", "")
        }
        out.append(new_r)
    return out

def parse_action_type_to_action_name(action_type, page_id_to_name, bookmark_id_to_name):
    """Convert Action Type IDs to Action Name values.
       Supports multiple actions separated by ';'.
    """
    if not action_type:
        return ""

    parts = [p.strip() for p in str(action_type).split(";") if p.strip()]
    out_parts = []

    for p in parts:
        if ":" not in p:
            continue
        label, val = p.split(":", 1)
        label = label.strip()
        action_id = val.strip()

        label_norm = _norm_text(label)

        if label_norm == _norm_text("tooltip"):
            resolved = page_id_to_name.get(action_id, "")
            if resolved:
                out_parts.append(f"Tooltip: {resolved}")

        elif label_norm == _norm_text("pagenavigation") or label_norm == _norm_text("pagenavigation"):
            resolved = page_id_to_name.get(action_id, "")
            if resolved:
                out_parts.append(f"Page Navigation: {resolved}")

        elif label_norm == _norm_text("bookmark"):
            resolved = bookmark_id_to_name.get(action_id, "")
            if resolved:
                out_parts.append(f"Bookmark: {resolved}")

    return "; ".join(out_parts)

def add_action_name_to_pages(pages_rows, page_id_to_name, bookmark_id_to_name):
    """Add Action Name column after Action Type."""
    out = []
    for r in pages_rows:
        action_type = r.get("Action Type", "")
        action_name = parse_action_type_to_action_name(action_type, page_id_to_name, bookmark_id_to_name)

        new_r = {
            "Page ID": r.get("Page ID", ""),
            "Page Name": r.get("Page Name", ""),
            "Visual ID": r.get("Visual ID", ""),
            "Visual Type": r.get("Visual Type", ""),
            "Action Type": action_type,
            "Action Name": action_name,
            "Visual Columns": r.get("Visual Columns", ""),
            "Visual Filters": r.get("Visual Filters", "")
        }
        out.append(new_r)
    return out

# Main

if __name__ == "__main__":
    # Update folder paths
    bookmarks_folder = r"C:\Path\To\Your\Bookmarks"
    pages_folder = r"C:\Path\To\Your\pages"
    output_excel = "pages and bookmarks log.xlsx"

    print("Reading Bookmarks")
    bookmark_rows = parse_bookmarks_folder(bookmarks_folder)
    print(f"Retrieved: {len(bookmark_rows)} Rows")

    print("Reading Pages")
    pages_rows = parse_pages_folder(pages_folder)
    print(f"Retrieved: {len(pages_rows)} Rows")

    # Build maps
    page_id_to_name, visual_id_to_page_name = build_page_maps(pages_rows)
    bookmark_id_to_name = build_bookmark_map(bookmark_rows)

    print("Adding Page Names to Bookmarks tab")
    bookmark_rows_added = add_page_names_to_bookmarks(bookmark_rows, visual_id_to_page_name)

    print("Adding Bookmark Names and Page Names on Pages tab")
    pages_rows_added = add_action_name_to_pages(pages_rows, page_id_to_name, bookmark_id_to_name)

    print("Generating File")
    df_bookmarks = pd.DataFrame(
        bookmark_rows_added,
        columns=[
            "Bookmark ID", "Bookmark Name", "Visual ID", "Page Name", "Visual Type",
            "Selected Visual", "Mode", "Applied Filters", "Slicer Selections"
        ]
    )

    df_pages = pd.DataFrame(
        pages_rows_added,
        columns=[
            "Page ID", "Page Name", "Visual ID", "Visual Type", "Action Type",
            "Action Name", "Visual Columns", "Visual Filters"
        ]
    )

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_bookmarks.to_excel(writer, sheet_name="Bookmarks", index=False)
        df_pages.to_excel(writer, sheet_name="Pages", index=False)

    print(f"Data extracted to {output_excel}")