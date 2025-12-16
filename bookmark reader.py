import os
import json
import pandas as pd
from typing import Any, Dict, List, Tuple

# Utility Functions

def stem_bookmark_id(filename: str) -> str:
    return os.path.splitext(os.path.splitext(filename)[0])[0]

def walk(obj: Any, path: Tuple[str, ...] = ()):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from walk(v, path + (str(k),))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from walk(v, path + (f"[{i}]",))
    else:
        yield (path, obj)

def find_first(obj: Any, keys: List[str]) -> Any:
    keys_lower = set(k.lower() for k in keys)
    for path, val in walk(obj):
        if path and path[-1].lower() in keys_lower:
            return val
    return None

def normalize_entity(entity: str) -> str:
    if not isinstance(entity, str):
        return str(entity) if entity else ""
    return ".".join(part for part in entity.split() if part)

def stringify_list(values: List[Any]) -> str:
    return "[" + ",".join(f"'{v}'" if v is not None else "null" for v in values) + "]"

# Applied Filters Logic

def extract_entity_property(filter_obj: Dict[str, Any]) -> Tuple[str, str]:
    entity = find_first(filter_obj, ["Entity", "Source", "Table"]) or ""
    entity = normalize_entity(entity)
    prop = find_first(filter_obj, ["Property", "Column", "Field"]) or ""
    return entity, prop

def extract_values(filter_obj: Dict[str, Any]) -> List[Any]:
    out = []
    expr = find_first(filter_obj, ["expression"]) or {}
    values_blocks = []
    if isinstance(expr, dict) and isinstance(expr.get("Values"), list):
        values_blocks.extend(expr["Values"])
    top_values = filter_obj.get("Values") or filter_obj.get("values")
    if isinstance(top_values, list):
        values_blocks.extend(top_values)
    accept_literals = []
    for path, val in walk(filter_obj):
        if path and path[-1].lower() == "value" and any(p.lower() == "literal" for p in path):
            accept_literals.append(val)

    def flatten(v):
        if isinstance(v, dict):
            lit = v.get("Literal")
            if isinstance(lit, dict) and "Value" in lit:
                return [lit["Value"]]
            if "Value" in v:
                return [v["Value"]]
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

    return [None if str(v).lower() == "null" else v for v in out]

def detect_operator(filter_obj: Dict[str, Any], values: List[Any]) -> Tuple[str, bool]:
    mode = find_first(filter_obj, ["mode"]) or ""
    where_texts = []
    for k in ["Where", "Condition", "where", "condition"]:
        val = find_first(filter_obj, [k])
        if isinstance(val, str):
            where_texts.append(val)
        elif isinstance(val, dict):
            for _, v in walk(val):
                if isinstance(v, str):
                    where_texts.append(v)
    blob = " ".join(where_texts).lower()
    is_negative = "not" in blob or any("not" in ".".join(path).lower() for path, _ in walk(filter_obj))

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
        return f"{lhs} {'≠' if is_negative else '='} '{values[0]}'"
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

# Slicer Selections Logic

def flatten_values(values_block):
    out = []
    if isinstance(values_block, list):
        for item in values_block:
            out.extend(flatten_values(item))
    elif isinstance(values_block, dict):
        lit = values_block.get("Literal")
        if isinstance(lit, dict) and "Value" in lit:
            out.append(lit["Value"].strip("'"))
    return out

def extract_slicer_selections(vdata):
    merge = (vdata.get("singleVisual", {}).get("objects", {}).get("merge", {}) or {})
    general_list = merge.get("general", [])
    if isinstance(general_list, dict):
        general_list = [general_list]

    selections = []
    for general in general_list:
        props = general.get("properties", {})
        fil = props.get("filter", {})
        deep = fil.get("filter", {})
        if not deep:
            continue

        where_list = deep.get("Where", [])
        if isinstance(where_list, dict):
            where_list = [where_list]

        for where_item in where_list:
            cond = where_item.get("Condition", {})
            in_block = cond.get("In", {})
            if not in_block:
                continue

            expressions = in_block.get("Expressions", [])
            if isinstance(expressions, dict):
                expressions = [expressions]

            prop_name = ""
            if expressions:
                col = expressions[0].get("Column", {})
                prop_name = col.get("Property", "")

            values = flatten_values(in_block.get("Values", []))

            if prop_name:
                if len(values) == 1:
                    selections.append(f"{prop_name} = '{values[0]}'")
                elif len(values) > 1:
                    vals = ",".join(f"'{v}'" for v in values)
                    selections.append(f"{prop_name} IN [{vals}]")
                else:
                    selections.append(prop_name)

    return "; ".join(selections) if selections else "None"

# Extract visual mode: Hidden
def extract_visual_mode(vdata: Dict[str, Any]) -> str:
 
     # Common placement in json
    mode = (
        vdata.get("singleVisual", {})
             .get("objects", {})
             .get("display", {})
             .get("mode")
    )
    if isinstance(mode, dict):
        # Some schemas embed as { expr: { Literal: { Value: '...' } } }
        literal = mode.get("expr", {}).get("Literal", {}).get("Value")
        if literal is not None:
            return str(literal).strip("'")
    elif isinstance(mode, str):
        return mode

    # Fallback: search anywhere in the visual block for a field named 'mode'
    found = find_first(vdata, ["mode"])
    if isinstance(found, dict):
        literal = found.get("expr", {}).get("Literal", {}).get("Value")
        if literal is not None:
            return str(literal).strip("'")
    elif isinstance(found, str):
        return found

    return ""

# Combine Applied Filters Logic + Slicer Selections Logic + Selected Visual Column + Mode

def extract_visual_rows(bookmark, filename):
    bookmark_id = stem_bookmark_id(filename)
    bookmark_name = bookmark.get("displayName", "")
    selected_visuals = set(bookmark.get("options", {}).get("targetVisualNames", []) +
                           bookmark.get("targetVisualNames", []) +
                           bookmark.get("explorationState", {}).get("options", {}).get("targetVisualNames", []))

    sections = bookmark.get("explorationState", {}).get("sections", {}) or {}
    active = bookmark.get("explorationState", {}).get("activeSection", "")
    visuals = sections.get(active, {}).get("visualContainers", {}) if active else {}

    rows = []
    for vid, vdata in visuals.items():
        vtype = (vdata.get("singleVisual", {}).get("visualType") or
                 vdata.get("visualType") or vdata.get("type") or "")
        vfilters = ((vdata.get("filters") or {}).get("byExpr") or [])
        applied_filters_str = summarize_filters([x for x in vfilters if isinstance(x, dict)])
        slicer_selections = extract_slicer_selections(vdata) if vtype == "slicer" else "None"
        selected_flag = "Yes" if str(vid) in selected_visuals else "No"

        mode_value = extract_visual_mode(vdata)

        rows.append({
            "Bookmark ID": bookmark_id,
            "Bookmark Name": bookmark_name,
            "Visual ID": vid,
            "Visual Type": vtype,
            "Selected Visual": selected_flag,
            "Applied Filters": applied_filters_str,
            "Slicer Selections": slicer_selections,
            "Mode": mode_value 
        })
    return rows

def parse_bookmarks_folder(folder_path, out_excel):
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

    df = pd.DataFrame(all_rows, columns=[
        "Bookmark ID", "Bookmark Name", "Visual ID", "Visual Type",
        "Selected Visual", "Mode", "Applied Filters", "Slicer Selections"
    ])
    df.to_excel(out_excel, index=False)
    print(f"Extraction complete! See '{out_excel}'.")

if __name__ == "__main__":
    folder = r"C:\Path\To\Your\Bookmarks"
    out_xlsx = "bookmarks log.xlsx"
    parse_bookmarks_folder(folder, out_xlsx)
