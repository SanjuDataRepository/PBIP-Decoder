# PBIP Decoder

import os
import re
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional
import pandas as pd

 
# USER CONFIGURATION (Plug & Play)
 
ROOT_FOLDER = r"Insert Folder Path Here"
 

# 1. SHARED UTILITIES & RECURSIVE WALKERS
 

def get_nested(data: Any, keys: List[Any], default: Any = None) -> Any:
    """Walk a nested dict using ordered keys. Case-insensitive."""
    for key in keys:
        if not isinstance(data, dict):
            return default

        if key in data:
            data = data[key]
            continue

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


def walk(obj: Any):
    """Depth-first generator over nested dict/list structures."""
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from walk(v)
    elif isinstance(obj, list):
        for item in obj:
            yield from walk(item)


def dedupe_preserve_order(seq: List[Any]) -> List[Any]:
    out, seen = [], set()
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def _get_any(d: dict, keys: list) -> Any:
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


def _norm_text(x: Any) -> str:
    if x is None:
        return ""
    s = str(x).strip().casefold()
    return re.sub(r"[\s_-]+", "", s)


def _eq_ci(a: Any, b: Any) -> bool:
    return _norm_text(a) == _norm_text(b)


def _norm_id(x: Any) -> str:
    if x is None:
        return ""
    return str(x).strip().casefold()


def extract_literal_str(val_node: Any) -> str:
    """Extracts unquoted string literal from expression nodes or raw strings."""
    if val_node is None:
        return ""
    if isinstance(val_node, (str, int, float)):
        s = str(val_node).strip()
        if (s.startswith("'") and s.endswith("'")) or (s.startswith('"') and s.endswith('"')):
            s = s[1:-1]
        return s
    if isinstance(val_node, dict):
        for k in ["expr", "Literal", "Value"]:
            if k in val_node:
                return extract_literal_str(val_node[k])
        if "Value" in val_node:
            return extract_literal_str(val_node["Value"])
    return ""


def format_entity_with_spaces(entity_name: str) -> str:
    """
    Standardized entity/table name formatter across both Bookmarks and Pages.
    Converts internal dot-delimited table names ('intensity.minutes') into spaces ('intensity minutes').
    """
    if not entity_name:
        return ""
    s = str(entity_name).strip()
    if "." in s and " " not in s:
        return " ".join(part for part in s.split(".") if part)
    return s


 
# 2. WORKSPACE AUTO-DISCOVERY MODULE
 

class WorkspaceDiscovery:
    """Recursively scans the root directory to find .pbix, bookmarks, and pages."""

    def __init__(self, root_dir: str):
        self.root_path = Path(root_dir).resolve()
        self.report_base_name: str = "PBIP_Report"
        self.pbix_file: Optional[Path] = None
        self.bookmarks_folder: Optional[Path] = None
        self.pages_folder: Optional[Path] = None

    def discover(self):
        if not self.root_path.exists():
            raise FileNotFoundError(f"Root path does not exist:\n{self.root_path}")

        print("=" * 70)
        print(f"[*] SCANNING WORKSPACE (NATIVE PBIP): {self.root_path}")
        print("=" * 70)

        # Discover top-level .pbix
        for item in self.root_path.iterdir():
            if item.is_file() and item.suffix.lower() == ".pbix":
                self.pbix_file = item
                self.report_base_name = item.stem
                print(f"  [+] Found PBIX: {item.name}")
                break

        # Case-insensitive recursive discovery of bookmarks and pages folders
        for root, dirs, _ in os.walk(self.root_path):
            root_path_obj = Path(root)
            for d in dirs:
                dl = d.lower()
                dir_full = root_path_obj / d
                if dl == "bookmarks" and self.bookmarks_folder is None:
                    if any(f.lower().endswith(".json") for f in os.listdir(dir_full)):
                        self.bookmarks_folder = dir_full
                        print(f"  [+] Discovered Bookmarks: {dir_full.relative_to(self.root_path)}")
                elif dl == "pages" and self.pages_folder is None:
                    if any(os.path.isdir(dir_full / sub) for sub in os.listdir(dir_full)):
                        self.pages_folder = dir_full
                        print(f"  [+] Discovered Pages: {dir_full.relative_to(self.root_path)}")

        print("=" * 70 + "\n")
        return self


 
# 3. NATIVE VISUAL TITLE & SOURCE CLASSIFIER
 

def resolve_visual_title_and_source(visual_data: Dict[str, Any],
                                    visual_type: str,
                                    visual_columns: List[str]) -> Tuple[str, str]:
    """
    Extracts the visual title and classifies where it was retrieved from:
      - 'Selection Pane (Renamed)'
      - 'Visual Header Title'
      - 'Button / Text Content'
      - 'Selection Pane (Default)'
      - 'Auto Generated Title'
    """
    vt_lower = visual_type.lower()

    # Standard default Selection Pane names for Power BI control/container visual types
    default_control_names = {
        "slicer": "Slicer",
        "shape": "Shape",
        "basicshape": "Shape",
        "image": "Image",
        "actionbutton": "Button",
        "textbox": "Text Box",
        "pagetag": "Page Navigator",
        "pagenavigator": "Page Navigator",
        "bookmarknavigator": "Bookmark Navigator"
    }

    # 1. Top-level displayName (Direct Selection Pane layer rename in PBIR)
    if visual_data.get("displayName"):
        return str(visual_data["displayName"]).strip(), "Selection Pane (Renamed)"
    if visual_data.get("visualGroup", {}).get("displayName"):
        return str(visual_data["visualGroup"]["displayName"]).strip(), "Selection Pane (Renamed)"

    is_control_or_shape = vt_lower in ["actionbutton", "image", "shape", "basicshape", "textbox", "slicer"]

    # 2. Check visualContainerObjects -> title (Formatting -> General -> Title)
    vco = visual_data.get("visualContainerObjects") or visual_data.get("visual", {}).get("visualContainerObjects") or {}
    if isinstance(vco, dict):
        title_objs = vco.get("title") or []
        if isinstance(title_objs, dict):
            title_objs = [title_objs]
        for t_obj in title_objs:
            if isinstance(t_obj, dict):
                props = t_obj.get("properties", {})
                show_val = extract_literal_str(props.get("show"))
                text_val = extract_literal_str(props.get("text") or props.get("title"))
                if text_val:
                    if show_val.lower() == "false" or is_control_or_shape:
                        return text_val, "Selection Pane (Renamed)"
                    return text_val, "Visual Header Title"

    # 3. Check visual.objects -> title or general
    objs = visual_data.get("objects") or visual_data.get("visual", {}).get("objects") or {}
    if isinstance(objs, dict):
        title_objs = objs.get("title") or []
        if isinstance(title_objs, dict):
            title_objs = [title_objs]
        for t_obj in title_objs:
            if isinstance(t_obj, dict):
                props = t_obj.get("properties", {})
                show_val = extract_literal_str(props.get("show"))
                text_val = extract_literal_str(props.get("text") or props.get("title"))
                if text_val:
                    if show_val.lower() == "false" or is_control_or_shape:
                        return text_val, "Selection Pane (Renamed)"
                    return text_val, "Visual Header Title"

        # Check buttonText or text inside buttons / cards / shapes
        for btn_key in ["buttonText", "text"]:
            b_list = objs.get(btn_key) or []
            if isinstance(b_list, dict):
                b_list = [b_list]
            for b_obj in b_list:
                if isinstance(b_obj, dict):
                    props = b_obj.get("properties", {})
                    text_val = extract_literal_str(props.get("text") or props.get("title"))
                    if text_val:
                        return text_val, "Button / Text Content"

    # 4. For controls/containers without an explicit title (Slicer, Shape, Image, Button):
    # In Power BI Selection Pane, these ALWAYS show their default type name, NEVER the bound column!
    if vt_lower in default_control_names:
        return default_control_names[vt_lower], "Selection Pane (Default)"

    # 5. For Charts / Data Visuals without explicit titles:
    # If projections exist, infer title from primary metric
    if visual_columns:
        primary_col = visual_columns[-1] if len(visual_columns) > 1 else visual_columns[0]
        field_part = primary_col.split(".")[-1]
        spaced = re.sub(r"([a-z])([A-Z])", r"\1 \2", field_part).strip()
        return spaced, "Auto Generated Title"

    # 6. Fallback Chart Type Default
    type_map = {
        "linechart": "Line Chart",
        "barchart": "Bar Chart",
        "columnchart": "Column Chart",
        "donutchart": "Donut Chart",
        "piechart": "Pie Chart",
        "clusteredbarchart": "Clustered Bar Chart",
        "clusteredcolumnchart": "Clustered Column Chart"
    }
    chart_name = type_map.get(vt_lower, visual_type.capitalize() if visual_type else "Visual")
    return chart_name, "Selection Pane (Default)"

# 4. BOOKMARK READER (Bookmarks Tab)
 
def stem_bookmark_id(filename: str) -> str:
    return os.path.splitext(os.path.splitext(filename)[0])[0]


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
    keys_lower = set(k.lower() for k in keys)
    for path, val in walk_values(obj):
        if path and path[-1].lower() in keys_lower:
            return val
    return None


def stringify_list(values: List[Any]) -> str:
    return "[" + ",".join(f"'{v}'" if v is not None else "null" for v in values) + "]"


def extract_entity_property(filter_obj: Dict[str, Any]) -> Tuple[str, str]:
    """Uses format_entity_with_spaces to match Pages tab formatting."""
    entity = find_first_value(filter_obj, ["Entity", "Source", "Table"]) or ""
    entity = format_entity_with_spaces(entity)
    prop = find_first_value(filter_obj, ["Property", "Column", "Field"]) or ""
    return entity, prop


def extract_values(filter_obj: Dict[str, Any]) -> List[Any]:
    out = []
    expr = find_first_value(filter_obj, ["expression"]) or {}
    values_blocks = []

    if isinstance(expr, dict):
        expr_vals = _get_any(expr, ["Values"])
        if isinstance(expr_vals, list):
            values_blocks.extend(expr_vals)

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
            if isinstance(lit, dict) and _has_any(lit, ["Value"]):
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
    return "; ".join(parts) if parts else ""


def flatten_values(values_block: Any) -> List[Any]:
    out = []
    if isinstance(values_block, list):
        for item in values_block:
            out.extend(flatten_values(item))
    elif isinstance(values_block, dict):
        lit = _get_any(values_block, ["Literal"])
        if isinstance(lit, dict) and _has_any(lit, ["Value"]):
            out.append(str(_get_any(lit, ["Value"])).strip("'"))
    return out


def extract_slicer_selections(vdata: Dict[str, Any]) -> str:
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

    return "; ".join(selections) if selections else ""


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


def extract_bookmark_visual_rows(bookmark: Dict[str, Any], filename: str) -> List[Dict[str, Any]]:
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

    target_blocks = []
    if isinstance(sections, dict):
        if active and active in sections:
            target_blocks.append(sections[active])
        else:
            target_blocks.extend(sections.values())

    rows = []
    for block in target_blocks:
        visuals = _get_any(block, ["visualContainers"]) or {}
        if isinstance(visuals, dict):
            items_to_process = list(visuals.items())
        elif isinstance(visuals, list):
            items_to_process = [(v.get("id") or v.get("name") or "", v) for v in visuals]
        else:
            items_to_process = []

        for vid, vdata in items_to_process:
            if not vid:
                continue

            vtype = (
                get_nested(vdata, ["singleVisual", "visualType"]) or
                _get_any(vdata, ["visualType"]) or
                _get_any(vdata, ["type"]) or
                ""
            )

            filt_block = _get_any(vdata, ["filters"]) or {}
            vfilters = _get_any(filt_block, ["byExpr"]) or []
            applied_filters_str = summarize_filters([x for x in vfilters if isinstance(x, dict)])

            slicer_selections = extract_slicer_selections(vdata) if _eq_ci(vtype, "slicer") else ""
            selected_flag = "Yes" if str(vid) in selected_visuals else "No"
            mode_value = extract_visual_mode(vdata)

            rows.append({
                "Bookmark ID": bookmark_id,
                "Bookmark Name": bookmark_name,
                "Visual ID": str(vid),
                "Visual Type": vtype,
                "Selected Visual": selected_flag,
                "Mode": mode_value if mode_value else None,
                "Applied Filters": applied_filters_str if applied_filters_str else None,
                "Slicer Selections": slicer_selections if slicer_selections else None
            })

    return rows


def parse_bookmarks_folder(bookmarks_folder: Optional[Path]) -> List[Dict[str, Any]]:
    if not bookmarks_folder or not bookmarks_folder.exists():
        return []
    all_rows = []
    for root, _, files in os.walk(bookmarks_folder):
        for filename in files:
            if filename.lower().endswith(".json"):
                try:
                    with open(os.path.join(root, filename), "r", encoding="utf-8-sig") as f:
                        bookmark = json.load(f)
                    all_rows.extend(extract_bookmark_visual_rows(bookmark, filename))
                except Exception as e:
                    print(f"[!] Error reading bookmark {filename}: {e}")
    return all_rows


def build_bookmark_id_to_name(bookmark_rows: List[Dict[str, Any]]) -> Dict[str, str]:
    m = {}
    for r in bookmark_rows:
        bid = _norm_id(r.get("Bookmark ID", ""))
        bname = str(r.get("Bookmark Name", "")).strip()
        if bid and bid not in m:
            m[bid] = bname
    return m

 
# 5. PAGES READER (Pages Tab)
 
def build_alias_map(visual_data: Dict[str, Any]) -> Dict[str, str]:
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


def to_table_column_from_fieldlike(field_like: dict, alias_map: Optional[dict] = None) -> Optional[str]:
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


def extract_literals(values_node: Any) -> List[Any]:
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


def is_valid_filter_value(v: Any) -> bool:
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


def stringify_value(v: Any) -> str:
    s = str(v)
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1]
    s = s.replace("'", "''")
    return f"'{s}'"


def extract_visual_columns(visual_data: Dict[str, Any], alias_map: Dict[str, str]) -> List[str]:
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


def find_table_column_in_condition(condition_node: Any, alias_map: Dict[str, str]) -> Optional[str]:
    for node in walk(condition_node):
        if isinstance(node, dict):
            col = _get_any(node, ["Column"])
            if isinstance(col, dict):
                tc = to_table_column_from_fieldlike(col, alias_map)
                if tc:
                    return tc
    return None


def parse_operator_and_values(cond_node: Any) -> Tuple[Optional[str], List[Any]]:
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
            if inner_op in ("IN", "="):
                return "NOT IN", inner_vals
            if inner_op == "BETWEEN":
                return "NOT BETWEEN", inner_vals
            return "NOT " + inner_op, inner_vals

    expr_node = _get_any(cond_node, ["Expression"])
    if isinstance(expr_node, dict):
        return parse_operator_and_values(expr_node)

    return None, []


def extract_visual_filters(visual_data: Dict[str, Any], alias_map: Dict[str, str]) -> List[str]:
    results = []

    def handle_filter_like(filter_like, field_hint=None):
        where = _get_any(filter_like, ["Where"])
        where_items = where if isinstance(where, list) else ([where] if isinstance(where, dict) else [])

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


def _literal_or_string(value_node: Any) -> Optional[str]:
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


def find_visual_actions(visual_data: Dict[str, Any]) -> List[str]:
    actions = []
    for node in walk(visual_data):
        if not isinstance(node, dict):
            continue

        visual_link = _get_any(node, ["visualLink", "action", "actions"])
        if not visual_link:
            continue

        items = visual_link if isinstance(visual_link, list) else [visual_link]
        for link_item in items:
            if not isinstance(link_item, dict):
                continue

            props = _get_any(link_item, ["properties"]) or link_item
            type_val = _literal_or_string(_get_any(props, ["type"]))
            type_val_norm = _norm_text(type_val)

            tooltip_id = _literal_or_string(_get_any(props, ["tooltip"]))
            if tooltip_id:
                actions.append(f"Tooltip: {tooltip_id}")

            if type_val_norm == _norm_text("pagenavigation"):
                nav_id = _literal_or_string(_get_any(props, ["navigationSection"]))
                actions.append(f"Page Navigation: {nav_id}" if nav_id else "Page Navigation")

            elif type_val_norm == _norm_text("bookmark"):
                bookmark_id = _literal_or_string(_get_any(props, ["bookmark"]))
                actions.append(f"Bookmark: {bookmark_id}" if bookmark_id else "Bookmark")

    return dedupe_preserve_order(actions)


def extract_visual_info(visual_json_path: Path) -> Tuple[str, str, str, str, str, str, str]:
    with open(visual_json_path, "r", encoding="utf-8-sig") as f:
        visual_data = json.load(f)

    alias_map = build_alias_map(visual_data)
    visual_id = _get_any(visual_data, ["name", "id"]) or ""
    visual_type = (
        get_nested(visual_data, ["visual", "visualType"]) or
        _get_any(visual_data, ["visualType", "type"]) or
        ""
    )

    action_type = "; ".join(find_visual_actions(visual_data))
    visual_columns = extract_visual_columns(visual_data, alias_map)
    visual_filters = extract_visual_filters(visual_data, alias_map)

    # Resolve title and classify its source
    visual_title, visual_source = resolve_visual_title_and_source(
        visual_data, str(visual_type), visual_columns
    )

    return (
        str(visual_id),
        str(visual_title),
        str(visual_source),
        str(visual_type),
        str(action_type),
        "; ".join(visual_columns),
        "; ".join(visual_filters),
    )


def extract_page_info(page_json_path: Path) -> Tuple[str, str]:
    with open(page_json_path, "r", encoding="utf-8-sig") as f:
        page_data = json.load(f)

    page_id = _get_any(page_data, ["name", "id"]) or ""
    page_name = _get_any(page_data, ["displayName"]) or ""
    return str(page_id), str(page_name)


def parse_pages_folder(pages_folder: Optional[Path]) -> List[Dict[str, Any]]:
    if not pages_folder or not pages_folder.exists():
        return []
    rows = []
    for page_folder in os.listdir(pages_folder):
        page_path = pages_folder / page_folder
        if not page_path.is_dir():
            continue

        page_json_path = page_path / "page.json"
        if not page_json_path.exists():
            continue

        page_id, page_name = extract_page_info(page_json_path)
        visuals_folder = page_path / "visuals"

        if visuals_folder.exists():
            for visual_item in os.listdir(visuals_folder):
                v_path = visuals_folder / visual_item
                target_json = None

                if v_path.is_dir() and (v_path / "visual.json").exists():
                    target_json = v_path / "visual.json"
                elif v_path.is_file() and v_path.suffix.lower() == ".json":
                    target_json = v_path

                if target_json:
                    try:
                        v_id, v_title, v_source, v_type, action_type, v_cols, v_filters = extract_visual_info(target_json)
                        rows.append({
                            "Page ID": page_id,
                            "Page Name": page_name,
                            "Visual ID": v_id,
                            "Visual Title": v_title,
                            "Visual Name Source": v_source,
                            "Visual Type": v_type,
                            "Action Type": action_type if action_type else None,
                            "Visual Columns": v_cols if v_cols else None,
                            "Visual Filters": v_filters if v_filters else None
                        })
                    except Exception as e:
                        print(f"[!] Error parsing visual {target_json.name}: {e}")
    return rows


def build_page_maps(pages_rows: List[Dict[str, Any]]) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str], Dict[str, str]]:
    page_id_to_name = {}
    visual_id_to_page_name = {}
    visual_id_to_title = {}
    visual_id_to_source = {}

    for r in pages_rows:
        pid = _norm_id(r.get("Page ID", ""))
        pname = r.get("Page Name", "")
        vid = _norm_id(r.get("Visual ID", ""))
        vtitle = r.get("Visual Title", "")
        vsource = r.get("Visual Name Source", "")

        if pid and pid not in page_id_to_name:
            page_id_to_name[pid] = pname
        if vid and vid not in visual_id_to_page_name:
            visual_id_to_page_name[vid] = pname
        if vid and vid not in visual_id_to_title:
            visual_id_to_title[vid] = vtitle
        if vid and vid not in visual_id_to_source:
            visual_id_to_source[vid] = vsource

    return page_id_to_name, visual_id_to_page_name, visual_id_to_title, visual_id_to_source


 
# 6. ENRICHMENT & MAPPING
 

def add_action_type_names(action_type: Any,
                          page_id_to_name: Dict[str, str],
                          bookmark_id_to_name: Dict[str, str]) -> Optional[str]:
    if not action_type or pd.isna(action_type):
        return None

    parts = [p.strip() for p in str(action_type).split(";") if p.strip()]
    out_parts = []

    for p in parts:
        if ":" not in p:
            out_parts.append(p)
            continue

        label, val = p.split(":", 1)
        label = label.strip()
        action_id_raw = val.strip()
        action_id = _norm_id(action_id_raw)
        label_norm = _norm_text(label)

        if label_norm == _norm_text("tooltip"):
            page_name = page_id_to_name.get(action_id, "")
            out_parts.append(f"Tooltip: {page_name}" if page_name else f"Tooltip: {action_id_raw}")

        elif label_norm == _norm_text("pagenavigation"):
            page_name = page_id_to_name.get(action_id, "")
            out_parts.append(f"Page Navigation: {page_name}" if page_name else f"Page Navigation: {action_id_raw}")

        elif label_norm == _norm_text("bookmark"):
            bm_name = bookmark_id_to_name.get(action_id, "")
            out_parts.append(f"Bookmark: {bm_name}" if bm_name else f"Bookmark: {action_id_raw}")
        else:
            out_parts.append(p)

    return "; ".join(out_parts) if out_parts else None


def add_page_and_visual_titles_to_bookmarks(bookmark_rows: List[Dict[str, Any]],
                                            visual_id_to_page_name: Dict[str, str],
                                            visual_title_map: Dict[str, str],
                                            visual_source_map: Dict[str, str]) -> List[Dict[str, Any]]:
    out = []
    for r in bookmark_rows:
        vid_raw = r.get("Visual ID", "")
        vid = _norm_id(vid_raw)

        page_name = visual_id_to_page_name.get(vid, "")
        vtitle = visual_title_map.get(vid, "")
        vsource = visual_source_map.get(vid, "")

        out.append({
            "Page Name": page_name,
            "Bookmark ID": r.get("Bookmark ID", ""),
            "Bookmark Name": r.get("Bookmark Name", ""),
            "Visual ID": vid_raw,
            "Visual Title": vtitle,
            "Visual Name Source": vsource,
            "Visual Type": r.get("Visual Type", ""),
            "Selected Visual": r.get("Selected Visual", "No"),
            "Mode": r.get("Mode"),
            "Applied Filters": r.get("Applied Filters"),
            "Slicer Selections": r.get("Slicer Selections")
        })
    return out


def add_action_names_to_pages(page_rows: List[Dict[str, Any]],
                              page_id_to_name: Dict[str, str],
                              bookmark_id_to_name: Dict[str, str]) -> List[Dict[str, Any]]:
    out = []
    for r in page_rows:
        action_type = r.get("Action Type")
        action_name = add_action_type_names(action_type, page_id_to_name, bookmark_id_to_name)

        out.append({
            "Page ID": r.get("Page ID", ""),
            "Page Name": r.get("Page Name", ""),
            "Visual ID": r.get("Visual ID", ""),
            "Visual Title": r.get("Visual Title", ""),
            "Visual Name Source": r.get("Visual Name Source", ""),
            "Visual Type": r.get("Visual Type", ""),
            "Action Type": action_type,
            "Action Name": action_name,
            "Visual Columns": r.get("Visual Columns"),
            "Visual Filters": r.get("Visual Filters")
        })
    return out


 
# 7. MAIN EXECUTION PIPELINE
 

def main():
    discovery = WorkspaceDiscovery(ROOT_FOLDER).discover()

    # Step 1: Read Pages & Resolve Visuals Natively
    print("[1/4] Reading Pages & Visuals (resolving visual titles & sources from PBIP)...")
    page_rows = parse_pages_folder(discovery.pages_folder)
    print(f"      -> Retrieved {len(page_rows)} total visual definitions across pages")

    # Step 2: Read Bookmarks
    print("[2/4] Reading Bookmarks folder...")
    bookmark_rows = parse_bookmarks_folder(discovery.bookmarks_folder)
    print(f"      -> Retrieved {len(bookmark_rows)} bookmark definition rows")

    # Step 3: Build Lookups & Enrich Cross-Tab Data
    print("[3/4] Cross-mapping Bookmarks, Pages, Actions, and Visual Sources...")
    page_id_to_name, visual_id_to_page_name, visual_title_map, visual_source_map = build_page_maps(page_rows)
    bookmark_id_to_name = build_bookmark_id_to_name(bookmark_rows)

    bookmark_rows_added = add_page_and_visual_titles_to_bookmarks(
        bookmark_rows, visual_id_to_page_name, visual_title_map, visual_source_map
    )
    page_rows_added = add_action_names_to_pages(
        page_rows, page_id_to_name, bookmark_id_to_name
    )

    df_bookmarks = pd.DataFrame(
        bookmark_rows_added,
        columns=[
            "Page Name",
            "Bookmark ID", "Bookmark Name",
            "Visual ID", "Visual Title", "Visual Name Source",
            "Visual Type", "Selected Visual", "Mode",
            "Applied Filters", "Slicer Selections"
        ]
    )

    df_pages = pd.DataFrame(
        page_rows_added,
        columns=[
            "Page ID", "Page Name",
            "Visual ID", "Visual Title", "Visual Name Source", "Visual Type",
            "Action Type", "Action Name",
            "Visual Columns", "Visual Filters"
        ]
    )

    # Visuals Tab: Includes Page ID, Page Name, Visual ID, Visual Title, Visual Name Source, Visual Type
    df_visuals = df_pages[[
        "Page ID", "Page Name", "Visual ID", "Visual Title", "Visual Name Source", "Visual Type"
    ]].drop_duplicates(
        subset=["Page ID", "Visual ID"]
    ).sort_values(["Page ID", "Visual ID"]).reset_index(drop=True)

    # Step 4: Export to Excel in the Root Directory
    output_filename = f"{discovery.report_base_name}_Decoded_Native_Report.xlsx"
    output_excel_path = discovery.root_path / output_filename
    print(f"[4/4] Exporting Excel workbook: {output_filename}...")

    try:
        writer = pd.ExcelWriter(output_excel_path, engine="openpyxl")
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{discovery.report_base_name}_Decoded_Native_{ts}.xlsx"
        output_excel_path = discovery.root_path / output_filename
        print(f"  [!] Primary file locked. Saving as: {output_filename}")
        writer = pd.ExcelWriter(output_excel_path, engine="openpyxl")

    with writer:
        df_bookmarks.to_excel(writer, sheet_name="Bookmarks", index=False)
        df_pages.to_excel(writer, sheet_name="Pages", index=False)
        df_visuals.to_excel(writer, sheet_name="Visuals", index=False)

        # Auto-adjust column widths
        for sheet_name, ws in writer.sheets.items():
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 55)

    print("\n" + "=" * 70)
    print("SUCCESS: Native PBIP decoding complete!")
    print(f"File location: {output_excel_path}")
    print(f"Summary: Bookmarks ({len(df_bookmarks)} rows), Pages ({len(df_pages)} rows), Visuals ({len(df_visuals)} rows)")
    print("=" * 70)


if __name__ == "__main__":
    main()