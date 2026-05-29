#!/usr/bin/env python3
"""
CRM liviano para prospectar grupos de Facebook.

La idea es trabajar de forma asistida: tu haces las busquedas en Facebook,
copias los datos visibles y este script limpia, deduplica y exporta un Excel.
No automatiza publicaciones ni acciones dentro de Facebook.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
import unicodedata
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse


APP_NAME = "facebook_group_crm"
DEFAULT_DATA_FILE = "facebook_groups.json"
DEFAULT_EXCEL_FILE = "facebook_groups.xlsx"
DEFAULT_MIN_MEMBERS = 100

STATUS_NEW = "Nuevo"
SCHEMA_VERSION = 1

HEADERS = [
    "ID",
    "Nombre del grupo",
    "URL",
    "Miembros/seguidores",
    "Texto miembros",
    "Ubicación",
    "Categoría",
    "Consulta origen",
    "Palabras coincidentes",
    "Publicaciones/día",
    "Estado",
    "Última publicación",
    "Mensajes generados",
    "Ventas generadas",
    "Prioridad",
    "Notas",
    "Creado",
    "Actualizado",
]

STOPWORDS = {
    "de",
    "del",
    "la",
    "las",
    "los",
    "el",
    "en",
    "y",
    "o",
    "a",
    "para",
    "por",
    "con",
    "mi",
    "mis",
    "tu",
    "tus",
    "un",
    "una",
    "unos",
    "unas",
    "grupo",
    "grupos",
    "facebook",
}

METADATA_PATTERNS = [
    r"^grupo\s+(publico|p[uú]blico|privado|private|public)\b",
    r"^(public|private)\s+group\b",
    r"\b(miembros|members|seguidores|followers)\b",
    r"^(unirte|join|joined|ver grupo|view group|publicaciones|posts)\b",
    r"^nuevo$",
    r"^popular cerca de ti$",
]

URL_PATTERN = re.compile(
    r"(?P<url>(?:https?://)?(?:www\.|m\.)?facebook\.com/groups/[^\s\"')<>]+)",
    re.IGNORECASE,
)

MEMBER_PATTERN = re.compile(
    r"(?P<number>\d+(?:[.,]\d+)?)\s*"
    r"(?P<unit>miembros?|members?|seguidores?|followers?|mill[oó]n(?:es)?|mil|k|m)?\b",
    re.IGNORECASE,
)

MEMBER_CONTEXT_WORDS = ("miembro", "member", "seguidor", "follower")
MEMBER_UNITS = {"miembro", "miembros", "member", "members", "seguidor", "seguidores", "follower", "followers"}
SCALE_UNITS = {"mil", "k", "m", "millon", "millones"}
NON_MEMBER_CONTEXT_WORDS = ("publicacion", "publicaciones", "post", "posts", "dia", "semana", "mes")
DAILY_POST_PATTERNS = [
    re.compile(
        r"(?P<number>\d+(?:[.,]\d+)?)\s*(?P<unit>mil|k)?\s*"
        r"(?:publicaci[oó]n|publicaciones|posts?)\s*(?:al|por|per|a|/)\s*(?:d[ií]a|day)\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?:publicaci[oó]n|publicaciones|posts?)\s*(?:al|por|per|a|/)\s*(?:d[ií]a|day)\D{0,16}"
        r"(?P<number>\d+(?:[.,]\d+)?)(?:\s*(?P<unit>mil|k))?",
        re.IGNORECASE,
    ),
]


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_value = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    ascii_value = ascii_value.lower()
    ascii_value = re.sub(r"https?://\S+", " ", ascii_value)
    ascii_value = re.sub(r"[^a-z0-9]+", " ", ascii_value)
    return re.sub(r"\s+", " ", ascii_value).strip()


def clean_group_name(name: str) -> str:
    cleaned = re.sub(r"\s+", " ", name or "").strip()
    noise_prefixes = [
        r"foto\s+del\s+perfil\s+de\s+",
        r"foto\s+de\s+perfil\s+de\s+",
        r"imagen\s+del\s+perfil\s+de\s+",
        r"profile\s+picture\s+of\s+",
        r"profile\s+photo\s+of\s+",
    ]
    for prefix in noise_prefixes:
        cleaned = re.sub(rf"^{prefix}", "", cleaned, flags=re.IGNORECASE).strip()
    return cleaned


def normalize_name(name: str) -> str:
    text = normalize_text(clean_group_name(name))
    text = re.sub(r"\b(oficial|grupo|facebook|fb)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_url(url: str) -> str:
    if not url:
        return ""

    cleaned = url.strip().rstrip(".,;)")
    if not cleaned.startswith(("http://", "https://")):
        cleaned = "https://" + cleaned

    parsed = urlparse(cleaned)
    host = parsed.netloc.lower().replace("m.facebook.com", "www.facebook.com")
    path_parts = [part for part in parsed.path.split("/") if part]

    if len(path_parts) >= 2 and path_parts[0].lower() == "groups":
        group_key = path_parts[1]
        return f"https://www.facebook.com/groups/{group_key}"

    return f"https://{host}{parsed.path}".rstrip("/")


def extract_url(text: str) -> str:
    match = URL_PATTERN.search(text or "")
    return normalize_url(match.group("url")) if match else ""


def parse_number(number_text: str, unit: str | None) -> int | None:
    raw = (number_text or "").strip()
    unit_normalized = normalize_text(unit or "")
    if not raw:
        return None

    multiplier = 1
    if unit_normalized in {"mil", "k"}:
        multiplier = 1_000
        numeric_text = raw.replace(",", ".")
    elif unit_normalized in {"m", "millon", "millones"}:
        multiplier = 1_000_000
        numeric_text = raw.replace(",", ".")
    else:
        # "12.345 miembros" o "12,345 members" son separadores de miles.
        if re.search(r"[.,]\d{3}$", raw):
            numeric_text = raw.replace(".", "").replace(",", "")
        else:
            numeric_text = raw.replace(",", ".")

    try:
        return int(round(float(numeric_text) * multiplier))
    except ValueError:
        return None


def has_member_context(text: str) -> bool:
    normalized = normalize_text(text)
    return any(word in normalized for word in MEMBER_CONTEXT_WORDS)


def nearby_text(text: str, start: int, end: int, radius: int = 28) -> str:
    return text[max(0, start - radius) : min(len(text), end + radius)]


def member_candidate_score(line: str, match: re.Match[str], unit: str | None, value: int) -> int:
    unit_normalized = normalize_text(unit or "")
    near = normalize_text(nearby_text(line, match.start(), match.end()))
    line_normalized = normalize_text(line)
    score = 0

    if unit_normalized in MEMBER_UNITS:
        score += 8
    if unit_normalized in SCALE_UNITS:
        score += 5
    if any(word in near for word in MEMBER_CONTEXT_WORDS):
        score += 6
    elif any(word in line_normalized for word in MEMBER_CONTEXT_WORDS):
        score += 2
    if any(word in near for word in NON_MEMBER_CONTEXT_WORDS) and not any(word in near for word in MEMBER_CONTEXT_WORDS):
        score -= 5
    if value >= 100:
        score += 1

    return score


def parse_members(value: str) -> tuple[int | None, str]:
    text = (value or "").strip()
    if not text:
        return None, ""

    # Acepta tambien "12000" cuando viene desde --members.
    if re.fullmatch(r"\d+(?:[.,]\d{3})*", text):
        parsed = parse_number(text, None)
        return parsed, text

    candidates: list[tuple[int, int, str]] = []
    for line in text.splitlines() or [text]:
        line_without_urls = URL_PATTERN.sub(" ", line)
        chunks = [chunk.strip() for chunk in re.split(r"\s*[·•|]\s*", line_without_urls) if chunk.strip()]
        chunks = chunks or [line_without_urls]

        for chunk in chunks:
            chunk_has_member_context = has_member_context(chunk)
            for match in MEMBER_PATTERN.finditer(chunk):
                unit = match.group("unit")
                unit_normalized = normalize_text(unit or "")
                if not unit and not chunk_has_member_context:
                    continue
                if unit_normalized not in MEMBER_UNITS | SCALE_UNITS and not chunk_has_member_context:
                    continue

                parsed = parse_number(match.group("number"), unit)
                if parsed is None:
                    continue

                score = member_candidate_score(chunk, match, unit, parsed)
                if score <= 0:
                    continue
                candidates.append((score, parsed, match.group(0).strip()))

    if candidates:
        candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
        _, parsed, members_text = candidates[0]
        return parsed, members_text

    return None, text


def parse_daily_posts(value: str) -> tuple[int | None, str]:
    text = (value or "").strip()
    if not text:
        return None, ""

    for line in text.splitlines() or [text]:
        line_without_urls = URL_PATTERN.sub(" ", line)
        chunks = [chunk.strip() for chunk in re.split(r"\s*[·•|]\s*", line_without_urls) if chunk.strip()]
        chunks = chunks or [line_without_urls]

        for chunk in chunks:
            normalized_chunk = normalize_text(chunk)
            if not any(word in normalized_chunk for word in ["publicacion", "post"]):
                continue
            if not any(word in normalized_chunk for word in ["dia", "day"]):
                continue

            for pattern in DAILY_POST_PATTERNS:
                match = pattern.search(chunk)
                if not match:
                    continue
                parsed = parse_number(match.group("number"), match.groupdict().get("unit"))
                if parsed is not None:
                    return parsed, match.group(0).strip()

    return None, ""


def keyword_tokens(query: str) -> list[str]:
    tokens = []
    for token in normalize_text(query).split():
        if len(token) >= 3 and token not in STOPWORDS:
            tokens.append(token)
    return sorted(set(tokens))


def matched_keywords(query: str, haystack: str) -> list[str]:
    normalized_haystack = normalize_text(haystack)
    return [token for token in keyword_tokens(query) if token in normalized_haystack]


def guess_category(*texts: str) -> str:
    combined = normalize_text(" ".join(texts))
    if re.search(r"\b(auto|autos|vehiculo|vehiculos|camioneta|camionetas|moto|motos|automotora)\b", combined):
        return "Autos"
    if re.search(r"\b(parcela|parcelas|terreno|terrenos|lote|lotes|campo|campos)\b", combined):
        return "Parcelas/Terrenos"
    if re.search(r"\b(electronica|celular|celulares|iphone|notebook|computador|computadores|pc|tablet)\b", combined):
        return "Electrónica"
    if re.search(r"\b(compra|venta|vendo|compro|permuto|clasificados|marketplace)\b", combined):
        return "Compra/Venta general"
    return ""


def priority_for_members(members: int | None) -> int:
    if members is None:
        return 0
    if members >= 100_000:
        return 5
    if members >= 50_000:
        return 4
    if members >= 20_000:
        return 3
    if members >= 5_000:
        return 2
    return 1


def load_database(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"schema_version": SCHEMA_VERSION, "groups": []}
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    data.setdefault("schema_version", SCHEMA_VERSION)
    data.setdefault("groups", [])
    return data


def save_database(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def next_group_id(groups: list[dict[str, Any]]) -> str:
    numbers = []
    for group in groups:
        match = re.match(r"grp_(\d+)$", str(group.get("id", "")))
        if match:
            numbers.append(int(match.group(1)))
    return f"grp_{max(numbers, default=0) + 1:04d}"


def find_existing(groups: list[dict[str, Any]], record: dict[str, Any]) -> dict[str, Any] | None:
    url = record.get("normalized_url") or ""
    name = record.get("normalized_name") or ""

    for group in groups:
        if url and url == group.get("normalized_url"):
            return group

    if name:
        for group in groups:
            if name == group.get("normalized_name"):
                return group

    return None


def below_min_members(record: dict[str, Any], min_members: int) -> bool:
    members = record.get("members")
    return members is not None and members < min_members


def merge_record(existing: dict[str, Any], incoming: dict[str, Any]) -> bool:
    changed = False

    fill_if_empty = [
        "name",
        "url",
        "normalized_url",
        "normalized_name",
        "members_text",
        "daily_posts_text",
        "location",
        "category",
        "notes",
    ]
    for field in fill_if_empty:
        if incoming.get(field) and not existing.get(field):
            existing[field] = incoming[field]
            changed = True

    incoming_members = incoming.get("members")
    current_members = existing.get("members")
    if incoming_members and (not current_members or incoming_members > current_members):
        existing["members"] = incoming_members
        existing["members_text"] = incoming.get("members_text", existing.get("members_text", ""))
        existing["priority"] = max(existing.get("priority", 0), priority_for_members(incoming_members))
        changed = True

    incoming_daily_posts = incoming.get("daily_posts")
    current_daily_posts = existing.get("daily_posts")
    if incoming_daily_posts is not None and (
        current_daily_posts is None or incoming_daily_posts > current_daily_posts
    ):
        existing["daily_posts"] = incoming_daily_posts
        existing["daily_posts_text"] = incoming.get("daily_posts_text", existing.get("daily_posts_text", ""))
        changed = True

    for field in ["source_query"]:
        incoming_value = incoming.get(field)
        existing_value = existing.get(field, "")
        if incoming_value and incoming_value not in existing_value.split(" | "):
            existing[field] = " | ".join(part for part in [existing_value, incoming_value] if part)
            changed = True

    existing_keywords = set(existing.get("matched_keywords", []))
    incoming_keywords = set(incoming.get("matched_keywords", []))
    if not incoming_keywords.issubset(existing_keywords):
        existing["matched_keywords"] = sorted(existing_keywords | incoming_keywords)
        changed = True

    if changed:
        existing["updated_at"] = now_iso()
    return changed


def build_record(
    *,
    name: str,
    url: str = "",
    members: str = "",
    query: str = "",
    location: str = "",
    category: str = "",
    status: str = STATUS_NEW,
    notes: str = "",
    raw_text: str = "",
) -> dict[str, Any]:
    clean_name = clean_group_name(name)
    clean_url = normalize_url(url) if url else extract_url(raw_text)
    parsed_members, members_text = parse_members(members or raw_text)
    daily_posts, daily_posts_text = parse_daily_posts(raw_text)
    category_value = category or guess_category(query, clean_name, raw_text)
    timestamp = now_iso()

    record = {
        "id": "",
        "name": clean_name,
        "url": clean_url,
        "normalized_url": clean_url,
        "normalized_name": normalize_name(clean_name),
        "members": parsed_members,
        "members_text": members_text,
        "daily_posts": daily_posts,
        "daily_posts_text": daily_posts_text,
        "location": location.strip(),
        "category": category_value,
        "source_query": query.strip(),
        "matched_keywords": matched_keywords(query, f"{clean_name}\n{raw_text}") if query else [],
        "status": status,
        "last_published_at": "",
        "messages_generated": 0,
        "sales_generated": 0,
        "priority": priority_for_members(parsed_members),
        "notes": notes.strip(),
        "created_at": timestamp,
        "updated_at": timestamp,
    }
    return record


def add_or_update(data: dict[str, Any], record: dict[str, Any], min_members: int = DEFAULT_MIN_MEMBERS) -> str:
    if not record.get("name"):
        return "ignored"

    if below_min_members(record, min_members):
        return "too_small"

    groups = data.setdefault("groups", [])
    existing = find_existing(groups, record)
    if existing:
        return "updated" if merge_record(existing, record) else "duplicate"

    record["id"] = next_group_id(groups)
    groups.append(record)
    return "added"


def is_metadata_line(line: str) -> bool:
    normalized = normalize_text(line)
    if not normalized:
        return True
    if URL_PATTERN.search(line):
        return True
    return any(re.search(pattern, normalized, re.IGNORECASE) for pattern in METADATA_PATTERNS)


def probable_name_line(line: str) -> bool:
    cleaned = line.strip()
    if len(cleaned) < 3:
        return False
    if is_metadata_line(cleaned):
        return False
    if cleaned.count(" ") > 14:
        return False
    return True


def parse_pipe_line(line: str, query: str, location: str, category: str) -> dict[str, Any] | None:
    parts = [part.strip() for part in line.split("|")]
    if not parts or not parts[0]:
        return None

    name = parts[0]
    url = parts[1] if len(parts) > 1 else ""
    members = parts[2] if len(parts) > 2 else ""
    notes = parts[3] if len(parts) > 3 else ""
    raw_text = line

    return build_record(
        name=name,
        url=url,
        members=members,
        query=query,
        location=location,
        category=category,
        notes=notes,
        raw_text=raw_text,
    )


def block_to_record(block: str, query: str, location: str, category: str) -> dict[str, Any] | None:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    if not lines:
        return None

    name = ""
    for line in lines:
        if probable_name_line(line):
            name = line
            break

    if not name:
        return None

    raw_text = "\n".join(lines)
    return build_record(
        name=name,
        members=raw_text,
        query=query,
        location=location,
        category=category,
        raw_text=raw_text,
    )


def blocks_from_text(text: str) -> list[str]:
    stripped = text.strip()
    if not stripped:
        return []

    blank_blocks = [block.strip() for block in re.split(r"\n\s*\n+", stripped) if block.strip()]
    if len(blank_blocks) > 1:
        return blank_blocks

    lines = [line.strip() for line in stripped.splitlines() if line.strip()]
    blocks: list[list[str]] = []
    current: list[str] = []

    for line in lines:
        starts_new = probable_name_line(line) and current and any(
            URL_PATTERN.search(item) or parse_members(item)[0] is not None for item in current
        )
        if starts_new:
            blocks.append(current)
            current = [line]
        else:
            current.append(line)

    if current:
        blocks.append(current)

    return ["\n".join(block) for block in blocks]


def parse_import_text(text: str, query: str, location: str, category: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    pipe_lines = [line for line in text.splitlines() if "|" in line and line.strip()]

    if pipe_lines:
        for line in pipe_lines:
            record = parse_pipe_line(line, query, location, category)
            if record:
                records.append(record)
        return records

    for block in blocks_from_text(text):
        record = block_to_record(block, query, location, category)
        if record:
            records.append(record)

    return records


def group_to_row(group: dict[str, Any]) -> list[Any]:
    return [
        group.get("id", ""),
        group.get("name", ""),
        group.get("url", ""),
        group.get("members", ""),
        group.get("members_text", ""),
        group.get("location", ""),
        group.get("category", ""),
        group.get("source_query", ""),
        ", ".join(group.get("matched_keywords", [])),
        group.get("daily_posts", ""),
        group.get("status", ""),
        group.get("last_published_at", ""),
        group.get("messages_generated", 0),
        group.get("sales_generated", 0),
        group.get("priority", 0),
        group.get("notes", ""),
        group.get("created_at", ""),
        group.get("updated_at", ""),
    ]


def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xlsx_cell(row_index: int, col_index: int, value: Any, header: bool = False) -> str:
    cell_ref = f"{column_letter(col_index)}{row_index}"
    style = ' s="1"' if header else ""

    if value is None:
        return f'<c r="{cell_ref}"{style}/>'

    if isinstance(value, int | float) and not isinstance(value, bool):
        return f'<c r="{cell_ref}"{style}><v>{value}</v></c>'

    escaped = html.escape(str(value), quote=False)
    preserve = ' xml:space="preserve"' if str(value).strip() != str(value) else ""
    return f'<c r="{cell_ref}" t="inlineStr"{style}><is><t{preserve}>{escaped}</t></is></c>'


def write_xlsx(path: Path, headers: list[str], rows: list[list[Any]]) -> None:
    all_rows = [headers, *rows]
    max_col = len(headers)
    max_row = len(all_rows)
    dimension = f"A1:{column_letter(max_col)}{max_row}"

    sheet_rows = []
    for row_index, row in enumerate(all_rows, start=1):
        cells = [
            xlsx_cell(row_index, col_index, value, header=row_index == 1)
            for col_index, value in enumerate(row, start=1)
        ]
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    col_widths = [12, 34, 42, 18, 18, 16, 18, 28, 24, 18, 18, 18, 18, 16, 10, 34, 22, 22]
    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(col_widths, start=1)
    )

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>{cols_xml}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Grupos" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    root_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>{APP_NAME}</dc:creator>
  <cp:lastModifiedBy>{APP_NAME}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
</Properties>'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml)
        xlsx.writestr("_rels/.rels", root_rels_xml)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        xlsx.writestr("xl/styles.xml", styles_xml)
        xlsx.writestr("docProps/core.xml", core_xml)
        xlsx.writestr("docProps/app.xml", app_xml)


def export_excel(data: dict[str, Any], output: Path) -> None:
    rows = [group_to_row(group) for group in sorted_groups(data["groups"], "priority")]
    write_xlsx(output, HEADERS, rows)


def sorted_groups(groups: list[dict[str, Any]], sort_key: str) -> list[dict[str, Any]]:
    if sort_key == "members":
        return sorted(groups, key=lambda item: item.get("members") or 0, reverse=True)
    if sort_key == "name":
        return sorted(groups, key=lambda item: normalize_text(item.get("name", "")))
    if sort_key == "oldest":
        return sorted(groups, key=lambda item: item.get("last_published_at") or "0000-00-00")
    return sorted(
        groups,
        key=lambda item: ((item.get("priority") or 0), (item.get("members") or 0), item.get("name", "")),
        reverse=True,
    )


def print_table(groups: list[dict[str, Any]], limit: int) -> None:
    rows = groups[:limit] if limit else groups
    if not rows:
        print("No hay grupos guardados todavia.")
        return

    print(f"{'ID':<9} {'Miembros':>10} {'Prio':>4} {'Estado':<22} Nombre")
    print("-" * 90)
    for group in rows:
        members = group.get("members")
        members_text = f"{members:,}".replace(",", ".") if members else "-"
        print(
            f"{group.get('id', ''):<9} {members_text:>10} {group.get('priority', 0):>4} "
            f"{group.get('status', '')[:22]:<22} {group.get('name', '')}"
        )


def read_stdin_or_file(path: str | None) -> str:
    if path:
        return Path(path).read_text(encoding="utf-8")

    if sys.stdin.isatty():
        print("Pega los resultados y termina con Ctrl+D:")
    return sys.stdin.read()


def empty_counters() -> dict[str, int]:
    return {"added": 0, "updated": 0, "duplicate": 0, "ignored": 0, "too_small": 0}


def get_min_members(args: argparse.Namespace) -> int:
    return int(getattr(args, "min_members", DEFAULT_MIN_MEMBERS) or DEFAULT_MIN_MEMBERS)


def format_counters(prefix: str, counters: dict[str, int], min_members: int) -> str:
    return (
        f"{prefix}: "
        f"{counters['added']} nuevos, {counters['updated']} actualizados, "
        f"{counters['duplicate']} duplicados, {counters['ignored']} ignorados, "
        f"{counters['too_small']} omitidos por menos de {min_members} miembros/seguidores."
    )


def command_add(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    data = load_database(data_path)
    min_members = get_min_members(args)
    record = build_record(
        name=args.name,
        url=args.url or "",
        members=args.members or "",
        query=args.query or "",
        location=args.location or "",
        category=args.category or "",
        status=args.status or STATUS_NEW,
        notes=args.notes or "",
        raw_text=" ".join(part for part in [args.name, args.url or "", args.members or "", args.query or ""] if part),
    )
    result = add_or_update(data, record, min_members)
    save_database(data_path, data)
    export_excel(data, Path(args.excel))
    if result == "too_small":
        print(f"omitido por menos de {min_members} miembros/seguidores: {record['name']}")
    else:
        print(f"{result}: {record['name']}")
    print(f"Excel actualizado: {Path(args.excel).resolve()}")
    return 0


def command_import_text(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    data = load_database(data_path)
    min_members = get_min_members(args)
    text = read_stdin_or_file(args.file)
    records = parse_import_text(text, args.query or "", args.location or "", args.category or "")

    counters = empty_counters()
    for record in records:
        result = add_or_update(data, record, min_members)
        counters[result] += 1

    save_database(data_path, data)
    export_excel(data, Path(args.excel))
    print(format_counters("Importacion lista", counters, min_members))
    print(f"Excel actualizado: {Path(args.excel).resolve()}")
    return 0


def print_capture_help() -> None:
    print(
        """
Comandos del modo captura:
  :save              Guarda la tanda pegada y actualiza el Excel
  :q                 Guarda lo pendiente y sale
  :query TEXTO       Cambia la busqueda origen
  :location TEXTO    Cambia la comuna/zona
  :category TEXTO    Cambia la categoria
  :list              Muestra el ranking actual
  :summary           Muestra resumen operativo
  :help              Muestra esta ayuda

Tambien puedes terminar una tanda con una linea que diga solo ---.
Formato rapido recomendado:
  Nombre del grupo | URL | 56 mil miembros
""".strip()
    )


def prompt_if_empty(label: str, value: str) -> str:
    if value:
        return value
    if not sys.stdin.isatty():
        return ""
    try:
        return input(f"{label}: ").strip()
    except EOFError:
        return ""


def save_batch(
    *,
    data: dict[str, Any],
    data_path: Path,
    excel_path: Path,
    batch_lines: list[str],
    query: str,
    location: str,
    category: str,
    min_members: int,
) -> dict[str, int]:
    text = "\n".join(batch_lines).strip()
    counters = empty_counters()
    if not text:
        return counters

    records = parse_import_text(text, query, location, category)
    for record in records:
        result = add_or_update(data, record, min_members)
        counters[result] += 1

    save_database(data_path, data)
    export_excel(data, excel_path)
    return counters


def print_save_result(counters: dict[str, int], excel_path: Path, min_members: int) -> None:
    total = sum(counters.values())
    if total == 0:
        print("No habia datos nuevos para guardar.")
        return
    print(format_counters("Guardado", counters, min_members))
    print(f"Excel actualizado: {excel_path.resolve()}")


def command_capture(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    excel_path = Path(args.excel)
    data = load_database(data_path)
    min_members = get_min_members(args)

    query = prompt_if_empty("Busqueda actual", args.query or "")
    location = prompt_if_empty("Comuna/zona", args.location or "")
    category = args.category or ""

    print()
    print("Modo captura iniciado. Pega resultados, usa :save para guardar y :q para salir.")
    print(f"Filtro activo: se omiten grupos con menos de {min_members} miembros/seguidores detectados.")
    print_capture_help()
    print()

    batch_lines: list[str] = []
    while True:
        try:
            line = input("> ")
        except EOFError:
            line = ":q"

        command = line.strip()
        if command == "---":
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            batch_lines.clear()
            print_save_result(counters, excel_path, min_members)
            continue

        if not command.startswith(":"):
            batch_lines.append(line)
            continue

        if command in {":q", ":quit", ":exit"}:
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            print_save_result(counters, excel_path, min_members)
            print("Captura finalizada.")
            return 0

        if command == ":save":
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            batch_lines.clear()
            print_save_result(counters, excel_path, min_members)
            continue

        if command.startswith(":query "):
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            batch_lines.clear()
            print_save_result(counters, excel_path, min_members)
            query = command.removeprefix(":query ").strip()
            print(f"Busqueda actual: {query}")
            continue

        if command.startswith(":location "):
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            batch_lines.clear()
            print_save_result(counters, excel_path, min_members)
            location = command.removeprefix(":location ").strip()
            print(f"Comuna/zona actual: {location}")
            continue

        if command.startswith(":category "):
            counters = save_batch(
                data=data,
                data_path=data_path,
                excel_path=excel_path,
                batch_lines=batch_lines,
                query=query,
                location=location,
                category=category,
                min_members=min_members,
            )
            batch_lines.clear()
            print_save_result(counters, excel_path, min_members)
            category = command.removeprefix(":category ").strip()
            print(f"Categoria actual: {category}")
            continue

        if command == ":list":
            print_table(sorted_groups(data["groups"], "priority"), 20)
            continue

        if command == ":summary":
            total = len(data["groups"])
            unpublished = sum(1 for group in data["groups"] if not group.get("last_published_at"))
            print(f"Grupos guardados: {total}")
            print(f"Nunca publicados: {unpublished}")
            print_table(sorted_groups(data["groups"], "priority"), 10)
            continue

        if command == ":help":
            print_capture_help()
            continue

        print("Comando no reconocido. Usa :help para ver opciones.")


def command_list(args: argparse.Namespace) -> int:
    data = load_database(Path(args.data))
    print_table(sorted_groups(data["groups"], args.sort), args.limit)
    return 0


def command_export(args: argparse.Namespace) -> int:
    data = load_database(Path(args.data))
    export_excel(data, Path(args.excel))
    print(f"Excel exportado: {Path(args.excel).resolve()}")
    return 0


def find_by_identifier(groups: list[dict[str, Any]], identifier: str) -> dict[str, Any] | None:
    normalized_identifier = normalize_name(identifier)
    normalized_url = normalize_url(identifier) if "facebook.com/groups/" in identifier else ""

    for group in groups:
        if identifier == group.get("id"):
            return group
        if normalized_url and normalized_url == group.get("normalized_url"):
            return group
        if normalized_identifier and normalized_identifier == group.get("normalized_name"):
            return group
    return None


def command_mark_posted(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    data = load_database(data_path)
    group = find_by_identifier(data["groups"], args.group)
    if not group:
        print(f"No encontre el grupo: {args.group}", file=sys.stderr)
        return 1

    group["last_published_at"] = args.date or date.today().isoformat()
    group["status"] = "Publicado esta semana"
    group["updated_at"] = now_iso()
    save_database(data_path, data)
    export_excel(data, Path(args.excel))
    print(f"Marcado como publicado: {group.get('name')} ({group['last_published_at']})")
    return 0


def command_summary(args: argparse.Namespace) -> int:
    data = load_database(Path(args.data))
    groups = data["groups"]
    total = len(groups)
    with_members = sum(1 for group in groups if group.get("members"))
    unpublished = [group for group in groups if not group.get("last_published_at")]
    top = sorted_groups(groups, "priority")[: args.limit]

    print(f"Grupos guardados: {total}")
    print(f"Con miembros registrados: {with_members}")
    print(f"Nunca publicados: {len(unpublished)}")
    print()
    print("Top prioridad:")
    print_table(top, args.limit)
    return 0


def command_clean_names(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    data = load_database(data_path)
    changed = 0
    for group in data["groups"]:
        old_name = group.get("name", "")
        new_name = clean_group_name(old_name)
        new_normalized_name = normalize_name(new_name)
        if new_name != old_name or group.get("normalized_name") != new_normalized_name:
            group["name"] = new_name
            group["normalized_name"] = new_normalized_name
            group["updated_at"] = now_iso()
            changed += 1

    if changed:
        save_database(data_path, data)
        export_excel(data, Path(args.excel))

    print(f"Nombres limpiados: {changed}")
    if changed:
        print(f"Excel actualizado: {Path(args.excel).resolve()}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Arma una base de grupos de Facebook con deduplicacion y exportacion a Excel."
    )
    parser.add_argument("--data", default=DEFAULT_DATA_FILE, help="Archivo JSON interno.")
    parser.add_argument("--excel", default=DEFAULT_EXCEL_FILE, help="Archivo .xlsx de salida.")
    parser.set_defaults(min_members=DEFAULT_MIN_MEMBERS)

    subparsers = parser.add_subparsers(dest="command", required=True)

    add_parser = subparsers.add_parser("add", help="Agrega un grupo manualmente.")
    add_parser.add_argument("--name", required=True, help="Nombre del grupo.")
    add_parser.add_argument("--url", default="", help="URL del grupo.")
    add_parser.add_argument("--members", default="", help='Miembros visibles, ej: "12 mil miembros".')
    add_parser.add_argument("--query", default="", help="Busqueda que origino el hallazgo.")
    add_parser.add_argument("--location", default="", help="Comuna/zona.")
    add_parser.add_argument("--category", default="", help="Categoria comercial.")
    add_parser.add_argument("--status", default=STATUS_NEW, help="Estado inicial.")
    add_parser.add_argument("--notes", default="", help="Notas libres.")
    add_parser.add_argument(
        "--min-members",
        type=int,
        default=DEFAULT_MIN_MEMBERS,
        help=f"Omite grupos con menos de esta cantidad detectada. Default: {DEFAULT_MIN_MEMBERS}.",
    )
    add_parser.set_defaults(func=command_add)

    import_parser = subparsers.add_parser("import-text", help="Importa resultados copiados o lineas separadas por |.")
    import_parser.add_argument("--file", help="Archivo .txt con resultados. Si se omite, lee desde stdin.")
    import_parser.add_argument("--query", default="", help="Busqueda que hiciste en Facebook.")
    import_parser.add_argument("--location", default="", help="Comuna/zona.")
    import_parser.add_argument("--category", default="", help="Categoria comercial.")
    import_parser.add_argument(
        "--min-members",
        type=int,
        default=DEFAULT_MIN_MEMBERS,
        help=f"Omite grupos con menos de esta cantidad detectada. Default: {DEFAULT_MIN_MEMBERS}.",
    )
    import_parser.set_defaults(func=command_import_text)

    capture_parser = subparsers.add_parser("capture", help="Modo interactivo: pega resultados y actualiza el Excel por tandas.")
    capture_parser.add_argument("--query", default="", help="Busqueda que estas haciendo en Facebook.")
    capture_parser.add_argument("--location", default="", help="Comuna/zona.")
    capture_parser.add_argument("--category", default="", help="Categoria comercial.")
    capture_parser.add_argument(
        "--min-members",
        type=int,
        default=DEFAULT_MIN_MEMBERS,
        help=f"Omite grupos con menos de esta cantidad detectada. Default: {DEFAULT_MIN_MEMBERS}.",
    )
    capture_parser.set_defaults(func=command_capture)

    list_parser = subparsers.add_parser("list", help="Muestra grupos guardados en consola.")
    list_parser.add_argument("--sort", choices=["priority", "members", "name", "oldest"], default="priority")
    list_parser.add_argument("--limit", type=int, default=30)
    list_parser.set_defaults(func=command_list)

    export_parser = subparsers.add_parser("export", help="Regenera el Excel desde el JSON.")
    export_parser.set_defaults(func=command_export)

    mark_parser = subparsers.add_parser("mark-posted", help="Marca un grupo como publicado hoy o en una fecha.")
    mark_parser.add_argument("group", help="ID, nombre exacto normalizado o URL del grupo.")
    mark_parser.add_argument("--date", default="", help="Fecha YYYY-MM-DD. Si se omite usa hoy.")
    mark_parser.set_defaults(func=command_mark_posted)

    summary_parser = subparsers.add_parser("summary", help="Resumen operativo.")
    summary_parser.add_argument("--limit", type=int, default=10)
    summary_parser.set_defaults(func=command_summary)

    clean_names_parser = subparsers.add_parser("clean-names", help="Limpia ruido comun en nombres guardados y regenera el Excel.")
    clean_names_parser.set_defaults(func=command_clean_names)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
