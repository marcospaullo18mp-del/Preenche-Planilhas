#!/usr/bin/env python3
import argparse
import re
from io import BytesIO
from pathlib import Path

import pdfplumber
import openpyxl

META_RE = re.compile(r"^META ESPEC[ÍI]FICA\s+(\d+)", re.IGNORECASE)
ITEM_RE = re.compile(
    r"^Item\s*(\d+)\s*(Planejado|Aprovado|Cancelado)?", re.IGNORECASE
)
ACTION_HEADER_KEY = "acao_art"
ACTION_HEADER_NUM_KEY = "acao_art_num"
ACTION_HEADER_PATTERN = re.compile(
    r"^Ação conforme Art\.\s*\d+º\s+da portaria nº 685$",
    re.IGNORECASE,
)
PLAN_SIGNATURE_RE = re.compile(
    r"\b([A-Z]{2})\s*-\s*([A-Z0-9]+)\s*-\s*(20\d{2})\b"
)
ART_PATTERN = re.compile(
    r"^Art\.?\s*(6|7|8)\s*º?\s*(?:\((\d+)\))?\s*:\s*(.*)",
    re.IGNORECASE,
)
ACTION_PATTERN = re.compile(r"^A[cç][aã]o:\s*(.*)", re.IGNORECASE)

CAPTURE_PATTERNS = [
    ("bem", re.compile(r"^(?:Bem|Material)/Servi[cç]o:\s*(.*)", re.IGNORECASE)),
    ("descricao", re.compile(r"^Descri[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("destinacao", re.compile(r"^Destina[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("unidade", re.compile(r"^Unidade de Medida:\s*(.*)", re.IGNORECASE)),
    ("quantidade", re.compile(r"^Qtd\.?\s*Planejada:\s*(.*)", re.IGNORECASE)),
    ("quantidade", re.compile(r"^Quantidade Planejada:\s*(.*)", re.IGNORECASE)),
    ("natureza", re.compile(r"^Natureza\s*\(ND\):\s*(.*)", re.IGNORECASE)),
    ("instituicao", re.compile(r"^Institui[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("valor_total", re.compile(r"^Valor Total:\s*(.*)", re.IGNORECASE)),
]

STOP_PATTERNS = [
    re.compile(r"^C[oó]d\.?\s*Senasp:", re.IGNORECASE),
    re.compile(r"^Valor Origin[aá]rio Planejado:", re.IGNORECASE),
    re.compile(r"^Valor Suplementar Planejado:", re.IGNORECASE),
    re.compile(r"^Valor Rendimento Planejado:", re.IGNORECASE),
]

OUTPUT_HEADERS = [
    "Número da Meta Específica",
    "Número do Item",
    "Ação conforme Art. 7º da portaria nº 685",
    "Material/Serviço",
    "Descrição",
    "Destinação",
    "Instituição",
    "Natureza da Despesa",
    "Quantidade Planejada",
    "Unidade de Medida",
    "Valor Planejado Total",
    "Status do Item",
]


def normalize(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def strip_currency(value: str) -> str:
    value = (value or "").replace("R$", "").strip()
    value = value.replace(".", "")
    value = re.sub(r"\s+", "", value)
    return value


def format_currency(value: str) -> str:
    value = strip_currency(value)
    if not value:
        return ""
    if "," in value:
        integer_part, decimal_part = value.split(",", 1)
    else:
        integer_part, decimal_part = value, "00"
    integer_part = re.sub(r"[^0-9]", "", integer_part)
    decimal_part = re.sub(r"[^0-9]", "", decimal_part)[:2].ljust(2, "0")
    integer_part = integer_part.lstrip("0") or "0"
    grouped = ""
    while integer_part:
        grouped = integer_part[-3:] + (f".{grouped}" if grouped else "")
        integer_part = integer_part[:-3]
    return f"R$ {grouped},{decimal_part}"


def parse_int(value: str):
    digits = re.sub(r"[^0-9]", "", value or "")
    return int(digits) if digits else ""


def normalize_pdf_text(text: str) -> str:
    text = text.replace("\x0c", "\n")
    text = re.sub(
        r"(META ESPEC[ÍI]FICA\s+\d+)", r"\n\1\n", text, flags=re.IGNORECASE
    )
    text = re.sub(
        r"(Item\s*\d+\s*(?:Planejado|Aprovado|Cancelado)?)",
        r"\n\1\n",
        text,
        flags=re.IGNORECASE,
    )
    return text


def clean_lines(text: str):
    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if re.match(r"^\d{2}/\d{2}/\d{4},", line):
            continue
        if "Planos de Aplicação" in line and re.search(r"\d{2}/\d{2}/\d{4}", line):
            continue
        if line.startswith("https://apps.mj.gov.br/"):
            continue
        lines.append(line)
    return lines


def extract_lines_from_pdf(pdf_path: Path):
    lines = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = normalize_pdf_text(page.extract_text() or "")
            lines.extend(clean_lines(text))
    return lines


def extract_lines_from_pdf_file(file_obj):
    file_obj.seek(0)
    lines = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = normalize_pdf_text(page.extract_text() or "")
            lines.extend(clean_lines(text))
    return lines


def extract_plan_signature(lines):
    max_lines = min(len(lines), 120)
    for idx in range(max_lines):
        line = (lines[idx] or "").strip()
        if not line:
            continue
        match = PLAN_SIGNATURE_RE.search(line.upper())
        if match:
            return {
                "sigla": match.group(2).upper(),
                "ano": int(match.group(3)),
                "raw_line": line,
            }
    return {"sigla": None, "ano": None, "raw_line": None}


def resolve_art_by_plan_rule(sigla, ano):
    if not sigla or not ano:
        return None
    sigla = str(sigla).upper()
    if sigla in {"ECV", "FISPDS", "RMVI"} and 2019 <= ano <= 2025:
        return "6"
    if sigla == "EVM" and 2023 <= ano <= 2025:
        return "7"
    if sigla in {"VPSP", "MQVPSP"} and 2019 <= ano <= 2025:
        return "8"
    return None


def parse_items(lines):
    items = []
    current_meta = None
    current_item = None
    current_status = None
    current_lines = []

    def flush():
        nonlocal current_item, current_lines, current_status
        if current_meta is None or current_item is None:
            return
        items.append({
            "meta": current_meta,
            "item": current_item,
            "status": current_status or "",
            "lines": current_lines[:],
        })
        current_item = None
        current_status = None
        current_lines = []

    for line in lines:
        meta_match = META_RE.match(line)
        if meta_match:
            flush()
            current_meta = int(meta_match.group(1))
            continue
        item_match = ITEM_RE.match(line)
        if item_match:
            flush()
            current_item = int(item_match.group(1))
            current_status = (item_match.group(2) or "").capitalize()
            current_lines = []
            continue
        if current_item is not None:
            current_lines.append(line)

    flush()
    return items


def extract_fields(item_lines):
    fields = {key: [] for key, _ in CAPTURE_PATTERNS}
    fields["acao"] = []
    fields["art"] = []
    fields["art_num"] = ""
    current_field = None

    for line in item_lines:
        matched = False
        for stop in STOP_PATTERNS:
            if stop.match(line):
                current_field = None
                matched = True
                break
        if matched:
            continue

        action_match = ACTION_PATTERN.match(line)
        if action_match:
            current_field = "acao"
            action_body = action_match.group(1).strip()
            if action_body:
                fields[current_field].append(action_body)
            continue

        art_match = ART_PATTERN.match(line)
        if art_match:
            current_field = "art"
            art_num = art_match.group(1)
            art_body = art_match.group(3).strip()
            if art_body:
                fields[current_field].append(art_body)
            fields["art_num"] = art_num
            continue

        for field, pattern in CAPTURE_PATTERNS:
            match = pattern.match(line)
            if match:
                current_field = field
                content = match.group(1).strip()
                if content:
                    fields[current_field].append(content)
                matched = True
                break
        if matched:
            continue

        if current_field:
            fields[current_field].append(line)

    for key in fields:
        fields[key] = normalize(" ".join(fields[key]))

    return fields


def build_material(bem, descricao, destinacao):
    parts = []
    if bem:
        parts.append(f"Bem/Serviço: {bem}")
    if descricao:
        parts.append(f"Descrição: {descricao}")
    if destinacao:
        parts.append(f"Destinação: {destinacao}")
    return " | ".join(parts)

def fill_worksheet(ws, rows, header_map):
    # Clear previous data (keep headers)
    max_col = max(header_map.values()) if header_map else ws.max_column
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=max_col):
        for cell in row:
            cell.value = None

    start_row = 3
    for idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            ws.cell(row=idx, column=col_idx, value=row_data.get(header, ""))


def get_template_header_info(template_path: Path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    headers = []
    header_map = {}
    for cell in ws[2]:
        if cell.value:
            header = str(cell.value).strip()
            headers.append(header)
            key = ACTION_HEADER_KEY if ACTION_HEADER_PATTERN.match(header) else header
            if key not in header_map:
                header_map[key] = cell.column
    if not header_map:
        headers = OUTPUT_HEADERS[:]
        header_map = {}
        for idx, header in enumerate(headers):
            key = ACTION_HEADER_KEY if ACTION_HEADER_PATTERN.match(header) else header
            if key not in header_map:
                header_map[key] = idx + 1
    return headers, header_map


def update_action_header(ws, rows, header_map, art_num_preferred=None):
    col_idx = header_map.get(ACTION_HEADER_KEY)
    if not col_idx or not rows:
        return
    art_num = art_num_preferred
    if not art_num:
        art_num = rows[0].get(ACTION_HEADER_NUM_KEY)
    if not art_num:
        return
    if str(art_num) not in {"6", "7", "8"}:
        return
    ws.cell(
        row=2,
        column=col_idx,
        value=f"Ação conforme Art. {art_num}º da portaria nº 685",
    )


def write_excel(
    template_path: Path,
    output_path: Path,
    rows,
    header_map,
    art_num_preferred=None,
):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    update_action_header(ws, rows, header_map, art_num_preferred=art_num_preferred)
    fill_worksheet(ws, rows, header_map)
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"
    ws.sheet_view.zoomScale = 100
    wb.save(output_path)


def generate_excel_bytes(template_path: Path, rows, header_map, art_num_preferred=None) -> bytes:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    update_action_header(ws, rows, header_map, art_num_preferred=art_num_preferred)
    fill_worksheet(ws, rows, header_map)
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"
    ws.sheet_view.zoomScale = 100
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_rows(parsed_items, header_map):
    has_descricao = "Descrição" in header_map
    has_destinacao = "Destinação" in header_map
    rows = []
    for item in parsed_items:
        fields = extract_fields(item["lines"])
        if has_descricao or has_destinacao:
            material = fields["bem"]
        else:
            material = build_material(fields["bem"], fields["descricao"], fields["destinacao"])
        valor_total = format_currency(fields["valor_total"])
        quantidade = parse_int(fields["quantidade"])
        row = {
            "Número da Meta Específica": item["meta"],
            "Número do Item": item["item"],
            ACTION_HEADER_KEY: fields["acao"] or fields["art"],
            ACTION_HEADER_NUM_KEY: fields["art_num"],
            "Material/Serviço": material,
            "Descrição": fields["descricao"] if has_descricao else "",
            "Destinação": fields["destinacao"] if has_destinacao else "",
            "Instituição": fields["instituicao"],
            "Natureza da Despesa": fields["natureza"],
            "Quantidade Planejada": quantidade,
            "Unidade de Medida": fields["unidade"],
            "Valor Planejado Total": valor_total,
            "Status do Item": item.get("status") or "Planejado",
        }
        rows.append(row)
    return rows


def main():
    parser = argparse.ArgumentParser(description="Preenche planilha a partir do PDF.")
    parser.add_argument("--pdf", default="Planos de Aplicação.pdf", help="PDF de entrada")
    parser.add_argument("--xlsx", default="Itens NT.xlsx", help="Planilha modelo")
    parser.add_argument("--output", default="Itens NT - preenchido.xlsx", help="Planilha de saída")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    xlsx_path = Path(args.xlsx)
    output_path = Path(args.output)

    if not pdf_path.exists():
        raise SystemExit(f"PDF não encontrado: {pdf_path}")
    if not xlsx_path.exists():
        raise SystemExit(f"Planilha não encontrada: {xlsx_path}")

    lines = extract_lines_from_pdf(pdf_path)
    parsed_items = parse_items(lines)
    if not parsed_items:
        raise SystemExit("Nenhum item encontrado no PDF.")

    signature = extract_plan_signature(lines)
    art_num_preferred = resolve_art_by_plan_rule(signature["sigla"], signature["ano"])
    _, header_map = get_template_header_info(xlsx_path)
    rows = build_rows(parsed_items, header_map)
    write_excel(
        xlsx_path,
        output_path,
        rows,
        header_map,
        art_num_preferred=art_num_preferred,
    )

    print(f"Itens extraídos: {len(rows)}")
    print(f"Arquivo gerado: {output_path}")


if __name__ == "__main__":
    main()
