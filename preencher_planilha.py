#!/usr/bin/env python3
import argparse
import re
from io import BytesIO
from pathlib import Path

import pdfplumber
import openpyxl

META_RE = re.compile(r"^META ESPEC[ÍI]FICA\s+(\d+)", re.IGNORECASE)
ITEM_RE = re.compile(r"^Item\s+(\d+)\s+Planejado", re.IGNORECASE)

CAPTURE_LABELS = {
    "Art. 7º (685):": "art",
    "Bem/Serviço:": "bem",
    "Descrição:": "descricao",
    "Destinação:": "destinacao",
    "Unidade de Medida:": "unidade",
    "Qtd. Planejada:": "quantidade",
    "Natureza (ND):": "natureza",
    "Instituição:": "instituicao",
    "Valor Total:": "valor_total",
}

STOP_LABELS = set(CAPTURE_LABELS.keys()) | {
    "Cód. Senasp:",
    "Valor Originário Planejado:",
    "Valor Suplementar Planejado:",
    "Valor Rendimento Planejado:",
}

OUTPUT_HEADERS = [
    "Número da Meta Específica",
    "Número do Item",
    "Ação conforme Art. 7º da portaria nº 685",
    "Material/Serviço",
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


def parse_int(value: str):
    digits = re.sub(r"[^0-9]", "", value or "")
    return int(digits) if digits else ""


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
            text = page.extract_text() or ""
            lines.extend(clean_lines(text))
    return lines


def extract_lines_from_pdf_file(file_obj):
    file_obj.seek(0)
    lines = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(clean_lines(text))
    return lines


def parse_items(lines):
    items = []
    current_meta = None
    current_item = None
    current_lines = []

    def flush():
        nonlocal current_item, current_lines
        if current_meta is None or current_item is None:
            return
        items.append({
            "meta": current_meta,
            "item": current_item,
            "lines": current_lines[:],
        })
        current_item = None
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
            current_lines = []
            continue
        if current_item is not None:
            current_lines.append(line)

    flush()
    return items


def extract_fields(item_lines):
    fields = {key: [] for key in CAPTURE_LABELS.values()}
    current_field = None

    for line in item_lines:
        label = None
        for candidate in STOP_LABELS:
            if line.startswith(candidate):
                label = candidate
                break

        if label:
            current_field = CAPTURE_LABELS.get(label)
            if current_field:
                content = line[len(label):].strip()
                if content:
                    fields[current_field].append(content)
            else:
                current_field = None
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

def fill_worksheet(ws, rows):
    # Clear previous data (keep headers)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=len(OUTPUT_HEADERS)):
        for cell in row:
            cell.value = None

    start_row = 3
    for idx, row_data in enumerate(rows, start=start_row):
        for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
            ws.cell(row=idx, column=col_idx, value=row_data.get(header, ""))


def write_excel(template_path: Path, output_path: Path, rows):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    fill_worksheet(ws, rows)
    wb.save(output_path)


def generate_excel_bytes(template_path: Path, rows) -> bytes:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    fill_worksheet(ws, rows)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_rows(parsed_items):
    rows = []
    for item in parsed_items:
        fields = extract_fields(item["lines"])
        material = build_material(fields["bem"], fields["descricao"], fields["destinacao"])
        valor_total = strip_currency(fields["valor_total"])
        quantidade = parse_int(fields["quantidade"])
        row = {
            "Número da Meta Específica": item["meta"],
            "Número do Item": item["item"],
            "Ação conforme Art. 7º da portaria nº 685": fields["art"],
            "Material/Serviço": material,
            "Instituição": fields["instituicao"],
            "Natureza da Despesa": fields["natureza"],
            "Quantidade Planejada": quantidade,
            "Unidade de Medida": fields["unidade"],
            "Valor Planejado Total": valor_total,
            "Status do Item": "Planejado",
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

    rows = build_rows(parsed_items)
    write_excel(xlsx_path, output_path, rows)

    print(f"Itens extraídos: {len(rows)}")
    print(f"Arquivo gerado: {output_path}")


if __name__ == "__main__":
    main()
