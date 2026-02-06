import base64
import streamlit as st
from pathlib import Path
from openpyxl.utils import get_column_letter

from preencher_planilha import (
    extract_lines_from_pdf_file,
    extract_plan_signature,
    resolve_art_by_plan_rule,
    extract_meta_especifica_sections,
    is_analysis_template_file,
    get_analysis_items_header_info,
    parse_items,
    build_rows,
    generate_excel_bytes,
    get_template_header_info,
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Planilha Base de Teste.xlsx"
LOGO_PATH = BASE_DIR / "Logo.png"

st.set_page_config(page_title="Preenche Planilhas", page_icon="📄", layout="centered")

st.markdown(
    """
    <style>
    .header {
      display: flex;
      align-items: center;
      gap: 16px;
    }
    .header-title {
      font-size: 1.6rem !important;
      font-weight: 600;
      margin: 0;
      line-height: 1.2;
    }
    .logo-wrap {
      width: 64px;
      height: 64px;
      border-radius: 16px;
      overflow: hidden;
      border: 1px solid #e6e6e6;
      flex: 0 0 auto;
    }
    .logo-wrap img {
      width: 64px;
      height: 64px;
      object-fit: cover;
      display: block;
    }
    .brand-bar {
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      height: 6px;
      border-radius: 999px;
      overflow: hidden;
      border: 1px solid #d8dbe0;
      margin-top: 0.6rem;
      margin-bottom: 14px;
      width: 699px;
    }
    .brand-bar span:nth-child(1) { background: #00b140; }
    .brand-bar span:nth-child(2) { background: #ff1b14; }
    .brand-bar span:nth-child(3) { background: #ffd200; }
    .brand-bar span:nth-child(4) { background: #1f4bff; }
    .brand-bar span:nth-child(5) { background: #ff1b14; }
    .app-subtitle { 
      margin: 0 0 16px 0;
    }
    div[data-testid="stDownloadButton"] button {
      background: #217346;
      border: 1px solid #1e6a40;
      color: #ffffff;
    }
    div[data-testid="stDownloadButton"] button:hover {
      background: #1b5e38;
      border-color: #1b5e38;
      color: #ffffff;
    }
    .blank-cells {
      border-collapse: collapse;
      width: auto;
    }
    .blank-cells th,
    .blank-cells td {
      border: 1px solid #e6e6e6;
      padding: 8px 12px;
      text-align: left;
    }
    .blank-cells th {
      background: #fafafa;
      font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

logo_b64 = ""
if LOGO_PATH.exists():
    logo_bytes = LOGO_PATH.read_bytes()
    logo_b64 = base64.b64encode(logo_bytes).decode("ascii")

logo_html = ""
if logo_b64:
    logo_html = f"""
      <div class="logo-wrap">
        <img src="data:image/png;base64,{logo_b64}" alt="Logo" />
      </div>
    """

st.markdown(
    f"""
    <div class="header">
      {logo_html}
      <h1 class="header-title">Gerador de Planilha de Itens - FAF</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="brand-bar">
      <span></span><span></span><span></span><span></span><span></span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<p class="app-subtitle">Faça upload do PDF do Plano de Aplicação e gere a planilha preenchida automaticamente.</p>',
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader("PDF do Plano", type=["pdf"])

if "result" not in st.session_state:
    st.session_state.result = None

if st.button("Processar", type="primary", disabled=uploaded_file is None):
    if not TEMPLATE_PATH.exists():
        st.error("Planilha modelo não encontrada no servidor.")
    else:
        try:
            with st.status("Processando PDF...", expanded=True) as status:
                status.write("Lendo PDF")
                lines = extract_lines_from_pdf_file(uploaded_file)
                analysis_mode = is_analysis_template_file(TEMPLATE_PATH)

                status.write("Extraindo itens")
                parsed_items = parse_items(lines)
                if not analysis_mode and not parsed_items:
                    status.update(label="Nenhum item encontrado.", state="error")
                    st.error("Nenhum item encontrado no PDF.")
                    st.session_state.result = None
                else:
                    status.write("Montando planilha")
                    signature = extract_plan_signature(lines)
                    art_num_preferred = resolve_art_by_plan_rule(
                        signature["sigla"], signature["ano"]
                    )
                    if analysis_mode:
                        sections = extract_meta_especifica_sections(lines)
                        header_row, _, items_header_map = get_analysis_items_header_info(
                            TEMPLATE_PATH
                        )
                        rows = build_rows(parsed_items, items_header_map)
                        excel_bytes = generate_excel_bytes(
                            TEMPLATE_PATH,
                            rows=rows,
                            header_map={},
                            art_num_preferred=art_num_preferred,
                            source_lines=lines,
                        )
                        missing_cells = set()
                        missing_rows = set()
                        start_row = (header_row + 1) if header_row else 3
                        for index, row_data in enumerate(rows):
                            excel_row = start_row + index
                            for header, col_index in items_header_map.items():
                                value = row_data.get(header)
                                if value is None or value == "":
                                    cell = f"{get_column_letter(col_index)}{excel_row}"
                                    missing_cells.add(cell)
                                    missing_rows.add(excel_row)
                        st.session_state.result = {
                            "mode": "analysis",
                            "rows": rows,
                            "excel_bytes": excel_bytes,
                            "meta_counts": {s["numero_meta"]: 1 for s in sections},
                            "missing_cells": sorted(missing_cells),
                            "missing_items_count": len(missing_rows),
                            "sections_count": len(sections),
                            "items_count": len(parsed_items),
                        }
                    else:
                        _, header_map = get_template_header_info(TEMPLATE_PATH)
                        rows = build_rows(parsed_items, header_map)
                        excel_bytes = generate_excel_bytes(
                            TEMPLATE_PATH,
                            rows,
                            header_map,
                            art_num_preferred=art_num_preferred,
                            source_lines=lines,
                        )

                        meta_counts = {}
                        missing_cells = set()
                        missing_rows = set()
                        start_row = 3
                        for index, row_data in enumerate(rows):
                            meta = row_data.get("Número da Meta Específica")
                            item_num = row_data.get("Número do Item")
                            meta_counts[meta] = meta_counts.get(meta, 0) + 1
                            excel_row = start_row + index
                            for header, col_index in header_map.items():
                                value = row_data.get(header)
                                if value is None or value == "":
                                    cell = f"{get_column_letter(col_index)}{excel_row}"
                                    missing_cells.add(cell)
                                    missing_rows.add(excel_row)

                        st.session_state.result = {
                            "mode": "items",
                            "rows": rows,
                            "excel_bytes": excel_bytes,
                            "meta_counts": meta_counts,
                            "missing_cells": sorted(missing_cells),
                            "missing_items_count": len(missing_rows),
                        }
                    status.update(label="Processamento concluído.", state="complete")
        except Exception as exc:
            st.exception(exc)

result = st.session_state.result
if result:
    mode = result.get("mode", "items")
    total_items = len(result["rows"])
    total_metas = len(result["meta_counts"])
    missing_count = result["missing_items_count"]
    missing_cells = result["missing_cells"]

    st.subheader("Resumo")
    summary_cols = st.columns(3)
    if mode == "analysis":
        summary_cols[0].metric("Metas encontradas", total_metas)
        summary_cols[1].metric("Itens extraídos (PDF)", result.get("items_count", 0))
        summary_cols[2].metric("Células em branco", len(missing_cells))
    else:
        summary_cols[0].metric("Itens extraídos", total_items)
        summary_cols[1].metric("Metas encontradas", total_metas)
        summary_cols[2].metric("Itens com campos faltantes", missing_count)

    if missing_count:
        st.warning("Alguns itens possuem campos em branco. Veja os detalhes abaixo.")

    st.download_button(
        "Baixar Planilha",
        data=result["excel_bytes"],
        file_name="Planilha de Itens.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    if missing_cells:
        st.subheader("Células em branco")
        rows_html = "".join(f"<tr><td>{cell}</td></tr>" for cell in missing_cells)
        st.markdown(
            f"""
            <table class="blank-cells">
              <thead><tr><th>Célula</th></tr></thead>
              <tbody>{rows_html}</tbody>
            </table>
            """,
            unsafe_allow_html=True,
        )

    # Preview e detalhes removidos conforme solicitado.
