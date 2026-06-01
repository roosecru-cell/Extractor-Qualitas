import streamlit as st
import pdfplumber
import re
import os
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import groupby

st.set_page_config(
    page_title="Extractor Refacciones – Quálitas",
    page_icon="🔧",
    layout="centered"
)

st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    h1 { color: #1a3a5c; }
    .stButton>button {
        background-color: #1a3a5c;
        color: white;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        font-weight: 600;
    }
    .stButton>button:hover { background-color: #c8392b; }
    .result-box {
        background: white;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        margin-top: 1rem;
        border-left: 4px solid #1a3a5c;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)

st.title("🔧 Extractor de Refacciones Qualitas")
st.caption("Valuaciones Quálitas  |  Sube uno o varios PDFs")

def get_ot(filename: str) -> str:
    return os.path.splitext(filename)[0]

def extract_refacciones(pdf_bytes: bytes, ot: str) -> list[dict]:
    partidas = []
    patron = re.compile(r'^(.+?)\s+(\d{3,6})\s+\$\s*([\d,]+\.\d{2})\s*$')

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            w, h = page.width, page.height
            left = page.crop((0, 0, w * 0.52, h))
            text = left.extract_text() or ""

            in_ref = False
            for line in text.split("\n"):
                ls = line.strip()

                if re.match(r'^REFACCIONES\b', ls):
                    in_ref = True
                    continue

                if in_ref and re.match(
                    r'^(PINTURA|MANO DE OBRA|SUBTOTAL|Subtotal)\b', ls, re.IGNORECASE
                ):
                    break

                if not in_ref:
                    continue

                if re.search(
                    r'DESCRIPCION|NO\. PARTE|Subtotal|IVA|Total|No Efectivo',
                    ls, re.IGNORECASE
                ):
                    continue

                m = patron.match(ls)
                if m and re.match(r'^\d+$', m.group(2)):
                    partidas.append({
                        "OT": ot,
                        "DESCRIPCION": m.group(1).strip(),
                        "NO. PARTE": m.group(2),
                        "MONTO": float(m.group(3).replace(",", "")),
                    })

    return partidas

def build_excel(all_rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Refacciones"

    AZUL_OSC   = "1A3A5C"
    AZUL_MED   = "2E6DA4"
    ROJO       = "C8392B"
    GRIS_CLARO = "F0F4F8"
    BLANCO     = "FFFFFF"

    thin = Side(style="thin", color="CCCCCC")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "REFACCIONES – VALUACIONES QUÁLITAS"
    c.font = Font(name="Arial", bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=AZUL_OSC)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, hdr in enumerate(["OT", "DESCRIPCIÓN", "NO. PARTE", "MONTO"], 1):
        c = ws.cell(2, col, hdr)
        c.font = Font(name="Arial", bold=True, size=10, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=AZUL_MED)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[2].height = 20

    sorted_rows = sorted(all_rows, key=lambda r: r["OT"])
    current_row = 3
    data_start = 3

    for ot_key, group in groupby(sorted_rows, key=lambda r: r["OT"]):
        group_list = list(group)
        first = current_row

        for i, row in enumerate(group_list):
            bg = BLANCO if i % 2 == 0 else GRIS_CLARO
            fill = PatternFill("solid", fgColor=bg)
            fnt = Font(name="Arial", size=10)

            for col in range(1, 5):
                c = ws.cell(current_row, col)
                c.fill = fill
                c.font = fnt
                c.border = borde

            ws.cell(current_row, 1).value = row["OT"]
            ws.cell(current_row, 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(current_row, 2).value = row["DESCRIPCION"]
            ws.cell(current_row, 2).alignment = Alignment(vertical="center")
            ws.cell(current_row, 3).value = row["NO. PARTE"]
            ws.cell(current_row, 3).alignment = Alignment(horizontal="center", vertical="center")
            c4 = ws.cell(current_row, 4)
            c4.value = row["MONTO"]
            c4.alignment = Alignment(horizontal="right", vertical="center")
            c4.number_format = '"$"#,##0.00'
            current_row += 1

        last = current_row - 1

        ws.merge_cells(f"A{current_row}:C{current_row}")
        sl = ws.cell(current_row, 1)
        sl.value = f"Subtotal OT {ot_key}"
        sl.font = Font(name="Arial", bold=True, size=10, color=BLANCO)
        sl.fill = PatternFill("solid", fgColor=AZUL_MED)
        sl.alignment = Alignment(horizontal="right", vertical="center")

        sv = ws.cell(current_row, 4)
        sv.value = f"=SUM(D{first}:D{last})"
        sv.font = Font(name="Arial", bold=True, size=10, color=BLANCO)
        sv.fill = PatternFill("solid", fgColor=AZUL_MED)
        sv.alignment = Alignment(horizontal="right", vertical="center")
        sv.number_format = '"$"#,##0.00'
        sv.border = borde
        ws.row_dimensions[current_row].height = 18
        current_row += 1

    ws.merge_cells(f"A{current_row}:C{current_row}")
    tl = ws.cell(current_row, 1)
    tl.value = "TOTAL GENERAL"
    tl.font = Font(name="Arial", bold=True, size=11, color=BLANCO)
    tl.fill = PatternFill("solid", fgColor=ROJO)
    tl.alignment = Alignment(horizontal="right", vertical="center")

    tv = ws.cell(current_row, 4)
    tv.value = f"=SUM(D{data_start}:D{current_row - 1})"
    tv.font = Font(name="Arial", bold=True, size=11, color=BLANCO)
    tv.fill = PatternFill("solid", fgColor=ROJO)
    tv.alignment = Alignment(horizontal="right", vertical="center")
    tv.number_format = '"$"#,##0.00'
    tv.border = borde
    ws.row_dimensions[current_row].height = 22

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

uploaded_files = st.file_uploader(
    "Sube uno o más PDFs de valuación Quálitas",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if uploaded_files:
    st.markdown(f"**{len(uploaded_files)} archivo(s) cargado(s):**")
    for f in uploaded_files:
        st.markdown(f"- `{f.name}`")

    if st.button("⚙️  Extraer y generar Excel"):
        all_rows = []
        errores = []
        progress = st.progress(0)

        for i, f in enumerate(uploaded_files):
            ot = get_ot(f.name)
            try:
                rows = extract_refacciones(f.read(), ot)
                if rows:
                    all_rows.extend(rows)
                else:
                    errores.append(f"⚠️ `{f.name}` — no se encontraron partidas de REFACCIONES.")
            except Exception as e:
                errores.append(f"❌ `{f.name}` — error: {e}")
            progress.progress((i + 1) / len(uploaded_files))

        for msg in errores:
            st.warning(msg)

        if all_rows:
            excel_bytes = build_excel(all_rows)
            total = sum(r["MONTO"] for r in all_rows)

            st.markdown('<div class="result-box">', unsafe_allow_html=True)
            st.markdown(f"✅ **{len(all_rows)} partidas** extraídas de **{len(uploaded_files)} PDF(s)**")
            st.markdown(f"💰 **Total general: ${total:,.2f}**")
            st.markdown('</div>', unsafe_allow_html=True)

            df = pd.DataFrame(all_rows)
            df_show = df.copy()
            df_show["MONTO"] = df_show["MONTO"].map(lambda x: f"${x:,.2f}")
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            out_name = (
                f"{get_ot(uploaded_files[0].name)}_refacciones.xlsx"
                if len(uploaded_files) == 1
                else "refacciones_qualitas.xlsx"
            )

            st.download_button(
                label="📥  Descargar Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.error("No se pudieron extraer partidas de ningún PDF.")
else:
    st.info("👆 Sube los PDFs de valuación para comenzar.")

st.markdown("---")
st.caption("Extractor Refacciones Quálitas")
