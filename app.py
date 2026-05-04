import streamlit as st
from docx import Document
from openpyxl import load_workbook
import unicodedata


# =========================
# NORMALIZAÇÃO
# =========================
def normalizar(texto):
    if not texto:
        return ""
    texto = str(texto).strip().lower()
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    return texto


def limpar_item_excel(texto):
    if not texto:
        return ""
    texto = str(texto).split("&")[0]
    return texto.strip()


# =========================
# DOCX
# =========================
def ler_docx(file):
    doc = Document(file)
    titulos = []

    for tabela in doc.tables:
        for row in tabela.rows:
            texto = row.cells[0].text.strip()
            if texto:
                titulos.append(texto)

    return titulos


# =========================
# EXCEL
# =========================
def analisar_excel(file, substancias):
    wb = load_workbook(file)
    sheets = wb.worksheets[:2]

    substancias_norm = {normalizar(s): s for s in substancias}
    resultado = {s: set() for s in substancias}

    for sheet in sheets:
        for row in sheet.iter_rows(min_row=6, values_only=True):
            if not row:
                continue

            row_clean = list(row)

            if len(row_clean) > 1:
                row_clean[1] = limpar_item_excel(row_clean[1])

            row_norm = {normalizar(cell) for cell in row_clean if cell}

            if len(row) < 7:
                continue

            area = row[6]
            if not area:
                continue

            area = str(area).strip()

            for norm_sub, original_sub in substancias_norm.items():
                if norm_sub in row_norm:
                    resultado[original_sub].add(area)

    output = [
        (s, sorted(list(a)))
        for s, a in resultado.items()
        if a
    ]

    return output


# =========================
# STREAMLIT UI
# =========================
st.title("🔬 Analisador de Substâncias")

st.write("Faz upload do ficheiro DOCX e XLSX")

docx_file = st.file_uploader("DOCX (substâncias)", type=["docx"])
xlsx_file = st.file_uploader("XLSX (consumos)", type=["xlsx"])

if docx_file and xlsx_file:
    if st.button("Analisar"):

        with st.spinner("A analisar ficheiros..."):

            substancias = ler_docx(docx_file)
            resultado = analisar_excel(xlsx_file, substancias)

        st.success("Análise concluída!")

        st.write("## Resultado")

        for substancia, areas in resultado:
            st.write(f"**{substancia}** → {', '.join(areas)}")

if __name__ == "__main__":
    import streamlit.web.cli as stcli
    import sys
    sys.argv = ["streamlit", "run", "app.py", "--server.port=10000"]
    sys.exit(stcli.main())
