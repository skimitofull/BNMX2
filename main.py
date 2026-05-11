import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import numpy as np
import io

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

MM_TO_PT = 2.83465

# Hoja carta vertical
PAGE_W_MM = 215.90
PAGE_H_MM = 279.40

PAGE_W_PT = PAGE_W_MM * MM_TO_PT
PAGE_H_PT = PAGE_H_MM * MM_TO_PT

# Fuente
FONT_NAME = "Helvetica"
FONT_SIZE = 9

# Posiciones horizontales medidas desde el PDF nuevo
X_DIA_MM = 8.78
X_MES_MM = 14.71
X_TEXTO_MM = 28.79

# Cantidades alineadas por su borde derecho
X_MONTO1_RIGHT_MM = 139.78
X_MONTO2_RIGHT_MM = 169.31
X_MONTO3_RIGHT_MM = 199.97

X_DIA_PT = X_DIA_MM * MM_TO_PT
X_MES_PT = X_MES_MM * MM_TO_PT
X_TEXTO_PT = X_TEXTO_MM * MM_TO_PT

X_MONTO1_RIGHT_PT = X_MONTO1_RIGHT_MM * MM_TO_PT
X_MONTO2_RIGHT_PT = X_MONTO2_RIGHT_MM * MM_TO_PT
X_MONTO3_RIGHT_PT = X_MONTO3_RIGHT_MM * MM_TO_PT

# Anchos de columnas para alinear importes
W_DIA_PT = 12
W_MES_PT = 28
W_TEXTO_PT = 92

W_MONTO1_PT = 55
W_MONTO2_PT = 55
W_MONTO3_PT = 55

X_MONTO1_PT = X_MONTO1_RIGHT_PT - W_MONTO1_PT
X_MONTO2_PT = X_MONTO2_RIGHT_PT - W_MONTO2_PT
X_MONTO3_PT = X_MONTO3_RIGHT_PT - W_MONTO3_PT

# Posiciones verticales
Y_START_MM = 50.83
Y_END_MM = 260.00

Y_START_PT = Y_START_MM * MM_TO_PT
Y_END_PT = Y_END_MM * MM_TO_PT

# Altura de línea
LINE_H_MM = 4.13
LINE_H_PT = LINE_H_MM * MM_TO_PT


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def clean_cell(val):
    if val is None:
        return ""
    if isinstance(val, float) and np.isnan(val):
        return ""
    sval = str(val).strip()
    if sval.lower() in ["nan", "none", "null"]:
        return ""
    return sval


def monto_cell(val):
    if val is None:
        return ""
    if isinstance(val, float) and np.isnan(val):
        return ""
    if isinstance(val, str) and val.strip().lower() in ["nan", "none", "null", ""]:
        return ""

    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace("$", "").strip()

        fval = float(val)

        if np.isnan(fval):
            return ""

        if fval == 0:
            return ""

        return f"{fval:,.2f}"
    except Exception:
        return ""


def parse_excel(df):
    """
    Espera archivo con 6 columnas:
    DIA, MES, XXXX, MONTO 1, MONTO 2, MONTO 3
    """

    df = df.iloc[:, :6].copy()

    df.columns = [
        "DIA",
        "MES",
        "XXXX",
        "MONTO 1",
        "MONTO 2",
        "MONTO 3"
    ]

    return df.fillna("")


def split_text(pdf, text, max_width):
    pdf.set_font(FONT_NAME, "", FONT_SIZE)

    text = clean_cell(text)

    if text == "":
        return [""]

    words = text.split(" ")
    lines = []
    current_line = ""

    for word in words:
        test_line = current_line + " " + word if current_line else word

        if pdf.get_string_width(test_line) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines if lines else [""]


# =========================================================
# CLASE PDF
# =========================================================

class EstadoCuentaPDF(FPDF):

    def __init__(self):
        super().__init__(unit="pt", format=(PAGE_W_PT, PAGE_H_PT))
        self.set_auto_page_break(False)
        self.alias_nb_pages()
        self.current_y = Y_START_PT

    def header(self):
        self.set_font(FONT_NAME, "", FONT_SIZE)

        # En este diseño no se dibuja encabezado nuevo.
        # Solo se inicia el área de movimientos.
        self.current_y = Y_START_PT

    def footer(self):
        pass

    def check_page_break(self, needed_height):
        if self.current_y + needed_height > Y_END_PT:
            self.add_page()
            self.current_y = Y_START_PT

    def add_movement_row(self, dia, mes, texto, monto1, monto2, monto3):
        dia_str = clean_cell(dia)
        mes_str = clean_cell(mes)
        texto_str = clean_cell(texto)

        monto1_str = monto_cell(monto1)
        monto2_str = monto_cell(monto2)
        monto3_str = monto_cell(monto3)

        texto_lines = split_text(self, texto_str, W_TEXTO_PT)
        row_height = len(texto_lines) * LINE_H_PT

        self.check_page_break(row_height)

        y = self.current_y

        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.set_text_color(0, 0, 0)

        # Día
        self.set_xy(X_DIA_PT, y)
        self.cell(W_DIA_PT, LINE_H_PT, dia_str, border=0, align="L")

        # Mes
        self.set_xy(X_MES_PT, y)
        self.cell(W_MES_PT, LINE_H_PT, mes_str, border=0, align="L")

        # Texto / XXXX
        for i, line in enumerate(texto_lines):
            self.set_xy(X_TEXTO_PT, y + (i * LINE_H_PT))
            self.cell(W_TEXTO_PT, LINE_H_PT, line, border=0, align="L")

        # Montos: se imprimen en la primera línea del movimiento
        self.set_xy(X_MONTO1_PT, y)
        self.cell(W_MONTO1_PT, LINE_H_PT, monto1_str, border=0, align="R")

        self.set_xy(X_MONTO2_PT, y)
        self.cell(W_MONTO2_PT, LINE_H_PT, monto2_str, border=0, align="R")

        self.set_xy(X_MONTO3_PT, y)
        self.cell(W_MONTO3_PT, LINE_H_PT, monto3_str, border=0, align="R")

        self.current_y += row_height


# =========================================================
# INTERFAZ STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Generador PDF Estado de Cuenta",
    layout="wide",
    page_icon="📄"
)

st.title("📄 Generador de Estado de Cuenta - Nuevo Diseño")
st.markdown("Carga un Excel con columnas: **DIA, MES, XXXX, MONTO 1, MONTO 2, MONTO 3**.")

excel_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xls"]
)

if excel_file:
    try:
        df_raw = pd.read_excel(excel_file)
        df = parse_excel(df_raw)

        st.success(f"Archivo cargado correctamente: {len(df)} filas.")
        st.dataframe(df.head(20), use_container_width=True)

        if st.button("Generar PDF", type="primary", use_container_width=True):
            try:
                pdf = EstadoCuentaPDF()
                pdf.add_page()

                for _, row in df.iterrows():
                    pdf.add_movement_row(
                        row["DIA"],
                        row["MES"],
                        row["XXXX"],
                        row["MONTO 1"],
                        row["MONTO 2"],
                        row["MONTO 3"]
                    )

                pdf_bytes = pdf.output(dest="S").encode("latin1")

                st.success("PDF generado correctamente.")

                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"Estado_Cuenta_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error al generar el PDF: {e}")

    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")

else:
    st.info("Sube un Excel para comenzar.")

    st.markdown("""
    ### Formato esperado

    | DIA | MES | XXXX | MONTO 1 | MONTO 2 | MONTO 3 |
    |---|---|---|---:|---:|---:|
    | 13 | OCT | XXXXXXXXXXXXXXXXXX | 1,118.00 |  | 6,810.23 |
    | 13 | OCT | XXXXXXXXXXXXXXXXXX |  | 3,300.00 | 10,110.23 |
    """)
