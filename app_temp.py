import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# --- CONFIGURACIÓN GENERAL ---
st.set_page_config(page_title="Sistema Despiece y Stock", layout="wide")
ARCHIVO_STOCK = "stock.csv"
os.remove(ARCHIVO_STOCK) if os.path.exists(ARCHIVO_STOCK) else None

# --- INICIALIZAR STOCK ---
def inicializar_stock():
    if not os.path.exists(ARCHIVO_STOCK):
        data = {
            "Categoría": [
                "Caño", "Caño", "Caño", "Caño", "Caño", "Caño", "Caño",
                "Chapa", "Chapa",
                "Goma espuma", "Tela", "Tela", "Madera",
                "Tornillo", "Tornillo", "Tornillo", "Tornillo", "Tornillo",
                "Regatones"
            ],
            "Tipo": [
                '1"', '3/4"', '1 1/2"', '15x25mm', '1/2"', 'Hierro del 6 liso', '10x20mm',
                'Diagonal 2 perforaciones', 'Rectangular 3 perforaciones',
                'Plancha', 'Tela asiento', 'Tela respaldo', 'Contratapa',
                '4½ x 35mm', '4½ x 45mm', '4½ x 40mm', '4½ x 16mm', '4½ x 20mm',
                'Regatones'
            ],
            "Unidad": [
                "m", "m", "m", "m", "m", "m", "m",
                "un", "un",
                "pliego", "m", "m", "un",
                "un", "un", "un", "un", "un",
                "un"
            ],
            "Stock": [
                200, 200, 200, 200, 200, 200, 200,
                100, 100,
                20, 100, 100, 100,
                500, 500, 500, 500, 500,
                200
            ]
        }
        pd.DataFrame(data).to_csv(ARCHIVO_STOCK, index=False)

inicializar_stock()

# --- BASE DE DESPIECE ---
despieces = {
    "Silla Franca": [
        ("Caño", '1"', 105, 2), ("Caño", '1"', 66, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 34, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.5),
        ("Tornillo", '4½ x 45mm', None, 2),
        ("Tornillo", '4½ x 40mm', None, 1),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Silla Mora": [
        ("Caño", '1"', 104, 2), ("Caño", '1"', 65, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 35, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.5),
        ("Tornillo", '4½ x 45mm', None, 1),
        ("Tornillo", '4½ x 40mm', None, 1),
        ("Tornillo", '4½ x 16mm', None, 1),
        ("Tornillo", '4½ x 20mm', None, 1),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Silla Xanaes": [
        ("Caño", '1"', 104, 2), ("Caño", '1"', 65, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 35, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.5),
        ("Tornillo", '4½ x 45mm', None, 1),
        ("Tornillo", '4½ x 40mm', None, 1),
        ("Tornillo", '4½ x 16mm', None, 1),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Silla Gala": [
        ("Caño", '1"', 105, 2), ("Caño", '1"', 66, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 34, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.6),
        ("Tornillo", '4½ x 35mm', None, 3),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Silla Eva": [
        ("Caño", '1"', 105, 2), ("Caño", '1"', 66, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 34, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.6),
        ("Tornillo", '4½ x 45mm', None, 2),
        ("Tornillo", '4½ x 35mm', None, 1),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Silla Luna": [
        ("Caño", '1"', 105, 2), ("Caño", '1"', 66, 2), ("Caño", '3/4"', 37, 2),
        ("Caño", 'Hierro del 6 liso', 34, 4),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.46),
        ("Tornillo", '4½ x 35mm', None, 3),
        ("Regatones", 'Regatones', None, 6)
    ],
    "Banquetas Umma": [
        ("Caño", '1 1/2"', 70, 4), ("Caño", '15x25mm', 32, 4),
        ("Caño", '1/2"', 40, 2), ("Caño", '10x20mm', 18, 1),
        ("Goma espuma", 'Plancha', None, 1),
        ("Tela", 'Tela asiento', None, 0.5),
        ("Tornillo", '4½ x 45mm', None, 1),
        ("Tornillo", '4½ x 16mm', None, 1),
        ("Tornillo", '4½ x 20mm', None, 1)
    ],
    "Mesas": [
        ("Caño", '1 1/2"', 75, 4), ("Caño", '15x25mm', 50, 4),
        ("Caño", '1/2"', 38, 2), ("Chapa", 'Rectangular 3 perforaciones', None, 2),
        ("Madera", 'Contratapa', None, 1), ("Tornillo", '4½ x 16mm', None, 6)
    ],
    "Percheros": [
        ("Caño", '3/4"', 40, 4), ("Caño", '3/4"', 50, 4),
        ("Caño", '3/4"', 50, 1), ("Caño", '1"', 130, 1),
        ("Chapa", 'Diagonal 2 perforaciones', None, 1),
        ("Tornillo", '4½ x 35mm', None, 3)
    ]
}

fracciones_goma_espuma = {"Silla Mora": 1/15}

# --- MENÚ LATERAL ---
menu = st.sidebar.radio("Menú principal", ["🧹 Despiece", "📦 Stock"])

# --- DESPIECE ---
if menu == "🧹 Despiece":
    st.header("🧹 Generador de Despiece")
    modelo = st.selectbox("📌 Elegí el modelo", list(despieces.keys()))
    cantidad = st.number_input("🔢 Cantidad a producir", min_value=1, value=1)
    cliente = st.text_input("👤 Cliente (opcional)", "")

    if st.button("⚙️ Generar Despiece"):
        df_stock = pd.read_csv(ARCHIVO_STOCK)
        data, simulado = [], []

        for categoria, tipo, largo, cant_unit in despieces[modelo]:
            total_usado = cant_unit * cantidad
            largo_total = total_usado * largo if largo else ""
            mostrar = cantidad if tipo == "Plancha" and categoria == "Goma espuma" else total_usado
            usado_real = cantidad * fracciones_goma_espuma.get(modelo, 1) if tipo == "Plancha" else total_usado
            stock_actual = df_stock.loc[df_stock["Tipo"] == tipo, "Stock"].values[0] if tipo in df_stock["Tipo"].values else 0
            stock_ok = stock_actual >= usado_real

            data.append({
                "Categoría": categoria, "Tipo": tipo, "Largo (cm)": largo or "",
                "Cantidad x unidad": cant_unit, "Total": mostrar, "Total Largo (cm)": largo_total,
                "Stock OK": "✅" if stock_ok else "❌"
            })
            simulado.append({
                "Tipo": tipo, "Stock actual": stock_actual, "Necesario": usado_real,
                "Stock final simulado": round(stock_actual - usado_real, 2)
            })

        df_resultado = pd.DataFrame(data)
        df_simulado = pd.DataFrame(simulado)
        st.markdown("### 🧮 Resultado del despiece")
        st.dataframe(df_resultado, use_container_width=True)
        st.markdown("### 🔍 Simulación de stock si se realiza esta producción")
        st.dataframe(df_simulado, use_container_width=True)

        def exportar_excel(df, modelo, cantidad, cliente):
            archivo = f"despiece_{modelo}_{cantidad}u.xlsx"
            df.to_excel(archivo, index=False)
            wb = load_workbook(archivo)
            ws = wb.active
            ws.insert_rows(1)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
            cell = ws.cell(row=1, column=1)
            cell.value = f"Pedido: {cantidad} {modelo}" + (f" – Cliente {cliente}" if cliente else "")
            cell.font = Font(size=18, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

            header_fill = PatternFill(start_color="D9D9D9", fill_type="solid")
            row_fill = PatternFill(start_color="F8F8F8", fill_type="solid")
            header_font = Font(bold=True, size=14)
            body_font = Font(size=13)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=2, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                ws.column_dimensions[get_column_letter(col)].width = 20

            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = body_font
                    cell.border = border
                    if cell.row % 2 == 1:
                        cell.fill = row_fill
                ws.row_dimensions[cell.row].height = 30

            ws.row_dimensions[2].height = 50
            ws.freeze_panes = "A3"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
            wb.save(archivo)
            return archivo

        archivo_excel = exportar_excel(df_resultado, modelo, cantidad, cliente)
        with open(archivo_excel, "rb") as f:
            st.download_button("🖨️ Imprimir Despiece", f, archivo_excel)

# --- STOCK ---
elif menu == "📦 Stock":
    st.header("📦 Stock Actual")
    df_stock = pd.read_csv(ARCHIVO_STOCK)
    st.markdown("### ✏️ Editar stock actual")
    df_editado = st.data_editor(df_stock, num_rows="dynamic", use_container_width=True)
    if st.button("📅 Guardar cambios"):
        df_editado.to_csv(ARCHIVO_STOCK, index=False)
        st.success("✅ Cambios guardados correctamente.")
