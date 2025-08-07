import streamlit as st
import pandas as pd
import math
from io import BytesIO

ARCHIVO_STOCK = "stock.csv"

# --- Funciones de stock ---
def cargar_stock():
    return pd.read_csv(ARCHIVO_STOCK)

def guardar_stock(df):
    df.to_csv(ARCHIVO_STOCK, index=False)

# --- Fracción de goma espuma por modelo ---
def calcular_fraccion_goma_espuma(modelo):
    fracciones = {
        "Silla Mora": 1/15,
        "Silla Franca": 1/15,
        "Silla Xanaes": 1/15,
        "Silla Gala": 1/12,
        "Silla Eva": 1/12,
        "Silla Luna": 1/20
    }
    return fracciones.get(modelo, 0)

# --- Cálculo de caños por barra de 6 m ---
def calcular_barras_usadas(cortes):
    total_cm = sum(cortes)
    return math.ceil(total_cm / 600)

# --- DESPIECE COMPLETO ACTUALIZADO ---
df_despiece = pd.DataFrame([
    # Silla Franca
    ["Silla Franca", "Caño", '1 1/2"', 42, 2],
    ["Silla Franca", "Caño", '1 1/2"', 85, 2],
    ["Silla Franca", "Caño", '15x25mm', 50, 2],
    ["Silla Franca", "Caño", '1/2"', 33.5, 2],
    ["Silla Franca", "Caño", '1/2"', 37, 1],
    ["Silla Franca", "Caño", '10x20mm', 35, 3],
    ["Silla Franca", "Tela", '55x46', 0, 1],
    ["Silla Franca", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Franca", "Tornillo", '4½ x 40mm', 0, 2],
    # Silla Luna
    ["Silla Luna", "Caño", '1"', 105, 2],
    ["Silla Luna", "Caño", '1"', 66, 2],
    ["Silla Luna", "Caño", '3/4"', 37, 2],
    ["Silla Luna", "Caño", 'Hierro del 6 liso', 34, 4],
    ["Silla Luna", "Tela", '20x46', 0, 1],
    ["Silla Luna", "Tornillo", '4½ x 35mm', 0, 3],
    # Silla Mora
    ["Silla Mora", "Caño", '1 1/2"', 42, 2],
    ["Silla Mora", "Caño", '1 1/2"', 85, 2],
    ["Silla Mora", "Caño", '15x25mm', 50, 2],
    ["Silla Mora", "Caño", '1/2"', 33.5, 2],
    ["Silla Mora", "Caño", '1/2"', 37, 1],
    ["Silla Mora", "Chapa", 'Diagonal 2 perforaciones', 0, 2],
    ["Silla Mora", "Tela", '55x46', 0, 1],
    ["Silla Mora", "Tela", 'Tela respaldo', 0, 1],
    ["Silla Mora", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 40mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 16mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 20mm', 0, 4],
    # Silla Xanaes
    ["Silla Xanaes", "Caño", '1 1/2"', 42, 2],
    ["Silla Xanaes", "Caño", '1 1/2"', 85, 2],
    ["Silla Xanaes", "Caño", '15x25mm', 50, 2],
    ["Silla Xanaes", "Caño", '1/2"', 33.5, 2],
    ["Silla Xanaes", "Caño", '1/2"', 37, 1],
    ["Silla Xanaes", "Chapa", 'Rectangular 3 perforaciones', 0, 2],
    ["Silla Xanaes", "Tela", '55x46', 0, 1],
    ["Silla Xanaes", "Tela", 'Tela respaldo', 0, 1],
    ["Silla Xanaes", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Xanaes", "Tornillo", '4½ x 40mm', 0, 2],
    ["Silla Xanaes", "Tornillo", '4½ x 16mm', 0, 2],
    # Silla Eva
    ["Silla Eva", "Caño", '1"', 43, 2],
    ["Silla Eva", "Caño", '1"', 87, 2],
    ["Silla Eva", "Caño", '15x25mm', 34.5, 3],
    ["Silla Eva", "Caño", '15x25mm', 35.5, 1],
    ["Silla Eva", "Caño", '1/2"', 35.5, 1],
    ["Silla Eva", "Caño", '1/2"', 35.5, 3],
    ["Silla Eva", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Eva", "Tornillo", '4½ x 35mm', 0, 1],
    # Silla Gala
    ["Silla Gala", "Caño", '1"', 71, 2],
    ["Silla Gala", "Caño", '1"', 100, 2],
    ["Silla Gala", "Caño", '15x25mm', 35, 1],
    ["Silla Gala", "Caño", '15x25mm', 31, 1],
    ["Silla Gala", "Caño", '1/2"', 31.5, 1],
    ["Silla Gala", "Caño", '1/2"', 31, 1],
    ["Silla Gala", "Tela", '53x55', 0, 1],
    ["Silla Gala", "Tela", 'Tela respaldo', 0, 1],
    ["Silla Gala", "Tornillo", '4½ x 35mm', 0, 4],
    # Banquetas Umma
    ["Banquetas Umma", "Caño", '1"', 114, 2],
    ["Banquetas Umma", "Caño", '1"', 70, 2],
    ["Banquetas Umma", "Caño", '3/4"', 35.5, 1],
    ["Banquetas Umma", "Caño", '3/4"', 29.5, 1],
    ["Banquetas Umma", "Caño", '15x25mm', 42, 2],
    ["Banquetas Umma", "Caño", '1/2"', 45.5, 2],
    ["Banquetas Umma", "Tela", '50x46', 0, 1],
    ["Banquetas Umma", "Tela", 'Tela respaldo', 0, 1],
    ["Banquetas Umma", "Tornillo", '4½ x 45mm', 0, 3],
    ["Banquetas Umma", "Tornillo", '4½ x 16mm', 0, 2],
    ["Banquetas Umma", "Tornillo", '4½ x 20mm', 0, 4],
    # Mesa
    ["Mesa", "Caño", '2"', 74.5, 4],
    ["Mesa", "Caño", '20x20', 60, 4],
    ["Mesa", "Caño", '20x40', 4, 4],
    ["Mesa", "Caño", '20x20', 1100, 4],
    ["Mesa", "Madera", 'Contratapa', 0, 1],
    # Perchero
    ["Perchero", "Caño", '3/4"', 40, 4],
    ["Perchero", "Caño", '3/4"', 50, 4],
    ["Perchero", "Caño", '3/4"', 50, 1],
    ["Perchero", "Caño", '1 1/4"', 130, 1],
], columns=["Modelo", "Categoría", "Tipo", "Largo (cm)", "Cantidad"])

st.set_page_config(page_title="Sistema de Despiece", layout="wide")
st.title("📐 Sistema de Despiece y Stock – DELARTE")

menu = st.sidebar.radio("Menú", ["📐 Despiece", "📦 Stock"])

if menu == "📐 Despiece":
    modelos = df_despiece["Modelo"].unique()
    modelo_seleccionado = st.selectbox("Seleccioná un modelo", modelos)
    cantidad = st.number_input("Cantidad a producir", min_value=1, step=1)
    cliente = st.text_input("Cliente")
    color_tela = st.text_input("Color de tapizado")
    color_cano = st.text_input("Color de caño")

    df_modelo = df_despiece[df_despiece["Modelo"] == modelo_seleccionado].copy()
    df_modelo["Total"] = df_modelo["Cantidad"] * cantidad

    # --- GOMA ESPUMA (visual: 1 por silla / stock: fraccionado) ---
    fraccion_goma = calcular_fraccion_goma_espuma(modelo_seleccionado)
    if fraccion_goma > 0:
        df_modelo = pd.concat([
            df_modelo,
            pd.DataFrame([{
                "Modelo": modelo_seleccionado,
                "Categoría": "Goma espuma",
                "Tipo": "Plancha",
                "Largo (cm)": 0,
                "Cantidad": 1,
                "Total": cantidad
            }])
        ], ignore_index=True)

    # --- Simulación de stock ---
    df_stock = cargar_stock()
    stock_dict = {}
    for _, fila in df_stock.iterrows():
        clave = (fila["Categoría"], fila["Tipo"])
        stock_dict[clave] = fila["Stock"]

    df_modelo["Stock Actual"] = df_modelo.apply(
        lambda x: stock_dict.get((x["Categoría"], x["Tipo"]), 0), axis=1)

    # --- Agrupación de consumo ---
    consumo_real = {}

    for _, row in df_modelo.iterrows():
        clave = (row["Categoría"], row["Tipo"])
        if row["Categoría"] == "Caño":
            largo_total = row["Total"] * row["Largo (cm)"]
            consumo_real[clave] = consumo_real.get(clave, 0) + largo_total
        elif row["Categoría"] == "Goma espuma":
            consumo_real[clave] = consumo_real.get(clave, 0) + (row["Total"] * fraccion_goma)
        else:
            consumo_real[clave] = consumo_real.get(clave, 0) + row["Total"]

    estados = []
    for _, row in df_modelo.iterrows():
        clave = (row["Categoría"], row["Tipo"])
        stock_actual = stock_dict.get(clave, 0)
        if clave in consumo_real:
            if row["Categoría"] == "Caño":
                caños_necesarios = consumo_real[clave] / 600
                estado = "✅ OK" if stock_actual >= caños_necesarios else "❌ Faltante"
            else:
                estado = "✅ OK" if stock_actual >= consumo_real[clave] else "❌ Faltante"
        else:
            estado = "✅ OK"
        estados.append(estado)

    df_modelo["Estado"] = estados

    # --- Visualización (sin A Consumir) ---
    st.subheader("🧩 Despiece y Simulación de Stock")
    columnas_vista = ["Categoría", "Tipo", "Largo (cm)", "Cantidad", "Total", "Stock Actual", "Estado"]
    df_vista = df_modelo[columnas_vista]
    st.dataframe(df_vista)

    # --- Guardado de stock (solo si OK) ---
    if st.button("💾 Guardar Producción y Actualizar Stock", key="guardar_produccion"):
        for clave, consumo in consumo_real.items():
            if clave in stock_dict:
                if clave[0] == "Caño":
                    caños_usados = consumo / 600
                    df_stock.loc[(df_stock["Categoría"] == clave[0]) &
                                 (df_stock["Tipo"] == clave[1]), "Stock"] -= caños_usados
                else:
                    df_stock.loc[(df_stock["Categoría"] == clave[0]) &
                                 (df_stock["Tipo"] == clave[1]), "Stock"] -= consumo
        guardar_stock(df_stock)
        st.success("✅ Producción guardada y stock actualizado.")

elif menu == "📦 Stock":
    st.subheader("📦 Stock de Materiales")
    df_stock = cargar_stock()
    edited_df = st.data_editor(df_stock, num_rows="dynamic", use_container_width=True)
    if st.button("💾 Guardar Cambios en Stock", key="guardar_stock"):
        guardar_stock(edited_df)
        st.success("✅ Cambios guardados correctamente.")

# --- EXPORTACIÓN A EXCEL ---
def exportar_excel(df, modelo, cantidad, cliente, color_tela, color_cano):
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, Alignment
    from openpyxl.worksheet.page import PageMargins

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Despiece"

    # Configuración de impresión
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.3, bottom=0.3)
    ws.print_options.horizontalCentered = True

    # Encabezado
    encabezado = [
        ["Cliente:", cliente],
        ["Modelo:", modelo],
        ["Cantidad:", cantidad],
        ["Color de tapizado:", color_tela],
        ["Color de caño:", color_cano]
    ]
    for i, fila in enumerate(encabezado, start=1):
        ws.cell(row=i, column=1, value=fila[0])
        ws.cell(row=i, column=2, value=fila[1])
        ws.cell(row=i, column=1).font = Font(bold=True)

    # Tabla de despiece (sin "A Consumir")
    columnas = ["Categoría", "Tipo", "Largo (cm)", "Cantidad", "Total", "Stock Actual", "Estado"]
    datos = df[columnas].copy()

    for r_idx, row in enumerate(dataframe_to_rows(datos, index=False, header=True), start=len(encabezado) + 2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
            ws.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")
            if r_idx == len(encabezado) + 2:
                ws.cell(row=r_idx, column=c_idx).font = Font(bold=True)

    # Ancho personalizado de columnas
    columnas_personalizadas = {
        1: 18, 2: 22, 3: 14, 4: 10, 5: 18,
        6: 16, 7: 14
    }
    for col, width in columnas_personalizadas.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- BOTÓN PARA IMPRIMIR DESPIECE ---
if menu == "📐 Despiece" and not df_modelo.empty:
    excel_bytes = exportar_excel(df_modelo, modelo_seleccionado, cantidad, cliente, color_tela, color_cano)
    st.download_button(
        label="🖨️ Imprimir Despiece",
        data=excel_bytes,
        file_name=f"Despiece_{modelo_seleccionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_excel"
    )
