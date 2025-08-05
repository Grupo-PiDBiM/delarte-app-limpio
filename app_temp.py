import streamlit as st
import pandas as pd
from io import BytesIO
import math

# --- ARCHIVO LOCAL ---
ARCHIVO_STOCK = "stock.csv"

def cargar_stock():
    return pd.read_csv(ARCHIVO_STOCK)

def guardar_stock(df):
    df.to_csv(ARCHIVO_STOCK, index=False)

def calcular_goma_espuma(modelo):
    fracciones = {
        "Silla Mora": 1/15, "Silla Franca": 1/15, "Silla Xanaes": 1/15,
        "Silla Gala": 1/12, "Silla Eva": 1/12, "Silla Luna": 1/20
    }
    return fracciones.get(modelo, 1)

# --- DESPIECE COMPLETO ---
df_despiece = pd.DataFrame([
    # Silla Franca
    ["Silla Franca", "Caño", '1"', 40, 2],
    ["Silla Franca", "Caño", '3/4"', 35, 2],
    ["Silla Franca", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Franca", "Tornillo", '4½ x 40mm', 0, 2],
    ["Silla Franca", "Tela", 'Tela asiento', 0, 1],
    ["Silla Franca", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Franca", "Regatones", 'Regatones', 0, 6],
    # Silla Luna
    ["Silla Luna", "Caño", '1"', 45, 2],
    ["Silla Luna", "Caño", '3/4"', 30, 2],
    ["Silla Luna", "Tornillo", '4½ x 35mm', 0, 3],
    ["Silla Luna", "Tela", 'Tela asiento', 0, 1],
    ["Silla Luna", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Luna", "Regatones", 'Regatones', 0, 6],
    # Silla Mora
    ["Silla Mora", "Caño", '1"', 42, 2],
    ["Silla Mora", "Caño", '3/4"', 32, 2],
    ["Silla Mora", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 40mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 16mm', 0, 2],
    ["Silla Mora", "Tornillo", '4½ x 20mm', 0, 4],
    ["Silla Mora", "Tela", 'Tela respaldo', 0, 1],
    ["Silla Mora", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Mora", "Regatones", 'Regatones', 0, 6],
    # Silla Xanaes
    ["Silla Xanaes", "Caño", '1"', 30, 2],
    ["Silla Xanaes", "Caño", '3/4"', 35, 2],
    ["Silla Xanaes", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Xanaes", "Tornillo", '4½ x 40mm', 0, 2],
    ["Silla Xanaes", "Tornillo", '4½ x 16mm', 0, 6],
    ["Silla Xanaes", "Tela", 'Tela respaldo', 0, 1],
    ["Silla Xanaes", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Xanaes", "Regatones", 'Regatones', 0, 6],
    # Silla Gala
    ["Silla Gala", "Caño", '1"', 44, 2],
    ["Silla Gala", "Caño", '3/4"', 36, 2],
    ["Silla Gala", "Tornillo", '4½ x 35mm', 0, 4],
    ["Silla Gala", "Tela", 'Tela asiento', 0, 1],
    ["Silla Gala", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Gala", "Regatones", 'Regatones', 0, 6],
    # Silla Eva
    ["Silla Eva", "Caño", '1"', 44, 2],
    ["Silla Eva", "Caño", '3/4"', 36, 2],
    ["Silla Eva", "Tornillo", '4½ x 45mm', 0, 2],
    ["Silla Eva", "Tornillo", '4½ x 35mm', 0, 1],
    ["Silla Eva", "Tela", 'Tela asiento', 0, 1],
    ["Silla Eva", "Goma espuma", 'Plancha', 0, 1],
    ["Silla Eva", "Regatones", 'Regatones', 0, 6],
    # Banquetas Umma
    ["Banquetas Umma", "Caño", '1 1/2"', 70, 4],
    ["Banquetas Umma", "Caño", '15x25mm', 32, 4],
    ["Banquetas Umma", "Caño", '1/2"', 40, 2],
    ["Banquetas Umma", "Caño", '10x20mm', 18, 1],
    ["Banquetas Umma", "Tornillo", '4½ x 45mm', 0, 3],
    ["Banquetas Umma", "Tornillo", '4½ x 16mm', 0, 2],
    ["Banquetas Umma", "Tornillo", '4½ x 20mm', 0, 4],
    ["Banquetas Umma", "Tela", 'Tela asiento', 0, 1],
    ["Banquetas Umma", "Goma espuma", 'Plancha', 0, 1],
    ["Banquetas Umma", "Regatones", 'Regatones', 0, 6],
    # Mesa
    ["Mesa", "Caño", '1 1/2"', 70, 4],
    ["Mesa", "Madera", 'Contratapa', 0, 1],
    ["Mesa", "Tornillo", '4½ x 45mm', 0, 8],
    # Perchero
    ["Perchero", "Caño", '1"', 150, 1],
    ["Perchero", "Caño", '3/4"', 45, 2],
    ["Perchero", "Chapa", 'Rectangular 3 perforaciones', 0, 1],
    ["Perchero", "Tornillo", '4½ x 16mm', 0, 4],
], columns=["Modelo", "Categoría", "Tipo", "Largo (cm)", "Cantidad"])

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Delarte App", layout="wide")
st.sidebar.title("🪑 Menú")
pagina = st.sidebar.radio("Ir a:", ["📐 Despiece", "📦 Stock"])

# --- SECCIÓN DESPIECE ---
if pagina == "📐 Despiece":
    st.title("📐 Generador de Despiece")

    modelo = st.selectbox("Seleccionar modelo", df_despiece["Modelo"].unique())
    cantidad = st.number_input("Cantidad a producir", min_value=1, value=1)
    cliente = st.text_input("Cliente", "")
    color_caño = st.selectbox("Color de caño", ["Negro", "Blanco", "Cromado", "Gris", "Otro"])
    color_tapizado = st.selectbox("Color de tapizado", ["Negro", "Azul", "Rojo", "Beige", "Otro"])
    titulo = st.text_input("Título para el Excel", f"Pedido – {modelo}")

    # Calcular despiece y totales
    df_modelo = df_despiece[df_despiece["Modelo"] == modelo].copy()
    df_modelo["Total"] = df_modelo["Cantidad"] * cantidad
    if "Goma espuma" in df_modelo["Categoría"].values:
        df_modelo.loc[df_modelo["Categoría"] == "Goma espuma", "Total"] = cantidad

    # Cargar stock y cruzar
    df_stock = cargar_stock()
    df_merge = pd.merge(df_modelo, df_stock, on=["Categoría", "Tipo"], how="left")
    df_merge["Estado"] = df_merge.apply(
        lambda row: "✅ OK" if row["Stock"] >= row["Total"] else "❌ Faltante", axis=1
    )

    st.subheader("📋 Despiece + Stock")
    st.dataframe(
        df_merge[["Categoría", "Tipo", "Largo (cm)", "Cantidad", "Total", "Stock", "Estado"]],
        use_container_width=True
    )

    # --- EXPORTAR A EXCEL ---
    def exportar_excel(df, modelo, cantidad, cliente, color_caño, color_tapizado):
        output = BytesIO()
        df_export = df[["Categoría", "Tipo", "Cantidad", "Total"]].copy()
        df_export.columns = ["Categoría", "Tipo", "Cantidad por unidad", "Total"]

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, startrow=6, sheet_name="Despiece")
            sheet = writer.sheets["Despiece"]

            sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
            sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
            sheet.page_setup.fitToPage = True
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.sheet_view.zoomScale = 130

            sheet["A1"] = f"Modelo: {modelo}"
            sheet["A2"] = f"Cantidad: {cantidad}"
            sheet["A3"] = f"Color de caño: {color_caño}"
            sheet["A4"] = f"Color de tapizado: {color_tapizado}"
            sheet["A5"] = f"Cliente: {cliente}"
            sheet["A6"] = ""

            columnas_personalizadas = {
                1: 24, 2: 28, 3: 24, 4: 22
            }
            for col, width in columnas_personalizadas.items():
                sheet.column_dimensions[chr(64 + col)].width = width

        return output.getvalue()

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="📄 Imprimir",
            data=exportar_excel(df_modelo, modelo, cantidad, cliente, color_caño, color_tapizado),
            file_name=f"{titulo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        if st.button("💾 Guardar Producción y Actualizar Stock"):
            df_stock = cargar_stock()

            # Descuento de caños por largo real
            df_caños = df_modelo[df_modelo["Categoría"] == "Caño"]
            tipos_caño = df_caños["Tipo"].unique()

            for tipo in tipos_caño:
                total_cm = 0
                for fila in df_caños[df_caños["Tipo"] == tipo].itertuples():
                    largo = fila._4
                    cantidad_por_unidad = fila.Cantidad
                    total_cortes = cantidad_por_unidad * cantidad
                    total_cm += largo * total_cortes
                fraccion_caños = total_cm / 600  # Largo de barra
                idx = df_stock[(df_stock["Categoría"] == "Caño") & (df_stock["Tipo"] == tipo)].index
                if not idx.empty:
                    df_stock.loc[idx, "Stock"] -= fraccion_caños

            # Descuento de otros materiales
            for fila in df_modelo.itertuples():
                cat, tipo, total = fila.Categoría, fila.Tipo, fila.Total
                if cat == "Caño":
                    continue
                elif cat == "Goma espuma":
                    total = cantidad * calcular_goma_espuma(modelo)
                idx = df_stock[(df_stock["Categoría"] == cat) & (df_stock["Tipo"] == tipo)].index
                if not idx.empty:
                    df_stock.loc[idx, "Stock"] -= total

            guardar_stock(df_stock)
            st.success("✅ Producción guardada y stock actualizado correctamente.")

# --- SECCIÓN STOCK ---
elif pagina == "📦 Stock":
    st.title("📦 Stock de Materiales (Editable)")
    df_stock = cargar_stock()

    st.subheader("📋 Tabla Editable")
    edited_df = st.data_editor(
        df_stock,
        use_container_width=True,
        num_rows="fixed",
        key="stock_editor",
        column_config={
            "Categoría": st.column_config.TextColumn(disabled=True),
            "Tipo": st.column_config.TextColumn(disabled=True),
            "Unidad": st.column_config.TextColumn(disabled=True),
            "Stock": st.column_config.NumberColumn("Stock", step=1),
        }
    )

    if st.button("💾 Guardar Cambios en Stock"):
        guardar_stock(edited_df)
        st.success("✅ Cambios guardados correctamente en stock.csv")
