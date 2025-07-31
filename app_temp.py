import streamlit as st
import pandas as pd

st.set_page_config(page_title="Sistema DELARTE", layout="wide")

st.title("🧠 Sistema de Despiece y Stock – DELARTE")

st.markdown("""
Bienvenido al sistema web de control de **despiece** y **stock** de productos.

👉 Acá vas a poder:
- Ver los insumos necesarios por modelo
- Ver el stock disponible
- Simular producción
""")

# Muestra ejemplo de modelo
modelos = ["Silla Luna", "Silla Franca", "Silla Mora"]
modelo = st.selectbox("Seleccioná un modelo", modelos)

cantidad = st.number_input("¿Cuántas unidades vas a producir?", min_value=0, max_value=100, value=0, step=1)

if st.button("🧮 Calcular despiece"):
    if cantidad == 0:
        st.warning("Ingresá una cantidad mayor a 0")
    else:
        st.success(f"Calculando despiece para {cantidad} unidades de {modelo}...")
        # Simulación de resultado (más adelante ponemos la lógica real)
        datos = {
            "Material": ["Caño 1''", "Tela", "Goma espuma"],
            "Cantidad por unidad": [2.0, 0.46, 1],
            "Total necesario": [2.0 * cantidad, 0.46 * cantidad, cantidad]
        }
        df = pd.DataFrame(datos)
        st.dataframe(df)
