import pandas as pd
import streamlit as st
from PIL import Image

def realizar_analisis(df):
    resultados = {}

    # Filtros según las órdenes proporcionadas con encabezados claramente definidos
    resultados['consulta_1'] = df[(df['Ente Comercial'] == 'JAVIER JESUS COLLADO VALDIVIESO-724') & (df['Comision total'] < 87)]
    resultados['consulta_2'] = df[(df['Abrev.Cía'] == 'Mapfre') & (df['Ente Comercial'] == 'VOLCAN, SALITRE Y LAVA SL-284') & (df['Comision total'] < 85)]
    # ... (el resto de tus consultas)
    resultados['consulta_22'] = df[(df['Comision total'] < 89) & (df['Clase'] == 'CARTERA') & (df['Ente Comercial'].isin(['Finisterre21', 'Soledad Muñoz', 'Garces Inversiones']))]

    return resultados

def guardar_resultados(resultados):
    with pd.ExcelWriter('resultado_analisis_streamlit.xlsx', engine='openpyxl') as writer:
        for nombre_consulta, resultado in resultados.items():
            resultado.to_excel(writer, sheet_name=nombre_consulta, index=False)

# Streamlit Application
def main():
    # Cargar el logo
    logo = Image.open("logo_empresa.png")
    st.image(logo, width=200)

    # Título y descripción
    st.title("JAVIMED: La Fórmula para Liquidaciones Sin Complicaciones")
    st.subheader("¡Facilita el proceso de análisis de liquidaciones de forma rápida y sencilla!")
    st.write("Sube un archivo Excel con las liquidaciones para analizarlo.")

    # Subir el archivo
    archivo = st.file_uploader("Subir archivo de Excel", type=["xlsx"])

    if archivo is not None:
        df = pd.read_excel(archivo, header=0)
        st.write("Datos cargados:")
        st.dataframe(df)

        if st.button("Realizar Análisis"):
            resultados = realizar_analisis(df)
            guardar_resultados(resultados)
            st.success("Análisis completado y archivo de resultados guardado.")

            # Ofrecer el archivo para descarga
            with open("resultado_analisis_streamlit.xlsx", "rb") as file:
                st.download_button(
                    label="Descargar archivo de resultados",
                    data=file,
                    file_name="resultado_analisis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
