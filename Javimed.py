import pandas as pd
import streamlit as st
from tkinter import Tk, filedialog
from PIL import Image

def cargar_archivo_excel():
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    archivo = filedialog.askopenfilename(title="Selecciona el archivo de liquidaciones", filetypes=[("Archivos Excel", "*.xlsx")])
    return archivo

def realizar_analisis(df):
    resultados = {}

    # Filtros según las órdenes proporcionadas con encabezados claramente definidos
    resultados['consulta_1'] = df[(df['Ente Comercial'] == 'JAVIER JESUS COLLADO VALDIVIESO-724') & (df['Comision total'] < 87)]
    resultados['consulta_2'] = df[(df['Abrev.Cía'] == 'Mapfre') & (df['Ente Comercial'] == 'VOLCAN, SALITRE Y LAVA SL-284') & (df['Comision total'] < 85)]
    resultados['consulta_3'] = df[(df['Abrev.Cía'] == 'Mapfre') & (df['Ente Comercial'] == 'CRISTINA PEÑAS BLANCO') & (df['Comision total'] < 83)]
    resultados['consulta_4'] = df[(df['Comision total'] > 86) & (~df['Ente Comercial'].isin(['FINISTERRE21 CORREDURIA DE SEGUROS, S.L.-410', 'JAVIER JESUS COLLADO VALDIVIESO-724', 'HERRERO BROKERS CORREDURIA DE SEGUROS SL-760', 'SOLEDAD MUÑOZ MARCOS-755']))]
    resultados['consulta_5'] = df[df['G+'] > 50]
    resultados['consulta_6'] = df[(df['G+'] < 50) & (df['Clase'] == 'PRODUCCION')]
    resultados['consulta_7'] = df[(df['Comision cedida'] > 6) & (df['Abrev.Cía'] == 'ADMIRAL')]
    resultados['consulta_8'] = df[(df['Comision cedida'] > 10) & (df['Abrev.Cía'] == 'CASER') & (df['Ramo Niv.2'].isin(['MIXTO REEM-CON.', 'GASTOS CONCERT.']))]
    resultados['consulta_9'] = df[(df['Comision cedida'] > 26) & (df['Abrev.Cía'] == 'OCASO') & (df['Ramo Niv.2'] == 'HOGAR')]
    resultados['consulta_10'] = df[(df['Comision cedida'] > 20) & (df['Abrev.Cía'] == 'OCASO') & (df['Ramo Niv.2'] == 'COMUNIDADES')]
    resultados['consulta_11'] = df[(df['Comision total'] >= 40) & (df['Comision total'] <= 43) & (df['Abrev.Cía'] == 'CASER') & (df['Ramo Niv.2'].isin(['MIXTO REEM-CON.', 'GASTOS CONCERT.']))]
    resultados['consulta_12'] = df[(df['Comision total'] >= 57) & (df['Comision total'] <= 60) & (df['Abrev.Cía'] == 'CASER') & (df['Ramo Niv.2'].isin(['MIXTO REEM-CON.', 'GASTOS CONCERT.']))]
    resultados['consulta_13'] = df[(df['Comision cedida'] >= 6.1) & (df['Comision cedida'] <= 6.9) & (df['Abrev.Cía'] == 'CASER') & (df['Ramo Niv.2'].isin(['MIXTO REEM-CON.', 'GASTOS CONCERT.']))]
    resultados['consulta_14'] = df[(df['Comision cedida'] > 12) & (df['Abrev.Cía'] == 'Axa') & (df['Ramo Niv.2'].isin(['TURISMOS/FURGO.', 'FURGONETA >700K', 'MOTOCICLETA', 'CAMION RIGIDO']))]
    resultados['consulta_15'] = df[(df['Comision cedida'] > 10.2) & (df['Abrev.Cía'] == 'Generali') & (df['Ramo Niv.2'].isin(['TURISMOS/FURGO.', 'FURGONETA >700K', 'MOTOCICLETA', 'CAMION RIGIDO']))]
    resultados['consulta_16'] = df[(df['Comision cedida'] > 21.25) & (df['Abrev.Cía'] == 'Generali') & (df['Ramo Niv.2'] == 'HOGAR')]
    resultados['consulta_17'] = df[(df['Comision cedida'] >= 3.1) & (df['Comision cedida'] <= 3.95) & (df['Abrev.Cía'] == 'Generali') & (df['Ramo Niv.2'].isin(['TURISMOS/FURGO.', 'FURGONETA >700K', 'MOTOCICLETA', 'CAMION RIGIDO']))]
    resultados['consulta_18'] = df[(df['Comision cedida'] >= 7.1) & (df['Comision cedida'] <= 7.95) & (df['Abrev.Cía'] == 'Generali') & (df['Ramo Niv.2'].isin(['TURISMOS/FURGO.', 'FURGONETA >700K', 'MOTOCICLETA', 'CAMION RIGIDO']))]
    resultados['consulta_19'] = df[(df['Comision total'] > 91) & (df['Ente Comercial'].isin(['Finisterre21', 'Herrero Broker', 'Soledad Muñoz', 'Garces Inversiones']))]
    resultados['consulta_20'] = df[(df['Comision total'] < 85) & (df['Ente Comercial'].isin(['Finisterre21', 'Herrero Broker', 'Soledad Muñoz', 'Garces Inversiones']))]
    resultados['consulta_21'] = df[(df['Comision total'] > 85) & (df['Clase'] == 'PRODUCCION') & (df['Ente Comercial'].isin(['Finisterre21', 'Soledad Muñoz', 'Garces Inversiones']))]
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
