import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds
import io

# 1. CONFIGURACIÓN E INTERFAZ (OCULTA MENÚS)
st.set_page_config(page_title="Conciliador Bancario GNB", layout="wide")

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.title("🏦 Concilia GNB @JuanS 🤖")
st.markdown("Diseñado para conciliar de Muchos a un bloque.")

# 2. CARGA DE ARCHIVO
archivo_subido = st.file_uploader("Sube tu archivo de Excel (.xlsx)", type=['xlsx'])

if archivo_subido is not None:
    try:
        df = pd.read_excel(archivo_subido)
        st.info("Archivo cargado con éxito. Procesando datos...")

        # Separar datos: Buscamos registros con Clave 40 para sumar lo que hay en Clave 50
        df_40 = df[df['Clave contabiliz.'] == 40].reset_index(drop=False)
        values_40 = df_40['Importe en moneda local'].values
        
        # El objetivo es la suma de todas las claves 50
        target_sum = df[df['Clave contabiliz.'] == 50]['Importe en moneda local'].sum()
        target_sum = round(float(target_sum), 2) # Evita errores de decimales infinitos
        
        st.write(f"**Objetivo de suma (Clave 50):** ${target_sum:,.2f}")
        st.write(f"**Registros candidatos (Clave 40):** {len(values_40)}")

        # 3. LÓGICA DE CONCILIACIÓN (MILP)
        if st.button("🚀 Ejecutar Conciliación"):
            with st.spinner('Buscando la combinación exacta...'):
                n = len(values_40)
                
                # Configurar modelo matemático
                c = np.zeros(n)
                A = np.atleast_2d(values_40)
                
                # Restricción: La suma debe ser exactamente igual al target_sum
                constraints = LinearConstraint(A, target_sum, target_sum)
                bounds = Bounds(0, 1)
                integrality = np.ones(n) # Variables binarias
                
                res = milp(c=c, constraints=constraints, bounds=bounds, integrality=integrality)
                
                if res.success:
                    # Marcar resultados en el DataFrame original
                    df['Conciliacion_Exacta'] = "No Conciliado"
                    selected_indices = np.where(np.round(res.x) == 1)[0]
                    selected_original_indices = df_40.iloc[selected_indices]['index'].values
                    
                    df.loc[selected_original_indices, 'Conciliacion_Exacta'] = "Conciliado (Grupo 40)"
                    df.loc[df['Clave contabiliz.'] == 50, 'Conciliacion_Exacta'] = "Conciliado (Grupo 50)"
                    
                    st.success(f"¡Logrado! Se encontraron {len(selected_indices)} registros que cuadran perfectamente.")
                    st.balloons()
                    
                    # 4. DESCARGA DE RESULTADOS
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 Descargar Excel Conciliado",
                        data=output.getvalue(),
                        file_name="Conciliacion_GNB_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se encontró una combinación exacta que sume esa cantidad.")
                    
    except Exception as e:
        st.error(f"Error técnico: Revisa que las columnas se llamen 'Clave contabiliz.' e 'Importe en moneda local'. Detalle: {e}")
