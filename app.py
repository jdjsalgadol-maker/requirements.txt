import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds
import io

# Configuración de la página
st.set_page_config(page_title="Conciliador Bancario GNB", layout="wide")

st.title("🏦 Concilia GNB @JuanS")
st.markdown("""
Diseñado para conciliar de Muchos a un bloque.
""")

# Botón para subir archivo
archivo_subido = st.file_uploader("Sube tu archivo de Excel (.xlsx)", type=['xlsx'])

if archivo_subido is not None:
    try:
        df = pd.read_excel(archivo_subido)
        st.info("Archivo cargado con éxito. Procesando datos...")

        # 1. Separar datos
        df_40 = df[df['Clave contabiliz.'] == 40].reset_index(drop=False)
        values_40 = df_40['Importe en moneda local'].values
        
        # El objetivo es la suma de todas las claves 50
        target_sum = df[df['Clave contabiliz.'] == 50]['Importe en moneda local'].sum()
        
        st.write(f"**Objetivo de suma (Clave 50):** ${target_sum:,.2f}")
        st.write(f"**Registros candidatos (Clave 40):** {len(values_40)}")

        if st.button("🚀 Ejecutar Conciliación"):
            with st.spinner('Buscando la combinación exacta...'):
                n = len(values_40)
                
                # Configurar modelo matemático (MILP)
                c = np.zeros(n)
                A = np.atleast_2d(values_40)
                lb = np.array([target_sum])
                ub = np.array([target_sum])
                
                constraints = LinearConstraint(A, lb, ub)
                bounds = Bounds(0, 1)
                integrality = np.ones(n)
                
                res = milp(c=c, constraints=constraints, bounds=bounds, integrality=integrality)
                
                if res.success:
                    # Marcar resultados
                    df['Conciliacion_Exacta'] = "No Conciliado"
                    selected_indices = np.where(np.round(res.x) == 1)[0]
                    selected_original_indices = df_40.iloc[selected_indices]['index'].values
                    
                    df.loc[selected_original_indices, 'Conciliacion_Exacta'] = "Conciliado (Grupo 40)"
                    df.loc[df['Clave contabiliz.'] == 50, 'Conciliacion_Exacta'] = "Conciliado (Grupo 50)"
                    
                    st.success(f"¡Logrado! Se encontraron {len(selected_indices)} registros que cuadran perfectamente.")
                    
                    # Preparar archivo para descargar
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
        st.error(f"Error técnico: Asegúrate de que las columnas se llamen 'Clave contabiliz.' e 'Importe en moneda local'. Error: {e}")
