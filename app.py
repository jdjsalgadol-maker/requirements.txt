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
st.markdown("Diseñado para conciliar de Muchos a un bloque filtrando por ventana de tiempo.")

# 2. CARGA DE ARCHIVO
archivo_subido = st.file_uploader("Sube tu archivo de Excel (.xlsx)", type=['xlsx'])

if archivo_subido is not None:
    try:
        df = pd.read_excel(archivo_subido)
        st.info("Archivo cargado con éxito. Procesando datos...")

        # --- NUEVA LÓGICA DE FILTRADO POR FECHA VALOR ---
        # 1. Convertimos la columna a tipo datetime (Asegúrate de que en tu Excel se llame exactamente 'Fecha valor')
        df['Fecha valor'] = pd.to_datetime(df['Fecha valor'], errors='coerce')
        
        # 2. Encontramos la última fecha del registro (Día Actual del set de datos)
        ultima_fecha = df['Fecha valor'].max()
        
        # 3. Calculamos la fecha límite (4 días hacia atrás)
        fecha_limite = ultima_fecha - pd.Timedelta(days=4)
        
        # El objetivo sigue siendo la suma de todas las claves 50
        target_sum = df[df['Clave contabiliz.'] == 50]['Importe en moneda local'].sum()
        target_sum = round(float(target_sum), 2) # Evita errores de decimales infinitos
        
        # 4. Separar datos aplicando el FILTRO DE FECHAS para los candidatos Clave 40
        df_40 = df[
            (df['Clave contabiliz.'] == 40) & 
            (df['Fecha valor'] >= fecha_limite) & 
            (df['Fecha valor'] <= ultima_fecha)
        ].reset_index(drop=False)
        
        values_40 = df_40['Importe en moneda local'].values
        
        # Mostramos la información del filtro en la pantalla para tranquilidad del usuario
        st.write(f"**Última fecha detectada en el archivo:** {ultima_fecha.strftime('%Y-%m-%d')}")
        st.write(f"**Rango de búsqueda activo (Fecha Valor):** {fecha_limite.strftime('%Y-%m-%d')} ➡️ {ultima_fecha.strftime('%Y-%m-%d')}")
        st.write(f"**Objetivo de suma (Clave 50):** ${target_sum:,.2f}")
        st.write(f"**Registros candidatos en este rango (Clave 40):** {len(values_40)}")

        # 3. LÓGICA DE CONCILIACIÓN (MILP)
        if len(values_40) == 0:
            st.warning("⚠️ No se encontraron registros con Clave 40 dentro del rango de 4 días especificado.")
        else:
            if st.button("🚀 Ejecutar Conciliación"):
                with st.spinner('Buscando la combinación exacta en el rango de fechas...'):
                    n = len(values_40)
                    
                    # Configurar modelo matemático
                    c = np.zeros(n)
                    A = np.atleast_2d(values_40)
                    
                    # Usamos una holgura de 1 centavo para evitar que el optimizador falle por decimales de flotantes
                    tolerancia = 0.01
                    constraints = LinearConstraint(A, target_sum - tolerancia, target_sum + tolerancia)
                    bounds = Bounds(0, 1)
                    integrality = np.ones(n) # Variables binarias
                    
                    # Límite de tiempo de 60 segundos por seguridad
                    options = {'time_limit': 60.0}
                    
                    res = milp(c=c, constraints=constraints, bounds=bounds, integrality=integrality, options=options)
                    
                    if res.success:
                        # Marcar resultados en el DataFrame original
                        df['Conciliacion_Exacta'] = "No Conciliado"
                        selected_indices = np.where(np.round(res.x) == 1)[0]
                        selected_original_indices = df_40.iloc[selected_indices]['index'].values
                        
                        df.loc[selected_original_indices, 'Conciliacion_Exacta'] = "Conciliado (Grupo 40)"
                        df.loc[df['Clave contabiliz.'] == 50, 'Conciliacion_Exacta'] = "Conciliado (Grupo 50)"
                        
                        st.success(f"¡Logrado! Se encontraron {len(selected_indices)} registros en el rango de fechas que cuadran perfectamente.")
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
                        st.error("No se encontró una combinación exacta en ese rango de 4 días que sume esa cantidad.")
                        
    except Exception as e:
        st.error(f"Error técnico: Revisa que las columnas se llamen 'Clave contabiliz.', 'Importe en moneda local' y 'Fecha valor'. Detalle: {e}")
