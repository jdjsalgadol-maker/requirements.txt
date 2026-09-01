import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io
import re
from datetime import timedelta
import gc

st.set_page_config(page_title="Organizador de Documentos", layout="wide")

st.title("📂 Organizador y Depurador de Documentos Contables")
st.write("Aplica automáticamente las 9 reglas de limpieza y formato.")

# --- BARRA LATERAL: PARÁMETROS DE EJECUCIÓN ---
st.sidebar.header("⚙️ Parámetros de Ejecución")
fecha_ej = st.sidebar.date_input("Fecha de Ejecución (Día actual)", pd.to_datetime("today"))

st.sidebar.subheader("Regla 7: Festivos")
hubo_festivo = st.sidebar.checkbox("¿Hubo un día festivo recientemente? (Resta 1 día hábil)")

st.sidebar.subheader("Regla 9: Depuración Opcional")
eliminar_fe_contab = st.sidebar.checkbox("Eliminar registros donde Fe.contabilización (Col D) sea igual al día de ejecución")

def calcular_dia_habil_anterior(fecha, festivo):
    d = fecha - timedelta(days=1)
    while d.weekday() > 4:  # 5=Sábado, 6=Domingo
        d -= timedelta(days=1)
    if festivo:
        d -= timedelta(days=1)
        while d.weekday() > 4:
            d -= timedelta(days=1)
    return d

fecha_habil_anterior = calcular_dia_habil_anterior(fecha_ej, hubo_festivo)
st.sidebar.info(f"Día hábil anterior calculado (Regla 7): **{fecha_habil_anterior.strftime('%Y-%m-%d')}**")

# --- CARGA DE ARCHIVO ---
archivo_subido = st.file_uploader("Sube el archivo Excel (Ej: Documento original.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    if st.button("🚀 Iniciar Organización", use_container_width=True):
        try:
            # PASO 1: Lectura ultrarrápida de datos
            st.info("⏳ Paso 1/4: Leyendo y analizando estructura del archivo...")
            df = pd.read_excel(archivo_subido)
            filas_iniciales = len(df)
            max_filas_reales = filas_iniciales + 1
            
            col_A = 'Asignación'
            col_B = 'Nº documento'
            col_D = 'Fe.contabilización'
            col_E = 'Fecha de documento'
            col_F = 'Fecha valor'
            col_G = 'Clave contabiliz.'
            col_I = 'Importe en moneda local'
            col_L = 'Clave referencia 3'

            # PASO 2: Escaneo de celdas amarillas protegido contra bloqueos de RAM
            st.info("⏳ Paso 2/4: Detectando registros en amarillo (Regla 5)...")
            archivo_subido.seek(0)
            wb = openpyxl.load_workbook(archivo_subido, data_only=True)
            ws = wb.active
            filas_amarillas = set()
            
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=max_filas_reales), start=2):
                for cell in row[:5]: # Revisa las primeras celdas de cada fila
                    if cell.fill and cell.fill.start_color and cell.fill.start_color.index != '00000000':
                        c_index = str(cell.fill.start_color.index).upper()
                        if 'FFFF00' in c_index or 'FFFFEE09' in c_index or c_index == '4':
                            filas_amarillas.add(row_idx - 2)
                            break
                            
            # Liberar memoria de openpyxl
            wb.close()
            del wb
            gc.collect()

            st.info("⏳ Paso 3/4: Rellenando bancos y aplicando reglas de depuración...")
            
            # REGLA 5: Eliminar amarillas
            df = df.drop(index=list(filas_amarillas)).reset_index(drop=True)

            # REGLA 3: Rellenar bancos (¡Se hace ANTES de eliminar las fechas para no perder los encabezados!)
            mapeo_bancos = {
                "1110056001": "CUENTA 1110056001", "1110056101": "BANCO DE BOGOTA",
                "1110056201": "BANCO DAVIBANK S.A.", "1110056301": "BANCOLOMBIA S.A.",
                "1110056401": "BANCO CAJA SOCIAL S.", "1110056501": "BANCO DAVIVIENDA S.A",
                "1110056601": "BANCO BILBAO VIZCAYA", "1110056701": "BANCO AGRARIO DE COL",
                "1120055001": "BANCO COMERCIAL AV V", "1120055101": "BANCO DE OCCIDENTE",
                "1120055301": "BANCO GNB SUDAMERIS",
            }
            
            bancos_procesados = []
            current_bank = ""
            for _, row in df.iterrows():
                asig_val = str(row.get(col_A, ""))
                banco_val = row.get(col_L, None)
                
                # 1. Chequear si la fila es el encabezado agrupador
                if "cuenta de mayor" in asig_val.lower():
                    m = re.search(r'(\d{6,})', asig_val)
                    if m:
                        current_bank = mapeo_bancos.get(m.group(1), f"CUENTA {m.group(1)} (sin mapear)")
                
                # 2. Si la fila ya trae explícitamente su banco, lo adopta
                if pd.notnull(banco_val) and str(banco_val).strip() != "" and str(banco_val).lower() != "nan":
                    current_bank = str(banco_val).strip()
                        
                bancos_procesados.append(current_bank)

            df[col_L] = bancos_procesados
            
            # Limpiar las filas agrupadoras (Cuenta de mayor) ahora que ya robaron el banco
            df = df[~df[col_A].astype(str).str.contains("cuenta de mayor", case=False, na=False)]

            # Estandarizar Fechas
            df[col_E] = pd.to_datetime(df[col_E], errors='coerce').dt.date
            df[col_F] = pd.to_datetime(df[col_F], errors='coerce').dt.date
            df[col_D] = pd.to_datetime(df[col_D], errors='coerce').dt.date

            # REGLA 1: Eliminar documentos con "Fecha de documento" igual al día que ejecuto
            df = df[df[col_E] != fecha_ej]

            # REGLA 9: Eliminar Columna D igual al día de ejecución (Opcional)
            if eliminar_fe_contab:
                df = df[df[col_D] != fecha_ej]

            # REGLA 2: Valores Absolutos (Positivos) en Columna I
            df[col_I] = pd.to_numeric(df[col_I], errors='coerce').fillna(0).abs()

            # REGLA 4: Eliminar documentos de más de 60 días atrás (Columna F)
            limite_60_dias = fecha_ej - timedelta(days=60)
            df = df[df[col_F] >= limite_60_dias]

            # REGLA 7: Eliminar OCCIDENTE / GNB SUDAMERIS con Clave 40 del día hábil anterior
            df[col_G] = df[col_G].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            bancos_r7 = ['BANCO DE OCCIDENTE', 'BANCO GNB SUDAMERIS']
            
            mask_r7 = (df[col_G] == '40') & (df[col_L].str.upper().str.contains('|'.join(bancos_r7), na=False)) & (df[col_F] == fecha_habil_anterior)
            df = df[~mask_r7]

            # REGLA 8: Ordenar por importe de menor a mayor
            df = df.sort_values(by=col_I, ascending=True).reset_index(drop=True)
            
            # Formatear fechas para Excel
            for c in ['Fe.contabilización', 'Fecha de documento', 'Fecha valor']:
                if c in df.columns:
                    df[c] = pd.to_datetime(df[c], errors='coerce').dt.strftime('%d/%m/%Y')

            # REGLA 6: EXPORTACIÓN CON FUENTE ROJA
            st.info("⏳ Paso 4/4: Generando archivo Excel con fuentes a color...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Documentos_Limpios')
                workbook  = writer.book
                worksheet = writer.sheets['Documentos_Limpios']
                
                formato_rojo = workbook.add_format({'font_color': 'red'})
                idx_col_g = df.columns.get_loc(col_G)
                
                for row_num in range(len(df)):
                    if str(df.iloc[row_num, idx_col_g]).strip() == '40':
                        worksheet.set_row(row_num + 1, None, formato_rojo)

            filas_finales = len(df)
            st.success("✅ ¡Organización completada exitosamente!")
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("Filas Iniciales (Crudo)", filas_iniciales)
            col2.metric("Filas Amarillas Eliminadas", len(filas_amarillas))
            col3.metric("Filas Finales Procesadas", filas_finales)

            st.download_button(
                label="📥 Descargar Documento Depurado (Despues.xlsx)",
                data=output.getvalue(),
                file_name="Despues.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
            st.exception(e)
