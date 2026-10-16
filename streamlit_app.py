import streamlit as st
import pandas as pd
import httpx
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configuración de la página web
st.set_page_config(page_title="Analizador de Posición PNG", page_icon="👓", layout="centered")

st.title("👓 Analizador Inteligente de Fotos por SKU")
st.write("Sube tu Excel. El sistema identificará las imágenes conflictivas y armará el reporte de dos columnas.")

# Componente para subir el archivo Excel
uploaded_file = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

# Función original para verificar el link (Simula navegador Chrome)
def verificar_url_sync(url, client):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        # Intento 1: HEAD (Rápido)
        response = client.head(url, headers=headers, follow_redirects=True)
        if response.status_code == 200:
            return True
        # Intento 2: GET (Por si bloquean solicitudes HEAD)
        response = client.get(url, headers=headers, follow_redirects=True)
        return response.status_code == 200
    except:
        return False

# Función encargada de revisar una fila manteniendo las posiciones numéricas de tus 12 fotos
def analizar_fila_inteligente(row, columna_sku, columnas_imagenes, client):
    current_sku = str(row.get(columna_sku, '')).strip()
    if current_sku.endswith('.0'):
        current_sku = current_sku[:-2]
        
    if not current_sku or current_sku.lower() in ['nan', '']:
        return []
        
    hallazgos_fila = []
    
    # Recorremos las 12 columnas de imágenes en el orden físico del Excel
    for indice_posicion, nombre_columna in enumerate(columnas_imagenes, start=1):
        valor_celda = row.get(nombre_columna)
        url = str(valor_celda).strip() if pd.notna(valor_celda) else ""
        
        if url and url.lower() != 'nan' and url.startswith("http"):
            # FILTRO: Si el link está caído (da error) o contiene explícitamente '.png' en su texto
            link_caido = not verificar_url_sync(url, client)
            contiene_png = ".png" in url.lower()
            
            if link_caido or contiene_png:
                hallazgos_fila.append({
                    "sku+posicion del link": f"{current_sku}_{indice_posicion}",
                    "link de la imagen png": url
                })
                
    return hallazgos_fila

if uploaded_file is not None:
    try:
        # Cargamos el Excel completo
        df = pd.read_excel(uploaded_file)
        
        # DETECCIÓN INTELIGENTE DE LA COLUMNA SKU
        columna_sku_detectada = None
        nombres_comunes_sku = ['sku', 'modelo', 'model', 'count', 'id', 'codigo', 'item', 'code']
        
        for col in df.columns:
            if str(col).strip().lower() in nombres_comunes_sku:
                columna_sku_detectada = col
                break
                
        if not columna_sku_detectada and len(df.columns) > 0:
            columna_sku_detectada = df.columns[0]
            
        if not columna_sku_detectada:
            st.error("❌ No se pudo determinar la columna del SKU. Asegúrate de que el Excel tenga datos.")
        else:
            # Separamos las 12 columnas de imágenes del identificador SKU
            columnas_de_imagenes = [col for col in df.columns if col != columna_sku_detectada]
            
            st.info(f"🔍 Usando la columna **'{columna_sku_detectada}'** como SKU. Analizando las {len(columnas_de_imagenes)} columnas de fotos.")
            
            if st.button("🚀 Comenzar Análisis Masivo"):
                reporte_datos = []
                total = len(df)
                
                progreso_barra = st.progress(0)
                texto_progreso = st.empty()
                
                filas_dict = df.to_dict(orient='records')
                
                # Cliente HTTP optimizado para hilos concurrentes
                limits = httpx.Limits(max_keepalive_connections=5, max_connections=15)
                
                with httpx.Client(verify=False, limits=limits) as client:
                    with ThreadPoolExecutor(max_workers=10) as executor:
                        
                        # Mapeo de hilos en paralelo
                        futuros = {
                            executor.submit(analizar_fila_inteligente, fila, columna_sku_detectada, columnas_de_imagenes, client): fila 
                            for fila in filas_dict
                        }
                        
                        # as_completed refresca la barra en tiempo real de forma fluida
                        for index, futuro in enumerate(as_completed(futuros)):
                            try:
                                resultados_fila = futuro.result()
                                if resultados_fila:
                                    reporte_datos.extend(resultados_fila)
                            except Exception:
                                pass
                            
                            porcentaje = (index + 1) / total
                            texto_progreso.text(f"Analizando fila {index + 1} de {total}...")
                            progreso_barra.progress(porcentaje)
                
                texto_progreso.text("¡Análisis masivo finalizado!")
                
                # Si encontramos las fotos culpables, armamos el Excel final solicitado
                if reporte_datos:
                    df_reporte = pd.DataFrame(reporte_datos)
                    
                    st.subheader("📊 Vista Previa del Reporte Solicitado")
                    st.dataframe(df_reporte, use_container_width=True)
                    
                    # Guardar el reporte en memoria con formato .xlsx (Excel) usando openpyxl
                    output_excel = BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_reporte.to_excel(writer, index=False, sheet_name='Reporte Posiciones')
                    excel_procesado = output_excel.getvalue()
                    
                    st.warning(f"⚠️ Se detectaron {len(df_reporte)} posiciones con inconvenientes.")
                    st.download_button(
                        label="📥 Descargar reporte_posiciones.xlsx",
                        data=excel_procesado,
                        file_name="reporte_posiciones.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("🎉 ¡Excelente! No se encontraron imágenes conflictivas en las posiciones revisadas.")
                    
    except Exception as e:
        st.error(f"Error general al procesar el archivo: {e}")
