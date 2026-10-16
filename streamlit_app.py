import streamlit as st
import pandas as pd
import httpx
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configuración de la página web
st.set_page_config(page_title="Analizador Universal de Inventario", page_icon="👓", layout="centered")

st.title("👓 Analizador Inteligente de Fotos por SKU")
st.write("Sube cualquier Excel. El sistema detectará automáticamente el SKU y todas las columnas con enlaces.")

uploaded_file = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

# 1. OPTIMIZACIÓN: Recibe el cliente HTTP reutilizable para no abrir/cerrar conexiones
def verificar_url_sync(url, client):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        # Intento 1: HEAD (Rápido y no descarga la imagen completa)
        response = client.head(url, headers=headers, follow_redirects=True)
        if response.status_code == 200:
            return True
        # Intento 2: GET (Por si el servidor bloquea HEAD)
        response = client.get(url, headers=headers, follow_redirects=True)
        return response.status_code == 200
    except:
        return False

# Revisa la fila de forma dinámica
def analizar_fila_inteligente(row_dict, columna_sku, client):
    current_sku = str(row_dict.get(columna_sku, '')).strip()
    if current_sku.endswith('.0'):
        current_sku = current_sku[:-2]
        
    if not current_sku or current_sku.lower() in ['nan', '']:
        return None
        
    sku_tiene_error = False
    
    for clave, valor in row_dict.items():
        if clave == columna_sku:
            continue
            
        url = str(valor).strip()
        if url and url.lower() != 'nan' and url.startswith("http"):
            # Pasamos el cliente reutilizable
            if not verificar_url_sync(url, client):
                sku_tiene_error = True
                break 
                
    if sku_tiene_error:
        return current_sku
    return None

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # DETECCIÓN INTELIGENTE DE LA COLUMNA SKU
        columna_sku_detectada = None
        nombres_comunes_sku = ['sku', 'modelo', 'model', 'count', 'id', 'codigo', 'item']
        
        for col in df.columns:
            if str(col).strip().lower() in nombres_comunes_sku:
                columna_sku_detectada = col
                break
                
        if not columna_sku_detectada and len(df.columns) > 0:
            columna_sku_detectada = df.columns[0]
            
        if not columna_sku_detectada:
            st.error("❌ No se pudo determinar la columna del SKU. Asegúrate de que el Excel tenga datos.")
        else:
            st.info(f"🔍 Sistema configurado automáticamente: Usando la columna **'{columna_sku_detectada}'** como identificador de SKU.")
            
            if st.button("🚀 Comenzar Análisis Masivo"):
                skus_rotos = []
                total = len(df)
                
                progreso_barra = st.progress(0)
                texto_progreso = st.empty()
                
                # Convertimos eficientemente a diccionarios
                filas_dict = df.to_dict(orient='records')
                
                # 2. OPTIMIZACIÓN: Un único cliente HTTP con límites de conexión reutilizables
                limits = httpx.Limits(max_keepalive_connections=5, max_connections=20)
                
                with httpx.Client(verify=False, timeout=5.0, limits=limits) as client:
                    # Subimos max_workers a 15-20 ya que la mayoría del tiempo es espera de red (I/O Bound)
                    with ThreadPoolExecutor(max_workers=15) as executor:
                        
                        # Mapeamos los futures
                        futuro_a_sku = {
                            executor.submit(analizar_fila_inteligente, fila, columna_sku_detectada, client): fila 
                            for fila in filas_dict
                        }
                        
                        # 3. OPTIMIZACIÓN: as_completed permite actualizar la barra EN TIEMPO REAL
                        for index, futuro in enumerate(as_completed(futuro_a_sku)):
                            try:
                                sku_detectado = futuro.result()
                                if sku_detectado:
                                    skus_rotos.append(sku_detectado)
                            except Exception as exc:
                                # Captura errores individuales de hilos para que no se caiga toda la app
                                pass
                            
                            # Actualización fluida de la interfaz
                            porcentaje = (index + 1) / total
                            texto_progreso.text(f"Analizando fila {index + 1} de {total}...")
                            progreso_barra.progress(porcentaje)
                
                texto_progreso.text("¡Análisis masivo finalizado!")
                
                skus_finales = list(set([s for s in skus_rotos if s]))
                
                if skus_finales:
                    st.warning(f"⚠️ Se encontraron {len(skus_finales)} SKUs con imágenes rotas.")
                    txt_data = "\n".join(skus_finales)
                    
                    st.download_button(
                        label="📥 Descargar skus_con_links_rotos.txt",
                        data=txt_data,
                        file_name="skus_con_links_rotos.txt",
                        mime="text/plain"
                    )
                else:
                    st.success("🎉 ¡Excelente! Absolutamente todos los enlaces del documento están activos.")
                    
    except Exception as e:
        st.error(f"Error general al procesar el archivo: {e}")
