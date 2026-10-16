import streamlit as st
import pandas as pd
import httpx
from io import BytesIO
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configuración estética de la aplicación
st.set_page_config(page_title="Detector de Posición PNG", page_icon="🎯", layout="centered")

st.title("🎯 Detector de Posición de Imágenes PNG")
st.write("Sube tu Excel de 12 columnas. El sistema unirá el SKU con la posición exacta de la imagen transparente.")

# Componente de subida de archivos
uploaded_file = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

def detectar_transparencia_en_url(url, client):
    """
    Descarga la imagen y evalúa bajo múltiples criterios si califica como PNG/WebP transparente
    o si su extensión la delata como imagen transparente.
    """
    # Criterio 0: Si por la pura URL ya vemos que es un formato que maneja transparencia
    url_clean = url.lower().split('?')[0]
    if url_clean.endswith('.png') or url_clean.endswith('.webp'):
        return True

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        response = client.get(url, headers=headers, timeout=6.0, follow_redirects=True)
        if response.status_code != 200:
            return False
            
        img = Image.open(BytesIO(response.content))
        
        # Criterio 1: Tiene un color de transparencia asignado en metadatos (Modo P / Paleta)
        if 'transparency' in img.info:
            return True
            
        # Criterio 2: Maneja canales nativos de transparencia (RGBA o LA)
        if img.mode in ('RGBA', 'LA'):
            alpha = img.convert('RGBA').split()[-1]
            extrema = alpha.getextrema()
            # Umbral flexible: Si al menos un píxel no es totalmente opaco (255), tiene transparencia
            if extrema[0] < 255:
                return True
                
        # Criterio 3: Formato explícito reportado por Pillow
        if img.format in ('PNG', 'WEBP'):
            return True
            
        return False
    except:
        return False

def procesar_fila_catalogo(row_dict, columna_sku, columnas_imagenes, client):
    """
    Escanea horizontalmente las columnas de imágenes de un SKU.
    Retorna una lista de hallazgos con el formato solicitado (SKU_Posición).
    """
    current_sku = str(row_dict.get(columna_sku, '')).strip()
    if current_sku.endswith('.0'):
        current_sku = current_sku[:-2]
        
    if not current_sku or current_sku.lower() in ['nan', '']:
        return []
        
    hallazgos_sku = []
    
    # Recorremos las columnas de imágenes manteniendo el orden/posición de las 12 columnas
    for indice_posicion, nombre_columna in enumerate(columnas_imagenes, start=1):
        valor_celda = row_dict.get(nombre_columna)
        url = str(valor_celda).strip() if pd.notna(valor_celda) else ""
        
        if url and url.lower() != 'nan' and url.startswith("http"):
            # Analizamos si la imagen en esta posición es la PNG transparente
            if detectar_transparencia_en_url(url, client):
                hallazgos_sku.append({
                    "sku+posicion del link": f"{current_sku}_{indice_posicion}",
                    "link de la imagen png": url
                })
                
    return hallazgos_sku

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Identificación automática de la columna identificadora (SKU)
        columna_sku_detectada = None
        nombres_comunes_sku = ['sku', 'modelo', 'model', 'id', 'codigo', 'item', 'artículo', 'code']
        
        for col in df.columns:
            if str(col).strip().lower() in nombres_comunes_sku:
                columna_sku_detectada = col
                break
                
        if not columna_sku_detectada and len(df.columns) > 0:
            columna_sku_detectada = df.columns[0]
            
        if not columna_sku_detectada:
            st.error("❌ No se pudo determinar la columna del SKU en el archivo.")
        else:
            # Las columnas de imágenes serán todas las demás columnas del Excel excluyendo el SKU
            columnas_de_imagenes = [col for col in df.columns if col != columna_sku_detectada]
            
            st.info(f"🔍 SKU identificado en columna: **'{columna_sku_detectada}'**. Escaneando las **{len(columnas_de_imagenes)}** columnas de imágenes restantes.")
            
            if st.button("🚀 Iniciar Escaneo Quirúrgico"):
                reporte_final_datos = []
                total_filas = len(df)
                
                progreso_barra = st.progress(0)
                texto_progreso = st.empty()
                
                filas_lista_dict = df.to_dict(orient='records')
                
                # Configuración de límites del cliente HTTP para evitar sobrecargas
                limits = httpx.Limits(max_keepalive_connections=5, max_connections=15)
                
                with httpx.Client(verify=False, limits=limits) as client:
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        
                        # Lanzamos los hilos concurrentes
                        futuros = {
                            executor.submit(procesar_fila_catalogo, fila, columna_sku_detectada, columnas_de_imagenes, client): fila 
                            for fila in filas_lista_dict
                        }
                        
                        for index, futuro in enumerate(as_completed(futuros)):
                            try:
                                resultados_fila = futuro.result()
                                if resultados_fila:
                                    reporte_final_datos.extend(resultados_fila)
                            except Exception:
                                pass
                            
                            # Actualización fluida de la barra de progreso
                            porcentaje = (index + 1) / total_filas
                            texto_progreso.text(f"Analizando producto {index + 1} de {total_filas}...")
                            progreso_barra.progress(porcentaje)
                
                texto_progreso.text("¡Procesamiento masivo finalizado!")
                
                # Generación del reporte si se encontraron culpables
                if reporte_final_datos:
                    df_reporte = pd.DataFrame(reporte_final_datos)
                    
                    st.subheader("📊 Vista Previa del Reporte de Transparentes")
                    st.dataframe(df_reporte, use_container_width=True)
                    
                    # Escritura directa a buffer de memoria en formato Excel (.xlsx)
                    output_excel = BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_reporte.to_excel(writer, index=False, sheet_name='Reporte PNG')
                    excel_procesado = output_excel.getvalue()
                    
                    st.success(f"⚠️ Se detectaron {len(df_reporte)} posiciones con imágenes transparentes.")
                    st.download_button(
                        label="📥 Descargar reporte_posiciones_png.xlsx",
                        data=excel_procesado,
                        file_name="reporte_posiciones_png.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("🎉 ¡Excelente! No se detectaron imágenes PNG con canales de transparencia en las celdas analizadas.")
                    
    except Exception as e:
        st.error(f"Error crítico al procesar el documento Excel: {e}")
