import streamlit as st
import pandas as pd
import httpx
from io import BytesIO
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configuración estética de la aplicación
st.set_page_config(page_title="Detector de Posición PNG", page_icon="🎯", layout="centered")

st.title("🎯 Detector de Posición de Imágenes PNG")
st.write("Sube tu Excel de 13 columnas (Columna A: SKU | Columnas B a M: Links de imágenes).")

# Componente de subida de archivos
uploaded_file = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

def detectar_transparencia_en_url(url, client):
    """
    Descarga la imagen y evalúa si califica como una imagen con fondo transparente real.
    """
    url_clean = url.lower().split('?')[0]
    
    # Si explícitamente termina en extensiones opacas comunes, evitamos la descarga
    if url_clean.endswith('.jpg') or url_clean.endswith('.jpeg'):
        return False

    # Criterio 0: Si por la pura URL ya vemos que es un formato que maneja transparencia nativa
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
        
        # Criterio 1: Color de transparencia en metadatos (Modo Paleta)
        if 'transparency' in img.info:
            return True
            
        # Criterio 2: Canales nativos de transparencia (RGBA o LA) analizados por umbral de ruido
        if img.mode in ('RGBA', 'LA'):
            alpha = img.convert('RGBA').split()[-1]
            datos_alfa = list(alpha.getdata())
            pixeles_totales = len(datos_alfa)
            pixeles_transparentes = sum(1 for p in datos_alfa if p < 250) 
            
            # Al menos el 1% de la imagen debe ser transparente para descartar ruido digital en JPGs falsos
            if (pixeles_transparentes / pixeles_totales) > 0.01:
                return True
                
        # Criterio 3: Formato explícito reportado por Pillow
        if img.format in ('PNG', 'WEBP'):
            return True
            
        return False
    except:
        return False

def procesar_fila_catalogo(row_dict, nombre_columna_sku, lista_columnas_imagenes, client):
    """
    Escanea horizontalmente las columnas de imágenes de un SKU (Columnas B a M).
    Retorna los hallazgos en formato SKU_Posición y su respectivo Link.
    """
    current_sku = str(row_dict.get(nombre_columna_sku, '')).strip()
    if current_sku.endswith('.0'):
        current_sku = current_sku[:-2]
        
    if not current_sku or current_sku.lower() in ['nan', '']:
        return []
        
    hallazgos_sku = []
    
    # Recorremos las columnas asignando el índice correspondiente (B=1, C=2, D=3...)
    for indice_posicion, nombre_columna in enumerate(lista_columnas_imagenes, start=1):
        valor_celda = row_dict.get(nombre_columna)
        url = str(valor_celda).strip() if pd.notna(valor_celda) else ""
        
        if url and url.lower() != 'nan' and url.startswith("http"):
            if detectar_transparencia_en_url(url, client):
                hallazgos_sku.append({
                    "sku+posicion del link": f"{current_sku}_{indice_posicion}",
                    "link de la imagen png": url
                })
                
    return hallazgos_sku

if uploaded_file is not None:
    try:
        # Leemos el Excel
        df = pd.read_excel(uploaded_file)
        
        if len(df.columns) < 2:
            st.error("❌ El archivo Excel debe tener al menos la columna de SKU y una columna de imagen.")
        else:
            # Requisito 1: La columna A (índice 0) siempre se toma como la columna del SKU
            columna_sku_detectada = df.columns[0]
            
            # Las columnas de la B a la M (del índice 1 en adelante) son las imágenes
            columnas_de_imagenes = list(df.columns[1:13]) 
            
            st.info(f"📋 **Columna SKU (A):** '{columna_sku_detectada}' | **Columnas de Imágenes (B a M):** {len(columnas_de_imagenes)} columnas detectadas.")
            
            if st.button("🚀 Iniciar Escaneo Quirúrgico"):
                reporte_final_datos = []
                total_filas = len(df)
                
                progreso_barra = st.progress(0)
                texto_progreso = st.empty()
                
                filas_lista_dict = df.to_dict(orient='records')
                
                # Configuración de límites HTTP
                limits = httpx.Limits(max_keepalive_connections=5, max_connections=15)
                
                with httpx.Client(verify=False, limits=limits) as client:
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        
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
                            
                            porcentaje = (index + 1) / total_filas
                            texto_progreso.text(f"Analizando producto {index + 1} de {total_filas}...")
                            progreso_barra.progress(porcentaje)
                
                texto_progreso.text("¡Procesamiento masivo finalizado!")
                
                # Requisito 2 y 3: Generación del reporte en Excel con exactamente dos columnas
                if reporte_final_datos:
                    df_reporte = pd.DataFrame(reporte_final_datos)
                    
                    # Aseguramos el orden exacto de las dos columnas solicitadas
                    df_reporte = df_reporte[["sku+posicion del link", "link de la imagen png"]]
                    
                    st.subheader("📊 Vista Previa del Reporte generado")
                    st.dataframe(df_reporte, use_container_width=True)
                    
                    # Escritura a buffer de memoria en formato Excel (.xlsx)
                    output_excel = BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_reporte.to_excel(writer, index=False, sheet_name='Reporte Transparencias')
                    excel_procesado = output_excel.getvalue()
                    
                    st.success(f"⚠️ Se detectaron {len(df_reporte)} links con transparencias erróneas.")
                    st.download_button(
                        label="📥 Descargar reporte_posiciones_png.xlsx",
                        data=excel_procesado,
                        file_name="reporte_posiciones_png.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("🎉 ¡Excelente! No se detectaron imágenes transparentes en el archivo.")
                    
    except Exception as e:
        st.error(f"Error crítico al procesar el documento Excel: {e}")
