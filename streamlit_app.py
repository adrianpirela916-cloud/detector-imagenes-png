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
    Descarga la imagen e inspecciona los canales binarios reales para buscar transparencias,
    sin importar qué extensión afirme tener la URL en el Excel.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        # Descarga forzatoria del archivo para validar el contenido real
        response = client.get(url, headers=headers, timeout=8.0, follow_redirects=True)
        if response.status_code != 200:
            return False
            
        img = Image.open(BytesIO(response.content))
        
        # Criterio 1: Color de transparencia asignado en metadatos (Modo Paleta)
        if 'transparency' in img.info:
            return True
            
        # Criterio 2: La imagen real se descargó con canales transparentes (RGBA o LA) o es PNG/WEBP nativo
        if img.mode in ('RGBA', 'LA') or img.format in ('PNG', 'WEBP'):
            alpha = img.convert('RGBA').split()[-1]
            colores = alpha.getcolors(maxcolors=65536)
            
            if colores:
                # Si arroja un solo color y este es 255 (totalmente opaco), no hay transparencia
                if len(colores) == 1 and colores[0][1] == 255:
                    return False
                
                # Contamos píxeles no opacos (valores menores a 255)
                pixeles_transparentes = sum(count for count, val in colores if val < 255)
                pixeles_totales = sum(count for count, val in colores)
                
                # Tolerancia ultra-baja: Si al menos el 0.05% de la imagen es transparente, se reporta
                if (pixeles_transparentes / pixeles_totales) > 0.0005:
                    return True
                    
        return False
    except:
        return False

def procesar_fila_catalogo(row_dict, nombre_columna_sku, lista_columnas_imagenes, client):
    """
    Escanea horizontalmente las columnas de imágenes de un SKU (Columnas B a M).
    """
    current_sku = str(row_dict.get(nombre_columna_sku, '')).strip()
    if current_sku.endswith('.0'):
        current_sku = current_sku[:-2]
        
    if not current_sku or current_sku.lower() in ['nan', '']:
        return []
        
    hallazgos_sku = []
    
    # Recorremos manteniendo la posición exacta solicitada (B=1, C=2, D=3...)
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
        df = pd.read_excel(uploaded_file)
        
        if len(df.columns) < 2:
            st.error("❌ El archivo Excel debe tener al menos la columna de SKU y una columna de imagen.")
        else:
            # Columna A (índice 0) siempre se toma como SKU
            columna_sku_detectada = df.columns[0]
            
            # Columnas de la B a la M (índices del 1 al 12)
            columnas_de_imagenes = list(df.columns[1:13]) 
            
            st.info(f"📋 **Columna SKU (A):** '{columna_sku_detectada}' | **Columnas de Imágenes (B a M):** {len(columnas_de_imagenes)} columnas asignadas.")
            
            if st.button("🚀 Iniciar Escaneo Quirúrgico"):
                reporte_final_datos = []
                total_filas = len(df)
                
                progreso_barra = st.progress(0)
                texto_progreso = st.empty()
                
                filas_lista_dict = df.to_dict(orient='records')
                
                # Configuración de conexiones simultáneas optimizada
                limits = httpx.Limits(max_keepalive_connections=10, max_connections=30)
                
                with httpx.Client(verify=False, limits=limits) as client:
                    with ThreadPoolExecutor(max_workers=5) as executor:
                        
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
                
                if reporte_final_datos:
                    df_reporte = pd.DataFrame(reporte_final_datos)
                    df_reporte = df_reporte[["sku+posicion del link", "link de la imagen png"]]
                    
                    st.subheader("📊 Vista Previa del Reporte generado")
                    st.dataframe(df_reporte, use_container_width=True)
                    
                    output_excel = BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_reporte.to_excel(writer, index=False, sheet_name='Reporte Transparencias')
                    excel_procesado = output_excel.getvalue()
                    
                    st.success(f"⚠️ Se detectaron {len(df_reporte)} links con transparencias ocultas.")
                    st.download_button(
                        label="📥 Descargar reporte_posiciones_png.xlsx",
                        data=excel_procesado,
                        file_name="reporte_posiciones_png.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("🎉 ¡Excelente! No se detectaron imágenes transparentes reales en el archivo.")
                    
    except Exception as e:
        st.error(f"Error crítico al procesar el documento Excel: {e}")
