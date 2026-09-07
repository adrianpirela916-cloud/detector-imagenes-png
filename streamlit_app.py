import streamlit as st
import pandas as pd
import requests
import io
from PIL import Image

st.set_page_config(page_title="Analizador de PNGs", page_icon="📸")
st.title("📊 Analizador Masivo de Imágenes (Fondos Transparentes)")
st.write("Sube tu archivo de Excel para detectar qué SKUs contienen imágenes PNG con transparencia real.")

# 1. Zona de carga del archivo
archivo_subido = st.file_uploader("Elige tu archivo de Excel (.xlsx)", type=["xlsx"])

def verificar_si_es_transparente_real(url):
    if not url or pd.isna(url) or not str(url).startswith("http"):
        return False
    try:
        response = requests.get(url, timeout=5, stream=True)
        if response.status_code != 200: 
            return False
        image_bytes = io.BytesIO(response.content)
        with Image.open(image_bytes) as img:
            if img.format != "PNG": 
                return False
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                alpha = img.convert("RGBA").split()[-1]
                if alpha.getextrema() == (255, 255): 
                    return False
                return True 
    except Exception:
        return False
    return False

if archivo_subido is not None:
    df = pd.read_excel(archivo_subido)
    st.success("✅ Archivo cargado correctamente.")
    
    if st.button("⚡ Empezar Análisis Masivo"):
        columna_sku = df.columns
        columnas_links = df.columns[1:13]
        
        skus_con_png_real = []
        
        # Barra de progreso visual en la página web
        barra_progreso = st.progress(0)
        texto_estado = st.empty()
        total_filas = len(df)
        
        for index, fila in df.iterrows():
            sku = str(fila[columna_sku]).strip()
            texto_estado.text(f"🔍 Analizando fila {index + 1} de {total_filas}... SKU: {sku}")
            barra_progreso.progress((index + 1) / total_filas)
            
            encontrado_transparente = False
            for col in columnas_links:
                url = fila[col]
                if verificar_si_es_transparente_real(url):
                    encontrado_transparente = True
                    break
                    
            if encontrado_transparente:
                skus_con_png_real.append(sku)
        
        texto_estado.text("¡Análisis completado con éxito! 🎉")
        
        # Generar el archivo de texto en memoria listo para descargar
        resultado_txt = "\n".join(skus_con_png_real)
        
        st.download_button(
            label="💾 Descargar Reporte (TXT)",
            data=resultado_txt,
            file_name="reporte_skus_png_real.txt",
            mime="text/plain"
        )
