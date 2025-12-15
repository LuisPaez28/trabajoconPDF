import streamlit as st
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import zipfile
import io
import os
import tempfile

# Configuración de la página
st.set_page_config(page_title="PDF Toolset Pro", page_icon="🛠️", layout="centered")

# --- LÓGICA 1: SEPARAR PDF (EN MEMORIA) ---
def procesar_separacion(archivo_upload):
    zip_buffer = io.BytesIO()
    try:
        reader = PdfReader(archivo_upload)
        nombre_base = os.path.splitext(archivo_upload.name)[0]
        total = len(reader.pages)
        
        my_bar = st.progress(0, text="Iniciando separación...")

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                pdf_bytes = io.BytesIO()
                writer.write(pdf_bytes)
                
                nombre_salida = f"{nombre_base}_pag{i + 1}.pdf"
                zf.writestr(nombre_salida, pdf_bytes.getvalue())
                
                my_bar.progress((i + 1) / total, text=f"Procesando página {i+1} de {total}")
        
        my_bar.empty()
        zip_buffer.seek(0)
        return zip_buffer, total
    except Exception as e:
        st.error(f"Error en separación: {e}")
        return None, 0

# --- LÓGICA 2: CONVERTIR A WORD (CON TEMPFILES) ---
def procesar_conversion_word(archivo_upload):
    docx_buffer = io.BytesIO()
    
    # Creamos un directorio temporal que se borra al salir del 'with'
    with tempfile.TemporaryDirectory() as temp_dir:
        ruta_pdf_temp = os.path.join(temp_dir, "input.pdf")
        ruta_docx_temp = os.path.join(temp_dir, "output.docx")
        
        # 1. Guardamos el archivo subido en el disco temporal
        with open(ruta_pdf_temp, "wb") as f:
            f.write(archivo_upload.getbuffer())
        
        try:
            # 2. Ejecutamos la conversión
            cv = Converter(ruta_pdf_temp)
            cv.convert(ruta_docx_temp, start=0, end=None)
            cv.close()
            
            # 3. Leemos el resultado de vuelta a la memoria
            with open(ruta_docx_temp, "rb") as f:
                docx_buffer.write(f.read())
            
            docx_buffer.seek(0)
            return docx_buffer
            
        except Exception as e:
            st.error(f"Error en conversión: {e}")
            return None

# --- INTERFAZ DE USUARIO (FRONTEND) ---

st.title("🛠️ PDF Toolset Pro")
st.markdown("Tu navaja suiza para gestión documental. **Seguro, rápido y sin límites.**")

# Creamos las pestañas
tab1, tab2 = st.tabs(["✂️ Separar PDF", "📝 Convertir a Word"])

# === PESTAÑA 1: SEPARADOR ===
with tab1:
    st.header("Separar por Páginas")
    file_split = st.file_uploader("Sube tu PDF para separar", type="pdf", key="u_split")
    
    if file_split:
        if st.button("Separar Ahora", type="primary"):
            with st.spinner("Cortando el documento..."):
                zip_result, count = procesar_separacion(file_split)
            
            if zip_result:
                st.success(f"¡Éxito! {count} páginas extraídas.")
                st.download_button(
                    "⬇️ Descargar ZIP", 
                    zip_result, 
                    file_name="paginas_separadas.zip", 
                    mime="application/zip"
                )

# === PESTAÑA 2: CONVERTIDOR ===
with tab2:
    st.header("Convertir PDF a Word")
    st.warning("Nota: Funciona mejor con PDFs creados digitalmente (no escaneados).")
    
    file_conv = st.file_uploader("Sube tu PDF para convertir", type="pdf", key="u_conv")
    
    if file_conv:
        if st.button("Convertir a Word", type="primary"):
            with st.spinner("Analizando diseño y convirtiendo... (esto puede tardar unos segundos)"):
                word_result = procesar_conversion_word(file_conv)
            
            if word_result:
                st.success("¡Conversión completada!")
                nombre_descarga = os.path.splitext(file_conv.name)[0] + ".docx"
                
                st.download_button(
                    "⬇️ Descargar Word (.docx)", 
                    word_result, 
                    file_name=nombre_descarga, 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

st.markdown("---")
st.caption("Sistema local seguro. Los archivos se procesan en memoria temporal.")