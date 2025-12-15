import streamlit as st
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import pdfplumber
import pandas as pd
import zipfile
import io
import os
import tempfile

# Configuración de la página
st.set_page_config(page_title="PDF Toolset Pro", page_icon="🛠️", layout="centered")

# ==========================================
# LÓGICA 1: SEPARAR PDF
# ==========================================
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

# ==========================================
# LÓGICA 2: CONVERTIR A WORD
# ==========================================
def procesar_conversion_word(archivo_upload):
    docx_buffer = io.BytesIO()
    with tempfile.TemporaryDirectory() as temp_dir:
        ruta_pdf_temp = os.path.join(temp_dir, "input.pdf")
        ruta_docx_temp = os.path.join(temp_dir, "output.docx")
        
        with open(ruta_pdf_temp, "wb") as f:
            f.write(archivo_upload.getbuffer())
        
        try:
            cv = Converter(ruta_pdf_temp)
            cv.convert(ruta_docx_temp, start=0, end=None)
            cv.close()
            
            with open(ruta_docx_temp, "rb") as f:
                docx_buffer.write(f.read())
            
            docx_buffer.seek(0)
            return docx_buffer
        except Exception as e:
            st.error(f"Error en conversión Word: {e}")
            return None

# ==========================================
# LÓGICA 3: CONVERTIR A EXCEL (UNIFICADO)
# ==========================================
def procesar_conversion_excel(archivo_upload):
    excel_buffer = io.BytesIO()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        ruta_pdf_temp = os.path.join(temp_dir, "input_excel.pdf")
        ruta_excel_temp = os.path.join(temp_dir, "output.xlsx")
        
        with open(ruta_pdf_temp, "wb") as f:
            f.write(archivo_upload.getbuffer())
            
        try:
            tablas_encontradas = False
            
            # Configuraciones de estrategias
            config_bordes = {"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 3}
            config_texto = {"vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 3}

            with pdfplumber.open(ruta_pdf_temp) as pdf:
                with pd.ExcelWriter(ruta_excel_temp, engine='openpyxl') as writer:
                    
                    row_counter = 0 # Contador para saber en qué fila escribir
                    hoja_unica = "Datos_Consolidados"
                    
                    for i, page in enumerate(pdf.pages):
                        # Intentamos extraer tablas
                        tables = page.extract_tables(config_bordes)
                        if not tables:
                            tables = page.extract_tables(config_texto)
                        
                        for table in tables:
                            # Limpieza de filas vacías
                            clean_table = [row for row in table if any(cell is not None and cell != '' for cell in row)]
                            
                            if clean_table:
                                if len(clean_table) > 1:
                                    df = pd.DataFrame(clean_table[1:], columns=clean_table[0])
                                else:
                                    df = pd.DataFrame(clean_table)
                                
                                # Escribimos en la misma hoja, pero desplazando la fila (startrow)
                                df.to_excel(writer, sheet_name=hoja_unica, startrow=row_counter, index=False)
                                
                                # Aumentamos el contador: filas de datos + 1 fila de encabezado + 2 filas de espacio libre
                                row_counter += len(df) + 3 
                                tablas_encontradas = True
            
            if not tablas_encontradas:
                return "NO_TABLES"

            with open(ruta_excel_temp, "rb") as f:
                excel_buffer.write(f.read())
                
            excel_buffer.seek(0)
            return excel_buffer

        except Exception as e:
            st.error(f"Error en conversión Excel: {e}")
            return None

# ==========================================
# INTERFAZ GRÁFICA (FRONTEND)
# ==========================================

st.title("🛠️ PDF Toolset Pro")
st.markdown("Tu navaja suiza para gestión documental. **Seguro, rápido y sin límites.**")

tab1, tab2, tab3 = st.tabs(["✂️ Separar PDF", "📝 A Word", "📊 A Excel"])

# === PESTAÑA 1: SEPARADOR ===
with tab1:
    st.header("Separar por Páginas")
    file_split = st.file_uploader("Sube tu PDF", type="pdf", key="u_split")
    if file_split and st.button("Separar Ahora", type="primary"):
        with st.spinner("Cortando..."):
            zip_result, count = procesar_separacion(file_split)
        if zip_result:
            st.success(f"¡Hecho! {count} páginas extraídas.")
            st.download_button("⬇ Descargar ZIP", zip_result, "paginas.zip", "application/zip")

# === PESTAÑA 2: WORD ===
with tab2:
    st.header("De PDF a Word")
    st.info("Ideal para cartas, contratos y textos.")
    file_word = st.file_uploader("Sube tu PDF", type="pdf", key="u_word")
    if file_word and st.button("Convertir a Word", type="primary"):
        with st.spinner("Convirtiendo..."):
            word_result = procesar_conversion_word(file_word)
        if word_result:
            st.success("¡Conversión lista!")
            st.download_button("⬇ Descargar Word", word_result, "documento.docx", 
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# === PESTAÑA 3: EXCEL ===
with tab3:
    st.header("De PDF a Excel")
    st.info("ℹ Todas las tablas detectadas se pondrán en **una sola hoja**, una debajo de la otra.")
    
    file_excel = st.file_uploader("Sube tu PDF", type="pdf", key="u_excel")
    
    if file_excel:
        if st.button("Extraer Tablas a Excel", type="primary"):
            with st.spinner("Unificando tablas en Excel..."):
                excel_result = procesar_conversion_excel(file_excel)
            
            if excel_result == "NO_TABLES":
                st.error("No detectamos tablas claras.")
            elif excel_result:
                st.success("¡Tablas extraídas y unificadas!")
                st.download_button(
                    "⬇️ Descargar Excel", 
                    excel_result, 
                    "tablas_unificadas.xlsx", 
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.caption("Sistema de procesamiento seguro")