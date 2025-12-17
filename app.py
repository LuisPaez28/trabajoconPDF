import streamlit as st
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import zipfile
import io
import os
import tempfile

# Configuración de la página (Título y Layout)
st.set_page_config(
    page_title="Dividir, unir o convertir PDF a Word",
    page_icon="🛠️",
    layout="centered"
)

# ==========================================
# LÓGICA 1: SEPARAR PDF (Split)
# ==========================================
def procesar_separacion(archivo_upload):
    """Separa un PDF en páginas individuales y devuelve un ZIP en memoria."""
    zip_buffer = io.BytesIO()
    
    try:
        reader = PdfReader(archivo_upload)
        nombre_base = os.path.splitext(archivo_upload.name)[0]
        total_paginas = len(reader.pages)
        
        # Barra de progreso
        barra = st.progress(0, text="Iniciando separación...")

        # Creamos un ZIP en memoria
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                
                # Escribimos la página en un buffer temporal
                pdf_bytes = io.BytesIO()
                writer.write(pdf_bytes)
                
                # Añadimos al ZIP
                nombre_salida = f"{nombre_base}_pag{i + 1}.pdf"
                zf.writestr(nombre_salida, pdf_bytes.getvalue())
                
                # Actualizar barra
                barra.progress((i + 1) / total_paginas, text=f"Procesando página {i+1} de {total_paginas}")

        barra.empty()
        zip_buffer.seek(0)
        return zip_buffer, total_paginas

    except Exception as e:
        st.error(f"Error al separar: {str(e)}")
        return None, 0

# ==========================================
# LÓGICA 2: UNIR PDF (Merge - ACTUALIZADA)
# ==========================================
def procesar_union(lista_archivos):
    """Recibe una lista de archivos y devuelve un solo PDF en memoria."""
    output_buffer = io.BytesIO()
    
    try:
        writer = PdfWriter()
        barra = st.progress(0, text="Uniendo archivos...")
        total = len(lista_archivos)

        for i, archivo in enumerate(lista_archivos):
            # IMPORTANTE: Aseguramos que el puntero de lectura esté al inicio
            archivo.seek(0)
            writer.append(archivo)
            barra.progress((i + 1) / total, text=f"Uniendo archivo {i+1} de {total}")
        
        writer.write(output_buffer)
        barra.empty()
        output_buffer.seek(0)
        return output_buffer

    except Exception as e:
        st.error(f"Error al unir: {str(e)}")
        return None

# ==========================================
# LÓGICA 3: CONVERTIR A WORD
# ==========================================
def procesar_word(archivo_upload):
    """
    Convierte PDF a Word.
    Nota: pdf2docx necesita archivos físicos, así que usamos tempfile.
    """
    docx_buffer = io.BytesIO()
    
    # Creamos un directorio temporal seguro que se borra al terminar
    with tempfile.TemporaryDirectory() as temp_dir:
        ruta_pdf_temp = os.path.join(temp_dir, "input.pdf")
        ruta_docx_temp = os.path.join(temp_dir, "output.docx")
        
        # 1. Guardar el archivo subido en el disco temporal
        with open(ruta_pdf_temp, "wb") as f:
            f.write(archivo_upload.getbuffer())
        
        try:
            # 2. Convertir
            cv = Converter(ruta_pdf_temp)
            cv.convert(ruta_docx_temp, start=0, end=None)
            cv.close()
            
            # 3. Leer el resultado de vuelta a memoria
            with open(ruta_docx_temp, "rb") as f:
                docx_buffer.write(f.read())
            
            docx_buffer.seek(0)
            return docx_buffer
            
        except Exception as e:
            st.error(f"Error en conversión Word: {str(e)}")
            return None

# ==========================================
# INTERFAZ GRÁFICA (Streamlit Frontend)
# ==========================================

st.title("Dividir, unir o convertir PDF a Word")
st.markdown("""
Herramienta todo en uno para gestionar tus PDFs.
*Procesamiento seguro en memoria (RAM).*
""")

# Pestañas
tab_split, tab_merge, tab_word = st.tabs(["✂️ Separar", "🔗 Unir", "📝 PDF a Word"])

# --- PESTAÑA 1: SEPARAR ---
with tab_split:
    st.header("Separar PDF por páginas")
    file_split = st.file_uploader("Sube tu PDF", type="pdf", key="split")
    
    if file_split:
        if st.button("Separar Ahora", type="primary"):
            zip_result, count = procesar_separacion(file_split)
            
            if zip_result:
                st.success(f"¡Listo! Se crearon {count} archivos.")
                st.download_button(
                    label="⬇️ Descargar ZIP con páginas",
                    data=zip_result,
                    file_name="paginas_separadas.zip",
                    mime="application/zip"
                )

# --- PESTAÑA 2: UNIR (CON REORDENAMIENTO) ---
with tab_merge:
    st.header("Unir múltiples PDFs")
    
    # 1. Subida de archivos
    files_merge = st.file_uploader(
        "1. Sube los PDFs (selecciona varios)", 
        type="pdf", 
        accept_multiple_files=True, 
        key="merge"
    )
    
    if files_merge:
        # 2. Mapa de archivos para poder reordenarlos
        # Creamos un diccionario {NombreDelArchivo: ObjetoArchivo}
        file_map = {f.name: f for f in files_merge}
        
        st.markdown("---")
        st.subheader("2. Ordenar archivos")
        st.info("👇 Arrastra o selecciona en el orden que quieres que aparezcan en el PDF final.")
        
        # Selector múltiple para definir el orden
        archivos_seleccionados_nombres = st.multiselect(
            "Secuencia de unión:",
            options=file_map.keys(),
            default=file_map.keys() # Por defecto aparecen todos en el orden de subida
        )
        
        # 3. Botón de Acción
        if st.button("Unir PDFs en este orden", type="primary"):
            if not archivos_seleccionados_nombres:
                st.warning("La lista para unir está vacía.")
            else:
                # Reconstruimos la lista de archivos basada en el orden visual
                lista_final_ordenada = [file_map[name] for name in archivos_seleccionados_nombres]
                
                pdf_merged = procesar_union(lista_final_ordenada)
                
                if pdf_merged:
                    st.success("¡Archivos unidos correctamente!")
                    st.download_button(
                        label="⬇️ Descargar PDF Unido",
                        data=pdf_merged,
                        file_name="documento_unido.pdf",
                        mime="application/pdf"
                    )

# --- PESTAÑA 3: WORD ---
with tab_word:
    st.header("Convertir PDF a Word")
    st.warning("Nota: Funciona mejor con PDFs de texto (no escaneados).")
    
    file_word = st.file_uploader("Sube tu PDF", type="pdf", key="word")
    
    if file_word:
        if st.button("Convertir a Word", type="primary"):
            with st.spinner("Analizando y convirtiendo... (puede tardar un poco)"):
                docx_result = procesar_word(file_word)
            
            if docx_result:
                st.success("¡Conversión completada!")
                nombre_descarga = os.path.splitext(file_word.name)[0] + ".docx"
                st.download_button(
                    label="⬇️ Descargar Word (.docx)",
                    data=docx_result,
                    file_name=nombre_descarga,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

st.markdown("---")
st.caption("Desarrollado por Luis Páez")