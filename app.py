import streamlit as st
import pandas as pd
from datetime import datetime
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

st.set_page_config(page_title="Inventario DDUMA - Oficios de Resguardo", layout="wide")

st.title("📋 Generador de Oficios de Ubicación de Bienes - DDUMA")
st.write("Sube tu archivo Excel de inventario general para filtrar automáticamente los bienes y generar el oficio formal.")

# 1. Cargar archivo Excel
uploaded_file = st.file_uploader("📂 Sube aquí tu archivo Excel de Inventario", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

        # Identificar la columna del resguardante
        col_resguardante = None
        for col in df_raw.columns:
            if "resguardante actual" in col.lower() or "resguardante_actual" in col.lower():
                col_resguardante = col
                break
        
        if col_resguardante is None:
            for col in df_raw.columns:
                if "actual" in col.lower() or "resguardante" in col.lower():
                    col_resguardante = col
                    break

        if col_resguardante is not None:
            # 2. Configuración de Destinatario
            st.subheader("Configuración del Destinatario y Oficio")
            
            col1, col2 = st.columns(2)

            with col1:
                num_oficio = st.text_input("Número de Oficio", value="0542/DDUMA/2026")
                fecha_oficio = st.date_input("Fecha del Oficio", value=datetime.now())
                remitente_nombre = st.text_input("Nombre Remitente (Titular DDUMA)", value="LIC. ROSENDO SÁNCHEZ PREVE")
                remitente_cargo = st.text_input("Cargo Remitente", value="DIRECTORA DE DESARROLLO URBANO Y MEDIO AMBIENTE")

            with col2:
                nombre_direccion_dest = st.text_input("Nombre de la Dirección Destinataria", value="CATASTRO")
                destinatario_nombre = st.text_input("Nombre del Director / Titular Destinatario", value="LIC. JOSÉ DOMINGO [APELLIDOS]")

            # Formatear el texto sin los dos puntos
            limpio_dest = nombre_direccion_dest.upper().replace("DIRECCIÓN DE", "").replace("DIRECCION DE", "").strip()
            area_asignada_texto = f"DIRECCION DE {limpio_dest}"

            # Filtrar y estandarizar datos
            resguardante_clean = df_raw[col_resguardante].astype(str).str.upper().str.strip()
            
            # Filtro flexible para capturar variantes ("CATASTRO", "DIRECCION DE CATASTRO", etc.)
            palabra_clave = limpio_dest
            filtro = resguardante_clean.str.contains(palabra_clave, na=False)
            df_filtrado = df_raw[filtro].copy()

            if len(df_filtrado) > 0:
                # Reemplazar el valor para que en la tabla diga exactamente "DIRECCION DE CATASTRO" (sin dos puntos)
                df_filtrado[col_resguardante] = area_asignada_texto
                
                st.success(f"✅ Se encontraron {len(df_filtrado)} bienes asignados a {area_asignada_texto}.")

                # Mostrar tabla con la corrección
                st.subheader("Bienes Filtrados para el Oficio")
                df_bienes_final = st.data_editor(df_filtrado, use_container_width=True)

                # 3. Función para Generar Documento Word Formal
                def generar_word_formal(df_data):
                    doc = docx.Document()

                    # Márgenes de la página
                    for section in doc.sections:
                        section.top_margin = Inches(0.8)
                        section.bottom_margin = Inches(0.8)
                        section.left_margin = Inches(0.8)
                        section.right_margin = Inches(0.8)

                    # Lema del año
                    p_lema = doc.add_paragraph('"2026, Año de Margarita Maza Parada"')
                    p_lema.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_lema.runs[0].font.italic = True
                    p_lema.runs[0].font.size = Pt(8.5)

                    # Encabezado (Metadatos)
                    p_meta = doc.add_paragraph()
                    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p_meta.add_run("DIRECCIÓN DE DESARROLLO URBANO Y MEDIO AMBIENTE\n").bold = True
                    p_meta.add_run(f"OFICIO: {num_oficio}\n").bold = True
                    p_meta.add_run(f"ASUNTO: Notificación y solicitud de regularización de resguardos de bienes muebles localizados físicamente en la {area_asignada_texto}.\n").bold = True

                    # Fecha
                    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                    fecha_fmt = f"{fecha_oficio.day} de {meses[fecha_oficio.month - 1]} del {fecha_oficio.year}"
                    p_fecha = doc.add_paragraph(f"San Francisco de Campeche, Camp., a {fecha_fmt}\n")

                    # Destinatario
                    p_dest = doc.add_paragraph()
                    p_dest.add_run(f"{destinatario_nombre}\n").bold = True
                    p_dest.add_run(f"DIRECTOR DE {limpio_dest}\n").bold = True
                    p_dest.add_run("P R E S E N T E .-")
                    p_dest.paragraph_format.space_after = Pt(12)

                    # Cuerpo formal
                    parrafos = [
                        "Me dirijo a usted de la manera más atenta con relación a las acciones de verificación, conciliación y depuración del inventario físico de los bienes muebles asignados a esta Dirección de Desarrollo Urbano y Medio Ambiente.",
                        f"Sobre el particular, hago de su conocimiento que, derivado de las inspecciones físicas realizadas por el personal a mi cargo, se identificó la presencia física de diversos bienes muebles patrimoniales ubicados en las áreas operativas e instalaciones de la {area_asignada_texto} a su digno cargo.",
                        "En virtud de lo anterior, y con el propósito de dar debido cumplimiento a los lineamientos normativos vigentes en materia de administración de bienes muebles patrimoniales del H. Ayuntamiento de Campeche, le solicito atentamente girar sus apreciables instrucciones a quien corresponda, a efecto de proceder con la formalización del cambio de resguardo e integración formal en el Catálogo General de Bienes Muebles del Municipio.",
                        "A continuación, se detalla la relación de los bienes muebles antes referidos:"
                    ]

                    for p_text in parrafos:
                        p = doc.add_paragraph(p_text)
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p.paragraph_format.space_after = Pt(6)
                        p.paragraph_format.line_spacing = 1.15

                    # Tabla de Bienes en Word
                    table = doc.add_table(rows=1, cols=len(df_data.columns))
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    # Encabezados
                    hdr_cells = table.rows[0].cells
                    for col_idx, col_name in enumerate(df_data.columns):
                        hdr_cells[col_idx].text = str(col_name).upper()
                        hdr_cells[col_idx].paragraphs[0].runs[0].font.bold = True
                        hdr_cells[col_idx].paragraphs[0].runs[0].font.size = Pt(8)

                    # Filas
                    for idx, row in df_data.iterrows():
                        row_cells = table.add_row().cells
                        for col_idx, col_name in enumerate(df_data.columns):
                            val = str(row[col_name]) if pd.notna(row[col_name]) else ""
                            row_cells[col_idx].text = val
                            row_cells[col_idx].paragraphs[0].runs[0].font.size = Pt(8)

                    # Cierre y Firma
                    p_cierre = doc.add_paragraph("\nSin otro particular por el momento, le reitero la seguridad de mi atenta y distinguida consideración.\n")
                    p_cierre.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    
                    p_atentamente = doc.add_paragraph("A T E N T A M E N T E\n\n\n\n")
                    p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_atentamente.add_run(f"{remitente_nombre}\n").bold = True
                    p_atentamente.add_run(f"{remitente_cargo}").bold = True

                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    return buffer

                # Botón de Descarga
                st.download_button(
                    label="📄 Descargar Oficio Formal (.docx)",
                    data=generar_word_formal(df_bienes_final),
                    file_name=f"Oficio_DDUMA_{limpio_dest.replace(' ', '_')}_{num_oficio.replace('/', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning(f"⚠️ No se encontraron registros para la dirección o palabra clave '{palabra_clave}'.")
        else:
            st.error("❌ No se encontró la columna de resguardante en el archivo Excel.")
            
    except Exception as e:
        st.error(f"Error al procesar el archivo Excel: {e}")
else:
    st.info("👆 Por favor sube tu archivo Excel de inventario para comenzar.")
