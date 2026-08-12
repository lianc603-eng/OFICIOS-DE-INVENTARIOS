import streamlit as st
import pandas as pd
from datetime import datetime
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

st.set_page_config(page_title="Inventario DDUMA - Oficios a Catastro", layout="wide")

st.title("📋 Generador de Oficios de Ubicación de Bienes - DDUMA a Catastro")
st.write("Sube tu archivo Excel de inventario general para filtrar automáticamente los bienes ubicados en Catastro y generar el oficio formal.")

# 1. Cargar archivo Excel
uploaded_file = st.file_uploader("📂 Sube aquí tu archivo Excel de Inventario", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leer el Excel
    df_raw = pd.read_excel(uploaded_file)
    
    # Normalizar texto para la búsqueda (quitar espacios de más y convertir a mayúsculas)
    # Ajusta 'Resguardante' o el nombre exacto de la columna en tu Excel si varía
    col_resguardante = None
    for col in df_raw.columns:
        if "resguardante" in col.lower() or "area" in col.lower() or "ubicacion" in col.lower():
            col_resguardante = col
            break

    if col_resguardante is None:
        col_resguardante = df_raw.columns[2] # Toma la 3ra columna por defecto si no la detecta por nombre

    # Limpieza: Unificar 'DIRECCIÓN DE CATASTRO', 'CATASTRO', 'DIRECCION DE CATASTRO', etc.
    resguardante_clean = df_raw[col_resguardante].astype(str).str.upper().str.strip()
    
    # Filtro: Detectar si dice 'CATASTRO' o 'DIRECCIÓN DE CATASTRO'
    filtro_catastro = resguardante_clean.str.contains("CATASTRO", na=False)
    df_filtrado = df_raw[filtro_catastro].copy()

    st.success(f"✅ Se encontraron {len(df_filtrado)} bienes asignados o ubicados en Catastro.")

    # 2. Formulario de Datos para el Oficio
    st.subheader("Configuración del Oficio")
    col1, col2 = st.columns(2)

    with col1:
        num_oficio = st.text_input("Número de Oficio", value="0542/DDUMA/2026")
        fecha_oficio = st.date_input("Fecha del Oficio", value=datetime.now())
        destinatario_nombre = st.text_input("Director de Catastro", value="LIC. JOSÉ DOMINGO [APELLIDOS]")

    with col2:
        remitente_nombre = st.text_input("Remitente (Titular DDUMA)", value="LIC. ROSENDO SÁNCHEZ PREVE")
        remitente_cargo = st.text_input("Cargo Remitente", value="DIRECTORA DE DESARROLLO URBANO Y MEDIO AMBIENTE")

    # Mostrar tabla editable para revisión previa
    st.subheader("Bienes Filtrados para el Oficio")
    df_bienes_final = st.data_editor(df_filtrado, use_container_width=True)

    # 3. Función para Generar Documento Word
    def generar_word_formal(df_data):
        doc = docx.Document()

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

        # Encabezado
        p_meta = doc.add_paragraph()
        p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_meta.add_run("DIRECCIÓN DE DESARROLLO URBANO Y MEDIO AMBIENTE\n").bold = True
        p_meta.add_run(f"OFICIO: {num_oficio}\n").bold = True
        p_meta.add_run("ASUNTO: Notificación y solicitud de regularización de resguardos de bienes muebles localizados físicamente en la Dirección de Catastro.\n").bold = True

        # Fecha
        p_fecha = doc.add_paragraph(f"San Francisco de Campeche, Camp., a {fecha_oficio.strftime('%d de %B del %Y')}\n")

        # Destinatario
        p_dest = doc.add_paragraph()
        p_dest.add_run(f"{destinatario_nombre}\n").bold = True
        p_dest.add_run("DIRECTOR DE CATASTRO\n").bold = True
        p_dest.add_run("P R E S E N T E .-")
        p_dest.paragraph_format.space_after = Pt(12)

        # Cuerpo
        parrafos = [
            "Me dirijo a usted de la manera más atenta con relación a las acciones de verificación, conciliación y depuración del inventario físico de los bienes muebles asignados a esta Dirección de Desarrollo Urbano y Medio Ambiente.",
            "Sobre el particular, hago de su conocimiento que, derivado de las inspecciones físicas realizadas por el personal a mi cargo, se identificó la presencia física de diversos bienes muebles patrimoniales ubicados en las áreas operativas e instalaciones de la Dirección de Catastro a su digno cargo.",
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
        
        # Encabezados de tabla
        hdr_cells = table.rows[0].cells
        for col_idx, col_name in enumerate(df_data.columns):
            hdr_cells[col_idx].text = str(col_name).upper()
            hdr_cells[col_idx].paragraphs[0].runs[0].font.bold = True
            hdr_cells[col_idx].paragraphs[0].runs[0].font.size = Pt(8)

        # Filas de datos
        for idx, row in df_data.iterrows():
            row_cells = table.add_row().cells
            for col_idx, col_name in enumerate(df_data.columns):
                row_cells[col_idx].text = str(row[col_name]) if pd.notna(row[col_name]) else ""

        # Cierre y Firma
        doc.add_paragraph("\nSin otro particular por el momento, le reitero la seguridad de mi atenta y distinguida consideración.\n")
        
        p_atentamente = doc.add_paragraph("A T E N T A M E N T E\n\n\n\n")
        p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_atentamente.add_run(f"{remitente_nombre}\n").bold = True
        p_atentamente.add_run(f"{remitente_cargo}").bold = True

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    # Botón para descargar el Word
    st.download_button(
        label="📄 Descargar Oficio Formal (.docx)",
        data=generar_word_formal(df_bienes_final),
        file_name=f"Oficio_DDUMA_Catastro_{num_oficio.replace('/', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.info("👆 Por favor sube tu archivo Excel de inventario para comenzar.")
