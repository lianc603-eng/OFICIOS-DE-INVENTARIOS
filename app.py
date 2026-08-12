import streamlit as st
import pandas as pd
from datetime import datetime
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

st.set_page_config(page_title="Inventario DDUMA - Formatos de Verificación", layout="wide")

st.title("📋 Control de Inventario - Formatos de Verificación en Campo")
st.write("Sube tu archivo Excel para filtrar por cualquier área o dirección, configurar destinatarios y generar formatos de verificación.")

# 1. Cargar archivo Excel
uploaded_file = st.file_uploader("📂 Sube aquí tu archivo Excel de Inventario", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

        # Identificar la columna del resguardante actual
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
            # Configuración general
            st.subheader("⚙️ Configuración del Destinatario y Oficio")
            
            col_cfg1, col_cfg2, col_cfg3 = st.columns(3)

            with col_cfg1:
                nombre_direccion_dest = st.text_input("Dirección / Area Destinataria", value="CATASTRO")
                limpio_dest = nombre_direccion_dest.upper().replace("DIRECCIÓN DE", "").replace("DIRECCION DE", "").strip()
                area_asignada_texto = f"DIRECCION DE {limpio_dest}"

                tipo_cargo_dest = st.selectbox(
                    "Cargo del Destinatario",
                    ["DIRECTOR", "DIRECTORA", "SUBDIRECTOR", "SUBDIRECTORA", "JEFE DE DEPARTAMENTO", "JEFA DE DEPARTAMENTO", "TITULAR", "OTRO"]
                )
                if tipo_cargo_dest == "OTRO":
                    cargo_dest_final = st.text_input("Escribe el cargo del destinatario", value="RESPONSABLE DE ÁREA").upper()
                else:
                    cargo_dest_final = f"{tipo_cargo_dest} DE {limpio_dest}"

            with col_cfg2:
                destinatario_nombre = st.text_input("Nombre del Destinatario / Titular", value="LIC. JOSÉ DOMINGO [APELLIDOS]")
                num_oficio = st.text_input("Número de Oficio", value="0542/DDUMA/2026")

            with col_cfg3:
                remitente_nombre = st.text_input("Remitente (Titular DDUMA)", value="LIC. ROSENDO SÁNCHEZ PREVE")
                remitente_cargo = st.text_input("Cargo Remitente", value="DIRECTOR DE DESARROLLO URBANO Y MEDIO AMBIENTE")
                fecha_oficio = st.date_input("Fecha del Oficio", value=datetime.now())

            st.markdown("---")

            # Filtrar datos de la dependencia seleccionada dinámicamente
            resguardante_clean = df_raw[col_resguardante].astype(str).str.upper().str.strip()
            filtro = resguardante_clean.str.contains(limpio_dest, na=False)
            df_filtrado = df_raw[filtro].copy()

            if len(df_filtrado) > 0:
                # Estandarizar la columna para que diga "DIRECCION DE [ÁREA]"
                df_filtrado[col_resguardante] = area_asignada_texto
                
                # ELIMINAR LA COLUMNA DE RESGUARDANTE ANTERIOR SI EXISTE
                cols_to_drop = [c for c in df_filtrado.columns if "anterior" in c.lower()]
                if cols_to_drop:
                    df_filtrado = df_filtrado.drop(columns=cols_to_drop)

                # RENOMBRAR LA COLUMNA 'RESGUARDANTE ACTUAL' A 'DIRECCIÓN'
                df_filtrado = df_filtrado.rename(columns={col_resguardante: "DIRECCIÓN"})

                # COLUMNA EXCLUSIVA PARA ANOTAR A PLUMA EL COMPAÑERO / OBSERVACIONES
                df_filtrado_verificacion = df_filtrado.copy()
                df_filtrado_verificacion["COMPAÑERO QUE LO TIENE / OBSERVACIONES"] = "____________________"

                st.success(f"✅ Se encontraron {len(df_filtrado)} bienes en {area_asignada_texto}.")

                # Pestañas: Plantilla Imprimible Excel vs Oficio Word
                tab1, tab2 = st.tabs(["🖨️ Formato Imprimible de Verificación (Excel)", "📄 Oficio Formal de Verificación (Word)"])

                # TAB 1: FORMATO PARA IMPRIMIR Y ANOTAR A MANO
                with tab1:
                    st.write("### Vista Previa de la Cédula de Verificación")
                    st.write("Esta tabla incluye la columna **'DIRECCIÓN'** y el espacio **'COMPAÑERO QUE LO TIENE / OBSERVACIONES'** para tomar notas en papel.")
                    
                    st.dataframe(df_filtrado_verificacion, use_container_width=True)

                    def convert_df_to_excel_print(df):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Cedula_Verificacion')
                        output.seek(0)
                        return output

                    st.download_button(
                        label="🟢 Descargar Excel de Verificación en Campo",
                        data=convert_df_to_excel_print(df_filtrado_verificacion),
                        file_name=f"Cedula_Verificacion_{limpio_dest}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # TAB 2: GENERACIÓN DEL OFICIO WORD DE VERIFICACIÓN
                with tab2:
                    st.write("### Generación del Oficio Formal para Verificación de Bienes")

                    def generar_word_verificacion(df_data):
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

                        # Metadatos
                        p_meta = doc.add_paragraph()
                        p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p_meta.add_run("DIRECCIÓN DE DESARROLLO URBANO Y MEDIO AMBIENTE\n").bold = True
                        p_meta.add_run(f"OFICIO: {num_oficio}\n").bold = True
                        p_meta.add_run(f"ASUNTO: Solicitud de verificación e inspección física de bienes muebles en sus instalaciones.\n").bold = True

                        # Fecha
                        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                        fecha_fmt = f"{fecha_oficio.day} de {meses[fecha_oficio.month - 1]} del {fecha_oficio.year}"
                        p_fecha = doc.add_paragraph(f"San Francisco de Campeche, Camp., a {fecha_fmt}\n")

                        # Destinatario dinámico
                        p_dest = doc.add_paragraph()
                        p_dest.add_run(f"{destinatario_nombre.upper()}\n").bold = True
                        p_dest.add_run(f"{cargo_dest_final.upper()}\n").bold = True
                        p_dest.add_run("P R E S E N T E .-")
                        p_dest.paragraph_format.space_after = Pt(12)

                        # Cuerpo
                        parrafos = [
                            "Me dirijo a usted de la manera más atenta en el marco de las actividades de control, seguimiento y actualización del inventario patrimonial de este H. Ayuntamiento.",
                            f"Sobre el particular, le solicito atentamente su valioso apoyo a efecto de realizar la verificación y constatación física en campo de los bienes muebles que a continuación se relacionan, los cuales se tienen ubicados preliminarmente en las instalaciones y áreas operativas de la {area_asignada_texto} a su digno cargo.",
                            "Agradeceré se sirva confirmar la localización física de dichos bienes e identificar al compañero o servidor público que los tiene actualmente bajo su uso u operatividad directa, a fin de mantener actualizados los registros institucionales.",
                            "La relación de los bienes sujetos a verificación física se detalla a continuación:"
                        ]

                        for p_text in parrafos:
                            p = doc.add_paragraph(p_text)
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            p.paragraph_format.space_after = Pt(6)
                            p.paragraph_format.line_spacing = 1.15

                        # Tabla de Bienes en Word
                        table = doc.add_table(rows=1, cols=len(df_data.columns))
                        table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        
                        hdr_cells = table.rows[0].cells
                        for col_idx, col_name in enumerate(df_data.columns):
                            hdr_cells[col_idx].text = str(col_name).upper()
                            hdr_cells[col_idx].paragraphs[0].runs[0].font.bold = True
                            hdr_cells[col_idx].paragraphs[0].runs[0].font.size = Pt(8)

                        for idx, row in df_data.iterrows():
                            row_cells = table.add_row().cells
                            for col_idx, col_name in enumerate(df_data.columns):
                                val = str(row[col_name]) if pd.notna(row[col_name]) else ""
                                row_cells[col_idx].text = val
                                row_cells[col_idx].paragraphs[0].runs[0].font.size = Pt(8)

                        # Cierre y Firma del Remitente
                        p_cierre = doc.add_paragraph("\nSin otro particular por el momento, le reitero la seguridad de mi atenta y distinguida consideración.\n")
                        p_cierre.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        p_atentamente = doc.add_paragraph("A T E N T A M E N T E\n\n\n\n")
                        p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_atentamente.add_run(f"{remitente_nombre.upper()}\n").bold = True
                        p_atentamente.add_run(f"{remitente_cargo.upper()}").bold = True

                        buffer = io.BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)
                        return buffer

                    st.download_button(
                        label="📄 Descargar Oficio de Verificación (.docx)",
                        data=generar_word_verificacion(df_filtrado_verificacion),
                        file_name=f"Oficio_Verificacion_{limpio_dest}_{num_oficio.replace('/', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.warning(f"⚠️ No se encontraron bienes asignados a la palabra clave '{limpio_dest}'. Prueba escribiendo otra área o dirección.")
        else:
            st.error("❌ No se encontró la columna de resguardante en el archivo Excel.")
            
    except Exception as e:
        st.error(f"Error al procesar el archivo Excel: {e}")
else:
    st.info("👆 Por favor sube tu archivo Excel de inventario para comenzar.")
