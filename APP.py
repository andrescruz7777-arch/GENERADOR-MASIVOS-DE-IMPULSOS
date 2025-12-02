import streamlit as st
import pandas as pd
from docx import Document

# ------------------------
# Funciones auxiliares
# ------------------------

def leer_excel(archivo_subido):
    try:
        df = pd.read_excel(archivo_subido)
        # Creamos un ID interno por fila para usar más adelante
        if "ID_INTERNO" not in df.columns:
            df.insert(0, "ID_INTERNO", range(1, len(df) + 1))
        return df, None
    except Exception as e:
        return None, f"Error al leer el Excel: {e}"

def leer_word_como_parrafos(archivo_subido):
    try:
        doc = Document(archivo_subido)
        # Filtramos párrafos vacíos para que la vista previa sea más limpia
        parrafos = [p.text for p in doc.paragraphs if p.text.strip() != ""]
        return parrafos, None
    except Exception as e:
        return None, f"Error al leer el Word: {e}"

# ------------------------
# Configuración básica de la app
# ------------------------

st.set_page_config(
    page_title="Generador de documentos judiciales",
    layout="wide"
)

st.title("📄 Generador de documentos judiciales – Fase 1")
st.caption("Paso 1: Cargar base en Excel y plantilla en Word con previsualización básica.")

# Estado de sesión
if "df_base" not in st.session_state:
    st.session_state.df_base = None

if "parrafos_plantilla" not in st.session_state:
    st.session_state.parrafos_plantilla = None

col1, col2 = st.columns(2)

# ------------------------
# Panel izquierdo: cargar Excel
# ------------------------
with col1:
    st.subheader("① Cargar base de datos (Excel)")

    archivo_excel = st.file_uploader(
        "Sube la base en Excel (.xlsx):",
        type=["xlsx", "xls"],
        key="uploader_excel"
    )

    if archivo_excel is not None:
        df, error = leer_excel(archivo_excel)
        if error:
            st.error(error)
        else:
            st.session_state.df_base = df
            st.success(f"Base cargada correctamente. Registros: {len(df)}")

            st.markdown("**Vista previa de las primeras filas:**")
            st.dataframe(df.head(10))

            st.markdown("**Columnas detectadas en la base:**")
            st.write(list(df.columns))

# ------------------------
# Panel derecho: cargar Word
# ------------------------
with col2:
    st.subheader("② Cargar plantilla (Word)")

    archivo_word = st.file_uploader(
        "Sube la plantilla en Word (.docx):",
        type=["docx"],
        key="uploader_word"
    )

    if archivo_word is not None:
        parrafos, error = leer_word_como_parrafos(archivo_word)
        if error:
            st.error(error)
        else:
            st.session_state.parrafos_plantilla = parrafos
            st.success("Plantilla Word cargada correctamente.")

            st.markdown("**Vista previa de los párrafos (solo texto):**")
            for i, p in enumerate(parrafos, start=1):
                st.markdown(f"**{i}.** {p}")

# ------------------------
# Estado general
# ------------------------
st.markdown("---")
st.subheader("Estado general")

if st.session_state.df_base is not None:
    st.success("✅ Base de datos cargada.")
else:
    st.warning("⚠️ Aún no has cargado la base de datos (Excel).")

if st.session_state.parrafos_plantilla is not None:
    st.success("✅ Plantilla Word cargada.")
else:
    st.warning("⚠️ Aún no has cargado la plantilla (Word).")

st.info("En el siguiente paso vamos a empezar el marcado de campos (JUZGADO, DEMANDANTE, etc.) sobre esta plantilla.")
