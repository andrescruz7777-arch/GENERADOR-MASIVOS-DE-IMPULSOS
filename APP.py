import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
import re
from docxtpl import DocxTemplate
import zipfile

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

st.title("📄 Generador de documentos judiciales")
st.caption("Fase inicial: combinar base en Excel + plantilla Word usando placeholders {{...}}.")

# Estado de sesión
if "df_base" not in st.session_state:
    st.session_state.df_base = None

if "parrafos_plantilla" not in st.session_state:
    st.session_state.parrafos_plantilla = None

if "plantilla_bytes" not in st.session_state:
    st.session_state.plantilla_bytes = None

if "mapeo_placeholders" not in st.session_state:
    st.session_state.mapeo_placeholders = {}

if "resultados_docx" not in st.session_state:
    st.session_state.resultados_docx = None

if "regla_nombre_archivo" not in st.session_state:
    st.session_state.regla_nombre_archivo = "Memorial_{{RADICADO}}_{{DEMANDADO}}.docx"

col1, col2 = st.columns(2)

# ------------------------
# PASO 1: Cargar Excel
# ------------------------
with col1:
    st.subheader("① Cargar base de datos (Excel)")

    archivo_excel = st.file_uploader(
        "Sube la base en Excel (.xlsx o .xls):",
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
# PASO 2: Cargar Word
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
            # Guardamos los bytes de la plantilla para usarlos al generar los .docx
            st.session_state.plantilla_bytes = archivo_word.getvalue()

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
    st.success("✅ Base de datos (Excel) cargada.")
else:
    st.warning("⚠️ Aún no has cargado la base de datos (Excel).")

if st.session_state.parrafos_plantilla is not None:
    st.success("✅ Plantilla Word cargada.")
else:
    st.warning("⚠️ Aún no has cargado la plantilla (Word).")

# ------------------------
# PASO 3: Marcado de campos ({{...}}) + previsualización
# ------------------------

st.markdown("---")
st.header("③ Marcado de campos en la plantilla y previsualización")

# Verificamos que ambos estén cargados
if st.session_state.df_base is None or st.session_state.parrafos_plantilla is None:
    st.warning("Carga primero la base de datos (Excel) y la plantilla (Word).")
    st.stop()

df = st.session_state.df_base
parrafos = st.session_state.parrafos_plantilla

# Diccionario global: nombre_placeholder -> columna_excel
if "mapeo_placeholders" not in st.session_state:
    st.session_state.mapeo_placeholders = {}

st.write(
    "Detectamos variables dentro del texto con el formato {{NOMBRE}}. "
    "Aquí puedes vincular cada variable a una columna de la base."
)

# Para también poder avisar si hay placeholders sin mapear
placeholders_detectados_global = set()

# ---- Marcado de variables por párrafo ----
for idx, p in enumerate(parrafos):
    # Buscar placeholders tipo {{ NOMBRE }} dentro del párrafo
    placeholders = re.findall(r"{{\s*([^}]+?)\s*}}", p)

    if not placeholders:
        continue

    placeholders_unicos = sorted(set(placeholders))
    placeholders_detectados_global.update(placeholders_unicos)

    with st.expander(f"Párrafo {idx+1}"):
        st.markdown("### Contenido del párrafo:")
        st.write(p)

        st.markdown("### Variables detectadas en este párrafo:")

        for ph in placeholders_unicos:
            # Valor actual si ya habíamos mapeado esta variable antes
            valor_actual = st.session_state.mapeo_placeholders.get(ph, "(No vincular)")

            opciones = ["(No vincular)"] + list(df.columns)

            # Determinar índice por defecto del selectbox
            if valor_actual in df.columns:
                index_default = opciones.index(valor_actual)
            else:
                index_default = 0

            col_select = st.selectbox(
                f"Vincular la variable '{{{{{ph}}}}}' a una columna de la base:",
                options=opciones,
                index=index_default,
                key=f"ph_{idx}_{ph}"
            )

            # Actualizar mapeo global
            if col_select != "(No vincular)":
                st.session_state.mapeo_placeholders[ph] = col_select
            else:
                # Si el usuario elige "No vincular", la quitamos del diccionario (si existía)
                if ph in st.session_state.mapeo_placeholders:
                    del st.session_state.mapeo_placeholders[ph]

st.markdown("### 📝 Resumen de variables vinculadas")

if st.session_state.mapeo_placeholders:
    st.write(st.session_state.mapeo_placeholders)
else:
    st.info("Aún no has vinculado ninguna variable {{...}} a columnas de la base.")

# Aviso de variables detectadas pero no mapeadas
no_mapeadas = placeholders_detectados_global.difference(st.session_state.mapeo_placeholders.keys())
if no_mapeadas:
    st.warning(
        f"Estas variables fueron detectadas en la plantilla pero no están vinculadas a ninguna columna: "
        f"{', '.join(sorted(no_mapeadas))}"
    )

# ------------------------
# Previsualización con una fila de ejemplo
# ------------------------
st.markdown("---")
st.subheader("👁️ Previsualización del documento con una fila de la base")

if not st.session_state.mapeo_placeholders:
    st.info("Primero vincula al menos una variable {{...}} a alguna columna para poder previsualizar.")
else:
    # Selector de fila de ejemplo
    total_filas = len(df)
    fila_idx = st.number_input(
        "Selecciona el número de fila de la base para usar como ejemplo (1 = primera fila):",
        min_value=1,
        max_value=total_filas,
        value=1,
        step=1
    )

    fila = df.iloc[fila_idx - 1]

    st.caption(f"Mostrando previsualización usando la fila {fila_idx} de {total_filas}.")

    # Generar previsualización de cada párrafo
    st.markdown("### Resultado previsualizado:")

    for i, p in enumerate(parrafos, start=1):
        texto = p

        # Reemplazar cada placeholder mapeado por el valor correspondiente de la fila
        for ph, col in st.session_state.mapeo_placeholders.items():
            if col in df.columns:
                valor = fila[col]
                if pd.isna(valor):
                    valor = ""
                # Reemplazamos cualquier variante {{ NOMBRE }} / {{NOMBRE}} / {{   NOMBRE   }}
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                texto = re.sub(patron, str(valor), texto)

        st.markdown(f"**Párrafo {i}:**")
        st.write(texto)

# ------------------------
# PASO 4: Nombre de archivo + generación de .docx
# ------------------------

st.markdown("---")
st.header("④ Nombre de archivo y generación de documentos (.docx)")

# Verificaciones previas
if st.session_state.df_base is None or st.session_state.parrafos_plantilla is None:
    st.warning("Carga primero la base de datos y la plantilla.")
    st.stop()

if st.session_state.plantilla_bytes is None:
    st.warning("No se encontró la plantilla original en memoria. Vuelve a cargar el archivo Word.")
    st.stop()

if "mapeo_placeholders" not in st.session_state or not st.session_state.mapeo_placeholders:
    st.warning("Primero debes vincular las variables {{...}} a columnas de la base en el paso ③.")
    st.stop()

df = st.session_state.df_base
mapeo = st.session_state.mapeo_placeholders

st.write("Con la configuración actual se generará un documento por cada fila de la base de datos.")

# ------------------------
# Regla de nombres de archivo personalizada
# ------------------------
st.markdown("### 🏷️ Nombre de archivo personalizado")

st.caption(
    "Define cómo se debe llamar cada documento generado. "
    "Puedes usar texto y variables de la base entre llaves, ejemplo:\n"
    "**Memorial_{{RADICADO}}_{{DEMANDADO}}.docx**"
)

regla_nombre = st.text_input(
    "Escribe la regla del nombre del archivo:",
    value=st.session_state.regla_nombre_archivo
)
st.session_state.regla_nombre_archivo = regla_nombre

# Detectar placeholders dentro del nombre
placeholders_nombre = re.findall(r"{{\s*([^}]+?)\s*}}", regla_nombre)

if placeholders_nombre:
    st.write("Variables detectadas en el nombre del archivo:")
    st.write(placeholders_nombre)
else:
    st.info("No se detectaron variables {{...}} en el nombre. Se usará el mismo nombre para todos los archivos (consecutivo).")

# ------------------------
# Botón para generar documentos
# ------------------------
if st.button("▶️ Generar documentos .docx"):
    plantilla_bytes = st.session_state.plantilla_bytes

    buffer_zip = BytesIO()
    zf = zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED)

    resultados = []
    progreso = st.progress(0)
    total = len(df)

    for idx, (_, fila) in enumerate(df.iterrows(), start=1):
        # Construimos el contexto para docxtpl: placeholder -> valor
        contexto = {}
        for ph, col in mapeo.items():
            if col in df.columns:
                valor = fila[col]
                if pd.isna(valor):
                    valor = ""
                contexto[ph] = str(valor)
            else:
                contexto[ph] = ""

        # Cargamos la plantilla desde memoria en cada iteración
        doc = DocxTemplate(BytesIO(plantilla_bytes))
        doc.render(contexto)

        # --- Construcción del nombre de archivo basado en la regla definida ---
        nombre_archivo = regla_nombre

        # Reemplazamos placeholders de la regla con los datos de la fila
        for ph, col in mapeo.items():
            if col in df.columns:
                valor = fila[col]
                if pd.isna(valor):
                    valor = ""
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                nombre_archivo = re.sub(patron, str(valor), nombre_archivo)

        # Si el usuario no incluyó .docx, lo agregamos
        if not nombre_archivo.lower().endswith(".docx"):
            nombre_archivo = nombre_archivo + ".docx"

        # Si después de reemplazar quedó vacío o solo .docx, usamos un fallback
        if nombre_archivo.strip() == ".docx":
            nombre_archivo = f"documento_{idx}.docx"

        # Guardamos en un buffer temporal
        doc_buffer = BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)

        # Añadimos al ZIP
        zf.writestr(nombre_archivo, doc_buffer.read())

        resultados.append({
            "fila": idx,
            "nombre_archivo": nombre_archivo
        })

        progreso.progress(idx / total)

    zf.close()
    buffer_zip.seek(0)

    st.success(f"Se generaron {len(resultados)} documentos .docx.")

    st.markdown("### Ejemplo de archivos generados:")
    st.dataframe(pd.DataFrame(resultados).head(10))

    st.download_button(
        label="⬇️ Descargar todos los documentos (.zip)",
        data=buffer_zip.getvalue(),
        file_name="documentos_generados.docx.zip",
        mime="application/zip"
    )

    # Guardamos el resumen en sesión por si lo necesitamos luego (para correos)
    st.session_state.resultados_docx = resultados
