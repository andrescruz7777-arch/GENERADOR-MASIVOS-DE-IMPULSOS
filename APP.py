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
# PASO 4: Nombre de archivo y generación de documentos (.docx)
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

    # Ver cuáles se pueden resolver (mapeo o columna directa) y cuáles no
    resolvibles = []
    no_resolvibles = []
    for ph in placeholders_nombre:
        if ph in mapeo or ph in df.columns:
            resolvibles.append(ph)
        else:
            no_resolvibles.append(ph)

    if resolvibles:
        st.success(f"Variables que se pueden resolver con la base: {', '.join(resolvibles)}")
    if no_resolvibles:
        st.warning(
            "Estas variables del nombre no tienen mapeo ni columna con el mismo nombre, "
            "y por ahora quedarán vacías en el nombre: "
            + ", ".join(no_resolvibles)
        )
else:
    st.info(
        "No se detectaron variables {{...}} en el nombre. "
        "Se usará el mismo nombre para todos los archivos con un consecutivo."
    )

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

        # 1) Reemplazamos usando las variables del NOMBRE (placeholders_nombre)
        for ph in placeholders_nombre:
            # Decidimos de dónde sacar el valor de esta variable de nombre
            if ph in mapeo and mapeo[ph] in df.columns:
                col = mapeo[ph]
            elif ph in df.columns:
                col = ph
            else:
                col = None

            if col is not None:
                valor = fila[col]
                if pd.isna(valor):
                    valor = ""
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                nombre_archivo = re.sub(patron, str(valor), nombre_archivo)
            else:
                # Si no hay columna/mapeo para esta variable del nombre, la dejamos vacía
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                nombre_archivo = re.sub(patron, "", nombre_archivo)

        # 2) Si el usuario no incluyó .docx, lo agregamos
        if not nombre_archivo.lower().endswith(".docx"):
            nombre_archivo = nombre_archivo + ".docx"

        # 3) Si después de reemplazar quedó vacío o solo .docx, usamos un fallback
        if nombre_archivo.strip() == ".docx" or nombre_archivo.strip() == "":
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
# ------------------------
# PASO 5: Cargar PDFs generados externamente y mapearlos
# ------------------------
st.markdown("---")
st.header("⑤ Cargar PDFs convertidos y asociarlos a cada registro")

if "resultados_pdf" not in st.session_state:
    st.session_state.resultados_pdf = None

# Verificaciones básicas
if st.session_state.df_base is None:
    st.warning("Carga primero la base de datos (Excel).")
    st.stop()

df = st.session_state.df_base
mapeo = st.session_state.mapeo_placeholders
plantilla_bytes = st.session_state.plantilla_bytes
regla_nombre = st.session_state.regla_nombre_archivo

st.write(
    "Después de generar los .docx y convertirlos a PDF por fuera, "
    "sube aquí un archivo .zip con todos los PDF. "
    "Los nombres de los PDF deben seguir la misma regla que usaste para los .docx, "
    "por ejemplo: **Memorial_{{RADICADO}}_{{DEMANDADO}}.pdf**"
)

zip_pdfs = st.file_uploader(
    "Sube el archivo ZIP con todos los PDF:",
    type=["zip"],
    key="uploader_zip_pdfs"
)

pdf_mapping = []

if zip_pdfs is not None:
    # Leemos el ZIP en memoria
    zip_bytes = zip_pdfs.read()
    zf_pdf = zipfile.ZipFile(BytesIO(zip_bytes), "r")

    nombres_archivos_zip = zf_pdf.namelist()
    st.write("Archivos detectados dentro del ZIP:")
    st.write(nombres_archivos_zip)

    # Detectamos placeholders del nombre (los mismos que en el paso 4)
    placeholders_nombre = re.findall(r"{{\s*([^}]+?)\s*}}", regla_nombre)

    # Recorremos cada fila del DF y calculamos el nombre de PDF esperado
    for idx, (_, fila) in enumerate(df.iterrows(), start=1):
        nombre_esperado = regla_nombre

        # Reemplazamos placeholders en el nombre usando mapeo o columna directa
        for ph in placeholders_nombre:
            if ph in mapeo and mapeo[ph] in df.columns:
                col = mapeo[ph]
            elif ph in df.columns:
                col = ph
            else:
                col = None

            if col is not None:
                valor = fila[col]
                if pd.isna(valor):
                    valor = ""
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                nombre_esperado = re.sub(patron, str(valor), nombre_esperado)
            else:
                patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
                nombre_esperado = re.sub(patron, "", nombre_esperado)

        # Ajustamos extensión a .pdf
        nombre_esperado = nombre_esperado.replace(".docx", "").replace(".DOCX", "")
        if not nombre_esperado.lower().endswith(".pdf"):
            nombre_esperado = nombre_esperado + ".pdf"

        # Buscamos si ese archivo existe en el ZIP (match exacto, case-sensitive simple)
        encontrado = nombre_esperado in nombres_archivos_zip

        pdf_mapping.append({
            "fila": idx,
            "nombre_esperado": nombre_esperado,
            "encontrado": encontrado
        })

    mapping_df = pd.DataFrame(pdf_mapping)

    st.markdown("### Resultado del mapeo PDFs ↔ base")
    st.dataframe(mapping_df)

    faltantes = mapping_df[~mapping_df["encontrado"]]
    if len(faltantes) > 0:
        st.warning(
            f"Hay {len(faltantes)} registros sin PDF encontrado. "
            "Verifica que los nombres en el ZIP coincidan con la regla."
        )
    else:
        st.success("Todos los registros tienen un PDF asociado en el ZIP.")

    # Guardamos en sesión: el zip original y el mapping
    st.session_state.pdf_zip_bytes = zip_bytes
    st.session_state.pdf_mapping = pdf_mapping
# ------------------------
# PASO 6: Configuración de correo y envío de prueba
# ------------------------
import smtplib
from email.message import EmailMessage

st.markdown("---")
st.header("⑥ Configuración de correo y envío de prueba")

if "pdf_zip_bytes" not in st.session_state or "pdf_mapping" not in st.session_state:
    st.info("Primero carga el ZIP con los PDFs en el paso ⑤ para poder adjuntarlos en correos.")
    st.stop()

df = st.session_state.df_base
pdf_zip_bytes = st.session_state.pdf_zip_bytes
pdf_mapping = st.session_state.pdf_mapping

# ------------------------
# Configuración SMTP
# ------------------------
st.subheader("Configuración SMTP (servidor de correo)")

smtp_host = st.text_input("Servidor SMTP (host):", value="smtp.gmail.com")
smtp_port = st.number_input("Puerto SMTP:", value=587, step=1)
smtp_user = st.text_input("Usuario / correo remitente:")
smtp_pass = st.text_input("Contraseña / app password:", type="password")
from_name = st.text_input("Nombre que verá el destinatario:", value="Área Judicial")
from_email = st.text_input("Correo FROM (si es distinto al usuario):", value="")

if from_email.strip() == "":
    from_email = smtp_user

st.caption("Usa un correo con permisos SMTP (por ejemplo, Gmail con app password o un buzón corporativo).")

# ------------------------
# Plantilla de correo (asunto y cuerpo)
# ------------------------
st.subheader("Plantilla de correo (uso de variables {{...}})")

asunto_template = st.text_input(
    "Asunto del correo (puedes usar {{RADICADO}}, {{DEMANDADO}}, etc.):",
    value="Memorial proceso {{RADICADO}} contra {{DEMANDADO}}"
)

cuerpo_template = st.text_area(
    "Cuerpo del correo:",
    value=(
        "Señor(a) {{JUZGADO}},\n\n"
        "Adjunto remito el memorial correspondiente al proceso {{RADICADO}} "
        "seguido por {{DEMANDANTE}} contra {{DEMANDADO}}.\n\n"
        "Cordialmente,\n"
        "{{ABOGADO}}\n"
        "{{TARJETA_PROFESIONAL}}"
    ),
    height=200
)

st.caption("Las variables funcionan igual que en la plantilla: se reemplazan con los datos de cada fila.")

# ------------------------
# Envío de prueba
# ------------------------
st.subheader("Envío de prueba (solo un registro)")

total_filas = len(df)
fila_prueba = st.number_input(
    "Fila de la base a usar para la prueba (1 = primera fila):",
    min_value=1,
    max_value=total_filas,
    value=1,
    step=1
)

correo_prueba = st.text_input("Enviar prueba a este correo:")

def construir_asunto_y_cuerpo(fila):
    texto_asunto = asunto_template
    texto_cuerpo = cuerpo_template

    # Reemplazamos usando columnas del DF (nombre de columna = placeholder)
    for col in df.columns:
        ph = col  # asumimos placeholder igual al nombre de la columna
        valor = fila[col]
        if pd.isna(valor):
            valor = ""
        patron = r"{{\s*" + re.escape(ph) + r"\s*}}"
        texto_asunto = re.sub(patron, str(valor), texto_asunto)
        texto_cuerpo = re.sub(patron, str(valor), texto_cuerpo)

    return texto_asunto, texto_cuerpo

def obtener_pdf_para_fila(idx_fila):
    # idx_fila viene 1-based, mapping también
    esperado = None
    encontrado = False
    nombre_zip = None

    for item in pdf_mapping:
        if item["fila"] == idx_fila:
            esperado = item["nombre_esperado"]
            encontrado = item["encontrado"]
            break

    if not encontrado or esperado is None:
        return None, None, False

    # Abrimos el ZIP y extraemos ese archivo
    zf_pdf = zipfile.ZipFile(BytesIO(pdf_zip_bytes), "r")
    try:
        data = zf_pdf.read(esperado)
        return esperado, data, True
    except KeyError:
        return esperado, None, False

if st.button("📨 Enviar correo de prueba"):
    if not smtp_host or not smtp_user or not smtp_pass:
        st.error("Configura primero SMTP (host, usuario y contraseña).")
    elif not correo_prueba:
        st.error("Indica un correo de destino para la prueba.")
    else:
        fila = df.iloc[fila_prueba - 1]
        asunto_final, cuerpo_final = construir_asunto_y_cuerpo(fila)

        nombre_pdf, pdf_bytes, ok_pdf = obtener_pdf_para_fila(fila_prueba)

        if not ok_pdf:
            st.error(f"No se encontró PDF para la fila {fila_prueba}. Nombre esperado: {nombre_pdf}")
        else:
            try:
                msg = EmailMessage()
                msg["Subject"] = asunto_final
                msg["From"] = f"{from_name} <{from_email}>"
                msg["To"] = correo_prueba
                msg.set_content(cuerpo_final)

                # Adjuntamos PDF
                msg.add_attachment(
                    pdf_bytes,
                    maintype="application",
                    subtype="pdf",
                    filename=nombre_pdf
                )

                with smtplib.SMTP(smtp_host, smtp_port) as server:
                    server.starttls()
                    server.login(smtp_user, smtp_pass)
                    server.send_message(msg)

                st.success(f"Correo de prueba enviado a {correo_prueba} con adjunto {nombre_pdf}.")
                st.code(f"Asunto: {asunto_final}\n\n{cuerpo_final}", language="text")

            except Exception as e:
                st.error(f"Error al enviar el correo de prueba: {e}")

