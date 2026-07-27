import io
import re
from datetime import datetime
import docx
from google import genai
import streamlit as st

# ==========================================
# CONFIGURACIÓN DE GEMINI API (SDK MODERNO)
# ==========================================
client = None
try:
    gemini_api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=gemini_api_key)
except Exception:
    st.warning(
        "⚠️ No se detectó 'GEMINI_API_KEY' en st.secrets. Agrega la clave para poder usar la IA."
    )

# ==========================================
# FUNCIONES HELPER - PARTE 1 (DOCX & TXT)
# ==========================================
def reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos):
    for buscar, reemplazo in mapa_reemplazos.items():
        if buscar in parrafo.text:
            parrafo.text = parrafo.text.replace(buscar, reemplazo)


def reemplazar_manteniendo_formato(doc, mapa_reemplazos):
    """Reemplaza los marcadores en párrafos y tablas del documento Word."""
    for parrafo in doc.paragraphs:
        reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos)


def obtener_valor_exacto(doc, etiqueta_buscada):
    """Busca el valor a la derecha o abajo de una etiqueta dada en las tablas."""
    etiqueta_clean = etiqueta_buscada.strip().lower()
    for tabla in doc.tables:
        for i_fila, fila in enumerate(tabla.rows):
            for i_celda, celda in enumerate(fila.cells):
                texto_celda = celda.text.strip().lower()
                if texto_celda == etiqueta_clean:
                    if i_celda + 1 < len(fila.cells):
                        val_derecha = fila.cells[i_celda + 1].text.strip()
                        if val_derecha and val_derecha.lower() != texto_celda:
                            return val_derecha
                    if i_fila + 1 < len(tabla.rows):
                        val_abajo = tabla.rows[i_fila + 1].cells[i_celda].text.strip()
                        if val_abajo:
                            return val_abajo
    return ""


def obtener_tarea_plan(doc, nombre_plan):
    """Obtiene el texto de los planes de ejecución o reversión en las tablas."""
    nombre_plan_clean = nombre_plan.strip().lower()
    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas_texto = []
            for celda in fila.cells:
                txt = celda.text.strip()
                if not celdas_texto or celdas_texto[-1] != txt:
                    celdas_texto.append(txt)
            for idx, txt in enumerate(celdas_texto):
                if nombre_plan_clean in txt.lower():
                    for posterior in celdas_texto[idx + 1 :]:
                        if posterior and posterior.lower() != txt.lower():
                            return posterior
    return ""


# ==========================================
# FUNCIÓN IA GEMINI - PARTE 2 (SQL RUC)
# ==========================================
def generar_queries_sql_con_gemini(
    texto_sunat,
    tipo_operacion="INSERT",
    ticket="123456",
    estado_previo="INACTIVO",
):
    prompt_sistema = f"""
    Eres un DBA experto en SQL Server para sistemas peruanos.
    Tu trabajo es procesar texto copiado de consultas RUC de SUNAT y generar exclusivamente scripts SQL.

    ESTRUCTURA DE LA TABLA `sunat_contribuyente`:
    - numero_ruc (varchar)
    - razon_social (varchar)
    - estado (varchar) -> Solo el estado (ej: 'ACTIVO', 'BAJA DE OFICIO')
    - condicion_domicilio (varchar) -> Solo la condición (ej: 'HABIDO', 'NO HABIDO')
    - ubigeo (varchar de 6 dígitos) -> CONVIERTE la cadena "DPTO - PROV - DIST" al código oficial de Ubigeo INE de 6 dígitos (ej: '140101', '150110', '140105').
    - tipo_via (varchar) -> ej: 'CAL.', 'AV.', 'JR.' (si aplica)
    - nombre_via (varchar) -> Nombre de la vía (si aplica)
    - numero (varchar) -> Número de puerta (si aplica)
    - interior (varchar) -> (si aplica)
    - lote (varchar) -> (si aplica)
    - departamento (varchar) -> (si aplica)
    - manzana (varchar) -> (si aplica)
    - kilometro (varchar) -> (si aplica)
    - tipo_zona (varchar) -> ej: 'URB.', 'A.H.' (si aplica)
    - codigo_zona (varchar) -> Nombre de la zona (si aplica)
    - estado_fila (varchar) -> 'I' si es INSERT, 'U' si es UPDATE.

    REGLAS DE OPERACIÓN:
    1. Responde ÚNICAMENTE con dos bloques separados por el delimitador "===ROLLBACK_SEPARADOR===":
       - Primero la consulta de PASE A PROD.
       - Luego el delimitador.
       - Finalmente la consulta de ROLLBACK.
    2. El encabezado del PASE debe iniciar con `--{ticket}`.
    3. La base de datos es `labroe_new`. Usa la sintaxis:
       use labroe_new
       go
    4. Omite los campos que sean NULL o no existan en la información dada.
    5. NO incluyas bloques de código Markdown (```sql), devuelve solo texto plano.
    6. Para ROLLBACK de UPDATE e INSERT: usa `update sunat_contribuyente set estado = '{estado_previo}', fecha_actualizacion = getdate() where numero_ruc = '...';`
    """

    prompt_usuario = f"""
    TIPO OPERACIÓN: {tipo_operacion}
    TEXTO SUNAT:
    {texto_sunat}
    """

    # ✅ Modelo oficial activo
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt_usuario,
        config={"system_instruction": prompt_sistema},
    )

    partes = response.text.split("===ROLLBACK_SEPARADOR===")
    q_prod = partes[0].strip() if len(partes) > 0 else response.text.strip()
    q_roll = partes[1].strip() if len(partes) > 1 else ""

    q_prod = re.sub(r"^```sql\n?|^```\n?", "", q_prod, flags=re.MULTILINE)
    q_prod = re.sub(r"\n?```$", "", q_prod, flags=re.MULTILINE)

    q_roll = re.sub(r"^```sql\n?|^```\n?", "", q_roll, flags=re.MULTILINE)
    q_roll = re.sub(r"\n?```$", "", q_roll, flags=re.MULTILINE)

    return q_prod, q_roll


# ==========================================
# INTERFAZ STREAMLIT (UNIFICADA)
# ==========================================
import zipfile

st.set_page_config(page_title="Gestor RUC & RFC", page_icon="📄", layout="wide")
st.title("📄 Sistema de Gestión Unificado RUC / RFC")

# Inicializar Session State para evitar reinicios al hacer descargas
if "resultado_procesado" not in st.session_state:
  st.session_state.resultado_procesado = None

# --- ENTRADA DE DATOS ---
st.header("1. Carga de Archivo y Datos")

col_left, col_right = st.columns(2)

with col_left:
  uploaded_file = st.file_uploader(
      "Sube la Solicitud de Cambio (.docx)", type=["docx"]
  )
  ticket_num = st.text_input("Ingresa el N° de Ticket:", placeholder="Ej: 12776")

with col_right:
  texto_ruc = st.text_area(
      "Pega la Consulta RUC copiada de SUNAT:", height=180
  )
  estado_previo = st.text_input(
      "Estado anterior para Rollback:",
      value="INACTIVO",
  )

# --- DETECCIÓN AUTOMÁTICA DEL TIPO DE OPERACIÓN ---
tipo_operacion = "INSERT"
if uploaded_file is not None:
  nombre_archivo = uploaded_file.name.lower()
  if "actualizar" in nombre_archivo:
    tipo_operacion = "UPDATE"
  elif "registrar" in nombre_archivo or "registro" in nombre_archivo:
    tipo_operacion = "INSERT"

  st.info(
      f"📌 **Tipo de Operación detectado automáticamente:** `{tipo_operacion}`"
  )

# --- BOTÓN DE PROCESAMIENTO ---
if st.button("🚀 Procesar Todo y Generar Archivos"):
  if not uploaded_file or not ticket_num or not texto_ruc:
    st.error("Por favor completa todos los campos requeridos.")
  elif not client:
    st.error("No se ha configurado 'GEMINI_API_KEY' en los secretos.")
  else:
    with st.spinner("Procesando documento Word y generando queries SQL..."):
      try:
        # 1. Procesar Word y TXT
        doc = docx.Document(uploaded_file)
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        reemplazos = {"[FECHA DE HOY]": fecha_actual, "TICKETNUM": ticket_num}
        reemplazar_manteniendo_formato(doc, reemplazos)

        detalle_cambio = obtener_valor_exacto(doc, "Descripción")
        plan_ejecucion = obtener_tarea_plan(doc, "ejecución")
        plan_reversion = obtener_tarea_plan(doc, "PLAN DE REVERSIÓN")
        descripcion_cambio = obtener_valor_exacto(doc, "MOTIVO DEL CAMBIO")
        analista_responsable = obtener_valor_exacto(doc, "Nombre")

        contenido_txt = f"""INFORMACION ADICIONAL

Detalle del Cambio/Despliegue:
-----------------------------------------------------------------
{detalle_cambio}

PLAN DE EJECUCIÓN
{plan_ejecucion}

PLAN DE REVERSIÓN (Roll-back)
{plan_reversion}

Descripción del cambio:
-----------------------------------------------------------------
{descripcion_cambio}

Analista/Especialista responsable del despliegue del cambio:
-----------------------------------------------------------------
{analista_responsable}

¿Existe Riesgo?:
-----------------------------------------------------------------
NINGUNO
"""
        buffer_docx = io.BytesIO()
        doc.save(buffer_docx)
        buffer_docx.seek(0)

        # 2. Generar SQL mediante Gemini AI
        q_prod, q_rollback = generar_queries_sql_con_gemini(
            texto_ruc, tipo_operacion, ticket_num, estado_previo
        )

        # Nombres dinámicos de archivos
        nombre_docx_out = f"{uploaded_file.name.replace('.docx', '')}_Procesado_{ticket_num}.docx"
        if tipo_operacion == "INSERT":
          nom_prod = f"query_RUC_registro - PASE A PROD ({ticket_num}).sql"
          nom_roll = f"query_RUC_registro - ROLLBACK ({ticket_num}).sql"
        else:
          nom_prod = f"query_RUC_actualizar - PASE A PROD ({ticket_num}).sql"
          nom_roll = f"query_RUC_actualizar - ROLLBACK ({ticket_num}).sql"

        # 3. Crear paquete ZIP con TODOS los archivos
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(
            zip_buffer, "w", zipfile.ZIP_DEFLATED
        ) as zip_file:
          zip_file.writestr(nombre_docx_out, buffer_docx.getvalue())
          zip_file.writestr("informacion_adicional.txt", contenido_txt)
          zip_file.writestr(nom_prod, q_prod)
          zip_file.writestr(nom_roll, q_rollback)

        zip_buffer.seek(0)

        # Guardar resultados en Session State
        st.session_state.resultado_procesado = {
            "q_prod": q_prod,
            "q_rollback": q_rollback,
            "zip_data": zip_buffer.getvalue(),
            "zip_name": f"Paquete_RFC_RUC_{ticket_num}.zip",
        }

      except Exception as e:
        st.error(f"Error durante el procesamiento: {e}")

# --- VISTA PREVIA Y DESCARGA EN BLOQUE ---
if st.session_state.resultado_procesado:
  st.divider()
  st.header("2. Vista Previa de Scripts SQL Generados")

  res = st.session_state.resultado_procesado

  col_p1, col_p2 = st.columns(2)
  with col_p1:
    st.subheader("Pase a Producción (PASE)")
    st.code(res["q_prod"], language="sql")

  with col_p2:
    st.subheader("Rollback (ROLLBACK)")
    st.code(res["q_rollback"], language="sql")

  st.divider()
  st.header("3. Confirmación y Descarga Unificada")

  st.download_button(
      label="📦 DESCARGAR TODOS LOS ARCHIVOS (.ZIP)",
      data=res["zip_data"],
      file_name=res["zip_name"],
      mime="application/zip",
      type="primary",
  )