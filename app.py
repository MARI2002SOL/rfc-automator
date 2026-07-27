import io
import re
from datetime import datetime
import docx
from google import genai
import streamlit as st


client = None
try:
  gemini_api_key = st.secrets["GEMINI_API_KEY"]
  client = genai.Client(api_key=gemini_api_key)
except Exception:
  st.warning(
      "⚠️ No se detectó 'GEMINI_API_KEY' en st.secrets. Agrega la clave para"
      " poder usar la IA."
  )


# ==========================================
# FUNCIONES HELPER - PARTE 1 (DOCX & TXT)
# ==========================================
def reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos):
  for buscar, reemplazo in mapa_reemplazos.items():
    if buscar in parrafo.text:
      parrafo.text = parrafo.text.replace(buscar, reemplazo)


def reemplazar_manteniendo_formato(doc, mapa_reemplazos):
  for parrafo in doc.paragraphs:
    reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos)
  for tabla in doc.tables:
    for fila in tabla.rows:
      for celda in fila.cells:
        for parrafo in celda.paragraphs:
          reemplazar_texto_en_parrafo(parrafo, mapa_reemplazos)


def obtener_valor_exacto(doc, etiqueta_buscada):
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
    texto_sunat, tipo_operacion="INSERT", ticket="123456", estado_previo="BAJA DE OFICIO"
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

    INSTRUCCIONES DE FORMATO:
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
    6. Para ROLLBACK de INSERT: usa `delete from sunat_contribuyente where numero_ruc = '...';`
    7. Para ROLLBACK de UPDATE: usa `update sunat_contribuyente set estado = '{estado_previo}', fecha_actualizacion = getdate() where numero_ruc = '...';`
    """

    prompt_usuario = f"""
    TIPO OPERACIÓN: {tipo_operacion}
    TEXTO SUNAT:
    {texto_sunat}
    """

    # Llamada a Gemini usando el modelo activo
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt_usuario,
        config={"system_instruction": prompt_sistema},
    )

    # Validar que la respuesta contenga texto antes de desempaquetar
    if not response or not hasattr(response, "text") or not response.text:
        raise ValueError("La API de Gemini devolvió una respuesta vacía o fue bloqueada por filtros.")

    texto_respuesta = response.text

    # Manejo seguro de la separación
    if "===ROLLBACK_SEPARADOR===" in texto_respuesta:
        partes = texto_respuesta.split("===ROLLBACK_SEPARADOR===")
        q_prod = partes[0].strip()
        q_roll = partes[1].strip() if len(partes) > 1 else ""
    else:
        q_prod = texto_respuesta.strip()
        q_roll = "-- No se generó query de rollback"

    # Limpieza de sintaxis markdown si la IA incluyó ```sql
    q_prod = re.sub(r"^```sql\n?|^```\n?", "", q_prod, flags=re.MULTILINE)
    q_prod = re.sub(r"\n?```$", "", q_prod, flags=re.MULTILINE)

    q_roll = re.sub(r"^```sql\n?|^```\n?", "", q_roll, flags=re.MULTILINE)
    q_roll = re.sub(r"\n?```$", "", q_roll, flags=re.MULTILINE)

    return q_prod, q_roll
  # Llamada a Gemini 2.5 Flash usando el cliente oficial
# Puedes ejecutar esto en tu terminal o agregar un st.write temporal en Streamlit:
if client:
  try:
    modelos = [m.name for m in client.models.list()]
    print("Modelos disponibles:", modelos)
  except Exception as e:
    print("Error al listar modelos:", e)

# ==========================================
# INTERFAZ STREAMLIT
# ==========================================
st.set_page_config(page_title="Gestor RUC & RFC", page_icon="📄", layout="wide")
st.title("📄 Sistema de Gestión RUC / RFC (Powered by Gemini AI)")

tab1, tab2 = st.tabs(["1. Completar Documento RFC", "2. Generador SQL RUC"])

# --- TAB 1: COMPLETAR DOCUMENTO RFC ---
with tab1:
  st.header("Completar Plantilla Word RFC")

  uploaded_file = st.file_uploader(
      "Sube la Solicitud de Cambio (.docx)", type=["docx"]
  )
  ticket_num = st.text_input("Ingresa el N° de Ticket:", placeholder="Ej: INC123456")

  if uploaded_file and ticket_num:
    if st.button("Procesar y Generar Archivos"):
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

      st.success("¡Documento procesado con éxito!")

      col1, col2 = st.columns(2)
      with col1:
        st.download_button(
            label="📥 Descargar Word Modificado",
            data=buffer_docx,
            file_name=f"RFC_Procesado_{ticket_num}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

      with col2:
        st.download_button(
            label="📥 Descargar TXT Información Adicional",
            data=contenido_txt,
            file_name="informacion_adicional.txt",
            mime="text/plain",
        )

# --- TAB 2: GENERADOR SQL RUC CON GEMINI ---
with tab2:
  st.header("Generador de Script SQL RUC con IA Gemini")

  texto_ruc = st.text_area(
      "Pega la Consulta RUC copiada de SUNAT:", height=180
  )
  col_a, col_b, col_c = st.columns(3)

  with col_a:
    tipo_op = st.radio("Tipo de Operación:", ["INSERT", "UPDATE"])
  with col_b:
    ticket_sql = st.text_input(
        "N° de Ticket para el script SQL:", value="INC 12776"
    )
  with col_c:
    estado_prev = "BAJA DE OFICIO"
    if tipo_op == "UPDATE":
      estado_prev = st.text_input(
          "Estado anterior (Solo Rollback Update):", value="BAJA DE OFICIO"
      )

  if st.button("🤖 Generar Queries con Gemini AI"):
    if not client:
      st.error(
          "No se ha configurado 'GEMINI_API_KEY' en la aplicación."
      )
    elif not texto_ruc or not ticket_sql:
      st.error("Por favor ingresa la consulta RUC y el número de Ticket.")
    else:
      with st.spinner("Procesando con Gemini AI..."):
        try:
          q_prod, q_rollback = generar_queries_sql_con_gemini(
              texto_ruc, tipo_op, ticket_sql, estado_prev
          )

          st.subheader("Pase a Producción (PASE)")
          st.code(q_prod, language="sql")

          st.subheader("Rollback (ROLLBACK)")
          st.code(q_rollback, language="sql")

          if tipo_op == "INSERT":
            nom_prod = f"query_RUC_registro - PASE A PROD ({ticket_sql}).sql"
            nom_roll = f"query_RUC_registro - ROLLBACK ({ticket_sql}).sql"
          else:
            nom_prod = f"query_RUC_actualizar - PASE A PROD ({ticket_sql}).sql"
            nom_roll = f"query_RUC_actualizar - ROLLBACK ({ticket_sql}).sql"

          col1, col2 = st.columns(2)
          with col1:
            st.download_button(
                label=f"📥 Descargar {nom_prod}",
                data=q_prod,
                file_name=nom_prod,
                mime="text/plain",
            )
          with col2:
            st.download_button(
                label=f"📥 Descargar {nom_roll}",
                data=q_rollback,
                file_name=nom_roll,
                mime="text/plain",
            )
        except Exception as e:
          st.error(f"Error al conectar con la API de Gemini: {e}")