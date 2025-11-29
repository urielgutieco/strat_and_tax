import os
import io
import zipfile
import random
import json
import base64
import logging
import tempfile
import pathlib
import re
import bcrypt
from flask import Flask, request, send_file, jsonify, abort
from flask_cors import CORS
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches

# Google Drive
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# -------------------------
# Config / Producción-safe
# -------------------------
app = Flask(__name__)

# Seguridad y límites
MAX_CONTENT_LENGTH_BYTES = int(os.getenv("MAX_CONTENT_LENGTH_BYTES", 20 * 1024 * 1024))  # 20 MB por default
ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg"}
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH_BYTES
app.config["JSON_SORT_KEYS"] = False

# CORS: permitir orígenes configurados (coma-separados) o ninguno (más seguro)
allowed_origins = os.getenv("ALLOWED_ORIGINS", "")
if allowed_origins:
    origins = [o.strip() for o in allowed_origins.split(",") if o.strip()]
    CORS(app, origins=origins)
else:
    # En producción conviene especificar ALLOWED_ORIGINS; si no existe, no habilitamos CORS globalmente
    CORS(app, resources={r"/generate-word": {"origins": []}})

# Logging en modo producción
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("generate-word-app")

# Aquí pones la contraseña que LE VAS A DAR a la usuaria (Ej: "ContraseñaProd123")
password_plano = b"u12345G" 
# Nota: La 'b' convierte la cadena a bytes, lo que requiere bcrypt.

# Generar el hash seguro
hashed_password = bcrypt.hashpw(password_plano, bcrypt.gensalt())

# Imprimir el hash para copiarlo
print(hashed_password.decode('utf-8'))
# Ubicación: Dentro de api/index.py

# Base de datos de usuarios con HASHES de contraseñas
USER_DB = {
    "nombre_administrador_seguro": b"$2b$12$tUa5Z8rF.E.q2H/i.q5U7.G.R6A1W9V4P3I0Y2X5Q8T7S6R5C4V", 
    
    # 🎯 Aquí pegas el hash generado en el paso 2
    "nombre_de_la_usuaria_prod": b"$2b$12$hGq9p0aQ4w7xS2zV.B.c.8A.D.E.F3G4H5I6J7K8L9M0N1O2P3Q4", 
}

# Ubicación: En api/index.py, después de la sección de Configuración de Seguridad.

@app.route('/login', methods=['POST'])
def login():
    """
    Verifica las credenciales utilizando bcrypt para comparar la contraseña ingresada
    con el hash almacenado en USER_DB. Si son válidas, emite un token JWT.
    """
    try:
        data = request.get_json()
        username = data.get('username')
        password = data.get('password')

        if not username or not password:
            return jsonify({"error": "Falta usuario o contraseña"}), 400

        # Obtener el hash almacenado del usuario (ya debe estar en formato bytes: b'...')
        stored_hash = USER_DB.get(username) 

        # 1. Verificar credenciales con bcrypt.checkpw
        # bcrypt.checkpw requiere que tanto la contraseña ingresada como el hash 
        # almacenado sean bytes (password.encode('utf-8') y stored_hash, respectivamente).
        if stored_hash and bcrypt.checkpw(password.encode('utf-8'), stored_hash):
            
            # 2. Generar Token JWT (Autenticación exitosa)
            expiration_time = datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRATION_HOURS)
            payload = {
                'sub': username, 
                'exp': expiration_time,
                'iat': datetime.now(timezone.utc)
            }
            token = jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)

            return jsonify({
                "message": "Login exitoso",
                "token": token
            }), 200
        else:
            # Fallo si el usuario no existe O el hash no coincide
            return jsonify({"error": "Usuario o contraseña inválidos"}), 401

    except Exception as e:
        logger.error("Error inesperado en login: %s", e)
        return jsonify({"error": "Error interno del servidor durante el login"}), 500


# Carpetas raíz (no usadas en serverless pero se mantienen por compatibilidad)
BASE_DIR = pathlib.Path(os.getenv("BASE_DIR", "."))

TEMPLATE_FOLDER_NAME = os.getenv("TEMPLATE_FOLDER_NAME", "template_word") # Asumiendo que "template_word" es el default
# BASE_DIR en Vercel es donde se encuentra el código.
TEMPLATE_FOLDER = pathlib.Path(__file__).parent / TEMPLATE_FOLDER_NAME
# Si usas el valor por defecto, las plantillas deben estar en api/template_word/
GENERATED_DOCS = BASE_DIR / 'template_gendocs'
GENERATED_ZIPS = BASE_DIR / 'template_genzips'

# Crear en caso de ejecución local; Vercel serverless no mantiene persistencia entre invocaciones.
TEMPLATE_FOLDER.mkdir(parents=True, exist_ok=True)
GENERATED_DOCS.mkdir(parents=True, exist_ok=True)
GENERATED_ZIPS.mkdir(parents=True, exist_ok=True)

# Mapeo servicio → carpeta (mismo que tenías)
SERVICIO_TO_DIR = {
    "Servicios de construccion de unidades unifamiliares": "construccion_unifamiliar",
    "Servicios de reparacion o ampliacion o remodelacion de viviendas unifamiliares": "reparacion_remodelacion_unifamiliar",
    "Servicio de remodelacion general de viviendas unifamiliares": "remodelacion_general",
    "Servicios de reparacion de casas moviles en el sitio": "reparacion_casas_moviles",
    "Servicios de construccion y reparacion de patios y terrazas": "patios_terrazas",
    "Servico de reparacion por daños ocasionados por fuego de viviendas unifamiliares": "reparacion_por_fuego",
    "Servicio de construccion de casas unifamiliares nuevas": "construccion_unifamiliar_nueva",
    "Servicio de instalacion de casas unifamiliares prefabricadas": "instalacion_prefabricadas",
    "Servicio de construccion de casas en la ciudad o casas jardin unifamiliares nuevas": "construccion_casas_ciudad_jardin",
    "Dasarrollo urbano": "desarrollo_urbano",
    "Servicio de planificacion de la ordenacion urbana": "planificacion_ordenacion_urbana",
    "Servicio de administracion de tierras urbanas": "administracion_tierras_urbanas",
    "Servicio de programacion de inversiones urbanas": "programacion_inversiones_urbanas",
    "Servicio de reestructuracion de barrios marginales": "reestructuracion_barrios_marginales",
    "Servicios de alumbrado urbano": "alumbrado_urbano",
    "Servicios de control o regulacion del desarrollo urbano": "control_desarrollo_urbano",
    "Servicios de estandares o regulacion de edificios urbanos": "estandares_regulacion_edificios",
    "Servicios comunitarios urbanos": "comunitarios_urbanos",
    "Servicios de administracion o gestion de proyectos o programas urbanos": "gestion_proyectos_programas_urbanos",
    "Ingenieria civil": "ingenieria_civil",
    "Ingenieria de carreteras": "ingenieria_carreteras",
    "Ingenieria deinfraestructura de instalaciones o fabricas": "infraestructura_instalaciones_fabricas",
    "Servicios de mantenimiento e instalacion de equipo pesado": "mantenimiento_instalacion_equipo_pesado",
    "Servicio de mantenimiento y reparacion de equipo pesado": "mantenimiento_reparacion_equipo_pesado",
}

# Plantillas (conservadas)
TEMPLATE_FILES = [
    'plantilla_solicitud.docx',
    '2.docx',
    '3.docx',
    '4.docx',
    '1.docx',
]

# -------------------------
# Utilidades
# -------------------------
def sanitize_filename(name: str) -> str:
    """Devuelve nombre de archivo seguro (sin caracteres raros)."""
    name = secure_filename(name)
    # quitar secuencias dobles de subrayado y limitar longitud
    name = re.sub(r'[_]{2,}', '_', name)
    return name[:200]

def replace_text_in_document(document, replacements):
    for paragraph in document.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, str(value))

def generate_single_document(template_filename, template_root, replacements, user_image_path=None, data=None):
    template_path = os.path.join(template_root, template_filename)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Plantilla '{template_filename}' no encontrada en '{template_root}'.")

    document = Document(template_path)
    replace_text_in_document(document, replacements)

    # Imagen
    if user_image_path and os.path.exists(user_image_path):
        try:
            document.add_paragraph()
            document.add_paragraph(
                data.get('nombre_completo_de_la_persona_que_firma_la_solicitud', 'N/A')
                if data else 'N/A'
            )
            document.add_picture(user_image_path, width=Inches(2.5))
        except Exception as ex:
            # Guardamos la excepción en log pero seguimos
            logger.warning("No se pudo insertar la imagen del usuario: %s", ex)
            document.add_paragraph("⚠ No se pudo insertar la imagen del usuario.")
    else:
        document.add_paragraph("⚠ Imagen de firma no encontrada en el servidor.")

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------
# Google Drive (Service Account) - seguro
# ------------------------
def _load_service_account_info():
    """
    Carga la info del service account desde:
    1) GOOGLE_SERVICE_ACCOUNT_JSON (contenido JSON) o
    2) GOOGLE_SERVICE_ACCOUNT_BASE64 (base64 del JSON)
    """
    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw_json:
        try:
            return json.loads(raw_json)
        except Exception as e:
            logger.error("GOOGLE_SERVICE_ACCOUNT_JSON no es JSON válido: %s", e)
            raise

    b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_BASE64")
    if b64:
        try:
            decoded = base64.b64decode(b64)
            return json.loads(decoded)
        except Exception as e:
            logger.error("GOOGLE_SERVICE_ACCOUNT_BASE64 inválido: %s", e)
            raise

    raise EnvironmentError("Falta GOOGLE_SERVICE_ACCOUNT_JSON o GOOGLE_SERVICE_ACCOUNT_BASE64")

def authenticate_and_upload_to_drive(file_name, zip_buffer):
    """
    Autenticación con Google Drive usando Service Account.
    Si la variable DISABLE_DRIVE_UPLOAD=1 está establecida, no sube (útil en pruebas).
    """
    if os.getenv("DISABLE_DRIVE_UPLOAD", "0") == "1":
        logger.info("Subida a Drive deshabilitada por variable DISABLE_DRIVE_UPLOAD.")
        return {"success": True, "message": "Subida a Drive deshabilitada (env var)"}

    try:
        service_account_info = _load_service_account_info()

        creds = Credentials.from_service_account_info(
            service_account_info,
            scopes=["https://www.googleapis.com/auth/drive.file"]
        )

        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        file_metadata = {
            'name': sanitize_filename(file_name)
        }

        zip_buffer.seek(0)
        media = MediaIoBaseUpload(zip_buffer, mimetype='application/zip', resumable=True)

        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        return {
            "success": True,
            "message": f"Archivo subido correctamente. ID: {uploaded_file.get('id')}"
        }

    except Exception as e:
        logger.exception("Error al subir a Drive: %s", e)
        return {"success": False, "message": f"Error al subir a Drive: {str(e)}"}
    

# ------------------------
# Endpoint principal (misma ruta y flujo)
# ------------------------
@app.route('/generate-word', methods=['POST'])
def generate_word():
    try:
        # form y archivos
        data = request.form.to_dict()
        uploaded_image = request.files.get("imagen_usuario")
        user_image_path = None

        # Validar datos mínimos
        if not data:
            return jsonify({"error": "No data received"}), 400

        servicio = data.get('servicio')
        if not servicio:
            return jsonify({"error": "Debes seleccionar un servicio."}), 400

        carpeta_servicio = SERVICIO_TO_DIR.get(servicio)
        if not carpeta_servicio:
            return jsonify({"error": f"No existe carpeta mapeada para el servicio: {servicio}"}), 404

        template_root = os.path.join(TEMPLATE_FOLDER, carpeta_servicio)
        if not os.path.isdir(template_root):
            return jsonify({"error": f"La carpeta de plantillas no existe: {template_root}"}), 404

        # manejar imagen subida (temporal)
        if uploaded_image and uploaded_image.filename:
            filename = sanitize_filename(uploaded_image.filename)
            ext = pathlib.Path(filename).suffix.lower()
            if ext not in ALLOWED_IMAGE_EXT:
                return jsonify({"error": "Tipo de archivo no permitido para imagen."}), 400

            # Guardar temporalmente
            tmp_dir = tempfile.TemporaryDirectory()
            tmp_path = pathlib.Path(tmp_dir.name) / filename
            uploaded_image.save(str(tmp_path))
            user_image_path = str(tmp_path)
            # nota: tmp_dir se eliminará cuando el objeto tmp_dir sea recolectado; mantenemos referencia en scope.

        # limpiar docs previos en carpeta (si existe) — seguí tu lógica
        try:
            for filename in os.listdir(GENERATED_DOCS):
                file_path = os.path.join(GENERATED_DOCS, filename)
                if os.path.isfile(file_path):
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        logger.warning("Error al eliminar %s: %s", file_path, e)
        except FileNotFoundError:
            # ok en serverless
            pass

        numero_de_contrato_unico = ''.join([str(random.randint(0, 9)) for _ in range(18)])
        descripcion_servicio = servicio

        replacements = {
            '${descripcion_del_servicio}': descripcion_servicio,
            '${razon_social}': data.get('razon_social', 'N/A'),
            '${r_f_c}': data.get('r_f_c', 'N/A'),
            '${domicilio_del_cliente}': data.get('domicilio_del_cliente', 'N/A'),
            '${telefono_del__cliente}': data.get('telefono_del__cliente', 'N/A'),
            '${correo_del_cliente}': data.get('correo_del_cliente', 'N/A'),
            '${fecha_de_inicio_del_servicio}': data.get('fecha_de_inicio_del_servicio', 'N/A'),
            '${fecha_de_conclusion_del_servicio}': data.get('fecha_de_conclusion_del_servicio', 'N/A'),
            '${monto_de_la_operacion_Sin_IVA}': data.get('monto_de_la_operacion_Sin_IVA', 'N/A'),
            '${forma_de_pago}': data.get('forma_de_pago', 'N/A'),
            '${cantidad}': data.get('cantidad', 'N/A'),
            '${unidad}': data.get('unidad', 'N/A'),
            '${numero_de_contrato}': numero_de_contrato_unico,
            '${fecha_de_operación}': data.get('fecha_de_operacion', 'N/A'),
            '${nombre_completo_de_la_persona_que_firma_la_solicitud}': data.get('nombre_completo_de_la_persona_que_firma_la_solicitud', 'N/A'),
            '${cargo_de_la_persona_que_firma_la_solicitud}': data.get('cargo_de_la_persona_que_firma_la_solicitud', 'N/A'),
            '${factura_relacionada_con_la_operación}': data.get('factura_relacionada_con_la_operación', 'N/A'),
            '${informe_si_cuenta_con_fotografias_videos_o_informacion_adicion}': data.get('informe_si_cuenta_con_fotografias_videos_o_informacion_adicion', 'N/A'),
            '${comentarios}': data.get('comentarios', 'N/A')
        }

        # Crear ZIP en memoria (BytesIO)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for template in TEMPLATE_FILES:
                try:
                    doc_buffer = generate_single_document(template, template_root, replacements, user_image_path, data)
                    base = os.path.splitext(template)[0]

                    rfc = data.get('r_f_c', 'N/A')
                    # filename seguro
                    output_filename = f"{sanitize_filename(base)}_{sanitize_filename(descripcion_servicio)}_{sanitize_filename(base)}_{numero_de_contrato_unico}_{sanitize_filename(rfc)}.docx"

                    # opcional: guardar en GENERATED_DOCS (si hay persistencia)
                    try:
                        output_path = GENERATED_DOCS / output_filename
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                        with open(output_path, "wb") as f:
                            f.write(doc_buffer.getvalue())
                    except Exception as e:
                        logger.debug("No se pudo guardar en disco (ok en serverless): %s", e)

                    zip_file.writestr(output_filename, doc_buffer.getvalue())

                except Exception as e:
                    logger.exception("Error generando documento %s: %s", template, e)
                    continue

        zip_buffer.seek(0)

        final_zip_name = f"{sanitize_filename(descripcion_servicio)}_{numero_de_contrato_unico}_{sanitize_filename(data.get('r_f_c', 'N/A'))}.zip"

        # Guardar zip temporalmente en servidor (opcional; en serverless es por poco tiempo)
        try:
            tmp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            tmp_zip.write(zip_buffer.getvalue())
            tmp_zip.flush()
            tmp_zip_name = tmp_zip.name
            tmp_zip.close()
            # mover a carpeta GENERARED_ZIPS si quieres persistencia local
            try:
                pathlib.Path(GENERATED_ZIPS).mkdir(parents=True, exist_ok=True)
                dest = pathlib.Path(GENERATED_ZIPS) / final_zip_name
                with open(tmp_zip_name, "rb") as s, open(dest, "wb") as d:
                    d.write(s.read())
            except Exception:
                # no crítico
                pass
        except Exception as e:
            logger.warning("No se pudo crear zip temporal en disco: %s", e)
            tmp_zip_name = None

        # ---------- SUBIR A DRIVE ----------
        zip_buffer.seek(0)
        upload_result = authenticate_and_upload_to_drive(final_zip_name, zip_buffer)
        logger.info("GOOGLE DRIVE: %s", upload_result.get("message"))

        # Preparar respuesta (descarga directa)
        zip_buffer.seek(0)
        response = send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=final_zip_name
        )

        # limpieza de temporales (si existen)
        try:
            if tmp_zip_name and os.path.exists(tmp_zip_name):
                os.remove(tmp_zip_name)
        except Exception:
            pass

        return response

    except Exception as e:
        logger.exception("Error interno en endpoint: %s", e)
        return jsonify({"error": f"Error interno: {str(e)}"}), 500


