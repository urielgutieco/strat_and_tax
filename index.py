# Código actualizado con los ajustes solicitados
# (Se respetó toda la estructura original. SOLO se corrigió la sección indicada.)

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
import jwt
import time
import boto3
from botocore.exceptions import ClientError
from mangum import Mangum  # si usas Mangum para Lambda
from googleapiclient.errors import HttpError
from flask import Response, make_response
from flask import Flask, request, send_file, jsonify, abort
from flask_cors import CORS
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import datetime, timedelta, timezone
from functools import wraps
from collections import defaultdict
from werkzeug.exceptions import RequestEntityTooLarge

# -------------------------
# Configuración Flask / Serverless
# -------------------------
app = Flask(__name__)

MAX_CONTENT_LENGTH_BYTES = int(os.getenv("MAX_CONTENT_LENGTH_BYTES", 20 * 1024 * 1024))
ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg"}
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH_BYTES
app.config["JSON_SORT_KEYS"] = False

TEMPLATE_FOLDER_NAME = os.getenv("TEMPLATE_FOLDER_NAME", "template_word")
TEMPLATE_FOLDER = pathlib.Path(__file__).parent / TEMPLATE_FOLDER_NAME

allowed_origins = os.getenv("ALLOWED_ORIGINS", "")
if allowed_origins:
    origins = [o.strip() for o in allowed_origins.split(",") if o.strip()]
    CORS(app, origins=origins)
else:
    CORS(app, resources={r"/generate-word": {"origins": []}})

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("generate-word-app")

# -------------------------
# JWT Config
# -------------------------
JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY", None)
if not JWT_SECRET_KEY:
    raise EnvironmentError("JWT_SECRET_KEY no configurada. Configúrala en AWS Lambda Environment Variables.")

JWT_ALGORITHM = "HS256"
JWT_EXPIRATION_HOURS = 1

# -------------------------
# Usuarios controlados
# -------------------------
USER_DB = {
    "usuario1": {
        "password_hash": b"$2b$12$hGq9p0aQ4w7xS2zV.B.c.8A.D.E.F3G4H5I6J7K8L9M0N1O2P3Q4",
        "role": "user"
    },
    "usuario2": {
        "password_hash": b"$2b$12$Xyz123Abc456Def789Ghi012Jkl345Mno678Pqr901Stu234Vwx",
        "role": "user"
    }
}

# -------------------------
# Rate limiting básico
# -------------------------
MAX_ATTEMPTS = 5
BLOCK_TIME_SECONDS = 300
login_attempts = defaultdict(lambda: {"count": 0, "last_attempt": 0, "blocked_until": 0})

def check_rate_limit(ip):
    entry = login_attempts[ip]
    now = time.time()
    if entry["blocked_until"] > now:
        return False, int(entry["blocked_until"] - now)
    return True, 0

def record_failed_attempt(ip):
    entry = login_attempts[ip]
    now = time.time()
    entry["count"] += 1
    entry["last_attempt"] = now
    if entry["count"] >= MAX_ATTEMPTS:
        entry["blocked_until"] = now + BLOCK_TIME_SECONDS

def reset_attempts(ip):
    login_attempts[ip] = {"count": 0, "last_attempt": 0, "blocked_until": 0}

# -------------------------
# Mapeos de Servicios
# -------------------------
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
    name = secure_filename(name)
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
    template_path = template_root / template_filename
    if not template_path.exists():
        raise FileNotFoundError(f"Plantilla '{template_filename}' no encontrada en '{template_root}'.")

    document = Document(template_path)
    replace_text_in_document(document, replacements)

    if user_image_path and os.path.exists(user_image_path):
        try:
            document.add_paragraph()
            document.add_paragraph(data.get('nombre_completo_de_la_persona_que_firma_la_solicitud', 'N/A') if data else 'N/A')
            document.add_picture(user_image_path, width=Inches(2.5))
        except Exception as ex:
            logger.warning("No se pudo insertar la imagen del usuario: %s", ex)
            document.add_paragraph("⚠ No se pudo insertar la imagen del usuario.")
    else:
        document.add_paragraph("⚠ Imagen de firma no encontrada en el servidor.")

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer


def get_secret_value(secret_name):
    if not secret_name:
        raise EnvironmentError("Secret name not provided for Secrets Manager")
    try:
        client = boto3.client("secretsmanager")
        resp = client.get_secret_value(SecretId=secret_name)
        if "SecretString" in resp and resp["SecretString"]:
            return resp["SecretString"]
        return resp["SecretBinary"]
    except ClientError as e:
        logger.exception("No se pudo obtener el secreto %s: %s", secret_name, e)
        raise

# -------------------------
# Google Drive Service Account
# -------------------------
def _load_service_account_info():
    sm_name = os.getenv("GOOGLE_SERVICE_ACCOUNT_SECRET_NAME")
    if sm_name:
        raw = get_secret_value(sm_name)
        return json.loads(raw)

    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw_json:
        return json.loads(raw_json)

    b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_BASE64")
    if b64:
        decoded = base64.b64decode(b64)
        return json.loads(decoded)

    raise EnvironmentError("Falta GOOGLE_SERVICE_ACCOUNT")


def authenticate_and_upload_to_drive(file_name, zip_buffer):
    if os.getenv("DISABLE_DRIVE_UPLOAD", "0") == "1":
        logger.info("Subida a Drive deshabilitada por variable DISABLE_DRIVE_UPLOAD.")
        return {"success": True, "message": "Subida deshabilitada"}

    try:
        service_account_info = _load_service_account_info()
    except Exception as e:
        logger.exception("No se pudo cargar info de service account: %s", e)
        return {"success": False, "message": str(e)}

    scopes = ["https://www.googleapis.com/auth/drive.file"]
    try:
        creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        file_metadata = {
            'name': sanitize_filename(file_name),
            'mimeType': 'application/zip'
        }

        drive_folder = os.getenv("DRIVE_FOLDER_ID")
        if drive_folder:
            file_metadata['parents'] = [drive_folder]

        zip_buffer.seek(0)
        media = MediaIoBaseUpload(zip_buffer, mimetype='application/zip', resumable=False)

        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                uploaded = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id'
                ).execute()
                return {"success": True, "message": f"Archivo subido. ID: {uploaded.get('id')}"}
            except HttpError as he:
                logger.warning("Error intento %d al subir a Drive: %s", attempt, he)
                if attempt == max_retries:
                    return {"success": False, "message": str(he)}
                time.sleep(2 ** attempt)
    except Exception as e:
        logger.exception("Error al autenticar o subir a Drive: %s", e)
        return {"success": False, "message": str(e)}


# -------------------------
# Decorador JWT
# -------------------------
def token_required(required_role='user'):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            token = None
            auth_header = request.headers.get('Authorization')
            if auth_header and auth_header.startswith('Bearer '):
                token = auth_header.split(' ')[1]

            try:
                data = jwt.decode(
                    token,
                    JWT_SECRET_KEY,
                    algorithms=[JWT_ALGORITHM],
                    audience='strat_and_tax_api',
                    issuer='strat_and_tax_server'
                )
                current_user = data['sub']
                role = data.get('role')

                if current_user not in USER_DB:
                    return jsonify({'error': 'Usuario no autorizado'}), 403
                if role != required_role:
                    return jsonify({'error': 'No tienes permisos'}), 403

            except jwt.ExpiredSignatureError:
                return jsonify({'error': 'Token expirado'}), 401
            except jwt.InvalidTokenError:
                return jsonify({'error': 'Token inválido'}), 403
            except Exception as e:
                logger.error("Error JWT: %s", e)
                return jsonify({'error': 'Error al validar token'}), 500

            return f(current_user, *args, **kwargs)
        return decorated
    return decorator

# -------------------------
# Login
# -------------------------
@app.route('/login', methods=['POST'])
def login():
    ip = request.remote_addr
    allowed, wait = check_rate_limit(ip)
    if not allowed:
        return jsonify({"error": f"Demasiados intentos, espera {wait}s"}), 429

    data = request.get_json()
    username = data.get('username')
    password = data.get('password')

    if not username or not password:
        return jsonify({"error": "Faltan datos"}), 400

    user = USER_DB.get(username)
    if user and bcrypt.checkpw(password.encode('utf-8'), user['password_hash']):
        reset_attempts(ip)
        issued = datetime.now(timezone.utc)
        exp = issued + timedelta(hours=JWT_EXPIRATION_HOURS)

        payload = {
            'sub': username,
            'role': user['role'],
            'iat': int(issued.timestamp()),
            'nbf': int(issued.timestamp()),
            'exp': int(exp.timestamp()),
            'aud': 'strat_and_tax_api',
            'iss': 'strat_and_tax_server'
        }

        token = jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)
        return jsonify({"message": "Login exitoso", "token": token}), 200

    record_failed_attempt(ip)
    return jsonify({"error": "Credenciales inválidas"}), 401

# -------------------------
# generate-word
# -------------------------
@app.route('/generate-word', methods=['POST'])
@token_required(required_role='user')
def generate_word(current_user):
    logger.info(f"Generación solicitada por: {current_user}")

    user_image_path = None

    try:
        data = request.form.to_dict()
        uploaded_image = request.files.get("imagen_usuario")

        if not data:
            return jsonify({"error": "No data"}), 400

        servicio = data.get('servicio')
        carpeta = SERVICIO_TO_DIR.get(servicio)
        if not carpeta:
            return jsonify({"error": f"Servicio desconocido: {servicio}"}), 404

        template_root = TEMPLATE_FOLDER / carpeta
        if not template_root.is_dir():
            return jsonify({"error": f"Carpeta no existe: {template_root}"}), 404

        # --- Manejo de imagen temporal ---
        if uploaded_image and uploaded_image.filename:
            filename = sanitize_filename(uploaded_image.filename)
            ext = pathlib.Path(filename).suffix.lower()
            if ext not in ALLOWED_IMAGE_EXT:
                return jsonify({"error": "Tipo de imagen no permitido"}), 400

            tmp_dir = pathlib.Path(tempfile.gettempdir()) / f"upload_{os.getpid()}"
            tmp_dir.mkdir(exist_ok=True)
            user_image_path = str(tmp_dir / filename)

            MAX_IMAGE_BYTES = int(os.getenv("MAX_IMAGE_BYTES", 5 * 1024 * 1024))
            uploaded_image.stream.seek(0, io.SEEK_END)
            size = uploaded_image.stream.tell()
            uploaded_image.stream.seek(0)
            if size > MAX_IMAGE_BYTES:
                return jsonify({"error": "Imagen demasiado grande"}), 413

            uploaded_image.save(user_image_path)

        numero = ''.join([str(random.randint(0, 9)) for _ in range(18)])

        replacements = {
            '${descripcion_del_servicio}': servicio,
            '${razon_social}': data.get('razon_social', 'N/A'),
            '${r_f_c}': data.get('r_f_c', 'N/A'),
            '${domicilio_del_cliente}': data.get('domicilio_del_cliente', 'N/A'),
            '${telefono_del_cliente}': data.get('telefono_del_cliente', 'N/A'),
            '${correo_del_cliente}': data.get('correo_del_cliente', 'N/A'),
            '${fecha_de_inicio_del_servicio}': data.get('fecha_de_inicio_del_servicio', 'N/A'),
            '${fecha_de_conclusion_del_servicio}': data.get('fecha_de_conclusion_del_servicio', 'N/A'),
            '${monto_de_la_operacion_Sin_IVA}': data.get('monto_de_la_operacion_in_iva', 'N/A'),
            '${forma_de_pago}': data.get('forma_de_pago', 'N/A'),
            '${cantidad}': data.get('cantidad', 'N/A'),
            '${unidad}': data.get('unidad', 'N/A'),
            '${numero_de_contrato}': numero,
            '${fecha_de_operación}': data.get('fecha_de_operacion', 'N/A'),
            '${nombre_completo_de_la_persona_que_firma_la_solicitud}': data.get('nombre_completo_de_la_persona_que_firma_la_solicitud', 'N/A'),
            '${cargo_de_la_persona_que_firma_la_solicitud}': data.get('cargo_de_la_persona_que_firma_la_solicitud', 'N/A'),
            '${factura_relacionada_con_la_operación}': data.get('factura_relacionada_con_la_operación', 'N/A'),
            '${informe_si_cuenta_con_fotografias_videos_o_informacion_adicion}': data.get('informe_si_cuenta_con_fotografias_videos_o_informacion_adicion', 'N/A'),
            '${comentarios}': data.get('comentarios', 'N/A')
        }

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for t in TEMPLATE_FILES:
                try:
                    doc_buf = generate_single_document(t, template_root, replacements, user_image_path, data)
                    base = os.path.splitext(t)[0]
                    rfc = data.get('r_f_c', 'N/A')
                    outname = f"{sanitize_filename(base)}_{sanitize_filename(servicio)}_{sanitize_filename(base)}_{numero}_{sanitize_filename(rfc)}.docx"
                    zipf.writestr(outname, doc_buf.getvalue())
                except Exception as e:
                    logger.exception("Error generando doc %s: %s", t, e)

        zip_buffer.seek(0)
        final_zip = f"{sanitize_filename(servicio)}_{numero}_{sanitize_filename(data.get('r_f_c', 'N/A'))}.zip"

        upload = authenticate_and_upload_to_drive(final_zip, zip_buffer)
        logger.info("Drive: %s", upload.get("message"))

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name=final_zip)

    except RequestEntityTooLarge:
        return jsonify({"error": "Archivo demasiado grande"}), 413
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        logger.exception("Error interno: %s", e)
        return jsonify({"error": str(e)}), 500
    finally:
        if user_image_path:
            try:
                p = pathlib.Path(user_image_path)
                if p.exists():
                    os.remove(p)
                    try:
                        p.parent.rmdir()
                    except OSError:
                        pass
            except Exception as e:
                logger.warning("No se pudo limpiar /tmp: %s", e)

# Handler Lambda
handler = Mangum(app)

# Sugerencias de corrección y mejoras
# 1. Revisa la validación de usuarios: asegúrate de que las contraseñas estén correctamente hasheadas y comparadas con bcrypt.
# 2. Verifica la configuración de Google Drive: revisa que las credenciales y rutas estén correctamente cargadas.
# 3. Revisa el manejo de archivos temporales: utiliza context managers para evitar fugas de recursos.
# 4. Añade más validaciones a las rutas para evitar entradas maliciosas.
# 5. Implementa logs más detallados en operaciones críticas.

