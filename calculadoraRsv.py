import os
import json
import csv
import sqlite3
import datetime
import time
import secrets
import base64
import hmac
import re
import calendar
import shutil
import tempfile
import unicodedata
from datetime import timezone, timedelta
from io import BytesIO, StringIO
from functools import wraps

try:
    from werkzeug.utils import secure_filename
except ImportError:

    def secure_filename(x):
        return ''.join(c for c in str(x) if c.isalnum() or c in '._- ') or 'file'
    
try:
    import requests

    _HAS_REQUESTS = True
except ImportError:
    requests = None
    _HAS_REQUESTS = False
    print("[!] 'requests' no instalado: sincronización de tasas deshabilitada.")

try:
    from openpyxl import load_workbook

    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False
    print(
        "[!] 'openpyxl' no instalado: importación de Excel (.xlsx) deshabilitada. Usa CSV."
    )

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    session,
    send_file,
    jsonify,
    abort,
)
from fpdf import FPDF
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

# 1. Movemos BASE_DIR arriba para poder usarlo en la configuración del secret_key
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# 2. Lógica del secret_key con archivo local para desarrollo
_rsv_secret = os.environ.get('RSV_SECRET', '').strip()
_debug_env = os.environ.get('RSV_DEBUG', '0') == '1'

if _rsv_secret and len(_rsv_secret) >= 16:
    app.secret_key = _rsv_secret
else:
    _dev_secret_file = os.path.join(BASE_DIR, ".dev_secret")
    if os.path.exists(_dev_secret_file):
        with open(_dev_secret_file, "r") as f:
            app.secret_key = f.read().strip()
    else:
        _sec = os.urandom(32).hex()
        with open(_dev_secret_file, "w") as f:
            f.write(_sec)
        app.secret_key = _sec
    print(
        "[!] RSV_SECRET no configurado. Usando clave local persistente en .dev_secret"
    )

# 3. El resto de tu configuración sigue igual
app.config['CSRF_ENABLED'] = False
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# (Asegúrate de borrar la línea original de BASE_DIR que estaba aquí abajo para no duplicarla)

DB_NAME = os.environ.get('RSV_DB_PATH') or os.path.join(BASE_DIR, "calculadora_rsv.db")
UPLOADS_DIR = os.path.join(BASE_DIR, "static", "uploads", "reminders")
os.makedirs(UPLOADS_DIR, exist_ok=True)

RECEPTION_PHOTOS_DIR = os.path.join(BASE_DIR, "static", "uploads", "receptions")
os.makedirs(RECEPTION_PHOTOS_DIR, exist_ok=True)

ALLOWED_PHOTO_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.gif'}
MAX_PHOTOS_PER_RECEPTION = 20
MAX_PHOTO_SIZE_MB = 5

PWA_NAME = "Presupuesto y Cotización RSV"
PWA_SHORT_NAME = "RSV"
PWA_DESCRIPTION = "Presupuestos, revisiones y gestión para taller automotriz"
PWA_THEME_COLOR = "#00E000"
PWA_BG_COLOR = "#0a0a0a"
VEH_TIMEZONE = timezone(timedelta(hours=-4))


def get_ve_time():
    return datetime.datetime.now(VEH_TIMEZONE)


def up(s):
    return str(s).strip().upper() if s else ''


def safe_float(value, default=0.0):
    if value is None:
        return default
    try:
        if isinstance(value, str):
            value = value.strip().replace(',', '.')
            if value == '':
                return default
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_int(value, default=0):
    if value is None:
        return default
    try:
        if isinstance(value, str):
            value = value.strip()
            if value == '':
                return default
        return int(float(value))
    except (ValueError, TypeError):
        return default


# ==========================================
# HELPERS DE CÉDULA / RIF VENEZOLANOS
# ==========================================
_RIF_RE = re.compile(r'^([VEJGPR])[\-\s]?(\d{7,10})[\-\s](\d)$', re.IGNORECASE)
# Cambiamos \d{5,10} por \d{6,10} para exigir mínimo 6 dígitos
_DOC_RE = re.compile(r'^([VEJGPR])[\-]?(\d{6,10})$', re.IGNORECASE)
_DIGITS_RE = re.compile(r'^\d{5,10}$')


def normalize_client_doc(raw):
    """
    Normaliza una cédula/RIF al formato PREFIX-NNNNN (o PREFIX-NNNNNNN-V para RIF).

    - 'v12345678'         -> 'V-12345678'
    - 'V 12.345.678'      -> 'V-12345678'
    - '12345678'          -> 'V-12345678'    (asume V)
    - 'J-40123456-7'      -> 'J-40123456-7'  (RIF con verificador)
    - ''  o basura        -> None
    """
    if raw is None:
        return None
    s = str(raw).strip().upper()
    if not s:
        return None

    # 1. RIF con dígito verificador separado por guión: J-40123456-7
    m = _RIF_RE.match(s)
    if m:
        return f"{m.group(1).upper()}-{m.group(2)}-{m.group(3)}"

    # Limpiar espacios y puntos para el resto
    s_clean = re.sub(r'[\s\.]', '', s)

    # 2. Formato simple PREFIX + dígitos: V-12345678, J40123456
    m = _DOC_RE.match(s_clean)
    if m:
        return f"{m.group(1).upper()}-{m.group(2)}"

    # 3. Solo dígitos -> asumir V (venezolano)
    if _DIGITS_RE.match(s_clean):
        return f"V-{s_clean}"

    return None


def is_temp_doc(doc):
    """True si es un placeholder SIN_DOC_... o está vacío."""
    return not doc or str(doc).startswith('SIN_DOC_')


def find_client_by_doc(doc):
    """Busca un cliente por cédula exacta. Devuelve la fila o None."""
    if not doc:
        return None
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM clients WHERE client_doc = ?", (doc,)).fetchone()
    conn.close()
    return row


def find_client_similar(name, lastname, exclude_doc=None):
    """
    Busca clientes con el mismo nombre+apellido (normalizado) para
    detectar posibles duplicados con distinta cédula.
    """
    full = f"{(name or '').strip()} {(lastname or '').strip()}".strip().lower()
    if not full:
        return []
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT * FROM clients
        WHERE LOWER(TRIM(client_name)) = ?
    """,
        (full,),
    ).fetchall()
    conn.close()
    if exclude_doc:
        rows = [r for r in rows if r['client_doc'] != exclude_doc]
    return rows


def get_csrf_token():
    token = session.get('csrf_token')
    if not token:
        token = secrets.token_urlsafe(32)
        session['csrf_token'] = token
    return token


@app.context_processor
def inject_csrf_token():
    return {'csrf_token': get_csrf_token}


@app.context_processor
def inject_role_flags():
    role = session.get('role', '')
    return {
        'is_readonly_reception': is_readonly_reception_role(role),
        'is_tablero_taller_default_role': is_tablero_taller_default_role(role),
        'is_budget_restricted': is_budget_restricted_role(
            role
        ),  # <-- AGREGAR ESTA LÍNEA
    }


_CSRF_EXEMPT_ENDPOINTS = {
    'login',
    'public_budget',
    'favicon',
    'pwa_manifest',
    'pwa_icon',
    'service_worker',
    'static',
}


@app.before_request
def csrf_protect():
    if request.method in ('GET', 'HEAD', 'OPTIONS', 'TRACE'):
        return
    if not app.config.get('CSRF_ENABLED', False):
        return
    if request.endpoint in _CSRF_EXEMPT_ENDPOINTS:
        return
    token = request.form.get('csrf_token') or request.headers.get('X-CSRF-Token')
    expected = session.get('csrf_token')
    if not expected or not token or not hmac.compare_digest(str(token), str(expected)):
        abort(403, description="Token CSRF inválido.")


ALL_SYSTEM_MODULES = [
    "🏠 Inicio",
    "📝 Nuevo Presupuesto",
    "📥 Recepciones",
    "🔬 Revisiones",
    "🔧 Tablero Taller",
    "📇 Clientes",
    "📁 Expedientes",
    "📜 Historial Clientes",
    "🚗 Vehículos",
    "📅 Citas",
    "📈 Ganancias",
    "📊 Reportes",
    "📂 Presupuestos",
    "📋 Historial de Acciones",
    "⚙️ Plantillas",
    "🔔 Recordatorios",
    "👥 Usuarios",
    "🛡️ Roles y Permisos",
    "📦 Inventario",
    "🛒 Compras",
    "📋 Comandas",
    "🛒 Solicitudes de Compra",
]

# ==========================================================
# PERMISOLOGÍA — Roles con restricciones especiales
# ==========================================================
# Roles que al iniciar sesión aterrizan directo en el Tablero Taller
# (y cuya entrada por "/" también redirige allí)
DEFAULT_TABLERO_TALLER_ROLES = {'JEFE SERVICIO TECNICO', 'SUPERVISOR SERV. TECNICO'}

# Roles que solo pueden VER las Recepciones (no crear / editar / eliminar)
READONLY_RECEPTION_ROLES = {'JEFE SERVICIO TECNICO', 'SUPERVISOR SERV. TECNICO'}


def _norm_role_name(role_name):
    """Normaliza el nombre del rol: mayúsculas, sin tildes, espacios simples."""
    if not role_name:
        return ''
    s = unicodedata.normalize('NFD', str(role_name).strip().upper())
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return ' '.join(s.split())


def is_readonly_reception_role(role_name=''):
    return _norm_role_name(role_name) in READONLY_RECEPTION_ROLES


def is_tablero_taller_default_role(role_name=''):
    return _norm_role_name(role_name) in DEFAULT_TABLERO_TALLER_ROLES


# Roles que NO pueden ver NADA de presupuestos
RESTRICTED_BUDGET_ROLES = {'JEFE SERVICIO TECNICO', 'SUPERVISOR SERV. TECNICO'}


def is_budget_restricted_role(role_name=''):
    return _norm_role_name(role_name) in RESTRICTED_BUDGET_ROLES


PAYMENT_METHODS = [
    "Efectivo USD",
    "Efectivo Bs",
    "Pago Móvil",
    "Transferencia Bs",
    "Zelle",
    "Binance USDT",
    "Punto de Venta",
    "Otro",
]

APPOINTMENT_STATUSES = [
    "Pendiente",
    "Confirmada",
    "Atendida",
    "Cancelada",
    "No Asistió",
]

REVISION_DEFAULT_ITEMS = [
    "LLAVES",
    "CONTROL",
    "GUANTERA",
    "AIRE ACONDICIONADO",
    "ANTENA",
    "LIMPIA PARABRISAS",
    "CLAXON",
    "RETROVISORES",
    "EMBLEMAS",
    "TAPICERIA",
    "ALFOMBRAS",
    "TASAS/ RINES",
    "TRIANGULO",
    "CAUCHO DE REPUESTO",
    "GATO",
    "EXTINTOR",
    "TAPA DE COMBUSTIBLE",
    "TAPA LIMPIA PARABRISAS",
    "TAPA REFRIGERANTE",
    "TAPA DE LIGA",
    "TAPA DE ACEITE",
    "VARILLA DE ACEITE",
    "LUZ CORTESIA",
    "LUZ INTERMITENTE",
    "LLAVE DE CRUZ",
    "CABLE AUXILIAR",
    "LUZ DELANTERA DER",
    "LUZ DELANTERA IZQ",
    "LUZ TRASERA DER",
    "LUZ TRASERA IZQ",
]


# ===== Recepciones: componentes que se responden "¿Tiene?" → Sí / No =====
def _norm_system(s):
    s = (s or '').strip().upper()
    return re.sub(r'\s*/\s*', '/', s)


REVISION_YESNO_ITEMS = {
    _norm_system(x)
    for x in [
        "ANTENA",
        "TASAS/RINES",
        "TRIANGULO",
        "CAUCHO DE REPUESTO",
        "GATO",
        "EXTINTOR",
        "TAPA DE COMBUSTIBLE",
        "TAPA LIMPIA PARABRISAS",
        "TAPA REFRIGERANTE",
        "TAPA DE LIGA",
        "TAPA DE ACEITE",
        "VARILLA DE ACEITE",
        "LLAVE DE CRUZ",
        "CABLE AUXILIAR",
    ]
}


def get_revision_options(system_name):
    """Devuelve las opciones de estado según el componente."""
    if _norm_system(system_name) in REVISION_YESNO_ITEMS:
        return ['Sí', 'No']
    return ['Bueno / Óptimo', 'Regular / Desgaste', 'Malo / Reemplazar', 'N/A']


MECHANIC_REVISION_CATEGORIES = [
    {
        'name': 'SUSPENSIÓN',
        'icon': '🔩',
        'items': [
            {
                'system': 'AMORTIGUADORES',
                'variants': ['DEL R', 'DEL L', 'TRA R', 'TRA L'],
            },
            {
                'system': 'BASE DE AMORTIGUADORES',
                'variants': ['DEL R', 'DEL L', 'TRA R', 'TRA L'],
            },
            {'system': 'ESPIRALES'},
            {'system': 'BALLESTAS', 'variants': ['VENCIDAS', 'ROTAS']},
            {'system': 'BARRA DE TORSIÓN'},
            {'system': 'BARRA ESTABILIZADORA', 'variants': ['DELANTERO', 'TRASERO']},
            {
                'system': 'GOMA DE BARRA ESTABILZADORA',
                'variants': ['DELANTERO', 'TRASERO'],
            },
            {'system': 'LÁPIZ', 'variants': ['DELANTERO', 'TRASERO']},
        ],
    },
    {
        'name': 'MOTOR',
        'icon': '⚙️',
        'items': [
            {
                'system': 'BASES DE MOTOR Y CAJA',
                'variants': ['DER', 'IZQ', 'CENTRAL DEL', 'CENTRAL TRA'],
            },
            {'system': 'CORREA DE TIEMPO'},
            {'system': 'CORREA DE ALTERNADOR'},
            {'system': 'CORREA DE COMPRESOR'},
            {'system': 'CORREA ÚNICA'},
            {'system': 'TERMOSTATO'},
            {'system': 'MANGUERAS'},
            {'system': 'RADIADOR'},
            {'system': 'EMPACADURAS', 'variants': ['TAPA VÁLVULAS', 'CÁMARA']},
            {
                'system': 'ESTOPERAS',
                'variants': ['CIGÜEÑAL DEL', 'CIGÜEÑAL TRA', 'LEVAS'],
            },
        ],
    },
    {
        'name': 'FRENOS',
        'icon': '🛑',
        'items': [
            {'system': 'PASTILLAS', 'variants': ['DELANTERO', 'TRASERO']},
            {'system': 'DISCOS', 'variants': ['RECTIFICAR DEL', 'RECTIFICAR TRA']},
            {'system': 'CALIPER', 'variants': ['DEL R', 'DEL L', 'TRA R', 'TRA L']},
            {
                'system': 'PASADORES DE CALIPER',
                'variants': ['DEL R', 'DEL L', 'TRA R', 'TRA L'],
            },
            {'system': 'TAMBORES', 'variants': ['RECTIFICAR DEL', 'RECTIFICAR TRA']},
            {'system': 'BANDAS'},
            {'system': 'CILINDROS DE FRENO', 'variants': ['DELANTERO', 'TRASERO']},
            {'system': 'BOMBA DE FRENO'},
            {'system': 'FUGA FLUIDOS'},
        ],
    },
    {
        'name': 'TREN DELANTERO',
        'icon': '🚙',
        'items': [
            {
                'system': 'MUÑON',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
            {
                'system': 'MOZO',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'MESETA',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
            {
                'system': 'BUJE',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
            {'system': 'RÓTULAS', 'variants': ['DERECHO', 'IZQUIERDO']},
            {
                'system': 'TERMINALES',
                'variants': [
                    'INT DER',
                    'INT IZQ',
                    'EXT DER',
                    'EXT IZQ',
                    'BRAZO LOCO',
                    'BRAZO PITMAN',
                ],
            },
            {'system': 'CAJETÍN DE DIRECCIÓN'},
        ],
    },
    {
        'name': 'TREN TRASERO',
        'icon': '🚗',
        'items': [
            {
                'system': 'MUÑON',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
            {'system': 'MOZO', 'variants': ['DER', 'IZQ']},
            {
                'system': 'MESETA',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
            {
                'system': 'BUJE',
                'variants': ['SUP DER', 'SUP IZQ', 'INF DER', 'INF IZQ'],
            },
        ],
    },
    {
        'name': 'SISTEMA DE PROPULSIÓN',
        'icon': '🔧',
        'items': [
            {
                'system': 'PUNTA DE TRÍPOIDE',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'GOMAS DE PUNTA DE TRIPOIDE',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'RODAMIENTOS',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'CUBOS DE RUEDA',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'TRICETA',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {
                'system': 'COPA CAJA',
                'variants': ['DEL DER', 'DEL IZQ', 'TRA DER', 'TRA IZQ'],
            },
            {'system': 'GOMA DE COPA CAJA', 'variants': ['DER', 'IZQ']},
            {'system': 'EMBRAGUE'},
            {'system': 'TRANSFER'},
            {'system': 'CRUCETA', 'variants': ['DELANTERO', 'TRASERO']},
            {'system': 'CARDÁN', 'variants': ['DELANTERO', 'TRASERO']},
            {'system': 'PUENTE DE CARDÁN'},
            {'system': 'DIFERENCIAL', 'variants': ['DELANTERO']},
            {'system': 'BUJES DE DIFERENCIAL', 'variants': ['DELANTERO', 'TRASERO']},
            {'system': 'PUNTAS DE EJE', 'variants': ['DER', 'IZQ']},
            {'system': 'JUNTAS HOMOCINÉTICA'},
        ],
    },
]

MREV_MARK_STATES = ['', '*', '✓', 'X']  # Ciclo: vacío → parcial → ok → malo

RECEPTION_CHECKLIST_UI = [
    {
        'title': 'Iluminación Delantera',
        'icon': '💡',
        'items': [
            {
                'name': 'Luz Delantera',
                'fields': ['LUZ DELANTERA IZQ', 'LUZ DELANTERA DER'],
            },
            {'name': 'Luz Alta', 'fields': ['LUZ ALTA IZQ', 'LUZ ALTA DER']},
            {'name': 'Luz Baja', 'fields': ['LUZ BAJA IZQ', 'LUZ BAJA DER']},
            {
                'name': 'Luz de Neblina',
                'fields': ['LUZ DE NEBLINA IZQ', 'LUZ DE NEBLINA DER'],
            },
        ],
    },
    {
        'title': 'Iluminación Trasera',
        'icon': '🔦',
        'items': [
            {'name': 'Luz Trasera', 'fields': ['LUZ TRASERA IZQ', 'LUZ TRASERA DER']},
            {
                'name': 'Luz de Freno',
                'fields': ['LUZ DE FRENO IZQ', 'LUZ DE FRENO DER'],
            },
            {
                'name': 'Luz de Retroceso',
                'fields': ['LUZ DE RETROCESO IZQ', 'LUZ DE RETROCESO DER'],
            },
        ],
    },
    {
        'title': 'Luces Interiores',
        'icon': '✨',
        'items': [
            {'name': 'Luz Cortesía', 'fields': ['LUZ CORTESIA']},
            {'name': 'Luz Intermitente', 'fields': ['LUZ INTERMITENTE']},
        ],
    },
    {
        'title': 'Equipamiento y Herramientas',
        'icon': '🧰',
        'items': [
            {'name': 'Llaves', 'fields': ['LLAVES']},
            {'name': 'Control', 'fields': ['CONTROL']},
            {'name': 'Gato', 'fields': ['GATO']},
            {'name': 'Triángulo', 'fields': ['TRIANGULO']},
            {'name': 'Caucho de Repuesto', 'fields': ['CAUCHO DE REPUESTO']},
            {'name': 'Extintor', 'fields': ['EXTINTOR']},
            {'name': 'Llave de Cruz', 'fields': ['LLAVE DE CRUZ']},
            {'name': 'Cable Auxiliar', 'fields': ['CABLE AUXILIAR']},
        ],
    },
    {
        'title': 'Exterior y Carrocería',
        'icon': '🚗',
        'items': [
            {'name': 'Antena', 'fields': ['ANTENA']},
            {'name': 'Limpia Parabrisas', 'fields': ['LIMPIA PARABRISAS']},
            {'name': 'Retrovisores', 'fields': ['RETROVISORES']},
            {'name': 'Emblemas', 'fields': ['EMBLEMAS']},
            {'name': 'Tapas/Rines', 'fields': ['TASAS/RINES']},
            {'name': 'Tapa Combustible', 'fields': ['TAPA DE COMBUSTIBLE']},
            {'name': 'Tapa Limpiaparabrisas', 'fields': ['TAPA LIMPIAPARABRISAS']},
            {'name': 'Tapa Refrigerante', 'fields': ['TAPA REFRIGERANTE']},
            {'name': 'Tapa de Liga', 'fields': ['TAPA DE LIGA']},
            {'name': 'Tapa de Aceite', 'fields': ['TAPA DE ACEITE']},
            {'name': 'Varilla de Aceite', 'fields': ['VARILLA DE ACEITE']},
        ],
    },
    {
        'title': 'Interior y Tapicería',
        'icon': '🪑',
        'items': [
            {'name': 'Guantera', 'fields': ['GUANTERA']},
            {'name': 'Aire Acondicionado', 'fields': ['AIRE ACONDICIONADO']},
            {'name': 'Claxon', 'fields': ['CLAXON']},
            {'name': 'Tapicería', 'fields': ['TAPICERIA']},
            {'name': 'Alfombras', 'fields': ['ALFOMBRAS']},
        ],
    },
]


RECEPTION_FIELDS_FLAT = []
for cat in RECEPTION_CHECKLIST_UI:
    for item in cat['items']:
        for f in item['fields']:
            RECEPTION_FIELDS_FLAT.append(f)


RECEPTION_MARK_STATES = ['', 'BUENO', 'MALO']

DASHBOARD_LIGHTS = [
    'Mantenimiento',
    'Luz de Cruce (Der / Izq)',
    'Freno de Mano',
    'Cinturón de Seguridad',
    'Aire Acondicionado',
    'ABS',
    'Limpia Parabrisas',
    'Stability Trance',
    'Luz de Aceite',
    '4WD',
    'Presión de Neumático',
    'Airbags',
    'Batería',
    'Luz de Motor',
    'Luces de Puertas Abiertas',
]

# Opciones del selector de gasolina (de izquierda a derecha)
FUEL_LEVEL_OPTIONS = ['Vacío (E)', '1/4', '1/2', '3/4', 'Lleno (F)']


def mechanic_revision_items_flat():
    """Devuelve la estructura completa con índices para el template."""
    flat = []
    for cat_idx, cat in enumerate(MECHANIC_REVISION_CATEGORIES):
        for item_idx, item in enumerate(cat['items']):
            variants = item.get('variants') or [None]
            for var_idx, variant in enumerate(variants):
                flat.append(
                    {
                        'cat_idx': cat_idx,
                        'item_idx': item_idx,
                        'var_idx': var_idx,
                        'category': cat['name'],
                        'icon': cat['icon'],
                        'system': item['system'],
                        'variant': variant,
                        'field_mark': f'mrev_mark_{cat_idx}_{item_idx}_{var_idx}',
                        'field_obs': f'mrev_obs_{cat_idx}_{item_idx}',
                    }
                )
    return flat


REMINDER_PRIORITIES = ["Alta", "Media", "Baja"]

_WORKSHOP_SELLERS_CACHE = None


def get_workshop_sellers():
    global _WORKSHOP_SELLERS_CACHE
    if _WORKSHOP_SELLERS_CACHE is not None:
        return _WORKSHOP_SELLERS_CACHE

    try:
        conn = get_db_connection()

        # Buscamos roles que contengan "VENTA" o "VENDEDOR" para capturar "JEFE VENTAS" y "VENDEDOR TALLER"
        rows = conn.execute("""
            SELECT username FROM users
            WHERE UPPER(role) LIKE '%VENTA%' OR UPPER(role) LIKE '%VENDEDOR%'
            ORDER BY id ASC
        """).fetchall()
        sellers = [r['username'] for r in rows]

        # Si no hay nadie con esos roles, usamos a los administradores como respaldo
        if not sellers:
            rows = conn.execute(
                "SELECT username FROM users WHERE UPPER(role) = 'ADMINISTRADOR' ORDER BY id ASC"
            ).fetchall()
            sellers = [r['username'] for r in rows]

        conn.close()
    except Exception:
        sellers = []

    _WORKSHOP_SELLERS_CACHE = sellers if sellers else ['admin']
    return _WORKSHOP_SELLERS_CACHE


WORKSHOP_MECHANICS = [
    {'key': 'efrain_charmel', 'name': 'EFRAIN CHARMEL', 'color': '#4bce97'},
    {'key': 'jose_diaz', 'name': 'JOSE DIAZ', 'color': '#f5cd47'},
    {'key': 'alexander_alvarez', 'name': 'ALEXANDER ALVAREZ', 'color': '#fea362'},
    {'key': 'luis_ramirez', 'name': 'LUIS RAMIREZ', 'color': '#e56910'},
    {'key': 'alfredo_garcia', 'name': 'ALFREDO GARCIA', 'color': '#f87168'},
    {'key': 'sleiher_huaman', 'name': 'SLEIHER HUAMAN', 'color': '#579dff'},
    {'key': 'diego_bravo', 'name': 'DIEGO BRAVO', 'color': '#b3bac5'},
    {'key': 'ivens_veloz', 'name': 'IVENS VELOZ', 'color': '#344563'},
]

WORKSHOP_REPAIR_STATUSES = [
    'Espera de Servicio Técnico',
    'Espera de Repuestos',
    'Espera de Aprob. Cliente',
    'Espera por Prueba de Campo',
    'Trabajo de Alineación',
    'Por Entregar',
]

WORKSHOP_DEFAULT_CHECKLIST = [
    {
        "text": "REVISIÓN Y LIMPIEZA DE VOLANTE Y PIEZAS MANCHADAS + CALIBRACIÓN NEUMÁTICA Y AJUSTE DE NEUMÁTICOS CON LLAVE DE FORMA MANUAL",
        "done": False,
    }
]

REMINDER_CATEGORIES_DEFAULT = [
    "Taller",
    "Administración",
    "Clientes",
    "Personal",
    "Compras",
    "Otro",
]
REMINDER_RECURRENCES = ["none", "daily", "weekly", "monthly"]


@app.route('/favicon.ico')
def favicon():
    return redirect('/icon-192.png', code=302)


@app.route('/manifest.json')
def pwa_manifest():
    manifest = {
        "name": PWA_NAME,
        "short_name": PWA_SHORT_NAME,
        "description": PWA_DESCRIPTION,
        "start_url": "/",
        "scope": "/",
        "display": "standalone",
        "orientation": "any",
        "background_color": PWA_BG_COLOR,
        "theme_color": PWA_THEME_COLOR,
        "lang": "es",
        "dir": "ltr",
        "categories": ["business", "productivity"],
        "icons": [
            {
                "src": "/icon-192.png",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any",
            },
            {
                "src": "/icon-512.png",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any",
            },
            {
                "src": "/icon-192.png",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "maskable",
            },
            {
                "src": "/icon-512.png",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "maskable",
            },
        ],
    }
    return (
        jsonify(manifest),
        200,
        {
            'Content-Type': 'application/manifest+json',
            'Cache-Control': 'public, max-age=3600',
        },
    )


@app.route('/icon-<int:size>.png')
def pwa_icon(size):
    if size not in (192, 512):
        abort(404)
    path = os.path.join(BASE_DIR, 'static', 'img', f'icon-{size}.png')
    if os.path.exists(path):
        return send_file(path, mimetype='image/png')
    fallback = os.path.join(BASE_DIR, 'static', 'img', 'logo.png')
    if os.path.exists(fallback):
        return send_file(fallback, mimetype='image/png')
    abort(404)


@app.route('/sw.js')
def service_worker():
    sw = r"""
const CACHE_VERSION = 'rsv-v6';
const CORE_ASSETS = ['/', '/manifest.json', '/icon-192.png', '/icon-512.png'];
self.addEventListener('install', function(event) {
    event.waitUntil(
        caches.open(CACHE_VERSION).then(function(cache) {
            return Promise.all(CORE_ASSETS.map(function(url) {
                return cache.add(url).catch(function() {});
            }));
        }).then(function() { return self.skipWaiting(); })
    );
});
self.addEventListener('activate', function(event) {
    event.waitUntil(
        caches.keys().then(function(keys) {
            return Promise.all(keys.filter(function(k) { return k !== CACHE_VERSION; })
                .map(function(k) { return caches.delete(k); }));
        }).then(function() { return self.clients.claim(); })
    );
});
self.addEventListener('fetch', function(event) {
    const req = event.request;
    if (req.method !== 'GET') return;
    const url = new URL(req.url);
    if (url.origin !== self.location.origin) return;
    if (url.pathname.startsWith('/api/')) return;
    if (url.pathname.startsWith('/admin/')) return;
    if (url.pathname.indexOf('/pdf/') !== -1) return;
    if (url.pathname.indexOf('/preview/') !== -1) return;
    if (url.pathname.startsWith('/static/') || /\.(css|js|png|jpg|jpeg|svg|woff2?)$/.test(url.pathname)) {
        event.respondWith(
            caches.match(req).then(function(cached) {
                if (cached) return cached;
                return fetch(req).then(function(res) {
                    if (res && res.status === 200) {
                        const copy = res.clone();
                        caches.open(CACHE_VERSION).then(function(c) { c.put(req, copy); });
                    }
                    return res;
                });
            })
        );
        return;
    }
    event.respondWith(
        fetch(req).then(function(res) {
            if (res && res.status === 200) {
                const copy = res.clone();
                caches.open(CACHE_VERSION).then(function(c) { c.put(req, copy); });
            }
            return res;
        }).catch(function() {
            return caches.match(req).then(function(cached) { return cached || caches.match('/'); });
        })
    );
});
"""
    return (
        sw,
        200,
        {
            'Content-Type': 'application/javascript',
            'Service-Worker-Allowed': '/',
            'Cache-Control': 'no-cache',
        },
    )


def get_db_connection():
    conn = sqlite3.connect(DB_NAME, timeout=20)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA busy_timeout=20000")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA cache_size=-64000")
        conn.execute("PRAGMA temp_store=MEMORY")
        conn.execute("PRAGMA mmap_size=134217728")
    except Exception:
        pass
    return conn


def check_permission(role_name, module_name):
    conn = get_db_connection()
    row = conn.execute(
        "SELECT 1 FROM role_permissions WHERE role_name = ? AND module_name = ?",
        (role_name, module_name),
    ).fetchone()
    conn.close()
    return row is not None


def check_action_permission(role_name, action_name):
    """Verifica si un rol tiene permiso para una acción específica."""
    if role_name == 'Administrador':
        return True  # El admin siempre puede todo
    conn = get_db_connection()
    # Reutilizamos la tabla role_permissions agregando el prefijo 'action:'
    row = conn.execute(
        "SELECT 1 FROM role_permissions WHERE role_name = ? AND module_name = ?",
        (role_name, f"action:{action_name}"),
    ).fetchone()
    conn.close()
    return row is not None


def block_mechanic(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        role = session.get('role', '').lower()
        if 'mecanico' in role or 'mecánico' in role:
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return (
                    jsonify(
                        {
                            'ok': False,
                            'message': 'Acción no permitida para el rol Mecánico.',
                        }
                    ),
                    403,
                )
            flash('⛔ No tienes permisos para modificar el tablero.', 'danger')
            return redirect(
                request.referrer or url_for('index', module='🔧 Tablero Taller')
            )
        return f(*args, **kwargs)

    return decorated_function


def block_readonly_reception(f):
    """Bloquea operaciones de ESCRITURA (POST/PUT/DELETE/PATCH) sobre
    Recepciones para los roles con acceso de solo lectura.
    Los GET pasan sin problema (para que puedan VER)."""

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if request.method in ('POST', 'PUT', 'DELETE', 'PATCH'):
            if is_readonly_reception_role(session.get('role', '')):
                msg = (
                    '⛔ Tu rol tiene acceso de SOLO LECTURA a las Recepciones. '
                    'No puedes crear ni modificar.'
                )
                if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    return jsonify({'ok': False, 'message': msg}), 403
                flash(msg, 'danger')
                return redirect(
                    request.referrer or url_for('index', module='📥 Recepciones')
                )
        return f(*args, **kwargs)

    return decorated_function


def block_budget_access(f):
    """Bloquea totalmente el acceso a Presupuestos para los roles restringidos."""

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if is_budget_restricted_role(session.get('role', '')):
            flash(
                '⛔ Tu rol no tiene acceso al módulo de Presupuestos / Órdenes.',
                'danger',
            )
            return redirect(url_for('index', module='🔧 Tablero Taller'))
        return f(*args, **kwargs)

    return decorated_function


def get_allowed_modules(role_name):
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT module_name FROM role_permissions WHERE role_name = ?", (role_name,)
    ).fetchall()
    conn.close()
    allowed = {r['module_name'] for r in rows}

    # Ocultar módulos de presupuestos y reportes a los roles de servicio técnico
    if is_budget_restricted_role(role_name):
        restricted = {
            "📝 Nuevo Presupuesto",
            "📂 Presupuestos",
            "📈 Ganancias",
            "📊 Reportes",
        }
        allowed.difference_update(restricted)

    return allowed


def get_user_by_username(username):
    if not username:
        return None
    conn = get_db_connection()
    user = conn.execute(
        "SELECT * FROM users WHERE username = ?", (username,)
    ).fetchone()
    conn.close()
    return user


def log_action(
    action, description, entity_type="", entity_id=None, username=None, conn=None
):
    """
    Registra una acción en el historial.

    - Si se pasa `conn`, usa ESA conexión (aprovecha la transacción en curso).
      En ese caso NO hace commit: lo hace el llamador.
    - Si no se pasa, abre una conexión propia, con timeout corto (3 s),
      y garantiza su cierre con try/finally incluso si falla.
    """
    own_conn = conn is None
    try:
        actor = username or session.get('username', 'Sistema')
        ip_address = ""
        try:
            ip_address = request.remote_addr or ""
        except RuntimeError:
            pass

        if own_conn:
            conn = sqlite3.connect(DB_NAME, timeout=3)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA foreign_keys=ON")

        try:
            conn.execute(
                """
                INSERT INTO action_logs
                    (date, username, action, entity_type, entity_id, description, ip_address)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    get_ve_time().strftime("%Y-%m-%d %H:%M"),
                    actor or "Sistema",
                    str(action)[:80],
                    str(entity_type)[:40],
                    entity_id,
                    str(description)[:500],
                    ip_address[:100],
                ),
            )
            if own_conn:
                conn.commit()
        finally:
            if own_conn:
                conn.close()

    except Exception as e:
        import sys

        print(f"[historial] No se pudo registrar: {e}", file=sys.stderr)


def generate_public_token():
    return secrets.token_urlsafe(16)


def compute_payment_status(budget):
    if budget['status'] == 'Anulado':
        return 'Anulado'
    total = float(budget['total_foreign'] or 0)
    paid = (
        float(budget['total_paid_foreign'] or 0)
        if 'total_paid_foreign' in budget.keys()
        else 0.0
    )
    if total <= 0:
        return 'Sin Pago'
    if paid <= 0.01:
        return 'Sin Pago'
    if paid >= total - 0.01:
        return 'Pagado'
    return 'Parcial'


def update_stock_for_budget(conn, budget_id, action, is_comanda=False):
    budget = conn.execute(
        "SELECT items_json, client_name FROM saved_budgets WHERE id = ?", (budget_id,)
    ).fetchone()
    if not budget:
        return

    try:
        items = json.loads(budget['items_json'] or '[]')
    except Exception:
        items = []

    if action == 'deduct' and not is_comanda:
        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")

        existing_comanda = conn.execute(
            "SELECT id FROM comandas WHERE budget_id = ?", (budget_id,)
        ).fetchone()
        if existing_comanda:
            return

        cursor = conn.execute(
            """
            INSERT INTO comandas (budget_id, client_name, status, created_at, updated_at)
            VALUES (?, ?, 'Pendiente de Almacén', ?, ?)
        """,
            (budget_id, budget['client_name'], now_str, now_str),
        )
        comanda_id = cursor.lastrowid

        for it in items:
            if it.get('type') == 'PRODUCTO':
                code = it.get('code', '').strip()
                name = it.get('name', 'PRODUCTO SIN NOMBRE').strip()
                qty = int(it.get('qty', 0))

                if not code or code == 'N/A':
                    code = f"S/N-{secrets.token_hex(2).upper()}"

                if qty > 0:
                    prod = conn.execute(
                        "SELECT id, stock, stock_1, stock_2 FROM inventory WHERE UPPER(code) = UPPER(?)",
                        (code,),
                    ).fetchone()
                    current_stock = float(prod['stock'] or 0) if prod else 0.0

                    if current_stock >= qty:
                        deduct_1 = min(float(prod['stock_1'] or 0), qty)
                        deduct_2 = qty - deduct_1
                        new_stock_1 = float(prod['stock_1'] or 0) - deduct_1
                        new_stock_2 = max(0.0, float(prod['stock_2'] or 0) - deduct_2)
                        new_stock = new_stock_1 + new_stock_2
                        conn.execute(
                            "UPDATE inventory SET stock = ?, stock_1 = ?, stock_2 = ? WHERE id = ?",
                            (new_stock, new_stock_1, new_stock_2, prod['id']),
                        )
                        item_status = 'En Stock'
                    else:
                        item_status = 'Faltante'

                    conn.execute(
                        """
                        INSERT INTO comanda_items (comanda_id, inv_code, name, qty, status)
                        VALUES (?, ?, ?, ?, ?)
                    """,
                        (comanda_id, code, name, qty, item_status),
                    )

    elif action == 'restore':
        existing = conn.execute(
            "SELECT id FROM comandas WHERE budget_id = ?", (budget_id,)
        ).fetchone()
        if existing:
            items_db = conn.execute(
                "SELECT * FROM comanda_items WHERE comanda_id = ? AND status IN ('En Stock', 'Entregado', 'Listo para Despacho')",
                (existing['id'],),
            ).fetchall()
            for it in items_db:
                conn.execute(
                    "UPDATE inventory SET stock = stock + ?, stock_1 = stock_1 + ? WHERE UPPER(code) = UPPER(?)",
                    (it['qty'], it['qty'], it['inv_code']),
                )

            conn.execute(
                "DELETE FROM comanda_items WHERE comanda_id = ?", (existing['id'],)
            )
            conn.execute("DELETE FROM comandas WHERE id = ?", (existing['id'],))


def wa_normalize_phone(phone):
    if not phone:
        return ''
    digits = ''.join(c for c in str(phone) if c.isdigit())
    if not digits:
        return ''
    if digits.startswith('58'):
        return digits
    if digits.startswith('0'):
        return '58' + digits[1:]
    return '58' + digits


def build_whatsapp_message(budget, public_url=''):
    moneda = budget['currency']
    total = float(budget['total_foreign'] or 0)
    msg = (
        f"Hola {budget['client_name']}, le compartimos el presupuesto de "
        f"*Ruedas y Suspensiones* 🚗\n\n"
        f"📋 Presupuesto #{budget['id']}\n"
        f"🚘 Vehículo: {budget['vehicle_name'] or '-'} {('(' + budget['vehicle_plate'] + ')') if budget['vehicle_plate'] else ''}\n"
        f"💰 Total: {total:.2f} {moneda}\n"
    )
    if public_url:
        msg += f"\n🔗 Ver y aprobar online:\n{public_url}\n"
    msg += "\nVálido por 3 días hábiles. ¡Gracias!"
    return msg


def build_whatsapp_message_revision(rev, public_url=''):
    """Mensaje WhatsApp para Recepciones/Revisiones."""
    veh_parts = []
    if rev['vehicle_brand']:
        veh_parts.append(rev['vehicle_brand'])
    if rev['vehicle_model']:
        veh_parts.append(rev['vehicle_model'])
    if rev['vehicle_year']:
        veh_parts.append(str(rev['vehicle_year']))
    veh_line = ' '.join(veh_parts).strip() or (rev['vehicle_name'] or 'Vehículo')
    if rev['vehicle_plate']:
        veh_line += f" ({rev['vehicle_plate']})"

    # Detectar tipo para personalizar el título
    titulo = "la revisión técnica"
    try:
        if json.loads(rev['reception_items_json'] or '[]'):
            titulo = "la planilla de inspección"
    except Exception:
        pass

    msg = (
        f"Hola {rev['client_name']}, le compartimos {titulo} de su vehículo "
        f"en *Ruedas y Suspensiones* 🚗\n\n"
        f"🚘 Vehículo: {veh_line}\n"
        f"📅 Fecha: {rev['date']}\n"
    )
    if public_url:
        msg += f"\n🔗 Ver online:\n{public_url}\n"
    msg += "\n¡Gracias por su preferencia!"
    return msg


def split_full_name(full, last=''):
    full = (full or '').strip()
    last = (last or '').strip()

    # Si tenemos el apellido explícito y coincide con el final
    if last and full.lower().endswith(last.lower()):
        return full[: -len(last)].strip(), last

    # Si no hay apellido, pero el nombre completo tiene espacios, separamos por el último espacio
    if not last and ' ' in full:
        parts = full.rsplit(' ', 1)
        return parts[0].strip(), parts[1].strip()

    return full, last


def recalc_cart_item(item):
    try:
        disc = float(item.get('discount', 0) or 0)
    except (ValueError, TypeError):
        disc = 0.0
    disc = max(0.0, min(100.0, disc))
    item['discount'] = disc
    try:
        qty = int(item.get('qty', 1) or 1)
    except (ValueError, TypeError):
        qty = 1
    try:
        unit_price = float(item.get('unit_price', 0) or 0)
    except (ValueError, TypeError):
        unit_price = 0.0
    unit_price = round(unit_price, 2)
    try:
        cost = float(item.get('cost', 0) or 0)
    except (ValueError, TypeError):
        cost = 0.0
    cost = round(cost, 2)
    gross = unit_price * qty
    disc_amt = round(gross * (disc / 100.0), 2)
    item['discount_amount'] = disc_amt
    item['gross_total'] = round(gross, 2)
    item['total'] = round(gross - disc_amt, 2)
    profit = (unit_price - cost) * qty - disc_amt if cost > 0 else gross - disc_amt
    item['profit'] = round(max(0.0, profit), 2)
    return item


def recalc_cart(cart):
    if not isinstance(cart, list):
        cart = []
    gross = discount = net = 0.0
    for it in cart:
        recalc_cart_item(it)
        gross += it.get('gross_total', 0.0)
        discount += it.get('discount_amount', 0.0)
        net += it.get('total', 0.0)
    effective = (discount / gross * 100.0) if gross > 0 else 0.0
    return {
        'gross': round(gross, 2),
        'discount': round(discount, 2),
        'net': round(net, 2),
        'effective_percent': round(effective, 2),
    }


def add_months_to_date(dt, months):
    month = dt.month - 1 + months
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, calendar.monthrange(year, month)[1])
    return dt.replace(year=year, month=month, day=day)


app.jinja_env.filters['enumerate'] = enumerate
app.jinja_env.globals['get_db_connection'] = get_db_connection
app.jinja_env.globals['check_permission'] = check_permission
app.jinja_env.globals['REVISION_DEFAULT_ITEMS'] = REVISION_DEFAULT_ITEMS
app.jinja_env.globals['REVISION_YESNO_ITEMS'] = REVISION_YESNO_ITEMS
app.jinja_env.globals['get_revision_options'] = get_revision_options
app.jinja_env.globals['MECHANIC_REVISION_CATEGORIES'] = MECHANIC_REVISION_CATEGORIES
app.jinja_env.globals['mechanic_revision_items_flat'] = mechanic_revision_items_flat
app.jinja_env.globals['PAYMENT_METHODS'] = PAYMENT_METHODS
app.jinja_env.globals['APPOINTMENT_STATUSES'] = APPOINTMENT_STATUSES
app.jinja_env.globals['compute_payment_status'] = compute_payment_status
app.jinja_env.globals['wa_normalize_phone'] = wa_normalize_phone
app.jinja_env.globals['build_whatsapp_message'] = build_whatsapp_message
app.jinja_env.globals['split_full_name'] = split_full_name
app.jinja_env.globals['REMINDER_PRIORITIES'] = REMINDER_PRIORITIES
app.jinja_env.globals['REMINDER_CATEGORIES_DEFAULT'] = REMINDER_CATEGORIES_DEFAULT
app.jinja_env.globals['REMINDER_RECURRENCES'] = REMINDER_RECURRENCES
app.jinja_env.globals['build_whatsapp_message_revision'] = (
    build_whatsapp_message_revision
)
app.jinja_env.globals['DASHBOARD_LIGHTS'] = DASHBOARD_LIGHTS
app.jinja_env.globals['FUEL_LEVEL_OPTIONS'] = FUEL_LEVEL_OPTIONS


def _from_json(value):
    try:
        return json.loads(value) if value else []
    except Exception:
        return []


app.jinja_env.filters['from_json'] = _from_json


# ==========================================
# HELPERS DE PAGINACIÓN Y VEHÍCULOS
# ==========================================
def paginate(total, page, per_page=20):
    page = max(1, safe_int(page, 1))
    total = max(0, safe_int(total, 0))
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = min(page, total_pages)
    offset = (page - 1) * per_page
    return {
        'page': page,
        'per_page': per_page,
        'total': total,
        'total_pages': total_pages,
        'offset': offset,
        'has_prev': page > 1,
        'has_next': page < total_pages,
        'start': offset + 1 if total > 0 else 0,
        'end': min(offset + per_page, total),
    }


def save_client_vehicle(
    client_doc,
    veh_name,
    veh_year,
    veh_plate,
    veh_brand,
    veh_color,
    veh_model,
    vin='',
    km='',
    notes='',
    conn=None,
):
    if not client_doc:
        return None

    veh_name = (veh_name or '').strip()
    veh_year = (veh_year or '').strip()
    veh_plate = up(veh_plate)
    veh_brand = up(veh_brand)
    veh_color = up(veh_color)
    veh_model = up(veh_model)
    vin = (vin or '').strip()
    km = (km or '').strip()
    notes = (notes or '').strip()

    if not any(
        [veh_name, veh_year, veh_plate, veh_brand, veh_color, veh_model, vin, km, notes]
    ):
        return None

    # 👇 CONTROL DE CONEXIÓN PARA EVITAR DB LOCKED
    own_conn = False
    if conn is None:
        conn = get_db_connection()
        own_conn = True

    try:
        existing = None
        if veh_plate:
            existing = conn.execute(
                "SELECT id FROM vehicles WHERE client_doc = ? AND UPPER(plate) = ? LIMIT 1",
                (client_doc, veh_plate.upper()),
            ).fetchone()
        if not existing and (veh_name or veh_brand or veh_model):
            existing = conn.execute(
                """
                SELECT id FROM vehicles WHERE client_doc = ?
                  AND UPPER(COALESCE(name,''))  = ?
                  AND UPPER(COALESCE(brand,'')) = ?
                  AND UPPER(COALESCE(model,'')) = ?
                  AND COALESCE(year,'')         = ?
                LIMIT 1
            """,
                (
                    client_doc,
                    veh_name.upper(),
                    veh_brand.upper(),
                    veh_model.upper(),
                    veh_year,
                ),
            ).fetchone()

        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")

        if existing:
            conn.execute(
                """
                UPDATE vehicles SET plate = ?, name = ?, brand = ?, model = ?, year = ?, color = ?,
                    vin = COALESCE(NULLIF(?, ''), vin),
                    km  = COALESCE(NULLIF(?, ''), km),
                    notes = COALESCE(NULLIF(?, ''), notes),
                    updated_at = ? WHERE id = ?
            """,
                (
                    veh_plate,
                    veh_name,
                    veh_brand,
                    veh_model,
                    veh_year,
                    veh_color,
                    vin,
                    km,
                    notes,
                    now_str,
                    existing['id'],
                ),
            )
            vid = existing['id']
        else:
            new_token = generate_public_token()
            cur = conn.execute(
                """
                INSERT INTO vehicles (client_doc, plate, name, brand, model, year, color,
                    vin, km, notes, created_at, updated_at, public_token)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    client_doc,
                    veh_plate,
                    veh_name,
                    veh_brand,
                    veh_model,
                    veh_year,
                    veh_color,
                    vin,
                    km,
                    notes,
                    now_str,
                    now_str,
                    new_token,
                ),
            )
            vid = cur.lastrowid

        # 👇 SOLO HACE COMMIT SI LA CONEXIÓN FUE CREADA AQUÍ
        if own_conn:
            conn.commit()

        return vid

    except Exception:
        # 👇 SOLO HACE ROLLBACK SI LA CONEXIÓN FUE CREADA AQUÍ
        if own_conn:
            conn.rollback()
        raise
    finally:
        # 👇 SOLO CIERRA SI LA CONEXIÓN FUE CREADA AQUÍ
        if own_conn:
            conn.close()


def save_or_update_client(
    doc,
    name,
    phone="",
    veh_name="",
    veh_year="",
    veh_plate="",
    veh_brand="",
    veh_color="",
    veh_model="",
    lastname="",
    conn=None,
):
    # Normalizar formato de cédula
    normalized = normalize_client_doc(doc)
    doc = normalized if normalized else (up(doc) if doc else '')

    name = str(name).strip() if name else ""
    lastname = str(lastname).strip() if lastname else ""
    phone = str(phone).strip() if phone else ""
    veh_name = str(veh_name).strip() if veh_name else ""
    veh_year = str(veh_year).strip() if veh_year else ""
    veh_plate = up(veh_plate)
    veh_brand = up(veh_brand)
    veh_color = up(veh_color)
    veh_model = up(veh_model)

    if not doc:
        doc = f"SIN_DOC_{int(time.time())}_{secrets.token_hex(3).upper()}"
    full_name = f"{name} {lastname}".strip() or "SIN NOMBRE"

    own_conn = False
    if conn is None:
        conn = get_db_connection()
        own_conn = True

    try:
        existing = conn.execute(
            "SELECT id FROM clients WHERE client_doc = ?", (doc,)
        ).fetchone()
        if existing:
            if any([veh_name, veh_year, veh_plate, veh_brand, veh_color, veh_model]):
                conn.execute(
                    """
                    UPDATE clients SET client_name = ?, client_lastname = ?, client_phone = ?,
                        vehicle_name = ?, vehicle_year = ?, vehicle_plate = ?,
                        vehicle_brand = ?, vehicle_color = ?, vehicle_model = ?
                    WHERE client_doc = ?
                """,
                    (
                        full_name,
                        lastname,
                        phone,
                        veh_name,
                        veh_year,
                        veh_plate,
                        veh_brand,
                        veh_color,
                        veh_model,
                        doc,
                    ),
                )
            else:
                conn.execute(
                    """
                    UPDATE clients SET client_name = ?, client_lastname = ?, client_phone = ?
                    WHERE client_doc = ?
                """,
                    (full_name, lastname, phone, doc),
                )
        else:
            conn.execute(
                """
                INSERT INTO clients (client_doc, client_name, client_lastname, client_phone,
                    vehicle_name, vehicle_year, vehicle_plate,
                    vehicle_brand, vehicle_color, vehicle_model)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    doc,
                    full_name,
                    lastname,
                    phone,
                    veh_name,
                    veh_year,
                    veh_plate,
                    veh_brand,
                    veh_color,
                    veh_model,
                ),
            )

        # FIX BUG #4: Guardar el vehículo ANTES de cerrar la conexión
        if any([veh_name, veh_year, veh_plate, veh_brand, veh_color, veh_model]):
            try:
                save_client_vehicle(
                    doc,
                    veh_name,
                    veh_year,
                    veh_plate,
                    veh_brand,
                    veh_color,
                    veh_model,
                    conn=conn,
                )
            except Exception as e:
                print(
                    f"[vehicles] Error al guardar vehículo en la misma transacción: {e}"
                )

        if own_conn:
            conn.commit()
    except Exception as e:
        if own_conn:
            conn.rollback()
        raise e
    finally:
        if own_conn:
            conn.close()

    return doc


def migrate_plaintext_passwords():
    conn = get_db_connection()
    try:
        rows = conn.execute("SELECT id, password FROM users").fetchall()
        migrated = 0
        for row in rows:
            pw = row['password'] or ""
            if not pw.startswith(("pbkdf2:", "scrypt:")):
                conn.execute(
                    "UPDATE users SET password = ? WHERE id = ?",
                    (generate_password_hash(pw), row['id']),
                )
                migrated += 1
        if migrated:
            conn.commit()
            print(f"[migración] {migrated} contraseña(s)")
    finally:
        conn.close()


def _ensure_column(cursor, table, column, coldef):
    cursor.execute(f"PRAGMA table_info({table})")
    existing = [r[1] for r in cursor.fetchall()]
    if column not in existing:
        try:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coldef}")
            print(f"[init_db] Columna agregada: {table}.{column}")
        except Exception as e:
            print(f"[init_db] No se pudo agregar {table}.{column}: {e}")


def purge_old_reminders():
    try:
        conn = get_db_connection()
        now = get_ve_time()
        now_str = now.strftime("%Y-%m-%d %H:%M:%S")
        threshold_completed = (now - timedelta(days=90)).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            """
            UPDATE reminders SET archived = 1, archived_at = ?
            WHERE archived = 0 AND is_completed = 1
              AND COALESCE(completed_at, date) != '' AND COALESCE(completed_at, date) < ?
        """,
            (now_str, threshold_completed),
        )
        threshold_archived = (now - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            """
            DELETE FROM reminders
            WHERE archived = 1 AND archived_at != '' AND archived_at < ?
        """,
            (threshold_archived,),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[purge_old_reminders] {e}")


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("PRAGMA journal_mode=WAL")
    except Exception as e:
        print(f"[init_db] No se pudo activar WAL: {e}")

    cursor.execute(
        "CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value REAL NOT NULL)"
    )
    cursor.execute("INSERT OR IGNORE INTO config (key, value) VALUES ('tasa_usd', 0.0)")
    cursor.execute("INSERT OR IGNORE INTO config (key, value) VALUES ('tasa_eur', 0.0)")

    cursor.execute("PRAGMA table_info(clients)")
    cols = [r[1] for r in cursor.fetchall()]

    if not cols:
        cursor.execute("""
            CREATE TABLE clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_doc TEXT UNIQUE NOT NULL,
                client_name TEXT NOT NULL,
                client_lastname TEXT DEFAULT '',
                client_phone TEXT DEFAULT '',
                vehicle_name TEXT DEFAULT '',
                vehicle_year TEXT DEFAULT '',
                vehicle_plate TEXT DEFAULT '',
                vehicle_brand TEXT DEFAULT '',
                vehicle_color TEXT DEFAULT '',
                vehicle_model TEXT DEFAULT ''
            )
        """)
    elif 'name' in cols or 'client_name' not in cols:
        cursor.execute("""
            CREATE TABLE clients_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_doc TEXT UNIQUE NOT NULL,
                client_name TEXT NOT NULL,
                client_lastname TEXT DEFAULT '',
                client_phone TEXT DEFAULT '',
                vehicle_name TEXT DEFAULT '',
                vehicle_year TEXT DEFAULT '',
                vehicle_plate TEXT DEFAULT '',
                vehicle_brand TEXT DEFAULT '',
                vehicle_color TEXT DEFAULT '',
                vehicle_model TEXT DEFAULT ''
            )
        """)
        name_col = 'client_name' if 'client_name' in cols else 'name'
        doc_col = 'client_doc' if 'client_doc' in cols else 'doc'
        phone_col = 'client_phone' if 'client_phone' in cols else 'phone'
        vname_col = 'vehicle_name' if 'vehicle_name' in cols else 'veh_name'
        vyear_col = 'vehicle_year' if 'vehicle_year' in cols else 'veh_year'
        vplate_col = 'vehicle_plate' if 'vehicle_plate' in cols else 'veh_plate'
        cursor.execute(f"""
            INSERT OR IGNORE INTO clients_new (client_doc, client_name, client_phone, vehicle_name, vehicle_year, vehicle_plate)
            SELECT COALESCE({doc_col}, 'DOC_' || id),
                   COALESCE({name_col}, 'Sin Nombre'),
                   COALESCE({phone_col}, ''),
                   COALESCE({vname_col}, ''),
                   COALESCE({vyear_col}, ''),
                   COALESCE({vplate_col}, '')
            FROM clients
        """)
        cursor.execute("DROP TABLE clients")
        cursor.execute("ALTER TABLE clients_new RENAME TO clients")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS saved_budgets (
            id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL,
            client_name TEXT NOT NULL, client_doc TEXT NOT NULL, client_phone TEXT NOT NULL,
            currency TEXT NOT NULL, bcv_rate REAL NOT NULL,
            total_foreign REAL NOT NULL, total_ves REAL NOT NULL,
            total_profit_foreign REAL DEFAULT 0.0, items_json TEXT NOT NULL,
            status TEXT DEFAULT 'Pendiente', vehicle_name TEXT DEFAULT '',
            vehicle_year TEXT DEFAULT '', vehicle_plate TEXT DEFAULT '',
            username TEXT DEFAULT '', budget_note TEXT DEFAULT '',
            base_imponible REAL DEFAULT 0.0, include_iva INTEGER DEFAULT 0,
            iva_amount REAL DEFAULT 0.0, discount_amount REAL DEFAULT 0.0
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS revisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL,
            client_name TEXT NOT NULL, client_doc TEXT NOT NULL, client_phone TEXT NOT NULL,
            vehicle_name TEXT DEFAULT '', vehicle_year TEXT DEFAULT '',
            vehicle_plate TEXT DEFAULT '', items_json TEXT NOT NULL,
            status TEXT DEFAULT 'Pendiente', username TEXT DEFAULT '',
            revision_note TEXT DEFAULT ''
        )
    """)

    # ===== MÓDULO DE COMPRAS =====
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            rif TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            address TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS purchases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            supplier_id INTEGER NOT NULL,
            invoice_num TEXT DEFAULT '',
            doc_type TEXT DEFAULT 'Factura Fiscal', -- NUEVO CAMPO
            total_foreign REAL DEFAULT 0.0,
            status TEXT DEFAULT 'Pendiente',
            credit_days INTEGER DEFAULT 0,
            due_date TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_by TEXT DEFAULT '',
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS purchase_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_id INTEGER NOT NULL,
            inv_id INTEGER,
            code TEXT DEFAULT '',
            name TEXT DEFAULT '',
            qty REAL DEFAULT 0.0,
            cost REAL DEFAULT 0.0,
            total REAL DEFAULT 0.0,
            FOREIGN KEY(purchase_id) REFERENCES purchases(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS comandas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            budget_id INTEGER NOT NULL,
            purchase_id INTEGER,
            client_name TEXT DEFAULT '',
            status TEXT DEFAULT 'Pendiente de Stock',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            notes TEXT DEFAULT '',
            FOREIGN KEY(budget_id) REFERENCES saved_budgets(id),
            FOREIGN KEY(purchase_id) REFERENCES purchases(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS comanda_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            comanda_id INTEGER NOT NULL,
            inv_code TEXT NOT NULL,
            name TEXT NOT NULL,
            qty REAL DEFAULT 0.0,
            status TEXT DEFAULT 'Faltante',
            FOREIGN KEY(comanda_id) REFERENCES comandas(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS comanda_ofertas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            supplier_id INTEGER,
            offered_cost REAL DEFAULT 0.0,
            offered_notes TEXT DEFAULT '',
            username TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            status TEXT DEFAULT 'Pendiente',
            FOREIGN KEY(item_id) REFERENCES comanda_items(id) ON DELETE CASCADE
        )
    """)

    def ensure(table, col, defn):
        _ensure_column(cursor, table, col, defn)

    # ... (dentro de init_db, justo después de crear las tablas de compras) ...

    ensure('purchases', 'doc_type', "TEXT DEFAULT 'Factura Fiscal'")
    ensure('purchase_items', 'extra_data', "TEXT DEFAULT '{}'")
    ensure('revisions', 'revision_note', "TEXT DEFAULT ''")
    ensure('purchases', 'comanda_id', 'INTEGER DEFAULT NULL')
    ensure('comandas', 'purchase_id', 'INTEGER DEFAULT NULL')
    ensure('comanda_items', 'supplier_id', 'INTEGER DEFAULT NULL')
    ensure('comanda_items', 'offered_cost', 'REAL DEFAULT 0.0')
    ensure('comanda_items', 'offered_notes', "TEXT DEFAULT ''")
    ensure('comanda_items', 'approved_by_seller', 'INTEGER DEFAULT 0')
    ensure('comanda_items', 'delivered_at', "TEXT DEFAULT ''")
    ensure('comanda_items', 'delivered_by', "TEXT DEFAULT ''")
    ensure('comanda_items', 'signature_data', "TEXT DEFAULT ''")

    # ===== Recepciones — Planilla de Inspección (columnas nuevas) =====
    for col, defn in [
        ('total_profit_foreign', 'REAL DEFAULT 0.0'),
        ('vehicle_name', "TEXT DEFAULT ''"),
        ('vehicle_year', "TEXT DEFAULT ''"),
        ('vehicle_plate', "TEXT DEFAULT ''"),
        ('username', "TEXT DEFAULT ''"),
        ('budget_note', "TEXT DEFAULT ''"),
        ('base_imponible', 'REAL DEFAULT 0.0'),
        ('include_iva', 'INTEGER DEFAULT 0'),
        ('iva_amount', 'REAL DEFAULT 0.0'),
        ('discount_amount', 'REAL DEFAULT 0.0'),
        ('discount_percent', 'REAL DEFAULT 0.0'),
        ('total_paid_foreign', 'REAL DEFAULT 0.0'),
        ('total_paid_ves', 'REAL DEFAULT 0.0'),
        ('public_token', "TEXT DEFAULT ''"),
        ('client_signature', "TEXT DEFAULT ''"),
        ('client_signature_date', "TEXT DEFAULT ''"),
        ('show_bs_equivalent', 'INTEGER DEFAULT 0'),  # ← NUEVO
    ]:
        ensure('saved_budgets', col, defn)
        ensure('saved_budgets', 'almacen_status', "TEXT DEFAULT 'Por procesar'")
        ensure('saved_budgets', 'budget_number', "TEXT DEFAULT ''")

    ensure('revisions', 'vehicle_km', "TEXT DEFAULT ''")
    ensure('revisions', 'reception_temperature', "TEXT DEFAULT ''")
    # ===== Token público para Recepciones/Revisiones =====
    # Columnas legadas que puedan existir en DBs antiguas
    ensure('revisions', 'details_json', "TEXT DEFAULT '[]'")
    ensure('revisions', 'public_token', "TEXT DEFAULT ''")
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_revisions_token ON revisions(public_token)"
    )

    # Backfill: generar token a las revisiones que no lo tengan
    try:
        rows_no_token = cursor.execute(
            "SELECT id FROM revisions WHERE COALESCE(public_token, '') = ''"
        ).fetchall()
        for r in rows_no_token:
            cursor.execute(
                "UPDATE revisions SET public_token = ? WHERE id = ?",
                (generate_public_token(), r['id']),
            )
        if rows_no_token:
            print(
                f"[init_db] Tokens públicos generados para {len(rows_no_token)} recepción(es)"
            )
    except Exception as e:
        print(f"[init_db] Backfill tokens revisiones: {e}")
        # ===== Revisión mecánica (checklist técnico) — columnas nuevas =====
    for col, defn in [
        ('mechanic_items_json', "TEXT DEFAULT '[]'"),
        ('mechanic_notes', "TEXT DEFAULT ''"),
        ('mechanic_username', "TEXT DEFAULT ''"),
        ('mechanic_date', "TEXT DEFAULT ''"),
        ('mechanic_status', "TEXT DEFAULT 'Pendiente'"),
    ]:
        ensure('revisions', col, defn)
    ensure('clients', 'client_lastname', "TEXT DEFAULT ''")

    for tbl in ('clients', 'saved_budgets', 'revisions'):
        ensure(tbl, 'vehicle_brand', "TEXT DEFAULT ''")
        ensure(tbl, 'vehicle_color', "TEXT DEFAULT ''")
        ensure(tbl, 'vehicle_model', "TEXT DEFAULT ''")

    for col, defn in [
        ('total_profit_foreign', 'REAL DEFAULT 0.0'),
        ('vehicle_name', "TEXT DEFAULT ''"),
        ('vehicle_year', "TEXT DEFAULT ''"),
        ('vehicle_plate', "TEXT DEFAULT ''"),
        ('username', "TEXT DEFAULT ''"),
        ('budget_note', "TEXT DEFAULT ''"),
        ('base_imponible', 'REAL DEFAULT 0.0'),
        ('include_iva', 'INTEGER DEFAULT 0'),
        ('iva_amount', 'REAL DEFAULT 0.0'),
        ('discount_amount', 'REAL DEFAULT 0.0'),
        ('discount_percent', 'REAL DEFAULT 0.0'),
        ('total_paid_foreign', 'REAL DEFAULT 0.0'),
        ('total_paid_ves', 'REAL DEFAULT 0.0'),
        ('public_token', "TEXT DEFAULT ''"),
        ('client_signature', "TEXT DEFAULT ''"),
        ('client_signature_date', "TEXT DEFAULT ''"),
    ]:
        ensure('saved_budgets', col, defn)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT, budget_id INTEGER NOT NULL,
            date TEXT NOT NULL, amount_foreign REAL NOT NULL, amount_ves REAL NOT NULL,
            currency TEXT NOT NULL, rate_bcv REAL NOT NULL, method TEXT NOT NULL,
            reference TEXT DEFAULT '', username TEXT DEFAULT '', note TEXT DEFAULT '',
            FOREIGN KEY (budget_id) REFERENCES saved_budgets(id) ON DELETE CASCADE
        )
    """)
    ensure('payments', 'original_amount', 'REAL DEFAULT 0.0')

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS service_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT, description TEXT NOT NULL,
            default_price REAL DEFAULT 0.0, currency TEXT DEFAULT 'USD',
            category TEXT DEFAULT '', sort_order INTEGER DEFAULT 0
        )
    """)

    for col, defn in [
        ('reception_items_json', "TEXT DEFAULT '[]'"),
        ('reception_fallas_json', "TEXT DEFAULT '[]'"),
        ('reception_razon_social', "TEXT DEFAULT ''"),
        ('reception_rif_ci', "TEXT DEFAULT ''"),
        ('reception_email', "TEXT DEFAULT ''"),
        ('reception_direccion_fiscal', "TEXT DEFAULT ''"),
        ('reception_es_retorno', "INTEGER DEFAULT 0"),
        ('reception_fecha', "TEXT DEFAULT ''"),
        ('reception_hora', "TEXT DEFAULT ''"),
        ('reception_am_pm', "TEXT DEFAULT ''"),
        ('reception_transmision', "TEXT DEFAULT ''"),
        ('reception_motor', "TEXT DEFAULT ''"),
        ('reception_clave', "TEXT DEFAULT ''"),
        ('reception_status', "TEXT DEFAULT 'Borrador'"),
        ('reception_sketch_json', "TEXT DEFAULT '{}'"),
        ('reception_lights_json', "TEXT DEFAULT '[]'"),
        ('reception_fuel_level', "TEXT DEFAULT ''"),
    ]:
        ensure('revisions', col, defn)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, time TEXT NOT NULL,
            client_name TEXT NOT NULL, client_doc TEXT DEFAULT '', client_phone TEXT DEFAULT '',
            vehicle_name TEXT DEFAULT '', vehicle_plate TEXT DEFAULT '',
            service_description TEXT DEFAULT '', status TEXT DEFAULT 'Pendiente',
            username TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL,
            title TEXT NOT NULL, content TEXT NOT NULL, date TEXT NOT NULL,
            is_seen INTEGER DEFAULT 0
        )
    """)
    for col, defn in [
        ('priority', "TEXT DEFAULT 'Media'"),
        ('category', "TEXT DEFAULT ''"),
        ('due_date', "TEXT DEFAULT ''"),
        ('attachment', "TEXT DEFAULT ''"),
        ('attachment_name', "TEXT DEFAULT ''"),
        ('attachment_is_image', 'INTEGER DEFAULT 0'),
        ('is_completed', 'INTEGER DEFAULT 0'),
        ('completed_at', "TEXT DEFAULT ''"),
        ('completed_by', "TEXT DEFAULT ''"),
        ('is_private', 'INTEGER DEFAULT 0'),
        ('client_doc', "TEXT DEFAULT ''"),
        ('client_name', "TEXT DEFAULT ''"),
        ('vehicle_plate', "TEXT DEFAULT ''"),
        ('assigned_to', "TEXT DEFAULT ''"),
        ('recurrence', "TEXT DEFAULT 'none'"),
        ('archived', 'INTEGER DEFAULT 0'),
        ('archived_at', "TEXT DEFAULT ''"),
        ('updated_at', "TEXT DEFAULT ''"),
    ]:
        ensure('reminders', col, defn)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_reminder_views (
            username TEXT NOT NULL, reminder_id INTEGER NOT NULL,
            PRIMARY KEY (username, reminder_id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL, role TEXT NOT NULL, first_name TEXT DEFAULT '',
            last_name TEXT DEFAULT '', phone TEXT DEFAULT '', show_in_pdf INTEGER DEFAULT 0
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS action_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL,
            username TEXT NOT NULL DEFAULT '', action TEXT NOT NULL,
            entity_type TEXT DEFAULT '', entity_id INTEGER,
            description TEXT NOT NULL DEFAULT '', ip_address TEXT DEFAULT ''
        )
    """)
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_action_logs_date ON action_logs(date DESC)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_action_logs_user ON action_logs(username)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_action_logs_action ON action_logs(action)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_reminders_visible ON reminders(archived, is_completed, assigned_to)"
    )

    cursor.execute(
        "CREATE TABLE IF NOT EXISTS roles (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL)"
    )
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS role_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT, role_name TEXT NOT NULL,
            module_name TEXT NOT NULL, UNIQUE(role_name, module_name)
        )
    """)

    # ===== INVENTARIO (Compra/Venta) =====
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            description TEXT DEFAULT '',
            brand TEXT DEFAULT '',
            category TEXT DEFAULT '',
            unit TEXT DEFAULT 'UND',
            cost REAL DEFAULT 0.0,
            price REAL DEFAULT 0.0,
            stock REAL DEFAULT 0.0,
            min_stock REAL DEFAULT 0.0,
            location TEXT DEFAULT '',
            supplier TEXT DEFAULT '',
            barcode TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_inventory_code    ON inventory(code)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_inventory_name    ON inventory(name)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_inventory_barcode ON inventory(barcode)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_inventory_cat     ON inventory(category)"
    )

    # ===== INVENTARIO — Nuevas columnas (Formulario Extendido) =====
    for col, defn in [
        ('linea', "TEXT DEFAULT ''"),
        ('division', "TEXT DEFAULT ''"),
        ('montaje', "TEXT DEFAULT ''"),
        ('id_casam', "TEXT DEFAULT ''"),
        ('id_modelo', "TEXT DEFAULT ''"),
        ('year_start', "TEXT DEFAULT ''"),
        ('year_end', "TEXT DEFAULT ''"),
        ('t_anio', "TEXT DEFAULT ''"),
        ('amperaje', "TEXT DEFAULT ''"),
        ('reserva', "TEXT DEFAULT ''"),
        ('ancho', "TEXT DEFAULT ''"),
        ('alto', "TEXT DEFAULT ''"),
        ('largo', "TEXT DEFAULT ''"),
        ('fecha_compra', "TEXT DEFAULT ''"),
        ('regmod', "TEXT DEFAULT ''"),
        ('almacen', "TEXT DEFAULT ''"),
        ('reporte', "TEXT DEFAULT ''"),
        ('temporal', "TEXT DEFAULT ''"),
        ('indice', "TEXT DEFAULT ''"),
        ('imagen', "TEXT DEFAULT ''"),
        ('sede_1', "TEXT DEFAULT 'RYS GR'"),
        ('stock_1', "REAL DEFAULT 0.0"),
        ('sede_2', "TEXT DEFAULT 'RYS VZLA'"),
        ('stock_2', "REAL DEFAULT 0.0"),
    ]:
        ensure('inventory', col, defn)

    cursor.execute("SELECT COUNT(*) FROM roles")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO roles (name) VALUES ('Administrador')")
        cursor.execute("INSERT INTO roles (name) VALUES ('Vendedor')")
        for m in ALL_SYSTEM_MODULES:
            cursor.execute(
                "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', ?)",
                (m,),
            )
        # En tu función init_db(), busca esta lista y agrega "🛒 Compras" al final:
        for m in [
            "🏠 Inicio",
            "📝 Nuevo Presupuesto",
            "📥 Recepciones",
            "📇 Clientes",
            "📜 Historial Clientes",
            "🚗 Vehículos",
            "📅 Citas",
            "📈 Ganancias",
            "📂 Presupuestos",
            "🔔 Recordatorios",
            "🛒 Compras",
        ]:  #
            cursor.execute(
                "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Vendedor', ?)",
                (m,),
            )

    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', '🏠 Inicio')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Vendedor', '🏠 Inicio')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', '📋 Historial de Acciones')"
    )

    # ===== Permisos Inventario =====
    for _role_inv in ('Administrador', 'Vendedor', 'Mecánico'):
        cursor.execute(
            "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES (?, ?)",
            (_role_inv, '📦 Inventario'),
        )

        # ===== Permisos Expedientes =====
    for _role_exp in ('Administrador', 'Vendedor', 'Mecánico'):
        cursor.execute(
            "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES (?, ?)",
            (_role_exp, '📁 Expedientes'),
        )
        # ===== Rol "Mecánico" (opcional) + permiso a "🔬 Revisiones" para Vendedor =====
    cursor.execute("INSERT OR IGNORE INTO roles (name) VALUES ('Mecánico')")
    for m in [
        "🏠 Inicio",
        "📥 Recepciones",
        "🔬 Revisiones",
        "📇 Clientes",
        "🚗 Vehículos",
    ]:
        cursor.execute(
            "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Mecánico', ?)",
            (m,),
        )
    # Dar acceso a Vendedor para que se pueda probar sin crear usuarios nuevos
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Vendedor', '🔬 Revisiones')"
    )

    # ===== Rol "Almacenista" =====
    cursor.execute("INSERT OR IGNORE INTO roles (name) VALUES ('Almacenista')")
    for m in [
        "🏠 Inicio",
        "📂 Presupuestos",
        "📦 Inventario",
        "🛒 Compras",
        "📋 Comandas",
    ]:  #
        cursor.execute(
            "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Almacenista', ?)",
            (m,),
        )
    # Dar también al administrador por si acaso
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', '🛒 Compras')"
    )
    # Dar el permiso de acción para el almacén
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Almacenista', 'action:almacen_toggle')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', 'action:almacen_toggle')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Administrador', '🛒 Solicitudes de Compra')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Almacenista', '🛒 Solicitudes de Compra')"
    )
    cursor.execute(
        "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES ('Vendedor', '🛒 Solicitudes de Compra')"
    )
    # ===== Migración: renombrar módulo "🔍 Revisiones" → "📥 Recepciones" =====
    try:
        old_roles = cursor.execute(
            "SELECT DISTINCT role_name FROM role_permissions WHERE module_name = '🔍 Revisiones'"
        ).fetchall()
        for row in old_roles:
            cursor.execute(
                "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES (?, '📥 Recepciones')",
                (row[0],),
            )
        cursor.execute(
            "DELETE FROM role_permissions WHERE module_name = '🔍 Revisiones'"
        )
        if old_roles:
            print(
                f"[init_db] Módulo renombrado a Recepciones en {len(old_roles)} rol(es)"
            )
    except Exception as e:
        print(f"[init_db] Migración permisos Recepciones: {e}")

    cursor.execute("SELECT COUNT(*) FROM service_templates")
    if cursor.fetchone()[0] == 0:
        for desc, price, curr, cat in [
            ("Alineación y balanceo", 25.0, "USD", "Suspensión"),
            ("Cambio de aceite y filtro", 40.0, "USD", "Mantenimiento"),
            ("Rotación de cauchos", 15.0, "USD", "Cauchos"),
            ("Reparación de pinchazo", 10.0, "USD", "Cauchos"),
            ("Montaje de amortiguador", 30.0, "USD", "Suspensión"),
            ("Revisión general", 0.0, "USD", "Diagnóstico"),
        ]:
            cursor.execute(
                "INSERT INTO service_templates (description, default_price, currency, category) VALUES (?, ?, ?, ?)",
                (desc, price, curr, cat),
            )

    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            """
            INSERT INTO users (username, password, role, first_name, last_name, phone, show_in_pdf)
            VALUES ('admin', ?, 'Administrador', 'Administrador', 'Principal', '0000-0000000', 1)
        """,
            (generate_password_hash('admin123'),),
        )

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vehicles (
            id INTEGER PRIMARY KEY AUTOINCREMENT, client_doc TEXT NOT NULL,
            plate TEXT DEFAULT '', name TEXT DEFAULT '', brand TEXT DEFAULT '',
            model TEXT DEFAULT '', year TEXT DEFAULT '', color TEXT DEFAULT '',
            vin TEXT DEFAULT '', km TEXT DEFAULT '', notes TEXT DEFAULT '',
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        )
    """)
    for col, defn in [
        ('client_doc', "TEXT DEFAULT ''"),
        ('plate', "TEXT DEFAULT ''"),
        ('name', "TEXT DEFAULT ''"),
        ('brand', "TEXT DEFAULT ''"),
        ('model', "TEXT DEFAULT ''"),
        ('year', "TEXT DEFAULT ''"),
        ('color', "TEXT DEFAULT ''"),
        ('vin', "TEXT DEFAULT ''"),
        ('km', "TEXT DEFAULT ''"),
        ('notes', "TEXT DEFAULT ''"),
        ('created_at', "TEXT DEFAULT ''"),
        ('updated_at', "TEXT DEFAULT ''"),
    ]:
        ensure('vehicles', col, defn)

    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_vehicles_client ON vehicles(client_doc)"
    )
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_vehicles_plate ON vehicles(plate)")

    ensure('saved_budgets', 'vehicle_id', 'INTEGER DEFAULT NULL')
    ensure('revisions', 'vehicle_id', 'INTEGER DEFAULT NULL')
    ensure('vehicles', 'public_token', "TEXT DEFAULT ''")
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_vehicles_token ON vehicles(public_token)"
    )
    # ===== Fotos de Recepción =====
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reception_photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            reception_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            original_name TEXT DEFAULT '',
            uploaded_by TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_reception_photos_rid ON reception_photos(reception_id)"
    )

    try:
        rows_no_token = cursor.execute(
            "SELECT id FROM vehicles WHERE COALESCE(public_token,'') = ''"
        ).fetchall()
        for r in rows_no_token:
            cursor.execute(
                "UPDATE vehicles SET public_token = ? WHERE id = ?",
                (generate_public_token(), r['id']),
            )
        if rows_no_token:
            print(
                f"[init_db] Tokens QR generados para {len(rows_no_token)} vehículo(s)"
            )
    except Exception as e:
        print(f"[init_db] Backfill tokens QR: {e}")

    try:
        cursor.execute("SELECT COUNT(*) FROM vehicles")
        if cursor.fetchone()[0] == 0:
            legacy_rows = cursor.execute("""
                SELECT client_doc, vehicle_name, vehicle_year, vehicle_plate,
                       vehicle_brand, vehicle_color, vehicle_model
                FROM clients
                WHERE COALESCE(vehicle_name,'') != '' OR COALESCE(vehicle_plate,'') != '' OR COALESCE(vehicle_brand,'') != ''
            """).fetchall()
            now_seed = get_ve_time().strftime("%Y-%m-%d %H:%M")
            for r in legacy_rows:
                cursor.execute(
                    """
                    INSERT INTO vehicles (client_doc, plate, name, brand, model, year, color, created_at, updated_at, public_token)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        r['client_doc'],
                        (r['vehicle_plate'] or '').upper(),
                        r['vehicle_name'] or '',
                        r['vehicle_brand'] or '',
                        r['vehicle_model'] or '',
                        r['vehicle_year'] or '',
                        r['vehicle_color'] or '',
                        now_seed,
                        now_seed,
                        generate_public_token(),
                    ),
                )
    except Exception as e:
        print(f"[init_db] Backfill: {e}")

    # ===== TABLERO TALLER (KANBAN) =====
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_columns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            slug TEXT UNIQUE NOT NULL,
            title TEXT NOT NULL,
            icon TEXT DEFAULT '',
            color TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_cards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            column_id INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            client_doc TEXT DEFAULT '',
            client_name TEXT DEFAULT '',
            client_phone TEXT DEFAULT '',
            vehicle_plate TEXT DEFAULT '',
            vehicle_name TEXT DEFAULT '',
            vehicle_brand TEXT DEFAULT '',
            vehicle_model TEXT DEFAULT '',
            vehicle_year TEXT DEFAULT '',
            assigned_to TEXT DEFAULT '',
            priority TEXT DEFAULT 'Media',
            entry_date TEXT DEFAULT '',
            due_date TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT DEFAULT '',
            created_by TEXT DEFAULT '',
            archived INTEGER DEFAULT 0
        )
    """)
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_workshop_cards_col ON workshop_cards(column_id, sort_order)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_workshop_cards_plate ON workshop_cards(vehicle_plate)"
    )

    # Asegurar columnas nuevas en workshop_cards
    ensure('workshop_cards', 'is_delivered', 'INTEGER DEFAULT 0')
    ensure('workshop_cards', 'labels_json', "TEXT DEFAULT '[]'")
    ensure('workshop_cards', 'checklist_json', "TEXT DEFAULT '[]'")
    ensure('workshop_cards', 'description', "TEXT DEFAULT ''")
    ensure('workshop_cards', 'title', "TEXT DEFAULT ''")
    ensure('workshop_cards', 'status', "TEXT DEFAULT 'taller'")
    ensure('workshop_cards', 'delivered_at', "TEXT DEFAULT ''")
    ensure(
        'workshop_cards', 'repair_status', "TEXT DEFAULT 'Espera de Servicio Técnico'"
    )

    # Migración: si hay tarjetas con is_delivered=1 sin status, marcarlas como 'entregado'
    try:
        cursor.execute("""
            UPDATE workshop_cards SET status = 'entregado'
            WHERE is_delivered = 1 AND (status IS NULL OR status = 'taller')
        """)
        if cursor.rowcount > 0:
            print(f"[init_db] {cursor.rowcount} tarjetas migradas a status='entregado'")
    except Exception as e:
        print(f"[init_db] Migración status: {e}")

    # ===== Comentarios y actividad del tablero =====
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_card_comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            card_id INTEGER NOT NULL,
            username TEXT NOT NULL DEFAULT '',
            type TEXT DEFAULT 'comment',
            content TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_wcc_card ON workshop_card_comments(card_id, id DESC)"
    )

    # ===== Migración: eliminar FK de workshop_cards (ya no usamos columnas reales) =====
    try:
        fk_rows = cursor.execute("PRAGMA foreign_key_list(workshop_cards)").fetchall()
        if fk_rows:
            print("[init_db] Migrando workshop_cards para remover FK...")
            cursor.execute("ALTER TABLE workshop_cards RENAME TO workshop_cards_old")
            cursor.execute("""
                CREATE TABLE workshop_cards (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    column_id INTEGER DEFAULT 0,
                    sort_order INTEGER DEFAULT 0,
                    client_doc TEXT DEFAULT '',
                    client_name TEXT DEFAULT '',
                    client_phone TEXT DEFAULT '',
                    vehicle_plate TEXT DEFAULT '',
                    vehicle_name TEXT DEFAULT '',
                    vehicle_brand TEXT DEFAULT '',
                    vehicle_model TEXT DEFAULT '',
                    vehicle_year TEXT DEFAULT '',
                    assigned_to TEXT DEFAULT '',
                    priority TEXT DEFAULT 'Media',
                    entry_date TEXT DEFAULT '',
                    due_date TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT DEFAULT '',
                    created_by TEXT DEFAULT '',
                    archived INTEGER DEFAULT 0,
                    is_delivered INTEGER DEFAULT 0,
                    labels_json TEXT DEFAULT '[]',
                    checklist_json TEXT DEFAULT '[]',
                    description TEXT DEFAULT '',
                    title TEXT DEFAULT '',
                    status TEXT DEFAULT 'taller'
                )
            """)
            cursor.execute("""
                INSERT INTO workshop_cards (
                    id, column_id, sort_order, client_doc, client_name, client_phone,
                    vehicle_plate, vehicle_name, vehicle_brand, vehicle_model, vehicle_year,
                    assigned_to, priority, entry_date, due_date, notes,
                    created_at, updated_at, created_by, archived,
                    is_delivered, labels_json, checklist_json, description, title, status
                )
                SELECT id, 0, sort_order, client_doc, client_name, client_phone,
                       vehicle_plate, vehicle_name, vehicle_brand, vehicle_model, vehicle_year,
                       assigned_to, priority, entry_date, due_date, notes,
                       created_at, updated_at, created_by, COALESCE(archived, 0),
                       COALESCE(is_delivered, 0),
                       COALESCE(labels_json, '[]'),
                       COALESCE(checklist_json, '[]'),
                       COALESCE(description, ''),
                       COALESCE(title, ''),
                       COALESCE(status, 'taller')
                FROM workshop_cards_old
            """)
            cursor.execute("DROP TABLE workshop_cards_old")
            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_workshop_cards_col ON workshop_cards(assigned_to, is_delivered, sort_order)"
            )
            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_workshop_cards_plate ON workshop_cards(vehicle_plate)"
            )
            print("[init_db] ✓ workshop_cards migrada sin FK")
    except Exception as e:
        print(f"[init_db] Migración workshop_cards: {e}")

    # ===== Migración: renombrar módulo "📂 Presupuestos" → "📂 Presupuestos" =====
    try:
        old_roles = cursor.execute(
            "SELECT DISTINCT role_name FROM role_permissions WHERE module_name = '📂 Historial'"
        ).fetchall()
        for row in old_roles:
            cursor.execute(
                "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES (?, '📂 Presupuestos')",
                (row[0],),
            )
        cursor.execute(
            "DELETE FROM role_permissions WHERE module_name = '📂 Historial'"
        )
        if old_roles:
            print(
                f"[init_db] Módulo renombrado a Presupuestos en {len(old_roles)} rol(es)"
            )
    except Exception as e:
        print(f"[init_db] Migración permisos Presupuestos: {e}")

    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_budgets_client ON saved_budgets(client_doc)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_budgets_date ON saved_budgets(date DESC)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_revisions_client ON revisions(client_doc)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_revisions_date ON revisions(date DESC)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_payments_budget ON payments(budget_id)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_wc_board ON workshop_cards(archived, status, assigned_to, sort_order)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_sb_almacen ON saved_budgets(almacen_status, status)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_sb_user_status ON saved_budgets(username, status, date DESC)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_rev_mech_status ON revisions(mechanic_status, mechanic_date)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_ci_status ON comanda_items(comanda_id, status)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_co_item ON comanda_ofertas(item_id, status)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_pay_date ON payments(budget_id, date)"
    )

    conn.commit()
    conn.close()


try:
    init_db()
    purge_old_reminders()
except Exception as e:
    import traceback

    print(f"[init_db] FALLO: {e}")
    traceback.print_exc()
    raise

try:
    migrate_plaintext_passwords()
except Exception as e:
    print(f"[migrate] FALLO: {e}")


def get_config(key):
    conn = get_db_connection()
    row = conn.execute("SELECT value FROM config WHERE key = ?", (key,)).fetchone()
    conn.close()
    return float(row['value']) if row else 0.0


def set_config(key, value):
    conn = get_db_connection()
    conn.execute("UPDATE config SET value = ? WHERE key = ?", (float(value), key))
    conn.commit()
    conn.close()


_rates_cache = {"ts": 0, "usd": None, "eur": None}
RATES_TTL = 900


def fetch_bcv_rates(force=False):
    if not _HAS_REQUESTS:
        return None, None
    now = time.time()
    if not force and (now - _rates_cache["ts"]) < RATES_TTL:
        return _rates_cache["usd"], _rates_cache["eur"]
    usd, eur = None, None
    try:
        r_usd = requests.get("https://ve.dolarapi.com/v1/dolares/oficial", timeout=3)
        if r_usd.status_code == 200:
            usd = round(float(r_usd.json().get("promedio")), 2)
    except Exception:
        pass
    try:
        r_eur = requests.get("https://ve.dolarapi.com/v1/euros/oficial", timeout=3)
        if r_eur.status_code == 200:
            eur = round(float(r_eur.json().get("promedio")), 2)
    except Exception:
        pass
    if usd is not None or eur is not None:
        _rates_cache.update({"ts": now, "usd": usd, "eur": eur})
    return usd, eur


class PDF(FPDF):
    def header(self):
        pass


def safe_txt(txt):
    if not txt:
        return ""
    return str(txt).encode('latin-1', 'replace').decode('latin-1')


def pdf_to_bytes(pdf):
    try:
        out = pdf.output(dest="S")
    except TypeError:
        out = pdf.output()
    if isinstance(out, bytes):
        return out
    if isinstance(out, bytearray):
        return bytes(out)
    if isinstance(out, str):
        return out.encode('latin-1', 'replace')
    return bytes(out)


def get_budget_number(conn, budget_id, date_str):
    """Calcula un número secuencial anual para la cotización: COT-YYYY-XXXX"""
    year = date_str[:4]
    row = conn.execute(
        """
        SELECT COUNT(*) as count
        FROM saved_budgets
        WHERE substr(date, 1, 4) = ? AND id <= ?
    """,
        (year, budget_id),
    ).fetchone()
    seq = row['count'] if row else 1
    return f"COT-{year}-{seq:04d}"


def generate_pdf(
    client,
    doc,
    phone,
    veh_name,
    veh_year,
    veh_plate,
    date,
    cart,
    currency,
    rate_bcv,
    total_usd,
    total_ves,
    budget_note,
    username="",
    base_imponible=0.0,
    include_iva=False,
    iva_amount=0.0,
    discount_amount=0.0,
    discount_percent=0.0,
    payments=None,
    payment_status='',
    signature_b64='',
    signature_date='',
    veh_brand='',
    veh_color='',
    veh_model='',
    budget_status='Pendiente',
    show_bs_equivalent=False,
    budget_number='',
):  # <-- NUEVO PARÁMETRO AQUÍ
    pdf = PDF()
    pdf.set_title(safe_txt(f"Presupuesto - {client}"))
    pdf.add_page()

    logo_path = os.path.join(BASE_DIR, "static", "img", "logopdf.png")
    if os.path.exists(logo_path):
        try:
            pdf.image(logo_path, 12, 5, 22)
        except Exception:
            pass

    pdf.set_font('Arial', 'B', 9)
    pdf.set_xy(42, 8)
    pdf.cell(80, 4, safe_txt("RUEDAS Y SUSPENSIONES "), 0, 1, 'L')
    pdf.set_font('Arial', '', 7.5)
    for line in [
        "RIF: J-40083124-5",
        "Dirección: Av. Bolívar Norte, Local cívico 123-31. Valencia Edo. Carabobo",
        "Teléfonos: (0241) 825.60.24 / 825.39.25",
        "Correo: ruedasysuspensionesgroup@hotmail.com",
    ]:
        pdf.set_x(42)
        pdf.cell(80, 3.5, safe_txt(line), 0, 1, 'L')

    show_advisor = False
    advisor_text = advisor_phone_text = ""
    if username:
        user = get_user_by_username(username)
        if user and user['show_in_pdf']:
            full_name = f"{user['first_name']} {user['last_name']}".strip()
            advisor_text = full_name if full_name else user['username']
            advisor_phone_text = user['phone'] or ""
            show_advisor = True

    # ===== NÚMERO DE COTIZACIÓN (NUEVO) =====
    pdf.set_xy(125, 8)
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(200, 0, 0)  # Color rojo para destacar el número
    pdf.cell(75, 5, safe_txt(f"COTIZACIÓN N° {budget_number}"), 0, 1, 'R')
    pdf.set_text_color(0, 0, 0)  # Volver a negro

    pdf.set_xy(125, 14)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(75, 4, safe_txt(f"FECHA: {date}"), 0, 1, 'R')
    pdf.set_xy(125, 18)
    pdf.cell(75, 4, safe_txt(f"CLIENTE: {client}"), 0, 1, 'R')
    pdf.set_xy(125, 22)
    pdf.set_font('Arial', '', 7.5)
    pdf.cell(75, 3.5, safe_txt(f"C.I / RIF: {doc}"), 0, 1, 'R')
    pdf.set_xy(125, 25.5)
    pdf.cell(75, 3.5, safe_txt(f"TELÉFONO CLIENTE: {phone}"), 0, 1, 'R')
    if show_advisor:
        pdf.set_xy(125, 29)
        pdf.cell(75, 3.5, safe_txt(f"ASESOR: {advisor_text}"), 0, 1, 'R')
        if advisor_phone_text:
            pdf.set_xy(125, 32.5)
            pdf.cell(
                75, 3.5, safe_txt(f"TELÉFONO ASESOR: {advisor_phone_text}"), 0, 1, 'R'
            )

    # ===== Banner de anulado =====
    if budget_status == 'Anulado':
        pdf.set_font('Arial', 'B', 14)
        pdf.set_text_color(200, 0, 0)
        pdf.cell(0, 8, safe_txt("*** PRESUPUESTO ANULADO ***"), 0, 1, 'C')
        pdf.set_text_color(0, 0, 0)
        pdf.ln(2)

    pdf.set_y(38)  # <-- Cambiado de 35 a 38 para que no choque con los datos

    if veh_name or veh_year or veh_plate or veh_brand or veh_model or veh_color:
        pdf.set_font('Arial', 'B', 8)
        veh_line = (
            f"Vehículo: {(veh_brand or '').strip()} {(veh_model or '').strip()}".strip()
        )
        if veh_name:
            veh_line += f" | Desc: {veh_name}"
        veh_line += f" | Año: {veh_year or '-'} | Color: {veh_color or '-'} | Placa: {veh_plate or '-'}"
        pdf.cell(0, 5, safe_txt(veh_line), 0, 1, 'L')
    pdf.ln(2)

    W_CODE = 22
    W_DESC = 78
    W_QTY = 15
    W_PRICE = 37.5
    W_TOTAL = 37.5
    LINE_H = 4.2
    PAD_Y = 1.6
    PAD_X = 1.2

    pdf.set_font('Arial', 'B', 9)
    pdf.cell(W_CODE, 7, safe_txt('Código'), 1, 0, 'C')
    pdf.cell(W_DESC, 7, safe_txt('Descripción'), 1, 0, 'L')
    pdf.cell(W_QTY, 7, safe_txt('Cant'), 1, 0, 'C')
    pdf.cell(W_PRICE, 7, safe_txt(f'P. Unit ({currency})'), 1, 0, 'R')
    pdf.cell(W_TOTAL, 7, safe_txt(f'Total ({currency})'), 1, 1, 'R')

    pdf.set_font('Arial', '', 8)
    subtotal_gross = 0.0
    left_x = pdf.l_margin
    inner_desc_w = W_DESC - 2 * PAD_X

    def _wrap_text_measured(text, max_w):
        text = str(text or '').replace('\n', ' ')
        words = text.split(' ')
        lines = []
        current = ''
        for w in words:
            if w == '':
                continue
            test = (current + ' ' + w) if current else w
            if pdf.get_string_width(test) <= max_w:
                current = test
            else:
                if current:
                    lines.append(current)
                if pdf.get_string_width(w) > max_w:
                    part = ''
                    for ch in w:
                        if pdf.get_string_width(part + ch) <= max_w:
                            part += ch
                        else:
                            if part:
                                lines.append(part)
                            part = ch
                    current = part
                else:
                    current = w
        if current:
            lines.append(current)
        return lines or ['']

    for item in cart:
        recalc_cart_item(item)
        code_str = str(item.get('code', 'N/A'))
        name_display = str(item['name'])
        disc_pct = float(item.get('discount', 0) or 0)
        if disc_pct > 0:
            name_display = f"{name_display}  [Desc {disc_pct:.1f}%]"

        wrapped_lines = _wrap_text_measured(name_display, inner_desc_w)[:6]
        lines_needed = len(wrapped_lines)
        row_h = max(7.0, lines_needed * LINE_H + 2 * PAD_Y)

        if pdf.get_y() + row_h > pdf.h - pdf.b_margin - 5:
            pdf.add_page()
            pdf.set_font('Arial', '', 8)

        y0 = pdf.get_y()

        x_code = left_x
        x_desc = x_code + W_CODE
        x_qty = x_desc + W_DESC
        x_price = x_qty + W_QTY
        x_total = x_price + W_PRICE

        pdf.rect(x_code, y0, W_CODE, row_h)
        pdf.rect(x_desc, y0, W_DESC, row_h)
        pdf.rect(x_qty, y0, W_QTY, row_h)
        pdf.rect(x_price, y0, W_PRICE, row_h)
        pdf.rect(x_total, y0, W_TOTAL, row_h)

        y_mid = y0 + (row_h - LINE_H) / 2.0

        pdf.set_xy(x_code, y_mid)
        pdf.cell(W_CODE, LINE_H, safe_txt(code_str[:12]), 0, 0, 'C')

        for i, line_txt in enumerate(wrapped_lines):
            pdf.set_xy(x_desc + PAD_X, y0 + PAD_Y + i * LINE_H)
            pdf.cell(inner_desc_w, LINE_H, safe_txt(line_txt), 0, 0, 'L')

        pdf.set_xy(x_qty, y_mid)
        pdf.cell(W_QTY, LINE_H, str(item['qty']), 0, 0, 'C')

        pdf.set_xy(x_price, y_mid)
        pdf.cell(W_PRICE, LINE_H, f"{float(item['unit_price']):.2f}", 0, 0, 'R')

        pdf.set_xy(x_total, y_mid)
        pdf.cell(W_TOTAL, LINE_H, f"{float(item['total']):.2f}", 0, 0, 'R')

        pdf.set_xy(left_x, y0 + row_h)

        subtotal_gross += float(
            item.get('gross_total', item['unit_price'] * item['qty'])
        )

    pdf.ln(4)
    pdf.set_font('Arial', 'I', 8)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(
        0,
        4,
        safe_txt(
            'Nota: Presupuesto sujeto a disponibilidad y cambios sin previo aviso.'
        ),
        0,
        'L',
    )
    if budget_note:
        pdf.ln(2)
        pdf.multi_cell(0, 4, safe_txt(f"Observación: {budget_note}"), 0, 'L')

    pdf.set_text_color(0, 0, 0)
    pdf.ln(6)
    pdf.set_font('Arial', '', 9)
    pdf.cell(
        0,
        5,
        safe_txt(f"Subtotal Bruto: {float(subtotal_gross):.2f} {currency}"),
        0,
        1,
        'R',
    )
    if discount_amount > 0:
        pdf.cell(
            0,
            5,
            safe_txt(f"Descuentos: -{float(discount_amount):.2f} {currency}"),
            0,
            1,
            'R',
        )
    pdf.cell(
        0,
        5,
        safe_txt(
            f"Base Imponible: {float(base_imponible if base_imponible > 0 else total_usd):.2f} {currency}"
        ),
        0,
        1,
        'R',
    )
    if include_iva:
        pdf.cell(
            0, 5, safe_txt(f"IVA (16%): {float(iva_amount):.2f} {currency}"), 0, 1, 'R'
        )
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 7, safe_txt(f"TOTAL: {float(total_usd):.2f} {currency}"), 0, 1, 'R')
    if show_bs_equivalent:
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(
            0,
            6,
            safe_txt(
                f"EQUIVALENTE EN BS (Tasa BCV: {float(rate_bcv):.2f}): {float(total_ves):.2f} Bs."
            ),
            0,
            1,
            'R',
        )

    if payments:
        pdf.ln(6)
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 5, safe_txt(f"ABONOS REGISTRADOS ({len(payments)}):"), 0, 1)
        pdf.set_font('Arial', '', 8)
        total_paid = 0.0
        for p in payments:
            orig_amt = (
                float(p['original_amount'])
                if 'original_amount' in p.keys() and p['original_amount']
                else float(p['amount_ves'])
            )
            if p['currency'] == currency:
                line = f"{p['date']} — {p['method']} — {float(p['amount_foreign']):.2f} {currency}"
            else:
                line = (
                    f"{p['date']} — {p['method']} — {float(p['amount_foreign']):.2f} {currency} "
                    f"(abonó {orig_amt:.2f} {p['currency']} @ {float(p['rate_bcv']):.2f})"
                )
            if p['reference']:
                line += f" (Ref: {p['reference']})"
            pdf.cell(0, 4, safe_txt(line), 0, 1)
            total_paid += float(p['amount_foreign'])
        pdf.set_font('Arial', 'B', 9)
        remaining = float(total_usd) - total_paid
        pdf.cell(
            0, 5, safe_txt(f"Total abonado: {total_paid:.2f} {currency}"), 0, 1, 'R'
        )
        if budget_status == 'Anulado':
            pdf.set_text_color(150, 150, 150)
            pdf.cell(0, 5, safe_txt("— Presupuesto anulado —"), 0, 1, 'R')
            pdf.set_text_color(0, 0, 0)
        elif remaining > 0.01:
            pdf.set_text_color(200, 0, 0)
            pdf.cell(
                0,
                5,
                safe_txt(f"SALDO PENDIENTE: {remaining:.2f} {currency}"),
                0,
                1,
                'R',
            )
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_text_color(0, 150, 0)
            pdf.cell(0, 5, safe_txt("✔ PAGADO EN SU TOTALIDAD"), 0, 1, 'R')
            pdf.set_text_color(0, 0, 0)
    elif payment_status == 'Pagado':
        pdf.ln(4)
        pdf.set_font('Arial', 'B', 10)
        pdf.set_text_color(0, 150, 0)
        pdf.cell(0, 5, safe_txt("✔ PAGADO EN SU TOTALIDAD"), 0, 1, 'R')
        pdf.set_text_color(0, 0, 0)

    if signature_b64:
        tmp_path = None
        try:
            _, b64data = (
                signature_b64.split(",", 1)
                if "," in signature_b64
                else ("", signature_b64)
            )
            sig_bytes = base64.b64decode(b64data)

            # Usar NamedTemporaryFile para generar un nombre único e irrepetible
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                tmp_file.write(sig_bytes)
                tmp_path = tmp_file.name

            pdf.ln(8)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(0, 5, safe_txt("FIRMA DEL CLIENTE:"), 0, 1, 'L')
            pdf.image(tmp_path, x=15, w=70)

            if signature_date:
                pdf.set_font('Arial', '', 8)
                pdf.cell(0, 4, safe_txt(f"Firmado: {signature_date}"), 0, 1, 'L')

        except Exception as e:
            print(f"[generate_pdf] Error procesando la firma: {e}")
        finally:
            # El bloque finally GARANTIZA que el archivo se borre SIEMPRE,
            # incluso si pdf.image() falla o hay un error inesperado.
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    return pdf_to_bytes(pdf)


def generate_revision_pdf(
    client,
    doc,
    phone,
    veh_name,
    veh_year,
    veh_plate,
    date,
    items,
    note,
    username="",
    veh_km="",
    veh_brand="",
    veh_color="",
    veh_model="",
):
    pdf = PDF()
    pdf.set_title(safe_txt(f"Hoja de Revisión - {client}"))
    pdf.add_page()
    logo_path = os.path.join(BASE_DIR, "static", "img", "logopdf.png")
    if os.path.exists(logo_path):
        try:
            pdf.image(logo_path, 12, 5, 22)
        except Exception:
            pass
    pdf.set_font('Arial', 'B', 10)
    pdf.set_xy(42, 8)
    pdf.cell(80, 4, safe_txt("RUEDAS Y SUSPENSIONES "), 0, 1, 'L')
    pdf.set_font('Arial', '', 7.5)
    pdf.set_x(42)
    pdf.cell(80, 3.5, safe_txt("RIF: J-40083124-5"), 0, 1, 'L')
    pdf.set_x(42)
    pdf.cell(80, 3.5, safe_txt("HOJA DE REVISIÓN Y DIAGNÓSTICO TÉCNICO"), 0, 1, 'L')
    pdf.set_xy(125, 8)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(75, 4, safe_txt(f"FECHA: {date}"), 0, 1, 'R')
    pdf.set_xy(125, 12)
    pdf.cell(75, 4, safe_txt(f"CLIENTE: {client}"), 0, 1, 'R')
    pdf.set_xy(125, 16)
    pdf.set_font('Arial', '', 7.5)
    pdf.cell(75, 3.5, safe_txt(f"C.I / RIF: {doc}"), 0, 1, 'R')
    pdf.set_xy(125, 19.5)
    pdf.cell(75, 3.5, safe_txt(f"TELÉFONO: {phone}"), 0, 1, 'R')
    pdf.set_xy(125, 23)
    pdf.cell(75, 3.5, safe_txt(f"INSPECTOR: {username}"), 0, 1, 'R')
    pdf.set_y(35)
    pdf.set_font('Arial', 'B', 9)
    km_txt = f" | Km: {veh_km}" if veh_km else ""
    veh_line = (
        f"Vehículo: {(veh_brand or '').strip()} {(veh_model or '').strip()}".strip()
    )
    if veh_name:
        veh_line += f" | Desc: {veh_name}"
    veh_line += f" | Año: {veh_year or '-'} | Color: {veh_color or '-'} | Placa: {veh_plate or '-'}{km_txt}"
    pdf.cell(0, 5, safe_txt(veh_line), 0, 1, 'L')
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(60, 7, safe_txt('Sistema / Componente'), 1, 0, 'L')
    pdf.cell(35, 7, safe_txt('Estado'), 1, 0, 'C')
    pdf.cell(95, 7, safe_txt('Observación / Diagnóstico'), 1, 1, 'L')
    pdf.set_font('Arial', '', 8)
    for item in items:
        pdf.cell(60, 7, safe_txt(str(item.get('system', ''))[:35]), 1, 0, 'L')
        pdf.cell(35, 7, safe_txt(str(item.get('status', ''))[:20]), 1, 0, 'C')
        pdf.cell(95, 7, safe_txt(str(item.get('obs', ''))[:60]), 1, 1, 'L')
    if note:
        pdf.ln(4)
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 5, safe_txt("Observaciones Generales:"), 0, 1, 'L')
        pdf.set_font('Arial', '', 8)
        pdf.multi_cell(0, 4, safe_txt(note), 0, 'L')
    return pdf_to_bytes(pdf)


def ves_to_str(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def generate_filtered_report_pdf(
    title_str, budgets, total_usd, total_eur, total_ves, users_map
):
    pdf = PDF()
    pdf.set_title(safe_txt(title_str))
    pdf.add_page()
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 8, safe_txt(title_str), 0, 1, 'C')
    pdf.ln(4)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(28, 7, safe_txt('Fecha'), 1)
    pdf.cell(38, 7, safe_txt('Cliente'), 1)
    pdf.cell(32, 7, safe_txt('Asesor / Vendedor'), 1)
    pdf.cell(15, 7, safe_txt('Moneda'), 1, 0, 'C')
    pdf.cell(20, 7, safe_txt('Tasa BCV'), 1, 0, 'R')
    pdf.cell(30, 7, safe_txt('Total Vendido'), 1, 0, 'R')
    pdf.cell(27, 7, safe_txt('Estado'), 1, 1, 'C')
    pdf.set_font('Arial', '', 8)
    for b in budgets:
        user_disp = users_map.get(b['username'], b['username'])
        pdf.cell(28, 7, safe_txt(b['date']), 1)
        pdf.cell(38, 7, safe_txt(b['client_name'])[:18], 1)
        pdf.cell(32, 7, safe_txt(user_disp)[:16], 1)
        pdf.cell(15, 7, safe_txt(b['currency']), 1, 0, 'C')
        pdf.cell(20, 7, f"{float(b['bcv_rate']):.2f}", 1, 0, 'R')
        pdf.cell(30, 7, f"{float(b['total_foreign']):.2f}", 1, 0, 'R')
        pdf.cell(27, 7, safe_txt(b['status']), 1, 1, 'C')
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 6, safe_txt('RESUMEN DE VENTAS:'), 0, 1)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, safe_txt(f'Total Vendido USD: USD {float(total_usd):.2f}'), 0, 1)
    pdf.cell(0, 5, safe_txt(f'Total Vendido EUR: EUR {float(total_eur):.2f}'), 0, 1)
    pdf.cell(
        0,
        5,
        safe_txt(
            f'Total Equivalente en Bolívares: {float(ves_to_str(total_ves)):.2f} Bs.'
        ),
        0,
        1,
    )
    return pdf_to_bytes(pdf)


# ==========================================
# CIERRE DIARIO — Reporte contable por día/vendedor
# ==========================================
def compute_daily_close(conn, date_str, seller_filter, users_map):
    """
    Calcula el cierre diario completo para la fecha y vendedor indicados.
    Devuelve un diccionario con: vendedores, métodos de pago, piezas, servicios,
    recepciones y modelos de carro más frecuentes.
    """

    def _seller_default(u):
        return {
            'username': u,
            'display_name': users_map.get(u, u) if u else 'Sin asignar',
            'budgets_total': 0,
            'budgets_approved': 0,
            'budgets_pending': 0,
            'budgets_draft': 0,
            'budgets_annulled': 0,
            'total_usd': 0.0,
            'total_eur': 0.0,
            'profit_usd': 0.0,
            'profit_eur': 0.0,
            'paid_usd': 0.0,
            'paid_eur': 0.0,
            'receptions': 0,
            'cards_created': 0,
            'cards_delivered': 0,
        }

    sellers = {}

    # ===== 1. Presupuestos del día =====
    q = "SELECT * FROM saved_budgets WHERE substr(date, 1, 10) = ?"
    p = [date_str]
    if seller_filter and seller_filter != 'Todos':
        q += " AND username = ?"
        p.append(seller_filter)
    q += " ORDER BY id ASC"
    budgets_today = conn.execute(q, p).fetchall()

    for b in budgets_today:
        u = b['username'] or 'Sin asignar'
        if u not in sellers:
            sellers[u] = _seller_default(u)
        s = sellers[u]
        s['budgets_total'] += 1
        st = b['status']
        if st == 'Aprobado':
            s['budgets_approved'] += 1
            if b['currency'] == 'USD':
                s['total_usd'] += float(b['total_foreign'] or 0)
                s['profit_usd'] += float(b['total_profit_foreign'] or 0)
            elif b['currency'] == 'EUR':
                s['total_eur'] += float(b['total_foreign'] or 0)
                s['profit_eur'] += float(b['total_profit_foreign'] or 0)
        elif st == 'Pendiente':
            s['budgets_pending'] += 1
        elif st == 'Borrador':
            s['budgets_draft'] += 1
        elif st == 'Anulado':
            s['budgets_annulled'] += 1
        if b['currency'] == 'USD':
            s['paid_usd'] += float(b['total_paid_foreign'] or 0)
        elif b['currency'] == 'EUR':
            s['paid_eur'] += float(b['total_paid_foreign'] or 0)

    # ===== 2. Recepciones del día =====
    rq = """
        SELECT id, client_name, client_doc, client_phone,
               vehicle_brand, vehicle_model, vehicle_plate, vehicle_year,
               reception_status, status, username, date, reception_fecha,
               reception_hora, reception_am_pm, reception_motor, vehicle_km
        FROM revisions
        WHERE (reception_fecha = ? OR substr(date, 1, 10) = ?)
    """
    rp = [date_str, date_str]
    if seller_filter and seller_filter != 'Todos':
        rq += " AND username = ?"
        rp.append(seller_filter)
    rq += " ORDER BY id DESC"
    receptions_today = conn.execute(rq, rp).fetchall()

    for r in receptions_today:
        u = r['username'] or 'Sin asignar'
        if u not in sellers:
            sellers[u] = _seller_default(u)
        sellers[u]['receptions'] += 1

    # ===== 3. Tarjetas taller creadas hoy =====
    cq = "SELECT created_by, COUNT(*) as c FROM workshop_cards WHERE substr(created_at, 1, 10) = ?"
    cp = [date_str]
    if seller_filter and seller_filter != 'Todos':
        cq += " AND created_by = ?"
        cp.append(seller_filter)
    cq += " GROUP BY created_by"
    for row in conn.execute(cq, cp).fetchall():
        u = row['created_by'] or 'Sin asignar'
        if u not in sellers:
            sellers[u] = _seller_default(u)
        sellers[u]['cards_created'] = row['c']

    # ===== 4. Tarjetas entregadas hoy =====
    dq = """
        SELECT assigned_to, COUNT(*) as c FROM workshop_cards
        WHERE status = 'entregado'
          AND substr(COALESCE(NULLIF(delivered_at,''), updated_at), 1, 10) = ?
    """
    dp = [date_str]
    if seller_filter and seller_filter != 'Todos':
        dq += " AND assigned_to = ?"
        dp.append(seller_filter)
    dq += " GROUP BY assigned_to"
    for row in conn.execute(dq, dp).fetchall():
        u = row['assigned_to'] or 'Sin asignar'
        if u not in sellers:
            sellers[u] = _seller_default(u)
        sellers[u]['cards_delivered'] = row['c']

    # ===== 5. Métodos de pago del día =====
    pq = """
        SELECT p.method, p.currency, p.amount_foreign, p.amount_ves, p.username
        FROM payments p
        WHERE substr(p.date, 1, 10) = ?
    """
    pp = [date_str]
    if seller_filter and seller_filter != 'Todos':
        pq += " AND p.username = ?"
        pp.append(seller_filter)
    payments_today = conn.execute(pq, pp).fetchall()

    pay_agg = {}
    for pay in payments_today:
        m = pay['method'] or 'Otro'
        if m not in pay_agg:
            pay_agg[m] = {'method': m, 'count': 0, 'total_bs': 0.0, 'currencies': {}}
        pay_agg[m]['count'] += 1
        pay_agg[m]['total_bs'] += float(pay['amount_ves'] or 0)
        cur = pay['currency'] or 'USD'
        pay_agg[m]['currencies'][cur] = pay_agg[m]['currencies'].get(cur, 0) + float(
            pay['amount_foreign'] or 0
        )

    # ===== 6. Piezas y servicios vendidos (de presupuestos APROBADOS) =====
    parts = {}
    services = {}
    for b in budgets_today:
        if b['status'] != 'Aprobado':
            continue
        try:
            items = json.loads(b['items_json'] or '[]')
        except Exception:
            items = []
        for it in items:
            typ = (it.get('type') or '').upper()
            code = it.get('code', '') or ''
            name = it.get('name', '') or ''
            qty = int(it.get('qty', 0) or 0)
            total = float(it.get('total', 0) or 0)
            currency = b['currency']
            key = (code, name)
            if typ == 'PRODUCTO':
                if key not in parts:
                    parts[key] = {
                        'code': code,
                        'name': name,
                        'qty': 0,
                        'total': 0.0,
                        'currency': currency,
                    }
                parts[key]['qty'] += qty
                parts[key]['total'] += total
            elif typ == 'LABOR':
                if key not in services:
                    services[key] = {
                        'code': code,
                        'name': name,
                        'qty': 0,
                        'total': 0.0,
                        'currency': currency,
                    }
                services[key]['qty'] += qty
                services[key]['total'] += total

    # ===== 7. Modelos de carro más frecuentes (de las recepciones del día) =====
    models_agg = {}
    for r in receptions_today:
        brand = (r['vehicle_brand'] or '').strip().upper()
        model = (r['vehicle_model'] or '').strip().upper()
        if not brand and not model:
            continue
        key = (brand, model)
        if key not in models_agg:
            models_agg[key] = {'brand': brand, 'model': model, 'count': 0}
        models_agg[key]['count'] += 1

    car_models = sorted(models_agg.values(), key=lambda x: x['count'], reverse=True)

    # ===== 8. Totales generales =====
    totals = {
        'budgets': len(budgets_today),
        'approved': sum(1 for b in budgets_today if b['status'] == 'Aprobado'),
        'pending': sum(1 for b in budgets_today if b['status'] == 'Pendiente'),
        'draft': sum(1 for b in budgets_today if b['status'] == 'Borrador'),
        'annulled': sum(1 for b in budgets_today if b['status'] == 'Anulado'),
        'total_usd': sum(
            float(b['total_foreign'] or 0)
            for b in budgets_today
            if b['status'] == 'Aprobado' and b['currency'] == 'USD'
        ),
        'total_eur': sum(
            float(b['total_foreign'] or 0)
            for b in budgets_today
            if b['status'] == 'Aprobado' and b['currency'] == 'EUR'
        ),
        'profit_usd': sum(
            float(b['total_profit_foreign'] or 0)
            for b in budgets_today
            if b['status'] == 'Aprobado' and b['currency'] == 'USD'
        ),
        'profit_eur': sum(
            float(b['total_profit_foreign'] or 0)
            for b in budgets_today
            if b['status'] == 'Aprobado' and b['currency'] == 'EUR'
        ),
        'paid_usd': sum(
            float(b['total_paid_foreign'] or 0)
            for b in budgets_today
            if b['currency'] == 'USD'
        ),
        'paid_eur': sum(
            float(b['total_paid_foreign'] or 0)
            for b in budgets_today
            if b['currency'] == 'EUR'
        ),
        'receptions': len(receptions_today),
        'cards_created': sum(s['cards_created'] for s in sellers.values()),
        'cards_delivered': sum(s['cards_delivered'] for s in sellers.values()),
        'payments_count': len(payments_today),
        'payments_bs': sum(v['total_bs'] for v in pay_agg.values()),
    }

    return {
        'date': date_str,
        'seller_filter': seller_filter,
        'sellers': sorted(
            sellers.values(),
            key=lambda x: (x['total_usd'] + x['total_eur']),
            reverse=True,
        ),
        'totals': totals,
        'payment_methods': sorted(
            pay_agg.values(), key=lambda x: x['count'], reverse=True
        ),
        'parts_sold': sorted(parts.values(), key=lambda x: x['qty'], reverse=True),
        'services_sold': sorted(
            services.values(), key=lambda x: x['qty'], reverse=True
        ),
        'receptions': [dict(r) for r in receptions_today],
        'car_models': car_models,
    }


# ==========================================
# HELPERS DE RECORDATORIOS
# ==========================================
def count_unseen_reminders(conn, username):
    return conn.execute(
        """
        SELECT COUNT(*) FROM reminders
        WHERE id NOT IN (SELECT reminder_id FROM user_reminder_views WHERE username = ?)
          AND (is_private = 0 OR username = ?)
          AND (assigned_to = '' OR assigned_to = ?)
          AND archived = 0 AND is_completed = 0
    """,
        (username, username, username),
    ).fetchone()[0]


def count_today_reminders(conn, username):
    today = get_ve_time().strftime("%Y-%m-%d")
    return conn.execute(
        """
        SELECT COUNT(*) FROM reminders
        WHERE (is_private = 0 OR username = ?)
          AND (assigned_to = '' OR assigned_to = ?)
          AND archived = 0 AND is_completed = 0
          AND substr(date, 1, 10) = ?
    """,
        (username, username, today),
    ).fetchone()[0]


def reminder_to_dict(row):
    try:
        keys = row.keys()
    except Exception:
        keys = []

    def g(k, d=''):
        return row[k] if k in keys and row[k] is not None else d

    return {
        'id': row['id'],
        'username': g('username'),
        'title': g('title'),
        'content': g('content'),
        'date': g('date'),
        'priority': g('priority', 'Media'),
        'category': g('category'),
        'due_date': g('due_date'),
        'attachment': g('attachment'),
        'attachment_name': g('attachment_name'),
        'attachment_is_image': int(g('attachment_is_image', 0) or 0),
        'is_completed': int(g('is_completed', 0) or 0),
        'completed_at': g('completed_at'),
        'completed_by': g('completed_by'),
        'is_private': int(g('is_private', 0) or 0),
        'client_doc': g('client_doc'),
        'client_name': g('client_name'),
        'vehicle_plate': g('vehicle_plate'),
        'assigned_to': g('assigned_to'),
        'recurrence': g('recurrence', 'none'),
        'archived': int(g('archived', 0) or 0),
        'archived_at': g('archived_at'),
        'updated_at': g('updated_at'),
    }


# ==========================================
# RUTAS DE AUTENTICACIÓN
# ==========================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u_input = request.form.get('username', '').strip()
        p_input = request.form.get('password', '')
        conn = get_db_connection()
        user = conn.execute(
            "SELECT * FROM users WHERE username = ?", (u_input,)
        ).fetchone()
        conn.close()
        if user and check_password_hash(user['password'], p_input):
            session['logged_in'] = True
            session['username'] = user['username']
            session['role'] = user['role']
            full_name = f"{user['first_name']} {user['last_name']}".strip()
            session['user_display_name'] = full_name if full_name else user['username']
            session['user_phone'] = user['phone']
            log_action(
                'Inicio de sesión',
                f'Inicio de sesión de {user["username"]}.',
                'Usuario',
                user['id'],
            )
            flash(f'¡Bienvenido, {session["user_display_name"]}!', 'success')

            # Los roles de Servicio Técnico aterrizan directo en el Tablero Taller
            if is_tablero_taller_default_role(user['role']):
                return redirect(url_for('index', module='🔧 Tablero Taller'))

            return redirect(url_for('index'))
        flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    username = session.get('username', '')
    display_name = session.get('user_display_name', username)
    if username:
        log_action(
            'Cierre de sesión',
            f'Cierre de sesión de {display_name}.',
            'Usuario',
            username=username,
        )
    session.clear()
    if username:
        flash(f'¡Hasta pronto, {display_name}!', 'success')
    return redirect(url_for('login'))


_users_cache = {'ts': 0.0, 'map': None, 'list': None}


def get_cached_users():
    now = time.time()
    if _users_cache['map'] is not None and (now - _users_cache['ts']) < 300:
        return _users_cache['map'], _users_cache['list']
    conn = get_db_connection()
    try:
        rows = conn.execute("SELECT * FROM users").fetchall()
    finally:
        conn.close()
    users_map = {}
    users_list = []
    for u in rows:
        full = f"{u['first_name']} {u['last_name']}".strip()
        if full and u['phone']:
            users_map[u['username']] = f"{full} ({u['phone']})"
        elif full:
            users_map[u['username']] = full
        else:
            users_map[u['username']] = u['username']
        users_list.append(dict(u))
    _users_cache['map'] = users_map
    _users_cache['list'] = users_list
    _users_cache['ts'] = now
    return users_map, users_list


def invalidate_users_cache():
    _users_cache['ts'] = 0.0
    _users_cache['map'] = None
    _users_cache['list'] = None


@app.route('/', defaults={'module': '🏠 Inicio'})
@app.route('/<module>')
def index(module):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    role = session['role']
    curr_username = session['username']

    # Los roles de Servicio Técnico que entran por "/" ven el Tablero Taller
    if is_tablero_taller_default_role(role) and module == '🏠 Inicio':
        return redirect(url_for('index', module='🔧 Tablero Taller'))

    allowed = get_allowed_modules(role)
    menu_options = [m for m in ALL_SYSTEM_MODULES if m in allowed]
    if not menu_options:
        menu_options = ["📝 Nuevo Presupuesto"]
    if module not in menu_options:
        module = menu_options[0]

    if 'cart' not in session:
        session['cart'] = []
    if 'budget_currency' not in session:
        session['budget_currency'] = 'USD'
    if 'budget_client' not in session:
        session['budget_client'] = {}

    tasa_usd = get_config('tasa_usd')
    tasa_eur = get_config('tasa_eur')
    current_date = get_ve_time().strftime("%Y-%m-%d")

    f_type = request.args.get('f_type', 'Todos')
    f_status = request.args.get('f_status', 'Todos')
    f_user = request.args.get('f_user', 'Todos')
    f_date = request.args.get('f_date', current_date)
    f_month = request.args.get('f_month', get_ve_time().strftime("%Y-%m"))
    f_start = request.args.get('f_start', '')
    h_search = request.args.get('h_search', '').strip()
    f_end = request.args.get('f_end', '')
    log_user = request.args.get('log_user', 'Todos')
    log_action_filter = request.args.get('log_action', 'Todos')
    selected_client_doc = request.args.get('selected_client_doc', '')
    vehicle_search = request.args.get('vehicle_search', '').strip()
    appointment_filter = request.args.get('appointment_filter', 'Proximas')

    c_search = request.args.get('c_search', '').strip()
    c_filter = request.args.get('c_filter', 'Todos')
    c_sort = request.args.get('c_sort', 'nombre')
    c_page = safe_int(request.args.get('c_page'), 1)

    v_search = request.args.get('v_search', '').strip()
    v_client = request.args.get('v_client', '').strip()
    v_sort = request.args.get('v_sort', 'recientes')
    v_page = safe_int(request.args.get('v_page'), 1)

    r_search = request.args.get('r_search', '').strip()
    r_priority = request.args.get('r_priority', 'Todas')
    r_category = request.args.get('r_category', 'Todas')
    r_status = request.args.get('r_status', 'Pendientes')
    r_author = request.args.get('r_author', 'Todos')
    r_sort = request.args.get('r_sort', 'recientes')

    conn = get_db_connection()

    reminders_count = count_unseen_reminders(conn, curr_username)
    reminders_today_count = count_today_reminders(conn, curr_username)

    appointments_count = conn.execute(
        """
        SELECT COUNT(*) FROM appointments
        WHERE status IN ('Pendiente', 'Confirmada') AND date >= ?
    """,
        (current_date,),
    ).fetchone()[0]

    users_map, _cached_users_list = get_cached_users()
    all_users = list(users_map.keys())

    all_clients = []
    c_pagination = None
    vehicle_list = []
    v_pagination = None
    vehicle_stats = None  # ← NUEVO
    pending_budgets = []
    seller_summary = []
    filtered_budgets = []
    revision_kpis = None
    mechanic_revision_kpis = None
    mr_pagination = None
    mr_status = 'Todos'
    recent_revisions = []
    r_pagination = None
    h_pagination = None
    workshop_columns = []
    cards_by_column = {}
    workshop_stats = None
    all_users_list = []
    workshop_mechanics = []
    # ⬇️ Inicializar variables del Historial del Taller (se pasan a render_template SIEMPRE)
    wb_view = request.args.get('wb_view', 'board')
    workshop_history = []
    workshop_history_stats = None
    hw_user = request.args.get('hw_user', 'Todos')
    hw_search = request.args.get('hw_search', '').strip()
    hw_date_from = request.args.get('hw_date_from', '')
    hw_date_to = request.args.get('hw_date_to', '')
    # ===== Inventario (defaults) =====
    inventory_list = []
    inv_pagination = None
    inventory_categories = []
    inventory_stats = None
    inv_search = request.args.get('inv_search', '').strip()
    inv_category_filter = request.args.get('inv_category', 'Todas')
    inv_page = safe_int(request.args.get('inv_page'), 1)
    # ===== Expedientes (defaults) =====
    expedientes_list = []
    exp_pagination = None
    cliente_exp = None
    exp_budgets = []
    exp_revisions = []
    exp_cards = []
    exp_appointments = []
    exp_vehicles = []
    exp_timeline = []
    exp_stats = None
    exp_search = request.args.get('exp_search', '').strip()
    exp_page = safe_int(request.args.get('exp_page'), 1)

    where_clause = " WHERE 1=1"
    params = []
    if f_user and f_user != 'Todos':
        where_clause += " AND username = ?"
        params.append(f_user)
    only_drafts = request.args.get('only_drafts') == '1'
    if only_drafts:
        where_clause += " AND status = 'Borrador'"
    if f_status and f_status != 'Todos':
        where_clause += " AND status = ?"
        params.append(f_status)
    if f_type == 'Dia' and f_date:
        where_clause += " AND date LIKE ?"
        params.append(f"{f_date}%")
    elif f_type == 'Mes' and f_month:
        where_clause += " AND date LIKE ?"
        params.append(f"{f_month}%")
    elif f_type == 'Rango' and f_start and f_end:
        where_clause += " AND substr(date, 1, 10) >= ? AND substr(date, 1, 10) <= ?"
        params.extend([f_start, f_end])

    if h_search:
        search_like = f"%{h_search}%"
        where_clause += """ AND (
            client_name LIKE ? OR
            client_doc LIKE ? OR
            vehicle_plate LIKE ? OR
            vehicle_brand LIKE ? OR
            vehicle_model LIKE ? OR
            CAST(id AS TEXT) LIKE ?
        )"""
        params.extend([search_like] * 6)

    # Paginación SOLO para Historial (Ganancias y Reportes necesitan todos los datos)
    h_page = safe_int(request.args.get('h_page'), 1)
    h_pagination = None

    if module == "📂 Presupuestos":
        total_h = conn.execute(
            f"SELECT COUNT(*) FROM saved_budgets{where_clause}", params
        ).fetchone()[0]
        h_pagination = paginate(total_h, h_page, per_page=25)
        all_budgets = conn.execute(
            f"SELECT * FROM saved_budgets{where_clause} ORDER BY id DESC LIMIT ? OFFSET ?",
            params + [h_pagination['per_page'], h_pagination['offset']],
        ).fetchall()
    else:
        all_budgets = conn.execute(
            f"SELECT * FROM saved_budgets{where_clause} ORDER BY id DESC", params
        ).fetchall()

    if module == "📈 Ganancias":
        filtered_budgets = [b for b in all_budgets if b['status'] == 'Aprobado']
        pending_budgets = [b for b in all_budgets if b['status'] == 'Pendiente']
    else:
        filtered_budgets = all_budgets

    reminders = []
    reminder_categories = []
    users = []
    roles_list = []
    all_revisions = []
    purchases_history = []
    mr_is_editing = False
    mr_revision = None
    client_budgets = []
    client_revisions = []
    templates = []
    appointments = []
    report_data = None
    action_logs = []
    action_types = []

    if module == "🔔 Recordatorios":
        where = "WHERE (is_private = 0 OR username = ?) AND (assigned_to = '' OR assigned_to = ?)"
        r_params = [curr_username, curr_username]
        if r_status == 'Pendientes':
            where += " AND archived = 0 AND is_completed = 0"
        elif r_status == 'Completados':
            where += " AND archived = 0 AND is_completed = 1"
        elif r_status == 'Archivados':
            where += " AND archived = 1"
        if r_priority != 'Todas':
            where += " AND priority = ?"
            r_params.append(r_priority)
        if r_category and r_category != 'Todas':
            where += " AND category = ?"
            r_params.append(r_category)
        if r_author != 'Todos':
            where += " AND username = ?"
            r_params.append(r_author)
        if r_search:
            like = f"%{r_search}%"
            where += " AND (title LIKE ? OR content LIKE ? OR client_name LIKE ? OR vehicle_plate LIKE ?)"
            r_params.extend([like, like, like, like])

        order_map = {
            'recientes': "id DESC",
            'antiguos': "id ASC",
            'prioridad': "CASE priority WHEN 'Alta' THEN 1 WHEN 'Media' THEN 2 ELSE 3 END ASC, id DESC",
            'vencimiento': "CASE WHEN due_date = '' OR due_date IS NULL THEN 1 ELSE 0 END ASC, due_date ASC, id DESC",
        }
        order_sql = order_map.get(r_sort, order_map['recientes'])
        reminders = conn.execute(
            f"SELECT * FROM reminders {where} ORDER BY {order_sql} LIMIT 500", r_params
        ).fetchall()

        reminder_categories = [
            r['category']
            for r in conn.execute(
                "SELECT DISTINCT category FROM reminders WHERE category != '' ORDER BY category ASC"
            ).fetchall()
        ]

        try:
            conn.execute(
                """
                INSERT OR IGNORE INTO user_reminder_views (username, reminder_id)
                SELECT ?, id FROM reminders WHERE archived = 0
            """,
                (curr_username,),
            )
            conn.commit()
        except Exception:
            pass

    elif module == "👥 Usuarios":
        users = conn.execute("SELECT * FROM users").fetchall()
        roles_list = conn.execute("SELECT * FROM roles").fetchall()
    elif module == "🛡️ Roles y Permisos":
        roles_list = conn.execute("SELECT * FROM roles").fetchall()
    elif module == "📥 Recepciones":
        r_search = request.args.get('r_search', '').strip()
        r_status_filter = request.args.get('r_status', '').strip()

        where_clauses = ["1=1"]
        params_kpi = []
        params_list = []

        if r_search:
            like = f"%{r_search}%"
            where_clauses.append(
                "(client_name LIKE ? OR client_doc LIKE ? OR vehicle_plate LIKE ? OR vehicle_brand LIKE ? OR vehicle_model LIKE ? OR client_phone LIKE ?)"
            )
            params_kpi.extend([like] * 6)
            params_list.extend([like] * 6)

        if r_status_filter:
            where_clauses.append("reception_status = ?")
            params_kpi.append(r_status_filter)
            params_list.append(r_status_filter)

        base_where = " WHERE " + " AND ".join(where_clauses)

        # KPIs agregados en SQL (ahora reflejan la búsqueda)
        kpi_row = conn.execute(
            f"""
            SELECT
                COUNT(*) as total,
                COALESCE(SUM(CASE WHEN status = 'Pendiente' THEN 1 ELSE 0 END), 0) as pendientes,
                COALESCE(SUM(CASE WHEN status = 'Aprobado'  THEN 1 ELSE 0 END), 0) as aprobadas,
                COALESCE(SUM(CASE WHEN status = 'Anulado'   THEN 1 ELSE 0 END), 0) as anuladas,
                COALESCE(SUM(CASE WHEN substr(date, 1, 7) = ? THEN 1 ELSE 0 END), 0) as este_mes
            FROM revisions {base_where}
        """,
            [current_date[:7]] + params_kpi,
        ).fetchone()

        hoy_count = conn.execute(
            f"SELECT COUNT(*) FROM revisions {base_where} AND substr(date, 1, 10) = ?",
            params_kpi + [current_date],
        ).fetchone()[0]

        revision_kpis = {
            'total': kpi_row['total'] or 0,
            'pendientes': kpi_row['pendientes'] or 0,
            'aprobadas': kpi_row['aprobadas'] or 0,
            'anuladas': kpi_row['anuladas'] or 0,
            'este_mes': kpi_row['este_mes'] or 0,
            'hoy': hoy_count,
        }

        # Top 5 recientes para el panel
        recent_revisions = conn.execute(
            f"SELECT * FROM revisions {base_where} ORDER BY id DESC LIMIT 5",
            params_list,
        ).fetchall()

        # Listado paginado (25 por página)
        r_page = safe_int(request.args.get('r_page'), 1)
        r_pagination = paginate(revision_kpis['total'], r_page, per_page=25)
        all_revisions = conn.execute(
            f"SELECT * FROM revisions {base_where} ORDER BY id DESC LIMIT ? OFFSET ?",
            params_list + [r_pagination['per_page'], r_pagination['offset']],
        ).fetchall()

        all_clients = [
            dict(r)
            for r in conn.execute(
                "SELECT client_doc, client_name, client_lastname, client_phone, "
                "vehicle_name, vehicle_year, vehicle_plate, vehicle_brand, vehicle_color, vehicle_model "
                "FROM clients ORDER BY client_name ASC"
            ).fetchall()
        ]

    elif module == "🔬 Revisiones":
        mr_status = request.args.get('mr_status', 'Todos')
        mr_search = request.args.get('mr_search', '').strip()

        mr_where = "WHERE 1=1"
        mr_params = []
        if mr_status == 'Sin revisión':
            mr_where += " AND COALESCE(mechanic_date,'') = ''"
        elif mr_status == 'Pendiente':
            mr_where += " AND COALESCE(mechanic_status,'') = 'Pendiente' AND COALESCE(mechanic_date,'') != ''"
        elif mr_status == 'En Proceso':
            mr_where += " AND COALESCE(mechanic_status,'') = 'En Proceso'"
        elif mr_status == 'Completado':
            mr_where += " AND COALESCE(mechanic_status,'') = 'Completado'"

        if mr_search:
            like = f"%{mr_search}%"
            mr_where += " AND (client_name LIKE ? OR client_doc LIKE ? OR vehicle_plate LIKE ? OR vehicle_brand LIKE ? OR vehicle_model LIKE ?)"
            mr_params.extend([like] * 5)

        # KPIs del módulo (reflejan la búsqueda)
        kpi_row = conn.execute(
            f"""
            SELECT
                COUNT(*) AS total,
                COALESCE(SUM(CASE WHEN COALESCE(mechanic_date,'') != '' THEN 1 ELSE 0 END), 0) AS con_revision,
                COALESCE(SUM(CASE WHEN COALESCE(mechanic_status,'') = 'Completado' THEN 1 ELSE 0 END), 0) AS completadas,
                COALESCE(SUM(CASE WHEN substr(date, 1, 7) = ? THEN 1 ELSE 0 END), 0) AS este_mes
            FROM revisions {mr_where}
        """,
            [current_date[:7]] + mr_params,
        ).fetchone()

        sin_revision = (kpi_row['total'] or 0) - (kpi_row['con_revision'] or 0)

        mechanic_revision_kpis = {
            'total': kpi_row['total'] or 0,
            'con_revision': kpi_row['con_revision'] or 0,
            'sin_revision': sin_revision,
            'completadas': kpi_row['completadas'] or 0,
            'este_mes': kpi_row['este_mes'] or 0,
        }

        mr_page = safe_int(request.args.get('mr_page'), 1)
        total_mr = conn.execute(
            f"SELECT COUNT(*) FROM revisions {mr_where}", mr_params
        ).fetchone()[0]
        mr_pagination = paginate(total_mr, mr_page, per_page=25)
        all_revisions = conn.execute(
            f"SELECT * FROM revisions {mr_where} ORDER BY id DESC LIMIT ? OFFSET ?",
            mr_params + [mr_pagination['per_page'], mr_pagination['offset']],
        ).fetchall()

        all_clients = [
            dict(r)
            for r in conn.execute(
                "SELECT client_doc, client_name, client_lastname, client_phone, "
                "vehicle_name, vehicle_year, vehicle_plate, vehicle_brand, vehicle_color, vehicle_model "
                "FROM clients ORDER BY client_name ASC"
            ).fetchall()
        ]

    elif module == "📜 Historial Clientes":
        all_clients = [
            dict(r)
            for r in conn.execute(
                "SELECT client_doc, client_name, client_lastname, client_phone FROM clients ORDER BY client_name ASC"
            ).fetchall()
        ]
        if selected_client_doc:
            client_budgets = conn.execute(
                "SELECT * FROM saved_budgets WHERE client_doc = ? ORDER BY id DESC",
                (selected_client_doc,),
            ).fetchall()
            client_revisions = conn.execute(
                "SELECT * FROM revisions WHERE client_doc = ? ORDER BY id DESC",
                (selected_client_doc,),
            ).fetchall()
    elif module == "📝 Nuevo Presupuesto":
        templates = conn.execute(
            "SELECT * FROM service_templates ORDER BY sort_order ASC, description ASC"
        ).fetchall()
        all_clients = [
            dict(r)
            for r in conn.execute(
                "SELECT client_doc, client_name, client_lastname, client_phone, "
                "vehicle_name, vehicle_year, vehicle_plate, vehicle_brand, vehicle_color, vehicle_model "
                "FROM clients ORDER BY client_name ASC"
            ).fetchall()
        ]
    elif module == "⚙️ Plantillas":
        templates = conn.execute(
            "SELECT * FROM service_templates ORDER BY sort_order ASC, description ASC"
        ).fetchall()
    elif module == "📅 Citas":
        today = current_date
        if appointment_filter == 'Hoy':
            appointments = conn.execute(
                "SELECT * FROM appointments WHERE date = ? ORDER BY time ASC", (today,)
            ).fetchall()
        elif appointment_filter == 'Semana':
            end_week = (get_ve_time() + timedelta(days=7)).strftime("%Y-%m-%d")
            appointments = conn.execute(
                "SELECT * FROM appointments WHERE date BETWEEN ? AND ? ORDER BY date ASC, time ASC",
                (today, end_week),
            ).fetchall()
        elif appointment_filter == 'Todas':
            appointments = conn.execute(
                "SELECT * FROM appointments ORDER BY date DESC, time ASC"
            ).fetchall()
        else:
            appointments = conn.execute(
                "SELECT * FROM appointments WHERE date >= ? AND status IN ('Pendiente','Confirmada') ORDER BY date ASC, time ASC",
                (today,),
            ).fetchall()
    elif module == "🚗 Vehículos":
        where = "WHERE 1=1"
        v_params = []
        if v_search:
            like = f"%{v_search.upper()}%"
            where += """ AND (
                UPPER(COALESCE(v.plate,'')) LIKE ? OR UPPER(COALESCE(v.name,''))  LIKE ?
                OR UPPER(COALESCE(v.brand,'')) LIKE ? OR UPPER(COALESCE(v.model,'')) LIKE ?
                OR UPPER(COALESCE(v.year,''))  LIKE ? OR UPPER(COALESCE(v.color,'')) LIKE ?
                OR UPPER(COALESCE(v.vin,''))   LIKE ? OR UPPER(COALESCE(v.km,'')) LIKE ?
                OR UPPER(COALESCE(c.client_name,'')) LIKE ?
                OR UPPER(COALESCE(c.client_lastname,'')) LIKE ?
                OR UPPER(COALESCE(c.client_doc,'')) LIKE ?
                OR UPPER(COALESCE(c.client_phone,'')) LIKE ?
            )"""
            v_params.extend([like] * 12)
        if v_client:
            where += " AND v.client_doc = ?"
            v_params.append(v_client)
        order_map = {
            'recientes': "v.updated_at DESC, v.id DESC",
            'placa': "v.plate ASC",
            'marca': "v.brand ASC, v.model ASC",
            'cliente': "c.client_name ASC, v.plate ASC",
        }
        order_sql = order_map.get(v_sort, order_map['recientes'])
        total_v = conn.execute(
            f"SELECT COUNT(*) FROM vehicles v LEFT JOIN clients c ON c.client_doc = v.client_doc {where}",
            v_params,
        ).fetchone()[0]
        v_pagination = paginate(total_v, v_page, per_page=24)

        vehicle_list = conn.execute(
            f"""
            SELECT v.*, c.client_name, c.client_lastname, c.client_phone,
                (SELECT COUNT(*) FROM saved_budgets
                    WHERE client_doc = v.client_doc
                      AND UPPER(COALESCE(vehicle_plate,'')) = UPPER(COALESCE(v.plate,''))
                      AND COALESCE(v.plate,'') != '') AS n_budgets,
                (SELECT COUNT(*) FROM revisions
                    WHERE client_doc = v.client_doc
                      AND UPPER(COALESCE(vehicle_plate,'')) = UPPER(COALESCE(v.plate,''))
                      AND COALESCE(v.plate,'') != '') AS n_receptions,
                (SELECT COUNT(*) FROM workshop_cards
                    WHERE client_doc = v.client_doc
                      AND UPPER(COALESCE(vehicle_plate,'')) = UPPER(COALESCE(v.plate,''))
                      AND COALESCE(v.plate,'') != ''
                      AND archived = 0) AS n_cards
            FROM vehicles v LEFT JOIN clients c ON c.client_doc = v.client_doc
            {where} ORDER BY {order_sql} LIMIT ? OFFSET ?
        """,
            v_params + [v_pagination['per_page'], v_pagination['offset']],
        ).fetchall()

        # ===== Stats generales =====
        try:
            vehicle_stats = {
                'total': conn.execute("SELECT COUNT(*) FROM vehicles").fetchone()[0],
                'with_plate': conn.execute(
                    "SELECT COUNT(*) FROM vehicles WHERE COALESCE(plate,'') != ''"
                ).fetchone()[0],
                'brands': conn.execute(
                    "SELECT COUNT(DISTINCT UPPER(brand)) FROM vehicles WHERE COALESCE(brand,'') != ''"
                ).fetchone()[0],
                'clients': conn.execute(
                    "SELECT COUNT(DISTINCT client_doc) FROM vehicles"
                ).fetchone()[0],
            }
        except Exception:
            vehicle_stats = {'total': 0, 'with_plate': 0, 'brands': 0, 'clients': 0}

        all_clients = [
            dict(r)
            for r in conn.execute(
                "SELECT client_doc, client_name, client_lastname, client_phone FROM clients ORDER BY client_name ASC"
            ).fetchall()
        ]

    elif module == "🔧 Tablero Taller":
        _sellers = get_workshop_sellers()
        placeholders = ','.join('?' * len(_sellers))
        lower_sellers = [s.lower() for s in _sellers]
        all_users_list = conn.execute(
            f"SELECT username, first_name, last_name FROM users WHERE LOWER(username) IN ({placeholders}) ORDER BY id ASC",
            lower_sellers,
        ).fetchall()

        workshop_columns = []
        workshop_mechanics = WORKSHOP_MECHANICS

        # 2 estados por vendedor
        WORKSHOP_STATUSES = [
            {'key': 'taller', 'label': 'EN TALLER', 'icon': '🔧', 'color': '#60a5fa'},
            {
                'key': 'entregado',
                'label': 'ENTREGADO',
                'icon': '✅',
                'color': '#10b981',
            },
        ]

        for u in all_users_list:
            full_name = f"{u['first_name']} {u['last_name']}".strip() or u['username']
            for st in WORKSHOP_STATUSES:
                workshop_columns.append(
                    {
                        'id': f"{st['key']}_{u['username']}",
                        'key': st['key'],
                        'username': u['username'],
                        'title': (
                            f"{st['label']} · {full_name.upper()}"
                            if st['key'] != 'taller'
                            else full_name.upper()
                        ),
                        'status': st['key'],
                        'icon': st['icon'],
                        'color': st['color'],
                    }
                )

        # Traer todas las tarjetas activas
        all_cards_db = conn.execute("""
            SELECT * FROM workshop_cards
            WHERE archived = 0
            ORDER BY sort_order ASC, id DESC
        """).fetchall()

        plates_set = set()
        for c_row in all_cards_db:
            p = (c_row['vehicle_plate'] or '').strip()
            if p:
                plates_set.add(p.upper())

        garantia_plates = set()
        if plates_set:
            plist = list(plates_set)
            ph = ','.join('?' * len(plist))
            gr = conn.execute(
                f"SELECT DISTINCT UPPER(TRIM(vehicle_plate)) AS p FROM revisions "
                f"WHERE reception_es_retorno = 1 AND UPPER(TRIM(vehicle_plate)) IN ({ph})",
                plist,
            ).fetchall()
            garantia_plates = {r['p'] for r in gr if r['p']}

        all_cards = []
        for card_row in all_cards_db:
            card_dict = dict(card_row)
            p_clean = (card_dict.get('vehicle_plate') or '').strip().upper()
            card_dict['is_warranty'] = bool(p_clean and p_clean in garantia_plates)
            all_cards.append(card_dict)

        # 🛑 NUEVO: Filtrar tarjetas para el rol Mecánico
        role_lower = role.lower()
        if 'mecanico' in role_lower or 'mecánico' in role_lower:
            mechanic_name = session.get('user_display_name', curr_username).upper()
            mechanic_username = curr_username.upper()
            filtered_cards = []
            for c in all_cards:
                try:
                    labels = json.loads(c['labels_json'] or '[]')
                    if any(
                        mechanic_name in (lbl.get('text', '') or '').upper()
                        or mechanic_username in (lbl.get('text', '') or '').upper()
                        for lbl in labels
                    ):
                        filtered_cards.append(c)
                except Exception:
                    pass
            all_cards = filtered_cards

        # Agrupar por columna virtual usando el campo status
        cards_by_column = {}
        for col in workshop_columns:
            cards_by_column[col['id']] = []

        # Mapear usuarios válidos para arreglar tarjetas viejas con minúsculas
        valid_users = {u['username'].lower(): u['username'] for u in all_users_list}

        for card in all_cards:
            assigned = card['assigned_to'] or ''
            if assigned and assigned.lower() in valid_users:
                assigned = valid_users[assigned.lower()]

            # Determinar status
            status = 'taller'
            try:
                s = card['status']
                if s in ('taller', 'prueba', 'entregado'):
                    status = s
                elif int(card['is_delivered'] or 0) == 1:
                    status = 'entregado'
            except Exception:
                if int(card['is_delivered'] or 0) == 1:
                    status = 'entregado'

            if not assigned:
                # Tarjeta sin asignar → se envía por defecto a la columna "En Taller" de Manuel
                default_seller = (
                    all_users_list[0]['username'] if all_users_list else 'admin'
                )
                key = f"taller_{default_seller}"
                if key in cards_by_column:
                    cards_by_column[key].append(card)
                continue

            key = f"{status}_{assigned}"
            if key in cards_by_column:
                cards_by_column[key].append(card)

        def _card_status(c):
            try:
                s = c['status']
                if s in ('taller', 'entregado'):  # Eliminado 'prueba'
                    return s
            except Exception:
                pass
            return 'entregado' if int(c['is_delivered'] or 0) == 1 else 'taller'

        hoy = current_date
        workshop_stats = {
            'total': sum(1 for c in all_cards if _card_status(c) != 'entregado'),
            'alta': sum(
                1
                for c in all_cards
                if c['priority'] == 'Alta' and _card_status(c) != 'entregado'
            ),
            'vencidas': sum(
                1
                for c in all_cards
                if c['due_date']
                and c['due_date'] < hoy
                and _card_status(c) != 'entregado'
            ),
            'entregadas': sum(1 for c in all_cards if _card_status(c) == 'entregado'),
        }

        # ===== Vista Historial =====
        wb_view = request.args.get('wb_view', 'board')
        workshop_history = []
        workshop_history_stats = None
        hw_user = request.args.get('hw_user', 'Todos')
        hw_search = request.args.get('hw_search', '').strip()
        hw_date_from = request.args.get('hw_date_from', '')
        hw_date_to = request.args.get('hw_date_to', '')

        if wb_view == 'history':
            hw_where = "WHERE status = 'entregado'"
            hw_params = []
            if hw_user and hw_user != 'Todos':
                hw_where += " AND assigned_to = ?"
                hw_params.append(hw_user)
            if hw_search:
                like = f"%{hw_search.upper()}%"
                hw_where += """ AND (
                    UPPER(COALESCE(vehicle_plate,'')) LIKE ? OR
                    UPPER(COALESCE(client_name,''))   LIKE ? OR
                    UPPER(COALESCE(vehicle_brand,'')) LIKE ? OR
                    UPPER(COALESCE(vehicle_model,'')) LIKE ?
                )"""
                hw_params.extend([like] * 4)
            if hw_date_from:
                hw_where += " AND substr(COALESCE(NULLIF(delivered_at,''), updated_at), 1, 10) >= ?"
                hw_params.append(hw_date_from)
            if hw_date_to:
                hw_where += " AND substr(COALESCE(NULLIF(delivered_at,''), updated_at), 1, 10) <= ?"
                hw_params.append(hw_date_to)

            workshop_history = conn.execute(
                f"""
                SELECT * FROM workshop_cards
                {hw_where}
                ORDER BY COALESCE(NULLIF(delivered_at,''), updated_at) DESC, id DESC
                LIMIT 500
            """,
                hw_params,
            ).fetchall()

            month_key_now = current_date[:7]
            this_month = 0
            for c in workshop_history:
                d = (
                    c['delivered_at']
                    if 'delivered_at' in c.keys() and c['delivered_at']
                    else (c['updated_at'] or '')
                )
                if d[:7] == month_key_now:
                    this_month += 1

            workshop_history_stats = {
                'total': len(workshop_history),
                'this_month': this_month,
                'filtered': bool(
                    hw_user != 'Todos' or hw_search or hw_date_from or hw_date_to
                ),
            }

    elif module == "📦 Inventario":
        inv_where = "WHERE 1=1"
        inv_params = []
        if inv_search:
            # Cada término debe coincidir sin importar el orden de escritura.
            inv_words = [w for w in re.split(r'\s+', inv_search.strip().upper()) if w]
            for word in inv_words:
                like = f"%{word}%"
                inv_where += """ AND (
                    UPPER(COALESCE(code,'')) LIKE ? OR UPPER(COALESCE(name,'')) LIKE ?
                    OR UPPER(COALESCE(description,'')) LIKE ? OR UPPER(COALESCE(barcode,'')) LIKE ?
                    OR UPPER(COALESCE(brand,'')) LIKE ? OR UPPER(COALESCE(category,'')) LIKE ?
                    OR UPPER(COALESCE(supplier,'')) LIKE ? OR UPPER(COALESCE(location,'')) LIKE ?
                )"""
                inv_params.extend([like] * 8)
        if inv_category_filter and inv_category_filter != 'Todas':
            inv_where += " AND category = ?"
            inv_params.append(inv_category_filter)

        total_inv = conn.execute(
            f"SELECT COUNT(*) FROM inventory {inv_where}", inv_params
        ).fetchone()[0]
        inv_pagination = paginate(total_inv, inv_page, per_page=50)
        inventory_list = conn.execute(
            f"""
            SELECT * FROM inventory {inv_where}
            ORDER BY name COLLATE NOCASE ASC LIMIT ? OFFSET ?
        """,
            inv_params + [inv_pagination['per_page'], inv_pagination['offset']],
        ).fetchall()

        inventory_categories = [
            r['category']
            for r in conn.execute(
                "SELECT DISTINCT category FROM inventory WHERE category != '' ORDER BY category ASC"
            ).fetchall()
        ]

        inventory_stats = {
            'total': conn.execute("SELECT COUNT(*) FROM inventory").fetchone()[0],
            'stock_total': conn.execute(
                "SELECT COALESCE(SUM(stock),0) FROM inventory"
            ).fetchone()[0],
            'low_stock': conn.execute(
                "SELECT COUNT(*) FROM inventory WHERE stock > 0 AND min_stock > 0 AND stock <= min_stock"
            ).fetchone()[0],
            'no_stock': conn.execute(
                "SELECT COUNT(*) FROM inventory WHERE stock <= 0"
            ).fetchone()[0],
        }

        # Cargar proveedores para el formulario
        suppliers_list = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM suppliers ORDER BY name ASC"
            ).fetchall()
        ]

    elif module in ("📋 Comandas", "🛒 Solicitudes de Compra"):
        # 🔄 ACTUALIZACIÓN AUTOMÁTICA: Convertir estados viejos a nuevos
        try:
            conn.execute(
                "UPDATE comanda_items SET status = 'Listo para Despacho' WHERE status = 'Aprobado por Cliente'"
            )
            conn.execute(
                "UPDATE comanda_items SET status = 'Ofertas Recibidas' WHERE status = 'Oferta Ingresada'"
            )
            conn.commit()
        except Exception:
            pass

        # Usamos LEFT JOIN para traer los datos del vehículo desde el presupuesto
        comandas_db = conn.execute("""
            SELECT c.*, b.date as budget_date, b.username as seller_username,
                   b.vehicle_brand, b.vehicle_model, b.vehicle_plate, b.vehicle_year
            FROM comandas c
            LEFT JOIN saved_budgets b ON c.budget_id = b.id
            ORDER BY c.id DESC
        """).fetchall()

        comandas_list = []
        solicitudes_pendientes = []
        solicitudes_procesadas = []

        for c in comandas_db:
            productos = conn.execute(
                "SELECT * FROM comanda_items WHERE comanda_id = ?", (c['id'],)
            ).fetchall()

            # Buscar ofertas para cada producto
            productos_con_ofertas = []
            for p in productos:
                ofertas = conn.execute(
                    """
                    SELECT o.*, s.name as supplier_name
                    FROM comanda_ofertas o
                    LEFT JOIN suppliers s ON o.supplier_id = s.id
                    WHERE o.item_id = ?
                """,
                    (p['id'],),
                ).fetchall()
                p_dict = dict(p)
                p_dict['ofertas'] = [dict(o) for o in ofertas]
                productos_con_ofertas.append(p_dict)

            data_obj = {'comanda': c, 'productos': productos_con_ofertas}
            comandas_list.append(data_obj)

            # Separar para Solicitudes de Compra
            if module == "🛒 Solicitudes de Compra":
                # 1. Procesados: Son los que ya se despacharon o rechazaron
                items_procesados = [
                    p
                    for p in productos_con_ofertas
                    if p['status']
                    in ('Listo para Despacho', 'Rechazado', 'Aprobado por Cliente')
                ]

                # 2. Pendientes: CUALQUIER OTRA COSA que no esté en stock ni procesada
                items_pendientes = [
                    p
                    for p in productos_con_ofertas
                    if p['status']
                    not in (
                        'En Stock',
                        'Listo para Despacho',
                        'Rechazado',
                        'Aprobado por Cliente',
                    )
                ]

                if items_pendientes:
                    solicitudes_pendientes.append(
                        {'comanda': c, 'productos': items_pendientes}
                    )
                if items_procesados:
                    solicitudes_procesadas.append(
                        {'comanda': c, 'productos': items_procesados}
                    )

        suppliers = conn.execute("SELECT * FROM suppliers ORDER BY name ASC").fetchall()

    elif module == "🛒 Compras":
        # Traer el historial de compras con el nombre del proveedor unido
        purchases_history = conn.execute("""
            SELECT p.*, s.name as supplier_name
            FROM purchases p
            LEFT JOIN suppliers s ON p.supplier_id = s.id
            ORDER BY p.id DESC LIMIT 100
        """).fetchall()

    elif module == "📇 Clientes":
        where = "WHERE 1=1"
        c_params = []
        if c_search:
            like = f"%{c_search.upper()}%"
            where += """ AND (UPPER(COALESCE(client_name,'')) LIKE ? OR UPPER(COALESCE(client_lastname,'')) LIKE ?
                            OR UPPER(COALESCE(client_doc,'')) LIKE ? OR UPPER(COALESCE(client_phone,'')) LIKE ?)"""
            c_params.extend([like, like, like, like])
        if c_filter == 'Con vehículos':
            where += " AND EXISTS (SELECT 1 FROM vehicles v WHERE v.client_doc = clients.client_doc)"
        elif c_filter == 'Sin vehículos':
            where += " AND NOT EXISTS (SELECT 1 FROM vehicles v WHERE v.client_doc = clients.client_doc)"
        order_map = {
            'nombre': "client_name COLLATE NOCASE ASC",
            'doc': "client_doc ASC",
            'recientes': "id DESC",
        }
        order_sql = order_map.get(c_sort, order_map['nombre'])
        total_c = conn.execute(
            f"SELECT COUNT(*) FROM clients {where}", c_params
        ).fetchone()[0]
        c_pagination = paginate(total_c, c_page, per_page=25)
        rows = conn.execute(
            f"""
            SELECT clients.*,
                   (SELECT COUNT(*) FROM vehicles v WHERE v.client_doc = clients.client_doc) AS vehicle_count,
                   (SELECT MAX(updated_at) FROM vehicles v WHERE v.client_doc = clients.client_doc) AS last_vehicle_update
            FROM clients {where} ORDER BY {order_sql} LIMIT ? OFFSET ?
        """,
            c_params + [c_pagination['per_page'], c_pagination['offset']],
        ).fetchall()
        all_clients = [dict(r) for r in rows]

    elif module == "📁 Expedientes":
        exp_where = "WHERE 1=1"
        exp_params = []
        if exp_search:
            like = f"%{exp_search.upper()}%"
            exp_where += """ AND (
                UPPER(COALESCE(c.client_name,'')) LIKE ? OR
                UPPER(COALESCE(c.client_lastname,'')) LIKE ? OR
                UPPER(COALESCE(c.client_doc,'')) LIKE ? OR
                UPPER(COALESCE(c.client_phone,'')) LIKE ?
            )"""
            exp_params.extend([like, like, like, like])

        if not selected_client_doc:
            # ---------- VISTA LISTA ----------
            total_exp = conn.execute(
                f"SELECT COUNT(*) FROM clients c {exp_where}", exp_params
            ).fetchone()[0]
            exp_pagination = paginate(total_exp, exp_page, per_page=24)

            # CORRECCIÓN: Indentación alineada a 12 espacios
            rows = conn.execute(
                f"""
                SELECT c.*,
                    (SELECT COUNT(*) FROM saved_budgets WHERE client_doc = c.client_doc) AS n_budgets,
                    (SELECT COUNT(*) FROM revisions WHERE client_doc = c.client_doc) AS n_revisions,
                    (SELECT COUNT(*) FROM workshop_cards WHERE client_doc = c.client_doc AND archived = 0) AS n_active_cards,
                    (SELECT COUNT(*) FROM vehicles WHERE client_doc = c.client_doc) AS n_vehicles,
                    (SELECT COALESCE(SUM(total_foreign),0) FROM saved_budgets
                        WHERE client_doc = c.client_doc AND status='Aprobado' AND currency='USD') AS spent_usd,
                    (SELECT COALESCE(SUM(total_foreign),0) FROM saved_budgets
                        WHERE client_doc = c.client_doc AND status='Aprobado' AND currency='EUR') AS spent_eur,
                    (SELECT MAX(date) FROM saved_budgets WHERE client_doc = c.client_doc) AS last_budget_date,
                    (SELECT MAX(date) FROM revisions WHERE client_doc = c.client_doc) AS last_revision_date,
                    (SELECT MAX(updated_at) FROM vehicles WHERE client_doc = c.client_doc) AS last_vehicle_update
                FROM clients c
                {exp_where}
                ORDER BY MAX(
                    COALESCE((SELECT MAX(date) FROM saved_budgets WHERE client_doc = c.client_doc), '1970-01-01'),
                    COALESCE((SELECT MAX(date) FROM revisions WHERE client_doc = c.client_doc), '1970-01-01'),
                    COALESCE((SELECT MAX(updated_at) FROM vehicles WHERE client_doc = c.client_doc), '1970-01-01')
                ) DESC, c.id DESC
                LIMIT ? OFFSET ?
            """,
                exp_params + [exp_pagination['per_page'], exp_pagination['offset']],
            ).fetchall()

            expedientes_list = [dict(r) for r in rows]

        else:
            # ---------- VISTA DETALLE ----------
            row_cli = conn.execute(
                "SELECT * FROM clients WHERE client_doc = ?", (selected_client_doc,)
            ).fetchone()
            if row_cli:
                cliente_exp = dict(row_cli)

                exp_budgets = [
                    dict(r)
                    for r in conn.execute(
                        "SELECT * FROM saved_budgets WHERE client_doc = ? ORDER BY id DESC",
                        (selected_client_doc,),
                    ).fetchall()
                ]
                exp_revisions = [
                    dict(r)
                    for r in conn.execute(
                        "SELECT * FROM revisions WHERE client_doc = ? ORDER BY id DESC",
                        (selected_client_doc,),
                    ).fetchall()
                ]
                exp_cards = [
                    dict(r)
                    for r in conn.execute(
                        "SELECT * FROM workshop_cards WHERE client_doc = ? ORDER BY id DESC",
                        (selected_client_doc,),
                    ).fetchall()
                ]
                exp_appointments = [
                    dict(r)
                    for r in conn.execute(
                        "SELECT * FROM appointments WHERE client_doc = ? ORDER BY id DESC",
                        (selected_client_doc,),
                    ).fetchall()
                ]
                exp_vehicles = [
                    dict(r)
                    for r in conn.execute(
                        "SELECT * FROM vehicles WHERE client_doc = ? ORDER BY id DESC",
                        (selected_client_doc,),
                    ).fetchall()
                ]

                # ---- Timeline unificado ----
                exp_timeline = []
                for b in exp_budgets:
                    exp_timeline.append(
                        {
                            'kind': 'budget',
                            'sort_key': b.get('date') or '',
                            'date': b.get('date') or '',
                            'id': b['id'],
                            'status': b.get('status') or '',
                            'amount': float(b.get('total_foreign') or 0),
                            'currency': b.get('currency') or 'USD',
                            'total_paid': float(b.get('total_paid_foreign') or 0),
                            'vehicle_brand': b.get('vehicle_brand') or '',
                            'vehicle_model': b.get('vehicle_model') or '',
                            'vehicle_plate': b.get('vehicle_plate') or '',
                            'vehicle_year': b.get('vehicle_year') or '',
                            'username': b.get('username') or '',
                        }
                    )
                for r in exp_revisions:
                    exp_timeline.append(
                        {
                            'kind': 'reception',
                            'sort_key': r.get('date') or '',
                            'date': r.get('date') or '',
                            'id': r['id'],
                            'status': r.get('reception_status')
                            or r.get('status')
                            or 'Borrador',
                            'vehicle_brand': r.get('vehicle_brand') or '',
                            'vehicle_model': r.get('vehicle_model') or '',
                            'vehicle_plate': r.get('vehicle_plate') or '',
                            'vehicle_year': r.get('vehicle_year') or '',
                            'username': r.get('username') or '',
                            'has_mechanic': bool(r.get('mechanic_date')),
                        }
                    )
                for c in exp_cards:
                    exp_timeline.append(
                        {
                            'kind': 'workshop',
                            'sort_key': c.get('created_at')
                            or c.get('entry_date')
                            or '',
                            'date': c.get('entry_date')
                            or (c.get('created_at') or '')[:10],
                            'id': c['id'],
                            'status': c.get('status') or 'taller',
                            'assigned_to': c.get('assigned_to') or '',
                            'priority': c.get('priority') or 'Media',
                            'archived': int(c.get('archived') or 0),
                            'vehicle_brand': c.get('vehicle_brand') or '',
                            'vehicle_model': c.get('vehicle_model') or '',
                            'vehicle_plate': c.get('vehicle_plate') or '',
                            'vehicle_year': c.get('vehicle_year') or '',
                        }
                    )
                for a in exp_appointments:
                    exp_timeline.append(
                        {
                            'kind': 'appointment',
                            'sort_key': a.get('date') or '',
                            'date': a.get('date') or '',
                            'time': a.get('time') or '',
                            'id': a['id'],
                            'status': a.get('status') or '',
                            'service': a.get('service_description') or '',
                            'username': a.get('username') or '',
                        }
                    )

                exp_timeline.sort(key=lambda x: x.get('sort_key') or '', reverse=True)

                # ---- Stats del expediente ----
                exp_stats = {
                    'n_budgets': len(exp_budgets),
                    'n_revisions': len(exp_revisions),
                    'n_cards': len(exp_cards),
                    'n_vehicles': len(exp_vehicles),
                    'n_appointments': len(exp_appointments),
                    'spent_usd': sum(
                        float(b.get('total_foreign') or 0)
                        for b in exp_budgets
                        if b.get('status') == 'Aprobado' and b.get('currency') == 'USD'
                    ),
                    'spent_eur': sum(
                        float(b.get('total_foreign') or 0)
                        for b in exp_budgets
                        if b.get('status') == 'Aprobado' and b.get('currency') == 'EUR'
                    ),
                    'paid_usd': sum(
                        float(b.get('total_paid_foreign') or 0)
                        for b in exp_budgets
                        if b.get('currency') == 'USD'
                    ),
                    'profit_usd': sum(
                        float(b.get('total_profit_foreign') or 0)
                        for b in exp_budgets
                        if b.get('status') == 'Aprobado' and b.get('currency') == 'USD'
                    ),
                }

    elif module == "📊 Reportes":
        # ===== NUEVO: sub-vista del módulo Reportes =====
        report_view = request.args.get('report_view', 'analytics')
        daily_close = None

        if report_view == 'daily_close':
            dc_date = request.args.get('dc_date', current_date)
            dc_user = request.args.get('dc_user', 'Todos')
            try:
                daily_close = compute_daily_close(conn, dc_date, dc_user, users_map)
            except Exception as _e:
                print(f"[daily_close] {_e}")
                daily_close = None
        else:
            # ===== Lógica original de Reportes (Analytics) =====
            if f_type == 'Dia' and f_date:
                date_where, date_param = "substr(date, 1, 10) = ?", f_date
            elif f_type == 'Mes' and f_month:
                date_where, date_param = "date LIKE ?", f"{f_month}%"
            elif f_type == 'Rango' and f_start and f_end:
                date_where, date_param = "substr(date, 1, 10) BETWEEN ? AND ?", (
                    f_start,
                    f_end,
                )
            else:
                date_where = "substr(date, 1, 10) >= ?"
                date_param = (get_ve_time() - timedelta(days=365)).strftime("%Y-%m-%d")
            prod_rows = conn.execute(
                f"SELECT items_json, currency FROM saved_budgets WHERE status = 'Aprobado' AND {date_where}",
                date_param if isinstance(date_param, tuple) else (date_param,),
            ).fetchall()
            product_stats = {}
            product_margin = {}
            for row in prod_rows:
                try:
                    items = json.loads(row['items_json'])
                    for it in items:
                        key = (it.get('name', ''), it.get('code', 'N/A'))
                        if key not in product_stats:
                            product_stats[key] = {
                                'qty': 0,
                                'total': 0.0,
                                'currency': row['currency'],
                            }
                        product_stats[key]['qty'] += int(it.get('qty', 0))
                        product_stats[key]['total'] += float(it.get('total', 0))
                        profit = float(it.get('profit', 0))
                        if key not in product_margin:
                            product_margin[key] = 0.0
                        product_margin[key] += profit
                except Exception:
                    pass
            top_products = sorted(
                product_stats.items(), key=lambda x: x[1]['qty'], reverse=True
            )[:20]
            top_margins = sorted(
                product_margin.items(), key=lambda x: x[1], reverse=True
            )[:20]
            conv_rows = conn.execute(
                f"SELECT status, COUNT(*) as cnt FROM saved_budgets WHERE {date_where} GROUP BY status",
                date_param if isinstance(date_param, tuple) else (date_param,),
            ).fetchall()
            conv = {r['status']: r['cnt'] for r in conv_rows}
            aprobados = conv.get('Aprobado', 0)
            pendientes = conv.get('Pendiente', 0)
            anulados = conv.get('Anulado', 0)
            tasa_conv = (
                (aprobados / (aprobados + pendientes + anulados) * 100)
                if (aprobados + pendientes + anulados) > 0
                else 0.0
            )
            cutoff = (get_ve_time() - timedelta(days=90)).strftime("%Y-%m-%d")
            inactive_clients = conn.execute(
                """
                SELECT client_doc, client_name, client_phone, MAX(substr(date, 1, 10)) as last_date
                FROM saved_budgets WHERE status = 'Aprobado'
                GROUP BY client_doc, client_name, client_phone
                HAVING last_date < ? ORDER BY last_date ASC LIMIT 100
            """,
                (cutoff,),
            ).fetchall()
            saldo_rows = conn.execute("""
                SELECT * FROM saved_budgets WHERE status = 'Aprobado'
                  AND (total_paid_foreign IS NULL OR total_paid_foreign < total_foreign - 0.01)
                ORDER BY id DESC LIMIT 100
            """).fetchall()
            report_data = {
                'top_products': top_products,
                'top_margins': top_margins,
                'tasa_conv': tasa_conv,
                'aprobados': aprobados,
                'pendientes': pendientes,
                'anulados': anulados,
                'inactive_clients': inactive_clients,
                'saldo_rows': saldo_rows,
            }

    elif module == "📋 Historial de Acciones":
        action_types = [
            row['action']
            for row in conn.execute(
                "SELECT DISTINCT action FROM action_logs ORDER BY action ASC"
            ).fetchall()
        ]
        log_query = "SELECT * FROM action_logs WHERE 1=1"
        log_params = []
        if log_user and log_user != 'Todos':
            log_query += " AND username = ?"
            log_params.append(log_user)
        if log_action_filter and log_action_filter != 'Todos':
            log_query += " AND action = ?"
            log_params.append(log_action_filter)
        log_query += " ORDER BY id DESC LIMIT 500"
        action_logs = conn.execute(log_query, log_params).fetchall()

    # ===== Dashboard (Inicio) =====
    dashboard_metrics = None
    if module == "🏠 Inicio":
        now_ve = get_ve_time()
        h = now_ve.hour
        greeting_txt = (
            "Buenos días"
            if 5 <= h < 12
            else ("Buenas tardes" if 12 <= h < 19 else "Buenas noches")
        )
        _days_es = [
            'lunes',
            'martes',
            'miércoles',
            'jueves',
            'viernes',
            'sábado',
            'domingo',
        ]
        _months_es = [
            'enero',
            'febrero',
            'marzo',
            'abril',
            'mayo',
            'junio',
            'julio',
            'agosto',
            'septiembre',
            'octubre',
            'noviembre',
            'diciembre',
        ]
        today_formatted = f"{_days_es[now_ve.weekday()].capitalize()}, {now_ve.day} de {_months_es[now_ve.month - 1]} de {now_ve.year}"

        month_key = current_date[:7]
        month_budgets = conn.execute(
            "SELECT * FROM saved_budgets WHERE status = 'Aprobado' AND substr(date, 1, 7) = ?",
            (month_key,),
        ).fetchall()

        month_sales_usd = sum(
            float(b['total_foreign']) for b in month_budgets if b['currency'] == 'USD'
        )
        month_sales_eur = sum(
            float(b['total_foreign']) for b in month_budgets if b['currency'] == 'EUR'
        )
        month_profit_usd = sum(
            float(b['total_profit_foreign'] or 0)
            for b in month_budgets
            if b['currency'] == 'USD'
        )
        month_profit_eur = sum(
            float(b['total_profit_foreign'] or 0)
            for b in month_budgets
            if b['currency'] == 'EUR'
        )
        month_paid_usd = sum(
            float(b['total_paid_foreign'] or 0)
            for b in month_budgets
            if b['currency'] == 'USD'
        )
        month_paid_eur = sum(
            float(b['total_paid_foreign'] or 0)
            for b in month_budgets
            if b['currency'] == 'EUR'
        )

        pending_count = conn.execute(
            "SELECT COUNT(*) FROM saved_budgets WHERE status = 'Pendiente'"
        ).fetchone()[0]
        clients_count = conn.execute("SELECT COUNT(*) FROM clients").fetchone()[0]
        vehicles_count = conn.execute("SELECT COUNT(*) FROM vehicles").fetchone()[0]

        recent_budgets = conn.execute(
            "SELECT * FROM saved_budgets ORDER BY id DESC LIMIT 6"
        ).fetchall()

        today_appointments = conn.execute(
            """
            SELECT * FROM appointments WHERE date = ? AND status IN ('Pendiente', 'Confirmada')
            ORDER BY time ASC LIMIT 5
        """,
            (current_date,),
        ).fetchall()

        daily_sales = []
        for i in range(6, -1, -1):
            d = (now_ve - timedelta(days=i)).strftime("%Y-%m-%d")
            row = conn.execute(
                """
                SELECT
                    COALESCE(SUM(CASE WHEN currency='USD' THEN total_foreign ELSE 0 END), 0) as usd,
                    COALESCE(SUM(CASE WHEN currency='EUR' THEN total_foreign ELSE 0 END), 0) as eur
                FROM saved_budgets
                WHERE status='Aprobado' AND substr(date,1,10)=?
            """,
                (d,),
            ).fetchone()
            tasa_usd = get_config('tasa_usd') or 1
            tasa_eur = get_config('tasa_eur') or 1
            total_equivalent = row['usd'] + (
                row['eur'] * (tasa_eur / tasa_usd if tasa_usd > 0 else 1)
            )
            daily_sales.append({'date': d, 'total': float(total_equivalent)})

        max_sale = max([d['total'] for d in daily_sales] or [0]) or 1

        dashboard_metrics = {
            'greeting': greeting_txt,
            'today_formatted': today_formatted,
            'month_sales_usd': month_sales_usd,
            'month_sales_eur': month_sales_eur,
            'month_profit_usd': month_profit_usd,
            'month_profit_eur': month_profit_eur,
            'month_paid_usd': month_paid_usd,
            'month_paid_eur': month_paid_eur,
            'pending_count': pending_count,
            'clients_count': clients_count,
            'vehicles_count': vehicles_count,
            'recent_budgets': recent_budgets,
            'today_appointments': today_appointments,
            'daily_sales': daily_sales,
            'max_sale': max_sale,
        }

    conn.close()

    user_disp = session.get('user_display_name', curr_username)
    user_phone = session.get('user_phone', '')

    # === CONEXIÓN DIRECTA SOLICITUDES -> COMPRAS ===
    load_offer_data = None
    load_offer_item_id = request.args.get('load_offer_item')
    if module == "🛒 Compras" and load_offer_item_id:
        conn_offer = get_db_connection()
        # Buscamos el ítem de la comanda y el nombre de su proveedor
        item_off = conn_offer.execute(
            """
            SELECT ci.*, s.name as supplier_name
            FROM comanda_items ci
            LEFT JOIN suppliers s ON ci.supplier_id = s.id
            WHERE ci.id = ?
        """,
            (load_offer_item_id,),
        ).fetchone()
        conn_offer.close()
        if item_off:
            load_offer_data = dict(item_off)

    return render_template(
        'base.html',
        menu_options=menu_options,
        current_module=module,
        username=curr_username,
        user_display_name=user_disp,
        user_phone=user_phone,
        role=role,
        tasa_usd=tasa_usd,
        tasa_eur=tasa_eur,
        budget_currency=session['budget_currency'],
        budget_client=session.get('budget_client', {}),
        all_modules=ALL_SYSTEM_MODULES,
        all_budgets=all_budgets,
        filtered_budgets=filtered_budgets,
        pending_budgets=pending_budgets,
        seller_summary=seller_summary,
        reminders=reminders,
        reminders_count=reminders_count,
        reminders_today_count=reminders_today_count,
        reminder_categories=reminder_categories,
        r_search=r_search,
        r_priority=r_priority,
        r_category=r_category,
        r_status=r_status,
        r_author=r_author,
        r_sort=r_sort,
        appointments_count=appointments_count,
        all_users=all_users,
        users_map=users_map,
        users=users,
        roles_list=roles_list,
        all_clients=all_clients,
        c_search=c_search,
        c_filter=c_filter,
        c_sort=c_sort,
        c_pagination=c_pagination,
        vehicle_list=vehicle_list,
        v_search=v_search,
        v_client=v_client,
        v_sort=v_sort,
        v_pagination=v_pagination,
        all_revisions=all_revisions,
        selected_client_doc=selected_client_doc,
        client_budgets=client_budgets,
        client_revisions=client_revisions,
        f_type=f_type,
        f_status=f_status,
        f_user=f_user,
        f_date=f_date,
        f_month=f_month,
        f_start=f_start,
        f_end=f_end,
        log_user=log_user,
        log_action_filter=log_action_filter,
        current_date=current_date,
        is_editing=False,
        is_editing_revision=False,
        templates=templates,
        appointments=appointments,
        appointment_filter=appointment_filter,
        vehicle_search=vehicle_search,
        report_data=report_data,
        action_logs=action_logs,
        action_types=action_types,
        dashboard_metrics=dashboard_metrics,
        only_drafts=only_drafts,
        workshop_columns=workshop_columns,
        cards_by_column=cards_by_column,
        workshop_stats=workshop_stats,
        wb_view=wb_view,
        workshop_history=workshop_history,
        workshop_history_stats=workshop_history_stats,
        # Le pasamos a JS los nombres exactos como están en la BD
        workshop_sellers=(
            [u['username'] for u in all_users_list]
            if all_users_list
            else get_workshop_sellers()
        ),
        hw_user=hw_user,
        hw_search=hw_search,
        hw_date_from=hw_date_from,
        hw_date_to=hw_date_to,
        all_users_list=[
            dict(u) for u in (all_users_list if module == '🔧 Tablero Taller' else [])
        ],
        workshop_mechanics=workshop_mechanics,
        workshop_repair_statuses=WORKSHOP_REPAIR_STATUSES,
        revision_kpis=revision_kpis,
        mechanic_revision_kpis=locals().get('mechanic_revision_kpis'),
        mr_pagination=locals().get('mr_pagination'),
        mr_status=locals().get('mr_status', 'Todos'),
        recent_revisions=recent_revisions,
        r_pagination=r_pagination,
        h_pagination=h_pagination,
        h_search=h_search,
        # ===== Inventario (NUEVO) =====
        inventory_list=inventory_list,
        inventory_categories=inventory_categories,
        inventory_stats=inventory_stats,
        inv_search=inv_search,
        inv_category_filter=inv_category_filter,
        inv_pagination=inv_pagination,
        # ===== Expedientes (NUEVO) =====
        expedientes_list=expedientes_list,
        exp_pagination=exp_pagination,
        cliente_exp=cliente_exp,
        exp_budgets=exp_budgets,
        exp_revisions=exp_revisions,
        exp_cards=exp_cards,
        exp_appointments=exp_appointments,
        exp_vehicles=exp_vehicles,
        exp_timeline=exp_timeline,
        exp_stats=exp_stats,
        exp_search=exp_search,
        # ===== Vehículos =====
        vehicle_stats=vehicle_stats,
        report_view=locals().get('report_view', 'analytics'),
        daily_close=locals().get('daily_close', None),
        purchases_history=purchases_history,
        comandas_list=locals().get('comandas_list', []),
        # 👇 AQUÍ ESTÁ EL FIX 👇
        solicitudes_pendientes=locals().get('solicitudes_pendientes', []),
        solicitudes_procesadas=locals().get('solicitudes_procesadas', []),
        # 👆 AQUÍ ESTÁ EL FIX 👆
        suppliers=locals().get('suppliers', []),
        load_offer_data=load_offer_data,
    )


# ==========================================
# RUTAS DE RECORDATORIOS
# ==========================================
def _save_reminder_attachment(file_storage):
    if not file_storage or not file_storage.filename:
        return '', '', 0
    filename = secure_filename(file_storage.filename)
    if not filename:
        return '', '', 0
    ts = int(time.time())
    unique = f"{ts}_{secrets.token_hex(4)}_{filename}"
    full_path = os.path.join(UPLOADS_DIR, unique)
    file_storage.save(full_path)
    is_image = (
        1
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp'))
        else 0
    )
    rel_path = f"/static/uploads/reminders/{unique}"
    return rel_path, filename, is_image


@app.route('/reminder/add', methods=['POST'])
def add_reminder():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    title = request.form.get('r_title', '').strip()
    content = request.form.get('r_content', '').strip()
    if not title or not content:
        flash('El título y contenido son obligatorios.', 'danger')
        return redirect(url_for('index', module='🔔 Recordatorios'))

    priority = request.form.get('r_priority', 'Media')
    if priority not in REMINDER_PRIORITIES:
        priority = 'Media'
    category = request.form.get('r_category', '').strip()
    due_date = request.form.get('r_due_date', '').strip()
    is_private = 1 if request.form.get('r_private') == '1' else 0
    assigned_to = request.form.get('r_assigned_to', '').strip()
    client_doc = up(request.form.get('r_client_doc', ''))
    client_name = request.form.get('r_client_name', '').strip()
    vehicle_plate = up(request.form.get('r_vehicle_plate', ''))
    recurrence = request.form.get('r_recurrence', 'none')
    if recurrence not in REMINDER_RECURRENCES:
        recurrence = 'none'

    attachment_path, attachment_name, attachment_is_image = '', '', 0
    try:
        if 'r_attachment' in request.files:
            attachment_path, attachment_name, attachment_is_image = (
                _save_reminder_attachment(request.files['r_attachment'])
            )
    except Exception as e:
        flash(f'No se pudo guardar el adjunto: {e}', 'danger')

    username = session.get('username', '')
    date_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db_connection()
    conn.execute(
        """
        INSERT INTO reminders (
            username, title, content, date, is_seen, priority, category, due_date,
            attachment, attachment_name, attachment_is_image,
            is_private, assigned_to, client_doc, client_name, vehicle_plate,
            recurrence, updated_at
        ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            username,
            title,
            content,
            date_str,
            priority,
            category,
            due_date,
            attachment_path,
            attachment_name,
            attachment_is_image,
            is_private,
            assigned_to,
            client_doc,
            client_name,
            vehicle_plate,
            recurrence,
            now_str,
        ),
    )
    reminder_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.commit()
    conn.close()

    try:
        c2 = get_db_connection()
        c2.execute(
            "INSERT OR IGNORE INTO user_reminder_views (username, reminder_id) VALUES (?, ?)",
            (username, reminder_id),
        )
        c2.commit()
        c2.close()
    except Exception:
        pass

    log_action(
        'Creación de recordatorio',
        f'Se creó el recordatorio "{title}".',
        'Recordatorio',
        reminder_id,
    )
    flash('Recordatorio creado exitosamente.', 'success')
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/edit/<int:id>', methods=['POST'])
def edit_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    rem = conn.execute("SELECT * FROM reminders WHERE id = ?", (id,)).fetchone()
    if not rem:
        conn.close()
        flash('Recordatorio no encontrado.', 'danger')
        return redirect(url_for('index', module='🔔 Recordatorios'))

    if session.get('role') != 'Administrador' and rem['username'] != session.get(
        'username'
    ):
        conn.close()
        flash(
            'Solo el creador o un Administrador pueden editar este recordatorio.',
            'danger',
        )
        return redirect(url_for('index', module='🔔 Recordatorios'))

    title = request.form.get('r_title', '').strip()
    content = request.form.get('r_content', '').strip()
    priority = request.form.get('r_priority', 'Media')
    if priority not in REMINDER_PRIORITIES:
        priority = 'Media'
    category = request.form.get('r_category', '').strip()
    due_date = request.form.get('r_due_date', '').strip()
    is_private = 1 if request.form.get('r_private') == '1' else 0
    assigned_to = request.form.get('r_assigned_to', '').strip()
    client_doc = up(request.form.get('r_client_doc', ''))
    client_name = request.form.get('r_client_name', '').strip()
    vehicle_plate = up(request.form.get('r_vehicle_plate', ''))
    recurrence = request.form.get('r_recurrence', 'none')
    if recurrence not in REMINDER_RECURRENCES:
        recurrence = 'none'

    remove_att = request.form.get('r_remove_attachment') == '1'

    new_att_path = rem['attachment'] if 'attachment' in rem.keys() else ''
    new_att_name = rem['attachment_name'] if 'attachment_name' in rem.keys() else ''
    new_att_is_image = (
        rem['attachment_is_image'] if 'attachment_is_image' in rem.keys() else 0
    )

    if remove_att and new_att_path:
        try:
            old_full = os.path.join(BASE_DIR, new_att_path.lstrip('/'))
            if os.path.exists(old_full):
                os.remove(old_full)
        except Exception:
            pass
        new_att_path, new_att_name, new_att_is_image = '', '', 0

    try:
        if 'r_attachment' in request.files and request.files['r_attachment'].filename:
            if new_att_path:
                try:
                    old_full = os.path.join(BASE_DIR, new_att_path.lstrip('/'))
                    if os.path.exists(old_full):
                        os.remove(old_full)
                except Exception:
                    pass
            new_att_path, new_att_name, new_att_is_image = _save_reminder_attachment(
                request.files['r_attachment']
            )
    except Exception as e:
        flash(f'No se pudo guardar el adjunto: {e}', 'danger')

    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """
        UPDATE reminders SET title = ?, content = ?, priority = ?, category = ?, due_date = ?,
            attachment = ?, attachment_name = ?, attachment_is_image = ?,
            is_private = ?, assigned_to = ?, client_doc = ?, client_name = ?,
            vehicle_plate = ?, recurrence = ?, updated_at = ?
        WHERE id = ?
    """,
        (
            title,
            content,
            priority,
            category,
            due_date,
            new_att_path,
            new_att_name,
            new_att_is_image,
            is_private,
            assigned_to,
            client_doc,
            client_name,
            vehicle_plate,
            recurrence,
            now_str,
            id,
        ),
    )
    conn.commit()
    conn.close()
    log_action(
        'Edición de recordatorio',
        f'Se editó el recordatorio #{id}.',
        'Recordatorio',
        id,
    )
    flash('Recordatorio actualizado.', 'success')
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/toggle_complete/<int:id>', methods=['POST'])
def toggle_complete_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    rem = conn.execute("SELECT * FROM reminders WHERE id = ?", (id,)).fetchone()
    if not rem:
        conn.close()
        flash('Recordatorio no encontrado.', 'danger')
        return redirect(url_for('index', module='🔔 Recordatorios'))

    is_completed = int(rem['is_completed'] or 0)
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    actor = session.get('username', '')

    if not is_completed:
        conn.execute(
            """
            UPDATE reminders SET is_completed = 1, completed_at = ?, completed_by = ?, updated_at = ?
            WHERE id = ?
        """,
            (now_str, actor, now_str, id),
        )
        conn.commit()

        recurrence = rem['recurrence'] if 'recurrence' in rem.keys() else 'none'
        if recurrence and recurrence != 'none':
            base_date = (
                rem['due_date']
                if 'due_date' in rem.keys() and rem['due_date']
                else rem['date'][:10]
            )
            try:
                base_dt = datetime.datetime.strptime(base_date[:10], "%Y-%m-%d")
            except Exception:
                base_dt = get_ve_time().replace(tzinfo=None)
            if recurrence == 'daily':
                next_dt = base_dt + timedelta(days=1)
            elif recurrence == 'weekly':
                next_dt = base_dt + timedelta(days=7)
            elif recurrence == 'monthly':
                next_dt = add_months_to_date(base_dt, 1)
            else:
                next_dt = base_dt
            next_date_str = next_dt.strftime("%Y-%m-%d")
            new_date = get_ve_time().strftime("%Y-%m-%d %H:%M")
            conn.execute(
                """
                INSERT INTO reminders (
                    username, title, content, date, is_seen, priority, category, due_date,
                    attachment, attachment_name, attachment_is_image,
                    is_private, assigned_to, client_doc, client_name, vehicle_plate,
                    recurrence, updated_at
                ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    rem['username'],
                    rem['title'],
                    rem['content'],
                    new_date,
                    rem['priority'] if 'priority' in rem.keys() else 'Media',
                    rem['category'] if 'category' in rem.keys() else '',
                    next_date_str,
                    rem['attachment'] if 'attachment' in rem.keys() else '',
                    rem['attachment_name'] if 'attachment_name' in rem.keys() else '',
                    (
                        rem['attachment_is_image']
                        if 'attachment_is_image' in rem.keys()
                        else 0
                    ),
                    rem['is_private'] if 'is_private' in rem.keys() else 0,
                    rem['assigned_to'] if 'assigned_to' in rem.keys() else '',
                    rem['client_doc'] if 'client_doc' in rem.keys() else '',
                    rem['client_name'] if 'client_name' in rem.keys() else '',
                    rem['vehicle_plate'] if 'vehicle_plate' in rem.keys() else '',
                    recurrence,
                    now_str,
                ),
            )
            conn.commit()
        flash('Recordatorio marcado como completado.', 'success')
    else:
        conn.execute(
            """
            UPDATE reminders SET is_completed = 0, completed_at = '', completed_by = '', updated_at = ?
            WHERE id = ?
        """,
            (now_str, id),
        )
        conn.commit()
        flash('Recordatorio reabierto.', 'success')

    conn.close()
    log_action(
        'Cambio de estado de recordatorio',
        f'Se {"completó" if not is_completed else "reabrió"} el recordatorio #{id}.',
        'Recordatorio',
        id,
    )
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/archive/<int:id>', methods=['POST'])
def archive_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE reminders SET archived = 1, archived_at = ?, updated_at = ? WHERE id = ?",
        (now_str, now_str, id),
    )
    conn.commit()
    conn.close()
    log_action(
        'Archivado de recordatorio',
        f'Se archivó el recordatorio #{id}.',
        'Recordatorio',
        id,
    )
    flash('Recordatorio archivado.', 'success')
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/unarchive/<int:id>', methods=['POST'])
def unarchive_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE reminders SET archived = 0, archived_at = '', updated_at = ? WHERE id = ?",
        (now_str, id),
    )
    conn.commit()
    conn.close()
    log_action(
        'Desarchivado', f'Se desarchivó el recordatorio #{id}.', 'Recordatorio', id
    )
    flash('Recordatorio restaurado.', 'success')
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/duplicate/<int:id>', methods=['POST'])
def duplicate_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    rem = conn.execute("SELECT * FROM reminders WHERE id = ?", (id,)).fetchone()
    if not rem:
        conn.close()
        flash('Recordatorio no encontrado.', 'danger')
        return redirect(url_for('index', module='🔔 Recordatorios'))

    username = session.get('username', '')
    date_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    new_title = f"[COPIA] {rem['title']}"
    conn.execute(
        """
        INSERT INTO reminders (
            username, title, content, date, is_seen, priority, category, due_date,
            attachment, attachment_name, attachment_is_image,
            is_private, assigned_to, client_doc, client_name, vehicle_plate,
            recurrence, updated_at
        ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            username,
            new_title,
            rem['content'],
            date_str,
            rem['priority'] if 'priority' in rem.keys() else 'Media',
            rem['category'] if 'category' in rem.keys() else '',
            rem['due_date'] if 'due_date' in rem.keys() else '',
            rem['attachment'] if 'attachment' in rem.keys() else '',
            rem['attachment_name'] if 'attachment_name' in rem.keys() else '',
            rem['attachment_is_image'] if 'attachment_is_image' in rem.keys() else 0,
            rem['is_private'] if 'is_private' in rem.keys() else 0,
            rem['assigned_to'] if 'assigned_to' in rem.keys() else '',
            rem['client_doc'] if 'client_doc' in rem.keys() else '',
            rem['client_name'] if 'client_name' in rem.keys() else '',
            rem['vehicle_plate'] if 'vehicle_plate' in rem.keys() else '',
            rem['recurrence'] if 'recurrence' in rem.keys() else 'none',
            now_str,
        ),
    )
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.commit()
    conn.close()
    try:
        c2 = get_db_connection()
        c2.execute(
            "INSERT OR IGNORE INTO user_reminder_views (username, reminder_id) VALUES (?, ?)",
            (username, new_id),
        )
        c2.commit()
        c2.close()
    except Exception:
        pass
    log_action(
        'Duplicado',
        f'Se duplicó el recordatorio #{id} → #{new_id}.',
        'Recordatorio',
        new_id,
    )
    flash('Recordatorio duplicado.', 'success')
    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminder/delete/<int:id>', methods=['POST'])
def delete_reminder(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    rem = conn.execute("SELECT * FROM reminders WHERE id = ?", (id,)).fetchone()
    deleted = False

    if rem:
        if session.get('role') == 'Administrador' or rem['username'] == session.get(
            'username'
        ):
            try:
                att = rem['attachment'] if 'attachment' in rem.keys() else ''
                if att:
                    full = os.path.join(BASE_DIR, att.lstrip('/'))
                    if os.path.exists(full):
                        os.remove(full)
            except Exception:
                pass
            conn.execute("DELETE FROM reminders WHERE id = ?", (id,))
            conn.execute("DELETE FROM user_reminder_views WHERE reminder_id = ?", (id,))
            conn.commit()
            deleted = True
        else:
            flash(
                'Solo el creador o un Administrador pueden eliminar este recordatorio.',
                'danger',
            )

    conn.close()

    if deleted:
        flash('Recordatorio eliminado.', 'success')
        log_action(
            'Eliminación de recordatorio',
            f'Se eliminó el recordatorio #{id}.',
            'Recordatorio',
            id,
        )

    return redirect(url_for('index', module='🔔 Recordatorios'))


@app.route('/reminders/mark_seen', methods=['POST'])
def mark_reminders_seen():
    if not session.get('logged_in'):
        return jsonify({"ok": False}), 401
    username = session['username']
    payload = request.get_json(silent=True) or {}
    reminder_ids = payload.get('reminder_ids', [])

    conn = get_db_connection()
    try:
        if reminder_ids:
            for rid in reminder_ids:
                try:
                    conn.execute(
                        "INSERT OR IGNORE INTO user_reminder_views (username, reminder_id) VALUES (?, ?)",
                        (username, int(rid)),
                    )
                except (ValueError, TypeError):
                    continue
        else:
            conn.execute(
                """
                INSERT OR IGNORE INTO user_reminder_views (username, reminder_id)
                SELECT ?, id FROM reminders WHERE archived = 0
            """,
                (username,),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route('/api/reminders/unseen', methods=['GET'])
def api_reminders_unseen():
    if not session.get('logged_in'):
        return jsonify({"ok": False, "items": []}), 401
    username = session['username']
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT * FROM reminders
        WHERE id NOT IN (SELECT reminder_id FROM user_reminder_views WHERE username = ?)
          AND (is_private = 0 OR username = ?)
          AND (assigned_to = '' OR assigned_to = ?)
          AND archived = 0 AND is_completed = 0
        ORDER BY CASE priority WHEN 'Alta' THEN 1 WHEN 'Media' THEN 2 ELSE 3 END ASC,
                 CASE WHEN due_date = '' OR due_date IS NULL THEN 1 ELSE 0 END ASC,
                 due_date ASC, id DESC
        LIMIT 100
    """,
        (username, username, username),
    ).fetchall()
    conn.close()
    items = [reminder_to_dict(r) for r in rows]
    return jsonify({"ok": True, "items": items, "count": len(items)})


# ==========================================
# PRESUPUESTOS Y PAGOS
# ==========================================
@app.route('/p/<token>', methods=['GET', 'POST'])
def public_budget(token):
    conn = get_db_connection()
    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE public_token = ?", (token,)
    ).fetchone()
    if not budget:
        conn.close()
        return "Presupuesto no encontrado.", 404
    payments = conn.execute(
        "SELECT * FROM payments WHERE budget_id = ? ORDER BY id ASC", (budget['id'],)
    ).fetchall()
    items = json.loads(budget['items_json'])
    conn.close()

    if request.method == 'POST':
        action = request.form.get('action', '')
        signature_data = request.form.get('signature_data', '')
        conn = get_db_connection()

        # ENVOLVEMOS EN TRY...FINALLY PARA GARANTIZAR EL CIERRE
        try:
            if action == 'approve':
                now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
                old_status = budget['status']
                conn.execute(
                    """UPDATE saved_budgets SET status = 'Aprobado', client_signature = ?, client_signature_date = ? WHERE id = ?""",
                    (signature_data, now_str, budget['id']),
                )

                if old_status != 'Aprobado':
                    update_stock_for_budget(conn, budget['id'], 'deduct')

                conn.commit()
                log_action(
                    'Aprobación',
                    f'Presupuesto #{budget["id"]} aprobado vía link público.',
                    'Presupuesto',
                    budget['id'],
                    username='Cliente',
                )
                flash('¡Presupuesto aprobado! Gracias.', 'success')
                return redirect(url_for('public_budget', token=token))

            elif action == 'reject':
                conn.execute(
                    "UPDATE saved_budgets SET status = 'Anulado' WHERE id = ?",
                    (budget['id'],),
                )
                conn.commit()
                flash('Presupuesto rechazado.', 'success')
                return redirect(url_for('public_budget', token=token))
        finally:
            # ESTO SE EJECUTA SIEMPRE, incluso si hubo un 'return' arriba o un error
            conn.close()

    conn = get_db_connection()
    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE public_token = ?", (token,)
    ).fetchone()
    payments = conn.execute(
        "SELECT * FROM payments WHERE budget_id = ? ORDER BY id ASC", (budget['id'],)
    ).fetchall()
    items = json.loads(budget['items_json'])
    conn.close()
    payment_status = compute_payment_status(budget)
    total_paid = float(budget['total_paid_foreign'] or 0)
    remaining = float(budget['total_foreign']) - total_paid
    return render_template(
        'presupuesto.html',
        budget=budget,
        items=items,
        payments=payments,
        payment_status=payment_status,
        total_paid=total_paid,
        remaining=remaining,
    )


@app.route('/budget/<int:budget_id>/payments', methods=['GET', 'POST'])
@block_budget_access
def budget_payments(budget_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (budget_id,)
    ).fetchone()
    if not budget:
        conn.close()
        flash('Presupuesto no encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))

    if request.method == 'POST':
        amount = safe_float(request.form.get('amount'), 0.0)
        currency = request.form.get('currency', budget['currency'])
        method = request.form.get('method', 'Efectivo USD')
        reference = request.form.get('reference', '').strip()
        note = request.form.get('note', '').strip()

        if amount > 0:
            tasa_u = get_config('tasa_usd')
            tasa_e = get_config('tasa_eur')
            base_currency = budget['currency']
            bcv_rate_budget = float(budget['bcv_rate'])

            if base_currency not in ('USD', 'EUR'):
                flash('Moneda base no válida.', 'danger')
                conn.close()
                return redirect(url_for('budget_payments', budget_id=budget_id))

            if currency == base_currency:
                rate_used = bcv_rate_budget
                amount_foreign = amount
                amount_ves = amount_foreign * rate_used
            elif currency == 'VES':
                tasa_dia = tasa_u if base_currency == 'USD' else tasa_e
                if not tasa_dia or tasa_dia <= 0:
                    flash('Tasa BCV del día no configurada.', 'danger')
                    conn.close()
                    return redirect(url_for('budget_payments', budget_id=budget_id))
                rate_used = float(tasa_dia)
                amount_foreign = amount / rate_used
                amount_ves = amount
            elif base_currency == 'USD' and currency == 'EUR':
                if tasa_u <= 0 or tasa_e <= 0:
                    flash('Tasas BCV no configuradas.', 'danger')
                    conn.close()
                    return redirect(url_for('budget_payments', budget_id=budget_id))
                rate_used = bcv_rate_budget
                amount_foreign = amount * (tasa_e / tasa_u)
                amount_ves = amount_foreign * rate_used
            elif base_currency == 'EUR' and currency == 'USD':
                if tasa_u <= 0 or tasa_e <= 0:
                    flash('Tasas BCV no configuradas.', 'danger')
                    conn.close()
                    return redirect(url_for('budget_payments', budget_id=budget_id))
                rate_used = bcv_rate_budget
                amount_foreign = amount * (tasa_u / tasa_e)
                amount_ves = amount_foreign * rate_used
            else:
                flash('Combinación no soportada.', 'danger')
                conn.close()
                return redirect(url_for('budget_payments', budget_id=budget_id))

            now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
            conn.execute(
                """
                INSERT INTO payments (budget_id, date, amount_foreign, amount_ves, currency, rate_bcv, method, reference, username, note, original_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    budget_id,
                    now_str,
                    amount_foreign,
                    amount_ves,
                    currency,
                    rate_used,
                    method,
                    reference,
                    session['username'],
                    note,
                    amount,
                ),
            )

            total_paid = conn.execute(
                "SELECT COALESCE(SUM(amount_foreign),0) as t FROM payments WHERE budget_id = ?",
                (budget_id,),
            ).fetchone()['t']
            total_paid_ves = conn.execute(
                "SELECT COALESCE(SUM(amount_ves),0) as t FROM payments WHERE budget_id = ?",
                (budget_id,),
            ).fetchone()['t']
            conn.execute(
                "UPDATE saved_budgets SET total_paid_foreign = ?, total_paid_ves = ? WHERE id = ?",
                (total_paid, total_paid_ves, budget_id),
            )

            total_foreign = float(budget['total_foreign'] or 0)
            if total_paid >= total_foreign - 0.01:
                b_status = conn.execute(
                    "SELECT status FROM saved_budgets WHERE id = ?", (budget_id,)
                ).fetchone()
                if b_status and b_status['status'] == 'Pendiente':
                    conn.execute(
                        "UPDATE saved_budgets SET status = 'Aprobado' WHERE id = ?",
                        (budget_id,),
                    )
                    update_stock_for_budget(conn, budget_id, 'deduct')

            conn.commit()
            conn.close()
            flash(f'Abono de {amount:.2f} {currency} registrado.', 'success')
        else:
            conn.close()
        return redirect(url_for('budget_payments', budget_id=budget_id))

    payments = conn.execute(
        "SELECT * FROM payments WHERE budget_id = ? ORDER BY id DESC", (budget_id,)
    ).fetchall()
    conn.close()
    try:
        items = json.loads(budget['items_json'])
    except Exception:
        items = []
    payment_status = compute_payment_status(budget)
    total_paid = float(budget['total_paid_foreign'] or 0)
    remaining = float(budget['total_foreign']) - total_paid
    return render_template(
        'presupuesto.html',
        budget=budget,
        items=items,
        payments=payments,
        payment_status=payment_status,
        total_paid=total_paid,
        remaining=remaining,
        tasa_usd=get_config('tasa_usd'),
        tasa_eur=get_config('tasa_eur'),
        show_payment_form=True,
    )


@app.route('/payment/<int:payment_id>/delete', methods=['POST'])
def delete_payment(payment_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    payment = conn.execute(
        "SELECT * FROM payments WHERE id = ?", (payment_id,)
    ).fetchone()
    if not payment:
        conn.close()
        flash('Pago no encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))
    budget_id = payment['budget_id']
    if session.get('role') != 'Administrador' and payment['username'] != session.get(
        'username'
    ):
        conn.close()
        flash('No tienes permiso.', 'danger')
        return redirect(url_for('budget_payments', budget_id=budget_id))
    conn.execute("DELETE FROM payments WHERE id = ?", (payment_id,))
    total_paid = conn.execute(
        "SELECT COALESCE(SUM(amount_foreign),0) as t FROM payments WHERE budget_id = ?",
        (budget_id,),
    ).fetchone()['t']
    total_paid_ves = conn.execute(
        "SELECT COALESCE(SUM(amount_ves),0) as t FROM payments WHERE budget_id = ?",
        (budget_id,),
    ).fetchone()['t']
    conn.execute(
        "UPDATE saved_budgets SET total_paid_foreign = ?, total_paid_ves = ? WHERE id = ?",
        (total_paid, total_paid_ves, budget_id),
    )
    conn.commit()
    conn.close()
    flash('Pago eliminado.', 'success')
    return redirect(url_for('budget_payments', budget_id=budget_id))


@app.route('/budget/<int:budget_id>/signature', methods=['GET', 'POST'])
@block_budget_access
def budget_signature(budget_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (budget_id,)
    ).fetchone()
    if not budget:
        conn.close()
        flash('Presupuesto no encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))
    if request.method == 'POST':
        signature_data = request.form.get('signature_data', '')
        if signature_data:
            now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
            conn.execute(
                """UPDATE saved_budgets SET client_signature = ?, client_signature_date = ?, status = 'Aprobado' WHERE id = ?""",
                (signature_data, now_str, budget_id),
            )
            conn.commit()
            conn.close()
            flash('Firma guardada.', 'success')
            return redirect(url_for('index', module='📂 Presupuestos'))
        conn.close()
        flash('Debe dibujar una firma.', 'danger')
        return redirect(url_for('budget_signature', budget_id=budget_id))
    conn.close()
    return render_template('firma.html', budget=budget)


@app.route('/template/add', methods=['POST'])
def add_template():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    desc = request.form.get('description', '').strip()
    if desc:
        conn = get_db_connection()
        conn.execute(
            "INSERT INTO service_templates (description, default_price, currency, category) VALUES (?, ?, ?, ?)",
            (
                desc,
                safe_float(request.form.get('default_price'), 0.0),
                request.form.get('currency', 'USD'),
                request.form.get('category', '').strip(),
            ),
        )
        conn.commit()
        conn.close()
        flash('Plantilla creada.', 'success')
    return redirect(url_for('index', module='⚙️ Plantillas'))


@app.route('/template/edit/<int:template_id>', methods=['POST'])
def edit_template(template_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    desc = request.form.get('description', '').strip()
    if desc:
        conn = get_db_connection()
        conn.execute(
            """UPDATE service_templates SET description = ?, default_price = ?, currency = ?, category = ? WHERE id = ?""",
            (
                desc,
                safe_float(request.form.get('default_price'), 0.0),
                request.form.get('currency', 'USD'),
                request.form.get('category', '').strip(),
                template_id,
            ),
        )
        conn.commit()
        conn.close()
        flash('Plantilla actualizada.', 'success')
    return redirect(url_for('index', module='⚙️ Plantillas'))


@app.route('/template/delete/<int:template_id>', methods=['POST'])
def delete_template(template_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute("DELETE FROM service_templates WHERE id = ?", (template_id,))
    conn.commit()
    conn.close()
    flash('Plantilla eliminada.', 'success')
    return redirect(url_for('index', module='⚙️ Plantillas'))


@app.route('/appointment/add', methods=['POST'])
def add_appointment():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute(
        """
        INSERT INTO appointments (date, time, client_name, client_doc, client_phone, vehicle_name, vehicle_plate, service_description, status, username, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente', ?, ?, ?)
    """,
        (
            request.form.get('apt_date', '').strip(),
            request.form.get('apt_time', '').strip(),
            request.form.get('apt_client_name', '').strip(),
            up(request.form.get('apt_client_doc', '')),
            request.form.get('apt_client_phone', '').strip(),
            request.form.get('apt_vehicle_name', '').strip(),
            up(request.form.get('apt_vehicle_plate', '')),
            request.form.get('apt_service', '').strip(),
            session['username'],
            request.form.get('apt_note', '').strip(),
            get_ve_time().strftime("%Y-%m-%d %H:%M"),
        ),
    )
    conn.commit()
    conn.close()
    flash('Cita agendada.', 'success')
    return redirect(url_for('index', module='📅 Citas'))


@app.route('/appointment/status/<int:apt_id>/<status>', methods=['POST'])
def change_appointment_status(apt_id, status):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    if status not in APPOINTMENT_STATUSES:
        flash('Estado inválido.', 'danger')
        return redirect(url_for('index', module='📅 Citas'))
    conn = get_db_connection()
    conn.execute("UPDATE appointments SET status = ? WHERE id = ?", (status, apt_id))
    conn.commit()
    conn.close()
    flash(f'Cita marcada como {status}.', 'success')
    return redirect(url_for('index', module='📅 Citas'))


@app.route('/appointment/delete/<int:apt_id>', methods=['POST'])
def delete_appointment(apt_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute("DELETE FROM appointments WHERE id = ?", (apt_id,))
    conn.commit()
    conn.close()
    flash('Cita eliminada.', 'success')
    return redirect(url_for('index', module='📅 Citas'))


@app.route('/export/budgets.csv')
def export_budgets_csv():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    f_type = request.args.get('f_type', 'Todos')
    f_user = request.args.get('f_user', 'Todos')
    f_date = request.args.get('f_date', '')
    f_month = request.args.get('f_month', '')
    f_start = request.args.get('f_start', '')
    f_end = request.args.get('f_end', '')

    query = "SELECT * FROM saved_budgets WHERE 1=1"
    params = []
    if f_user and f_user != 'Todos':
        query += " AND username = ?"
        params.append(f_user)
    if f_type == 'Dia' and f_date:
        query += " AND date LIKE ?"
        params.append(f"{f_date}%")
    elif f_type == 'Mes' and f_month:
        query += " AND date LIKE ?"
        params.append(f"{f_month}%")
    elif f_type == 'Rango' and f_start and f_end:
        query += " AND substr(date, 1, 10) >= ? AND substr(date, 1, 10) <= ?"
        params.extend([f_start, f_end])
    query += " ORDER BY id DESC"

    conn = get_db_connection()
    rows = conn.execute(query, params).fetchall()
    conn.close()

    si = StringIO()
    si.write('\ufeff')
    writer = csv.writer(si, delimiter=';')
    writer.writerow(
        [
            'ID',
            'Fecha',
            'Cliente',
            'C.I/RIF',
            'Teléfono',
            'Vehículo',
            'Año',
            'Placa',
            'Vendedor',
            'Moneda',
            'Tasa BCV',
            'Subtotal',
            'Descuento %',
            'Descuento Monto',
            'Base Imponible',
            'IVA',
            'Total Divisa',
            'Total Bs',
            'Ganancia',
            'Abonado',
            'Saldo',
            'Estado',
            'Estado Pago',
            'Nota',
        ]
    )
    for b in rows:
        try:
            items = json.loads(b['items_json'])
            subtotal = sum(
                float(it.get('gross_total', it.get('total', 0))) for it in items
            )
        except Exception:
            subtotal = 0.0
        total_paid = float(b['total_paid_foreign'] or 0)
        saldo = float(b['total_foreign']) - total_paid
        pstatus = compute_payment_status(b)
        writer.writerow(
            [
                b['id'],
                b['date'],
                b['client_name'],
                b['client_doc'],
                b['client_phone'],
                b['vehicle_name'],
                b['vehicle_year'],
                b['vehicle_plate'],
                b['username'],
                b['currency'],
                f"{float(b['bcv_rate']):.2f}",
                f"{subtotal:.2f}",
                f"{float(b['discount_percent'] or 0):.2f}",
                f"{float(b['discount_amount'] or 0):.2f}",
                f"{float(b['base_imponible'] or 0):.2f}",
                f"{float(b['iva_amount'] or 0):.2f}",
                f"{float(b['total_foreign']):.2f}",
                f"{float(b['total_ves']):.2f}",
                f"{float(b['total_profit_foreign'] or 0):.2f}",
                f"{total_paid:.2f}",
                f"{saldo:.2f}",
                b['status'],
                pstatus,
                b['budget_note'],
            ]
        )
    csv_bytes = si.getvalue().encode('utf-8')
    return send_file(
        BytesIO(csv_bytes),
        mimetype='text/csv; charset=utf-8',
        as_attachment=True,
        download_name=f"presupuestos_{get_ve_time().strftime('%Y%m%d_%H%M')}.csv",
    )


@app.route('/update_rates', methods=['POST'])
def update_rates():
    usd = request.form.get('inp_usd')
    eur = request.form.get('inp_eur')
    if usd:
        set_config('tasa_usd', usd)
    if eur:
        set_config('tasa_eur', eur)
    flash('¡Tasas actualizadas!', 'success')
    return redirect(url_for('index'))


@app.route('/api/rates/check', methods=['GET'])
def check_rates_api():
    if not session.get('logged_in'):
        return jsonify({"error": "no autenticado"}), 401
    tasa_usd = get_config('tasa_usd')
    tasa_eur = get_config('tasa_eur')
    updated_usd = updated_eur = False
    try:
        api_usd, api_eur = fetch_bcv_rates(force=True)
    except Exception:
        api_usd, api_eur = None, None
    if api_usd and abs(api_usd - tasa_usd) > 0.001:
        set_config('tasa_usd', api_usd)
        tasa_usd = api_usd
        updated_usd = True
    if api_eur and abs(api_eur - tasa_eur) > 0.001:
        set_config('tasa_eur', api_eur)
        tasa_eur = api_eur
        updated_eur = True
    return jsonify(
        {
            "tasa_usd": tasa_usd,
            "tasa_eur": tasa_eur,
            "updated_usd": updated_usd,
            "updated_eur": updated_eur,
            "updated": updated_usd or updated_eur,
        }
    )


@app.route('/set_budget_currency', methods=['POST'])
def set_budget_currency():
    curr = request.form.get('currency', 'USD')
    session['budget_currency'] = curr
    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


@app.route('/client/save_direct', methods=['POST'])
def client_save_direct():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    step = request.form.get('step', 'client')

    if step == 'client':
        doc = up(request.form.get('client_doc', ''))
        name = request.form.get('client_name', '').strip()
        lastname = request.form.get('client_lastname', '').strip()
        phone = request.form.get('client_phone', '').strip()
        if not name and not lastname:
            flash('Indica al menos el nombre o apellido.', 'danger')
            return redirect(url_for('index', module='📝 Nuevo Presupuesto'))
        try:
            real_doc = doc
            if doc:
                conn = get_db_connection()
                existing = conn.execute(
                    "SELECT vehicle_name, vehicle_year, vehicle_plate, vehicle_brand, vehicle_color, vehicle_model "
                    "FROM clients WHERE client_doc = ?",
                    (doc,),
                ).fetchone()
                conn.close()
            else:
                existing = None
            if existing:
                save_or_update_client(
                    doc,
                    name,
                    phone,
                    existing['vehicle_name'] or '',
                    existing['vehicle_year'] or '',
                    existing['vehicle_plate'] or '',
                    existing['vehicle_brand'] or '',
                    existing['vehicle_color'] or '',
                    existing['vehicle_model'] or '',
                    lastname=lastname,
                )
                bc = dict(session.get('budget_client', {}))
                bc.update(
                    {
                        'veh_name': existing['vehicle_name'] or '',
                        'veh_year': existing['vehicle_year'] or '',
                        'veh_plate': existing['vehicle_plate'] or '',
                        'veh_brand': existing['vehicle_brand'] or '',
                        'veh_color': existing['vehicle_color'] or '',
                        'veh_model': existing['vehicle_model'] or '',
                    }
                )
                session['budget_client'] = bc
            else:
                real_doc = save_or_update_client(doc, name, phone, lastname=lastname)
            bc = dict(session.get('budget_client', {}))
            bc['client_doc'] = real_doc
            bc['client_name'] = f"{name} {lastname}".strip() or "SIN NOMBRE"
            bc['client_lastname'] = lastname
            bc['client_phone'] = phone
            session['budget_client'] = bc
            flash('✓ Cliente guardado.', 'success')
        except Exception as e:
            flash(f'Error: {e}', 'danger')

    elif step == 'vehicle':
        bc = dict(session.get('budget_client', {}))
        doc = up(request.form.get('client_doc', '')) or bc.get('client_doc', '')
        name = bc.get('client_name', '')
        lastname = bc.get('client_lastname', '')
        if not doc:
            flash('Primero guarda el cliente.', 'danger')
            return redirect(url_for('index', module='📝 Nuevo Presupuesto'))
        veh_name = request.form.get('veh_name', '').strip()
        veh_brand = up(request.form.get('veh_brand', ''))
        veh_model = up(request.form.get('veh_model', ''))
        veh_color = up(request.form.get('veh_color', ''))
        veh_year = request.form.get('veh_year', '').strip()
        veh_plate = up(request.form.get('veh_plate', ''))
        if not any([veh_name, veh_brand, veh_model, veh_color, veh_year, veh_plate]):
            flash('No se ingresó información del vehículo.', 'danger')
            return redirect(url_for('index', module='📝 Nuevo Presupuesto'))
        try:
            save_or_update_client(
                doc,
                name,
                bc.get('client_phone', ''),
                veh_name,
                veh_year,
                veh_plate,
                veh_brand,
                veh_color,
                veh_model,
                lastname=lastname,
            )
            bc.update(
                {
                    'veh_name': veh_name,
                    'veh_brand': veh_brand,
                    'veh_model': veh_model,
                    'veh_color': veh_color,
                    'veh_year': veh_year,
                    'veh_plate': veh_plate,
                }
            )
            session['budget_client'] = bc
            flash('✓ Vehículo registrado.', 'success')
        except Exception as e:
            flash(f'Error: {e}', 'danger')

    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


def save_temp_budget_client(form_data):
    current = dict(session.get('budget_client', {}))
    for key in [
        'client_name',
        'client_lastname',
        'client_doc',
        'client_phone',
        'veh_name',
        'veh_year',
        'veh_plate',
        'veh_brand',
        'veh_color',
        'veh_model',
        'budget_note',
        'discount_percent',
        'include_iva',
    ]:
        if key in form_data:
            current[key] = form_data.get(key, '')
    session['budget_client'] = current


@app.route('/cart/add_product', methods=['POST'])
@block_budget_access
def cart_add_product():
    cart = session.get('cart', [])
    if len(cart) >= 100:
        flash('El carrito alcanzó el límite.', 'danger')
        return redirect(url_for('index', module='📝 Nuevo Presupuesto'))
    save_temp_budget_client(request.form)
    code = request.form.get('p_code', '').strip()
    name = request.form.get('p_name', '').strip()
    cost_price = max(0.0, safe_float(request.form.get('p_cost'), 0.0))
    margin_pct = max(0.0, safe_float(request.form.get('p_margin'), 30.0))
    discount = max(0.0, min(100.0, safe_float(request.form.get('p_discount'), 0.0)))
    qty = max(1, safe_int(request.form.get('qty'), 1))
    user_price = safe_float(request.form.get('p_price'), 0.0)
    selling_price = (
        user_price if user_price > 0 else cost_price * (1 + (margin_pct / 100.0))
    )

    item = {
        "type": "PRODUCTO",
        "code": code if code else "N/A",
        "name": name,
        "qty": qty,
        "cost": float(cost_price),
        "unit_price": float(selling_price),
        "discount": float(discount),
        "profit": 0.0,
        "total": 0.0,
    }
    recalc_cart_item(item)
    cart.append(item)
    session['cart'] = cart
    totals = recalc_cart(cart)
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify(
            {
                'ok': True,
                'message': 'Producto agregado.',
                'cart': cart,
                'subtotal': totals['net'],
                'totals': totals,
            }
        )
    flash('Producto agregado.', 'success')
    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


@app.route('/cart/add_labor', methods=['POST'])
@block_budget_access
def cart_add_labor():
    save_temp_budget_client(request.form)
    desc = request.form.get('labor_desc')
    qty = max(1, safe_int(request.form.get('q_labor'), 1))
    price = safe_float(request.form.get('labor_price'), 0.0)
    discount = max(0.0, min(100.0, safe_float(request.form.get('labor_discount'), 0.0)))
    moneda_labor = request.form.get('labor_curr', 'USD')
    moneda_base = session.get('budget_currency', 'USD')
    if desc:
        tasa_u = get_config('tasa_usd')
        tasa_e = get_config('tasa_eur')
        precio = price
        if moneda_labor != moneda_base:
            if (
                moneda_base == "USD"
                and moneda_labor == "EUR"
                and tasa_u > 0
                and tasa_e > 0
            ):
                precio = price * (tasa_e / tasa_u)
            elif (
                moneda_base == "EUR"
                and moneda_labor == "USD"
                and tasa_e > 0
                and tasa_u > 0
            ):
                precio = price * (tasa_u / tasa_e)
        item = {
            "type": "LABOR",
            "code": "SERV",
            "name": desc,
            "qty": qty,
            "cost": 0.0,
            "unit_price": float(precio),
            "discount": float(discount),
            "profit": 0.0,
            "total": 0.0,
        }
        recalc_cart_item(item)
        cart = session.get('cart', [])
        cart.append(item)
        session['cart'] = cart
        totals = recalc_cart(cart)
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify(
                {
                    'ok': True,
                    'message': 'Servicio agregado.',
                    'cart': cart,
                    'subtotal': totals['net'],
                    'totals': totals,
                }
            )
    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


@app.route('/cart/remove/<int:index>', methods=['POST'])
@block_budget_access
def cart_remove(index):
    cart = session.get('cart', [])
    if 0 <= index < len(cart):
        cart.pop(index)
        session['cart'] = cart
    totals = recalc_cart(cart)
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify(
            {'ok': True, 'cart': cart, 'subtotal': totals['net'], 'totals': totals}
        )
    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


@app.route('/cart/set_discount/<int:index>', methods=['POST'])
@block_budget_access
def cart_set_discount(index):
    cart = session.get('cart', [])
    if 0 <= index < len(cart):
        try:
            disc = float(request.form.get('discount', 0) or 0)
        except (ValueError, TypeError):
            disc = 0.0
        disc = max(0.0, min(100.0, disc))
        cart[index]['discount'] = disc
        recalc_cart_item(cart[index])
        session['cart'] = cart
    totals = recalc_cart(cart)
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify(
            {'ok': True, 'cart': cart, 'subtotal': totals['net'], 'totals': totals}
        )
    return redirect(url_for('index', module='📝 Nuevo Presupuesto'))


@app.route('/budget/save', methods=['POST'])
@block_budget_access
def save_budget():
    if not session.get('logged_in') or not session.get('username'):
        flash('⚠️ Tu sesión expiró. Inicia sesión nuevamente.', 'danger')
        return redirect(url_for('login'))

    current_username = session.get('username', 'Sistema')
    cart = session.get('cart', [])

    if not cart:
        flash('El presupuesto está vacío.', 'danger')
        return redirect(url_for('index', module='📝 Nuevo Presupuesto'))

    client = request.form.get('client_name', '').strip()
    client_lastname = request.form.get('client_lastname', '').strip()
    raw_doc = request.form.get('client_doc', '')
    phone = request.form.get('client_phone', '').strip()

    is_draft = not (client or client_lastname)

    if is_draft:
        full_client = "SIN ASIGNAR"
        doc = None
    else:
        full_client = f"{client} {client_lastname}".strip()
        doc = normalize_client_doc(raw_doc)
        if not doc:
            flash(
                '❌ Cédula/RIF inválida. Formatos válidos: V-12345678 · E-12345678 · J-40123456-7.',
                'danger',
            )
            return redirect(url_for('index', module='📝 Nuevo Presupuesto'))
        try:
            doc = save_or_update_client(doc, client, phone, lastname=client_lastname)
        except Exception as e:
            flash(f'Error cliente: {e}', 'danger')
            return redirect(url_for('index', module='📝 Nuevo Presupuesto'))

    veh_name = request.form.get('veh_name', '').strip()
    veh_year = request.form.get('veh_year', '').strip()
    veh_plate = up(request.form.get('veh_plate', ''))
    veh_brand = up(request.form.get('veh_brand', ''))
    veh_color = up(request.form.get('veh_color', ''))
    veh_model = up(request.form.get('veh_model', ''))

    currency = session.get('budget_currency', 'USD')
    tasa_bcv = safe_float(
        get_config('tasa_usd') if currency == 'USD' else get_config('tasa_eur'), 0.0
    )
    budget_note = request.form.get('budget_note', '').strip()

    totals = recalc_cart(cart)
    base_imponible = totals['net']
    discount_amount = totals['discount']
    discount_percent = totals['effective_percent']
    include_iva = 1 if request.form.get('include_iva') == '1' else 0
    show_bs_equivalent = 1 if request.form.get('show_bs_equivalent') == '1' else 0
    iva_amount = base_imponible * 0.16 if include_iva == 1 else 0.0
    total_divisa = base_imponible + iva_amount
    raw_profit = sum(item.get('profit', item.get('total', 0.0)) for item in cart)
    ganancia_total = max(0.0, raw_profit)
    total_ves = total_divisa * tasa_bcv
    items_json = json.dumps(cart)
    fecha_actual = get_ve_time().strftime("%Y-%m-%d %H:%M")
    public_token = generate_public_token()

    conn = get_db_connection()
    year_str = get_ve_time().strftime("%Y")
    last_budget = conn.execute(
        "SELECT budget_number FROM saved_budgets WHERE budget_number LIKE ? ORDER BY id DESC LIMIT 1",
        (f"COT-{year_str}-%",),
    ).fetchone()

    next_seq = 1
    if last_budget and last_budget['budget_number']:
        try:
            next_seq = int(last_budget['budget_number'].split('-')[-1]) + 1
        except ValueError:
            pass
    new_budget_num = f"COT-{year_str}-{next_seq:04d}"

    # FIX BUG #5: Incluir budget_number en el INSERT
    cursor = conn.execute(
        """
        INSERT INTO saved_budgets (
            date, client_name, client_doc, client_phone, currency, bcv_rate,
            total_foreign, total_ves, total_profit_foreign, items_json, status,
            vehicle_name, vehicle_year, vehicle_plate,
            vehicle_brand, vehicle_color, vehicle_model,
            username, budget_note,
            base_imponible, include_iva, iva_amount, discount_amount, discount_percent,
            total_paid_foreign, total_paid_ves, public_token,
            client_signature, client_signature_date, show_bs_equivalent, budget_number
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, 0.0, ?, '', '', ?, ?)
    """,
        (
            fecha_actual,
            full_client,
            doc or '',
            phone,
            currency,
            tasa_bcv,
            total_divisa,
            total_ves,
            ganancia_total,
            items_json,
            'Borrador' if is_draft else 'Pendiente',
            veh_name,
            veh_year,
            veh_plate,
            veh_brand,
            veh_color,
            veh_model,
            current_username,
            budget_note,
            base_imponible,
            include_iva,
            iva_amount,
            discount_amount,
            discount_percent,
            public_token,
            show_bs_equivalent,
            new_budget_num,
        ),
    )
    budget_id = cursor.lastrowid
    conn.commit()
    conn.close()

    session['cart'] = []
    session['budget_client'] = {}

    if is_draft:
        log_action(
            'Creación de borrador',
            f'Borrador #{budget_id} sin cliente asignado.',
            'Presupuesto',
            budget_id,
        )
        flash(
            f'💾 Borrador #{budget_id} guardado. Asígnale un cliente desde el historial.',
            'success',
        )
    else:
        log_action(
            'Creación de presupuesto',
            f'Presupuesto #{budget_id} para {full_client}.',
            'Presupuesto',
            budget_id,
        )
        flash(f'✓ Presupuesto #{budget_id} guardado para {full_client}.', 'success')

    return redirect(url_for('index', module='📂 Presupuestos'))


@app.route('/client/add', methods=['POST'])
def add_client():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    raw_doc = request.form.get('client_doc', '')
    name = request.form.get('client_name', '').strip()
    lastname = request.form.get('client_lastname', '').strip()
    phone = request.form.get('client_phone', '').strip()

    # === Validación 1: nombre o apellido ===
    if not name and not lastname:
        flash('Indica al menos el nombre o apellido.', 'danger')
        return redirect(url_for('index', module='📇 Clientes'))

    # === Validación 2: cédula obligatoria y con formato válido ===
    doc = normalize_client_doc(raw_doc)
    if not doc:
        flash(
            '❌ La cédula/RIF es obligatoria. Formatos válidos: '
            'V-12345678 · E-12345678 · J-40123456-7 · G-12345678-9 · P-1234567.',
            'danger',
        )
        return redirect(url_for('index', module='📇 Clientes'))

    # === Validación 3: duplicado exacto por cédula ===
    existing = find_client_by_doc(doc)
    if existing:
        flash(
            f'⚠️ Ya existe un cliente con la cédula {doc}: '
            f'{existing["client_name"]}. Búscalo en la lista y edítalo allí.',
            'danger',
        )
        return redirect(url_for('index', module='📇 Clientes'))

    # === Validación 4: mismo nombre+apellido con otra cédula (solo aviso) ===
    similar = find_client_similar(name, lastname, exclude_doc=doc)
    if similar:
        flash(
            f'⚠️ Aviso: ya existe "{similar[0]["client_name"]}" con cédula '
            f'{similar[0]["client_doc"]}. Se creó un cliente NUEVO con '
            f'{doc}. Si fue un error, elimina el duplicado manualmente.',
            'warning',
        )

    # === Insertar ===
    try:
        save_or_update_client(doc, name, phone, lastname=lastname)
        flash(f'✓ Cliente {name} {lastname} ({doc}) registrado.', 'success')
        log_action(
            'Creación de cliente', f'Cliente {name} {lastname} ({doc}).', 'Cliente'
        )
    except Exception as e:
        flash(f'Error al guardar: {e}', 'danger')

    # Redirigir al expediente recién creado
    try:
        return redirect(
            url_for('index', module='📁 Expedientes', selected_client_doc=doc)
        )
    except Exception:
        return redirect(url_for('index', module='📁 Expedientes'))


@app.route('/client/edit/<int:id>', methods=['POST'])
def edit_client(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    doc = up(request.form.get('client_doc', ''))
    name = request.form.get('client_name', '').strip()
    lastname = request.form.get('client_lastname', '').strip()
    phone = request.form.get('client_phone', '').strip()
    if not doc:
        doc = f"SIN_DOC_{int(time.time())}_{secrets.token_hex(3).upper()}"
    full_name = f"{name} {lastname}".strip() or "SIN NOMBRE"
    conn = get_db_connection()
    conn.execute(
        "UPDATE clients SET client_doc = ?, client_name = ?, client_lastname = ?, client_phone = ? WHERE id = ?",
        (doc, full_name, lastname, phone, id),
    )
    conn.commit()
    conn.close()
    flash('Cliente actualizado.', 'success')
    return redirect(url_for('index', module='📁 Expedientes', selected_client_doc=doc))


@app.route('/client/delete/<int:id>', methods=['POST'])
def delete_client(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute("SELECT client_doc FROM clients WHERE id = ?", (id,)).fetchone()
    if row:
        conn.execute("DELETE FROM vehicles WHERE client_doc = ?", (row['client_doc'],))
    conn.execute("DELETE FROM clients WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    flash('Cliente eliminado.', 'success')
    return redirect(url_for('index', module='📁 Expedientes'))


@app.route('/vehicle/add', methods=['POST'])
def add_vehicle():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    client_doc = up(request.form.get('client_doc', ''))
    if not client_doc:
        flash('Selecciona un cliente.', 'danger')
        return redirect(url_for('index', module='🚗 Vehículos'))
    vid = save_client_vehicle(
        client_doc,
        request.form.get('veh_name', '').strip(),
        request.form.get('veh_year', '').strip(),
        up(request.form.get('veh_plate', '')),
        up(request.form.get('veh_brand', '')),
        up(request.form.get('veh_color', '')),
        up(request.form.get('veh_model', '')),
        vin=request.form.get('veh_vin', '').strip(),
        km=request.form.get('veh_km', '').strip(),
        notes=request.form.get('veh_notes', '').strip(),
    )
    flash(
        'Vehículo registrado.' if vid else 'Sin datos.', 'success' if vid else 'danger'
    )
    return redirect(
        url_for('index', module='📁 Expedientes', selected_client_doc=client_doc)
    )


@app.route('/vehicle/edit/<int:vid>', methods=['POST'])
def edit_vehicle(vid):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    veh = conn.execute("SELECT * FROM vehicles WHERE id = ?", (vid,)).fetchone()
    if not veh:
        conn.close()
        flash('Vehículo no encontrado.', 'danger')
        return redirect(url_for('index', module='📁 Expedientes'))

    # Conservar VIN y notas si el formulario no los trae
    vin_val = request.form.get('veh_vin')
    if vin_val is None:
        vin_val = veh['vin'] or ''
    notes_val = request.form.get('veh_notes')
    if notes_val is None:
        notes_val = veh['notes'] or ''

    conn.execute(
        """
        UPDATE vehicles SET client_doc = ?, plate = ?, name = ?, brand = ?, model = ?,
            year = ?, color = ?, vin = ?, km = ?, notes = ?, updated_at = ? WHERE id = ?
    """,
        (
            up(request.form.get('client_doc', veh['client_doc'])),
            up(request.form.get('veh_plate', '')),
            request.form.get('veh_name', '').strip(),
            up(request.form.get('veh_brand', '')),
            up(request.form.get('veh_model', '')),
            request.form.get('veh_year', '').strip(),
            up(request.form.get('veh_color', '')),
            vin_val.strip(),
            request.form.get('veh_km', '').strip(),
            notes_val.strip(),
            get_ve_time().strftime("%Y-%m-%d %H:%M"),
            vid,
        ),
    )
    conn.commit()
    conn.close()
    flash('Vehículo actualizado.', 'success')
    target_doc = up(
        request.form.get('client_doc', veh['client_doc'] if veh else '')
    ) or (veh['client_doc'] if veh else '')
    return redirect(
        url_for('index', module='📁 Expedientes', selected_client_doc=target_doc)
    )


@app.route('/vehicle/delete/<int:vid>', methods=['POST'])
def delete_vehicle(vid):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute("DELETE FROM vehicles WHERE id = ?", (vid,))
    conn.commit()
    conn.close()
    flash('Vehículo eliminado.', 'success')
    return redirect(request.referrer or url_for('index', module='📁 Expedientes'))


@app.route('/vehicle/set_primary/<int:vid>', methods=['POST'])
def set_primary_vehicle(vid):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    veh = conn.execute("SELECT * FROM vehicles WHERE id = ?", (vid,)).fetchone()
    if veh:
        conn.execute(
            """UPDATE clients SET vehicle_name = ?, vehicle_year = ?, vehicle_plate = ?,
                        vehicle_brand = ?, vehicle_color = ?, vehicle_model = ? WHERE client_doc = ?""",
            (
                veh['name'],
                veh['year'],
                veh['plate'],
                veh['brand'],
                veh['color'],
                veh['model'],
                veh['client_doc'],
            ),
        )
        conn.commit()
        flash('Vehículo marcado como principal.', 'success')
    conn.close()
    return redirect(request.referrer or url_for('index', module='📁 Expedientes'))


@app.route('/revision/save', methods=['POST'])
@block_readonly_reception
def save_revision():
    """Crea una recepción con datos mínimos y redirige a la planilla completa."""
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    try:
        client = request.form.get('client_name', '').strip()
        client_lastname = request.form.get('client_lastname', '').strip()
        raw_doc = request.form.get('client_doc', '')
        phone = request.form.get('client_phone', '').strip()
        full_client = f"{client} {client_lastname}".strip() or "SIN NOMBRE"

        if not client and not client_lastname:
            flash('❌ Indica al menos el nombre o apellido del cliente.', 'danger')
            return redirect(url_for('index', module='📥 Recepciones'))

        doc = normalize_client_doc(raw_doc)
        if not doc:
            flash(
                '❌ La cédula/RIF es obligatoria. Formatos válidos: V-12345678 · E-12345678 · J-40123456-7.',
                'danger',
            )
            return redirect(url_for('index', module='📥 Recepciones'))

        veh_name = request.form.get('veh_name', '').strip()
        veh_year = request.form.get('veh_year', '').strip()
        veh_plate = up(request.form.get('veh_plate', ''))
        veh_km = request.form.get('veh_km', '').strip()
        veh_brand = up(request.form.get('veh_brand', ''))
        veh_color = up(request.form.get('veh_color', ''))
        veh_model = up(request.form.get('veh_model', ''))
        note = request.form.get('revision_note', '').strip()

        try:
            doc = save_or_update_client(
                doc,
                client,
                phone,
                veh_name,
                veh_year,
                veh_plate,
                veh_brand,
                veh_color,
                veh_model,
                lastname=client_lastname,
            )
        except Exception as e:
            print(f"[save_revision] Error save_or_update_client: {e}")
            flash(f'Error al guardar el cliente: {e}', 'danger')
            return redirect(url_for('index', module='📥 Recepciones'))

        # Items iniciales vacíos
        initial_items = [{'sys': f, 'mark': ''} for f in RECEPTION_FIELDS_FLAT]

        fecha = get_ve_time().strftime("%Y-%m-%d %H:%M")
        fecha_only = get_ve_time().strftime("%Y-%m-%d")
        hora_str = get_ve_time().strftime("%H:%M")
        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
        public_token = generate_public_token()

        conn = get_db_connection()
        try:
            # ===== Detectar columnas NOT NULL "extra" que no controlamos =====
            cur_schema = conn.cursor()
            cur_schema.execute("PRAGMA table_info(revisions)")
            schema_cols = cur_schema.fetchall()

            known_columns = {
                'id',
                'date',
                'client_name',
                'client_doc',
                'client_phone',
                'vehicle_name',
                'vehicle_year',
                'vehicle_plate',
                'vehicle_km',
                'vehicle_brand',
                'vehicle_color',
                'vehicle_model',
                'items_json',
                'status',
                'username',
                'revision_note',
                'reception_items_json',
                'reception_fallas_json',
                'reception_razon_social',
                'reception_rif_ci',
                'reception_email',
                'reception_direccion_fiscal',
                'reception_es_retorno',
                'reception_fecha',
                'reception_hora',
                'reception_am_pm',
                'reception_transmision',
                'reception_motor',
                'reception_clave',
                'reception_status',
                'public_token',
                'mechanic_items_json',
                'mechanic_notes',
                'mechanic_username',
                'mechanic_date',
                'mechanic_status',
                'vehicle_id',
            }

            extra_not_null = []
            for c in schema_cols:
                c_name = c[1]
                c_notnull = c[3]
                c_default = c[4]
                if c_name not in known_columns and c_notnull == 1 and c_default is None:
                    extra_not_null.append(c_name)

            print(
                f"[save_revision] Columnas extra NOT NULL detectadas: {extra_not_null}"
            )

            # ===== Construir INSERT dinámico =====
            base_columns = [
                'date',
                'client_name',
                'client_doc',
                'client_phone',
                'vehicle_name',
                'vehicle_year',
                'vehicle_plate',
                'vehicle_km',
                'vehicle_brand',
                'vehicle_color',
                'vehicle_model',
                'items_json',
                'status',
                'username',
                'revision_note',
                'reception_items_json',
                'reception_fallas_json',
                'reception_razon_social',
                'reception_rif_ci',
                'reception_email',
                'reception_direccion_fiscal',
                'reception_es_retorno',
                'reception_fecha',
                'reception_hora',
                'reception_am_pm',
                'reception_transmision',
                'reception_motor',
                'reception_clave',
                'reception_status',
                'public_token',
            ]

            base_values = [
                fecha,
                full_client,
                doc,
                phone,
                veh_name,
                veh_year,
                veh_plate,
                veh_km,
                veh_brand,
                veh_color,
                veh_model,
                json.dumps(
                    [
                        {'system': it['sys'], 'status': 'Bueno / Óptimo', 'obs': ''}
                        for it in initial_items
                    ]
                ),
                'Pendiente',
                session['username'],
                note,
                json.dumps(initial_items),
                '[]',
                '',
                '',
                '',
                '',
                0,
                fecha_only,
                hora_str,
                '',
                '',
                '',
                '',
                'Borrador',
                public_token,
            ]

            # Columnas extra (drift del schema): dar valor seguro
            for ec in extra_not_null:
                # Heurística: nombres con "json" o "data" → '[]'; el resto → ''
                val = '[]' if ('json' in ec.lower() or 'data' in ec.lower()) else ''
                base_columns.append(ec)
                base_values.append(val)

            placeholders = ', '.join(['?'] * len(base_columns))
            sql = f"INSERT INTO revisions ({', '.join(base_columns)}) VALUES ({placeholders})"

            cursor = conn.execute(sql, base_values)
            new_id = cursor.lastrowid

            # ===== Crear tarjeta automáticamente en el Tablero Taller =====
            try:
                # Título: "MARCA MODELO AÑO · PLACA"
                title_parts = [p for p in [veh_brand, veh_model, veh_year] if p]
                card_title = ' '.join(title_parts).strip()
                if veh_plate:
                    card_title = (
                        (card_title + ' · ' + veh_plate) if card_title else veh_plate
                    )
                if not card_title:
                    card_title = full_client

                # Checklist inicial
                checklist_items = [
                    {
                        "text": "REVISIÓN Y LIMPIEZA DE VOLANTE Y PIEZAS MANCHADAS + CALIBRACIÓN NEUMÁTICA Y AJUSTE DE NEUMÁTICOS CON LLAVE DE FORMA MANUAL",
                        "done": False,
                    }
                ]

                description_lines = [f"📥 Recepción #{new_id}"]
                if note:
                    description_lines.append(note)
                card_description = '\n'.join(description_lines)

                now_card = get_ve_time().strftime("%Y-%m-%d %H:%M")
                today_card = get_ve_time().strftime("%Y-%m-%d")

                cursor2 = conn.execute(
                    """
                    INSERT INTO workshop_cards (
                        column_id, sort_order, client_doc, client_name, client_phone,
                        vehicle_plate, vehicle_name, vehicle_brand, vehicle_model, vehicle_year,
                        assigned_to, priority, entry_date, due_date, notes, description,
                        created_at, updated_at, created_by, is_delivered, checklist_json, status,
                        title
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, 'taller', ?)
                """,
                    (
                        0,
                        0,
                        doc,
                        full_client,
                        phone,
                        veh_plate,
                        veh_name,
                        veh_brand,
                        veh_model,
                        veh_year,
                        '',  # assigned_to vacío → sin asignar
                        'Media',
                        today_card,
                        '',  # due_date
                        '',  # notes
                        card_description,
                        now_card,
                        now_card,
                        session['username'],
                        json.dumps(checklist_items),
                        card_title,
                    ),
                )
                card_id = cursor2.lastrowid

                print(
                    f"[save_revision] ✓ Tarjeta #{card_id} creada en Tablero Taller (sin asignar)"
                )
            except Exception as e:
                print(f"[save_revision] Error creando tarjeta taller: {e}")

            conn.commit()
        except Exception as e:
            conn.rollback()
            conn.close()
            import traceback

            traceback.print_exc()
            print(f"[save_revision] Error en INSERT: {e}")
            flash(f'Error al crear recepción: {e}', 'danger')
            return redirect(url_for('index', module='📥 Recepciones'))
        conn.close()

        log_action(
            'Creación de recepción',
            f'Recepción #{new_id} para {full_client} ({doc}).',
            'Recepción',
            new_id,
        )
        flash(f'✓ Recepción #{new_id} creada. Completa la planilla.', 'success')
        return redirect(url_for('reception_edit', reception_id=new_id))

    except Exception as e:
        import traceback

        traceback.print_exc()
        print(f"[save_revision] Error inesperado: {e}")
        flash(f'Error inesperado: {e}', 'danger')
        return redirect(url_for('index', module='📥 Recepciones'))


@app.route('/revision/status/<int:id>/<status>', methods=['POST'])
@block_readonly_reception
def change_revision_status(id, status):
    conn = get_db_connection()
    conn.execute("UPDATE revisions SET status = ? WHERE id = ?", (status, id))
    conn.commit()
    conn.close()
    flash(f'Estado: {status}.', 'success')
    return redirect(request.referrer or url_for('index', module='📥 Recepciones'))


@app.route('/revision/delete/<int:id>', methods=['POST'])
@block_readonly_reception
def delete_revision(id):
    conn = get_db_connection()
    r = conn.execute("SELECT * FROM revisions WHERE id = ?", (id,)).fetchone()
    if r and (
        session.get('role') == 'Administrador'
        or r['username'] == session.get('username')
    ):

        # 1. Borrar fotos del disco duro
        try:
            folder_path = os.path.join(RECEPTION_PHOTOS_DIR, str(id))
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
        except Exception as e:
            print(f"[delete_revision] Error borrando fotos: {e}")

        # 2. Borrar registros de fotos en DB
        conn.execute("DELETE FROM reception_photos WHERE reception_id = ?", (id,))

        # 3. Borrar la tarjeta del Tablero Taller (si existe y tiene la misma placa/cliente)
        if r['vehicle_plate']:
            conn.execute(
                """
                DELETE FROM workshop_cards
                WHERE UPPER(TRIM(vehicle_plate)) = UPPER(TRIM(?))
                  AND client_doc = ?
            """,
                (r['vehicle_plate'], r['client_doc']),
            )

        # 4. Borrar la revisión
        conn.execute("DELETE FROM revisions WHERE id = ?", (id,))
        conn.commit()
        flash('Revisión y sus fotos eliminadas.', 'success')
    conn.close()
    return redirect(request.referrer or url_for('index', module='📥 Recepciones'))


@app.route('/budget/pdf/<int:id>')
def download_budget_pdf(id):
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    payments = (
        conn.execute(
            "SELECT * FROM payments WHERE budget_id = ? ORDER BY id ASC", (id,)
        ).fetchall()
        if row
        else []
    )

    # Usar el número guardado, o generarlo si es un presupuesto antiguo
    if row:
        budget_number = (
            row['budget_number']
            if 'budget_number' in row.keys() and row['budget_number']
            else get_budget_number(conn, row['id'], row['date'])
        )
    else:
        budget_number = ""

    conn.close()
    if row:
        cart = json.loads(row['items_json'])
        pdf_bytes = generate_pdf(
            row['client_name'],
            row['client_doc'],
            row['client_phone'],
            row['vehicle_name'],
            row['vehicle_year'],
            row['vehicle_plate'],
            row['date'],
            cart,
            row['currency'],
            float(row['bcv_rate']),
            float(row['total_foreign']),
            float(row['total_ves']),
            row['budget_note'],
            username=row['username'],
            base_imponible=float(row['base_imponible'] or row['total_foreign']),
            include_iva=bool(row['include_iva']),
            iva_amount=float(row['iva_amount'] or 0.0),
            discount_amount=float(row['discount_amount'] or 0.0),
            discount_percent=float(row['discount_percent'] or 0.0),
            payments=payments,
            payment_status=compute_payment_status(row),
            signature_b64=row['client_signature'] or '',
            signature_date=row['client_signature_date'] or '',
            veh_brand=row['vehicle_brand'] if 'vehicle_brand' in row.keys() else '',
            veh_color=row['vehicle_color'] if 'vehicle_color' in row.keys() else '',
            veh_model=row['vehicle_model'] if 'vehicle_model' in row.keys() else '',
            budget_status=row['status'],
            show_bs_equivalent=bool(
                row['show_bs_equivalent'] if 'show_bs_equivalent' in row.keys() else 0
            ),
            budget_number=budget_number,
        )  # <-- PASAR EL NÚMERO AQUÍ
        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Presupuesto_{row['client_name']}.pdf",
        )
    return "No encontrado", 404


import os
from io import BytesIO
from flask import send_file, flash, redirect, url_for, session


@app.route('/budget/export_galac/<int:id>')
def export_galac(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    budget = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    if not budget:
        conn.close()
        flash('Presupuesto no encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))

    client_lastname = ''
    if budget['client_doc']:
        c = conn.execute(
            "SELECT client_lastname FROM clients WHERE client_doc = ?",
            (budget['client_doc'],),
        ).fetchone()
        if c:
            client_lastname = c['client_lastname'] or ''

    try:
        items = json.loads(budget['items_json'] or '[]')
    except Exception:
        items = []
    finally:
        conn.close()

    lineas = []

    num_factura = str(budget['id']).zfill(8)
    fecha = budget['date'][:10]

    try:
        fecha_dt = datetime.datetime.strptime(fecha, "%Y-%m-%d")
        fecha_galac = fecha_dt.strftime("%d/%m/%Y")
    except Exception:
        fecha_galac = fecha

    cod_cliente = budget['client_doc'] or 'V-00000000'
    cod_vendedor = '01'
    observaciones = f"{budget['client_name']} {client_lastname}".strip()

    base_imponible = float(budget['base_imponible'] or 0)
    iva = float(budget['iva_amount'] or 0)
    total_factura = float(budget['total_foreign'] or 0)
    total_renglones = sum(float(it.get('total', 0)) for it in items)

    def fmt_num(n):
        return f"{n:.2f}"

    moneda_galac = 'DS' if budget['currency'] == 'USD' else 'BS'
    status_galac = (
        'BOR'
        if budget['status'] == 'Borrador'
        else ('ANU' if budget['status'] == 'Anulado' else 'EMI')
    )
    tipo_doc = 'FAC'

    cabecera = [
        num_factura,
        fecha_galac,
        cod_cliente,
        cod_vendedor,
        observaciones,
        "0.00",
        fmt_num(base_imponible),
        fmt_num(total_renglones),
        fmt_num(iva),
        fmt_num(total_factura),
        "0.00",
        moneda_galac,
        "PR1",
        "UNICO",
        status_galac,
        tipo_doc,
    ]

    lineas.append("\t".join(cabecera))

    for it in items:
        descripcion = it.get('name', '')
        cantidad = it.get('qty', 1)
        precio_sin_iva = float(it.get('unit_price', 0))
        alicuota = "16.00" if float(budget['iva_amount'] or 0) > 0 else "0.00"
        cod_articulo = it.get('code', 'SERV')

        renglon = [
            descripcion,
            alicuota,
            fmt_num(cantidad),
            fmt_num(precio_sin_iva),
            fmt_num(precio_sin_iva * (1 + float(alicuota) / 100)),
            "0.00",
            fmt_num(precio_sin_iva * cantidad),
            "100.00",
            "",
            "",
            "",
            cod_articulo,
        ]
        lineas.append("\t".join(renglon))

    contenido = "\r\n".join(lineas)

    try:
        contenido_bytes = contenido.encode('ascii', errors='replace')
    except Exception:
        contenido_bytes = contenido.encode('latin-1', errors='replace')

    filename = f"Galac_Factura_{budget['id']}.txt"

    log_action(
        'Exportación a Galac',
        f'Presupuesto #{id} exportado para Galac 30.',
        'Presupuesto',
        id,
    )

    return send_file(
        BytesIO(contenido_bytes),
        mimetype='text/plain',
        as_attachment=True,
        download_name=filename,
    )


@app.route('/budget/preview/<int:id>')
def preview_budget_pdf(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    payments = (
        conn.execute(
            "SELECT * FROM payments WHERE budget_id = ? ORDER BY id ASC", (id,)
        ).fetchall()
        if row
        else []
    )

    # FIX BUG #8: Obtener el número de cotización
    budget_number = ""
    if row:
        budget_number = (
            row['budget_number']
            if 'budget_number' in row.keys() and row['budget_number']
            else get_budget_number(conn, row['id'], row['date'])
        )

    conn.close()

    if row:
        cart = json.loads(row['items_json'])
        pdf_bytes = generate_pdf(
            row['client_name'],
            row['client_doc'],
            row['client_phone'],
            row['vehicle_name'],
            row['vehicle_year'],
            row['vehicle_plate'],
            row['date'],
            cart,
            row['currency'],
            float(row['bcv_rate']),
            float(row['total_foreign']),
            float(row['total_ves']),
            row['budget_note'],
            username=row['username'],
            base_imponible=float(row['base_imponible'] or row['total_foreign']),
            include_iva=bool(row['include_iva']),
            iva_amount=float(row['iva_amount'] or 0.0),
            discount_amount=float(row['discount_amount'] or 0.0),
            discount_percent=float(row['discount_percent'] or 0.0),
            payments=payments,
            payment_status=compute_payment_status(row),
            signature_b64=row['client_signature'] or '',
            signature_date=row['client_signature_date'] or '',
            veh_brand=row['vehicle_brand'] if 'vehicle_brand' in row.keys() else '',
            veh_color=row['vehicle_color'] if 'vehicle_color' in row.keys() else '',
            veh_model=row['vehicle_model'] if 'vehicle_model' in row.keys() else '',
            budget_status=row['status'],
            show_bs_equivalent=bool(
                row['show_bs_equivalent'] if 'show_bs_equivalent' in row.keys() else 0
            ),
            budget_number=budget_number,
        )  # FIX BUG #8
        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            download_name=f"Presupuesto_{row['client_name']}.pdf",
        )
    return "No encontrado", 404


def sync_budget_from_request(conn, editing_id):
    client_name = request.form.get('client_name')
    if client_name is None:
        return
    client_name = client_name.strip()
    client_lastname = request.form.get('client_lastname', '').strip()
    raw_doc = request.form.get('client_doc', '')
    client_phone = request.form.get('client_phone', '').strip()
    veh_name = request.form.get('veh_name', '').strip()
    veh_year = request.form.get('veh_year', '').strip()
    veh_plate = up(request.form.get('veh_plate', ''))
    veh_brand = up(request.form.get('veh_brand', ''))
    veh_color = up(request.form.get('veh_color', ''))
    veh_model = up(request.form.get('veh_model', ''))
    budget_note = request.form.get('budget_note', '').strip()
    status = request.form.get('status', 'Pendiente')

    budget_row = conn.execute(
        "SELECT status FROM saved_budgets WHERE id = ?", (editing_id,)
    ).fetchone()
    old_status = budget_row['status'] if budget_row else 'Pendiente'

    if client_name or client_lastname:
        full_client = f"{client_name} {client_lastname}".strip() or "SIN NOMBRE"
        normalized = normalize_client_doc(raw_doc)
        if normalized:
            client_doc = save_or_update_client(
                normalized,
                client_name,
                client_phone,
                veh_name,
                veh_year,
                veh_plate,
                veh_brand,
                veh_color,
                veh_model,
                lastname=client_lastname,
                conn=conn,
            )
        else:
            client_doc = ''
    else:
        full_client = "SIN ASIGNAR"
        client_doc = ''
        status = 'Borrador'

    item_codes = request.form.getlist('item_code')
    item_names = request.form.getlist('item_name')
    item_qtys = request.form.getlist('item_qty')
    item_prices = request.form.getlist('item_price')
    item_discounts = request.form.getlist('item_discount')

    budget_full = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (editing_id,)
    ).fetchone()
    existing_items = json.loads(budget_full['items_json']) if budget_full else []

    items = []
    for i in range(len(item_names)):
        code = item_codes[i].strip() if i < len(item_codes) else 'N/A'
        name = item_names[i].strip()
        qty = safe_int(item_qtys[i] if i < len(item_qtys) else None, 1)
        unit_price = safe_float(item_prices[i] if i < len(item_prices) else None, 0.0)
        item_disc = safe_float(
            item_discounts[i] if i < len(item_discounts) else None, 0.0
        )
        orig = existing_items[i] if i < len(existing_items) else {}
        new_item = {
            "type": orig.get('type', 'PRODUCTO'),
            "code": code,
            "name": name,
            "qty": qty,
            "cost": float(orig.get('cost', 0.0)),
            "unit_price": float(unit_price),
            "discount": float(item_disc),
            "profit": 0.0,
            "total": 0.0,
        }
        recalc_cart_item(new_item)
        items.append(new_item)

    totals = recalc_cart(items)
    base_imponible = totals['net']
    discount_amount = totals['discount']
    discount_percent = totals['effective_percent']
    include_iva = 1 if request.form.get('include_iva') == '1' else 0
    show_bs_equivalent = 1 if request.form.get('show_bs_equivalent') == '1' else 0
    iva_amount = base_imponible * 0.16 if include_iva == 1 else 0.0
    total_divisa = base_imponible + iva_amount

    raw_profit = sum(item.get('profit', item.get('total', 0.0)) for item in items)
    ganancia_total = max(0.0, raw_profit)

    tasa_bcv = float(budget_full['bcv_rate']) if budget_full else 1.0
    total_ves = total_divisa * tasa_bcv

    items_json = json.dumps(items)

    conn.execute(
        """
        UPDATE saved_budgets SET
            client_name = ?, client_doc = ?, client_phone = ?,
            vehicle_name = ?, vehicle_year = ?, vehicle_plate = ?,
            vehicle_brand = ?, vehicle_color = ?, vehicle_model = ?,
            budget_note = ?, status = ?, items_json = ?,
            total_foreign = ?, total_ves = ?, total_profit_foreign = ?,
            base_imponible = ?, include_iva = ?, iva_amount = ?,
            discount_amount = ?, discount_percent = ?,
            show_bs_equivalent = ?
        WHERE id = ?
    """,
        (
            full_client,
            client_doc,
            client_phone,
            veh_name,
            veh_year,
            veh_plate,
            veh_brand,
            veh_color,
            veh_model,
            budget_note,
            status,
            items_json,
            total_divisa,
            total_ves,
            ganancia_total,
            base_imponible,
            include_iva,
            iva_amount,
            discount_amount,
            discount_percent,
            show_bs_equivalent,
            editing_id,
        ),
    )

    if old_status != 'Aprobado' and status == 'Aprobado':
        update_stock_for_budget(conn, editing_id, 'deduct')
    elif old_status == 'Aprobado' and status != 'Aprobado':
        update_stock_for_budget(conn, editing_id, 'restore')

    conn.commit()


@app.route('/budget/edit/<int:id>', methods=['GET', 'POST'])
@block_budget_access
def edit_budget(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    budget = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    if not budget:
        conn.close()
        flash('No encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))
    if request.method == 'POST':
        try:
            sync_budget_from_request(conn, id)
            flash('¡Presupuesto actualizado!', 'success')
        except Exception as e:
            print(f"[edit_budget POST] Error: {e}")
            flash(f'Error al guardar el presupuesto: {e}', 'danger')
        finally:
            conn.close()

        return redirect(url_for('index', module='📂 Presupuestos'))

    is_draft = budget['client_name'] == 'SIN ASIGNAR'
    items = json.loads(budget['items_json'])
    for it in items:
        recalc_cart_item(it)
    curr_u = session['username']
    reminders_count = count_unseen_reminders(conn, curr_u)
    reminders_today_count = count_today_reminders(conn, curr_u)
    appointments_count = conn.execute(
        """SELECT COUNT(*) FROM appointments WHERE status IN ('Pendiente','Confirmada') AND date >= ?""",
        (get_ve_time().strftime("%Y-%m-%d"),),
    ).fetchone()[0]
    all_clients_rows = conn.execute(
        "SELECT * FROM clients ORDER BY client_name ASC"
    ).fetchall()
    all_clients = [dict(c) for c in all_clients_rows]

    client_lastname_db = ''
    if budget['client_doc']:
        c_row = conn.execute(
            "SELECT client_lastname FROM clients WHERE client_doc = ?",
            (budget['client_doc'],),
        ).fetchone()
        if c_row:
            client_lastname_db = c_row['client_lastname'] or ''

    conn.close()

    role = session['role']
    allowed = get_allowed_modules(role)
    menu_options = [m for m in ALL_SYSTEM_MODULES if m in allowed]

    return render_template(
        'base.html',
        menu_options=menu_options,
        current_module="📂 Presupuestos",
        username=session['username'],
        is_editing=True,
        is_editing_revision=False,
        budget=budget,
        items=items,
        is_draft_budget=is_draft,
        user_display_name=session.get('user_display_name', curr_u),
        user_phone=session.get('user_phone', ''),
        role=role,
        tasa_usd=get_config('tasa_usd'),
        tasa_eur=get_config('tasa_eur'),
        budget_currency=budget['currency'],
        budget_client={},
        all_modules=ALL_SYSTEM_MODULES,
        all_budgets=[],
        filtered_budgets=[],
        pending_budgets=[],
        seller_summary=[],
        reminders=[],
        reminders_count=reminders_count,
        reminders_today_count=reminders_today_count,
        reminder_categories=[],
        r_search='',
        r_priority='Todas',
        r_category='Todas',
        r_status='Pendientes',
        r_author='Todos',
        r_sort='recientes',
        appointments_count=appointments_count,
        all_users=[],
        users_map={},
        users=[],
        roles_list=[],
        all_clients=all_clients,
        c_search='',
        c_filter='Todos',
        c_sort='nombre',
        c_pagination=None,
        vehicle_list=[],
        v_search='',
        v_client='',
        v_sort='recientes',
        v_pagination=None,
        all_revisions=[],
        selected_client_doc='',
        client_budgets=[],
        client_revisions=[],
        f_type='Todos',
        f_user='Todos',
        f_date='',
        f_month='',
        f_start='',
        f_end='',
        current_date=get_ve_time().strftime("%Y-%m-%d"),
        templates=[],
        appointments=[],
        appointment_filter='Proximas',
        vehicle_search='',
        report_data=None,
        action_logs=[],
        action_types=[],
        log_user='Todos',
        log_action_filter='Todos',
        dashboard_metrics=None,
        revision_kpis={},
        recent_revisions=[],
        workshop_columns=[],
        cards_by_column={},
        workshop_stats=None,
        workshop_mechanics=[],
        r_pagination=None,
        h_pagination=None,
        budget_client_lastname=client_lastname_db,
        h_search='',  # FIX BUG #11
        f_status='Todos',
    )  # FIX BUG #11


@app.route('/budget/edit/add_product', methods=['POST'])
def edit_cart_add_product():
    editing_id = request.form.get('editing_id')
    if not editing_id:
        return redirect(url_for('index', module='📂 Presupuestos'))
    conn = get_db_connection()
    sync_budget_from_request(conn, editing_id)

    code = request.form.get('p_code', '').strip()
    name = request.form.get('p_name', '').strip()
    cost_price = max(0.0, safe_float(request.form.get('p_cost'), 0.0))
    margin_pct = max(0.0, safe_float(request.form.get('p_margin'), 30.0))
    discount = max(0.0, min(100.0, safe_float(request.form.get('p_discount'), 0.0)))
    qty = max(1, safe_int(request.form.get('qty'), 1))
    user_price = safe_float(request.form.get('p_price'), 0.0)
    selling_price = (
        user_price if user_price > 0 else cost_price * (1 + (margin_pct / 100.0))
    )

    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (editing_id,)
    ).fetchone()
    if not budget:
        conn.close()
        return redirect(url_for('index', module='📂 Presupuestos'))

    items = json.loads(budget['items_json'])
    ni = {
        "type": "PRODUCTO",
        "code": code if code else "N/A",
        "name": name,
        "qty": qty,
        "cost": float(cost_price),
        "unit_price": float(selling_price),
        "discount": float(discount),
        "profit": 0.0,
        "total": 0.0,
    }
    recalc_cart_item(ni)
    items.append(ni)
    totals = recalc_cart(items)
    include_iva = int(budget['include_iva'] or 0)
    base_imponible = totals['net']
    iva_amount = base_imponible * 0.16 if include_iva == 1 else 0.0
    total_divisa = base_imponible + iva_amount
    raw_profit = sum(item.get('profit', 0) for item in items)
    tasa_bcv = float(budget['bcv_rate'])
    total_ves = total_divisa * tasa_bcv

    conn.execute(
        """UPDATE saved_budgets SET items_json = ?, total_foreign = ?, total_ves = ?, total_profit_foreign = ?,
                    base_imponible = ?, iva_amount = ?, discount_amount = ?, discount_percent = ? WHERE id = ?""",
        (
            json.dumps(items),
            total_divisa,
            total_ves,
            max(0.0, raw_profit),
            base_imponible,
            iva_amount,
            totals['discount'],
            totals['effective_percent'],
            editing_id,
        ),
    )
    conn.commit()
    conn.close()
    flash('Producto agregado.', 'success')
    return redirect(url_for('edit_budget', id=editing_id))


@app.route('/budget/edit/add_labor', methods=['POST'])
def edit_cart_add_labor():
    editing_id = request.form.get('editing_id')
    if not editing_id:
        return redirect(url_for('index', module='📂 Presupuestos'))
    conn = get_db_connection()
    sync_budget_from_request(conn, editing_id)

    desc = request.form.get('labor_desc', '').strip()
    qty = max(1, safe_int(request.form.get('q_labor'), 1))
    price = safe_float(request.form.get('labor_price'), 0.0)
    discount = max(0.0, min(100.0, safe_float(request.form.get('labor_discount'), 0.0)))
    moneda_labor = request.form.get('labor_curr', 'USD')

    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (editing_id,)
    ).fetchone()
    if not budget:
        conn.close()
        return redirect(url_for('index', module='📂 Presupuestos'))

    moneda_base = budget['currency']
    tasa_u = get_config('tasa_usd')
    tasa_e = get_config('tasa_eur')
    precio = price
    if moneda_labor != moneda_base:
        if moneda_base == "USD" and moneda_labor == "EUR" and tasa_u > 0 and tasa_e > 0:
            precio = price * (tasa_e / tasa_u)
        elif (
            moneda_base == "EUR" and moneda_labor == "USD" and tasa_e > 0 and tasa_u > 0
        ):
            precio = price * (tasa_u / tasa_e)

    items = json.loads(budget['items_json'])
    ni = {
        "type": "LABOR",
        "code": "SERV",
        "name": desc,
        "qty": qty,
        "cost": 0.0,
        "unit_price": float(precio),
        "discount": float(discount),
        "profit": 0.0,
        "total": 0.0,
    }
    recalc_cart_item(ni)
    items.append(ni)
    totals = recalc_cart(items)
    include_iva = int(budget['include_iva'] or 0)
    base_imponible = totals['net']
    iva_amount = base_imponible * 0.16 if include_iva == 1 else 0.0
    total_divisa = base_imponible + iva_amount
    raw_profit = sum(item.get('profit', 0) for item in items)
    tasa_bcv = float(budget['bcv_rate'])
    total_ves = total_divisa * tasa_bcv

    conn.execute(
        """UPDATE saved_budgets SET items_json = ?, total_foreign = ?, total_ves = ?, total_profit_foreign = ?,
                    base_imponible = ?, iva_amount = ?, discount_amount = ?, discount_percent = ? WHERE id = ?""",
        (
            json.dumps(items),
            total_divisa,
            total_ves,
            max(0.0, raw_profit),
            base_imponible,
            iva_amount,
            totals['discount'],
            totals['effective_percent'],
            editing_id,
        ),
    )
    conn.commit()
    conn.close()
    flash('Servicio agregado.', 'success')
    return redirect(url_for('edit_budget', id=editing_id))


@app.route('/budget/edit/remove/<int:index>', methods=['POST'])
def edit_cart_remove(index):
    editing_id = request.form.get('editing_id')
    if not editing_id:
        return redirect(url_for('index', module='📂 Presupuestos'))

    conn = get_db_connection()
    sync_budget_from_request(conn, editing_id)

    budget = conn.execute(
        "SELECT * FROM saved_budgets WHERE id = ?", (editing_id,)
    ).fetchone()
    if not budget:
        conn.close()
        return redirect(url_for('index', module='📂 Presupuestos'))

    items = json.loads(budget['items_json'])
    if 0 <= index < len(items):
        items.pop(index)

    totals = recalc_cart(items)
    include_iva = int(budget['include_iva'] or 0)
    base_imponible = totals['net']
    iva_amount = base_imponible * 0.16 if include_iva == 1 else 0.0
    total_divisa = base_imponible + iva_amount
    raw_profit = sum(item.get('profit', 0) for item in items)
    tasa_bcv = float(budget['bcv_rate'])
    total_ves = total_divisa * tasa_bcv

    conn.execute(
        """UPDATE saved_budgets SET items_json = ?, total_foreign = ?, total_ves = ?, total_profit_foreign = ?,
                    base_imponible = ?, iva_amount = ?, discount_amount = ?, discount_percent = ? WHERE id = ?""",
        (
            json.dumps(items),
            total_divisa,
            total_ves,
            max(0.0, raw_profit),
            base_imponible,
            iva_amount,
            totals['discount'],
            totals['effective_percent'],
            editing_id,
        ),
    )
    conn.commit()
    conn.close()
    flash('Ítem eliminado.', 'success')
    return redirect(url_for('edit_budget', id=editing_id))


@app.route('/budget/status/<int:id>/<status>', methods=['POST'])
@block_budget_access
def change_budget_status(id, status):
    if status not in ('Borrador', 'Pendiente', 'Aprobado', 'Anulado'):
        flash('Estado inválido.', 'danger')
        return redirect(request.referrer or url_for('index'))

    conn = get_db_connection()
    # FIX BUG #9: Envolver en try/finally
    try:
        old_budget = conn.execute(
            "SELECT status FROM saved_budgets WHERE id = ?", (id,)
        ).fetchone()
        old_status = old_budget['status'] if old_budget else ''

        conn.execute("UPDATE saved_budgets SET status = ? WHERE id = ?", (status, id))

        if old_status != 'Aprobado' and status == 'Aprobado':
            update_stock_for_budget(conn, id, 'deduct')
            log_action(
                'Venta Aprobada',
                f'Presupuesto #{id} aprobado. Stock descontado.',
                'Presupuesto',
                id,
                conn=conn,
            )
        elif old_status == 'Aprobado' and status != 'Aprobado':
            update_stock_for_budget(conn, id, 'restore')
            log_action(
                'Venta Anulada/Modificada',
                f'Presupuesto #{id} ya no está aprobado. Stock restaurado.',
                'Presupuesto',
                id,
                conn=conn,
            )

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[change_budget_status] Error: {e}")
        flash('Error al cambiar el estado del presupuesto.', 'danger')
    finally:
        conn.close()

    flash(f'Presupuesto: {status}.', 'success')
    return redirect(request.referrer or url_for('index', module='📂 Presupuestos'))


@app.route('/budget/toggle_almacen/<int:id>', methods=['POST'])
@block_budget_access
def toggle_almacen_status(id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    # 🔒 BLOQUEO DE SEGURIDAD: Solo roles con permiso pueden cambiar el estatus
    if not check_action_permission(session.get('role', ''), 'almacen_toggle'):
        return (
            jsonify(
                {
                    'ok': False,
                    'message': 'No tienes permiso para modificar el estatus del almacén.',
                }
            ),
            403,
        )

    conn = get_db_connection()
    row = conn.execute(
        "SELECT almacen_status FROM saved_budgets WHERE id = ?", (id,)
    ).fetchone()
    if row:
        current = row['almacen_status'] or 'Por procesar'
        new_status = 'Procesado' if current == 'Por procesar' else 'Por procesar'

        conn.execute(
            "UPDATE saved_budgets SET almacen_status = ? WHERE id = ?", (new_status, id)
        )
        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'status': new_status})

    conn.close()
    return jsonify({'ok': False, 'message': 'No encontrado'}), 404


@app.route('/budget/toggle_payment/<int:id>', methods=['POST'])
def toggle_budget_payment(id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    conn = get_db_connection()
    budget = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    if not budget:
        conn.close()
        return jsonify({'ok': False, 'message': 'No encontrado'}), 404

    current_paid = float(budget['total_paid_foreign'] or 0)
    total = float(budget['total_foreign'] or 0)
    rate = float(budget['bcv_rate'] or 1)

    if current_paid >= total - 0.01:
        # Estaba Pagado -> Pasa a "Sin Pago"
        new_paid = 0.0
        new_paid_ves = 0.0
        # Limpiamos los pagos registrados para mantener consistencia contable
        conn.execute("DELETE FROM payments WHERE budget_id = ?", (id,))
    else:
        # Estaba Sin Pago o Parcial -> Pasa a "Pagado"
        new_paid = total
        new_paid_ves = total * rate

    conn.execute(
        "UPDATE saved_budgets SET total_paid_foreign = ?, total_paid_ves = ? WHERE id = ?",
        (new_paid, new_paid_ves, id),
    )

    # Actualizar estatus si es necesario
    if new_paid > 0 and budget['status'] == 'Pendiente':
        conn.execute("UPDATE saved_budgets SET status = 'Aprobado' WHERE id = ?", (id,))
        # 👇 NUEVO: Generar comanda y descontar stock
        update_stock_for_budget(conn, id, 'deduct')
        # 👆 FIN NUEVO

    conn.commit()

    # Calcular el nuevo estado de pago
    new_status = compute_payment_status(
        {
            'total_foreign': total,
            'total_paid_foreign': new_paid,
            'status': budget['status'],
        }
    )
    conn.close()

    return jsonify({'ok': True, 'status': new_status})


@app.route('/budget/<int:id>/generate_public_token', methods=['POST'])
def generate_budget_public_token(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    budget = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    if budget:
        token = budget['public_token'] or generate_public_token()
        conn.execute(
            "UPDATE saved_budgets SET public_token = ? WHERE id = ?", (token, id)
        )
        conn.commit()
        flash('Link público generado.', 'success')
    conn.close()
    return redirect(request.referrer or url_for('index', module='📂 Presupuestos'))


@app.route('/budget/duplicate/<int:id>', methods=['POST'])
@block_budget_access
def duplicate_budget(id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    orig = conn.execute("SELECT * FROM saved_budgets WHERE id = ?", (id,)).fetchone()
    if not orig:
        conn.close()
        flash('Presupuesto no encontrado.', 'danger')
        return redirect(url_for('index', module='📂 Presupuestos'))

    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
    new_token = generate_public_token()
    current_user = session.get('username', '')

    def _g(key, default=''):
        return orig[key] if key in orig.keys() and orig[key] is not None else default

    year_str = get_ve_time().strftime("%Y")
    last_budget = conn.execute(
        "SELECT budget_number FROM saved_budgets WHERE budget_number LIKE ? ORDER BY id DESC LIMIT 1",
        (f"COT-{year_str}-%",),
    ).fetchone()
    next_seq = 1
    if last_budget and last_budget['budget_number']:
        try:
            next_seq = int(last_budget['budget_number'].split('-')[-1]) + 1
        except ValueError:
            pass
    new_budget_num = f"COT-{year_str}-{next_seq:04d}"

    # FIX BUG #5: Incluir budget_number
    cursor = conn.execute(
        """
        INSERT INTO saved_budgets (
            date, client_name, client_doc, client_phone, currency, bcv_rate,
            total_foreign, total_ves, total_profit_foreign, items_json, status,
            vehicle_name, vehicle_year, vehicle_plate,
            vehicle_brand, vehicle_color, vehicle_model,
            username, budget_note,
            base_imponible, include_iva, iva_amount, discount_amount, discount_percent,
            total_paid_foreign, total_paid_ves, public_token,
            client_signature, client_signature_date, show_bs_equivalent, budget_number
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente',
                  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, 0.0, ?, '', '', ?, ?)
    """,
        (
            now_str,
            _g('client_name'),
            _g('client_doc'),
            _g('client_phone'),
            _g('currency', 'USD'),
            float(_g('bcv_rate', 0.0)),
            float(_g('total_foreign', 0.0)),
            float(_g('total_ves', 0.0)),
            float(_g('total_profit_foreign', 0.0)),
            _g('items_json', '[]'),
            _g('vehicle_name'),
            _g('vehicle_year'),
            _g('vehicle_plate'),
            _g('vehicle_brand'),
            _g('vehicle_color'),
            _g('vehicle_model'),
            current_user,
            _g('budget_note'),
            float(_g('base_imponible', 0.0)),
            int(_g('include_iva', 0) or 0),
            float(_g('iva_amount', 0.0)),
            float(_g('discount_amount', 0.0)),
            float(_g('discount_percent', 0.0)),
            new_token,
            int(_g('show_bs_equivalent', 0) or 0),
            new_budget_num,
        ),
    )
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()

    log_action(
        'Duplicado de presupuesto',
        f'Se duplicó el presupuesto #{id} → #{new_id}.',
        'Presupuesto',
        new_id,
    )
    flash(
        f'✓ Presupuesto #{id} duplicado como #{new_id}. Ahora puedes quitar los ítems que correspondan.',
        'success',
    )
    return redirect(url_for('edit_budget', id=new_id))


@app.route('/budget/delete/<int:id>', methods=['POST'])
@block_budget_access
def delete_budget(id):
    conn = get_db_connection()
    try:
        b = conn.execute(
            "SELECT username, status, items_json FROM saved_budgets WHERE id = ?", (id,)
        ).fetchone()
        if b and (
            session.get('role') == 'Administrador'
            or b['username'] == session.get('username')
        ):

            # 1. Restaurar el stock de TODOS los productos (incluso si la comanda ya se procesó)
            if b['status'] == 'Aprobado':
                try:
                    items = json.loads(b['items_json'] or '[]')
                    for it in items:
                        if it.get('type') == 'PRODUCTO' and it.get('qty', 0) > 0:
                            code = it.get('code', '')
                            if code and code != 'N/A':
                                conn.execute(
                                    """
                                    UPDATE inventory
                                    SET stock = stock + ?,
                                        stock_1 = stock_1 + ?,
                                        stock_2 = stock_2 + ?
                                    WHERE UPPER(code) = UPPER(?)
                                """,
                                    (it['qty'], it['qty'], it['qty'], code),
                                )
                except Exception as e:
                    print(f"Error restaurando stock al borrar: {e}")

            # 2. Eliminar los pagos asociados
            conn.execute("DELETE FROM payments WHERE budget_id = ?", (id,))

            # 3. Eliminar las comandas asociadas
            comandas = conn.execute(
                "SELECT id FROM comandas WHERE budget_id = ?", (id,)
            ).fetchall()
            for c in comandas:
                conn.execute(
                    "DELETE FROM comanda_items WHERE comanda_id = ?", (c['id'],)
                )
            conn.execute("DELETE FROM comandas WHERE budget_id = ?", (id,))

            # 4. Eliminar el presupuesto
            conn.execute("DELETE FROM saved_budgets WHERE id = ?", (id,))

            conn.commit()
            flash('Presupuesto eliminado y stock restaurado exitosamente.', 'success')
        else:
            flash('No tienes permiso para eliminar este presupuesto.', 'danger')

    except Exception as e:
        conn.rollback()
        flash(f'Error al eliminar el presupuesto: {e}', 'danger')
    finally:
        conn.close()

    return redirect(request.referrer or url_for('index', module='📂 Presupuestos'))


@app.route('/user/add', methods=['POST'])
def add_user():
    u = request.form.get('new_username', '').strip()
    p = request.form.get('new_password', '').strip()
    fn = request.form.get('new_first_name', '').strip()
    ln = request.form.get('new_last_name', '').strip()
    ph = request.form.get('new_phone', '').strip()
    r = request.form.get('new_role', '').strip()
    sp = 1 if request.form.get('show_in_pdf') == '1' else 0

    if u and p and fn and ln and ph:
        try:
            conn = get_db_connection()
            conn.execute(
                """INSERT INTO users (username, password, role, first_name, last_name, phone, show_in_pdf)
                            VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (u, generate_password_hash(p), r, fn, ln, ph, sp),
            )
            conn.commit()
            conn.close()
            flash('Usuario creado.', 'success')
        except sqlite3.IntegrityError:
            flash('Usuario ya existe.', 'danger')

    invalidate_users_cache()  # <-- LÍNEA NUEVA
    return redirect(url_for('index', module='👥 Usuarios'))


@app.route('/user/edit/<int:id>', methods=['POST'])
def edit_user(id):
    if not session.get('logged_in') or session.get('role') != 'Administrador':
        flash('Solo admins.', 'danger')
        return redirect(url_for('index', module='👥 Usuarios'))

    u = request.form.get('edit_username', '').strip()
    p = request.form.get('edit_password', '').strip()
    fn = request.form.get('edit_first_name', '').strip()
    ln = request.form.get('edit_last_name', '').strip()
    ph = request.form.get('edit_phone', '').strip()
    r = request.form.get('edit_role', '').strip()
    sp = 1 if request.form.get('edit_show_in_pdf') == '1' else 0

    conn = get_db_connection()
    t = conn.execute("SELECT * FROM users WHERE id = ?", (id,)).fetchone()

    if t:
        if t['username'].lower() == 'admin':
            r = 'Administrador'
        try:
            if p:
                conn.execute(
                    """UPDATE users SET username = ?, password = ?, role = ?, first_name = ?, last_name = ?, phone = ?, show_in_pdf = ? WHERE id = ?""",
                    (u, generate_password_hash(p), r, fn, ln, ph, sp, id),
                )
            else:
                conn.execute(
                    """UPDATE users SET username = ?, role = ?, first_name = ?, last_name = ?, phone = ?, show_in_pdf = ? WHERE id = ?""",
                    (u, r, fn, ln, ph, sp, id),
                )
            conn.commit()
            flash(f'Usuario {u} actualizado.', 'success')
        except sqlite3.IntegrityError:
            flash('Usuario en uso.', 'danger')

    conn.close()

    invalidate_users_cache()  # <-- LÍNEA NUEVA
    return redirect(url_for('index', module='👥 Usuarios'))


@app.route('/user/delete/<int:id>', methods=['POST'])
def delete_user(id):
    if not session.get('logged_in') or session.get('role') != 'Administrador':
        flash('Solo admins.', 'danger')
        return redirect(url_for('index', module='👥 Usuarios'))

    conn = get_db_connection()
    user = conn.execute("SELECT username FROM users WHERE id = ?", (id,)).fetchone()

    if user and user['username'].lower() == 'admin':
        flash('El admin principal no se puede eliminar.', 'danger')
    elif user:
        conn.execute("DELETE FROM users WHERE id = ?", (id,))
        conn.commit()
        flash('Usuario eliminado.', 'success')

    conn.close()

    invalidate_users_cache()  # <-- LÍNEA NUEVA
    return redirect(url_for('index', module='👥 Usuarios'))


@app.route('/role/add', methods=['POST'])
def add_role():
    name = request.form.get('role_name', '').strip()
    if name:
        try:
            conn = get_db_connection()
            conn.execute("INSERT INTO roles (name) VALUES (?)", (name,))
            conn.commit()
            conn.close()
            flash('Rol creado.', 'success')
        except sqlite3.IntegrityError:
            flash('Rol existe.', 'danger')
    return redirect(url_for('index', module='🛡️ Roles y Permisos'))


@app.route('/role/permissions', methods=['POST'])
def update_role_permissions():
    role = request.form.get('role_name')
    modules = request.form.getlist('modules')
    conn = get_db_connection()
    # Guardar los permisos action:* existentes
    action_rows = conn.execute(
        "SELECT module_name FROM role_permissions WHERE role_name = ? AND module_name LIKE 'action:%'",
        (role,),
    ).fetchall()
    action_perms = [r['module_name'] for r in action_rows]

    conn.execute("DELETE FROM role_permissions WHERE role_name = ?", (role,))
    for m in modules:
        conn.execute(
            "INSERT INTO role_permissions (role_name, module_name) VALUES (?, ?)",
            (role, m),
        )
    for a in action_perms:  # reinsertar
        conn.execute(
            "INSERT OR IGNORE INTO role_permissions (role_name, module_name) VALUES (?, ?)",
            (role, a),
        )
    conn.commit()
    conn.close()
    flash('Permisos actualizados.', 'success')
    return redirect(url_for('index', module='🛡️ Roles y Permisos'))


@app.route('/role/edit/<old_name>', methods=['POST'])
def edit_role_name(old_name):
    if not session.get('logged_in') or session.get('role') != 'Administrador':
        flash('Solo administradores pueden editar roles.', 'danger')
        return redirect(url_for('index', module='🛡️ Roles y Permisos'))

    if old_name.lower() == 'administrador':
        flash('El rol "Administrador" no se puede renombrar.', 'danger')
        return redirect(url_for('index', module='🛡️ Roles y Permisos'))

    new_name = request.form.get('new_role_name', '').strip()
    if new_name and new_name != old_name:
        conn = get_db_connection()
        try:
            conn.execute(
                "UPDATE roles SET name = ? WHERE name = ?", (new_name, old_name)
            )
            conn.execute(
                "UPDATE role_permissions SET role_name = ? WHERE role_name = ?",
                (new_name, old_name),
            )
            conn.execute(
                "UPDATE users SET role = ? WHERE role = ?", (new_name, old_name)
            )
            conn.commit()
        except sqlite3.IntegrityError:
            conn.rollback()
            flash('El nombre del rol ya existe.', 'danger')
            return redirect(url_for('index', module='🛡️ Roles y Permisos'))
        finally:
            conn.close()

        flash(f'✓ Rol renombrado a "{new_name}".', 'success')
        log_action(
            'Edición de Rol', f'Rol {old_name} renombrado a {new_name}.', 'Roles'
        )  # ← SIN conn, fuera del try

    return redirect(url_for('index', module='🛡️ Roles y Permisos'))


@app.route('/role/delete/<role>', methods=['POST'])
def delete_role(role):
    if role.strip().lower() != 'administrador':
        conn = get_db_connection()
        conn.execute("DELETE FROM roles WHERE name = ?", (role,))
        conn.execute("DELETE FROM role_permissions WHERE role_name = ?", (role,))
        conn.commit()
        conn.close()
        flash('Rol eliminado.', 'success')
    return redirect(url_for('index', module='🛡️ Roles y Permisos'))


@app.route('/admin/backup')
def admin_backup():
    if not session.get('logged_in') or session.get('role') != 'Administrador':
        return redirect(url_for('login'))
    return send_file(
        DB_NAME,
        as_attachment=True,
        download_name=f"backup_{get_ve_time().strftime('%Y%m%d_%H%M')}.db",
    )


def _wb_get_column_meta(column_id):
    """
    Dado un column_id:
    - 'sin_asignar'         → ('', 'taller')     (bandeja de entrada)
    - 'taller_juan'         → ('juan', 'taller')
    - 'prueba_juan'         → ('juan', 'prueba')
    - 'entregado_juan'      → ('juan', 'entregado')
    """
    if not column_id:
        return (None, None)
    if column_id == 'sin_asignar':
        return ('', 'taller')
    if '_' not in column_id:
        return (None, None)
    parts = column_id.split('_', 1)
    if len(parts) != 2:
        return (None, None)
    kind, username = parts
    if kind in ('taller', 'prueba', 'entregado'):
        return (username, kind)
    if kind == 'seller':
        return (username, 'taller')
    if kind == 'delivered':
        return (username, 'entregado')
    return (None, None)


def _wb_log_activity(card_id, content, username=None, conn=None):
    own_conn = conn is None
    try:
        actor = username or session.get('username', 'Sistema')
        if own_conn:
            conn = get_db_connection()
        conn.execute(
            """
            INSERT INTO workshop_card_comments (card_id, username, type, content, created_at)
            VALUES (?, ?, 'activity', ?, ?)
        """,
            (
                card_id,
                actor,
                str(content)[:500],
                get_ve_time().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        if own_conn:
            conn.commit()
    except Exception as e:
        print(f"[_wb_log_activity] {e}")
    finally:
        if own_conn and conn:
            conn.close()


def _wb_log_comment(card_id, content, username=None, conn=None):
    own_conn = conn is None
    try:
        actor = username or session.get('username', 'Sistema')
        if own_conn:
            conn = get_db_connection()
        conn.execute(
            """
            INSERT INTO workshop_card_comments (card_id, username, type, content, created_at)
            VALUES (?, ?, 'comment', ?, ?)
        """,
            (
                card_id,
                actor,
                str(content)[:2000],
                get_ve_time().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        if own_conn:
            conn.commit()
    except Exception as e:
        print(f"[_wb_log_comment] {e}")
    finally:
        if own_conn and conn:
            conn.close()


def _wb_log_comment(card_id, content, username=None):
    """Registra un comentario del usuario."""
    try:
        actor = username or session.get('username', 'Sistema')
        conn = get_db_connection()
        conn.execute(
            """
            INSERT INTO workshop_card_comments (card_id, username, type, content, created_at)
            VALUES (?, ?, 'comment', ?, ?)
        """,
            (
                card_id,
                actor,
                str(content)[:2000],
                get_ve_time().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[_wb_log_comment] {e}")


@app.route('/workshop/column/<username>/<int:is_delivered>/archive', methods=['POST'])
def workshop_column_archive(username, is_delivered):
    """Archiva todas las tarjetas entregadas de un vendedor (o restaura)."""
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if is_delivered != 1:
        flash('Solo se pueden archivar columnas de entregados.', 'warning')
        return redirect(url_for('index', module='🔧 Tablero Taller'))

    conn = get_db_connection()
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    cursor = conn.execute(
        """
        UPDATE workshop_cards SET archived = 1, updated_at = ?
        WHERE assigned_to = ? AND is_delivered = 1 AND archived = 0
    """,
        (now_str, username),
    )
    count = cursor.rowcount
    conn.commit()
    conn.close()

    log_action(
        'Archivado masivo de tarjetas',
        f'Se archivaron {count} tarjetas entregadas de {username}.',
        'Taller',
    )
    flash(f'✓ {count} tarjeta(s) archivada(s).', 'success')
    return redirect(url_for('index', module='🔧 Tablero Taller'))


@app.route('/workshop/board/checksum')
def workshop_board_checksum():
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    try:
        row = conn.execute(
            "SELECT COUNT(*) AS c, COALESCE(MAX(updated_at),'') AS m, COALESCE(SUM(id),0) AS s "
            "FROM workshop_cards WHERE archived = 0"
        ).fetchone()
        row2 = conn.execute(
            "SELECT COUNT(*) AS c, COALESCE(MAX(created_at),'') AS m "
            "FROM workshop_card_comments"
        ).fetchone()
        checksum = f"{row['c'] or 0}:{row['m'] or ''}:{row['s'] or 0}:{row2['c'] or 0}:{row2['m'] or ''}"
        return jsonify({'ok': True, 'checksum': checksum})
    finally:
        conn.close()


@app.route('/workshop/card/add', methods=['POST'])
@block_mechanic
def workshop_card_add():

    if not session.get('logged_in'):
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'ok': False, 'message': 'No autenticado'}), 401
        return redirect(url_for('login'))

    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    # Si no se especifica, por defecto va para Manuel
    _sellers = get_workshop_sellers()  # <-- Añadir
    default_user = (
        _sellers[0] if _sellers else session.get('username', '')
    )  # <-- Usar _sellers
    assigned_to = request.form.get('w_assigned_to', '').strip() or default_user

    client_name = request.form.get('w_client_name', '').strip()
    client_doc = up(request.form.get('w_client_doc', ''))
    client_phone = request.form.get('w_client_phone', '').strip()
    veh_plate = up(request.form.get('w_vehicle_plate', ''))
    veh_name = request.form.get('w_vehicle_name', '').strip()
    veh_brand = up(request.form.get('w_vehicle_brand', ''))
    veh_model = up(request.form.get('w_vehicle_model', ''))
    veh_year = request.form.get('w_vehicle_year', '').strip()
    priority = request.form.get('w_priority', 'Media')
    if priority not in REMINDER_PRIORITIES:
        priority = 'Media'
    entry_date = request.form.get('w_entry_date', '').strip() or get_ve_time().strftime(
        "%Y-%m-%d"
    )
    due_date = request.form.get('w_due_date', '').strip()
    notes = request.form.get('w_notes', '').strip()
    description = request.form.get('w_description', '').strip()

    conn = get_db_connection()
    try:
        max_sort = conn.execute(
            "SELECT COALESCE(MAX(sort_order), 0) FROM workshop_cards WHERE assigned_to = ? AND is_delivered = 0",
            (assigned_to,),
        ).fetchone()[0]
        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")

        default_checklist_json = json.dumps(WORKSHOP_DEFAULT_CHECKLIST)

        cursor = conn.execute(
            """
            INSERT INTO workshop_cards (
                column_id, sort_order, client_doc, client_name, client_phone,
                vehicle_plate, vehicle_name, vehicle_brand, vehicle_model, vehicle_year,
                assigned_to, priority, entry_date, due_date, notes, description,
                created_at, updated_at, created_by, is_delivered, checklist_json, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, 'taller')
        """,
            (
                0,
                max_sort + 1,
                client_doc,
                client_name or 'SIN NOMBRE',
                client_phone,
                veh_plate,
                veh_name,
                veh_brand,
                veh_model,
                veh_year,
                assigned_to,
                priority,
                entry_date,
                due_date,
                notes,
                description,
                now_str,
                now_str,
                session.get('username', ''),
                default_checklist_json,
            ),
        )
        card_id = cursor.lastrowid
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        if is_ajax:
            return jsonify({'ok': False, 'message': str(e)}), 500
        flash(f'Error al crear tarjeta: {e}', 'danger')
        return redirect(url_for('index', module='🔧 Tablero Taller'))
    finally:
        conn.close()

    log_action(
        'Creación de tarjeta taller',
        f'Tarjeta #{card_id} — {veh_plate or client_name or "sin datos"}.',
        'Taller',
        card_id,
    )
    _wb_log_activity(card_id, 'Creó la tarjeta')

    if is_ajax:
        return jsonify(
            {
                'ok': True,
                'card': {
                    'id': card_id,
                    'column_id': f"taller_{assigned_to}",  # ← ANTES: f"seller_{assigned_to}"
                    'client_name': client_name or 'SIN NOMBRE',
                    'client_doc': client_doc,
                    'client_phone': client_phone,
                    'vehicle_plate': veh_plate,
                    'vehicle_brand': veh_brand,
                    'vehicle_model': veh_model,
                    'vehicle_year': veh_year,
                    'vehicle_name': veh_name,
                    'assigned_to': assigned_to,
                    'priority': priority,
                    'entry_date': entry_date,
                    'due_date': due_date,
                    'notes': notes,
                    'description': description,
                    'labels': [],
                    'checklist': WORKSHOP_DEFAULT_CHECKLIST,
                    'is_delivered': 0,
                    'status': 'taller',
                },
            }
        )

    flash('✓ Tarjeta agregada al tablero.', 'success')
    return redirect(url_for('index', module='🔧 Tablero Taller'))


@app.route('/workshop/card/edit/<int:card_id>', methods=['POST'])
@block_mechanic
def workshop_card_edit(card_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    card = conn.execute(
        "SELECT * FROM workshop_cards WHERE id = ?", (card_id,)
    ).fetchone()
    if not card:
        conn.close()
        flash('Tarjeta no encontrada.', 'danger')
        return redirect(url_for('index', module='🔧 Tablero Taller'))

    priority = request.form.get('w_priority', 'Media')
    if priority not in REMINDER_PRIORITIES:
        priority = 'Media'
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")

    conn.execute(
        """
        UPDATE workshop_cards SET
            client_doc = ?, client_name = ?, client_phone = ?,
            vehicle_plate = ?, vehicle_name = ?, vehicle_brand = ?, vehicle_model = ?, vehicle_year = ?,
            assigned_to = ?, priority = ?, entry_date = ?, due_date = ?, notes = ?, description = ?,
            updated_at = ?
        WHERE id = ?
    """,
        (
            up(request.form.get('w_client_doc', '')),
            request.form.get('w_client_name', '').strip() or 'SIN NOMBRE',
            request.form.get('w_client_phone', '').strip(),
            up(request.form.get('w_vehicle_plate', '')),
            request.form.get('w_vehicle_name', '').strip(),
            up(request.form.get('w_vehicle_brand', '')),
            up(request.form.get('w_vehicle_model', '')),
            request.form.get('w_vehicle_year', '').strip(),
            request.form.get('w_assigned_to', '').strip() or card['assigned_to'],
            priority,
            request.form.get('w_entry_date', '').strip(),
            request.form.get('w_due_date', '').strip(),
            request.form.get('w_notes', '').strip(),
            request.form.get('w_description', '').strip(),
            now_str,
            card_id,
        ),
    )
    conn.commit()
    conn.close()

    log_action(
        'Edición de tarjeta taller', f'Tarjeta #{card_id} editada.', 'Taller', card_id
    )
    _wb_log_activity(card_id, 'Editó los datos de la tarjeta')
    flash('✓ Tarjeta actualizada.', 'success')
    return redirect(url_for('index', module='🔧 Tablero Taller'))


@app.route('/workshop/card/delete/<int:card_id>', methods=['POST'])
@block_mechanic
def workshop_card_delete(card_id):
    if not session.get('logged_in'):
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'ok': False, 'message': 'No autenticado'}), 401
        return redirect(url_for('login'))

    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    conn = get_db_connection()
    conn.execute("DELETE FROM workshop_card_comments WHERE card_id = ?", (card_id,))
    conn.execute("DELETE FROM workshop_cards WHERE id = ?", (card_id,))
    conn.commit()
    conn.close()

    log_action(
        'Eliminación de tarjeta taller',
        f'Tarjeta #{card_id} eliminada.',
        'Taller',
        card_id,
    )

    if is_ajax:
        return jsonify({'ok': True, 'card_id': card_id})

    flash('Tarjeta eliminada.', 'success')
    return redirect(url_for('index', module='🔧 Tablero Taller'))


@app.route('/workshop/card/move/<int:card_id>', methods=['POST'])
@block_mechanic
def workshop_card_move(card_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    new_column = request.form.get('column_id', '').strip()
    order_str = request.form.get('order', '').strip()

    username, status = _wb_get_column_meta(new_column)
    if username is None or status not in ('taller', 'prueba', 'entregado'):
        return jsonify({'ok': False, 'message': 'Columna inválida'}), 400

    conn = get_db_connection()
    card = conn.execute(
        "SELECT * FROM workshop_cards WHERE id = ?", (card_id,)
    ).fetchone()
    if not card:
        conn.close()
        return jsonify({'ok': False, 'message': 'Tarjeta no existe'}), 400
    if username:
        u = conn.execute(
            "SELECT username FROM users WHERE username = ?", (username,)
        ).fetchone()
        if not u:
            conn.close()
            return jsonify({'ok': False, 'message': 'Usuario no existe'}), 400

    is_delivered = 1 if status == 'entregado' else 0
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")

    # Determinar delivered_at: setear al entrar a entregado, limpiar al salir
    prev_status = (
        card['status'] if 'status' in card.keys() and card['status'] else 'taller'
    )
    if status == 'entregado' and prev_status != 'entregado':
        delivered_at = now_str
    elif status != 'entregado':
        delivered_at = ''
    else:
        delivered_at = (
            card['delivered_at']
            if 'delivered_at' in card.keys() and card['delivered_at']
            else now_str
        )

    try:
        conn.execute(
            "UPDATE workshop_cards SET assigned_to = ?, status = ?, is_delivered = ?, "
            "delivered_at = ?, updated_at = ? WHERE id = ?",
            (username, status, is_delivered, delivered_at, now_str, card_id),
        )
        if order_str:
            ids = [int(x) for x in order_str.split(',') if x.strip().isdigit()]
            for idx, cid in enumerate(ids):
                conn.execute(
                    "UPDATE workshop_cards SET sort_order = ? WHERE id = ? AND assigned_to = ? AND status = ?",
                    (idx, cid, username, status),
                )
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'ok': False, 'message': str(e)}), 500
    conn.close()

    status_labels = {
        'taller': 'En Taller',
        'prueba': 'Prueba de Campo',
        'entregado': 'Entregado',
    }
    log_action(
        'Movimiento de tarjeta taller',
        f'Tarjeta #{card_id} movida a "{status_labels[status]}" de {username}.',
        'Taller',
        card_id,
    )
    _wb_log_activity(card_id, f'Cambió estado a "{status_labels[status]}"')

    return jsonify(
        {
            'ok': True,
            'column': new_column,
            'username': username,
            'status': status,
            'is_delivered': is_delivered,
        }
    )


@app.route('/workshop/column/reorder', methods=['POST'])
@block_mechanic
def workshop_column_reorder():
    """Reordena todas las tarjetas de una columna según una lista de IDs."""
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    column_id = request.form.get('column_id', '').strip()
    order_str = request.form.get('order', '').strip()

    username, status = _wb_get_column_meta(column_id)
    if not username or status not in ('taller', 'prueba', 'entregado'):
        return jsonify({'ok': False, 'message': 'Columna inválida'}), 400
    if not order_str:
        return jsonify({'ok': False, 'message': 'Orden vacío'}), 400

    conn = get_db_connection()
    try:
        ids = [int(x) for x in order_str.split(',') if x.strip().isdigit()]
        for idx, cid in enumerate(ids):
            conn.execute(
                "UPDATE workshop_cards SET sort_order = ? WHERE id = ? AND assigned_to = ? AND status = ?",
                (idx, cid, username, status),
            )
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'ok': False, 'message': str(e)}), 500
    conn.close()

    return jsonify({'ok': True, 'count': len(ids)})


@app.route('/workshop/card/<int:card_id>/detail_data')
def workshop_card_detail_data(card_id):
    """Devuelve JSON con el detalle completo de la tarjeta."""
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    card = conn.execute(
        "SELECT * FROM workshop_cards WHERE id = ?", (card_id,)
    ).fetchone()

    # 👇 NUEVO: Verificar si es garantía 👇
    is_warranty = False
    if card and card['vehicle_plate']:
        garantia_row = conn.execute(
            """
            SELECT id FROM revisions
            WHERE UPPER(TRIM(vehicle_plate)) = UPPER(TRIM(?))
              AND reception_es_retorno = 1
            ORDER BY id DESC LIMIT 1
        """,
            (card['vehicle_plate'],),
        ).fetchone()
        if garantia_row:
            is_warranty = True
    # 👆 FIN NUEVO 👆

    conn.close()
    if not card:
        return jsonify({'ok': False, 'message': 'No encontrada'}), 404
    try:
        labels = json.loads(card['labels_json'] or '[]')
    except Exception:
        labels = []
    try:
        checklist = json.loads(card['checklist_json'] or '[]')
    except Exception:
        checklist = []
    return jsonify(
        {
            'ok': True,
            'id': card['id'],
            'title': card['title'] if 'title' in card.keys() else '',
            'client_name': card['client_name'] or '',
            'client_doc': card['client_doc'] or '',
            'client_phone': card['client_phone'] or '',
            'vehicle_plate': card['vehicle_plate'] or '',
            'vehicle_brand': card['vehicle_brand'] or '',
            'vehicle_model': card['vehicle_model'] or '',
            'vehicle_year': card['vehicle_year'] or '',
            'vehicle_name': card['vehicle_name'] or '',
            'assigned_to': card['assigned_to'] or '',
            'priority': card['priority'] or 'Media',
            'entry_date': card['entry_date'] or '',
            'due_date': card['due_date'] or '',
            'notes': card['notes'] or '',
            'description': card['description'] or '',
            'is_delivered': int(card['is_delivered'] or 0),
            'status': (
                card['status']
                if 'status' in card.keys() and card['status']
                else 'taller'
            ),
            'labels': labels,
            'checklist': checklist,
            'repair_status': (
                card['repair_status']
                if 'repair_status' in card.keys()
                else 'Espera de Servicio Técnico'
            ),
            'is_warranty': is_warranty,
        }
    )


@app.route('/workshop/card/<int:card_id>/save', methods=['POST'])
def workshop_card_save(card_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    data = request.get_json(silent=True) or {}

    # 🛑 BLOQUEO PARA MECÁNICOS: Solo pueden guardar el checklist
    role = session.get('role', '').lower()
    if 'mecanico' in role or 'mecánico' in role:
        data = {'checklist': data.get('checklist', [])}

    conn = get_db_connection()
    card = conn.execute(
        "SELECT * FROM workshop_cards WHERE id = ?", (card_id,)
    ).fetchone()
    # ... (deja el resto del código de esta función exactamente igual) ...
    if not card:
        conn.close()
        return jsonify({'ok': False, 'message': 'Tarjeta no encontrada'}), 404

    allowed_text = [
        'title',
        'description',
        'notes',
        'entry_date',
        'due_date',
        'priority',
        'assigned_to',
        'client_name',
        'client_doc',
        'client_phone',
        'vehicle_plate',
        'vehicle_brand',
        'vehicle_model',
        'vehicle_year',
        'vehicle_name',
        'status',
        'repair_status',
    ]

    updates = []
    params = []
    for field in allowed_text:
        if field in data:
            value = data[field]
            if isinstance(value, str):
                value = value.strip()
            if field == 'priority' and value not in REMINDER_PRIORITIES:
                value = 'Media'
            if field == 'status':
                if value not in ('taller', 'prueba', 'entregado'):
                    value = 'taller'
                updates.append("status = ?")
                params.append(value)
                # Mantener is_delivered en sync
                is_deliv_val = 1 if value == 'entregado' else 0
                updates.append("is_delivered = ?")
                params.append(is_deliv_val)
            else:
                updates.append(f"{field} = ?")
                params.append(value)

    # JSON fields
    if 'labels' in data:
        try:
            labels_json = json.dumps(data['labels'])
        except Exception:
            labels_json = '[]'
        updates.append("labels_json = ?")
        params.append(labels_json)

    if 'checklist' in data:
        try:
            new_checklist = data['checklist']
            if isinstance(new_checklist, list):
                # PROTECCIÓN CONTRA BORRADO ACCIDENTAL:
                # Si el checklist llega VACÍO pero el JSON trae otros campos (ej. priority, description),
                # significa que el frontend se desincronizó y no cargó el checklist a tiempo.
                # En ese caso, ignoramos la orden de borrar y conservamos el que está en la BD.
                if len(new_checklist) == 0 and len(data) > 1:
                    pass
                else:
                    # Si viene vacío pero es el único campo enviado, es porque el usuario
                    # eliminó el último ítem a propósito. Lo guardamos.
                    checklist_json = json.dumps(new_checklist)
                    updates.append("checklist_json = ?")
                    params.append(checklist_json)
        except Exception as e:
            print(f"[workshop_card_save] Error procesando checklist: {e}")

    if updates:
        updates.append("updated_at = ?")
        params.append(get_ve_time().strftime("%Y-%m-%d %H:%M"))
        params.append(card_id)
        try:
            conn.execute(
                f"UPDATE workshop_cards SET {', '.join(updates)} WHERE id = ?", params
            )

            if 'priority' in data:
                _wb_log_activity(
                    card_id, f'Cambió prioridad a "{data["priority"]}"', conn=conn
                )
            if 'assigned_to' in data:
                _wb_log_activity(
                    card_id, f'Asignó la tarjeta a {data["assigned_to"]}', conn=conn
                )
            if 'due_date' in data:
                _wb_log_activity(
                    card_id,
                    f'Cambió fecha de entrega a {data["due_date"] or "(vacía)"}',
                    conn=conn,
                )
            if 'entry_date' in data:
                _wb_log_activity(
                    card_id,
                    f'Cambió fecha de ingreso a {data["entry_date"] or "(vacía)"}',
                    conn=conn,
                )
            if 'description' in data:
                _wb_log_activity(card_id, 'Editó la descripción', conn=conn)
            if 'title' in data:
                _wb_log_activity(card_id, 'Cambió el título', conn=conn)
            if 'labels' in data:
                _wb_log_activity(card_id, 'Actualizó las etiquetas', conn=conn)
            if 'checklist' in data:
                _wb_log_activity(card_id, 'Actualizó el checklist', conn=conn)

            conn.commit()
        except Exception as e:
            conn.rollback()
            conn.close()
            return jsonify({'ok': False, 'message': str(e)}), 500

    conn.close()
    return jsonify({'ok': True})


@app.route('/workshop/card/<int:card_id>/comments', methods=['GET'])
def workshop_card_comments(card_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT * FROM workshop_card_comments
        WHERE card_id = ?
        ORDER BY id DESC
        LIMIT 200
    """,
        (card_id,),
    ).fetchall()
    conn.close()
    return jsonify(
        {
            'ok': True,
            'items': [
                {
                    'id': r['id'],
                    'username': r['username'],
                    'type': r['type'],
                    'content': r['content'],
                    'created_at': r['created_at'],
                }
                for r in rows
            ],
        }
    )


@app.route('/workshop/card/<int:card_id>/comment', methods=['POST'])
def workshop_card_add_comment(card_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    data = request.get_json(silent=True) or {}
    content = (data.get('content') or '').strip()
    if not content:
        return jsonify({'ok': False, 'message': 'Comentario vacío'}), 400

    conn = get_db_connection()
    card = conn.execute(
        "SELECT id FROM workshop_cards WHERE id = ?", (card_id,)
    ).fetchone()
    conn.close()
    if not card:
        return jsonify({'ok': False, 'message': 'Tarjeta no encontrada'}), 404

    _wb_log_comment(card_id, content)
    return jsonify({'ok': True})


# ==========================================
# REVISIONES MECÁNICAS — Rutas
# ==========================================
@app.route('/mechanic_revision/edit/<int:reception_id>', methods=['GET', 'POST'])
def mechanic_revision_edit(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    rev = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    if not rev:
        conn.close()
        flash('Recepción no encontrada.', 'danger')
        return redirect(url_for('index', module='🔬 Revisiones'))

    if request.method == 'POST':
        items = []
        for cat_idx, cat in enumerate(MECHANIC_REVISION_CATEGORIES):
            for item_idx, item in enumerate(cat['items']):
                variants = item.get('variants') or [None]
                obs = request.form.get(f'mrev_obs_{cat_idx}_{item_idx}', '').strip()
                for var_idx, variant in enumerate(variants):
                    mark = request.form.get(
                        f'mrev_mark_{cat_idx}_{item_idx}_{var_idx}', ''
                    )
                    if mark not in MREV_MARK_STATES:
                        mark = ''
                    items.append(
                        {
                            'cat': cat['name'],
                            'sys': item['system'],
                            'var': variant or '',
                            'cat_idx': cat_idx,
                            'item_idx': item_idx,
                            'var_idx': var_idx,
                            'mark': mark,
                            'obs': obs,
                        }
                    )

        notes = request.form.get('mechanic_notes', '').strip()
        new_status = request.form.get('mechanic_status', 'Pendiente')
        if new_status not in ('Pendiente', 'En Proceso', 'Completado'):
            new_status = 'Pendiente'

        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")
        actor = session.get('username', '')

        conn.execute(
            """
            UPDATE revisions SET
                mechanic_items_json = ?,
                mechanic_notes = ?,
                mechanic_username = ?,
                mechanic_date = COALESCE(NULLIF(mechanic_date, ''), ?),
                mechanic_status = ?
            WHERE id = ?
        """,
            (
                json.dumps(items, ensure_ascii=False),
                notes,
                actor,
                now_str,
                new_status,
                reception_id,
            ),
        )
        conn.commit()
        conn.close()

        log_action(
            'Revisión mecánica guardada',
            f'Revisión mecánica de la recepción #{reception_id} (estado: {new_status}).',
            'Recepción',
            reception_id,
        )
        flash('✓ Revisión técnica guardada.', 'success')
        return redirect(url_for('index', module='🔬 Revisiones'))
    # -------- GET: render --------
    try:
        existing_items = json.loads(rev['mechanic_items_json'] or '[]')
    except Exception:
        existing_items = []

    # Mapa dual: por índice (preferido) y por string (fallback para datos viejos)
    existing_by_idx = {}
    existing_by_str = {}
    for it in existing_items:
        if not isinstance(it, dict):
            continue
        mark_val = it.get('mark', '') or ''
        obs_val = it.get('obs', '') or ''

        # Por índice (nuevo formato)
        if 'cat_idx' in it and 'item_idx' in it and 'var_idx' in it:
            try:
                key_idx = (int(it['cat_idx']), int(it['item_idx']), int(it['var_idx']))
                existing_by_idx[key_idx] = {'mark': mark_val, 'obs': obs_val}
            except (ValueError, TypeError):
                pass

        # Por string (formato original)
        if 'sys' in it:
            key_str = (it.get('cat', ''), it.get('sys', ''), it.get('var', ''))
            existing_by_str[key_str] = {'mark': mark_val, 'obs': obs_val}
        elif 'status' in it:
            # Formato aún más viejo
            key_str = (it.get('category', ''), it.get('system', ''), '')
            existing_by_str[key_str] = {'mark': '', 'obs': obs_val}

    # Construir estructura para el template
    categories_with_items = []
    for cat_idx, cat in enumerate(MECHANIC_REVISION_CATEGORIES):
        cat_items = []
        for item_idx, item in enumerate(cat['items']):
            variants_data = []
            variants = item.get('variants') or [None]
            obs_val = ''
            for var_idx, variant in enumerate(variants):
                key_idx = (cat_idx, item_idx, var_idx)
                key_str = (cat['name'], item['system'], variant or '')
                found = existing_by_idx.get(key_idx) or existing_by_str.get(key_str, {})
                mark_val = found.get('mark', '')
                if mark_val not in MREV_MARK_STATES:
                    mark_val = ''
                if found.get('obs'):
                    obs_val = found['obs']
                variants_data.append(
                    {
                        'label': variant or '',
                        'field_mark': f'mrev_mark_{cat_idx}_{item_idx}_{var_idx}',
                        'mark': mark_val,
                    }
                )
            cat_items.append(
                {
                    'system': item['system'],
                    'variants': variants_data,
                    'field_obs': f'mrev_obs_{cat_idx}_{item_idx}',
                    'obs': obs_val,
                    'has_variants': bool(item.get('variants')),
                }
            )
        categories_with_items.append(
            {
                'name': cat['name'],
                'icon': cat['icon'],
                'components': cat_items,
            }
        )

    conn.close()
    mr_photos = _get_reception_photos(reception_id)

    role = session['role']
    allowed = get_allowed_modules(role)
    menu_options = [m for m in ALL_SYSTEM_MODULES if m in allowed]
    user_disp = session.get('user_display_name', session['username'])

    return render_template(
        'base.html',
        menu_options=menu_options,
        current_module="🔬 Revisiones",
        username=session['username'],
        user_display_name=user_disp,
        user_phone=session.get('user_phone', ''),
        role=role,
        tasa_usd=get_config('tasa_usd'),
        tasa_eur=get_config('tasa_eur'),
        budget_currency='USD',
        budget_client={},
        all_modules=ALL_SYSTEM_MODULES,
        all_budgets=[],
        filtered_budgets=[],
        pending_budgets=[],
        seller_summary=[],
        reminders=[],
        reminders_count=0,
        reminders_today_count=0,
        reminder_categories=[],
        r_search='',
        r_priority='Todas',
        r_category='Todas',
        r_status='Pendientes',
        r_author='Todos',
        r_sort='recientes',
        appointments_count=0,
        all_users=[],
        users_map={},
        users=[],
        roles_list=[],
        all_clients=[],
        c_search='',
        c_filter='Todos',
        c_sort='nombre',
        c_pagination=None,
        vehicle_list=[],
        v_search='',
        v_client='',
        v_sort='recientes',
        v_pagination=None,
        all_revisions=[],
        selected_client_doc='',
        client_budgets=[],
        client_revisions=[],
        f_type='Todos',
        f_user='Todos',
        f_date='',
        f_month='',
        f_start='',
        f_end='',
        log_user='Todos',
        log_action_filter='Todos',
        current_date=get_ve_time().strftime("%Y-%m-%d"),
        is_editing=False,
        is_editing_revision=False,
        templates=[],
        appointments=[],
        appointment_filter='Proximas',
        vehicle_search='',
        report_data=None,
        action_logs=[],
        action_types=[],
        dashboard_metrics=None,
        only_drafts=False,
        workshop_columns=[],
        cards_by_column={},
        workshop_stats=None,
        wb_view='board',
        workshop_history=[],
        workshop_history_stats=None,
        hw_user='Todos',
        hw_search='',
        hw_date_from='',
        hw_date_to='',
        all_users_list=[],
        workshop_mechanics=[],
        revision_kpis=None,
        recent_revisions=[],
        r_pagination=None,
        h_pagination=None,
        mr_is_editing=True,
        mr_revision=rev,
        mr_categories=categories_with_items,
        mr_photos=mr_photos,
        mechanic_revision_kpis=None,
        mr_pagination=None,
        mr_status='Todos',
    )


@app.route('/mechanic_revision/preview/<int:reception_id>')
@block_readonly_reception
def mechanic_revision_preview(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    conn.close()
    if not row:
        return "No encontrada", 404
    try:
        items = json.loads(row['mechanic_items_json'] or '[]')
    except Exception:
        items = []
    pdf_bytes = generate_mechanic_revision_pdf(
        client=row['client_name'],
        doc=row['client_doc'],
        phone=row['client_phone'],
        veh_name=row['vehicle_name'],
        veh_year=row['vehicle_year'],
        veh_plate=row['vehicle_plate'],
        date=(row['mechanic_date'] or row['date']),
        items=items,
        notes=(row['mechanic_notes'] if 'mechanic_notes' in row.keys() else '') or '',
        username=(row['mechanic_username'] if 'mechanic_username' in row.keys() else '')
        or '',
        veh_km=row['vehicle_km'] if 'vehicle_km' in row.keys() else '',
        veh_brand=row['vehicle_brand'] if 'vehicle_brand' in row.keys() else '',
        veh_color=row['vehicle_color'] if 'vehicle_color' in row.keys() else '',
        veh_model=row['vehicle_model'] if 'vehicle_model' in row.keys() else '',
        categories=MECHANIC_REVISION_CATEGORIES,
        reception_id=reception_id,
    )
    filename = (
        f"Revision_Mecanica_{(row['client_name'] or 'cliente').replace(' ', '_')}.pdf"
    )
    return send_file(
        BytesIO(pdf_bytes), mimetype='application/pdf', download_name=filename
    )


@app.route('/mechanic_revision/pdf/<int:reception_id>')
def mechanic_revision_pdf(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    conn.close()
    if not row:
        return "No encontrada", 404
    try:
        items = json.loads(row['mechanic_items_json'] or '[]')
    except Exception:
        items = []
    pdf_bytes = generate_mechanic_revision_pdf(
        client=row['client_name'],
        doc=row['client_doc'],
        phone=row['client_phone'],
        veh_name=row['vehicle_name'],
        veh_year=row['vehicle_year'],
        veh_plate=row['vehicle_plate'],
        date=(row['mechanic_date'] or row['date']),
        items=items,
        notes=(row['mechanic_notes'] if 'mechanic_notes' in row.keys() else '') or '',
        username=(row['mechanic_username'] if 'mechanic_username' in row.keys() else '')
        or '',
        veh_km=row['vehicle_km'] if 'vehicle_km' in row.keys() else '',
        veh_brand=row['vehicle_brand'] if 'vehicle_brand' in row.keys() else '',
        veh_color=row['vehicle_color'] if 'vehicle_color' in row.keys() else '',
        veh_model=row['vehicle_model'] if 'vehicle_model' in row.keys() else '',
        categories=MECHANIC_REVISION_CATEGORIES,
        reception_id=reception_id,  # ← agrega esta línea
    )
    filename = (
        f"Revision_Mecanica_{(row['client_name'] or 'cliente').replace(' ', '_')}.pdf"
    )
    return send_file(
        BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename,
    )


@app.route('/mechanic_revision/delete/<int:reception_id>', methods=['POST'])
def mechanic_revision_delete(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute(
        """
        UPDATE revisions SET
            mechanic_items_json = '[]',
            mechanic_notes = '',
            mechanic_username = '',
            mechanic_date = '',
            mechanic_status = 'Pendiente'
        WHERE id = ?
    """,
        (reception_id,),
    )
    conn.commit()
    conn.close()
    log_action(
        'Revisión mecánica eliminada',
        f'Se eliminó la revisión mecánica de la recepción #{reception_id}.',
        'Recepción',
        reception_id,
    )
    flash('Revisión técnica eliminada.', 'success')
    return redirect(url_for('index', module='🔬 Revisiones'))


def _mrev_draw_category_col(
    pdf, col_x, y_start, col_w, cats_slice, marks_by_idx, marks_by_str, cat_start_idx=0
):
    """Dibuja una columna de categorías. Solo muestra lo que está marcado."""

    def _s(v, d=''):
        if v is None:
            return d
        try:
            return str(v)
        except Exception:
            return d

    y = y_start
    sys_w = 40.5
    box_sz = 3.4
    var_pad = 1.2
    var_gap = 3.2
    ROW_H = 4.2
    CAT_HEADER_H = 5.5
    CAT_GAP = 2.5

    def _lookup(cat_idx, item_idx, var_idx, cat_name, sys_name, var_str):
        m = marks_by_idx.get((cat_idx, item_idx, var_idx))
        if m:
            return m
        m = marks_by_str.get((cat_name, sys_name, var_str))
        if m:
            return m
        if hasattr(_lookup, '_norm'):
            norm = _lookup._norm
        else:

            def _norm(s):
                if s is None:
                    return ''
                s = str(s).strip().upper()
                s = unicodedata.normalize('NFD', s)
                s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
                return ' '.join(s.split())

            _lookup._norm = _norm
            norm = _norm
        return marks_by_str.get(
            ('__norm__', norm(cat_name), norm(sys_name), norm(var_str)), ''
        ) or marks_by_str.get(
            ('__norm2__', norm(cat_name) + '|' + norm(sys_name) + '|' + norm(var_str)),
            '',
        )

    for local_cat_idx, cat in enumerate(cats_slice):
        if not isinstance(cat, dict):
            continue
        cat_idx = cat_start_idx + local_cat_idx
        cat_name = _s(cat.get('name', ''))
        cat_items = cat.get('items', [])

        # ===== Filtrar ítems con al menos una variante marcada =====
        filtered_items = []
        for item_idx, it in enumerate(cat_items):
            if not isinstance(it, dict):
                continue
            sys_name = _s(it.get('system', ''))
            variants = it.get('variants') or ['']
            marked_variants = []
            for var_idx, v in enumerate(variants):
                v_str = _s(v)
                mark = _lookup(cat_idx, item_idx, var_idx, cat_name, sys_name, v_str)
                if mark in ('*', '✓', 'X'):
                    marked_variants.append((v_str, var_idx, mark))
            if marked_variants:
                filtered_items.append(
                    {
                        'item_idx': item_idx,
                        'sys_name': sys_name,
                        'variants': marked_variants,
                    }
                )

        # Si la categoría no tiene ningún ítem marcado → se omite completa
        if not filtered_items:
            continue

        # ----- Header de la categoría -----
        pdf.set_xy(col_x, y)
        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_fill_color(230, 230, 230)
        pdf.set_draw_color(60, 60, 60)
        pdf.cell(col_w, CAT_HEADER_H, safe_txt(f"  {cat_name}"), 1, 0, 'L', True)

        for i, sym in enumerate(['*', 'V', 'X']):
            bx = col_x + col_w - 3 * (box_sz + 1) - 3 + i * (box_sz + 1)
            by = y + (CAT_HEADER_H - box_sz) / 2.0
            pdf.set_draw_color(40, 40, 40)
            pdf.rect(bx, by, box_sz, box_sz)
            pdf.set_xy(bx, by - 0.15)
            pdf.set_font('Arial', 'B', 6)
            pdf.cell(box_sz, box_sz, safe_txt(sym), 0, 0, 'C')

        y += CAT_HEADER_H

        avail_w = col_w - sys_w - 1.5

        for entry in filtered_items:
            sys_name = entry['sys_name']
            variants = entry['variants']

            pdf.set_font('Arial', '', 6)
            v_items = []
            for v_str, var_idx, mark in variants:
                w_label = pdf.get_string_width(v_str) + 1.8
                w_total = w_label + var_pad + box_sz + var_gap
                v_items.append((v_str, var_idx, mark, w_label, w_total))

            # Repartir variantes en filas si no caben en una sola
            rows_of_vars = []
            cur_row, cur_w = [], 0.0
            for item_v in v_items:
                w_total = item_v[4]
                if cur_w + w_total <= avail_w or not cur_row:
                    cur_row.append(item_v)
                    cur_w += w_total
                else:
                    rows_of_vars.append(cur_row)
                    cur_row = [item_v]
                    cur_w = w_total
            if cur_row:
                rows_of_vars.append(cur_row)

            for r_idx, row_vars in enumerate(rows_of_vars):
                pdf.set_draw_color(150, 150, 150)
                pdf.rect(col_x, y, col_w, ROW_H)

                if r_idx == 0:
                    pdf.set_xy(col_x + 1.5, y + 1.4)
                    pdf.set_font('Arial', 'B', 6.2)
                    pdf.cell(sys_w - 2, 3, safe_txt(sys_name[:34]), 0, 0, 'L')
                else:
                    pdf.set_draw_color(190, 190, 190)
                    pdf.line(col_x + sys_w, y, col_x + sys_w, y + ROW_H)

                vx = col_x + sys_w + 0.5
                for v_str, var_idx, mark, w_label, w_total in row_vars:
                    pdf.set_xy(vx, y + 1.4)
                    pdf.set_font('Arial', '', 6)
                    pdf.cell(w_label, 3, safe_txt(v_str), 0, 0, 'L')

                    bx = vx + w_label + var_pad
                    by = y + (ROW_H - box_sz) / 2.0
                    pdf.set_draw_color(40, 40, 40)
                    pdf.rect(bx, by, box_sz, box_sz)

                    sym = {'*': '*', '✓': 'V', 'X': 'X'}.get(mark, mark)
                    pdf.set_xy(bx, by - 0.15)
                    pdf.set_font('Arial', 'B', 6.8)
                    pdf.cell(box_sz, box_sz, safe_txt(sym), 0, 0, 'C')

                    vx += w_total

                y += ROW_H

        y += CAT_GAP

    return y


def generate_mechanic_revision_pdf(
    client,
    doc,
    phone,
    veh_name,
    veh_year,
    veh_plate,
    date,
    items,
    notes,
    username,
    veh_km='',
    veh_brand='',
    veh_color='',
    veh_model='',
    categories=None,
    reception_id=None,
):
    """
    PDF horizontal tamaño Carta que replica la planilla física de Ruedas y Suspensiones.
    """

    def _s(v, d=''):
        if v is None:
            return d
        try:
            return str(v)
        except Exception:
            return d

    try:
        # Letter horizontal: 279.4 x 215.9 mm
        pdf = PDF(orientation='L', unit='mm', format='letter')
        pdf.set_auto_page_break(False)
        pdf.set_margins(6, 6, 6)
        pdf.add_page()
        pdf.set_title(safe_txt(f"Orden de Revisión - {_s(client)}"))

        MARGIN = 4
        TOTAL_W = 279.4 - 2 * MARGIN
        COL_W = 134
        COL_LX = MARGIN
        COL_RX = MARGIN + COL_W + 3.4

        # ================== ENCABEZADO ==================
        logo_path = os.path.join(BASE_DIR, "static", "img", "logopdf.png")
        if os.path.exists(logo_path):
            try:
                pdf.image(logo_path, MARGIN + 0.5, 5, 10)
            except Exception:
                pass

        # Empresa
        pdf.set_font('Arial', 'B', 13)
        pdf.set_xy(MARGIN + 15, 4)
        pdf.cell(120, 5, safe_txt("RUEDAS Y SUSPENSIONES"), 0, 0, 'L')

        pdf.set_font('Arial', '', 6)
        pdf.set_xy(MARGIN + 15, 9.5)
        pdf.cell(120, 2.5, safe_txt("RIF: J-40083124-5"), 0, 0, 'L')
        pdf.set_xy(MARGIN + 15, 11.8)
        pdf.cell(
            120,
            2.5,
            safe_txt("Av. Bolívar Norte, Local cívico 123-31, Valencia Edo. Carabobo"),
            0,
            0,
            'L',
        )
        pdf.set_xy(MARGIN + 15, 14.1)
        pdf.cell(
            120,
            2.5,
            safe_txt(
                "Telf: (0241) 825.60.24 / 825.39.25   ·   Asistencia Vial VIP 0414.189.2294"
            ),
            0,
            0,
            'L',
        )

        # Orden de revisión (arriba derecha)
        pdf.set_font('Arial', 'B', 11)
        pdf.set_xy(180, 5)
        pdf.cell(90, 5, safe_txt("ORDEN DE REVISIÓN"), 0, 0, 'C')

        pdf.set_font('Arial', 'B', 14)
        pdf.set_text_color(200, 0, 0)
        if reception_id:
            pdf.set_xy(230, 5)
            pdf.cell(45, 5, safe_txt(f"N° {int(reception_id):06d}"), 0, 0, 'R')
        pdf.set_text_color(0, 0, 0)

        pdf.set_font('Arial', '', 8)
        pdf.set_xy(180, 12)
        pdf.cell(45, 4, safe_txt(f"Fecha: {_s(date)}"), 0, 0, 'L')
        pdf.set_xy(230, 12)
        pdf.cell(45, 4, safe_txt(f"Mecánico: {_s(username) or '—'}"), 0, 0, 'R')

        # Línea separadora
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.3)
        pdf.line(MARGIN, 17.5, MARGIN + TOTAL_W, 17.5)

        # ================== DATOS DEL VEHÍCULO ==================
        y = 18
        pdf.set_xy(MARGIN, y)
        pdf.set_font('Arial', 'B', 8)
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(TOTAL_W, 5, safe_txt("  DATOS DEL VEHÍCULO"), 1, 0, 'L', True)

        y += 5
        col_h = 5
        wq = TOTAL_W / 4.0  # ancho de cada bloque (Marca, Modelo, Año, Color)

        campos_fila1 = [
            ('Marca', _s(veh_brand) or '—'),
            ('Modelo', _s(veh_model) or '—'),
            ('Año', _s(veh_year) or '—'),
            ('Color', _s(veh_color) or '—'),
        ]
        campos_fila2 = [
            ('Km', _s(veh_km) or '—'),
            ('Placa', _s(veh_plate) or '—'),
            ('Cliente', _s(client)),
            ('C.I/RIF', _s(doc)),
        ]

        for fila in (campos_fila1, campos_fila2):
            pdf.set_xy(MARGIN, y)
            for lbl, val in fila:
                pdf.set_font('Arial', 'B', 6.5)
                pdf.cell(wq * 0.32, col_h, safe_txt(f" {lbl}:"), 1, 0, 'L')
                pdf.set_font('Arial', '', 7)
                pdf.cell(wq * 0.68, col_h, safe_txt(f" {val}"), 1, 0, 'L')
            y += col_h

        y += 1

        # ================== ÍNDICE DE MARCAS ==================
        marks_by_idx = {}
        marks_by_str = {}
        for it in items or []:
            if not isinstance(it, dict):
                continue
            mark_val = _s(it.get('mark', ''))
            # Por índice
            if 'cat_idx' in it and 'item_idx' in it and 'var_idx' in it:
                try:
                    marks_by_idx[
                        (int(it['cat_idx']), int(it['item_idx']), int(it['var_idx']))
                    ] = mark_val
                except (ValueError, TypeError):
                    pass
            # Por string
            key_str = (
                _s(it.get('cat', '')),
                _s(it.get('sys', '')),
                _s(it.get('var', '')),
            )
            marks_by_str[key_str] = mark_val

            # Por string normalizado (una sola entrada con tuple)
            def _norm(s):
                if s is None:
                    return ''
                s = str(s).strip().upper()
                s = unicodedata.normalize('NFD', s)
                s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
                return ' '.join(s.split())

            marks_by_str[
                (
                    '__norm__',
                    _norm(it.get('cat', '')),
                    _norm(it.get('sys', '')),
                    _norm(it.get('var', '')),
                )
            ] = mark_val

        # ================== CATEGORÍAS ==================
        cats = categories if isinstance(categories, list) else []
        left_cats = cats[:3]
        right_cats = cats[3:6]

        end_left = _mrev_draw_category_col(
            pdf, COL_LX, y, COL_W, left_cats, marks_by_idx, marks_by_str, 0
        )
        end_right = _mrev_draw_category_col(
            pdf, COL_RX, y, COL_W, right_cats, marks_by_idx, marks_by_str, 3
        )
        y = max(end_left, end_right) + 2

        # ================== OBSERVACIONES ==================
        pdf.set_xy(MARGIN, y)
        pdf.set_font('Arial', 'B', 8)
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(TOTAL_W, 5, safe_txt("  OBSERVACIONES"), 1, 0, 'L', True)

        y += 5
        box_h = 11
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(MARGIN, y, TOTAL_W, box_h)
        notes_str = _s(notes).strip()
        if notes_str:
            pdf.set_xy(MARGIN + 2, y + 1.5)
            pdf.set_font('Arial', '', 7.5)
            try:
                pdf.multi_cell(TOTAL_W - 4, 4, safe_txt(notes_str), 0, 'L')
            except Exception:
                pdf.cell(TOTAL_W - 4, 4, safe_txt(notes_str[:200]), 0, 0, 'L')
        y += box_h + 1.5

        # ================== FIRMAS ==================
        y += 2
        firm_h = 12
        firm_w = TOTAL_W / 3.0

        # --- Fila 1: "RECEPCIÓN DE VEHÍCULO" + 2 firmas ---
        pdf.set_draw_color(0, 0, 0)
        pdf.set_fill_color(220, 220, 220)
        pdf.rect(MARGIN, y, firm_w, firm_h, 'DF')
        pdf.set_xy(MARGIN, y + 2)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(firm_w, 4, safe_txt("RECEPCIÓN DE VEHÍCULO"), 0, 0, 'C')

        pdf.rect(MARGIN + firm_w, y, firm_w, firm_h)
        pdf.set_xy(MARGIN + firm_w, y + 1)
        pdf.set_font('Arial', 'B', 6.5)
        pdf.cell(firm_w, 3, safe_txt("FIRMA CLIENTE"), 0, 0, 'C')
        pdf.line(
            MARGIN + firm_w + 6, y + firm_h - 3, MARGIN + 2 * firm_w - 6, y + firm_h - 3
        )

        pdf.rect(MARGIN + 2 * firm_w, y, firm_w, firm_h)
        pdf.set_xy(MARGIN + 2 * firm_w, y + 1)
        pdf.set_font('Arial', 'B', 6.5)
        pdf.cell(firm_w, 3, safe_txt("FIRMA EJECUTIVO"), 0, 0, 'C')
        pdf.line(
            MARGIN + 2 * firm_w + 6,
            y + firm_h - 3,
            MARGIN + 3 * firm_w - 6,
            y + firm_h - 3,
        )

        y += firm_h + 1.5

        # --- Fila 2: "ENTREGA DE VEHÍCULO" + 2 firmas ---
        pdf.set_draw_color(0, 0, 0)
        pdf.set_fill_color(220, 220, 220)
        pdf.rect(MARGIN, y, firm_w, firm_h, 'DF')
        pdf.set_xy(MARGIN, y + 2)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(firm_w, 4, safe_txt("ENTREGA DE VEHÍCULO"), 0, 0, 'C')

        pdf.rect(MARGIN + firm_w, y, firm_w, firm_h)
        pdf.set_xy(MARGIN + firm_w, y + 1)
        pdf.set_font('Arial', 'B', 6.5)
        pdf.cell(firm_w, 3, safe_txt("RECIBE CONFORME"), 0, 0, 'C')
        pdf.line(
            MARGIN + firm_w + 6, y + firm_h - 3, MARGIN + 2 * firm_w - 6, y + firm_h - 3
        )

        pdf.rect(MARGIN + 2 * firm_w, y, firm_w, firm_h)
        pdf.set_xy(MARGIN + 2 * firm_w, y + 1)
        pdf.set_font('Arial', 'B', 6.5)
        pdf.cell(firm_w, 3, safe_txt("FECHA / HORA"), 0, 0, 'C')
        pdf.line(
            MARGIN + 2 * firm_w + 6,
            y + firm_h - 3,
            MARGIN + 3 * firm_w - 6,
            y + firm_h - 3,
        )

        y += firm_h + 1

        # ================== LEYENDA ==================
        pdf.set_xy(MARGIN, y)
        pdf.set_font('Arial', 'I', 6.5)
        pdf.cell(
            TOTAL_W,
            3.5,
            safe_txt("NOTA: Esta revisión es visual, no cubre daños ocultos."),
            0,
            1,
            'C',
        )
        pdf.set_xy(MARGIN, y + 3.5)
        pdf.cell(
            TOTAL_W,
            3.5,
            safe_txt(
                "* Presenta desgaste parcial, reparable.    V Se encuentra en buen estado.    X Mal estado, debe ser reemplazado."
            ),
            0,
            1,
            'C',
        )

        y += 8

        # ================== PIE ==================
        pdf.set_xy(MARGIN, y)
        pdf.set_font('Arial', 'I', 7)
        pdf.cell(TOTAL_W, 3.5, safe_txt("www.ruedasysuspensiones.com"), 0, 0, 'C')

        return pdf_to_bytes(pdf)

    except Exception as e:
        import traceback

        print(f"[PDF mecánica] ERROR generando PDF: {e}")
        traceback.print_exc()
        try:
            pdf = PDF(orientation='L', unit='mm', format='letter')
            pdf.add_page()
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, safe_txt("Error al generar el PDF"), 0, 1, 'C')
            pdf.set_font('Arial', '', 9)
            pdf.multi_cell(0, 5, safe_txt(str(e)), 0, 'L')
            return pdf_to_bytes(pdf)
        except Exception:
            pdf = PDF()
            pdf.add_page()
            return pdf_to_bytes(pdf)


# ==========================================
# RECEPCIONES — Planilla de Inspección (página completa)
# ==========================================
@app.route('/reception_edit/<int:reception_id>', methods=['GET', 'POST'])
@block_readonly_reception
def reception_edit(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    rev = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    if not rev:
        conn.close()
        flash('Recepción no encontrada.', 'danger')
        return redirect(url_for('index', module='📥 Recepciones'))

    # -------- POST --------
    if request.method == 'POST':
        # Datos básicos
        client_name = request.form.get('client_name', '').strip()
        client_lastname = request.form.get('client_lastname', '').strip()
        client_doc = up(request.form.get('client_doc', ''))
        client_phone = request.form.get('client_phone', '').strip()
        full_client = f"{client_name} {client_lastname}".strip() or "SIN NOMBRE"

        razon_social = request.form.get('reception_razon_social', '').strip()
        rif_ci = request.form.get('reception_rif_ci', '').strip()
        email = request.form.get('reception_email', '').strip()
        direccion_fiscal = request.form.get('reception_direccion_fiscal', '').strip()

        veh_brand = up(request.form.get('veh_brand', ''))
        veh_color = up(request.form.get('veh_color', ''))
        veh_model = up(request.form.get('veh_model', ''))
        veh_year = request.form.get('veh_year', '').strip()
        veh_km = request.form.get('veh_km', '').strip()
        veh_plate = up(request.form.get('veh_plate', ''))
        transmision = request.form.get('reception_transmision', '').strip()
        motor = request.form.get('reception_motor', '').strip()
        clave = request.form.get('reception_clave', '').strip()

        # Fallas 1-5
        fallas = []
        for i in range(1, 6):
            f = request.form.get(f'fallas_{i}', '').strip()
            fallas.append(f)

        es_retorno = 1 if request.form.get('reception_es_retorno') == '1' else 0

        # Items checklist
        items = []
        for idx, sys_name in enumerate(RECEPTION_FIELDS_FLAT):
            mark = request.form.get(f'reception_mark_{idx}', '')
            if mark not in RECEPTION_MARK_STATES:
                mark = ''
            items.append({'sys': sys_name, 'mark': mark})

        # ===== Luces del Tablero (indicadores de alerta) =====
        dash_lights_on = []
        for idx, light_name in enumerate(DASHBOARD_LIGHTS):
            val = request.form.get(f'dash_light_{idx}', '')
            if val == 'ON':
                dash_lights_on.append(light_name)

        fuel_level = request.form.get('reception_fuel_level', '').strip()
        if fuel_level and fuel_level not in FUEL_LEVEL_OPTIONS:
            fuel_level = ''
        temperature_raw = (
            request.form.get('reception_temperature', '').strip().replace(',', '.')
        )
        try:
            temperature = str(float(temperature_raw)) if temperature_raw else ''
            if temperature and not (-40 <= float(temperature) <= 200):
                temperature = ''
        except (TypeError, ValueError):
            temperature = ''

        # ===== Boceto de daños =====
        reception_sketch_json = request.form.get('reception_sketch_json', '{}') or '{}'
        reception_status = request.form.get('reception_status', 'Pendiente')
        if reception_status not in ('Borrador', 'Pendiente', 'Completado'):
            reception_status = 'Borrador'

        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
        fecha_hora = get_ve_time()
        fecha_str = fecha_hora.strftime("%Y-%m-%d")
        hora_str = fecha_hora.strftime("%I:%M")
        am_pm = fecha_hora.strftime("%p").replace("AM", "AM").replace("PM", "PM")

        # Guardar/actualizar cliente y registrar vehículo en el garaje
        try:
            if client_doc:
                save_or_update_client(
                    doc=client_doc,
                    name=client_name,
                    phone=client_phone,
                    veh_name='',  # La planilla no tiene un campo "nombre de vehículo", se deja vacío
                    veh_year=veh_year,
                    veh_plate=veh_plate,
                    veh_brand=veh_brand,
                    veh_color=veh_color,
                    veh_model=veh_model,
                    lastname=client_lastname,
                )
        except Exception as e:
            print(f"[reception_edit] Error guardando cliente/vehículo: {e}")

        conn.execute(
            """
            UPDATE revisions SET
                client_name = ?, client_doc = ?, client_phone = ?,
                vehicle_brand = ?, vehicle_color = ?, vehicle_model = ?,
                vehicle_year = ?, vehicle_km = ?, vehicle_plate = ?,
                reception_items_json = ?,
                reception_fallas_json = ?,
                reception_razon_social = ?,
                reception_rif_ci = ?,
                reception_email = ?,
                reception_direccion_fiscal = ?,
                reception_es_retorno = ?,
                reception_fecha = ?,
                reception_hora = ?,
                reception_am_pm = ?,
                reception_transmision = ?,
                reception_motor = ?,
                reception_clave = ?,
                reception_status = ?,
                reception_sketch_json = ?,
                reception_lights_json = ?,
                reception_fuel_level = ?,
                reception_temperature = ?
            WHERE id = ?
        """,
            (
                full_client,
                client_doc,
                client_phone,
                veh_brand,
                veh_color,
                veh_model,
                veh_year,
                veh_km,
                veh_plate,
                json.dumps(items, ensure_ascii=False),
                json.dumps(fallas, ensure_ascii=False),
                razon_social,
                rif_ci,
                email,
                direccion_fiscal,
                es_retorno,
                fecha_str,
                hora_str,
                am_pm,
                transmision,
                motor,
                clave,
                reception_status,
                reception_sketch_json,
                json.dumps(dash_lights_on, ensure_ascii=False),
                fuel_level,
                temperature,
                reception_id,
            ),
        )
        conn.commit()
        conn.close()

        log_action(
            'Recepción guardada',
            f'Planilla de recepción #{reception_id} guardada.',
            'Recepción',
            reception_id,
        )
        flash('✓ Recepción guardada.', 'success')
        return redirect(url_for('index', module='📥 Recepciones'))

    # -------- GET --------
    try:
        items_data = json.loads(rev['reception_items_json'] or '[]')
    except Exception:
        items_data = []
    try:
        fallas_data = json.loads(rev['reception_fallas_json'] or '[]')
    except Exception:
        fallas_data = []
    while len(fallas_data) < 5:
        fallas_data.append('')

    # Mapa por sys_name → mark
    marks_map = {
        it.get('sys'): it.get('mark', '') for it in items_data if isinstance(it, dict)
    }

    # Construir estructura para el template
    ui_data = []
    flat_idx = 0
    for cat in RECEPTION_CHECKLIST_UI:
        cat_items = []
        for item in cat['items']:
            item_fields = []
            for f_name in item['fields']:
                label = 'Izq' if 'IZQ' in f_name else ('Der' if 'DER' in f_name else '')
                item_fields.append(
                    {
                        'flat_idx': flat_idx,
                        'mark': marks_map.get(f_name, ''),
                        'label': label,
                    }
                )
                flat_idx += 1
            cat_items.append({'name': item['name'], 'fields': item_fields})
        ui_data.append({'title': cat['title'], 'icon': cat['icon'], 'items': cat_items})

    # ===== Fotos de la recepción =====
    reception_photos = _get_reception_photos(reception_id)
    # ===== Boceto de daños =====
    reception_sketch_json = (
        rev['reception_sketch_json']
        if 'reception_sketch_json' in rev.keys() and rev['reception_sketch_json']
        else '{}'
    )

    # 🔥 NUEVO: Obtener el apellido real desde la base de datos
    client_lastname_db = ''
    if rev['client_doc']:
        c_row = conn.execute(
            "SELECT client_lastname FROM clients WHERE client_doc = ?",
            (rev['client_doc'],),
        ).fetchone()
        if c_row:
            client_lastname_db = c_row['client_lastname'] or ''

    # Separar el nombre de forma inteligente (usando la nueva función)
    parts = split_full_name(rev['client_name'], client_lastname_db)
    first_name = parts[0]
    lastname = parts[1]

    # ===== Luces del Tablero =====
    try:
        reception_lights_active = json.loads(rev['reception_lights_json'] or '[]')
    except Exception:
        reception_lights_active = []
    if not isinstance(reception_lights_active, list):
        reception_lights_active = []

    reception_fuel_level = ''
    try:
        reception_fuel_level = (
            (rev['reception_fuel_level'] or '')
            if 'reception_fuel_level' in rev.keys()
            else ''
        )
    except Exception:
        reception_fuel_level = ''

    first_name = parts[0] if len(parts) > 1 else (rev['client_name'] or '')
    lastname = parts[1] if len(parts) > 1 else ''

    conn.close()

    role = session['role']
    allowed = get_allowed_modules(role)
    menu_options = [m for m in ALL_SYSTEM_MODULES if m in allowed]
    user_disp = session.get('user_display_name', session['username'])

    return render_template(
        'base.html',
        menu_options=menu_options,
        current_module="📥 Recepciones",
        username=session['username'],
        user_display_name=user_disp,
        user_phone=session.get('user_phone', ''),
        role=role,
        tasa_usd=get_config('tasa_usd'),
        tasa_eur=get_config('tasa_eur'),
        budget_currency='USD',
        budget_client={},
        all_modules=ALL_SYSTEM_MODULES,
        all_budgets=[],
        filtered_budgets=[],
        pending_budgets=[],
        seller_summary=[],
        reminders=[],
        reminders_count=0,
        reminders_today_count=0,
        reminder_categories=[],
        r_search='',
        r_priority='Todas',
        r_category='Todas',
        r_status='Pendientes',
        r_author='Todos',
        r_sort='recientes',
        appointments_count=0,
        all_users=[],
        users_map={},
        users=[],
        roles_list=[],
        all_clients=[],
        c_search='',
        c_filter='Todos',
        c_sort='nombre',
        c_pagination=None,
        vehicle_list=[],
        v_search='',
        v_client='',
        v_sort='recientes',
        v_pagination=None,
        all_revisions=[],
        selected_client_doc='',
        client_budgets=[],
        client_revisions=[],
        f_type='Todos',
        f_user='Todos',
        f_date='',
        f_month='',
        f_start='',
        f_end='',
        log_user='Todos',
        log_action_filter='Todos',
        current_date=get_ve_time().strftime("%Y-%m-%d"),
        is_editing=False,
        is_editing_revision=False,
        templates=[],
        appointments=[],
        appointment_filter='Proximas',
        vehicle_search='',
        report_data=None,
        action_logs=[],
        action_types=[],
        dashboard_metrics=None,
        only_drafts=False,
        workshop_columns=[],
        cards_by_column={},
        workshop_stats=None,
        wb_view='board',
        workshop_history=[],
        workshop_history_stats=None,
        hw_user='Todos',
        hw_search='',
        hw_date_from='',
        hw_date_to='',
        all_users_list=[],
        workshop_mechanics=[],
        revision_kpis=None,
        recent_revisions=[],
        r_pagination=None,
        h_pagination=None,
        mr_is_editing=False,
        mr_revision=None,
        mr_categories=[],
        mr_checklist_flat=[],
        mechanic_revision_kpis=None,
        mr_pagination=None,
        mr_status='Todos',
        reception_is_editing=True,
        reception_data=rev,
        reception_first_name=first_name,
        reception_lastname=lastname,
        reception_ui_data=ui_data,
        reception_photos=reception_photos,
        reception_sketch_json=reception_sketch_json,
        reception_fallas=fallas_data,
        reception_lights_active=reception_lights_active,
        reception_fuel_level=reception_fuel_level,
        reception_temperature=(
            rev['reception_temperature']
            if 'reception_temperature' in rev.keys()
            else ''
        ),
        client_lastname_db=client_lastname_db,
    )  # <-- FIX BUG #10


@app.route('/reception_preview/<int:reception_id>')
def reception_preview(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    conn.close()
    if not row:
        return "No encontrada", 404
    try:
        items = json.loads(row['reception_items_json'] or '[]')
    except Exception:
        items = []
    try:
        fallas = json.loads(row['reception_fallas_json'] or '[]')
    except Exception:
        fallas = []
    pdf_bytes = generate_reception_pdf(row, items, fallas)
    filename = (
        f"Planilla_Inspeccion_{(row['client_name'] or 'cliente').replace(' ', '_')}.pdf"
    )
    return send_file(
        BytesIO(pdf_bytes), mimetype='application/pdf', download_name=filename
    )


@app.route('/reception_pdf/<int:reception_id>')
def reception_pdf(reception_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    conn.close()
    if not row:
        return "No encontrada", 404
    try:
        items = json.loads(row['reception_items_json'] or '[]')
    except Exception:
        items = []
    try:
        fallas = json.loads(row['reception_fallas_json'] or '[]')
    except Exception:
        fallas = []
    pdf_bytes = generate_reception_pdf(row, items, fallas)
    filename = (
        f"Planilla_Inspeccion_{(row['client_name'] or 'cliente').replace(' ', '_')}.pdf"
    )
    return send_file(
        BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename,
    )


def generate_reception_pdf(row, items, fallas):
    """Genera el PDF de la Planilla de Inspección (Letter horizontal, 1 hoja)."""

    def _s(v, d=''):
        if v is None:
            return d
        try:
            return str(v)
        except Exception:
            return d

    def _g(key, default=''):
        try:
            return row[key] if key in row.keys() and row[key] is not None else default
        except Exception:
            return default

    try:
        pdf = PDF(orientation='L', unit='mm', format='letter')
        pdf.set_auto_page_break(False)
        pdf.set_margins(5, 5, 5)
        pdf.add_page()
        pdf.set_title(safe_txt(f"Planilla Inspección - {_g('client_name')}"))

        M = 5
        TW = 279.4 - 2 * M  # 269.4

        # ================== ENCABEZADO ==================
        logo_path = os.path.join(BASE_DIR, "static", "img", "logopdf.png")
        if os.path.exists(logo_path):
            try:
                pdf.image(logo_path, M + 0.5, 4, 11)
            except Exception:
                pass

        pdf.set_font('Arial', 'B', 12)
        pdf.set_xy(M + 14, 4)
        pdf.cell(95, 5, safe_txt("RUEDAS Y SUSPENSIONES"), 0, 0, 'L')

        pdf.set_font('Arial', '', 5.5)
        pdf.set_xy(M + 14, 9)
        pdf.cell(95, 2.5, safe_txt("RIF: J-40083124-5"), 0, 0, 'L')
        pdf.set_xy(M + 14, 11.2)
        pdf.cell(
            95,
            2.5,
            safe_txt("Av. Bolívar Norte, Local 123-31, Valencia Edo. Carabobo"),
            0,
            0,
            'L',
        )
        pdf.set_xy(M + 14, 13.4)
        pdf.cell(
            95,
            2.5,
            safe_txt(
                "Telf: (0241) 825.60.24 / 825.39.25 · Asistencia Vial VIP 0414-189-2294"
            ),
            0,
            0,
            'L',
        )

        # Cuadro superior derecho
        pdf.set_font('Arial', 'B', 10)
        pdf.set_xy(200, 4)
        pdf.cell(50, 5, safe_txt("PLANILLA DE RECEPCIÓN A-"), 0, 0, 'C')
        pdf.set_font('Arial', 'B', 12)
        pdf.set_text_color(200, 0, 0)
        pdf.set_xy(250, 4)
        pdf.cell(24, 5, safe_txt(f"{int(_g('id', 0)):06d}"), 0, 0, 'R')
        pdf.set_text_color(0, 0, 0)

        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_xy(200, 10)
        pdf.cell(15, 3.5, safe_txt("Fecha:"), 0, 0, 'L')
        pdf.set_font('Arial', '', 7.5)
        pdf.cell(
            59, 3.5, safe_txt(_g('reception_fecha') or _g('date', '')[:10]), 0, 0, 'L'
        )

        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_xy(200, 13.5)
        pdf.cell(15, 3.5, safe_txt("Hora:"), 0, 0, 'L')
        pdf.set_font('Arial', '', 7.5)
        hora_val = _s(_g('reception_hora', ''))
        pdf.cell(35, 3.5, safe_txt(hora_val), 0, 0, 'L')
        pdf.cell(8, 3.5, safe_txt("AM"), 1, 0, 'C')
        pdf.cell(8, 3.5, safe_txt("PM"), 1, 0, 'C')
        pdf.ln(0)

        pdf.set_draw_color(0, 0, 0)
        pdf.line(M, 17, M + TW, 17)

        # ================== DATOS PERSONALES ==================
        y = 18
        lbl_w = 55
        half_w = (TW - lbl_w) / 2.0

        def _bar_row(y0, label, campos):
            """Dibuja una fila con barra gris a la izquierda + N campos."""
            pdf.set_xy(M, y0)
            pdf.set_font('Arial', 'B', 7.5)
            pdf.set_fill_color(220, 220, 220)
            pdf.set_draw_color(0, 0, 0)
            pdf.cell(lbl_w, 5.5, safe_txt(f" {label}"), 1, 0, 'L', True)
            n = max(len(campos), 1)
            w_each = (TW - lbl_w) / n
            # Si son 5+ campos, achicar label; si son 2-4, usar proporción generosa
            if n >= 5:
                lab_ratio, val_ratio, lab_font = 0.42, 0.58, 5.8
            elif n == 4:
                lab_ratio, val_ratio, lab_font = 0.36, 0.64, 6.4
            else:
                lab_ratio, val_ratio, lab_font = 0.32, 0.68, 6.5
            for lbl, val in campos:
                pdf.set_font('Arial', 'B', lab_font)
                pdf.cell(w_each * lab_ratio, 5.5, safe_txt(f" {lbl}:"), 1, 0, 'L')
                pdf.set_font('Arial', '', 7)
                pdf.cell(w_each * val_ratio, 5.5, safe_txt(f" {val}"), 1, 0, 'L')

        # Fila 1: Nombre y Apellido | Cédula
        _bar_row(
            y,
            "DATOS PERSONALES",
            [("Nombre y Apellido", _g('client_name')), ("Cédula", _g('client_doc'))],
        )
        y += 5.5

        # Fila 2: Razón Social | RIF/C.I
        _bar_row(
            y,
            "DATOS FACTURACIÓN",
            [
                ("Razón Social", _g('reception_razon_social')),
                ("Rif/C.I", _g('reception_rif_ci')),
            ],
        )
        y += 5.5

        # Fila 3: Teléfonos | Email | Dirección Fiscal
        _bar_row(
            y,
            "Teléfonos:",
            [("", _g('client_phone')), ("Email", _g('reception_email'))],
        )
        y += 5.5
        # Dirección fiscal a ancho completo
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_fill_color(220, 220, 220)
        pdf.cell(lbl_w, 5.5, safe_txt(" Dirección Fiscal:"), 1, 0, 'L', True)
        pdf.set_font('Arial', '', 7)
        pdf.cell(
            TW - lbl_w, 5.5, safe_txt(f" {_g('reception_direccion_fiscal')}"), 1, 0, 'L'
        )
        y += 5.5

        # Fila: Marca | Color | Modelo | Año
        _bar_row(
            y,
            "DATOS DEL VEHÍCULO",
            [
                ("Marca", _g('vehicle_brand')),
                ("Color", _g('vehicle_color')),
                ("Modelo", _g('vehicle_model')),
                ("Año", _g('vehicle_year')),
            ],
        )
        y += 5.5
        # Fila: Km | Placa | Transmisión | Motor | Clave
        _bar_row(
            y,
            " ",
            [
                ("KM", _g('vehicle_km')),
                ("Placa", _g('vehicle_plate')),
                ("Transmisión", _g('reception_transmision')),
                ("Motor", _g('reception_motor')),
                ("Clave", _g('reception_clave')),
            ],
        )
        y += 5.5

        # ================== OBSERVACIONES / FALLAS ==================
        # Barra de título (full width)
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_fill_color(220, 220, 220)
        pdf.set_draw_color(0, 0, 0)
        pdf.cell(
            TW, 5, safe_txt("  OBSERVACIONES / FALLAS DECLARADAS"), 1, 1, 'L', True
        )
        y += 5

        # 5 líneas enumeradas 1..5
        num_w = 8
        for i in range(5):
            pdf.set_xy(M, y)
            pdf.set_font('Arial', 'B', 7)
            pdf.cell(num_w, 4.5, safe_txt(f" {i+1}."), 1, 0, 'L')
            pdf.set_font('Arial', '', 7.5)
            pdf.cell(
                TW - num_w,
                4.5,
                safe_txt(f" {fallas[i] if len(fallas) > i else ''}"),
                1,
                1,
                'L',
            )
            y += 4.5

        # ================== LUCES DEL TABLERO + GASOLINA ==================
        y += 1
        # Título de la sección
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'B', 7.5)
        pdf.set_fill_color(220, 220, 220)
        pdf.set_draw_color(0, 0, 0)
        pdf.cell(
            TW,
            5,
            safe_txt(
                "  LUCES DEL TABLERO (INDICADORES ENCENDIDOS)  ·  ⛽ NIVEL DE GASOLINA"
            ),
            1,
            1,
            'L',
            True,
        )
        y += 5

        # Recuperar luces encendidas y nivel de gasolina
        try:
            lights_on = json.loads(_g('reception_lights_json', '[]') or '[]')
        except Exception:
            lights_on = []
        if not isinstance(lights_on, list):
            lights_on = []
        fuel_level_val = _s(_g('reception_fuel_level', ''))

        # Layout: 5 columnas x 3 filas = 15 celdas para las 15 luces
        total_lights = len(DASHBOARD_LIGHTS)
        cols_n = 5
        rows_n = (total_lights + cols_n - 1) // cols_n  # 3 filas
        cell_h = 5.0
        cell_w = TW / cols_n

        box_sz = 3.0
        for i, light_name in enumerate(DASHBOARD_LIGHTS):
            row_i = i // cols_n
            col_i = i % cols_n
            cx = M + col_i * cell_w
            cy = y + row_i * cell_h

            pdf.set_draw_color(160, 160, 160)
            pdf.rect(cx, cy, cell_w, cell_h)

            # Casilla
            bx = cx + 1.5
            by = cy + (cell_h - box_sz) / 2.0
            pdf.set_draw_color(50, 50, 50)
            pdf.rect(bx, by, box_sz, box_sz)

            is_on = light_name in lights_on
            if is_on:
                pdf.set_xy(bx, by - 0.5)
                pdf.set_font('Arial', 'B', 8)
                pdf.set_text_color(200, 0, 0)
                pdf.cell(box_sz, box_sz + 0.5, safe_txt("X"), 0, 0, 'C')
                pdf.set_text_color(0, 0, 0)

            # Nombre de la luz
            pdf.set_xy(bx + box_sz + 1.5, cy + 1.4)
            pdf.set_font('Arial', 'B' if is_on else '', 6.0)
            label_max_w = cell_w - (box_sz + 4)
            label = light_name
            while pdf.get_string_width(label) > label_max_w and len(label) > 4:
                label = label[:-2]
            pdf.cell(label_max_w, 3, safe_txt(label), 0, 0, 'L')

        y += rows_n * cell_h

        # Fila del nivel de gasolina
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'B', 6.5)
        pdf.set_fill_color(240, 240, 240)
        pdf.cell(45, 5, safe_txt("  ⛽ Nivel de Gasolina:"), 1, 0, 'L', True)
        pdf.set_font('Arial', 'B', 7)
        fuel_display = fuel_level_val if fuel_level_val else '— Sin especificar —'
        pdf.cell(TW - 45, 5, safe_txt(f"  {fuel_display}"), 1, 1, 'L')

        # Pequeña barra visual con el nivel de gasolina
        if fuel_level_val and fuel_level_val in FUEL_LEVEL_OPTIONS:
            lvl_idx = FUEL_LEVEL_OPTIONS.index(fuel_level_val)
            bar_x = M + 47
            bar_y = y + 1.0
            bar_w = 60
            bar_h = 3.0
            pdf.set_draw_color(120, 120, 120)
            pdf.rect(bar_x, bar_y, bar_w, bar_h)
            fill_w = bar_w * (lvl_idx + 1) / len(FUEL_LEVEL_OPTIONS)
            pdf.set_fill_color(230, 180, 30)
            pdf.rect(bar_x, bar_y, fill_w, bar_h, 'F')

        y += 5

        # ================== ÍNDICE DE MARCAS ==================
        marks = {}
        for it in items or []:
            if isinstance(it, dict):
                marks[it.get('sys', '')] = it.get('mark', '') or ''

        # ================== CHECKLIST (solo ítems marcados) ==================
        y += 2

        gap = 3
        item_h = 4.5
        head_h = 5.5
        mark_w = 17
        box_sz = 3.4

        # Filtrar: solo ítems con marca BUENO o MALO
        all_marked = [
            f for f in RECEPTION_FIELDS_FLAT if marks.get(f) in ('BUENO', 'MALO')
        ]

        # Dividir en dos columnas para el PDF
        half_idx = len(all_marked) // 2 + len(all_marked) % 2
        left_items = all_marked[:half_idx]
        right_items = all_marked[half_idx:]

        columns_to_draw = []
        if left_items:
            columns_to_draw.append((left_items, 0, RECEPTION_CHECKLIST_UI[0]['title']))
        if right_items:
            columns_to_draw.append((right_items, 1, RECEPTION_CHECKLIST_UI[1]['title']))

        if not columns_to_draw:
            # Nada marcado → mensaje compacto
            pdf.set_xy(M, y)
            pdf.set_font('Arial', 'I', 7)
            pdf.set_draw_color(180, 180, 180)
            pdf.cell(
                TW, 8, safe_txt("(Sin ítems marcados en esta revisión)"), 1, 1, 'C'
            )
            y += 9
        else:
            # 1 columna con ítems → ancho completo; 2 columnas → mitad y mitad
            if len(columns_to_draw) == 1:
                col_xs = [M]
                col_ws = [TW]
            else:
                half_w = (TW - gap) / 2.0
                col_xs = [M, M + half_w + gap]
                col_ws = [half_w, half_w]

            y_start_grid = y
            y_max = y
            for idx, (sys_list, cidx, ttl) in enumerate(columns_to_draw):
                col_x = col_xs[idx]
                this_w = col_ws[idx]
                this_sys_w = this_w - 2 * mark_w
                cy = y_start_grid

                # Header de la columna
                pdf.set_xy(col_x, cy)
                pdf.set_font('Arial', 'B', 7)
                pdf.set_fill_color(220, 220, 220)
                pdf.set_draw_color(0, 0, 0)
                pdf.cell(
                    this_w,
                    head_h,
                    safe_txt(f"  REVISIÓN DEL VEHÍCULO - {ttl}"),
                    1,
                    0,
                    'L',
                    True,
                )
                cy += head_h

                # Sub-header BUENO / MALO
                pdf.set_xy(col_x, cy)
                pdf.set_font('Arial', 'B', 6)
                pdf.set_fill_color(240, 240, 240)
                pdf.cell(this_sys_w, 4.5, safe_txt(""), 1, 0, 'L', True)
                pdf.cell(mark_w, 4.5, safe_txt("BUENO / Sí"), 1, 0, 'C', True)
                pdf.cell(mark_w, 4.5, safe_txt("MALO / No"), 1, 0, 'C', True)
                cy += 4.5

                # Filas de ítems (solo los marcados)
                for sys_name in sys_list:
                    mark = marks.get(sys_name, '')
                    cy_box = cy + (item_h - box_sz) / 2.0

                    pdf.set_xy(col_x, cy)
                    pdf.set_font('Arial', '', 6.5)
                    pdf.set_draw_color(150, 150, 150)
                    pdf.cell(this_sys_w, item_h, safe_txt(f"  {sys_name}"), 1, 0, 'L')

                    # Celda BUENO
                    bx_bueno = col_x + this_sys_w
                    pdf.set_draw_color(150, 150, 150)
                    pdf.rect(bx_bueno, cy, mark_w, item_h)
                    cx_b = bx_bueno + (mark_w - box_sz) / 2.0
                    pdf.set_draw_color(60, 60, 60)
                    pdf.rect(cx_b, cy_box, box_sz, box_sz)
                    if mark == 'BUENO':
                        pdf.set_xy(cx_b, cy_box - 0.45)
                        pdf.set_font('Arial', 'B', 8)
                        pdf.cell(box_sz, box_sz + 0.45, safe_txt("V"), 0, 0, 'C')

                    # Celda MALO
                    bx_malo = bx_bueno + mark_w
                    pdf.set_draw_color(150, 150, 150)
                    pdf.rect(bx_malo, cy, mark_w, item_h)
                    cx_m = bx_malo + (mark_w - box_sz) / 2.0
                    pdf.set_draw_color(60, 60, 60)
                    pdf.rect(cx_m, cy_box, box_sz, box_sz)
                    if mark == 'MALO':
                        pdf.set_xy(cx_m, cy_box - 0.45)
                        pdf.set_font('Arial', 'B', 8)
                        pdf.cell(box_sz, box_sz + 0.45, safe_txt("X"), 0, 0, 'C')

                    cy += item_h

                y_max = max(y_max, cy)

            y = y_max

        # ================== RETORNO ==================
        y += 2
        # Caja gris para el título a la izquierda
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'B', 7)
        pdf.set_fill_color(220, 220, 220)
        pdf.set_draw_color(0, 0, 0)
        pdf.cell(
            lbl_w, 5.5, safe_txt(" ¿Retorno por trabajo ya realizado?"), 1, 0, 'L', True
        )

        # Validación robusta (por si viene vacío o nulo)
        es_ret_bool = int(_g('reception_es_retorno', 0) or 0) == 1

        # Coordenadas para la casilla SÍ
        cb_x = M + lbl_w + 6
        cb_y = y + 1.0
        cb_sz = 3.5

        # Dibujar casilla SÍ
        pdf.set_draw_color(60, 60, 60)
        pdf.rect(cb_x, cb_y, cb_sz, cb_sz)
        if es_ret_bool:
            pdf.set_xy(cb_x, cb_y - 0.45)
            pdf.set_font('Arial', 'B', 8)
            pdf.cell(cb_sz, cb_sz + 0.45, safe_txt("X"), 0, 0, 'C')

        # Texto "SÍ"
        pdf.set_xy(cb_x + cb_sz + 2, y)
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(12, 5.5, safe_txt("SÍ"), 0, 0, 'L')

        # Coordenadas para la casilla NO
        cb_x2 = cb_x + cb_sz + 18

        # Dibujar casilla NO
        pdf.rect(cb_x2, cb_y, cb_sz, cb_sz)
        if not es_ret_bool:
            pdf.set_xy(cb_x2, cb_y - 0.45)
            pdf.set_font('Arial', 'B', 8)
            pdf.cell(cb_sz, cb_sz + 0.45, safe_txt("X"), 0, 0, 'C')

        # Texto "NO"
        pdf.set_xy(cb_x2 + cb_sz + 2, y)
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(12, 5.5, safe_txt("NO"), 0, 0, 'L')

        # Borde exterior para cerrar la fila completa
        pdf.rect(M + lbl_w, y, TW - lbl_w, 5.5)

        y += 5.5

        # ================== FIRMAS ==================
        y += 2
        firm_h = 12
        firm_w = TW / 3.0
        labels_firm = ["RECEPCIÓN DE VEHÍCULO", "FIRMA CLIENTE", "FIRMA EJECUTIVO"]
        for i, lbl in enumerate(labels_firm):
            x = M + i * firm_w
            pdf.set_draw_color(0, 0, 0)
            # La primera columna es el TÍTULO de la fila → fondo gris
            if i == 0:
                pdf.set_fill_color(220, 220, 220)
                pdf.rect(x, y, firm_w, firm_h, 'DF')
                pdf.set_xy(x, y + 2)
                pdf.set_font('Arial', 'B', 8)
                pdf.cell(firm_w, 4, safe_txt(lbl), 0, 0, 'C')
            else:
                pdf.rect(x, y, firm_w, firm_h)
                pdf.set_xy(x, y + 1)
                pdf.set_font('Arial', 'B', 6)
                pdf.cell(firm_w, 3, safe_txt(lbl), 0, 0, 'C')
                pdf.line(x + 6, y + firm_h - 3, x + firm_w - 6, y + firm_h - 3)
        y += firm_h + 1.5

        # Segunda fila de firmas
        labels_firm2 = ["ENTREGA DE VEHÍCULO", "RECIBE CONFORME", "FECHA / HORA"]
        for i, lbl in enumerate(labels_firm2):
            x = M + i * firm_w
            pdf.set_draw_color(0, 0, 0)
            # La primera columna es el TÍTULO de la fila → fondo gris
            if i == 0:
                pdf.set_fill_color(220, 220, 220)
                pdf.rect(x, y, firm_w, firm_h, 'DF')
                pdf.set_xy(x, y + 2)
                pdf.set_font('Arial', 'B', 8)
                pdf.cell(firm_w, 4, safe_txt(lbl), 0, 0, 'C')
            else:
                pdf.rect(x, y, firm_w, firm_h)
                pdf.set_xy(x, y + 1)
                pdf.set_font('Arial', 'B', 6)
                pdf.cell(firm_w, 3, safe_txt(lbl), 0, 0, 'C')
                pdf.line(x + 6, y + firm_h - 3, x + firm_w - 6, y + firm_h - 3)
        y += firm_h + 1

        # ================== PIE ==================
        pdf.set_xy(M, y)
        pdf.set_font('Arial', 'I', 5.5)
        pie = (
            "Mediante la presente autorizo los trabajos de reparación y/o instalación de repuestos o materiales necesarios, "
            "las piezas a instalar no tienen garantía a menos que se especifique lo contrario. La empresa no se hace responsable "
            "por pérdidas, daños al vehículo u objetos dejados en él. Así mismo cualquier daño originado por falla de piezas "
            "o repuestos del proveedor (no somos responsables). Autorizo a la empresa a probar el vehículo en la vía pública."
        )
        try:
            pdf.multi_cell(TW, 2.5, safe_txt(pie), 0, 'J')
        except Exception:
            pdf.cell(TW, 2.5, safe_txt(pie[:200]), 0, 0, 'L')

        return pdf_to_bytes(pdf)

    except Exception as e:
        import traceback

        print(f"[PDF recepción] ERROR: {e}")
        traceback.print_exc()
        try:
            pdf = PDF(orientation='L', unit='mm', format='letter')
            pdf.add_page()
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, safe_txt("Error al generar el PDF"), 0, 1, 'C')
            pdf.set_font('Arial', '', 9)
            pdf.multi_cell(0, 5, safe_txt(str(e)), 0, 'L')
            return pdf_to_bytes(pdf)
        except Exception:
            pdf = PDF()
            pdf.add_page()
            return pdf_to_bytes(pdf)


# ==========================================
# PREVIEW PÚBLICO DE RECEPCIONES/REVISIONES
# ==========================================
@app.route('/r/<token>', methods=['GET'])
def public_revision(token):
    """Vista pública (sin login) del preview moderno de una recepción o revisión mecánica."""
    conn = get_db_connection()
    rev = conn.execute(
        "SELECT * FROM revisions WHERE public_token = ?", (token,)
    ).fetchone()
    if not rev:
        conn.close()
        return "Documento no encontrado o link expirado.", 404

    # Detectar qué tipo de documento es
    try:
        reception_items = json.loads(rev['reception_items_json'] or '[]')
    except Exception:
        reception_items = []

    try:
        mechanic_items = json.loads(rev['mechanic_items_json'] or '[]')
    except Exception:
        mechanic_items = []

    try:
        fallas = json.loads(rev['reception_fallas_json'] or '[]')
    except Exception:
        fallas = []

    try:
        legacy_items = json.loads(rev['items_json'] or '[]')
    except Exception:
        legacy_items = []

    has_reception = bool(reception_items)
    has_mechanic = bool(mechanic_items)

    reception_columns_render = []
    if has_reception:
        marks_map = {
            it.get('sys'): it.get('mark', '')
            for it in reception_items
            if isinstance(it, dict)
        }
        for cat in RECEPTION_CHECKLIST_UI:
            col_items = []
            for item in cat['items']:
                item_fields = []
                for f_name in item['fields']:
                    mk = marks_map.get(f_name, '')
                    if mk in ('BUENO', 'MALO'):
                        label = (
                            'Izq'
                            if 'IZQ' in f_name
                            else ('Der' if 'DER' in f_name else '')
                        )
                        item_fields.append({'label': label, 'mark': mk})
                if item_fields:
                    col_items.append({'name': item['name'], 'fields': item_fields})
            if col_items:
                reception_columns_render.append(
                    {
                        'title': cat['title'],
                        'icon': cat['icon'],
                        'items': col_items,
                    }
                )
    # Si no hay ítems marcados, se considera que no hay recepción para mostrar
    has_reception = bool(reception_columns_render)

    # Preparar estructura agrupada para la revisión mecánica (solo marcados)
    mechanic_categories_render = []
    if has_mechanic:
        by_idx = {}
        for it in mechanic_items:
            if not isinstance(it, dict):
                continue
            if 'cat_idx' in it and 'item_idx' in it and 'var_idx' in it:
                try:
                    by_idx[
                        (int(it['cat_idx']), int(it['item_idx']), int(it['var_idx']))
                    ] = it
                except (ValueError, TypeError):
                    pass

        by_str = {}
        for it in mechanic_items:
            if isinstance(it, dict) and 'sys' in it:
                by_str[(it.get('cat', ''), it.get('sys', ''), it.get('var', ''))] = it

        for cat_idx, cat in enumerate(MECHANIC_REVISION_CATEGORIES):
            cat_rows = []
            for item_idx, item in enumerate(cat['items']):
                variants = item.get('variants') or [None]
                variants_render = []
                obs_val = ''
                for var_idx, variant in enumerate(variants):
                    it = (
                        by_idx.get((cat_idx, item_idx, var_idx))
                        or by_str.get((cat['name'], item['system'], variant or ''))
                        or {}
                    )
                    mark = it.get('mark', '') or ''
                    if mark not in ('*', '✓', 'X'):
                        continue  # solo variantes marcadas
                    if it.get('obs'):
                        obs_val = it['obs']
                    variants_render.append(
                        {
                            'label': variant or '',
                            'mark': mark,
                        }
                    )
                if variants_render:
                    cat_rows.append(
                        {
                            'system': item['system'],
                            'variants': variants_render,
                            'obs': obs_val,
                            'has_variants': bool(item.get('variants')),
                        }
                    )
            if cat_rows:
                mechanic_categories_render.append(
                    {
                        'name': cat['name'],
                        'icon': cat['icon'],
                        'components': cat_rows,
                    }
                )
    # Si no hay componentes marcados, no se considera revisión mecánica para mostrar
    has_mechanic = bool(mechanic_categories_render)

    conn.close()
    # ===== Fotos para el preview público =====
    reception_photos = _get_reception_photos(rev['id'])
    # ===== Boceto de daños =====
    try:
        reception_sketch = json.loads(rev['reception_sketch_json'] or '{}')
    except Exception:
        reception_sketch = {}

    # ===== Luces del Tablero =====
    try:
        lights_active = json.loads(rev['reception_lights_json'] or '[]')
    except Exception:
        lights_active = []
    if not isinstance(lights_active, list):
        lights_active = []

    # ===== Nivel de Gasolina =====
    fuel_level = ''
    try:
        fuel_level = (
            (rev['reception_fuel_level'] or '')
            if 'reception_fuel_level' in rev.keys()
            else ''
        )
    except Exception:
        fuel_level = ''

    # Solo mostrar luces ENCENDIDAS (las apagadas se asumen normales)
    lights_on_render = [l for l in DASHBOARD_LIGHTS if l in lights_active]

    # ===== Luces del Tablero + Gasolina =====
    try:
        lights_active = json.loads(rev['reception_lights_json'] or '[]')
    except Exception:
        lights_active = []
    if not isinstance(lights_active, list):
        lights_active = []

    fuel_level = ''
    try:
        fuel_level = (
            (rev['reception_fuel_level'] or '')
            if 'reception_fuel_level' in rev.keys()
            else ''
        )
    except Exception:
        fuel_level = ''

    return render_template(
        'revision_public.html',
        rev=rev,
        has_reception=has_reception,
        has_mechanic=has_mechanic,
        reception_columns=reception_columns_render,
        mechanic_categories=mechanic_categories_render,
        fallas=fallas,
        legacy_items=legacy_items,
        reception_photos=reception_photos,
        reception_sketch=reception_sketch,
        # ⬇️ NUEVOS
        reception_lights_active=lights_active,
        reception_lights_on_render=lights_on_render,
        reception_fuel_level=fuel_level,
        fuel_level_options=FUEL_LEVEL_OPTIONS,
    )


@app.route('/revision/<int:revision_id>/generate_public_token', methods=['POST'])
@block_readonly_reception
def generate_revision_public_token(revision_id):
    """Genera (si no existe) el token público de una recepción/revisión."""
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    conn = get_db_connection()
    rev = conn.execute(
        "SELECT * FROM revisions WHERE id = ?", (revision_id,)
    ).fetchone()
    if rev:
        token = (
            rev['public_token']
            if 'public_token' in rev.keys() and rev['public_token']
            else generate_public_token()
        )
        conn.execute(
            "UPDATE revisions SET public_token = ? WHERE id = ?", (token, revision_id)
        )
        conn.commit()
        flash('✓ Link público listo. Ya puedes compartirlo.', 'success')
    conn.close()
    return redirect(request.referrer or url_for('index', module='📥 Recepciones'))


# ==========================================
# FOTOS DE RECEPCIÓN
# ==========================================
@app.route('/reception/<int:reception_id>/photo/upload', methods=['POST'])
@block_readonly_reception
def reception_photo_upload(reception_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    conn = get_db_connection()
    rev = conn.execute(
        "SELECT id FROM revisions WHERE id = ?", (reception_id,)
    ).fetchone()
    if not rev:
        conn.close()
        return jsonify({'ok': False, 'message': 'Recepción no encontrada'}), 404

    current_count = conn.execute(
        "SELECT COUNT(*) FROM reception_photos WHERE reception_id = ?", (reception_id,)
    ).fetchone()[0]

    files = request.files.getlist('photos')
    valid_files = [f for f in files if f and f.filename]
    if not valid_files:
        conn.close()
        return jsonify({'ok': False, 'message': 'No se seleccionaron archivos'}), 400

    reception_dir = os.path.join(RECEPTION_PHOTOS_DIR, str(reception_id))
    os.makedirs(reception_dir, exist_ok=True)

    uploaded = []
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    actor = session.get('username', '')

    for f in valid_files:
        if current_count + len(uploaded) >= MAX_PHOTOS_PER_RECEPTION:
            break
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_PHOTO_EXTS:
            continue

        # Verificar tamaño (leyendo desde el stream)
        try:
            f.seek(0, 2)
            size = f.tell()
            f.seek(0)
        except Exception:
            size = 0
        if size > MAX_PHOTO_SIZE_MB * 1024 * 1024:
            continue

        ts = int(time.time() * 1000)
        unique_name = f"{ts}_{secrets.token_hex(4)}{ext}"
        full_path = os.path.join(reception_dir, unique_name)

        try:
            f.save(full_path)
        except Exception as e:
            print(f"[photo upload] {e}")
            continue

        rel_path = f"/static/uploads/receptions/{reception_id}/{unique_name}"
        conn.execute(
            """
            INSERT INTO reception_photos (reception_id, filename, original_name, uploaded_by, created_at)
            VALUES (?, ?, ?, ?, ?)
        """,
            (
                reception_id,
                rel_path,
                secure_filename(f.filename) or unique_name,
                actor,
                now_str,
            ),
        )
        uploaded.append(
            {
                'filename': rel_path,
                'original_name': secure_filename(f.filename) or unique_name,
            }
        )

    conn.commit()
    conn.close()

    if uploaded:
        log_action(
            'Fotos subidas',
            f'{len(uploaded)} foto(s) a la recepción #{reception_id}.',
            'Recepción',
            reception_id,
        )

    return jsonify({'ok': True, 'photos': uploaded, 'count': len(uploaded)})


@app.route('/reception/photo/<int:photo_id>/delete', methods=['POST'])
@block_readonly_reception
def reception_photo_delete(photo_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    conn = get_db_connection()
    photo = conn.execute(
        "SELECT * FROM reception_photos WHERE id = ?", (photo_id,)
    ).fetchone()
    if not photo:
        conn.close()
        return jsonify({'ok': False, 'message': 'Foto no encontrada'}), 404

    try:
        full_path = os.path.join(BASE_DIR, photo['filename'].lstrip('/'))
        if os.path.exists(full_path):
            os.remove(full_path)
    except Exception as e:
        print(f"[photo delete] {e}")

    conn.execute("DELETE FROM reception_photos WHERE id = ?", (photo_id,))
    conn.commit()
    conn.close()

    log_action(
        'Foto eliminada',
        f'Foto #{photo_id} de la recepción #{photo["reception_id"]}.',
        'Recepción',
        photo['reception_id'],
    )

    return jsonify({'ok': True})


def _get_reception_photos(reception_id):
    """Devuelve la lista de fotos de una recepción como dicts."""
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT * FROM reception_photos WHERE reception_id = ? ORDER BY id ASC
    """,
        (reception_id,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ==========================================
# INVENTARIO — Compra/Venta
# ==========================================

# Sinónimos de columnas para detectar automáticamente los encabezados
_INV_COL_MAP = {
    'code': ['codigo', 'code', 'cod', 'sku', 'ref', 'referencia', 'codigo producto'],
    'name': [
        'nombre',
        'name',
        'descripcion',
        'producto',
        'articulo',
        'detalle producto',
    ],
    'brand': ['marca', 'brand', 'fabricante'],
    'category': ['categoria', 'category', 'grupo', 'familia', 'linea'],
    'unit': ['unidad', 'unit', 'um', 'medida', 'unidad medida'],
    'cost': ['costo', 'cost', 'precio costo', 'p costo', 'compra', 'precio compra'],
    'price': ['precio', 'price', 'precio venta', 'pvp', 'venta', 'p venta'],
    'stock': [
        'stock',
        'cantidad',
        'qty',
        'existencia',
        'existencias',
        'inventario',
        'disponible',
    ],
    'min_stock': ['minimo', 'min stock', 'stock min', 'stock minimo', 'punto reorden'],
    'location': ['ubicacion', 'location', 'estante', 'pasillo', 'almacen', 'deposito'],
    'supplier': ['proveedor', 'supplier', 'suplidor'],
    'barcode': ['barcode', 'codigo barras', 'ean', 'codigo de barras'],
    'description': [
        'observacion',
        'observaciones',
        'especificacion',
        'especificaciones',
    ],
    'notes': ['nota', 'notas', 'comentario', 'comentarios'],
}


def _normalize_header(h):
    if h is None:
        return ''
    s = str(h).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    return ' '.join(s.split())


def _detect_col_index(headers_norm, key):
    sinos = _INV_COL_MAP.get(key, [])
    for i, h in enumerate(headers_norm):
        if h in sinos:
            return i
    for i, h in enumerate(headers_norm):
        for s in sinos:
            if s and s in h:
                return i
    return -1


def _parse_inventory_file(file_storage):
    """Devuelve (headers, rows) donde rows es lista de listas."""
    filename = (file_storage.filename or '').lower()
    content = file_storage.read()
    file_storage.seek(0)

    if filename.endswith(('.csv', '.txt')):
        text = None
        for enc in ('utf-8-sig', 'utf-8', 'latin-1'):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError('No se pudo decodificar el CSV.')
        sample = text[:4096]
        delim = ';' if sample.count(';') > sample.count(',') else ','
        reader = csv.reader(StringIO(text), delimiter=delim)
        rows = [r for r in reader]
        if not rows:
            return [], []
        return rows[0], rows[1:]

    if filename.endswith(('.xlsx', '.xlsm')):
        if not _HAS_OPENPYXL:
            raise ValueError(
                "Para importar Excel (.xlsx) instala openpyxl: pip install openpyxl. "
                "Alternativamente guarda tu archivo como CSV."
            )
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        if not rows:
            return [], []
        return rows[0], rows[1:]

    if filename.endswith('.xls'):
        raise ValueError(
            "Formato .xls no soportado. Guarda el archivo como .xlsx o .csv."
        )

    raise ValueError("Formato no soportado. Usa .xlsx o .csv.")


@app.route('/inventory/add', methods=['POST'])
def inventory_add():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    code = up(request.form.get('inv_code', ''))
    name = request.form.get('inv_name', '').strip()
    if not code or not name:
        flash('Código y descripción son obligatorios.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    imagen_url = ''
    if 'inv_imagen' in request.files:
        file = request.files['inv_imagen']
        if file and file.filename:
            ext = os.path.splitext(file.filename)[1].lower()
            if ext in {'.jpg', '.jpeg', '.png', '.webp', '.gif'}:
                os.makedirs(
                    os.path.join(BASE_DIR, 'static', 'uploads', 'inventory'),
                    exist_ok=True,
                )
                filename = f"{int(time.time())}_{secrets.token_hex(4)}{ext}"
                file.save(
                    os.path.join(BASE_DIR, 'static', 'uploads', 'inventory', filename)
                )
                imagen_url = f"/static/uploads/inventory/{filename}"

    stock_1 = safe_float(request.form.get('inv_stock_1'), 0.0)
    stock_2 = safe_float(request.form.get('inv_stock_2'), 0.0)
    stock_total = stock_1 + stock_2

    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    current_user = session.get('username', '')

    conn = get_db_connection()
    try:
        conn.execute(
            """
            INSERT INTO inventory (code, name, description, brand, category, unit,
                cost, price, stock, min_stock, location, supplier, barcode, notes,
                is_active, created_at, updated_at, linea, division, montaje, id_casam,
                id_modelo, year_start, year_end, t_anio, amperaje, reserva, ancho, alto,
                largo, fecha_compra, regmod, almacen, reporte, temporal, indice, imagen,
                sede_1, stock_1, sede_2, stock_2)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                code,
                name,
                request.form.get('inv_description', '').strip(),
                up(request.form.get('inv_brand', '')),
                request.form.get('inv_category', '').strip(),
                request.form.get('inv_unit', 'UND').strip() or 'UND',
                safe_float(request.form.get('inv_cost'), 0.0),
                safe_float(request.form.get('inv_price'), 0.0),
                stock_total,
                safe_float(request.form.get('inv_min_stock'), 0.0),
                request.form.get('inv_location', '').strip(),
                request.form.get('inv_supplier', '').strip(),
                '',
                '',
                now_str,
                now_str,
                request.form.get('inv_linea', '').strip(),
                request.form.get('inv_division', '').strip(),
                request.form.get('inv_montaje', '').strip(),
                request.form.get('inv_id_casam', '').strip(),
                request.form.get('inv_id_modelo', '').strip(),
                request.form.get('inv_year_start', '').strip(),
                request.form.get('inv_year_end', '').strip(),
                request.form.get('inv_t_anio', '').strip(),
                request.form.get('inv_amperaje', '').strip(),
                request.form.get('inv_reserva', '').strip(),
                request.form.get('inv_ancho', '').strip(),
                request.form.get('inv_alto', '').strip(),
                request.form.get('inv_largo', '').strip(),
                request.form.get('inv_fecha_compra', '').strip(),
                current_user,
                request.form.get('inv_almacen', '').strip(),
                request.form.get('inv_reporte', '').strip(),
                request.form.get('inv_temporal', '').strip(),
                request.form.get('inv_indice', '').strip(),
                imagen_url,
                request.form.get('inv_sede_1', 'RSG'),
                stock_1,
                request.form.get('inv_sede_2', 'RSV'),
                stock_2,
            ),
        )

        inv_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

        if request.form.get('add_purchase') == '1':
            supplier_id = safe_int(request.form.get('p_supplier_id'), 0)
            invoice_num = request.form.get('p_invoice_num', '').strip()
            credit_days = safe_int(request.form.get('p_credit_days'), 0)
            p_date = request.form.get('p_date', '').strip() or get_ve_time().strftime(
                "%Y-%m-%d"
            )
            cost_val = safe_float(request.form.get('inv_cost'), 0.0)

            if supplier_id > 0:
                due_date = ''
                # FIX BUG #14: Envolver en try/except
                if credit_days > 0:
                    try:
                        base_date = datetime.datetime.strptime(p_date, "%Y-%m-%d")
                        due_date = (base_date + timedelta(days=credit_days)).strftime(
                            "%Y-%m-%d"
                        )
                    except ValueError:
                        print(f"[inventory_add] Fecha inválida para compra: {p_date}")
                        due_date = ''

                cursor_p = conn.execute(
                    """
                    INSERT INTO purchases (date, supplier_id, invoice_num, total_foreign, status, credit_days, due_date, notes, created_by)
                    VALUES (?, ?, ?, ?, 'Aprobado', ?, ?, 'Ingreso directo desde Inventario', ?)
                """,
                    (
                        p_date,
                        supplier_id,
                        invoice_num,
                        cost_val * stock_total,
                        credit_days,
                        due_date,
                        current_user,
                    ),
                )

                purchase_id = cursor_p.lastrowid

                conn.execute(
                    """
                    INSERT INTO purchase_items (purchase_id, inv_id, code, name, qty, cost, total)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        purchase_id,
                        inv_id,
                        code,
                        name,
                        stock_total,
                        cost_val,
                        cost_val * stock_total,
                    ),
                )

                log_action(
                    'Compra Integrada',
                    f'Factura {invoice_num} registrada al crear producto {code}.',
                    'Compras',
                    purchase_id,
                )

        conn.commit()
        log_action('Producto agregado al inventario', f'{code} — {name}', 'Inventario')
        flash(f'✓ Producto "{name}" agregado exitosamente.', 'success')
    except sqlite3.IntegrityError:
        flash(f'El código {code} ya existe.', 'danger')
    except Exception as e:
        conn.rollback()
        flash(f'Error: {e}', 'danger')
    conn.close()
    return redirect(url_for('index', module='📦 Inventario'))


@app.route('/inventory/edit/<int:inv_id>', methods=['POST'])
def inventory_edit(inv_id):
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    if not session.get('logged_in'):
        if is_ajax:
            return jsonify({'ok': False, 'message': 'No autenticado'}), 401
        return redirect(url_for('login'))

    conn = get_db_connection()
    prod = conn.execute("SELECT * FROM inventory WHERE id = ?", (inv_id,)).fetchone()
    if not prod:
        conn.close()
        if is_ajax:
            return jsonify({'ok': False, 'message': 'Producto no encontrado'}), 404
        flash('Producto no encontrado.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    old_stock = float(prod['stock'] or 0)
    old_min = float(prod['min_stock'] or 0)

    imagen_url = prod['imagen'] if 'imagen' in prod.keys() and prod['imagen'] else ''
    if 'inv_imagen' in request.files:
        file = request.files['inv_imagen']
        if file and file.filename:
            ext = os.path.splitext(file.filename)[1].lower()
            if ext in {'.jpg', '.jpeg', '.png', '.webp', '.gif'}:
                os.makedirs(
                    os.path.join(BASE_DIR, 'static', 'uploads', 'inventory'),
                    exist_ok=True,
                )
                filename = f"{int(time.time())}_{secrets.token_hex(4)}{ext}"
                file.save(
                    os.path.join(BASE_DIR, 'static', 'uploads', 'inventory', filename)
                )
                if imagen_url:
                    old_path = os.path.join(BASE_DIR, imagen_url.lstrip('/'))
                    if os.path.exists(old_path):
                        try:
                            os.remove(old_path)
                        except Exception:
                            pass
                imagen_url = f"/static/uploads/inventory/{filename}"

    stock_1 = safe_float(request.form.get('inv_stock_1'), 0.0)
    stock_2 = safe_float(request.form.get('inv_stock_2'), 0.0)
    stock_total = stock_1 + stock_2

    code = up(request.form.get('inv_code', '') or prod['code'])
    name = request.form.get('inv_name', '').strip() or prod['name']
    min_stock = safe_float(request.form.get('inv_min_stock'), 0.0)
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    current_user = session.get('username', '')

    try:
        conn.execute(
            """
            UPDATE inventory SET
                code = ?, name = ?, brand = ?, category = ?, unit = ?,
                cost = ?, price = ?, stock = ?, min_stock = ?, location = ?, supplier = ?,
                updated_at = ?, linea = ?, division = ?, montaje = ?, id_casam = ?,
                id_modelo = ?, year_start = ?, year_end = ?, t_anio = ?, amperaje = ?,
                reserva = ?, ancho = ?, alto = ?, largo = ?, fecha_compra = ?, regmod = ?,
                almacen = ?, reporte = ?, temporal = ?, indice = ?, imagen = ?,
                sede_1 = ?, stock_1 = ?, sede_2 = ?, stock_2 = ?
            WHERE id = ?
        """,
            (
                code,
                name,
                up(request.form.get('inv_brand', '')),
                request.form.get('inv_category', '').strip(),
                request.form.get('inv_unit', 'UND').strip() or 'UND',
                safe_float(request.form.get('inv_cost'), 0.0),
                safe_float(request.form.get('inv_price'), 0.0),
                stock_total,
                min_stock,
                request.form.get('inv_location', '').strip(),
                request.form.get('inv_supplier', '').strip(),
                now_str,
                request.form.get('inv_linea', '').strip(),
                request.form.get('inv_division', '').strip(),
                request.form.get('inv_montaje', '').strip(),
                request.form.get('inv_id_casam', '').strip(),
                request.form.get('inv_id_modelo', '').strip(),
                request.form.get('inv_year_start', '').strip(),
                request.form.get('inv_year_end', '').strip(),
                request.form.get('inv_t_anio', '').strip(),
                request.form.get('inv_amperaje', '').strip(),
                request.form.get('inv_reserva', '').strip(),
                request.form.get('inv_ancho', '').strip(),
                request.form.get('inv_alto', '').strip(),
                request.form.get('inv_largo', '').strip(),
                request.form.get('inv_fecha_compra', '').strip(),
                current_user,
                request.form.get('inv_almacen', '').strip(),
                request.form.get('inv_reporte', '').strip(),
                request.form.get('inv_temporal', '').strip(),
                request.form.get('inv_indice', '').strip(),
                imagen_url,
                request.form.get('inv_sede_1', 'RSG'),
                stock_1,
                request.form.get('inv_sede_2', 'RSV'),
                stock_2,
                inv_id,
            ),
        )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.rollback()
        conn.close()
        if is_ajax:
            return jsonify({'ok': False, 'message': 'El código ya está en uso.'}), 400
        flash(f'El código "{code}" ya existe.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))
    except Exception as e:
        conn.rollback()
        conn.close()
        if is_ajax:
            return jsonify({'ok': False, 'message': str(e)}), 500
        flash(f'Error: {e}', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))
    else:
        conn.close()

    log_action(
        'Edición de inventario',
        f'Producto #{inv_id} ({code} — {name}) actualizado.',
        'Inventario',
        inv_id,
    )

    if is_ajax:
        was_low = old_stock > 0 and old_min > 0 and old_stock <= old_min
        is_low = stock_total > 0 and min_stock > 0 and stock_total <= min_stock
        was_none = old_stock <= 0
        is_none = stock_total <= 0
        return jsonify(
            {
                'ok': True,
                'updated_id': inv_id,
                'product': {
                    'id': inv_id,
                    'code': code,
                    'name': name,
                    'stock': stock_total,
                },
                'deltas': {
                    'stock_delta': stock_total - old_stock,
                    'low_delta': (1 if is_low else 0) - (1 if was_low else 0),
                    'no_delta': (1 if is_none else 0) - (1 if was_none else 0),
                },
            }
        )

    flash('✓ Producto actualizado.', 'success')
    return redirect(url_for('index', module='📦 Inventario'))


@app.route('/inventory/delete/<int:inv_id>', methods=['POST'])
def inventory_delete(inv_id):
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    if not session.get('logged_in'):
        if is_ajax:
            return jsonify({'ok': False, 'message': 'No autenticado'}), 401
        return redirect(url_for('login'))

    conn = get_db_connection()
    prod = conn.execute("SELECT * FROM inventory WHERE id = ?", (inv_id,)).fetchone()
    if not prod:
        conn.close()
        if is_ajax:
            return jsonify({'ok': False, 'message': 'Producto no encontrado'}), 404
        flash('Producto no encontrado.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    code = prod['code']
    name = prod['name']
    stock_val = float(prod['stock'] or 0)
    min_stock_val = float(prod['min_stock'] or 0)

    conn.execute("DELETE FROM inventory WHERE id = ?", (inv_id,))
    conn.commit()
    conn.close()

    log_action(
        'Eliminación de inventario',
        f'Producto {code} — {name} eliminado.',
        'Inventario',
        inv_id,
    )

    if is_ajax:
        return jsonify(
            {
                'ok': True,
                'deleted_id': inv_id,
                'stock': stock_val,
                'min_stock': min_stock_val,
                'was_low': (
                    stock_val > 0 and min_stock_val > 0 and stock_val <= min_stock_val
                ),
                'was_none': (stock_val <= 0),
            }
        )

    flash('Producto eliminado.', 'success')
    return redirect(url_for('index', module='📦 Inventario'))


@app.route('/inventory/clear_all', methods=['POST'])
def inventory_clear_all():
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    if not session.get('logged_in'):
        if is_ajax:
            return jsonify({'ok': False, 'message': 'No autenticado'}), 401
        return redirect(url_for('login'))

    if session.get('role') != 'Administrador':
        msg = 'Solo el Administrador puede vaciar el inventario.'
        if is_ajax:
            return jsonify({'ok': False, 'message': msg}), 403
        flash(msg, 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    conn = get_db_connection()
    total = 0
    backup_name = ''
    try:
        rows = conn.execute("SELECT * FROM inventory ORDER BY id ASC").fetchall()
        total = len(rows)

        backup_dir = os.path.join(BASE_DIR, 'static', 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        backup_name = f"inventario_backup_{get_ve_time().strftime('%Y%m%d_%H%M%S')}.csv"
        backup_path = os.path.join(backup_dir, backup_name)

        try:
            with open(backup_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(
                    [
                        'id',
                        'code',
                        'name',
                        'description',
                        'brand',
                        'category',
                        'unit',
                        'cost',
                        'price',
                        'stock',
                        'min_stock',
                        'location',
                        'supplier',
                        'barcode',
                        'notes',
                        'is_active',
                        'created_at',
                        'updated_at',
                    ]
                )
                for r in rows:
                    writer.writerow(
                        [
                            r['id'],
                            r['code'],
                            r['name'],
                            r['description'] or '',
                            r['brand'] or '',
                            r['category'] or '',
                            r['unit'] or '',
                            r['cost'] or 0,
                            r['price'] or 0,
                            r['stock'] or 0,
                            r['min_stock'] or 0,
                            r['location'] or '',
                            r['supplier'] or '',
                            r['barcode'] or '',
                            r['notes'] or '',
                            r['is_active'] or 1,
                            r['created_at'] or '',
                            r['updated_at'] or '',
                        ]
                    )
        except Exception as e:
            print(f"[clear_all] No se pudo crear backup: {e}")
            backup_name = ''

        conn.execute("DELETE FROM inventory")
        conn.commit()
    except Exception as e:
        conn.rollback()
        import traceback

        traceback.print_exc()
        if is_ajax:
            return jsonify({'ok': False, 'message': str(e)}), 500
        flash(f'Error al vaciar el inventario: {e}', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))
    finally:
        conn.close()

    log_action(
        'Vaciado total del inventario',
        f'Se eliminaron {total} productos del inventario. Backup: {backup_name or "no creado"}.',
        'Inventario',
    )

    if is_ajax:
        return jsonify(
            {
                'ok': True,
                'deleted': total,
                'backup_name': backup_name,
            }
        )

    flash(f'✓ Inventario vaciado. {total} productos eliminados.', 'success')
    return redirect(url_for('index', module='📦 Inventario'))


@app.route('/inventory/import', methods=['POST'])
def inventory_import():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if 'inv_file' not in request.files or not request.files['inv_file'].filename:
        flash('Selecciona un archivo Excel (.xlsx) o CSV.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    file_storage = request.files['inv_file']
    update_existing = request.form.get('inv_update_existing') == '1'
    clear_first = request.form.get('inv_clear_first') == '1'

    try:
        headers, rows = _parse_inventory_file(file_storage)
    except Exception as e:
        flash(f'No se pudo leer el archivo: {e}', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    if not headers:
        flash('El archivo está vacío.', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))

    headers_norm = [_normalize_header(h) for h in headers]
    idx_code = _detect_col_index(headers_norm, 'code')
    idx_name = _detect_col_index(headers_norm, 'name')
    if idx_code < 0 or idx_name < 0:
        flash(
            f'No se detectaron las columnas obligatorias. Se requiere al menos "Código" y "Nombre". '
            f'Encabezados encontrados: {headers}',
            'danger',
        )
        return redirect(url_for('index', module='📦 Inventario'))

    idx_brand = _detect_col_index(headers_norm, 'brand')
    idx_cat = _detect_col_index(headers_norm, 'category')
    idx_unit = _detect_col_index(headers_norm, 'unit')
    idx_cost = _detect_col_index(headers_norm, 'cost')
    idx_price = _detect_col_index(headers_norm, 'price')
    idx_stock = _detect_col_index(headers_norm, 'stock')
    idx_min = _detect_col_index(headers_norm, 'min_stock')
    idx_loc = _detect_col_index(headers_norm, 'location')
    idx_sup = _detect_col_index(headers_norm, 'supplier')
    idx_bar = _detect_col_index(headers_norm, 'barcode')
    idx_desc = _detect_col_index(headers_norm, 'description')
    idx_notes = _detect_col_index(headers_norm, 'notes')

    def _cell(row, idx):
        if idx < 0 or idx >= len(row):
            return ''
        v = row[idx]
        return '' if v is None else str(v).strip()

    def _num(row, idx):
        if idx < 0 or idx >= len(row):
            return 0.0
        v = row[idx]
        if v is None or v == '':
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace('$', '').replace(' ', '')
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '.')
        try:
            return float(s)
        except ValueError:
            return 0.0

    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db_connection()
    inserted = updated = skipped = 0
    try:
        if clear_first:
            conn.execute("DELETE FROM inventory")
            conn.commit()

        for row in rows:
            if not row:
                continue
            code = up(_cell(row, idx_code))
            name = _cell(row, idx_name)
            if not code or not name:
                skipped += 1
                continue

            existing = conn.execute(
                "SELECT id FROM inventory WHERE code = ?", (code,)
            ).fetchone()

            if existing:
                if update_existing:
                    conn.execute(
                        """
                        UPDATE inventory SET name = ?, description = ?, brand = ?, category = ?,
                            unit = ?, cost = ?, price = ?, stock = ?, min_stock = ?,
                            location = ?, supplier = ?, barcode = ?, notes = ?, updated_at = ?
                        WHERE id = ?
                    """,
                        (
                            name,
                            _cell(row, idx_desc),
                            up(_cell(row, idx_brand)),
                            _cell(row, idx_cat),
                            _cell(row, idx_unit) or 'UND',
                            _num(row, idx_cost),
                            _num(row, idx_price),
                            _num(row, idx_stock),
                            _num(row, idx_min),
                            _cell(row, idx_loc),
                            _cell(row, idx_sup),
                            _cell(row, idx_bar),
                            _cell(row, idx_notes),
                            now_str,
                            existing['id'],
                        ),
                    )
                    updated += 1
                else:
                    skipped += 1
            else:
                conn.execute(
                    """
                    INSERT INTO inventory (code, name, description, brand, category, unit,
                        cost, price, stock, min_stock, location, supplier, barcode, notes,
                        is_active, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                """,
                    (
                        code,
                        name,
                        _cell(row, idx_desc),
                        up(_cell(row, idx_brand)),
                        _cell(row, idx_cat),
                        _cell(row, idx_unit) or 'UND',
                        _num(row, idx_cost),
                        _num(row, idx_price),
                        _num(row, idx_stock),
                        _num(row, idx_min),
                        _cell(row, idx_loc),
                        _cell(row, idx_sup),
                        _cell(row, idx_bar),
                        _cell(row, idx_notes),
                        now_str,
                        now_str,
                    ),
                )
                inserted += 1
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        import traceback

        traceback.print_exc()
        flash(f'Error al importar: {e}', 'danger')
        return redirect(url_for('index', module='📦 Inventario'))
    conn.close()

    log_action(
        'Importación de inventario',
        f'Insertados {inserted}, actualizados {updated}, omitidos {skipped}.',
        'Inventario',
    )
    flash(
        f'✓ Importación: {inserted} nuevos · {updated} actualizados · {skipped} omitidos.',
        'success',
    )
    return redirect(url_for('index', module='📦 Inventario'))


@app.route('/api/inventory/search')
def api_inventory_search():
    """Búsqueda AJAX para autocompletar en Nuevo Presupuesto.

    Parámetros:
      q      → término(s) de búsqueda
      mode   → 'all' (por defecto) busca en todos los campos
               'code' busca SÓLO en el campo código

    Multi-palabra e independiente del orden:
      "CIC AMORTIGUADOR"  ≡  "AMORTIGUADOR CIC"  ≡  "amortiguador cic"
    """
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'items': []}), 401

    q = (request.args.get('q') or '').strip()
    mode = (request.args.get('mode') or 'all').strip().lower()
    if mode not in ('all', 'code'):
        mode = 'all'

    limit = safe_int(request.args.get('limit'), 300)
    if limit <= 0:
        limit = 300
    limit = min(limit, 1000)
    if not q:
        return jsonify({'ok': True, 'items': []})

    words = [w.upper() for w in re.split(r'\s+', q) if w.strip()]
    if not words:
        return jsonify({'ok': True, 'items': []})

    where_parts = []
    params_where = []

    if mode == 'code':
        # Sólo por código
        for w in words:
            where_parts.append("UPPER(COALESCE(code,'')) LIKE ?")
            params_where.append(f"%{w}%")
    else:
        # Todos los campos (comportamiento original)
        for w in words:
            like_w = f"%{w}%"
            where_parts.append(
                "(UPPER(COALESCE(code,'')) LIKE ? OR UPPER(COALESCE(name,'')) LIKE ? "
                " OR UPPER(COALESCE(barcode,'')) LIKE ? OR UPPER(COALESCE(brand,'')) LIKE ? "
                " OR UPPER(COALESCE(category,'')) LIKE ? OR UPPER(COALESCE(description,'')) LIKE ? "
                " OR UPPER(COALESCE(supplier,'')) LIKE ? OR UPPER(COALESCE(location,'')) LIKE ?)"
            )
            params_where.extend([like_w] * 8)

    where_sql = " AND ".join(where_parts) if where_parts else "1=1"

    q_up = q.upper()
    like_prefix = f"{q_up}%"

    sql = f"""
        SELECT id, code, name, brand, category, unit, cost, price, stock,
               description, supplier, location, barcode,
               CASE
                   WHEN UPPER(code) = ? THEN 0
                   WHEN UPPER(code) LIKE ? THEN 1
                   WHEN UPPER(barcode) = ? THEN 2
                   WHEN UPPER(name) LIKE ? THEN 3
                   ELSE 5
               END AS rank
        FROM inventory
        WHERE is_active = 1 AND ({where_sql})
        ORDER BY stock DESC, rank ASC, name ASC
        LIMIT ?
    """

    params = [q_up, like_prefix, q_up, like_prefix] + params_where + [limit]

    conn = get_db_connection()
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    items = [
        {
            'id': r['id'],
            'code': r['code'],
            'name': r['name'],
            'brand': r['brand'] or '',
            'category': r['category'] or '',
            'unit': r['unit'] or 'UND',
            'cost': float(r['cost'] or 0),
            'price': float(r['price'] or 0),
            'stock': float(r['stock'] or 0),
            'description': r['description'] or '',
            'supplier': r['supplier'] or '',
            'location': r['location'] or '',
            'barcode': r['barcode'] or '',
        }
        for r in rows
    ]
    return jsonify({'ok': True, 'items': items})


@app.route('/api/vehicle_by_plate')
def api_vehicle_by_plate():
    """Busca un vehículo por placa para autocompletar la recepción."""
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401

    # Normalizar la placa: solo letras y números, sin guiones ni espacios
    plate = re.sub(r'[^A-Z0-9]', '', up(request.args.get('plate', '')))
    if not plate:
        return jsonify({'ok': False, 'message': 'Placa vacía'})

    conn = get_db_connection()
    # Buscamos en la tabla de vehículos y unimos con clientes para traer los datos completos
    row = conn.execute(
        """
        SELECT v.plate, v.brand, v.model, v.year, v.color, v.name,
               c.client_name, c.client_lastname, c.client_doc, c.client_phone
        FROM vehicles v
        LEFT JOIN clients c ON c.client_doc = v.client_doc
        WHERE REPLACE(REPLACE(UPPER(v.plate), '-', ''), ' ', '') = ?
        ORDER BY v.id DESC LIMIT 1
    """,
        (plate,),
    ).fetchone()
    conn.close()

    if row:
        return jsonify(
            {
                'ok': True,
                'vehicle': {
                    'plate': row['plate'],
                    'brand': row['brand'],
                    'model': row['model'],
                    'year': row['year'],
                    'color': row['color'],
                    'name': row['name'],
                    'client_name': row['client_name'],
                    'client_lastname': row['client_lastname'],
                    'client_doc': row['client_doc'],
                    'client_phone': row['client_phone'],
                },
            }
        )
    return jsonify({'ok': False, 'message': 'Vehículo no encontrado'})


@app.route('/inventory/add_stock/<int:inv_id>', methods=['POST'])
def inventory_add_stock(inv_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False, 'message': 'No autenticado'}), 401

    qty = safe_float(request.form.get('qty'), 0.0)
    if qty <= 0:
        return jsonify({'ok': False, 'message': 'Cantidad inválida'}), 400

    conn = get_db_connection()
    prod = conn.execute(
        "SELECT stock, stock_1 FROM inventory WHERE id = ?", (inv_id,)
    ).fetchone()
    if not prod:
        conn.close()
        return jsonify({'ok': False, 'message': 'Producto no encontrado'}), 404

    new_stock = float(prod['stock'] or 0) + qty
    new_stock_1 = float(prod['stock_1'] or 0) + qty
    conn.execute(
        """
        UPDATE inventory
        SET stock = ?, stock_1 = ?, updated_at = ?
        WHERE id = ?
    """,
        (new_stock, new_stock_1, get_ve_time().strftime("%Y-%m-%d %H:%M:%S"), inv_id),
    )
    conn.commit()
    conn.close()

    log_action(
        'Entrada de Inventario',
        f'Se sumaron {qty} unidades al producto #{inv_id}.',
        'Inventario',
        inv_id,
    )
    return jsonify({'ok': True, 'new_stock': new_stock})


# ==========================================
# MÓDULO DE COMPRAS (PROVEEDORES Y FACTURAS)
# ==========================================


@app.route('/supplier/add', methods=['POST'])
def add_supplier():
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401 if is_ajax else redirect(url_for('login'))

    name = request.form.get('sup_name', '').strip()
    rif = request.form.get('sup_rif', '').strip()
    phone = request.form.get('sup_phone', '').strip()
    address = request.form.get('sup_address', '').strip()

    if not name:
        if is_ajax:
            return jsonify(
                {'ok': False, 'message': 'El nombre del proveedor es obligatorio.'}
            )
        flash('El nombre del proveedor es obligatorio.', 'danger')
        return redirect(url_for('index', module='🛒 Compras'))

    conn = get_db_connection()
    count = conn.execute("SELECT COUNT(*) FROM suppliers").fetchone()[0]
    code = f"PROV-{count + 1:04d}"

    try:
        cursor = conn.execute(
            """
            INSERT INTO suppliers (code, name, rif, phone, address, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """,
            (code, name, rif, phone, address, get_ve_time().strftime("%Y-%m-%d %H:%M")),
        )
        conn.commit()
        new_id = cursor.lastrowid

        if is_ajax:
            return jsonify(
                {
                    'ok': True,
                    'supplier': {'id': new_id, 'code': code, 'name': name, 'rif': rif},
                }
            )

        flash(f'Proveedor {name} registrado con código {code}.', 'success')
    except Exception as e:
        if is_ajax:
            return jsonify({'ok': False, 'message': str(e)})
        flash(f'Error: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('index', module='🛒 Compras'))


@app.route('/suppliers/import', methods=['POST'])
def suppliers_import():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if 'sup_file' not in request.files or not request.files['sup_file'].filename:
        flash('Selecciona un archivo Excel (.xlsx).', 'danger')
        return redirect(url_for('index', module='🛒 Compras'))

    file_storage = request.files['sup_file']

    if not file_storage.filename.lower().endswith(('.xlsx', '.xls')):
        flash('El archivo debe ser formato Excel (.xlsx).', 'danger')
        return redirect(url_for('index', module='🛒 Compras'))

    if not _HAS_OPENPYXL:
        flash(
            'La librería openpyxl no está instalada en el servidor. Ejecuta: pip install openpyxl',
            'danger',
        )
        return redirect(url_for('index', module='🛒 Compras'))

    try:
        content = file_storage.read()
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
        ws = wb.active

        agregados = 0
        omitidos = 0

        conn = get_db_connection()

        # Asumimos que la primera fila es el encabezado
        # Los datos empiezan en la fila 2
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue

            # Basado en tu archivo, la col A tiene el Número, la B el Nombre y la C el RIF
            # Ajustamos las posiciones (índices 1 y 2) si el número está en la columna 0
            nombre = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ''
            rif = str(row[2]).strip().upper() if len(row) > 2 and row[2] else ''

            # Si por alguna razón el número no está en la col 0 y el nombre está en la col 0
            if not nombre and row[0]:
                nombre = str(row[0]).strip().upper()

            if not nombre or nombre == 'NONE' or nombre == 'NOMBRES DE PROVEEDORES':
                continue

            # Verificar si ya existe (por nombre o RIF) para no duplicar
            existente = conn.execute(
                "SELECT id FROM suppliers WHERE UPPER(name) = ? OR (rif != '' AND UPPER(rif) = ?)",
                (nombre, rif),
            ).fetchone()

            if existente:
                omitidos += 1
                continue

            # Generar código de proveedor auto-incremental
            count = conn.execute("SELECT COUNT(*) FROM suppliers").fetchone()[0]
            code = f"PROV-{count + 1:04d}"

            # Insertar en la base de datos
            conn.execute(
                """
                INSERT INTO suppliers (code, name, rif, created_at)
                VALUES (?, ?, ?, ?)
            """,
                (code, nombre, rif, get_ve_time().strftime("%Y-%m-%d %H:%M")),
            )

            agregados += 1

        conn.commit()
        conn.close()
        wb.close()

        flash(
            f'✓ Importación exitosa: {agregados} proveedores agregados, {omitidos} omitidos (ya existían).',
            'success',
        )
        log_action(
            'Importación de Proveedores',
            f'Se importaron {agregados} proveedores desde Excel.',
            'Compras',
        )

    except Exception as e:
        import traceback

        traceback.print_exc()
        flash(f'Error al importar el archivo: {e}', 'danger')

    return redirect(url_for('index', module='🛒 Compras'))


@app.route('/purchase/add', methods=['POST'])
def add_purchase():
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401

    supplier_id = safe_int(request.form.get('supplier_id'), 0)
    invoice_num = request.form.get('invoice_num', '').strip()
    doc_type = request.form.get('doc_type', 'Factura Fiscal')
    date_str = request.form.get('purchase_date', '').strip() or get_ve_time().strftime(
        "%Y-%m-%d"
    )
    credit_days = safe_int(request.form.get('credit_days'), 0)
    notes = request.form.get('purchase_notes', '').strip()

    # 👇 LEER EL STATUS DEL BOTÓN PRESIONADO 👇
    new_status = request.form.get('status', 'Pendiente')
    if new_status not in ('Pendiente', 'Aprobado'):
        new_status = 'Pendiente'

    raw_comanda_id = request.form.get('comanda_id', '').strip()
    comanda_id = int(raw_comanda_id) if raw_comanda_id else None

    due_date = ''
    if credit_days > 0:
        base_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        due_date = (base_date + timedelta(days=credit_days)).strftime("%Y-%m-%d")

    items_json = request.form.get('items_json', '[]')
    try:
        items = json.loads(items_json)
    except Exception:
        items = []

    if not supplier_id or not items:
        flash(
            'Debes seleccionar un proveedor y agregar al menos un producto.', 'danger'
        )
        return redirect(url_for('index', module='🛒 Compras'))

    total = sum(float(it.get('total', 0)) for it in items)

    conn = get_db_connection()
    purchase_id = None
    try:
        cursor = conn.cursor()
        # Guardamos siempre como 'Pendiente' primero
        cursor.execute(
            """
            INSERT INTO purchases (date, supplier_id, invoice_num, doc_type, total_foreign, status, credit_days, due_date, notes, created_by, comanda_id)
            VALUES (?, ?, ?, ?, ?, 'Pendiente', ?, ?, ?, ?, ?)
        """,
            (
                date_str,
                supplier_id,
                invoice_num,
                doc_type,
                total,
                credit_days,
                due_date,
                notes,
                session['username'],
                comanda_id,
            ),
        )
        purchase_id = cursor.lastrowid

        for it in items:
            extra_data_str = (
                json.dumps(it.get('extra_data', {})) if it.get('extra_data') else '{}'
            )
            cursor.execute(
                """
                INSERT INTO purchase_items (purchase_id, inv_id, code, name, qty, cost, total, extra_data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    purchase_id,
                    it.get('inv_id'),
                    it.get('code'),
                    it.get('name'),
                    it.get('qty'),
                    it.get('cost'),
                    it.get('total'),
                    extra_data_str,
                ),
            )

        # 👇 SI EL USUARIO ELIGIÓ APROBAR, SUMAMOS EL STOCK AQUÍ 👇
        if new_status == 'Aprobado':
            now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")

            # 1. Sumar Stock
            items_db = conn.execute(
                "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
            ).fetchall()
            for it in items_db:
                code = it['code']
                if code and code != 'N/A':
                    prod = conn.execute(
                        "SELECT id FROM inventory WHERE UPPER(code) = UPPER(?)", (code,)
                    ).fetchone()
                    try:
                        extra = (
                            json.loads(it['extra_data'])
                            if 'extra_data' in it.keys() and it['extra_data']
                            else {}
                        )
                    except Exception:
                        extra = {}

                    if prod:
                        conn.execute(
                            """
                            UPDATE inventory
                            SET stock = COALESCE(stock, 0) + ?, cost = ?, stock_1 = COALESCE(stock_1, 0) + ?
                            WHERE id = ?
                        """,
                            (it['qty'], it['cost'], it['qty'], prod['id']),
                        )
                    else:
                        brand = extra.get('brand', '')
                        category = extra.get('category', '')
                        linea = extra.get('linea', '')
                        montaje = extra.get('montaje', '')
                        year_start = extra.get('year_start', '')
                        year_end = extra.get('year_end', '')
                        amperaje = extra.get('amperaje', '')
                        ancho = extra.get('ancho', '')
                        alto = extra.get('alto', '')
                        largo = extra.get('largo', '')
                        almacen = extra.get('almacen', '')
                        indice = extra.get('indice', '')
                        sede_1 = extra.get('sede_1', 'RSG')
                        stock_1 = float(extra.get('stock_1', it['qty']))
                        sede_2 = extra.get('sede_2', 'RSV')
                        stock_2 = float(extra.get('stock_2', 0))
                        suggested_price = float(it['cost']) * 1.30

                        conn.execute(
                            """
                            INSERT INTO inventory (code, name, brand, category, linea, montaje,
                            year_start, year_end, amperaje, ancho, alto, largo, almacen, indice,
                            cost, price, stock, created_at, updated_at, sede_1, stock_1, sede_2, stock_2)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                            (
                                code.upper(),
                                it['name'],
                                brand.upper(),
                                category.upper(),
                                linea.upper(),
                                montaje.upper(),
                                year_start,
                                year_end,
                                amperaje,
                                ancho,
                                alto,
                                largo,
                                almacen.upper(),
                                indice,
                                it['cost'],
                                suggested_price,
                                it['qty'],
                                now_str,
                                now_str,
                                sede_1,
                                stock_1,
                                sede_2,
                                stock_2,
                            ),
                        )

            # 2. Cambiar status a Aprobado
            conn.execute(
                "UPDATE purchases SET status = 'Aprobado' WHERE id = ?", (purchase_id,)
            )

            if comanda_id:
                conn.execute(
                    "UPDATE comandas SET status = 'Completada', purchase_id = ?, updated_at = ? WHERE id = ?",
                    (purchase_id, now_str, comanda_id),
                )
                budget = conn.execute(
                    "SELECT budget_id FROM comandas WHERE id = ?", (comanda_id,)
                ).fetchone()
                if budget:
                    # FIX BUG #13: Marcar los ítems como En Stock en la comanda
                    conn.execute(
                        "UPDATE comanda_items SET status = 'En Stock' WHERE comanda_id = ? AND status = 'Faltante'",
                        (comanda_id,),
                    )

        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f'Error al registrar compra: {e}', 'danger')
        return redirect(url_for('index', module='🛒 Compras'))
    finally:
        conn.close()

    log_action(
        'Compra Registrada',
        f'{doc_type} {invoice_num} por ${total:.2f} (Estado: {new_status})',
        'Compras',
        purchase_id,
    )
    flash('✓ Factura de compra registrada exitosamente.', 'success')
    return redirect(url_for('index', module='🛒 Compras'))


@app.route('/purchase/status/<int:purchase_id>/<status>', methods=['POST'])
def change_purchase_status(purchase_id, status):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    if status not in ('Pendiente', 'Aprobado', 'Pagado', 'Anulado'):
        flash('Estado inválido.', 'danger')
        return redirect(url_for('index', module='🛒 Compras'))

    conn = get_db_connection()
    p = conn.execute(
        "SELECT status, comanda_id FROM purchases WHERE id = ?", (purchase_id,)
    ).fetchone()
    old_status = p['status'] if p else ''
    comanda_id = p['comanda_id'] if p else None

    conn.execute("UPDATE purchases SET status = ? WHERE id = ?", (status, purchase_id))

    if old_status != 'Aprobado' and status == 'Aprobado':
        items = conn.execute(
            "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
        ).fetchall()
        now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")

        for it in items:
            code = it['code']
            if code and code != 'N/A':
                prod = conn.execute(
                    "SELECT id FROM inventory WHERE UPPER(code) = UPPER(?)", (code,)
                ).fetchone()

                try:
                    extra = (
                        json.loads(it['extra_data'])
                        if 'extra_data' in it.keys() and it['extra_data']
                        else {}
                    )
                except Exception:
                    extra = {}

                if prod:
                    conn.execute(
                        """
                        UPDATE inventory
                        SET stock = COALESCE(stock, 0) + ?, cost = ?, stock_1 = COALESCE(stock_1, 0) + ?
                        WHERE id = ?
                    """,
                        (it['qty'], it['cost'], it['qty'], prod['id']),
                    )
                else:
                    brand = extra.get('brand', '')
                    category = extra.get('category', '')
                    linea = extra.get('linea', '')
                    montaje = extra.get('montaje', '')
                    year_start = extra.get('year_start', '')
                    year_end = extra.get('year_end', '')
                    amperaje = extra.get('amperaje', '')
                    ancho = extra.get('ancho', '')
                    alto = extra.get('alto', '')
                    largo = extra.get('largo', '')
                    almacen = extra.get('almacen', '')
                    indice = extra.get('indice', '')
                    sede_1 = extra.get('sede_1', 'RSG')
                    stock_1 = float(extra.get('stock_1', it['qty']))
                    sede_2 = extra.get('sede_2', 'RSV')
                    stock_2 = float(extra.get('stock_2', 0))
                    suggested_price = float(it['cost']) * 1.30

                    conn.execute(
                        """
                        INSERT INTO inventory (code, name, brand, category, linea, montaje,
                        year_start, year_end, amperaje, ancho, alto, largo, almacen, indice,
                        cost, price, stock, created_at, updated_at, sede_1, stock_1, sede_2, stock_2)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            code.upper(),
                            it['name'],
                            brand.upper(),
                            category.upper(),
                            linea.upper(),
                            montaje.upper(),
                            year_start,
                            year_end,
                            amperaje,
                            ancho,
                            alto,
                            largo,
                            almacen.upper(),
                            indice,
                            it['cost'],
                            suggested_price,
                            it['qty'],
                            now_str,
                            now_str,
                            sede_1,
                            stock_1,
                            sede_2,
                            stock_2,
                        ),
                    )

        if comanda_id:
            conn.execute(
                "UPDATE comandas SET status = 'Completada', purchase_id = ?, updated_at = ? WHERE id = ?",
                (purchase_id, now_str, comanda_id),
            )
            budget = conn.execute(
                "SELECT budget_id FROM comandas WHERE id = ?", (comanda_id,)
            ).fetchone()
            if budget:
                # FIX BUG #13: Marcar los ítems como En Stock en la comanda
                conn.execute(
                    "UPDATE comanda_items SET status = 'En Stock' WHERE comanda_id = ? AND status = 'Faltante'",
                    (comanda_id,),
                )

        log_action(
            'Compra Aprobada',
            f'Factura #{purchase_id} aprobada. Stock actualizado/creado.',
            'Compras',
            purchase_id,
        )

    elif old_status == 'Aprobado' and status == 'Anulado':
        items = conn.execute(
            "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
        ).fetchall()
        for it in items:
            if it['code'] and it['code'] != 'N/A':
                conn.execute(
                    "UPDATE inventory SET stock = COALESCE(stock, 0) - ? WHERE UPPER(code) = UPPER(?)",
                    (it['qty'], it['code']),
                )
        log_action(
            'Compra Anulada',
            f'Factura #{purchase_id} anulada. Stock restado.',
            'Compras',
            purchase_id,
        )

    conn.commit()
    conn.close()
    flash(f'Factura marcada como {status}.', 'success')
    return redirect(url_for('index', module='🛒 Compras'))


@app.route('/purchase/pdf/<int:purchase_id>')
def download_purchase_pdf(purchase_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = get_db_connection()
    # Usamos LEFT JOIN por si el proveedor fue borrado
    p = conn.execute(
        """
        SELECT p.*, s.name as sup_name, s.rif as sup_rif, s.phone as sup_phone, s.address as sup_address
        FROM purchases p
        LEFT JOIN suppliers s ON p.supplier_id = s.id
        WHERE p.id = ?
    """,
        (purchase_id,),
    ).fetchone()

    items = conn.execute(
        "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
    ).fetchall()
    conn.close()

    if not p:
        return "Factura no encontrada", 404

    try:
        pdf = PDF()
        pdf.add_page()

        # --- ENCABEZADO ---
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(
            0, 10, safe_txt("RUEDAS Y SUSPENSIONES - CONTROL DE COMPRAS"), 0, 1, 'C'
        )

        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 5, safe_txt(f"FACTURA DE COMPRA #{p['id']:06d}"), 0, 1, 'R')

        pdf.set_font('Arial', '', 9)
        pdf.cell(
            0,
            5,
            safe_txt(f"Fecha: {p['date'] or '—'}  |  Estatus: {p['status'] or '—'}"),
            0,
            1,
            'R',
        )

        pdf.ln(5)

        # --- PROVEEDOR ---
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 5, safe_txt("DATOS DEL PROVEEDOR"), 0, 1, 'L')
        pdf.set_font('Arial', '', 9)
        pdf.cell(0, 5, safe_txt(f"Razón Social: {p['sup_name'] or '—'}"), 0, 1)
        pdf.cell(
            0,
            5,
            safe_txt(
                f"RIF: {p['sup_rif'] or '—'}  |  Teléfono: {p['sup_phone'] or '—'}"
            ),
            0,
            1,
        )
        pdf.cell(0, 5, safe_txt(f"Dirección: {p['sup_address'] or '—'}"), 0, 1)

        pdf.ln(5)

        # --- DATOS FACTURA ---
        pdf.cell(
            0,
            5,
            safe_txt(
                f"N° Factura Proveedor: {p['invoice_num'] or '—'}  |  Días de Crédito: {p['credit_days'] or 0}"
            ),
            0,
            1,
        )
        if p['due_date']:
            pdf.cell(0, 5, safe_txt(f"Fecha de Vencimiento: {p['due_date']}"), 0, 1)

        pdf.ln(5)

        # --- TABLA DE ÍTEMS ---
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(25, 7, safe_txt('Código'), 1)
        pdf.cell(90, 7, safe_txt('Descripción'), 1)
        pdf.cell(20, 7, safe_txt('Cant.'), 1, 0, 'C')
        pdf.cell(30, 7, safe_txt('Costo Unit.'), 1, 0, 'R')
        pdf.cell(25, 7, safe_txt('Total'), 1, 1, 'R')

        pdf.set_font('Arial', '', 8)
        for it in items:
            # Validaciones anti-crash: Si el campo es None, lo convierte a '' o 0.0
            c_code = str(it['code'] or '')[:15]
            c_name = str(it['name'] or '')[:50]
            c_qty = float(it['qty'] or 0)
            c_cost = float(it['cost'] or 0)
            c_total = float(it['total'] or 0)

            pdf.cell(25, 6, safe_txt(c_code), 1)
            pdf.cell(90, 6, safe_txt(c_name), 1)
            pdf.cell(20, 6, str(c_qty), 1, 0, 'C')
            pdf.cell(30, 6, f"{c_cost:.2f}", 1, 0, 'R')
            pdf.cell(25, 6, f"{c_total:.2f}", 1, 1, 'R')

        pdf.ln(4)
        pdf.set_font('Arial', 'B', 10)
        total_val = float(p['total_foreign'] or 0)
        pdf.cell(0, 6, safe_txt(f"TOTAL FACTURA: ${total_val:.2f}"), 0, 1, 'R')

        pdf_bytes = pdf_to_bytes(pdf)
        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Compra_{p['id']:06d}.pdf",
        )

    except Exception as e:
        # Si algo falla, en vez de crashear toda la app, muestra un mensaje
        import traceback

        traceback.print_exc()
        return f"Error interno al generar PDF: {e}", 500


@app.route('/api/suppliers/list')
def api_suppliers_list():
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM suppliers ORDER BY name ASC").fetchall()
    conn.close()
    return jsonify({'ok': True, 'suppliers': [dict(r) for r in rows]})


@app.route('/api/comanda/<int:comanda_id>')
def api_get_comanda(comanda_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    comanda = conn.execute(
        "SELECT * FROM comandas WHERE id = ?", (comanda_id,)
    ).fetchone()
    if not comanda:
        conn.close()
        return jsonify({'ok': False, 'message': 'Comanda no encontrada'}), 404

    items = conn.execute(
        "SELECT * FROM comanda_items WHERE comanda_id = ?", (comanda_id,)
    ).fetchall()
    items_list = []
    for it in items:
        prod = conn.execute(
            "SELECT cost, price, id FROM inventory WHERE UPPER(code) = UPPER(?)",
            (it['inv_code'],),
        ).fetchone()
        cost = float(prod['cost']) if prod else 0.0
        inv_id = prod['id'] if prod else None
        items_list.append(
            {
                'inv_id': inv_id,
                'code': it['inv_code'],
                'name': it['name'],
                'qty': it['qty'],
                'cost': cost,
                'total': cost * it['qty'],
            }
        )
    conn.close()
    return jsonify({'ok': True, 'comanda': dict(comanda), 'items': items_list})


@app.route('/api/comanda/add_oferta', methods=['POST'])
def api_comanda_add_oferta():
    if not session.get('logged_in'):
        return jsonify({'ok': False})
    item_id = safe_int(request.form.get('item_id'))
    supplier_id = safe_int(request.form.get('supplier_id'))
    cost = safe_float(request.form.get('offered_cost'))
    notes = request.form.get('offered_notes', '').strip()
    username = session.get('username', '')

    conn = get_db_connection()
    # Guardar oferta
    conn.execute(
        """
        INSERT INTO comanda_ofertas (item_id, supplier_id, offered_cost, offered_notes, username, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """,
        (
            item_id,
            supplier_id,
            cost,
            notes,
            username,
            get_ve_time().strftime("%Y-%m-%d %H:%M"),
        ),
    )

    # Actualizar estado del item
    conn.execute(
        "UPDATE comanda_items SET status = 'Ofertas Recibidas' WHERE id = ?", (item_id,)
    )
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/comanda/select_oferta', methods=['POST'])
def api_comanda_select_oferta():
    if not session.get('logged_in'):
        return jsonify({'ok': False})
    oferta_id = safe_int(request.form.get('oferta_id'))
    item_id = safe_int(request.form.get('item_id'))

    conn = get_db_connection()
    oferta = conn.execute(
        "SELECT * FROM comanda_ofertas WHERE id = ?", (oferta_id,)
    ).fetchone()

    if oferta:
        # Marcar esta oferta como seleccionada y las demás como rechazadas
        conn.execute(
            "UPDATE comanda_ofertas SET status = 'Seleccionada' WHERE id = ?",
            (oferta_id,),
        )
        conn.execute(
            "UPDATE comanda_ofertas SET status = 'Descartada' WHERE item_id = ? AND id != ?",
            (item_id, oferta_id),
        )

        # Pasar los datos a la tabla principal del item para que el módulo de Compras los lea
        conn.execute(
            """
            UPDATE comanda_items
            SET status = 'Listo para Despacho', supplier_id = ?, offered_cost = ?, offered_notes = ?, approved_by_seller = 1
            WHERE id = ?
        """,
            (
                oferta['supplier_id'],
                oferta['offered_cost'],
                oferta['offered_notes'],
                item_id,
            ),
        )
        conn.commit()

    conn.close()
    return jsonify({'ok': True})


@app.route('/api/comanda/rechazar_item', methods=['POST'])
def api_comanda_rechazar_item():
    if not session.get('logged_in'):
        return jsonify({'ok': False})
    item_id = safe_int(request.form.get('item_id'))
    conn = get_db_connection()
    conn.execute(
        "UPDATE comanda_items SET status = 'Rechazado' WHERE id = ?", (item_id,)
    )
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/comanda/deliver_item', methods=['POST'])
def api_comanda_deliver_item():
    if not session.get('logged_in'):
        return jsonify({'ok': False})
    item_id = safe_int(request.form.get('item_id'))
    signature = request.form.get('signature', '')
    actor = session.get('username', '')
    now_str = get_ve_time().strftime("%Y-%m-%d %H:%M")

    if not signature:
        return jsonify({'ok': False, 'message': 'La firma es obligatoria.'})

    conn = get_db_connection()
    try:
        conn.execute(
            """
            UPDATE comanda_items
            SET status = 'Entregado', delivered_at = ?, delivered_by = ?, signature_data = ?
            WHERE id = ?
        """,
            (now_str, actor, signature, item_id),
        )
        conn.commit()
        return jsonify({'ok': True})
    except Exception as e:
        conn.rollback()
        return jsonify({'ok': False, 'message': str(e)})
    finally:
        conn.close()


@app.route('/api/comanda_item/<int:item_id>')
def api_get_comanda_item(item_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False})
    conn = get_db_connection()
    item = conn.execute(
        """
        SELECT ci.*, s.name as supplier_name
        FROM comanda_items ci
        LEFT JOIN suppliers s ON ci.supplier_id = s.id
        WHERE ci.id = ?
    """,
        (item_id,),
    ).fetchone()
    conn.close()
    if item:
        return jsonify({'ok': True, 'item': dict(item)})
    return jsonify({'ok': False})


@app.route('/api/purchase/next_order_num')
def api_next_order_num():
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401

    # Obtenemos el año actual
    year = datetime.datetime.now().year
    prefix = f"ORD-{year}-"

    conn = get_db_connection()
    # Buscamos el último número generado este año
    row = conn.execute(
        "SELECT invoice_num FROM purchases WHERE invoice_num LIKE ? ORDER BY id DESC LIMIT 1",
        (prefix + '%',),
    ).fetchone()
    conn.close()

    if row and row['invoice_num']:
        try:
            # Extraemos la secuencia (ej: de ORD-2026-0005 extrae 0005)
            last_seq = int(row['invoice_num'].replace(prefix, ''))
            next_seq = last_seq + 1
        except ValueError:
            next_seq = 1
    else:
        next_seq = 1

    # Formateamos a 4 dígitos (ej: ORD-2026-0001)
    next_num = f"{prefix}{next_seq:04d}"

    return jsonify({'ok': True, 'next_num': next_num})


@app.route('/api/purchase/<int:purchase_id>')
def api_get_purchase(purchase_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    p = conn.execute("SELECT * FROM purchases WHERE id = ?", (purchase_id,)).fetchone()
    items = conn.execute(
        "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
    ).fetchall()
    conn.close()
    if not p:
        return jsonify({'ok': False, 'message': 'No encontrada'})

    return jsonify({'ok': True, 'purchase': dict(p), 'items': [dict(i) for i in items]})


@app.route('/purchase/delete/<int:purchase_id>', methods=['POST'])
def delete_purchase(purchase_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()
    p = conn.execute("SELECT * FROM purchases WHERE id = ?", (purchase_id,)).fetchone()

    if not p:
        conn.close()
        return jsonify({'ok': False, 'message': 'No encontrada'}), 404

    error_msg = None
    try:
        if p['status'] == 'Aprobado':
            items = conn.execute(
                "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
            ).fetchall()
            for it in items:
                if it['code'] and it['code'] != 'N/A':
                    conn.execute(
                        "UPDATE inventory SET stock = COALESCE(stock, 0) - ? WHERE UPPER(code) = UPPER(?)",
                        (it['qty'], it['code']),
                    )

        conn.execute("DELETE FROM purchase_items WHERE purchase_id = ?", (purchase_id,))
        conn.execute("DELETE FROM purchases WHERE id = ?", (purchase_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        error_msg = str(e)
    finally:
        conn.close()

    if error_msg:
        return jsonify({'ok': False, 'message': error_msg}), 500

    log_action(
        'Compra Eliminada',
        f'Se eliminó la compra #{purchase_id}',
        'Compras',
        purchase_id,
    )
    return jsonify({'ok': True})


@app.route('/purchase/edit/<int:purchase_id>', methods=['POST'])
def edit_purchase(purchase_id):
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401
    conn = get_db_connection()

    supplier_id = safe_int(request.form.get('supplier_id'), 0)
    invoice_num = request.form.get('invoice_num', '').strip()
    doc_type = request.form.get('doc_type', 'Factura Fiscal')
    date_str = request.form.get('purchase_date', '').strip()
    credit_days = safe_int(request.form.get('credit_days'), 0)
    notes = request.form.get('purchase_notes', '').strip()
    items_json = request.form.get('items_json', '[]')
    new_status = request.form.get('status', 'Pendiente')

    try:
        items = json.loads(items_json)
    except:
        items = []

    total = sum(float(it.get('total', 0)) for it in items)

    due_date = ''
    if credit_days > 0:
        try:
            base_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            due_date = (base_date + timedelta(days=credit_days)).strftime("%Y-%m-%d")
        except ValueError:
            pass

    old = conn.execute(
        "SELECT status FROM purchases WHERE id = ?", (purchase_id,)
    ).fetchone()
    old_status = old['status'] if old else 'Pendiente'

    try:
        if old_status == 'Aprobado':
            old_items = conn.execute(
                "SELECT * FROM purchase_items WHERE purchase_id = ?", (purchase_id,)
            ).fetchall()
            for it in old_items:
                if it['code'] and it['code'] != 'N/A':
                    conn.execute(
                        "UPDATE inventory SET stock = COALESCE(stock, 0) - ? WHERE UPPER(code) = UPPER(?)",
                        (it['qty'], it['code']),
                    )

        conn.execute(
            """
            UPDATE purchases SET date=?, supplier_id=?, invoice_num=?, doc_type=?, total_foreign=?, credit_days=?, due_date=?, notes=?, status=?
            WHERE id=?
        """,
            (
                date_str,
                supplier_id,
                invoice_num,
                doc_type,
                total,
                credit_days,
                due_date,
                notes,
                new_status,
                purchase_id,
            ),
        )

        conn.execute("DELETE FROM purchase_items WHERE purchase_id = ?", (purchase_id,))
        for it in items:
            extra_data_str = (
                json.dumps(it.get('extra_data', {})) if it.get('extra_data') else '{}'
            )
            conn.execute(
                """
                INSERT INTO purchase_items (purchase_id, inv_id, code, name, qty, cost, total, extra_data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    purchase_id,
                    it.get('inv_id'),
                    it.get('code'),
                    it.get('name'),
                    it.get('qty'),
                    it.get('cost'),
                    it.get('total'),
                    extra_data_str,
                ),
            )

        if new_status == 'Aprobado':
            for it in items:
                code = it.get('code')
                if code and code != 'N/A':
                    prod = conn.execute(
                        "SELECT id FROM inventory WHERE UPPER(code) = UPPER(?)", (code,)
                    ).fetchone()
                    if prod:
                        conn.execute(
                            "UPDATE inventory SET stock = COALESCE(stock, 0) + ?, cost = ? WHERE id = ?",
                            (it['qty'], it['cost'], prod['id']),
                        )
                    else:
                        extra = it.get('extra_data', {})
                        if extra:
                            brand = extra.get('brand', '')
                            category = extra.get('category', '')
                            linea = extra.get('linea', '')
                            montaje = extra.get('montaje', '')
                            year_start = extra.get('year_start', '')
                            year_end = extra.get('year_end', '')
                            amperaje = extra.get('amperaje', '')
                            ancho = extra.get('ancho', '')
                            alto = extra.get('alto', '')
                            largo = extra.get('largo', '')
                            almacen = extra.get('almacen', '')
                            indice = extra.get('indice', '')
                            sede_1 = extra.get('sede_1', 'RSG')
                            stock_1 = float(extra.get('stock_1', it['qty']))
                            sede_2 = extra.get('sede_2', 'RSV')
                            stock_2 = float(extra.get('stock_2', 0))
                            suggested_price = float(it['cost']) * 1.30
                            now_str = get_ve_time().strftime("%Y-%m-%d %H:%M:%S")

                            conn.execute(
                                """
                                INSERT INTO inventory (code, name, brand, category, linea, montaje,
                                year_start, year_end, amperaje, ancho, alto, largo, almacen, indice,
                                cost, price, stock, created_at, updated_at, sede_1, stock_1, sede_2, stock_2)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                                (
                                    code.upper(),
                                    it['name'],
                                    brand.upper(),
                                    category.upper(),
                                    linea.upper(),
                                    montaje.upper(),
                                    year_start,
                                    year_end,
                                    amperaje,
                                    ancho,
                                    alto,
                                    largo,
                                    almacen.upper(),
                                    indice,
                                    it['cost'],
                                    suggested_price,
                                    it['qty'],
                                    now_str,
                                    now_str,
                                    sede_1,
                                    stock_1,
                                    sede_2,
                                    stock_2,
                                ),
                            )

        conn.commit()
        return jsonify({'ok': True})
    except Exception as e:
        conn.rollback()
        return jsonify({'ok': False, 'message': str(e)}), 500
    finally:
        conn.close()


def generate_purchase_search_pdf(purchases, filters_desc):
    """Genera un PDF con el reporte de compras filtradas."""
    try:
        pdf = PDF(orientation='L', unit='mm', format='letter')
        pdf.add_page()

        # Encabezado
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(
            0, 10, safe_txt("RUEDAS Y SUSPENSIONES - REPORTE DE COMPRAS"), 0, 1, 'C'
        )

        pdf.set_font('Arial', 'I', 9)
        pdf.cell(0, 5, safe_txt(f"Filtros aplicados: {filters_desc}"), 0, 1, 'C')
        pdf.ln(5)

        # Tabla Header
        pdf.set_font('Arial', 'B', 8)
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(20, 7, safe_txt('ID'), 1, 0, 'C', True)
        pdf.cell(30, 7, safe_txt('Fecha'), 1, 0, 'C', True)
        pdf.cell(70, 7, safe_txt('Proveedor'), 1, 0, 'L', True)
        pdf.cell(40, 7, safe_txt('N° Factura'), 1, 0, 'L', True)
        pdf.cell(30, 7, safe_txt('Estatus'), 1, 0, 'C', True)
        pdf.cell(35, 7, safe_txt('Total'), 1, 1, 'R', True)

        # Filas
        pdf.set_font('Arial', '', 8)
        total_global = 0
        for p in purchases:
            total_val = float(p['total_foreign'] or 0)
            total_global += total_val

            pdf.cell(20, 6, safe_txt(str(p['id'])), 1, 0, 'C')
            pdf.cell(30, 6, safe_txt(p['date']), 1, 0, 'C')
            pdf.cell(
                70,
                6,
                safe_txt(str(p['supplier_name'] or 'Desconocido')[:40]),
                1,
                0,
                'L',
            )
            pdf.cell(40, 6, safe_txt(str(p['invoice_num'] or '—')[:20]), 1, 0, 'L')
            pdf.cell(30, 6, safe_txt(p['status']), 1, 0, 'C')
            pdf.cell(35, 6, safe_txt(f"${total_val:.2f}"), 1, 1, 'R')

        # Total General
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(190, 8, safe_txt("TOTAL GENERAL:"), 0, 0, 'R')
        pdf.cell(35, 8, safe_txt(f"${total_global:.2f}"), 1, 1, 'R')

        return pdf_to_bytes(pdf)
    except Exception as e:
        print(f"[PDF Compras] Error: {e}")
        return None


@app.route('/api/purchase/search', methods=['GET'])
def api_search_purchases():
    if not session.get('logged_in'):
        return jsonify({'ok': False}), 401

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    supplier_id = request.args.get('supplier_id', '')
    status = request.args.get('status', '')
    invoice_num = request.args.get('invoice_num', '').strip()
    doc_type = request.args.get('doc_type', '')  # <-- NUEVO

    query = """
        SELECT p.*, s.name as supplier_name
        FROM purchases p
        LEFT JOIN suppliers s ON p.supplier_id = s.id
        WHERE 1=1
    """
    params = []

    if date_from:
        query += " AND p.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND p.date <= ?"
        params.append(date_to)
    if supplier_id:
        query += " AND p.supplier_id = ?"
        params.append(supplier_id)
    if status:
        query += " AND p.status = ?"
        params.append(status)
    if invoice_num:
        query += " AND p.invoice_num LIKE ?"
        params.append(f"%{invoice_num}%")
    if doc_type:  # <-- NUEVO
        query += " AND p.doc_type = ?"
        params.append(doc_type)

    query += " ORDER BY p.id DESC"

    conn = get_db_connection()
    rows = conn.execute(query, params).fetchall()
    conn.close()

    return jsonify({'ok': True, 'items': [dict(r) for r in rows]})


@app.route('/purchase/report/pdf', methods=['GET'])
def download_purchase_report_pdf():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    supplier_id = request.args.get('supplier_id', '')
    status = request.args.get('status', '')
    invoice_num = request.args.get('invoice_num', '').strip()
    doc_type = request.args.get('doc_type', '').strip()  # <-- NUEVA LÍNEA

    query = """
        SELECT p.*, s.name as supplier_name
        FROM purchases p
        LEFT JOIN suppliers s ON p.supplier_id = s.id
        WHERE 1=1
    """
    params = []
    desc_parts = []

    if date_from:
        query += " AND p.date >= ?"
        params.append(date_from)
        desc_parts.append(f"Desde: {date_from}")
    if date_to:
        query += " AND p.date <= ?"
        params.append(date_to)
        desc_parts.append(f"Hasta: {date_to}")
    if supplier_id:
        query += " AND p.supplier_id = ?"
        params.append(supplier_id)
        desc_parts.append("Proveedor filtrado")
    if status:
        query += " AND p.status = ?"
        params.append(status)
        desc_parts.append(f"Estatus: {status}")
    if invoice_num:
        query += " AND p.invoice_num LIKE ?"
        params.append(f"%{invoice_num}%")
        desc_parts.append(f"Factura: {invoice_num}")
    if doc_type:  # <-- NUEVO FILTRO
        query += " AND p.doc_type = ?"
        params.append(doc_type)
        desc_parts.append(f"Documento: {doc_type}")

    query += " ORDER BY p.id DESC"

    conn = get_db_connection()
    rows = conn.execute(query, params).fetchall()
    conn.close()

    filters_desc = " | ".join(desc_parts) if desc_parts else "Todas las compras"

    pdf_bytes = generate_purchase_search_pdf(rows, filters_desc)
    if pdf_bytes:
        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Reporte_Compras.pdf",
        )
    return "Error al generar PDF", 500


@app.route('/download_filtered_report_pdf')
def download_filtered_report_pdf():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    f_type = request.args.get('f_type', 'Todos')
    f_status = request.args.get('f_status', 'Todos')
    f_user = request.args.get('f_user', 'Todos')
    f_date = request.args.get('f_date', '')
    f_month = request.args.get('f_month', '')
    f_start = request.args.get('f_start', '')
    f_end = request.args.get('f_end', '')
    report_type = request.args.get('report_type', 'reportes')

    # Construir la misma consulta que usa el index
    where_clause = " WHERE 1=1"
    params = []
    if f_user and f_user != 'Todos':
        where_clause += " AND username = ?"
        params.append(f_user)
    if f_status and f_status != 'Todos':
        where_clause += " AND status = ?"
        params.append(f_status)
    if f_type == 'Dia' and f_date:
        where_clause += " AND date LIKE ?"
        params.append(f"{f_date}%")
    elif f_type == 'Mes' and f_month:
        where_clause += " AND date LIKE ?"
        params.append(f"{f_month}%")
    elif f_type == 'Rango' and f_start and f_end:
        where_clause += " AND substr(date, 1, 10) >= ? AND substr(date, 1, 10) <= ?"
        params.extend([f_start, f_end])

    conn = get_db_connection()
    all_budgets = conn.execute(
        f"SELECT * FROM saved_budgets{where_clause} ORDER BY id DESC", params
    ).fetchall()

    # Calcular totales para el PDF
    total_usd = sum(
        float(b['total_foreign'] or 0) for b in all_budgets if b['currency'] == 'USD'
    )
    total_eur = sum(
        float(b['total_foreign'] or 0) for b in all_budgets if b['currency'] == 'EUR'
    )
    total_ves = sum(float(b['total_ves'] or 0) for b in all_budgets)

    # Mapa de usuarios
    all_users_rows = conn.execute("SELECT * FROM users").fetchall()
    conn.close()

    users_map = {}
    for u in all_users_rows:
        full = f"{u['first_name']} {u['last_name']}".strip()
        if full and u['phone']:
            users_map[u['username']] = f"{full} ({u['phone']})"
        elif full:
            users_map[u['username']] = full
        else:
            users_map[u['username']] = u['username']

    # Determinar título del reporte
    title_str = (
        "Reporte de Ganancias" if report_type == 'ganancias' else "Reporte General"
    )
    if f_user != 'Todos':
        title_str += f" - {users_map.get(f_user, f_user)}"

    # Si es de ganancias, solo mostrar los aprobados
    if report_type == 'ganancias':
        all_budgets = [b for b in all_budgets if b['status'] == 'Aprobado']

    # Generar el PDF usando la función que ya tienes
    pdf_bytes = generate_filtered_report_pdf(
        title_str, all_budgets, total_usd, total_eur, total_ves, users_map
    )

    filename = f"{title_str.replace(' ', '_').replace('/', '_')}.pdf"
    return send_file(
        BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename,
    )


if __name__ == '__main__':
    debug_mode = os.environ.get('RSV_DEBUG', '0') == '1'
    app.run(host='0.0.0.0', port=5000, debug=debug_mode)