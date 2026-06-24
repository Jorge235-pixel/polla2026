from __future__ import annotations

import os
import sqlite3
import socket
import re
import random
import shutil
import unicodedata
from datetime import datetime, timedelta
from uuid import uuid4
from collections import OrderedDict
from functools import wraps
from io import BytesIO

from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'cambia-esta-clave-2026')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = '/var/data/polla_mundial_2026.db'
UPLOAD_FOLDER = '/var/data/uploads'
SEED_DB_PATH = os.path.join(BASE_DIR, 'polla_mundial_2026.db')
SEED_UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

POINTS_EXACT = 5
POINTS_OUTCOME = 3
POINTS_WRONG = -1
PREDICTION_LOCK_MINUTES = 35
MAX_SCORE_VALUE = 15
REGISTRATION_CLOSE = datetime(2026, 6, 10, 21, 35, 00)

AUTHORIZED_PARTICIPANTS_PATH = os.environ.get('AUTHORIZED_PARTICIPANTS_PATH') or os.path.join(BASE_DIR, 'participantes_autorizados.csv')


def normalize_text_for_lookup(value: str) -> str:
    """Normaliza texto para comparar nombres sin depender de mayúsculas, acentos o comas."""
    value = (value or '').strip().casefold()
    value = ''.join(ch for ch in unicodedata.normalize('NFKD', value) if not unicodedata.combining(ch))
    value = re.sub(r'[^a-z0-9]+', ' ', value)
    return ' '.join(value.split())


def normalize_name_tokens(value: str) -> str:
    """Compara nombres por tokens ordenados: acepta 'APELLIDO, NOMBRE' o 'Nombre Apellido'."""
    normalized = normalize_text_for_lookup(value)
    tokens = [token for token in normalized.split() if token]
    return ' '.join(sorted(tokens))


def load_authorized_participants() -> dict[str, str]:
    """Carga la lista autorizada desde participantes_autorizados.csv.

    Formato esperado:
    full_name,email
    MESA HERNANDEZ, CINDY KATHERINE,ck.mesa@uniandes.edu.co
    """
    participants = {}
    if not os.path.exists(AUTHORIZED_PARTICIPANTS_PATH):
        return participants
    import csv
    with open(AUTHORIZED_PARTICIPANTS_PATH, newline='', encoding='utf-8-sig') as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            email = (row.get('email') or row.get('Correo electrónico laboral') or '').strip().lower()
            full_name = (row.get('full_name') or row.get('Apellido y Nombre') or row.get('nombre') or '').strip()
            if email and full_name:
                participants[email] = full_name
    return participants


def is_authorized_participant(full_name: str, email: str) -> tuple[bool, str | None]:
    """Valida que correo y nombre pertenezcan a la misma persona de la lista autorizada."""
    email = (email or '').strip().lower()
    authorized = load_authorized_participants()
    listed_name = authorized.get(email)
    if not listed_name:
        return False, None
    if normalize_name_tokens(full_name) != normalize_name_tokens(listed_name):
        return False, listed_name
    return True, listed_name


# Fases permitidas para creación manual desde el panel administrador.
# IMPORTANTE: las fases eliminatorias ya no se crean automáticamente.
MANUAL_PHASES = [
    'Dieciseisavos de final',
    'Octavos de final',
    'Cuartos de final',
    'Semifinal',
    'Tercer puesto',
    'Final',
    'Eliminación',
    'Grupos',
]


def is_valid_full_name(full_name: str) -> bool:
    """Valida nombres con letras, espacios, acentos, ñ, guiones y apóstrofes."""
    return bool(re.fullmatch(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'\- ]{2,80}", full_name or ''))


def is_strong_password(password: str) -> bool:
    """Valida contraseña fuerte: 8-30 caracteres, mayúscula, minúscula, número y símbolo."""
    if not password or not re.fullmatch(r"(?=.{8,30}$)(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z0-9]).+", password):
        return False
    return True


def parse_non_negative_int(value: str | None, max_value: int = MAX_SCORE_VALUE) -> int:
    """Convierte un marcador y rechaza letras, signos, decimales, negativos o valores exagerados."""
    value = (value or '').strip()
    if not re.fullmatch(r'\d+', value):
        raise ValueError('Solo se permiten números enteros positivos o cero.')
    number = int(value)
    if number > max_value:
        raise ValueError(f'El marcador máximo permitido es {max_value}.')
    return number


# Compatibilidad con código anterior que llamaba esta función.
def is_valid_alphanumeric_password(password: str) -> bool:
    return is_strong_password(password)

TEAM_FLAGS = {
    'México': 'mx', 'Sudáfrica': 'za', 'República de Corea': 'kr', 'Corea del Sur': 'kr', 'República Checa': 'cz',
    'Canadá': 'ca', 'Bosnia y Herzegovina': 'ba', 'Catar': 'qa', 'Qatar': 'qa', 'Suiza': 'ch',
    'Brasil': 'br', 'Marruecos': 'ma', 'Haití': 'ht', 'Escocia': 'gb-sct',
    'Estados Unidos': 'us', 'Paraguay': 'py', 'Australia': 'au', 'Turquía': 'tr',
    'Alemania': 'de', 'Curazao': 'cw', 'Costa de Marfil': 'ci', 'Ecuador': 'ec',
    'Países Bajos': 'nl', 'Japón': 'jp', 'Suecia': 'se', 'Túnez': 'tn',
    'Bélgica': 'be', 'Egipto': 'eg', 'Irán': 'ir', 'Nueva Zelanda': 'nz',
    'España': 'es', 'Cabo Verde': 'cv', 'Arabia Saudí': 'sa', 'Arabia Saudita': 'sa', 'Uruguay': 'uy',
    'Francia': 'fr', 'Senegal': 'sn', 'Irak': 'iq', 'Noruega': 'no',
    'Argentina': 'ar', 'Argelia': 'dz', 'Austria': 'at', 'Jordania': 'jo',
    'Portugal': 'pt', 'RD de Congo': 'cd', 'RD del Congo/Jamaica': 'cd', 'Jamaica/RD del Congo': 'jm',
    'Jamaica/RD de Congo': 'jm', 'República Democrática del Congo': 'cd',
    'Uzbekistán': 'uz', 'Colombia': 'co', 'Inglaterra': 'gb-eng', 'Croacia': 'hr',
    'Ghana': 'gh', 'Panamá': 'pa',
}

TEAM_FLAG_FALLBACK = {
    'México': '🇲🇽', 'Sudáfrica': '🇿🇦', 'República de Corea': '🇰🇷', 'Corea del Sur': '🇰🇷', 'República Checa': '🇨🇿',
    'Canadá': '🇨🇦', 'Bosnia y Herzegovina': '🇧🇦', 'Catar': '🇶🇦', 'Qatar': '🇶🇦', 'Suiza': '🇨🇭',
    'Brasil': '🇧🇷', 'Marruecos': '🇲🇦', 'Haití': '🇭🇹', 'Escocia': '🏴',
    'Estados Unidos': '🇺🇸', 'Paraguay': '🇵🇾', 'Australia': '🇦🇺', 'Turquía': '🇹🇷',
    'Alemania': '🇩🇪', 'Curazao': '🇨🇼', 'Costa de Marfil': '🇨🇮', 'Ecuador': '🇪🇨',
    'Países Bajos': '🇳🇱', 'Japón': '🇯🇵', 'Suecia': '🇸🇪', 'Túnez': '🇹🇳',
    'Bélgica': '🇧🇪', 'Egipto': '🇪🇬', 'Irán': '🇮🇷', 'Nueva Zelanda': '🇳🇿',
    'España': '🇪🇸', 'Cabo Verde': '🇨🇻', 'Arabia Saudí': '🇸🇦', 'Arabia Saudita': '🇸🇦', 'Uruguay': '🇺🇾',
    'Francia': '🇫🇷', 'Senegal': '🇸🇳', 'Irak': '🇮🇶', 'Noruega': '🇳🇴',
    'Argentina': '🇦🇷', 'Argelia': '🇩🇿', 'Austria': '🇦🇹', 'Jordania': '🇯🇴',
    'Portugal': '🇵🇹', 'RD de Congo': '🇨🇩', 'RD del Congo/Jamaica': '🇨🇩/🇯🇲', 'Jamaica/RD del Congo': '🇯🇲/🇨🇩',
    'Jamaica/RD de Congo': '🇯🇲/🇨🇩', 'República Democrática del Congo': '🇨🇩',
    'Uzbekistán': '🇺🇿', 'Colombia': '🇨🇴', 'Inglaterra': '🏴', 'Croacia': '🇭🇷',
    'Ghana': '🇬🇭', 'Panamá': '🇵🇦',
}


GROUP_MATCHES = [
    ('Grupo A', '2026-06-11 14:00', 'México', 'Sudáfrica', 'Ciudad de México'),
    ('Grupo A', '2026-06-11 21:00', 'Corea del Sur', 'República Checa', 'Guadalajara'),
    ('Grupo B', '2026-06-12 14:00', 'Canadá', 'Bosnia y Herzegovina', 'Toronto'),
    ('Grupo D', '2026-06-12 20:00', 'Estados Unidos', 'Paraguay', 'Los Ángeles'),
    ('Grupo B', '2026-06-13 14:00', 'Qatar', 'Suiza', 'San Francisco'),
    ('Grupo C', '2026-06-13 17:00', 'Brasil', 'Marruecos', 'Nueva Jersey'),
    ('Grupo C', '2026-06-13 20:00', 'Haití', 'Escocia', 'Boston'),
    ('Grupo D', '2026-06-13 23:00', 'Australia', 'Turquía', 'Vancouver'),
    ('Grupo E', '2026-06-14 12:00', 'Alemania', 'Curazao', 'Houston'),
    ('Grupo F', '2026-06-14 15:00', 'Países Bajos', 'Japón', 'Dallas'),
    ('Grupo E', '2026-06-14 18:00', 'Costa de Marfil', 'Ecuador', 'Philadelphia'),
    ('Grupo F', '2026-06-14 21:00', 'Suecia', 'Túnez', 'Monterrey'),
    ('Grupo H', '2026-06-15 11:00', 'España', 'Cabo Verde', 'Atlanta'),
    ('Grupo G', '2026-06-15 14:00', 'Bélgica', 'Egipto', 'Seattle'),
    ('Grupo H', '2026-06-15 17:00', 'Arabia Saudita', 'Uruguay', 'Miami'),
    ('Grupo G', '2026-06-15 20:00', 'Irán', 'Nueva Zelanda', 'Los Ángeles'),
    ('Grupo I', '2026-06-16 14:00', 'Francia', 'Senegal', 'Nueva Jersey'),
    ('Grupo I', '2026-06-16 17:00', 'Irak', 'Noruega', 'Boston'),
    ('Grupo J', '2026-06-16 20:00', 'Argentina', 'Argelia', 'Kansas City'),
    ('Grupo J', '2026-06-16 23:00', 'Austria', 'Jordania', 'San Francisco'),
    ('Grupo K', '2026-06-17 12:00', 'Portugal', 'RD de Congo', 'Houston'),
    ('Grupo L', '2026-06-17 15:00', 'Inglaterra', 'Croacia', 'Dallas'),
    ('Grupo L', '2026-06-17 18:00', 'Ghana', 'Panamá', 'Toronto'),
    ('Grupo K', '2026-06-17 21:00', 'Uzbekistán', 'Colombia', 'Ciudad de México'),
    ('Grupo A', '2026-06-18 11:00', 'República Checa', 'Sudáfrica', 'Atlanta'),
    ('Grupo B', '2026-06-18 14:00', 'Suiza', 'Bosnia y Herzegovina', 'Los Ángeles'),
    ('Grupo B', '2026-06-18 17:00', 'Canadá', 'Qatar', 'Vancouver'),
    ('Grupo A', '2026-06-18 20:00', 'México', 'Corea del Sur', 'Guadalajara'),
    ('Grupo D', '2026-06-19 14:00', 'Estados Unidos', 'Australia', 'Seattle'),
    ('Grupo C', '2026-06-19 17:00', 'Escocia', 'Marruecos', 'Boston'),
    ('Grupo C', '2026-06-19 19:30', 'Brasil', 'Haití', 'Philadelphia'),
    ('Grupo D', '2026-06-19 22:00', 'Turquía', 'Paraguay', 'San Francisco'),
    ('Grupo F', '2026-06-20 12:00', 'Países Bajos', 'Suecia', 'Houston'),
    ('Grupo E', '2026-06-20 15:00', 'Alemania', 'Costa de Marfil', 'Toronto'),
    ('Grupo E', '2026-06-20 21:00', 'Ecuador', 'Curazao', 'Kansas City'),
    ('Grupo F', '2026-06-20 23:00', 'Túnez', 'Japón', 'Monterrey'),
    ('Grupo H', '2026-06-21 11:00', 'España', 'Arabia Saudita', 'Atlanta'),
    ('Grupo G', '2026-06-21 14:00', 'Bélgica', 'Irán', 'Los Ángeles'),
    ('Grupo H', '2026-06-21 17:00', 'Uruguay', 'Cabo Verde', 'Miami'),
    ('Grupo G', '2026-06-21 20:00', 'Nueva Zelanda', 'Egipto', 'Vancouver'),
    ('Grupo J', '2026-06-22 12:00', 'Argentina', 'Austria', 'Dallas'),
    ('Grupo I', '2026-06-22 16:00', 'Francia', 'Irak', 'Philadelphia'),
    ('Grupo I', '2026-06-22 19:00', 'Noruega', 'Senegal', 'Nueva Jersey'),
    ('Grupo J', '2026-06-22 22:00', 'Jordania', 'Argelia', 'San Francisco'),
    ('Grupo K', '2026-06-23 12:00', 'Portugal', 'Uzbekistán', 'Houston'),
    ('Grupo L', '2026-06-23 15:00', 'Inglaterra', 'Ghana', 'Boston'),
    ('Grupo L', '2026-06-23 18:00', 'Panamá', 'Croacia', 'Toronto'),
    ('Grupo K', '2026-06-23 21:00', 'Colombia', 'RD de Congo', 'Guadalajara'),
    ('Grupo B', '2026-06-24 14:00', 'Suiza', 'Canadá', 'Vancouver'),
    ('Grupo B', '2026-06-24 14:00', 'Bosnia y Herzegovina', 'Qatar', 'Lumen Field, Seattle'),
    ('Grupo C', '2026-06-24 17:00', 'Marruecos', 'Haití', 'Atlanta'),
    ('Grupo C', '2026-06-24 17:00', 'Brasil', 'Escocia', 'Miami'),
    ('Grupo A', '2026-06-24 20:00', 'Sudáfrica', 'Corea del Sur', 'Monterrey'),
    ('Grupo A', '2026-06-24 20:00', 'República Checa', 'México', 'Ciudad de México'),
    ('Grupo E', '2026-06-25 15:00', 'Curazao', 'Costa de Marfil', 'Philadelphia'),
    ('Grupo E', '2026-06-25 15:00', 'Ecuador', 'Alemania', 'Nueva Jersey'),
    ('Grupo F', '2026-06-25 18:00', 'Japón', 'Suecia', 'Dallas'),
    ('Grupo F', '2026-06-25 18:00', 'Túnez', 'Países Bajos', 'Kansas City'),
    ('Grupo D', '2026-06-25 21:00', 'Paraguay', 'Australia', 'San Francisco'),
    ('Grupo D', '2026-06-25 21:00', 'Turquía', 'Estados Unidos', 'Los Ángeles'),
    ('Grupo I', '2026-06-26 14:00', 'Noruega', 'Francia', 'Boston'),
    ('Grupo I', '2026-06-26 14:00', 'Senegal', 'Irak', 'Toronto'),
    ('Grupo H', '2026-06-26 19:00', 'Cabo Verde', 'Arabia Saudita', 'Houston'),
    ('Grupo H', '2026-06-26 19:00', 'Uruguay', 'España', 'Guadalajara'),
    ('Grupo G', '2026-06-26 22:00', 'Egipto', 'Irán', 'Seattle'),
    ('Grupo G', '2026-06-26 22:00', 'Nueva Zelanda', 'Bélgica', 'Vancouver'),
    ('Grupo L', '2026-06-27 16:00', 'Croacia', 'Ghana', 'Philadelphia'),
    ('Grupo L', '2026-06-27 16:00', 'Panamá', 'Inglaterra', 'Nueva Jersey'),
    ('Grupo K', '2026-06-27 18:30', 'Colombia', 'Portugal', 'Miami'),
    ('Grupo K', '2026-06-27 18:30', 'RD de Congo', 'Uzbekistán', 'Atlanta'),
    ('Grupo J', '2026-06-27 21:00', 'Argelia', 'Austria', 'Kansas City'),
    ('Grupo J', '2026-06-27 21:00', 'Jordania', 'Argentina', 'Dallas'),
]


def _team_key(team_name: str) -> str:
    """Clave flexible para reconocer equipos aunque cambien mayúsculas, acentos o espacios."""
    import unicodedata
    raw = (team_name or '').strip().casefold()
    raw = ' '.join(raw.split())
    return ''.join(ch for ch in unicodedata.normalize('NFKD', raw) if not unicodedata.combining(ch))


def official_team_names() -> list[str]:
    # IMPORTANTE: el asistente para crear la siguiente fase debe listar solo
    # los equipos reales definidos en el calendario de grupos.
    #
    # Antes se mezclaban aquí también las llaves de TEAM_FLAGS y
    # TEAM_FLAG_FALLBACK. Esos diccionarios incluyen alias/compatibilidades
    # para mostrar banderas, por ejemplo: Arabia Saudí, Catar, República de
    # Corea, República Democrática del Congo y combinaciones Jamaica/Congo.
    # Al unirlos con los equipos oficiales, esos alias terminaban apareciendo
    # como opciones adicionales en el desplegable del asistente.
    return sorted({name for _, _, home, away, _ in GROUP_MATCHES for name in (home, away)})


def normalize_team_name(team_name: str) -> str:
    """Mantiene texto manual, pero corrige nombres oficiales conocidos para conservar banderas."""
    cleaned = ' '.join((team_name or '').strip().split())
    aliases = {
        'purtugal': 'Portugal',
        'portugal': 'Portugal',
        'brasil': 'Brasil',
        'brazil': 'Brasil',
        'mexico': 'México',
        'argentina': 'Argentina',
        'colombia': 'Colombia',
        'espana': 'España',
        'spain': 'España',
        'estados unidos': 'Estados Unidos',
        'usa': 'Estados Unidos',
        'eeuu': 'Estados Unidos',
        'ee.uu.': 'Estados Unidos',
        'holanda': 'Países Bajos',
        'paises bajos': 'Países Bajos',
        'arabia saudita': 'Arabia Saudita',
        'arabia saudi': 'Arabia Saudí',
        'qatar': 'Qatar',
        'catar': 'Catar',
        'corea del sur': 'Corea del Sur',
        'republica de corea': 'República de Corea',
        'republica checa': 'República Checa',
        'rd de congo': 'RD de Congo',
        'rd del congo': 'RD de Congo',
        'republica democratica del congo': 'República Democrática del Congo',
        'curazao': 'Curazao',
    }
    key = _team_key(cleaned)
    if key in aliases:
        return aliases[key]
    for official in official_team_names():
        if _team_key(official) == key:
            return official
    return cleaned


def get_team_from_form(field_name: str) -> str:
    manual = request.form.get(field_name, '')
    selected = request.form.get(f'{field_name}_select', '')
    # El campo manual manda si el admin escribió algo; el select sirve como ayuda segura.
    chosen = manual.strip() or selected.strip()
    return normalize_team_name(chosen)


def team_flag(team_name: str) -> str:
    return TEAM_FLAG_FALLBACK.get(normalize_team_name(team_name), '🏳️')


def team_flag_url(team_name: str) -> str | None:
    code = TEAM_FLAGS.get(normalize_team_name(team_name))
    if not code:
        return None
    if code in {'gb-sct', 'gb-eng'}:
        return f"https://flagcdn.com/w40/gb.png"
    return f"https://flagcdn.com/w40/{code}.png"


def team_with_flag(team_name: str) -> str:
    return f"{team_flag(team_name)} {normalize_team_name(team_name)}"






def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_photo(file_storage):
    if not file_storage or not file_storage.filename:
        return None
    if not allowed_file(file_storage.filename):
        return None
    original = secure_filename(file_storage.filename)
    ext = original.rsplit('.', 1)[1].lower()
    filename = f"{uuid4().hex}.{ext}"
    file_storage.save(os.path.join(UPLOAD_FOLDER, filename))
    return filename


def photo_url(filename: str | None) -> str | None:
    if not filename:
        return None
    return url_for('uploaded_file', filename=filename)


@app.route('/uploads/<path:filename>')
def uploaded_file(filename: str):
    return send_from_directory(UPLOAD_FOLDER, filename)


def parse_match_dt(value: str) -> datetime | None:
    try:
        return datetime.strptime(value, '%Y-%m-%d %H:%M')
    except Exception:
        return None


def match_is_open(match) -> bool:
    if match['is_finished']:
        return False
    dt = parse_match_dt(match['match_datetime'])
    if not dt:
        return True
    return datetime.now() < (dt - timedelta(minutes=PREDICTION_LOCK_MINUTES))


def prediction_status_label(pred, match) -> str:
    if not match['is_finished']:
        return 'Pendiente'
    if pred is None or pred['predicted_home_score'] is None:
        return 'Sin pronóstico'
    ph, pa = pred['predicted_home_score'], pred['predicted_away_score']
    rh, ra = match['home_score'], match['away_score']
    if ph == rh and pa == ra:
        return f'Marcador exacto (+{POINTS_EXACT})'
    predicted_outcome = (ph > pa) - (ph < pa)
    real_outcome = (rh > ra) - (rh < ra)
    if predicted_outcome == real_outcome:
        if real_outcome == 0:
            return f'Acertó empate (+{POINTS_OUTCOME})'
        return f'Acertó ganador (+{POINTS_OUTCOME})'
    return f'No acertó ({POINTS_WRONG})'


def prediction_points_from_values(pred_home, pred_away, match) -> int:
    if not match['is_finished'] or pred_home is None or pred_away is None:
        return 0
    if pred_home == match['home_score'] and pred_away == match['away_score']:
        return POINTS_EXACT
    predicted_outcome = (pred_home > pred_away) - (pred_home < pred_away)
    real_outcome = (match['home_score'] > match['away_score']) - (match['home_score'] < match['away_score'])
    return POINTS_OUTCOME if predicted_outcome == real_outcome else POINTS_WRONG



def prediction_explanation_from_values(pred_home, pred_away, match) -> str:
    """Texto claro para participante y administrador sobre el puntaje obtenido."""
    if pred_home is None or pred_away is None:
        return 'Aún no has registrado pronóstico para este partido.'
    if not match['is_finished']:
        return 'Partido pendiente: tus puntos se calculan cuando el administrador registre el resultado final.'
    rh, ra = match['home_score'], match['away_score']
    if pred_home == rh and pred_away == ra:
        return f'Ganaste {POINTS_EXACT} puntos porque acertaste el marcador exacto: {rh} - {ra}.'
    predicted_outcome = (pred_home > pred_away) - (pred_home < pred_away)
    real_outcome = (rh > ra) - (rh < ra)
    if predicted_outcome == real_outcome:
        if real_outcome == 0:
            return f'Ganaste {POINTS_OUTCOME} puntos porque acertaste que el partido terminaba empatado.'
        winner = match['home_team'] if rh > ra else match['away_team']
        return f'Ganaste {POINTS_OUTCOME} puntos porque acertaste el ganador: {winner}.'
    return f'Perdiste {abs(POINTS_WRONG)} punto porque no acertaste nada. El resultado oficial fue {rh} - {ra}.'


def prediction_badge_text_from_values(pred_home, pred_away, match) -> str:
    pts = prediction_points_from_values(pred_home, pred_away, match)
    if pred_home is None or pred_away is None:
        return 'Sin pronóstico'
    if not match['is_finished']:
        return 'Pendiente'
    if pts == POINTS_EXACT:
        return f'✅ Exacto (+{POINTS_EXACT})'
    if pts == POINTS_OUTCOME:
        ph, pa = pred_home, pred_away
        return f'🟡 Empate (+{POINTS_OUTCOME})' if ph == pa else f'🟡 Ganador (+{POINTS_OUTCOME})'
    return f'❌ Falló ({POINTS_WRONG})'

def phase_class(phase: str | None) -> str:
    phase = (phase or '').strip().lower()
    if 'grupo' in phase:
        return 'phase-grupos'
    if 'octavos' in phase:
        return 'phase-octavos'
    if 'cuartos' in phase:
        return 'phase-cuartos'
    if 'semi' in phase:
        return 'phase-semifinal'
    if 'tercer' in phase:
        return 'phase-tercer'
    if 'final' in phase:
        return 'phase-final'
    return 'phase-otro'


@app.template_filter('dt')
def format_dt(value: str) -> str:
    try:
        return datetime.strptime(value, '%Y-%m-%d %H:%M').strftime('%d/%m/%Y %H:%M')
    except Exception:
        return value


@app.template_filter('date_only')
def format_date_only(value: str) -> str:
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return value


@app.context_processor
def inject_helpers():
    return {
        'team_flag': team_flag,
        'team_flag_url': team_flag_url,
        'team_with_flag': team_with_flag,
        'photo_url': photo_url,
        'match_is_open': match_is_open,
        'phase_class': phase_class,
        'is_elimination_phase': is_elimination_phase,
        'registrations_open': registrations_open,
        'registration_close_text': REGISTRATION_CLOSE.strftime('%d/%m/%Y %H:%M'),
        'now_text': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'prediction_lock_minutes': PREDICTION_LOCK_MINUTES,
        'prediction_status_label': prediction_status_label,
        'prediction_points_from_values': prediction_points_from_values,
        'prediction_explanation_from_values': prediction_explanation_from_values,
        'prediction_badge_text_from_values': prediction_badge_text_from_values,
    }


def get_db() -> sqlite3.Connection:
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()


def query_db(query: str, args=(), one=False):
    cur = get_db().execute(query, args)
    rows = cur.fetchall()
    cur.close()
    return (rows[0] if rows else None) if one else rows


def execute_db(query: str, args=()):
    db = get_db()
    cur = db.execute(query, args)
    db.commit()
    lastrowid = cur.lastrowid
    cur.close()
    return lastrowid


def init_db():
    # En Render, /var/data empieza vacío. Si hay una base incluida en el proyecto,
    # se copia una sola vez al disco persistente.
    if DB_PATH != SEED_DB_PATH and not os.path.exists(DB_PATH) and os.path.exists(SEED_DB_PATH):
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        shutil.copy2(SEED_DB_PATH, DB_PATH)

    if UPLOAD_FOLDER != SEED_UPLOAD_FOLDER and os.path.exists(SEED_UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        for name in os.listdir(SEED_UPLOAD_FOLDER):
            src = os.path.join(SEED_UPLOAD_FOLDER, name)
            dst = os.path.join(UPLOAD_FOLDER, name)
            if os.path.isfile(src) and not os.path.exists(dst):
                shutil.copy2(src, dst)

    db = sqlite3.connect(DB_PATH)
    db.executescript(
        '''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            is_admin INTEGER NOT NULL DEFAULT 0,
            photo TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS matches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phase TEXT NOT NULL DEFAULT 'Grupos',
            group_name TEXT,
            match_datetime TEXT NOT NULL,
            home_team TEXT NOT NULL,
            away_team TEXT NOT NULL,
            stadium TEXT,
            home_score INTEGER,
            away_score INTEGER,
            winner_team TEXT,
            penalty_home INTEGER,
            penalty_away INTEGER,
            is_finished INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS predictions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            match_id INTEGER NOT NULL,
            predicted_home_score INTEGER NOT NULL,
            predicted_away_score INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(user_id, match_id),
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(match_id) REFERENCES matches(id)
        );
        '''
    )
    try:
        user_cols = [row[1] for row in db.execute("PRAGMA table_info(users)").fetchall()]
        if 'photo' not in user_cols:
            db.execute("ALTER TABLE users ADD COLUMN photo TEXT")
            db.commit()
    except Exception:
        pass

    # Migración segura: agrega campos para definir ganador manual en eliminatorias empatadas.
    try:
        match_cols = [row[1] for row in db.execute("PRAGMA table_info(matches)").fetchall()]
        if 'winner_team' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN winner_team TEXT")
        if 'penalty_home' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN penalty_home INTEGER")
        if 'penalty_away' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN penalty_away INTEGER")
        # Campos para bracket manual: a qué partido pasa el clasificado y en qué casilla se ubica.
        if 'next_match_id' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN next_match_id INTEGER")
        if 'next_match_slot' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN next_match_slot TEXT")
        if 'classified_team' not in match_cols:
            db.execute("ALTER TABLE matches ADD COLUMN classified_team TEXT")
        db.commit()
    except Exception:
        pass

    db.commit()

    admin = db.execute("SELECT id FROM users WHERE email = ?", ('admin@polla2026.local',)).fetchone()
    if not admin:
        db.execute(
            "INSERT INTO users (full_name, email, password_hash, is_admin, created_at) VALUES (?, ?, ?, ?, ?)",
            ('Administrador', 'admin@polla2026.local', generate_password_hash('Admin2026*'), 1, datetime.now().isoformat(timespec='seconds')),
        )

    current_matches = db.execute('SELECT COUNT(*) FROM matches').fetchone()[0]
    if current_matches == 0:
        for group_name, dt, home, away, stadium in GROUP_MATCHES:
            db.execute(
                '''INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium)
                   VALUES (?, ?, ?, ?, ?, ?)''',
                ('Grupos', group_name, dt, home, away, stadium),
            )
    db.commit()
    db.close()


def reset_database_for_production():
    """Limpia datos de prueba y deja la base lista para iniciar la polla.

    Conserva el usuario administrador, borra participantes y pronósticos,
    reinicia todos los partidos a la programación base de grupos y elimina
    fases creadas automáticamente durante pruebas.
    """
    db = sqlite3.connect(DB_PATH)
    try:
        db.execute('PRAGMA foreign_keys = OFF')
        db.execute('BEGIN')
        db.execute('DELETE FROM predictions')
        db.execute('DELETE FROM users WHERE is_admin = 0')
        db.execute('DELETE FROM matches')
        try:
            db.execute("DELETE FROM sqlite_sequence WHERE name IN ('users', 'matches', 'predictions')")
        except Exception:
            pass
        for group_name, dt, home, away, stadium in GROUP_MATCHES:
            db.execute(
                """INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium,
                                         home_score, away_score, winner_team, penalty_home, penalty_away, is_finished)
                   VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, NULL, NULL, NULL, 0)""",
                ('Grupos', group_name, dt, home, away, stadium),
            )
        admin = db.execute("SELECT id FROM users WHERE email = ?", ('admin@polla2026.local',)).fetchone()
        if not admin:
            db.execute(
                "INSERT INTO users (full_name, email, password_hash, is_admin, created_at) VALUES (?, ?, ?, ?, ?)",
                ('Administrador', 'admin@polla2026.local', generate_password_hash('Admin2026*'), 1, datetime.now().isoformat(timespec='seconds')),
            )
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def registrations_open() -> bool:
    return datetime.now() <= REGISTRATION_CLOSE


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if 'user_id' not in session:
            flash('Debes iniciar sesión para continuar.', 'warning')
            return redirect(url_for('login'))
        user = current_user()
        if not user:
            session.clear()
            flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
            return redirect(url_for('login'))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if 'user_id' not in session:
            flash('Debes iniciar sesión para continuar.', 'warning')
            return redirect(url_for('login'))
        user = current_user()
        if not user:
            session.clear()
            flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
            return redirect(url_for('login'))
        if not bool(user['is_admin']):
            session['is_admin'] = False
            flash('No tienes permisos para ingresar a esta sección.', 'danger')
            return redirect(url_for('dashboard'))
        session['is_admin'] = True
        return view(*args, **kwargs)
    return wrapped


def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    user = query_db('SELECT * FROM users WHERE id = ?', (uid,), one=True)
    if user:
        session['user_name'] = user['full_name']
        session['is_admin'] = bool(user['is_admin'])
        return user
    return None


def available_match_dates():
    rows = query_db("SELECT DISTINCT substr(match_datetime,1,10) AS match_date FROM matches ORDER BY match_date")
    return [row['match_date'] for row in rows]


def get_selected_date() -> str:
    return (request.args.get('fecha') or '').strip()


def build_matches_query(base_select: str):
    selected_date = get_selected_date()
    if selected_date:
        return base_select + " WHERE substr(m.match_datetime,1,10) = ? ORDER BY m.match_datetime, m.id", (selected_date,)
    return base_select + " ORDER BY m.match_datetime, m.id", ()



def phase_sort_priority(phase: str | None) -> int:
    """Prioridad visual para el panel admin: fases más avanzadas primero."""
    phase = (phase or '').strip().lower()
    if 'final' in phase and 'dieciseis' not in phase and 'octavos' not in phase and 'cuartos' not in phase and 'semi' not in phase:
        return 100
    if 'tercer' in phase:
        return 95
    if 'semi' in phase:
        return 90
    if 'cuartos' in phase:
        return 80
    if 'octavos' in phase:
        return 70
    if 'dieciseis' in phase or 'dieciseisavos' in phase:
        return 60
    if 'grupo' in phase:
        return 10
    return 50


def sort_matches_for_admin(matches):
    """
    Ordena el panel de resultados sin borrar partidos:
    1) partidos pendientes primero,
    2) fases siguientes/eliminatorias antes que grupos,
    3) luego fecha y hora,
    4) partidos ya jugados quedan abajo.
    """
    return sorted(
        matches,
        key=lambda m: (
            1 if m['is_finished'] else 0,
            -phase_sort_priority(m['phase']),
            parse_match_dt(m['match_datetime']) or datetime.max,
            m['id'],
        )
    )



def sort_matches_for_participant(matches):
    """Muestra primero partidos disponibles/futuros y manda al final los ya jugados."""
    return sorted(
        matches,
        key=lambda m: (
            1 if m['is_finished'] else 0,
            0 if (m['phase'] or '').strip() == 'Grupos' else 1,
            parse_match_dt(m['match_datetime']) or datetime.max,
            m['id'],
        )
    )

def group_matches_by_day(matches):
    grouped = OrderedDict()
    for match in matches:
        dt = parse_match_dt(match['match_datetime'])
        if dt:
            key = dt.strftime('%Y-%m-%d')
            label = dt.strftime('%A %d/%m/%Y')
        else:
            key = match['match_datetime']
            label = match['match_datetime']
        if key not in grouped:
            grouped[key] = {'key': key, 'label': label, 'matches': []}
        grouped[key]['matches'].append(match)
    return list(grouped.values())


def score_prediction(pred, match):
    return prediction_points_from_values(pred['predicted_home_score'], pred['predicted_away_score'], match)




def outcome_points_for_team(goals_for: int, goals_against: int) -> int:
    if goals_for > goals_against:
        return 3
    if goals_for == goals_against:
        return 1
    return 0


def group_standings():
    """Calcula la tabla de posiciones por grupo con partidos terminados."""
    standings = OrderedDict()
    matches = query_db("SELECT * FROM matches WHERE phase = 'Grupos' ORDER BY group_name, match_datetime, id")
    for m in matches:
        group = m['group_name'] or 'Sin grupo'
        standings.setdefault(group, {})
        for team in (m['home_team'], m['away_team']):
            standings[group].setdefault(team, {
                'group': group, 'team': team, 'pj': 0, 'pg': 0, 'pe': 0, 'pp': 0,
                'gf': 0, 'gc': 0, 'gd': 0, 'pts': 0
            })
        if not m['is_finished']:
            continue
        h, a = m['home_team'], m['away_team']
        hs, as_ = m['home_score'], m['away_score']
        standings[group][h]['pj'] += 1
        standings[group][a]['pj'] += 1
        standings[group][h]['gf'] += hs
        standings[group][h]['gc'] += as_
        standings[group][a]['gf'] += as_
        standings[group][a]['gc'] += hs
        standings[group][h]['gd'] = standings[group][h]['gf'] - standings[group][h]['gc']
        standings[group][a]['gd'] = standings[group][a]['gf'] - standings[group][a]['gc']
        standings[group][h]['pts'] += outcome_points_for_team(hs, as_)
        standings[group][a]['pts'] += outcome_points_for_team(as_, hs)
        if hs > as_:
            standings[group][h]['pg'] += 1; standings[group][a]['pp'] += 1
        elif hs < as_:
            standings[group][a]['pg'] += 1; standings[group][h]['pp'] += 1
        else:
            standings[group][h]['pe'] += 1; standings[group][a]['pe'] += 1
    ordered = OrderedDict()
    for group, teams in standings.items():
        ordered[group] = sorted(teams.values(), key=lambda r: (-r['pts'], -r['gd'], -r['gf'], r['team']))
    return ordered


def all_group_matches_finished() -> bool:
    total = query_db("SELECT COUNT(*) AS c FROM matches WHERE phase = 'Grupos'", one=True)['c']
    pending = query_db("SELECT COUNT(*) AS c FROM matches WHERE phase = 'Grupos' AND is_finished = 0", one=True)['c']
    return total > 0 and pending == 0


def create_group_next_phase_if_ready() -> int:
    """Crea automáticamente la ronda de 32 cuando todos los grupos tengan resultados."""
    if not all_group_matches_finished():
        return 0
    existing = query_db("SELECT COUNT(*) AS c FROM matches WHERE phase = 'Dieciseisavos de final'", one=True)['c']
    if existing:
        return 0
    standings = group_standings()
    first_second = []
    thirds = []
    for group, rows in standings.items():
        if len(rows) >= 1:
            first_second.append(rows[0])
        if len(rows) >= 2:
            first_second.append(rows[1])
        if len(rows) >= 3:
            thirds.append(rows[2])
    thirds = sorted(thirds, key=lambda r: (-r['pts'], -r['gd'], -r['gf'], r['team']))[:8]
    qualifiers = sorted(first_second + thirds, key=lambda r: (-r['pts'], -r['gd'], -r['gf'], r['team']))[:32]
    if len(qualifiers) < 2:
        return 0
    start = datetime(2026, 6, 28, 15, 0)
    created = 0
    for i in range(len(qualifiers) // 2):
        home = qualifiers[i]['team']
        away = qualifiers[-(i + 1)]['team']
        dt = start + timedelta(hours=3 * created)
        execute_db("""INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium)
                      VALUES (?, ?, ?, ?, ?, ?)""",
                   ('Dieciseisavos de final', 'Clasificación automática', dt.strftime('%Y-%m-%d %H:%M'), home, away, 'Por confirmar'))
        created += 1
    return created


ELIMINATION_ORDER = ['Dieciseisavos de final', 'Octavos de final', 'Cuartos de final', 'Semifinal']
NEXT_PHASE = {
    'Dieciseisavos de final': 'Octavos de final',
    'Octavos de final': 'Cuartos de final',
    'Cuartos de final': 'Semifinal',
    'Semifinal': 'Final',
}
NEXT_PHASE_DATES = {
    'Octavos de final': datetime(2026, 7, 4, 15, 0),
    'Cuartos de final': datetime(2026, 7, 9, 15, 0),
    'Semifinal': datetime(2026, 7, 14, 15, 0),
    'Final': datetime(2026, 7, 19, 15, 0),
    'Tercer puesto': datetime(2026, 7, 18, 15, 0),
}


def is_elimination_phase(phase: str | None) -> bool:
    return (phase or '').strip() != 'Grupos'


def match_winner(match):
    if not match['is_finished']:
        return None
    if match['home_score'] > match['away_score']:
        return match['home_team']
    if match['away_score'] > match['home_score']:
        return match['away_team']
    # En eliminatorias puede haber empate en tiempo reglamentario; el admin define quién pasa.
    try:
        winner_team = match['winner_team']
    except Exception:
        winner_team = None
    if winner_team in (match['home_team'], match['away_team']):
        return winner_team
    return None


def match_loser(match):
    if not match['is_finished']:
        return None
    if match['home_score'] > match['away_score']:
        return match['away_team']
    if match['away_score'] > match['home_score']:
        return match['home_team']
    winner = match_winner(match)
    if winner == match['home_team']:
        return match['away_team']
    if winner == match['away_team']:
        return match['home_team']
    return None



PHASE_PREVIOUS = {
    'Octavos de final': 'Dieciseisavos de final',
    'Cuartos de final': 'Octavos de final',
    'Semifinal': 'Cuartos de final',
    'Final': 'Semifinal',
}

PHASE_MATCH_COUNTS = {
    'Dieciseisavos de final': 16,
    'Octavos de final': 8,
    'Cuartos de final': 4,
    'Semifinal': 2,
    'Tercer puesto': 1,
    'Final': 1,
}


def finished_matches_for_phase(phase: str):
    return query_db(
        """SELECT * FROM matches
             WHERE phase = ? AND is_finished = 1
             ORDER BY match_datetime, id""",
        (phase,),
    )


def completed_phase_winners(phase: str) -> list[str]:
    """Ganadores reales de una fase ya terminada, respetando penales si hubo empate."""
    winners = []
    seen = set()
    for match in finished_matches_for_phase(phase):
        winner = match_winner(match)
        if winner and winner not in seen:
            winners.append(winner)
            seen.add(winner)
    return winners


def completed_phase_losers(phase: str) -> list[str]:
    """Perdedores de una fase terminada; se usa para el partido de tercer puesto."""
    losers = []
    seen = set()
    for match in finished_matches_for_phase(phase):
        loser = match_loser(match)
        if loser and loser not in seen:
            losers.append(loser)
            seen.add(loser)
    return losers


def eligible_teams_for_phase(phase: str) -> list[str]:
    """Equipos que deben aparecer en el asistente según la fase seleccionada.

    - Dieciseisavos: permite escoger entre los equipos oficiales/clasificados desde grupos.
    - Octavos: solo ganadores de Dieciseisavos.
    - Cuartos: solo ganadores de Octavos.
    - Semifinal: solo ganadores de Cuartos.
    - Final: solo ganadores de Semifinal.
    - Tercer puesto: perdedores de Semifinal.
    """
    phase = (phase or '').strip()
    if phase == 'Dieciseisavos de final':
        return official_team_names()
    if phase == 'Tercer puesto':
        return completed_phase_losers('Semifinal')
    previous = PHASE_PREVIOUS.get(phase)
    if previous:
        return completed_phase_winners(previous)
    return official_team_names()


def eligible_teams_by_phase() -> dict:
    return {phase: eligible_teams_for_phase(phase) for phase in PHASE_MATCH_COUNTS.keys()}


def elimination_phases_for_bracket() -> list[str]:
    return [
        'Dieciseisavos de final',
        'Octavos de final',
        'Cuartos de final',
        'Semifinal',
        'Final',
        'Tercer puesto',
    ]


def bracket_matches_by_phase() -> dict:
    """Devuelve las eliminatorias agrupadas por fase para mostrar llaves visuales."""
    data = {}
    for phase in elimination_phases_for_bracket():
        data[phase] = query_db(
            """SELECT * FROM matches WHERE phase = ? ORDER BY match_datetime, id""",
            (phase,),
        )
    return data


def score_with_penalties(match, side: str) -> str:
    """Muestra el marcador normal y, si aplica, los penales entre paréntesis."""
    if not match or not match['is_finished']:
        return '-'
    if side == 'home':
        base = match['home_score']
        pens = match['penalty_home'] if 'penalty_home' in match.keys() else None
    else:
        base = match['away_score']
        pens = match['penalty_away'] if 'penalty_away' in match.keys() else None
    if pens is not None:
        return f"{base} ({pens})"
    return str(base)


app.jinja_env.globals['score_with_penalties'] = score_with_penalties


def open_destination_matches(current_match_id: int | None = None):
    """Partidos de fases eliminatorias que pueden recibir un clasificado manual."""
    rows = query_db(
        """SELECT id, phase, match_datetime, home_team, away_team
             FROM matches
            WHERE phase <> 'Grupos'
            ORDER BY match_datetime, id"""
    )
    options = []
    for r in rows:
        if current_match_id and r['id'] == current_match_id:
            continue
        options.append(r)
    return options


def update_classified_destination(match_id: int, winner_team: str, next_match_id, next_match_slot: str | None):
    """Mueve manualmente el clasificado a la siguiente llave si el admin eligió destino."""
    if not next_match_id or not next_match_slot:
        execute_db('UPDATE matches SET next_match_id = NULL, next_match_slot = NULL, classified_team = ? WHERE id = ?', (winner_team, match_id))
        return
    try:
        next_id = int(next_match_id)
    except Exception:
        raise ValueError('Destino inválido.')
    if next_match_slot not in ('home', 'away'):
        raise ValueError('Casilla inválida para el clasificado.')

    destination = query_db('SELECT * FROM matches WHERE id = ?', (next_id,), one=True)
    if not destination:
        raise ValueError('El partido destino no existe.')
    if destination['phase'] == 'Grupos':
        raise ValueError('El destino no puede ser un partido de grupos.')
    if destination['is_finished']:
        raise ValueError('No se puede enviar un clasificado a un partido que ya está cerrado.')

    slot_column = 'home_team' if next_match_slot == 'home' else 'away_team'
    current_value = (destination[slot_column] or '').strip()
    if current_value and current_value.lower() not in ('a definir', 'por definir', 'pendiente') and current_value != winner_team:
        raise ValueError(f'La casilla destino ya tiene asignado a {current_value}. Edita ese partido si deseas cambiarlo.')

    execute_db(f'UPDATE matches SET {slot_column} = ? WHERE id = ?', (winner_team, next_id))
    execute_db('UPDATE matches SET next_match_id = ?, next_match_slot = ?, classified_team = ? WHERE id = ?', (next_id, next_match_slot, winner_team, match_id))

def create_next_elimination_phase(current_phase: str) -> int:
    rows = query_db('SELECT * FROM matches WHERE phase = ? ORDER BY match_datetime, id', (current_phase,))
    if not rows or any(not r['is_finished'] for r in rows):
        return 0
    winners = [match_winner(r) for r in rows]
    if any(w is None for w in winners):
        return 0
    next_phase = NEXT_PHASE.get(current_phase)
    if not next_phase:
        return 0
    existing = query_db('SELECT COUNT(*) AS c FROM matches WHERE phase = ?', (next_phase,), one=True)['c']
    if existing:
        return 0
    created = 0
    start = NEXT_PHASE_DATES.get(next_phase, datetime.now() + timedelta(days=1))
    if current_phase == 'Semifinal':
        losers = [match_loser(r) for r in rows]
        if len(winners) >= 2:
            execute_db("""INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium)
                          VALUES (?, ?, ?, ?, ?, ?)""",
                       ('Final', 'Ganadores semifinal', start.strftime('%Y-%m-%d %H:%M'), winners[0], winners[1], 'Por confirmar'))
            created += 1
        if len(losers) >= 2 and all(losers):
            dt3 = NEXT_PHASE_DATES['Tercer puesto']
            execute_db("""INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium)
                          VALUES (?, ?, ?, ?, ?, ?)""",
                       ('Tercer puesto', 'Perdedores semifinal', dt3.strftime('%Y-%m-%d %H:%M'), losers[0], losers[1], 'Por confirmar'))
            created += 1
        return created
    for i in range(0, len(winners), 2):
        if i + 1 >= len(winners):
            break
        dt = start + timedelta(hours=3 * created)
        execute_db("""INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium)
                      VALUES (?, ?, ?, ?, ?, ?)""",
                   (next_phase, 'Ganadores fase anterior', dt.strftime('%Y-%m-%d %H:%M'), winners[i], winners[i+1], 'Por confirmar'))
        created += 1
    return created


def auto_create_next_phase_after_result() -> int:
    """Desactiva la generación automática de fases eliminatorias.

    Decisión de operación para Mundial 2026:
    - El sistema conserva resultados, pronósticos, puntos y ranking automáticos.
    - Los clasificados, cruces y nuevas fases se crean manualmente desde el panel admin.
    - Esto evita llaves mal armadas cuando termina la fase de grupos o una fase eliminatoria.
    """
    return 0


def admin_prediction_rows(selected_date: str = ''):
    sql = """
        SELECT p.*, u.full_name, u.email, u.photo,
               m.phase, m.group_name, m.match_datetime, m.home_team, m.away_team,
               m.home_score, m.away_score, m.is_finished
        FROM predictions p
        JOIN users u ON u.id = p.user_id
        JOIN matches m ON m.id = p.match_id
    """
    args = []
    if selected_date:
        sql += ' WHERE substr(m.match_datetime,1,10) = ?'
        args.append(selected_date)
    sql += ' ORDER BY m.match_datetime DESC, u.full_name'
    rows = []
    for r in query_db(sql, tuple(args)):
        d = dict(r)
        d['points'] = prediction_points_from_values(d['predicted_home_score'], d['predicted_away_score'], d)
        d['status_label'] = prediction_status_label(d, d)
        rows.append(d)
    return rows

def leaderboard_rows():
    users = query_db('SELECT id, full_name, email, photo FROM users WHERE is_admin = 0 ORDER BY full_name')
    matches = {m['id']: m for m in query_db('SELECT * FROM matches')}
    rows = []
    for user in users:
        preds = query_db('SELECT * FROM predictions WHERE user_id = ?', (user['id'],))
        total = exact = winners = draws = failures = 0
        counted_predictions = 0
        for pred in preds:
            match = matches.get(pred['match_id'])
            if not match:
                continue
            pts = score_prediction(pred, match)
            total += pts
            if match['is_finished']:
                counted_predictions += 1
                ph, pa = pred['predicted_home_score'], pred['predicted_away_score']
                rh, ra = match['home_score'], match['away_score']
                predicted_outcome = (ph > pa) - (ph < pa)
                real_outcome = (rh > ra) - (rh < ra)
                if ph == rh and pa == ra:
                    exact += 1
                elif predicted_outcome == real_outcome and real_outcome == 0:
                    draws += 1
                elif predicted_outcome == real_outcome:
                    winners += 1
                else:
                    failures += 1
        acertados = winners + draws
        rows.append({
            'user_id': user['id'], 'full_name': user['full_name'], 'email': user['email'], 'photo': user['photo'], 'points': total,
            'exact': exact, 'winners': winners, 'draws': draws, 'failures': failures, 'hits': acertados,
            'predictions': len(preds), 'counted_predictions': counted_predictions,
        })
    rows.sort(key=lambda x: (-x['points'], -x['exact'], -x['hits'], x['full_name'].lower()))
    for idx, row in enumerate(rows, start=1):
        row['position'] = idx
    return rows


@app.route('/')
def index():
    selected_date = get_selected_date()
    if selected_date:
        next_matches = query_db('SELECT * FROM matches WHERE substr(match_datetime,1,10) = ? ORDER BY match_datetime, id', (selected_date,))
    else:
        next_matches = query_db('SELECT * FROM matches ORDER BY match_datetime, id')
    ranking = leaderboard_rows()[:10]
    matches_by_day = group_matches_by_day(next_matches)
    return render_template(
        'index.html',
        next_matches=next_matches,
        matches_by_day=matches_by_day,
        ranking=ranking,
        available_dates=available_match_dates(),
        selected_date=selected_date,
    )


@app.route('/register', methods=['GET', 'POST'])
def register():
    if not registrations_open():
        flash('Las inscripciones estuvieron habilitadas hasta el 09/06/2026 23:59.', 'warning')
        return render_template('register.html', registration_closed=True)

    if request.method == 'POST':
        full_name = request.form['full_name'].strip()
        email = request.form['email'].strip().lower()
        password = request.form['password']
        confirm = request.form['confirm_password']
        if not full_name or not email or not password:
            flash('Todos los campos son obligatorios.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        if not is_valid_full_name(full_name):
            flash('El nombre solo debe contener letras y espacios. No uses números ni símbolos.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        authorized_ok, listed_name = is_authorized_participant(full_name, email)
        if not authorized_ok:
            if listed_name:
                flash('El correo existe en la lista autorizada, pero el nombre no coincide. Escríbelo tal como aparece en la lista oficial.', 'danger')
            else:
                flash('No estás autorizado para inscribirte. El nombre y correo deben aparecer en la lista oficial.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        # Guarda el nombre oficial de la lista para evitar duplicados por variaciones de escritura.
        full_name = listed_name
        if password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        if not is_strong_password(password):
            flash('La contraseña debe tener de 8 a 30 caracteres e incluir mayúscula, minúscula, número y símbolo.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        existing = query_db('SELECT id FROM users WHERE email = ?', (email,), one=True)
        if existing:
            flash('Este correo ya está registrado.', 'warning')
            return render_template('register.html', registration_closed=not registrations_open())
        photo_file = request.files.get('photo')
        if photo_file and photo_file.filename and not allowed_file(photo_file.filename):
            flash('Archivo no permitido. La foto debe ser .jpg, .jpeg o .png.', 'danger')
            return render_template('register.html', registration_closed=not registrations_open())
        photo_filename = save_uploaded_photo(photo_file)
        execute_db(
            'INSERT INTO users (full_name, email, password_hash, is_admin, photo, created_at) VALUES (?, ?, ?, 0, ?, ?)',
            (full_name, email, generate_password_hash(password), photo_filename, datetime.now().isoformat(timespec='seconds')),
        )
        flash('Registro exitoso. Ya puedes iniciar sesión.', 'success')
        return redirect(url_for('login'))
    return render_template('register.html', registration_closed=not registrations_open())


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email'].strip().lower()
        password = request.form['password']
        user = query_db('SELECT * FROM users WHERE email = ?', (email,), one=True)
        if not user or not check_password_hash(user['password_hash'], password):
            flash('Correo o contraseña incorrectos.', 'danger')
            return render_template('login.html')
        session['user_id'] = user['id']
        session['user_name'] = user['full_name']
        session['is_admin'] = bool(user['is_admin'])
        flash(f'Bienvenido, {user["full_name"]}.', 'success')
        return redirect(url_for('dashboard'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('Sesión cerrada correctamente.', 'info')
    return redirect(url_for('index'))


@app.route('/dashboard')
@login_required
def dashboard():
    user = current_user()
    if not user:
        session.clear()
        flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
        return redirect(url_for('login'))
    selected_date = get_selected_date()
    params = [user['id']]
    sql = '''
        SELECT m.*, p.predicted_home_score, p.predicted_away_score, p.id AS prediction_id
        FROM matches m
        LEFT JOIN predictions p ON p.match_id = m.id AND p.user_id = ?
    '''
    if selected_date:
        sql += ' WHERE substr(m.match_datetime,1,10) = ?'
        params.append(selected_date)
    sql += ' ORDER BY m.match_datetime, m.id'
    matches = sort_matches_for_participant(query_db(sql, tuple(params)))
    matches_by_day = group_matches_by_day(matches)

    # Total visible para el participante: suma solamente partidos finalizados
    # y explica en pantalla si ganó 5, 3 o 0 puntos por cada partido.
    total_points = sum(
        prediction_points_from_values(m['predicted_home_score'], m['predicted_away_score'], m)
        for m in matches
    )

    bracket_data = bracket_matches_by_phase()
    standings = group_standings()

    return render_template(
        'dashboard.html', matches=matches, matches_by_day=matches_by_day, user=user,
        available_dates=available_match_dates(), selected_date=selected_date,
        total_points=total_points, bracket_data=bracket_data, standings=standings,
    )


@app.route('/predict/<int:match_id>', methods=['POST'])
@login_required
def predict(match_id: int):
    user = current_user()
    if not user:
        session.clear()
        flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
        return redirect(url_for('login'))
    match = query_db('SELECT * FROM matches WHERE id = ?', (match_id,), one=True)
    if not match:
        flash('Partido no encontrado.', 'danger')
        return redirect(url_for('dashboard'))
    if not match_is_open(match):
        flash('Este partido ya está cerrado para pronósticos. El bloqueo inicia 1 día antes del partido.', 'warning')
        return redirect(url_for('dashboard', fecha=request.args.get('fecha') or ''))

    try:
        home = parse_non_negative_int(request.form.get('predicted_home_score'))
        away = parse_non_negative_int(request.form.get('predicted_away_score'))
    except Exception:
        flash('Debes ingresar marcadores válidos entre 0 y 15.', 'danger')
        return redirect(url_for('dashboard'))

    existing = query_db('SELECT id FROM predictions WHERE user_id = ? AND match_id = ?', (user['id'], match_id), one=True)
    if existing:
        execute_db(
            'UPDATE predictions SET predicted_home_score = ?, predicted_away_score = ?, created_at = ? WHERE id = ?',
            (home, away, datetime.now().isoformat(timespec='seconds'), existing['id']),
        )
        flash('Pronóstico corregido correctamente. Puedes modificarlo hasta 1 día antes del partido.', 'success')
    else:
        execute_db(
            '''INSERT INTO predictions (user_id, match_id, predicted_home_score, predicted_away_score, created_at)
               VALUES (?, ?, ?, ?, ?)''',
            (user['id'], match_id, home, away, datetime.now().isoformat(timespec='seconds')),
        )
        flash('Pronóstico guardado correctamente. Puedes modificarlo hasta 1 día antes del partido.', 'success')
    return redirect(url_for('dashboard', fecha=request.args.get('fecha') or '', _anchor=f'match-{match_id}'))


@app.route('/predictions/save-all', methods=['POST'])
@login_required
def save_all_predictions():
    """Guarda de forma masiva todos los pronósticos diligenciados en la pantalla.

    Soluciona el problema de perder marcadores escritos en otros partidos al guardar uno solo.
    Solo se guardan filas con los dos marcadores completos y partidos que sigan abiertos.
    """
    user = current_user()
    if not user:
        session.clear()
        flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
        return redirect(url_for('login'))

    selected_date = request.form.get('fecha') or request.args.get('fecha') or ''
    matches = query_db('SELECT * FROM matches ORDER BY match_datetime, id')
    saved = 0
    skipped = 0
    now = datetime.now().isoformat(timespec='seconds')

    for match in matches:
        home_raw = request.form.get(f'pred_home_{match["id"]}')
        away_raw = request.form.get(f'pred_away_{match["id"]}')
        if (home_raw in (None, '') and away_raw in (None, '')):
            continue
        if home_raw in (None, '') or away_raw in (None, ''):
            skipped += 1
            continue
        if not match_is_open(match):
            skipped += 1
            continue
        try:
            home = parse_non_negative_int(home_raw)
            away = parse_non_negative_int(away_raw)
        except Exception:
            skipped += 1
            continue
        existing = query_db('SELECT id FROM predictions WHERE user_id = ? AND match_id = ?', (user['id'], match['id']), one=True)
        if existing:
            execute_db('UPDATE predictions SET predicted_home_score = ?, predicted_away_score = ?, created_at = ? WHERE id = ?', (home, away, now, existing['id']))
            saved += 1
            continue
        execute_db('INSERT INTO predictions (user_id, match_id, predicted_home_score, predicted_away_score, created_at) VALUES (?, ?, ?, ?, ?)', (user['id'], match['id'], home, away, now))
        saved += 1

    if saved:
        flash(f'Se guardaron {saved} pronósticos correctamente sin borrar los demás campos diligenciados.', 'success')
    if skipped:
        flash(f'{skipped} filas no se guardaron porque estaban incompletas, cerradas o tenían valores inválidos.', 'warning')
    if not saved and not skipped:
        flash('No había marcadores diligenciados para guardar.', 'info')
    return redirect(url_for('dashboard', fecha=selected_date, _anchor='detalle-pronosticos'))


@app.route('/predictions/random', methods=['POST'])
@login_required
def random_predictions():
    """Genera pronósticos aleatorios para el usuario actual sin tocar resultados oficiales.

    Solo afecta los partidos que todavía están abiertos para pronosticar.
    Si ya existía un pronóstico abierto, lo reemplaza para que el usuario pueda
    hacer pruebas completas rápidamente.
    """
    user = current_user()
    if not user:
        session.clear()
        flash('Tu sesión ya no es válida. Inicia sesión nuevamente.', 'warning')
        return redirect(url_for('login'))

    selected_date = request.form.get('fecha') or ''
    matches = query_db('SELECT * FROM matches ORDER BY match_datetime, id')
    created = 0
    updated = 0
    now = datetime.now().isoformat(timespec='seconds')

    for match in matches:
        if not match_is_open(match):
            continue

        home = random.randint(0, 5)
        away = random.randint(0, 5)

        existing = query_db(
            'SELECT id FROM predictions WHERE user_id = ? AND match_id = ?',
            (user['id'], match['id']),
            one=True,
        )
        if existing:
            continue
        execute_db(
            '''INSERT INTO predictions (user_id, match_id, predicted_home_score, predicted_away_score, created_at)
               VALUES (?, ?, ?, ?, ?)''',
            (user['id'], match['id'], home, away, now),
        )
        created += 1

    total = created + updated
    if total:
        flash(f'Se generaron pronósticos automáticos para {total} partidos abiertos que aún no tenían marcador. Nuevos: {created}. Ya guardados: no se modificaron.', 'success')
    else:
        flash('No hay partidos abiertos para generar pronósticos automáticos.', 'info')

    return redirect(url_for('dashboard', fecha=selected_date, _anchor='detalle-pronosticos'))


@app.route('/admin/prediction/<int:prediction_id>/update', methods=['POST'])
@admin_required
def update_prediction_admin(prediction_id: int):
    """Permite que solo el administrador corrija un pronóstico bloqueado del participante."""
    selected_date = request.form.get('fecha') or request.args.get('fecha') or ''
    pred = query_db('SELECT * FROM predictions WHERE id = ?', (prediction_id,), one=True)
    if not pred:
        flash('Pronóstico no encontrado.', 'danger')
        return redirect(url_for('admin_panel', fecha=selected_date))
    try:
        home = parse_non_negative_int(request.form.get('predicted_home_score'))
        away = parse_non_negative_int(request.form.get('predicted_away_score'))
    except Exception:
        flash('Debes ingresar marcadores válidos entre 0 y 15 para corregir el pronóstico.', 'danger')
        return redirect(url_for('admin_panel', fecha=selected_date, _anchor='pronosticos-usuarios'))
    execute_db(
        'UPDATE predictions SET predicted_home_score = ?, predicted_away_score = ?, created_at = ? WHERE id = ?',
        (home, away, datetime.now().isoformat(timespec='seconds'), prediction_id),
    )
    flash('Pronóstico corregido por el administrador correctamente.', 'success')
    return redirect(url_for('admin_panel', fecha=selected_date, _anchor='pronosticos-usuarios'))


@app.route('/ranking')
def ranking():
    rows = leaderboard_rows()
    first = rows[0] if len(rows) > 0 else None
    second = rows[1] if len(rows) > 1 else None
    is_admin = bool(session.get('is_admin'))
    return render_template('ranking.html', rows=rows, first=first, second=second, is_admin=is_admin)


@app.route('/export/ranking.xlsx')
@admin_required
def export_ranking():
    rows = leaderboard_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ranking Polla'
    ws.append(['Posición', 'Participante', 'Correo', 'Puntos', 'Marcadores exactos', 'Fallos', 'Aciertos', 'Pronósticos registrados'])
    for row in rows:
        ws.append([row['position'], row['full_name'], row['email'], row['points'], row['exact'], row['failures'], row['hits'], row['predictions']])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='ranking_polla_mundial_2026.xlsx')



def create_empty_bracket_skeleton() -> int:
    """Crea llaves vacías con A definir para probar el bracket manual."""
    schedule = [
        ('Dieciseisavos de final', 16, datetime(2026, 6, 28, 14, 0)),
        ('Octavos de final', 8, datetime(2026, 7, 4, 12, 0)),
        ('Cuartos de final', 4, datetime(2026, 7, 9, 15, 0)),
        ('Semifinal', 2, datetime(2026, 7, 14, 15, 0)),
        ('Tercer puesto', 1, datetime(2026, 7, 18, 15, 0)),
        ('Final', 1, datetime(2026, 7, 19, 15, 0)),
    ]
    created = 0
    for phase, qty, start in schedule:
        existing = query_db('SELECT COUNT(*) AS c FROM matches WHERE phase = ?', (phase,), one=True)['c']
        if existing:
            continue
        for i in range(qty):
            dt = start + timedelta(hours=4 * i)
            execute_db(
                """INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium,
                                          home_score, away_score, winner_team, penalty_home, penalty_away, is_finished)
                   VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, NULL, NULL, NULL, 0)""",
                (phase, f'Llave {i + 1}', dt.strftime('%Y-%m-%d %H:%M'), 'A definir', 'A definir', 'Por confirmar'),
            )
            created += 1
    return created


@app.route('/admin/create-bracket-skeleton', methods=['POST'])
@admin_required
def create_bracket_skeleton():
    created = create_empty_bracket_skeleton()
    if created:
        flash(f'Se crearon {created} partidos vacíos para probar el bracket manual.', 'success')
    else:
        flash('Las llaves vacías ya existen. No se duplicaron partidos.', 'info')
    return redirect(url_for('admin_panel'))

@app.route('/admin')
@admin_required
def admin_panel():
    users = query_db('SELECT id, full_name, email, photo, created_at FROM users WHERE is_admin = 0 ORDER BY created_at DESC')
    selected_date = get_selected_date()
    if selected_date:
        matches = query_db('SELECT * FROM matches WHERE substr(match_datetime,1,10) = ? ORDER BY match_datetime, id', (selected_date,))
    else:
        matches = query_db('SELECT * FROM matches ORDER BY match_datetime, id')
    matches = sort_matches_for_admin(matches)
    ranking = leaderboard_rows()
    prediction_rows = admin_prediction_rows(selected_date)
    standings = group_standings()
    matches_by_day = group_matches_by_day(matches)
    bracket_data = bracket_matches_by_phase()
    destination_matches = open_destination_matches()
    return render_template(
        'admin.html', users=users, matches=matches, matches_by_day=matches_by_day, ranking=ranking,
        prediction_rows=prediction_rows, standings=standings, bracket_data=bracket_data,
        destination_matches=destination_matches,
        available_dates=available_match_dates(), selected_date=selected_date,
        team_names=official_team_names(), manual_phases=MANUAL_PHASES,
        eligible_teams_by_phase=eligible_teams_by_phase(), phase_match_counts=PHASE_MATCH_COUNTS,
    )


@app.route('/admin/result/<int:match_id>', methods=['POST'])
@admin_required
def save_result(match_id: int):
    def back_to_match():
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or '', _anchor=f'admin-match-{match_id}'))
    match = query_db('SELECT * FROM matches WHERE id = ?', (match_id,), one=True)
    if not match:
        flash('Partido no encontrado.', 'danger')
        return back_to_match()
    try:
        home_score = parse_non_negative_int(request.form.get('home_score'))
        away_score = parse_non_negative_int(request.form.get('away_score'))
    except Exception:
        flash('Marcador inválido. Usa números entre 0 y 15.', 'danger')
        return back_to_match()

    winner_team = None
    penalty_home = None
    penalty_away = None

    if home_score == away_score and is_elimination_phase(match['phase']):
        winner_team = (request.form.get('winner_team') or '').strip()
        if winner_team not in (match['home_team'], match['away_team']):
            flash('Este partido eliminatorio quedó empatado. Selecciona el equipo ganador para poder crear la siguiente fase.', 'warning')
            return back_to_match()

        ph = request.form.get('penalty_home')
        pa = request.form.get('penalty_away')
        if ph in (None, '') or pa in (None, ''):
            flash('Este partido eliminatorio quedó empatado. Debes escribir los penales de ambos equipos.', 'warning')
            return back_to_match()
        try:
            penalty_home = parse_non_negative_int(ph, max_value=30)
            penalty_away = parse_non_negative_int(pa, max_value=30)
            if penalty_home == penalty_away:
                raise ValueError
        except Exception:
            flash('Penales inválidos. Escribe valores numéricos y asegúrate de que haya un ganador por penales.', 'danger')
            return back_to_match()

        penalty_winner = match['home_team'] if penalty_home > penalty_away else match['away_team']
        if winner_team != penalty_winner:
            flash('El equipo seleccionado como ganador no coincide con el marcador de penales.', 'warning')
            return back_to_match()
    elif home_score > away_score:
        winner_team = match['home_team']
    elif away_score > home_score:
        winner_team = match['away_team']

    execute_db(
        'UPDATE matches SET home_score = ?, away_score = ?, winner_team = ?, penalty_home = ?, penalty_away = ?, is_finished = 1 WHERE id = ?',
        (home_score, away_score, winner_team, penalty_home, penalty_away, match_id),
    )
    next_match_id = request.form.get('next_match_id')
    next_match_slot = request.form.get('next_match_slot')
    if is_elimination_phase(match['phase']) and winner_team and next_match_id and next_match_slot:
        try:
            update_classified_destination(match_id, winner_team, next_match_id, next_match_slot)
            flash('Resultado guardado y clasificado enviado correctamente a la siguiente fase.', 'success')
        except ValueError as exc:
            flash(f'Resultado guardado, pero no se pudo mover el clasificado: {exc}', 'warning')
            return back_to_match()
    auto_create_next_phase_after_result()
    flash('Resultado guardado correctamente. Los puntos y el ranking fueron actualizados automáticamente. Las siguientes fases se manejan manualmente desde el panel administrador.', 'success')
    return back_to_match()



@app.route('/admin/classify/<int:match_id>', methods=['POST'])
@admin_required
def classify_match_winner(match_id: int):
    """Permite escoger manualmente a qué llave pasa el ganador de un partido ya cerrado."""
    match = query_db('SELECT * FROM matches WHERE id = ?', (match_id,), one=True)
    if not match:
        flash('Partido no encontrado.', 'danger')
        return redirect(url_for('admin_panel'))
    if not is_elimination_phase(match['phase']):
        flash('La clasificación manual solo aplica para fases eliminatorias.', 'warning')
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))
    winner_team = match_winner(match)
    if not winner_team:
        flash('Primero debes cerrar el partido y definir el ganador.', 'warning')
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))
    next_match_id = request.form.get('next_match_id')
    next_match_slot = request.form.get('next_match_slot')
    try:
        update_classified_destination(match_id, winner_team, next_match_id, next_match_slot)
    except ValueError as exc:
        flash(str(exc), 'danger')
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))
    flash(f'{winner_team} fue enviado correctamente a la siguiente fase.', 'success')
    return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))

@app.route('/admin/reopen/<int:match_id>', methods=['POST'])
@admin_required
def reopen_match(match_id: int):
    execute_db('UPDATE matches SET is_finished = 0, home_score = NULL, away_score = NULL, winner_team = NULL, penalty_home = NULL, penalty_away = NULL, classified_team = NULL WHERE id = ?', (match_id,))
    flash('Partido reabierto para ajustes.', 'info')
    return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))


@app.route('/admin/add-match', methods=['POST'])
@admin_required
def add_match():
    """Crea partidos manuales para cualquier fase sin generar llaves automáticas."""
    phase = (request.form.get('phase') or '').strip()
    group_name = (request.form.get('group_name') or '').strip() or None
    match_datetime = (request.form.get('match_datetime') or '').strip().replace('T', ' ')
    if len(match_datetime) > 16:
        match_datetime = match_datetime[:16]
    home_team = get_team_from_form('home_team')
    away_team = get_team_from_form('away_team')
    stadium = (request.form.get('stadium') or '').strip() or 'Por confirmar'

    if not phase:
        flash('Debes seleccionar o escribir la fase del partido.', 'danger')
        return redirect(url_for('admin_panel'))
    if not match_datetime or not home_team or not away_team:
        flash('Debes completar fecha, hora, equipo local y equipo visitante.', 'danger')
        return redirect(url_for('admin_panel'))
    both_placeholder = home_team.lower() in ('a definir', 'por definir', 'pendiente') and away_team.lower() in ('a definir', 'por definir', 'pendiente')
    if home_team == away_team and not both_placeholder:
        flash('El equipo local y visitante no pueden ser iguales.', 'warning')
        return redirect(url_for('admin_panel'))
    if not parse_match_dt(match_datetime):
        flash('Fecha y hora inválidas. Usa una fecha válida desde el calendario.', 'danger')
        return redirect(url_for('admin_panel'))

    duplicate = query_db(
        """
        SELECT id FROM matches
         WHERE lower(trim(phase)) = lower(trim(?))
           AND (
                (lower(trim(home_team)) = lower(trim(?)) AND lower(trim(away_team)) = lower(trim(?)))
             OR (lower(trim(home_team)) = lower(trim(?)) AND lower(trim(away_team)) = lower(trim(?)))
           )
         LIMIT 1
        """,
        (phase, home_team, away_team, away_team, home_team),
        one=True,
    )
    if duplicate and not both_placeholder:
        flash('Ese cruce ya existe en la misma fase. Revisa antes de crear otro partido igual.', 'warning')
        return redirect(url_for('admin_panel'))

    execute_db(
        """INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium,
                                home_score, away_score, winner_team, penalty_home, penalty_away, is_finished)
           VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, NULL, NULL, NULL, 0)""",
        (phase, group_name, match_datetime, home_team, away_team, stadium),
    )
    flash('Partido manual creado correctamente. No se generó ninguna fase automática.', 'success')
    return redirect(url_for('admin_panel'))



@app.route('/admin/create-phase-wizard', methods=['POST'])
@admin_required
def create_phase_wizard():
    """Asistente para crear una fase completa escogiendo equipos, fecha y hora por partido.

    Cada cruce exige su propia fecha/hora para cumplir la creación manual solicitada.
    """
    phase = (request.form.get('wizard_phase') or '').strip()
    if not phase:
        flash('Selecciona la fase que quieres crear.', 'danger')
        return redirect(url_for('admin_panel'))
    if phase == 'Grupos':
        flash('Este asistente es solo para eliminatorias. La fase de grupos ya está cargada.', 'warning')
        return redirect(url_for('admin_panel'))
    allowed_teams = eligible_teams_for_phase(phase)
    allowed_set = set(allowed_teams)
    if phase != 'Dieciseisavos de final' and not allowed_teams:
        previous = PHASE_PREVIOUS.get(phase, 'la fase anterior')
        if phase == 'Tercer puesto':
            previous = 'Semifinal'
        flash(f'Primero debes terminar {previous} y guardar sus ganadores antes de crear {phase}.', 'warning')
        return redirect(url_for('admin_panel'))

    homes = request.form.getlist('wizard_home[]')
    aways = request.form.getlist('wizard_away[]')
    datetimes = request.form.getlist('wizard_datetime[]')

    created = 0
    errors = []
    seen = set()
    used_teams = set()
    max_rows = max(len(homes), len(aways), len(datetimes))

    for i in range(max_rows):
        home_team = normalize_team_name(homes[i] if i < len(homes) else '')
        away_team = normalize_team_name(aways[i] if i < len(aways) else '')
        raw_datetime = (datetimes[i] if i < len(datetimes) else '').strip().replace('T', ' ')
        if len(raw_datetime) > 16:
            raw_datetime = raw_datetime[:16]

        # Fila totalmente vacía: se ignora para que el admin pueda crear solo los partidos que necesite.
        if not home_team and not away_team:
            continue

        row_label = f'Partido {i + 1}'
        if not home_team or not away_team:
            errors.append(f'{row_label}: completa equipo local y visitante.')
            continue
        if not raw_datetime or not parse_match_dt(raw_datetime):
            errors.append(f'{row_label}: selecciona una fecha y hora válida.')
            continue
        if home_team not in allowed_set or away_team not in allowed_set:
            errors.append(f'{row_label}: solo puedes escoger equipos habilitados para {phase}.')
            continue
        if home_team == away_team:
            errors.append(f'{row_label}: el equipo local y visitante no pueden ser iguales.')
            continue
        if home_team in used_teams or away_team in used_teams:
            errors.append(f'{row_label}: un equipo no puede quedar en dos partidos de la misma fase.')
            continue
        used_teams.add(home_team)
        used_teams.add(away_team)

        match_datetime = raw_datetime

        key = tuple(sorted([home_team.lower(), away_team.lower()])) + (phase.lower(),)
        if key in seen:
            errors.append(f'{row_label}: cruce repetido en el formulario.')
            continue
        seen.add(key)

        duplicate = query_db(
            """
            SELECT id FROM matches
             WHERE lower(trim(phase)) = lower(trim(?))
               AND (
                    (lower(trim(home_team)) = lower(trim(?)) AND lower(trim(away_team)) = lower(trim(?)))
                 OR (lower(trim(home_team)) = lower(trim(?)) AND lower(trim(away_team)) = lower(trim(?)))
               )
             LIMIT 1
            """,
            (phase, home_team, away_team, away_team, home_team),
            one=True,
        )
        if duplicate:
            errors.append(f'{row_label}: {home_team} vs {away_team} ya existe en {phase}.')
            continue

        execute_db(
            """INSERT INTO matches (phase, group_name, match_datetime, home_team, away_team, stadium,
                                    home_score, away_score, winner_team, penalty_home, penalty_away, is_finished)
               VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, NULL, NULL, NULL, 0)""",
            (phase, f'Llave {i + 1}', match_datetime, home_team, away_team, 'Por confirmar'),
        )
        created += 1

    if created:
        flash(f'Se crearon {created} partidos para {phase} con la fecha y hora indicada en cada cruce. Ya aparecen en el bracket y en cargar resultados.', 'success')
    if errors:
        flash('Algunas filas no se guardaron: ' + ' | '.join(errors[:4]), 'warning')
    if not created and not errors:
        flash('No se creó ningún partido. Debes llenar al menos una fila.', 'warning')
    return redirect(url_for('admin_panel'))

@app.route('/admin/match/<int:match_id>/edit', methods=['GET', 'POST'])
@admin_required
def edit_match(match_id: int):
    """Permite corregir equipos, fase, grupo, fecha, hora y estadio sin borrar resultados ni pronósticos."""
    match = query_db('SELECT * FROM matches WHERE id = ?', (match_id,), one=True)
    if not match:
        flash('Partido no encontrado.', 'danger')
        return redirect(url_for('admin_panel'))

    if request.method == 'POST':
        phase = (request.form.get('phase') or '').strip() or 'Grupos'
        group_name = (request.form.get('group_name') or '').strip() or None
        raw_datetime = (request.form.get('match_datetime') or '').strip()
        home_team = get_team_from_form('home_team')
        away_team = get_team_from_form('away_team')
        stadium = (request.form.get('stadium') or '').strip() or 'Por confirmar'
        keep_result = request.form.get('keep_result') == '1'

        match_datetime = raw_datetime.replace('T', ' ')
        if len(match_datetime) > 16:
            match_datetime = match_datetime[:16]

        if not phase or not match_datetime or not home_team or not away_team:
            flash('Debes completar fase, fecha/hora y los dos equipos.', 'danger')
            return redirect(url_for('edit_match', match_id=match_id, fecha=request.args.get('fecha') or ''))
        both_placeholder = home_team.lower() in ('a definir', 'por definir', 'pendiente') and away_team.lower() in ('a definir', 'por definir', 'pendiente')
        if home_team == away_team and not both_placeholder:
            flash('El equipo local y visitante no pueden ser iguales.', 'warning')
            return redirect(url_for('edit_match', match_id=match_id, fecha=request.args.get('fecha') or ''))
        if not parse_match_dt(match_datetime):
            flash('Fecha y hora inválidas. Usa el formato 2026-06-28 15:00.', 'danger')
            return redirect(url_for('edit_match', match_id=match_id, fecha=request.args.get('fecha') or ''))

        if keep_result:
            execute_db("""
                UPDATE matches
                   SET phase = ?, group_name = ?, match_datetime = ?, home_team = ?, away_team = ?, stadium = ?
                 WHERE id = ?
            """, (phase, group_name, match_datetime, home_team, away_team, stadium, match_id))
        else:
            execute_db("""
                UPDATE matches
                   SET phase = ?, group_name = ?, match_datetime = ?, home_team = ?, away_team = ?, stadium = ?,
                       home_score = NULL, away_score = NULL, winner_team = NULL,
                       penalty_home = NULL, penalty_away = NULL, is_finished = 0
                 WHERE id = ?
            """, (phase, group_name, match_datetime, home_team, away_team, stadium, match_id))

        flash('Partido actualizado correctamente.', 'success')
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))

    match_dict = dict(match)
    match_dict['match_datetime_form'] = (match_dict.get('match_datetime') or '').replace(' ', 'T')
    return render_template('edit_match.html', match=match_dict, team_names=official_team_names(), selected_date=request.args.get('fecha') or '')



@app.route('/admin/simulate-group-results', methods=['POST'])
@admin_required
def simulate_group_results():
    """Herramienta de pruebas: llena marcadores aleatorios solo en la fase de grupos.

    No borra partidos, no toca participantes, no crea fases automáticas.
    Sirve para probar tablas, ranking y luego escoger manualmente clasificados.
    """
    rows = query_db("SELECT * FROM matches WHERE phase = 'Grupos' ORDER BY match_datetime, id")
    if not rows:
        flash('No hay partidos de grupos para simular.', 'warning')
        return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))

    updated = 0
    for m in rows:
        # Marcadores realistas para pruebas: 0 a 4 goles.
        home_score = random.randint(0, 4)
        away_score = random.randint(0, 4)
        winner_team = None
        if home_score > away_score:
            winner_team = m['home_team']
        elif away_score > home_score:
            winner_team = m['away_team']
        execute_db(
            """UPDATE matches
                  SET home_score = ?, away_score = ?, winner_team = ?,
                      penalty_home = NULL, penalty_away = NULL, is_finished = 1,
                      next_match_id = NULL, next_match_slot = NULL, classified_team = NULL
                WHERE id = ?""",
            (home_score, away_score, winner_team, m['id']),
        )
        updated += 1

    flash(f'Se generaron marcadores aleatorios para {updated} partidos de grupos. Ya puedes revisar tablas y probar clasificados manuales.', 'success')
    return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))


@app.route('/admin/clear-results-only', methods=['POST'])
@admin_required
def clear_results_only():
    """Limpia únicamente marcadores y estados de partidos, sin borrar fixture ni usuarios."""
    execute_db(
        """UPDATE matches
              SET home_score = NULL, away_score = NULL, winner_team = NULL,
                  penalty_home = NULL, penalty_away = NULL, is_finished = 0,
                  next_match_id = NULL, next_match_slot = NULL, classified_team = NULL"""
    )
    flash('Marcadores limpiados correctamente. Se conservaron equipos, fechas, usuarios y partidos.', 'success')
    return redirect(url_for('admin_panel', fecha=request.args.get('fecha') or ''))

@app.route('/admin/clean-database', methods=['POST'])
@admin_required
def clean_database():
    confirmation = (request.form.get('confirmation') or '').strip().upper()
    if confirmation != 'LIMPIAR':
        flash('Limpieza cancelada. Para limpiar la base debes escribir LIMPIAR exactamente.', 'warning')
        return redirect(url_for('admin_panel'))

    try:
        reset_database_for_production()
    except Exception:
        app.logger.exception('Error limpiando base de datos')
        flash('No se pudo limpiar la base de datos. No se aplicaron cambios.', 'danger')
        return redirect(url_for('admin_panel'))

    flash('Base de datos limpia: se borraron participantes y pronósticos de prueba, y los partidos volvieron al calendario inicial.', 'success')
    return redirect(url_for('admin_panel'))


@app.route('/admin/user/<int:user_id>/edit', methods=['GET', 'POST'])
@admin_required
def edit_user(user_id: int):
    user = query_db('SELECT * FROM users WHERE id = ? AND is_admin = 0', (user_id,), one=True)
    if not user:
        flash('Participante no encontrado.', 'danger')
        return redirect(url_for('admin_panel'))

    if request.method == 'POST':
        full_name = (request.form.get('full_name') or '').strip()
        email = (request.form.get('email') or '').strip().lower()
        new_password = (request.form.get('password') or '').strip()

        if not full_name or not email:
            flash('Nombre y correo son obligatorios.', 'danger')
            return render_template('edit_user.html', user=user)
        if not is_valid_full_name(full_name):
            flash('El nombre solo debe contener letras y espacios. No uses números ni símbolos.', 'danger')
            return render_template('edit_user.html', user=user)

        existing = query_db('SELECT id FROM users WHERE email = ? AND id <> ?', (email, user_id), one=True)
        if existing:
            flash('Ese correo ya está registrado por otro participante.', 'warning')
            return render_template('edit_user.html', user=user)

        if new_password:
            if not is_strong_password(new_password):
                flash('La nueva contraseña debe tener de 8 a 30 caracteres e incluir mayúscula, minúscula, número y símbolo.', 'danger')
                return render_template('edit_user.html', user=user)
            execute_db(
                'UPDATE users SET full_name = ?, email = ?, password_hash = ? WHERE id = ?',
                (full_name, email, generate_password_hash(new_password), user_id),
            )
            flash('Participante actualizado y contraseña restablecida.', 'success')
        else:
            execute_db(
                'UPDATE users SET full_name = ?, email = ? WHERE id = ?',
                (full_name, email, user_id),
            )
            flash('Participante actualizado correctamente.', 'success')
        return redirect(url_for('admin_panel'))

    return render_template('edit_user.html', user=user)


@app.route('/admin/user/<int:user_id>/delete', methods=['POST'])
@admin_required
def delete_user(user_id: int):
    user = query_db('SELECT * FROM users WHERE id = ? AND is_admin = 0', (user_id,), one=True)
    if not user:
        flash('Participante no encontrado.', 'danger')
        return redirect(url_for('admin_panel'))

    execute_db('DELETE FROM predictions WHERE user_id = ?', (user_id,))
    execute_db('DELETE FROM users WHERE id = ?', (user_id,))
    flash(f'Participante eliminado: {user["full_name"]}.', 'info')
    return redirect(url_for('admin_panel'))


# Inicializa la base de datos también cuando Render carga la app con Gunicorn.
init_db()

if __name__ == '__main__':
    # init_db() ya se ejecutó arriba; se deja la app lista para correr localmente.

    local_ip = get_local_ip()
    print('\n' + '=' * 70)
    print('Polla Mundial 2026 iniciada como servidor local')
    print(f'Acceso en este equipo: http://127.0.0.1:5055')
    print(f'Acceso desde otros equipos de la misma red: http://{local_ip}:5055')
    print('Si otro equipo no puede entrar, abra el puerto 5055 en el firewall de Windows.')
    print('=' * 70 + '\n')
    port = int(os.environ.get('PORT', 5055))
    app.run(host='0.0.0.0', port=port, debug=False)
