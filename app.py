import math
import time
import uuid
import unicodedata
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo
from typing import Dict, Optional, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError

try:
    from streamlit_js_eval import get_geolocation
except Exception:
    get_geolocation = None


# ============================================================
# CONFIG
# ============================================================

TZ = ZoneInfo("America/Sao_Paulo")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
APP_TITLE = st.secrets.get("APP_TITLE", "LIIVV Checklist")

DEFAULT_DISTANCE_METERS = 200
DEFAULT_CHECKIN_MINUTES = 30
DEFAULT_ALERT_EMAIL = "operacao@liivv.com.br"


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp {
        background: #EFE7DD;
    }

    [data-testid="stHeader"] {
        background: rgba(239, 231, 221, 0.95);
    }

    .main-title {
        color: #0E2A47;
        font-size: 30px;
        font-weight: 900;
        margin-bottom: 2px;
    }

    .subtitle {
        color: #6B7785;
        font-size: 14px;
        margin-bottom: 20px;
    }

    .liivv-card {
        background: #ffffff;
        border: 1px solid #D7CFC3;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        margin-bottom: 16px;
    }

    .metric-card {
        border-radius: 16px;
        padding: 16px;
        color: white;
        min-height: 110px;
        margin-bottom: 10px;
    }

    .metric-title {
        font-size: 14px;
        opacity: 0.95;
        font-weight: 700;
    }

    .metric-value {
        font-size: 30px;
        font-weight: 900;
        margin-top: 8px;
    }

    .metric-sub {
        font-size: 12px;
        opacity: 0.92;
        margin-top: 6px;
    }

    .task-card {
        border-radius: 16px;
        padding: 14px;
        margin: 10px 0;
        border: 1px solid rgba(0,0,0,0.08);
    }

    .task-title {
        font-size: 15px;
        font-weight: 900;
        color: #111827;
    }

    .task-detail {
        font-size: 13px;
        color: #374151;
        margin-top: 6px;
    }

    .task-meta {
        font-size: 12px;
        color: #4B5563;
        margin-top: 6px;
    }

    .pill {
        display: inline-block;
        padding: 4px 9px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 800;
        margin-top: 8px;
    }

    .small-muted {
        color: #6B7785;
        font-size: 13px;
    }

    div[data-testid="stButton"] > button {
        border-radius: 10px;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# GOOGLE SHEETS
# ============================================================

def retryable(fn, tries=5, base_sleep=0.7, max_sleep=8.0):
    last = None

    for i in range(tries):
        try:
            return fn()
        except APIError as exc:
            last = exc
            msg = str(exc)
            quota = "429" in msg or "Quota exceeded" in msg or "RESOURCE_EXHAUSTED" in msg

            if not quota and i >= 1:
                raise

            time.sleep(min(max_sleep, base_sleep * (2 ** i)))

    raise last


@st.cache_resource
def get_client():
    if "google_service_account" in st.secrets:
        creds_dict = dict(st.secrets["google_service_account"])
    elif "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    else:
        raise RuntimeError("Secrets precisa ter [google_service_account] ou [gcp_service_account].")

    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    if not SPREADSHEET_ID:
        st.error("SPREADSHEET_ID não configurado nos Secrets.")
        st.stop()

    try:
        return retryable(lambda: get_client().open_by_key(SPREADSHEET_ID))
    except Exception as exc:
        st.error(
            "Não foi possível abrir a planilha. "
            "Verifique o SPREADSHEET_ID e compartilhe a planilha com o e-mail da service account."
        )
        st.exception(exc)
        st.stop()


@st.cache_data(ttl=45)
def read_sheet(sheet_name: str) -> pd.DataFrame:
    sh = get_spreadsheet()

    try:
        ws = retryable(lambda: sh.worksheet(sheet_name))
    except Exception:
        return pd.DataFrame()

    values = retryable(lambda: ws.get_all_records())
    return pd.DataFrame(values)


def get_worksheet(sheet_name: str):
    sh = get_spreadsheet()
    return retryable(lambda: sh.worksheet(sheet_name))


def append_row(sheet_name: str, row: list):
    ws = get_worksheet(sheet_name)
    retryable(lambda: ws.append_row(row, value_input_option="USER_ENTERED"))
    st.cache_data.clear()


def update_cell_by_key(sheet_name: str, key_col: str, key_value: str, target_col: str, value):
    df = read_sheet(sheet_name)

    if df.empty:
        return False

    if key_col not in df.columns or target_col not in df.columns:
        return False

    match = df[df[key_col].astype(str).str.strip().str.lower() == str(key_value).strip().lower()]

    if match.empty:
        return False

    row_index = int(match.index[0]) + 2
    col_index = list(df.columns).index(target_col) + 1

    ws = get_worksheet(sheet_name)
    retryable(lambda: ws.update_cell(row_index, col_index, value))
    st.cache_data.clear()

    return True


# ============================================================
# HELPERS
# ============================================================

def now_sp() -> datetime:
    return datetime.now(TZ)


def date_str(dt: Optional[datetime] = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%Y-%m-%d")


def time_str(dt: Optional[datetime] = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%H:%M:%S")


def datetime_str(dt: Optional[datetime] = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def weekday_pt(d: date) -> str:
    names = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    return names[d.weekday()]


def strip_accents(text: str) -> str:
    text = str(text or "").strip()
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def norm(value) -> str:
    return strip_accents(str(value or "")).strip().lower()


def id_new(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:10].upper()}"


def to_float(value, default=None):
    try:
        if value is None or value == "":
            return default

        return float(str(value).replace(",", "."))
    except Exception:
        return default


def haversine_meters(lat1, lon1, lat2, lon2) -> float:
    radius = 6_371_000

    phi1 = math.radians(float(lat1))
    phi2 = math.radians(float(lat2))

    d_phi = math.radians(float(lat2) - float(lat1))
    d_lambda = math.radians(float(lon2) - float(lon1))

    a = (
        math.sin(d_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    )

    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return radius * c


def unidade_aplica(valor_unidade: str, unidade_id: str) -> bool:
    return norm(valor_unidade) in {"todas", norm(unidade_id)}


def get_param(parametros: pd.DataFrame, unidade_id: str, parametro: str, default):
    if parametros.empty:
        return default

    required = {"unidade_id", "parametro", "valor"}

    if not required.issubset(set(parametros.columns)):
        return default

    df = parametros.copy()
    df["unidade_id_norm"] = df["unidade_id"].astype(str).map(norm)
    df["parametro_norm"] = df["parametro"].astype(str).map(norm)

    specific = df[
        (df["unidade_id_norm"] == norm(unidade_id))
        & (df["parametro_norm"] == norm(parametro))
    ]

    if not specific.empty and str(specific.iloc[0]["valor"]).strip() != "":
        return specific.iloc[0]["valor"]

    general = df[
        (df["unidade_id_norm"] == "todas")
        & (df["parametro_norm"] == norm(parametro))
    ]

    if not general.empty and str(general.iloc[0]["valor"]).strip() != "":
        return general.iloc[0]["valor"]

    return default


# ============================================================
# DATA
# ============================================================

def load_data():
    return {
        "UNIDADES": read_sheet("UNIDADES"),
        "POSICOES": read_sheet("POSICOES"),
        "PESSOAS": read_sheet("PESSOAS"),
        "PESSOA_UNIDADE_POSICAO": read_sheet("PESSOA_UNIDADE_POSICAO"),
        "USUARIOS_APP": read_sheet("USUARIOS_APP"),
        "CHECKPOINTS": read_sheet("CHECKPOINTS"),
        "CHECKLIST_GERAL_PADRAO": read_sheet("CHECKLIST_GERAL_PADRAO"),
        "CHECKLIST_POSICAO_PADRAO": read_sheet("CHECKLIST_POSICAO_PADRAO"),
        "AGENDA_ATENDIMENTOS": read_sheet("AGENDA_ATENDIMENTOS"),
        "CHECKINS": read_sheet("CHECKINS"),
        "RESPOSTAS_CHECKLIST": read_sheet("RESPOSTAS_CHECKLIST"),
        "ALERTAS": read_sheet("ALERTAS"),
        "PARAMETROS": read_sheet("PARAMETROS"),
    }


def require_columns(df: pd.DataFrame, cols: list[str], sheet_name: str):
    missing = [c for c in cols if c not in df.columns]

    if missing:
        st.error(f"A aba {sheet_name} está sem colunas obrigatórias: {', '.join(missing)}")
        st.stop()


# ============================================================
# AUTH
# ============================================================

def authenticate(login: str, senha: str):
    data = load_data()

    usuarios = data["USUARIOS_APP"]
    pessoas = data["PESSOAS"]
    vinculos = data["PESSOA_UNIDADE_POSICAO"]
    unidades = data["UNIDADES"]
    posicoes = data["POSICOES"]

    require_columns(
        usuarios,
        ["usuario_id", "unidade_id", "pessoa_id", "login", "senha", "perfil", "ativa"],
        "USUARIOS_APP",
    )
    require_columns(pessoas, ["pessoa_id", "nome", "ativa"], "PESSOAS")
    require_columns(
        vinculos,
        ["vinculo_id", "unidade_id", "pessoa_id", "posicao_id", "ativa"],
        "PESSOA_UNIDADE_POSICAO",
    )
    require_columns(unidades, ["unidade_id", "unidade_nome", "ativa"], "UNIDADES")
    require_columns(posicoes, ["posicao_id", "posicao_nome", "ativa"], "POSICOES")

    user = usuarios[
        (usuarios["login"].astype(str).map(norm) == norm(login))
        & (usuarios["senha"].astype(str) == str(senha))
        & (usuarios["ativa"].astype(str).map(norm) == "sim")
    ]

    if user.empty:
        return None, "Login ou senha inválidos."

    user_row = user.iloc[0].to_dict()
    pessoa_id = str(user_row["pessoa_id"])

    pessoa = pessoas[
        (pessoas["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (pessoas["ativa"].astype(str).map(norm) == "sim")
    ]

    if pessoa.empty:
        return None, "Pessoa inativa ou não encontrada."

    pessoa_row = pessoa.iloc[0].to_dict()

    vinculos_df = vinculos[
        (vinculos["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (vinculos["ativa"].astype(str).map(norm) == "sim")
    ].copy()

    vinculos_out = []

    for _, v in vinculos_df.iterrows():
        unidade_id = str(v["unidade_id"])
        posicao_id = str(v["posicao_id"])

        unidade = unidades[unidades["unidade_id"].astype(str).map(norm) == norm(unidade_id)]
        posicao = posicoes[posicoes["posicao_id"].astype(str).map(norm) == norm(posicao_id)]

        if unidade.empty:
            continue

        if norm(unidade.iloc[0].get("ativa", "")) != "sim":
            continue

        if not posicao.empty and norm(posicao.iloc[0].get("ativa", "")) != "sim":
            continue

        vinculos_out.append(
            {
                "vinculo_id": str(v["vinculo_id"]),
                "unidade_id": unidade_id,
                "unidade_nome": str(unidade.iloc[0].get("unidade_nome", unidade_id)),
                "posicao_id": posicao_id,
                "posicao_nome": str(posicao.iloc[0].get("posicao_nome", posicao_id))
                if not posicao.empty
                else posicao_id,
            }
        )

    if not vinculos_out:
        return None, "Usuário sem vínculo ativo com unidade."

    update_cell_by_key(
        "USUARIOS_APP",
        "usuario_id",
        user_row["usuario_id"],
        "ultimo_login",
        datetime_str(),
    )

    session = {
        "usuario_id": str(user_row["usuario_id"]),
        "pessoa_id": pessoa_id,
        "nome": str(pessoa_row["nome"]),
        "perfil": str(user_row.get("perfil", "operacao")),
        "vinculos": vinculos_out,
    }

    return session, None


def logout():
    for key in ["user", "logged", "page", "vinculo_label", "default_checkpoint"]:
        if key in st.session_state:
            del st.session_state[key]

    st.rerun()


# ============================================================
# BUSINESS LOOKUPS
# ============================================================

def get_vinculo(data: dict, pessoa_id: str, unidade_id: str):
    vinculos = data["PESSOA_UNIDADE_POSICAO"]

    if vinculos.empty:
        return None

    match = vinculos[
        (vinculos["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (vinculos["unidade_id"].astype(str).map(norm) == norm(unidade_id))
        & (vinculos["ativa"].astype(str).map(norm) == "sim")
    ]

    if match.empty:
        return None

    return match.iloc[0].to_dict()


def get_unidade(data: dict, unidade_id: str):
    unidades = data["UNIDADES"]

    if unidades.empty:
        return None

    match = unidades[unidades["unidade_id"].astype(str).map(norm) == norm(unidade_id)]

    if match.empty:
        return None

    return match.iloc[0].to_dict()


def get_pessoa(data: dict, pessoa_id: str):
    pessoas = data["PESSOAS"]

    if pessoas.empty:
        return None

    match = pessoas[pessoas["pessoa_id"].astype(str).map(norm) == norm(pessoa_id)]

    if match.empty:
        return None

    return match.iloc[0].to_dict()


def get_posicao(data: dict, posicao_id: str):
    posicoes = data["POSICOES"]

    if posicoes.empty:
        return None

    match = posicoes[posicoes["posicao_id"].astype(str).map(norm) == norm(posicao_id)]

    if match.empty:
        return None

    return match.iloc[0].to_dict()


# ============================================================
# CHECK-IN
# ============================================================

def evaluate_checkin_time(data: dict, pessoa_id: str, unidade_id: str):
    agenda = data["AGENDA_ATENDIMENTOS"]
    parametros = data["PARAMETROS"]

    antecedencia = int(
        to_float(
            get_param(
                parametros,
                unidade_id,
                "CHECKIN_ANTECEDENCIA_MINUTOS",
                DEFAULT_CHECKIN_MINUTES,
            ),
            DEFAULT_CHECKIN_MINUTES,
        )
    )

    today = date_str()
    now = now_sp()

    if agenda.empty:
        return {
            "status_horario": "SEM_AGENDA",
            "agenda_id": "",
            "mensagem": "Não há agenda cadastrada.",
        }

    required = ["unidade_id", "agenda_id", "data", "horario_inicio", "pessoa_id", "status"]

    for c in required:
        if c not in agenda.columns:
            return {
                "status_horario": "SEM_AGENDA",
                "agenda_id": "",
                "mensagem": f"A aba AGENDA_ATENDIMENTOS está sem a coluna {c}.",
            }

    df = agenda[
        (agenda["unidade_id"].astype(str).map(norm) == norm(unidade_id))
        & (agenda["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (agenda["data"].astype(str).str[:10] == today)
        & (agenda["status"].astype(str).map(norm) != "cancelado")
    ].copy()

    if df.empty:
        return {
            "status_horario": "SEM_AGENDA",
            "agenda_id": "",
            "mensagem": "Não há atendimento cadastrado para esta pessoa nesta unidade hoje.",
        }

    df = df.sort_values("horario_inicio")
    first = df.iloc[0].to_dict()
    horario_inicio = str(first["horario_inicio"])[:5]

    try:
        dt_inicio = datetime.strptime(f"{today} {horario_inicio}", "%Y-%m-%d %H:%M").replace(tzinfo=TZ)
    except Exception:
        return {
            "status_horario": "ERRO_AGENDA",
            "agenda_id": str(first.get("agenda_id", "")),
            "mensagem": "Horário de agenda inválido.",
        }

    limite = dt_inicio - timedelta(minutes=antecedencia)

    if now <= limite:
        return {
            "status_horario": "OK",
            "agenda_id": str(first.get("agenda_id", "")),
            "mensagem": "Check-in realizado dentro do prazo.",
        }

    return {
        "status_horario": "ATRASADO",
        "agenda_id": str(first.get("agenda_id", "")),
        "mensagem": f"Check-in realizado após o limite de {antecedencia} minutos antes do primeiro atendimento.",
    }


def create_alert(
    unidade_id: str,
    tipo_alerta: str,
    severidade: str,
    pessoa_id: str,
    nome_pessoa: str,
    unidade_nome: str,
    agenda_id: str,
    mensagem: str,
    email: str,
):
    append_row(
        "ALERTAS",
        [
            unidade_id,
            id_new("ALT"),
            datetime_str(),
            tipo_alerta,
            severidade,
            pessoa_id,
            nome_pessoa,
            unidade_nome,
            agenda_id,
            mensagem,
            "PENDENTE",
            email or DEFAULT_ALERT_EMAIL,
            "",
            "",
        ],
    )


def register_checkin(pessoa_id: str, unidade_id: str, lat: float, lon: float):
    data = load_data()

    pessoa = get_pessoa(data, pessoa_id)
    unidade = get_unidade(data, unidade_id)
    vinculo = get_vinculo(data, pessoa_id, unidade_id)

    if not pessoa:
        raise ValueError("Pessoa não encontrada.")

    if not unidade:
        raise ValueError("Unidade não encontrada.")

    if not vinculo:
        raise ValueError("Pessoa não possui vínculo ativo com esta unidade.")

    posicao = get_posicao(data, str(vinculo["posicao_id"]))

    unidade_lat = to_float(unidade.get("latitude"))
    unidade_lon = to_float(unidade.get("longitude"))

    if unidade_lat is None or unidade_lon is None:
        raise ValueError("Geolocalização da unidade não cadastrada na aba UNIDADES.")

    pessoa_lat = to_float(lat)
    pessoa_lon = to_float(lon)

    if pessoa_lat is None or pessoa_lon is None:
        raise ValueError("Latitude ou longitude da pessoa inválida.")

    parametros = data["PARAMETROS"]

    raio = to_float(unidade.get("raio_permitido_metros"))

    if raio is None:
        raio = to_float(
            get_param(parametros, unidade_id, "DISTANCIA_MAXIMA_METROS", DEFAULT_DISTANCE_METERS),
            DEFAULT_DISTANCE_METERS,
        )

    distancia = haversine_meters(pessoa_lat, pessoa_lon, unidade_lat, unidade_lon)
    status_distancia = "OK" if distancia <= raio else "FORA_DO_RAIO"

    horario = evaluate_checkin_time(data, pessoa_id, unidade_id)

    alerta_gerado = "não"
    email_alerta = str(unidade.get("email_alerta") or DEFAULT_ALERT_EMAIL)

    nome_pessoa = str(pessoa.get("nome", pessoa_id))
    unidade_nome = str(unidade.get("unidade_nome", unidade_id))
    posicao_id = str(vinculo.get("posicao_id", ""))
    posicao_nome = str(posicao.get("posicao_nome", posicao_id)) if posicao else posicao_id

    if status_distancia != "OK":
        alerta_gerado = "sim"

        create_alert(
            unidade_id=unidade_id,
            tipo_alerta="CHECKIN_FORA_DO_RAIO",
            severidade="ALTA",
            pessoa_id=pessoa_id,
            nome_pessoa=nome_pessoa,
            unidade_nome=unidade_nome,
            agenda_id=horario.get("agenda_id", ""),
            mensagem=(
                f"Check-in fora do raio permitido. "
                f"Distância calculada: {distancia:.1f}m. "
                f"Limite permitido: {raio:.1f}m."
            ),
            email=email_alerta,
        )

    if horario["status_horario"] == "ATRASADO":
        alerta_gerado = "sim"

        create_alert(
            unidade_id=unidade_id,
            tipo_alerta="CHECKIN_ATRASADO",
            severidade="MEDIA",
            pessoa_id=pessoa_id,
            nome_pessoa=nome_pessoa,
            unidade_nome=unidade_nome,
            agenda_id=horario.get("agenda_id", ""),
            mensagem=horario.get("mensagem", ""),
            email=email_alerta,
        )

    checkin_id = id_new("CHK")

    append_row(
        "CHECKINS",
        [
            unidade_id,
            checkin_id,
            datetime_str(),
            date_str(),
            time_str(),
            pessoa_id,
            nome_pessoa,
            unidade_nome,
            posicao_id,
            posicao_nome,
            pessoa_lat,
            pessoa_lon,
            unidade_lat,
            unidade_lon,
            round(distancia, 1),
            raio,
            status_distancia,
            horario["status_horario"],
            alerta_gerado,
            horario.get("agenda_id", ""),
            horario.get("mensagem", ""),
        ],
    )

    return {
        "checkin_id": checkin_id,
        "distancia_metros": distancia,
        "raio_permitido_metros": raio,
        "status_distancia": status_distancia,
        "status_horario": horario["status_horario"],
        "alerta_gerado": alerta_gerado,
        "unidade_nome": unidade_nome,
        "posicao_nome": posicao_nome,
        "mensagem": horario.get("mensagem", ""),
    }


# ============================================================
# CHECKLIST STATUS
# ============================================================

def item_key(
    unidade_id: str,
    checkpoint_id: str,
    tipo: str,
    item_id: str,
    pessoa_id: str = "",
) -> Tuple[str, str, str, str, str]:
    return (
        norm(unidade_id),
        norm(checkpoint_id),
        norm(tipo),
        norm(item_id),
        norm(pessoa_id),
    )


def latest_response_map(respostas: pd.DataFrame, day_iso: str) -> Dict[Tuple[str, str, str, str, str], dict]:
    if respostas.empty:
        return {}

    required = {"unidade_id", "data", "checkpoint_id", "tipo_checklist", "item_id", "resposta", "timestamp"}

    if not required.issubset(set(respostas.columns)):
        return {}

    df = respostas.copy()
    df["data"] = df["data"].astype(str).str[:10]
    df = df[df["data"] == day_iso].copy()

    if df.empty:
        return {}

    df["ts_dt"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["ts_dt"]).sort_values("ts_dt")

    if "pessoa_id" not in df.columns:
        df["pessoa_id"] = ""

    latest = df.groupby(
        ["unidade_id", "checkpoint_id", "tipo_checklist", "item_id", "pessoa_id"],
        as_index=False,
    ).tail(1)

    mp = {}

    for _, r in latest.iterrows():
        key = item_key(
            r["unidade_id"],
            r["checkpoint_id"],
            r["tipo_checklist"],
            r["item_id"],
            r.get("pessoa_id", ""),
        )
        mp[key] = r.to_dict()

    return mp


def normalize_response_status(value: str) -> str:
    s = norm(value).replace(" ", "_")

    if s in {"ok", "conforme", "feito", "concluido"}:
        return "OK"

    if s in {"nao_ok", "não_ok", "n_ok", "nok", "nao_conforme", "não_conforme"}:
        return "NAO_OK"

    if s in {"nao_aplicavel", "não_aplicavel", "n_a", "na"}:
        return "NA"

    if s in {"pendente", "desmarcar", ""}:
        return "PENDENTE"

    return "PENDENTE"


def effective_status(response_row: Optional[dict]) -> str:
    if not response_row:
        return "PENDENTE"

    return normalize_response_status(str(response_row.get("resposta", "")))


def card_palette(status: str) -> Tuple[str, str, str]:
    s = (status or "").upper()

    if s == "OK":
        return "#d1fae5", "#065f46", "Concluído"

    if s == "NAO_OK":
        return "#fee2e2", "#991b1b", "Não OK"

    if s == "NA":
        return "#e0e7ff", "#3730a3", "N/A"

    return "#f3f4f6", "#374151", "Pendente"


def metric_palette(ok: int, nok: int, pending: int, total: int):
    if total == 0:
        return "#1f2937"

    if nok > 0:
        return "#7a1f2b"

    if pending > 0 and ok > 0:
        return "#8b6b12"

    if pending > 0 and ok == 0:
        return "#1f2937"

    return "#0b6a5a"


# ============================================================
# CHECKLIST DATA
# ============================================================

def get_checklist_items(data: dict, pessoa_id: str, unidade_id: str, checkpoint_id: str, tipo: str) -> pd.DataFrame:
    vinculo = get_vinculo(data, pessoa_id, unidade_id)

    if not vinculo:
        return pd.DataFrame()

    if tipo == "geral":
        df = data["CHECKLIST_GERAL_PADRAO"]

        if df.empty:
            return pd.DataFrame()

        return df[
            df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
            & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint_id))
            & (df["ativo"].astype(str).map(norm) == "sim")
        ].copy()

    df = data["CHECKLIST_POSICAO_PADRAO"]

    if df.empty:
        return pd.DataFrame()

    posicao_id = str(vinculo["posicao_id"])

    return df[
        df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
        & (df["posicao_id"].astype(str).map(norm) == norm(posicao_id))
        & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint_id))
        & (df["ativo"].astype(str).map(norm) == "sim")
    ].copy()


def find_checklist_item(data: dict, item_id: str, unidade_id: str):
    for sheet in ["CHECKLIST_GERAL_PADRAO", "CHECKLIST_POSICAO_PADRAO"]:
        df = data[sheet]

        if df.empty or "item_padrao_id" not in df.columns:
            continue

        found = df[
            (df["item_padrao_id"].astype(str).map(norm) == norm(item_id))
            & df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
        ]

        if not found.empty:
            return found.iloc[0].to_dict()

    return None


def register_checklist_response(
    unidade_id: str,
    pessoa_id: str,
    checkpoint_id: str,
    tipo_checklist: str,
    item_id: str,
    resposta: str,
    observacao: str = "",
    evidencia_url: str = "",
):
    data = load_data()

    pessoa = get_pessoa(data, pessoa_id)
    unidade = get_unidade(data, unidade_id)
    vinculo = get_vinculo(data, pessoa_id, unidade_id)

    if not pessoa:
        raise ValueError("Pessoa não encontrada.")

    if not unidade:
        raise ValueError("Unidade não encontrada.")

    if not vinculo:
        raise ValueError("Pessoa não possui vínculo ativo com esta unidade.")

    posicao = get_posicao(data, str(vinculo["posicao_id"]))
    item = find_checklist_item(data, item_id, unidade_id)

    nome_pessoa = str(pessoa.get("nome", pessoa_id))
    unidade_nome = str(unidade.get("unidade_nome", unidade_id))
    posicao_id = str(vinculo.get("posicao_id", ""))
    posicao_nome = str(posicao.get("posicao_nome", posicao_id)) if posicao else posicao_id
    item_nome = str(item.get("item", item_id)) if item else item_id

    status = "CONFORME"
    alerta = "não"

    normalized = normalize_response_status(resposta)

    if normalized == "NAO_OK":
        status = "NÃO CONFORME"
        alerta = "sim"

        create_alert(
            unidade_id=unidade_id,
            tipo_alerta="CHECKLIST_NAO_CONFORME",
            severidade="MEDIA",
            pessoa_id=pessoa_id,
            nome_pessoa=nome_pessoa,
            unidade_nome=unidade_nome,
            agenda_id="",
            mensagem=f"Item não conforme no checklist. Item: {item_nome}. Observação: {observacao or ''}",
            email=str(unidade.get("email_alerta") or DEFAULT_ALERT_EMAIL),
        )

    elif normalized == "NA":
        status = "NÃO APLICÁVEL"

    elif normalized == "PENDENTE":
        status = "PENDENTE"

    resposta_id = id_new("RSP")

    append_row(
        "RESPOSTAS_CHECKLIST",
        [
            unidade_id,
            resposta_id,
            datetime_str(),
            date_str(),
            checkpoint_id,
            tipo_checklist,
            pessoa_id,
            nome_pessoa,
            unidade_nome,
            posicao_id,
            posicao_nome,
            item_id,
            item_nome,
            resposta,
            observacao or "",
            evidencia_url or "",
            status,
            alerta,
        ],
    )

    return {
        "resposta_id": resposta_id,
        "status": status,
        "alerta_gerado": alerta,
    }


# ============================================================
# UI LOGIN
# ============================================================

def render_login():
    left, mid, right = st.columns([1, 1.2, 1])

    with mid:
        st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
        st.markdown('<div class="main-title">LIIVV Checklist</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="subtitle">Acesso operacional para check-in e checklist por unidade.</div>',
            unsafe_allow_html=True,
        )

        with st.form("login_form"):
            login = st.text_input("Login")
            senha = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True)

        if submitted:
            if not login or not senha:
                st.warning("Informe login e senha.")
            else:
                session, error = authenticate(login, senha)

                if error:
                    st.error(error)
                else:
                    st.session_state["logged"] = True
                    st.session_state["user"] = session
                    st.session_state["page"] = "Iniciar trabalho"
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI NAVIGATION
# ============================================================

def select_vinculo(user: dict):
    vinculos = user.get("vinculos", [])

    if not vinculos:
        st.error("Usuário sem vínculo ativo.")
        st.stop()

    labels = [f"{v['unidade_nome']} | {v['posicao_nome']}" for v in vinculos]

    if "vinculo_label" not in st.session_state:
        st.session_state["vinculo_label"] = labels[0]

    selected_label = st.sidebar.selectbox(
        "Unidade / posição",
        labels,
        index=labels.index(st.session_state["vinculo_label"])
        if st.session_state["vinculo_label"] in labels
        else 0,
    )

    st.session_state["vinculo_label"] = selected_label

    return vinculos[labels.index(selected_label)]


def get_current_checkpoint():
    now = now_sp().time()

    if now.hour < 12:
        return "CP0700", "07:00 - Abertura"

    if now.hour < 18:
        return "CP1400", "14:00 - Meio do dia"

    return "CP2000", "20:00 - Fechamento"


def user_has_checkin_today(user: dict, vinculo: dict) -> bool:
    data = load_data()
    checkins = data["CHECKINS"]

    if checkins.empty:
        return False

    required = {"unidade_id", "data", "pessoa_id"}

    if not required.issubset(set(checkins.columns)):
        return False

    today = date_str()

    df = checkins[
        (checkins["unidade_id"].astype(str).map(norm) == norm(vinculo["unidade_id"]))
        & (checkins["pessoa_id"].astype(str).map(norm) == norm(user["pessoa_id"]))
        & (checkins["data"].astype(str).str[:10] == today)
    ]

    return not df.empty


# ============================================================
# UI START WORK
# ============================================================

def render_start_work(user: dict, vinculo: dict):
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.markdown('<div class="main-title">Iniciar trabalho</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtitle">{vinculo["unidade_nome"]} | {vinculo["posicao_nome"]}</div>',
        unsafe_allow_html=True,
    )

    checkin_ok = user_has_checkin_today(user, vinculo)
    cp_id, cp_label = get_current_checkpoint()

    c1, c2, c3 = st.columns(3)
    c1.metric("Colaboradora", user["nome"])
    c2.metric("Unidade", vinculo["unidade_nome"])
    c3.metric("Checkpoint sugerido", cp_label)

    if checkin_ok:
        st.success("Check-in de hoje já realizado. Você já pode preencher o checklist.")
    else:
        st.warning("Antes de preencher o checklist, registre seu check-in com geolocalização.")

    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
        st.subheader("1. Check-in")

        if checkin_ok:
            st.info("Check-in já registrado hoje.")
        else:
            st.caption("Use o botão abaixo para registrar sua localização na unidade.")

        if st.button("Abrir check-in", type="primary", use_container_width=True):
            st.session_state["page"] = "Check-in"
            st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
        st.subheader("2. Minhas atividades")

        st.caption("Abra sua lista por posição e marque cada item como OK, Não OK ou N/A.")

        if st.button("Abrir minhas atividades", type="primary", use_container_width=True):
            st.session_state["page"] = "Checklist"
            st.session_state["default_checkpoint"] = cp_id
            st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    render_mini_dashboard(user, vinculo)


def render_mini_dashboard(user: dict, vinculo: dict):
    data = load_data()
    respostas = data["RESPOSTAS_CHECKLIST"]

    unidade_id = vinculo["unidade_id"]
    pessoa_id = user["pessoa_id"]
    today = date_str()
    response_map = latest_response_map(respostas, today)

    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Resumo rápido do dia")

    checkpoint_options = [
        ("CP0700", "07:00"),
        ("CP1400", "14:00"),
        ("CP2000", "20:00"),
    ]

    total_all = 0
    ok_all = 0
    nok_all = 0
    pending_all = 0
    na_all = 0

    for checkpoint_id, _ in checkpoint_options:
        for tipo in ["geral", "posicao"]:
            df_items = get_checklist_items(data, pessoa_id, unidade_id, checkpoint_id, tipo)

            for _, item in df_items.iterrows():
                item_id = str(item.get("item_padrao_id", ""))
                key = item_key(unidade_id, checkpoint_id, tipo, item_id, pessoa_id)
                status = effective_status(response_map.get(key))

                total_all += 1

                if status == "OK":
                    ok_all += 1
                elif status == "NAO_OK":
                    nok_all += 1
                elif status == "NA":
                    na_all += 1
                else:
                    pending_all += 1

    pct = int(round((ok_all / total_all) * 100, 0)) if total_all else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Progresso", f"{pct}%")
    c2.metric("OK", ok_all)
    c3.metric("Não OK", nok_all)
    c4.metric("N/A", na_all)
    c5.metric("Pendentes", pending_all)

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI CHECK-IN
# ============================================================

def render_checkin(user: dict, vinculo: dict):
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Check-in com geolocalização")

    c1, c2, c3 = st.columns(3)
    c1.metric("Pessoa", user["nome"])
    c2.metric("Unidade", vinculo["unidade_nome"])
    c3.metric("Posição", vinculo["posicao_nome"])

    if user_has_checkin_today(user, vinculo):
        st.success("Você já realizou check-in hoje nesta unidade.")

    st.caption(
        "Permita o acesso à localização no navegador. "
        "Em prédios corporativos, o GPS pode oscilar. O raio permitido é configurado na aba UNIDADES."
    )

    location = None

    if get_geolocation is not None:
        try:
            location = get_geolocation()
        except Exception:
            location = None

    default_lat = None
    default_lon = None

    if location and isinstance(location, dict) and location.get("coords"):
        default_lat = location["coords"].get("latitude")
        default_lon = location["coords"].get("longitude")

    with st.expander("Ver localização capturada", expanded=False):
        st.write(location if location else "Localização ainda não capturada ou bloqueada.")

    manual = st.checkbox(
        "Informar latitude e longitude manualmente",
        value=not bool(default_lat and default_lon),
    )

    if manual:
        lat = st.number_input("Latitude atual", value=0.0, format="%.8f")
        lon = st.number_input("Longitude atual", value=0.0, format="%.8f")
    else:
        lat = default_lat
        lon = default_lon
        st.success(f"Localização capturada: {lat}, {lon}")

    if st.button("Registrar check-in", type="primary", use_container_width=True):
        try:
            if not lat or not lon:
                st.warning("Latitude e longitude não disponíveis.")
            else:
                result = register_checkin(
                    pessoa_id=user["pessoa_id"],
                    unidade_id=vinculo["unidade_id"],
                    lat=lat,
                    lon=lon,
                )

                if result["alerta_gerado"] == "sim":
                    st.warning("Check-in registrado, mas houve alerta.")
                else:
                    st.success("Check-in registrado com sucesso.")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Distância", f"{result['distancia_metros']:.1f} m")
                c2.metric("Raio permitido", f"{result['raio_permitido_metros']} m")
                c3.metric("Status distância", result["status_distancia"])
                c4.metric("Status horário", result["status_horario"])
                st.caption(result.get("mensagem", ""))

        except Exception as exc:
            st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI DASHBOARD
# ============================================================

def render_dashboard(user: dict, vinculo: dict):
    data = load_data()
    respostas = data["RESPOSTAS_CHECKLIST"]

    unidade_id = vinculo["unidade_id"]
    pessoa_id = user["pessoa_id"]
    today = date_str()
    response_map = latest_response_map(respostas, today)

    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Dashboard operacional")
    st.info(f"Resumo do dia: {weekday_pt(now_sp().date())} | {today}")

    checkpoint_options = [
        ("CP0700", "07:00 - Abertura"),
        ("CP1400", "14:00 - Meio do dia"),
        ("CP2000", "20:00 - Fechamento"),
    ]

    rows = []

    for checkpoint_id, checkpoint_label in checkpoint_options:
        for tipo in ["geral", "posicao"]:
            df_items = get_checklist_items(data, pessoa_id, unidade_id, checkpoint_id, tipo)
            total = len(df_items)

            ok = 0
            nok = 0
            na = 0
            pending = 0

            for _, item in df_items.iterrows():
                item_id = str(item.get("item_padrao_id", ""))
                key = item_key(unidade_id, checkpoint_id, tipo, item_id, pessoa_id)
                status = effective_status(response_map.get(key))

                if status == "OK":
                    ok += 1
                elif status == "NAO_OK":
                    nok += 1
                elif status == "NA":
                    na += 1
                else:
                    pending += 1

            pct = int(round((ok / total) * 100, 0)) if total else 0

            rows.append(
                {
                    "checkpoint_id": checkpoint_id,
                    "checkpoint": checkpoint_label,
                    "tipo": "Geral" if tipo == "geral" else "Posição",
                    "total": total,
                    "ok": ok,
                    "nok": nok,
                    "na": na,
                    "pending": pending,
                    "pct": pct,
                }
            )

    cols = st.columns(3)

    for i, row in enumerate(rows):
        bg = metric_palette(row["ok"], row["nok"], row["pending"], row["total"])

        with cols[i % 3]:
            st.markdown(
                f"""
                <div class="metric-card" style="background:{bg};">
                    <div class="metric-title">{row['checkpoint']} | {row['tipo']}</div>
                    <div class="metric-value">{row['pct']}%</div>
                    <div class="metric-sub">
                        OK {row['ok']}/{row['total']} | Não OK {row['nok']} | N/A {row['na']} | Pend. {row['pending']}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI CHECKLIST
# ============================================================

def render_checklist(user: dict, vinculo: dict):
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Minhas atividades")

    unidade_id = vinculo["unidade_id"]
    pessoa_id = user["pessoa_id"]
    today = date_str()

    data = load_data()
    respostas = data["RESPOSTAS_CHECKLIST"]
    response_map = latest_response_map(respostas, today)

    st.info(f"Atividades do dia: {weekday_pt(now_sp().date())} | {today}")

    current_cp, _ = get_current_checkpoint()

    checkpoint_values = {
        "CP0700 | 07:00 - Abertura": "CP0700",
        "CP1400 | 14:00 - Meio do dia": "CP1400",
        "CP2000 | 20:00 - Fechamento": "CP2000",
    }

    labels = list(checkpoint_values.keys())

    default_checkpoint = st.session_state.get("default_checkpoint", current_cp)
    default_label = next(
        (label for label, value in checkpoint_values.items() if value == default_checkpoint),
        labels[0],
    )

    c1, c2 = st.columns([1, 1])

    with c1:
        checkpoint_label = st.selectbox(
            "Escolha o checkpoint",
            labels,
            index=labels.index(default_label),
        )
        checkpoint_id = checkpoint_values[checkpoint_label]

    with c2:
        tipo_label = st.selectbox(
            "Tipo de checklist",
            ["Por posição", "Geral da unidade"],
        )
        tipo = "posicao" if tipo_label == "Por posição" else "geral"

    checklist = get_checklist_items(data, pessoa_id, unidade_id, checkpoint_id, tipo)

    if checklist.empty:
        st.info("Nenhuma atividade encontrada para esta combinação.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    total = len(checklist)
    ok = 0
    nok = 0
    na = 0
    pending = 0

    for _, item in checklist.iterrows():
        item_id = str(item.get("item_padrao_id", ""))
        key = item_key(unidade_id, checkpoint_id, tipo, item_id, pessoa_id)
        status = effective_status(response_map.get(key))

        if status == "OK":
            ok += 1
        elif status == "NAO_OK":
            nok += 1
        elif status == "NA":
            na += 1
        else:
            pending += 1

    pct = int(round((ok / total) * 100, 0)) if total else 0

    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("Progresso", f"{pct}%")
    p2.metric("OK", ok)
    p3.metric("Não OK", nok)
    p4.metric("N/A", na)
    p5.metric("Pendentes", pending)

    if st.button("Atualizar atividades", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.caption("Marque cada atividade. Use observação quando houver algo fora do padrão.")

    for _, item in checklist.iterrows():
        item_id = str(item.get("item_padrao_id", ""))
        item_nome = str(item.get("item", ""))
        detalhe = str(item.get("detalhe", ""))
        servico = str(item.get("servico", ""))
        area = str(item.get("area", ""))
        obrigatorio = str(item.get("obrigatorio", ""))
        evidencia_requerida = str(item.get("evidencia_requerida", ""))

        key = item_key(unidade_id, checkpoint_id, tipo, item_id, pessoa_id)
        latest = response_map.get(key)
        status = effective_status(latest)
        bg, fg, label = card_palette(status)

        obs_key = f"obs_{unidade_id}_{checkpoint_id}_{tipo}_{item_id}"
        evid_key = f"evid_{unidade_id}_{checkpoint_id}_{tipo}_{item_id}"

        st.markdown(
            f"""
            <div class="task-card" style="background:{bg};">
                <div class="task-title">{item_nome}</div>
                <div class="task-detail">{detalhe}</div>
                <div class="task-meta">
                    <b>Serviço/Área:</b> {servico or area or "-"} |
                    <b>Obrigatório:</b> {obrigatorio or "-"} |
                    <b>Evidência:</b> {evidencia_requerida or "-"}
                </div>
                <span class="pill" style="background:{fg}; color:white;">{label}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.expander("Observação / evidência", expanded=False):
            st.text_area("Observação", key=obs_key)
            st.text_input("URL de evidência", key=evid_key)

        b1, b2, b3, b4 = st.columns(4)

        with b1:
            if st.button("OK", key=f"ok_{key}", use_container_width=True):
                register_checklist_response(
                    unidade_id=unidade_id,
                    pessoa_id=pessoa_id,
                    checkpoint_id=checkpoint_id,
                    tipo_checklist=tipo,
                    item_id=item_id,
                    resposta="OK",
                    observacao=st.session_state.get(obs_key, ""),
                    evidencia_url=st.session_state.get(evid_key, ""),
                )
                st.cache_data.clear()
                st.rerun()

        with b2:
            if st.button("Não OK", key=f"nok_{key}", use_container_width=True):
                register_checklist_response(
                    unidade_id=unidade_id,
                    pessoa_id=pessoa_id,
                    checkpoint_id=checkpoint_id,
                    tipo_checklist=tipo,
                    item_id=item_id,
                    resposta="NÃO OK",
                    observacao=st.session_state.get(obs_key, ""),
                    evidencia_url=st.session_state.get(evid_key, ""),
                )
                st.cache_data.clear()
                st.rerun()

        with b3:
            if st.button("N/A", key=f"na_{key}", use_container_width=True):
                register_checklist_response(
                    unidade_id=unidade_id,
                    pessoa_id=pessoa_id,
                    checkpoint_id=checkpoint_id,
                    tipo_checklist=tipo,
                    item_id=item_id,
                    resposta="NÃO APLICÁVEL",
                    observacao=st.session_state.get(obs_key, ""),
                    evidencia_url=st.session_state.get(evid_key, ""),
                )
                st.cache_data.clear()
                st.rerun()

        with b4:
            if st.button("Desmarcar", key=f"rst_{key}", use_container_width=True):
                register_checklist_response(
                    unidade_id=unidade_id,
                    pessoa_id=pessoa_id,
                    checkpoint_id=checkpoint_id,
                    tipo_checklist=tipo,
                    item_id=item_id,
                    resposta="PENDENTE",
                    observacao="Desmarcado pelo usuário",
                    evidencia_url="",
                )
                st.cache_data.clear()
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI ADMIN
# ============================================================

def render_admin(user: dict, vinculo: dict):
    if norm(user.get("perfil")) != "admin":
        st.info("Área disponível apenas para perfil admin.")
        return

    data = load_data()
    unidade_id = vinculo["unidade_id"]

    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Painel admin")

    tab_alertas, tab_checkins, tab_respostas, tab_unidades = st.tabs(
        ["Alertas", "Check-ins", "Respostas", "Unidades"]
    )

    with tab_alertas:
        df = data["ALERTAS"]

        if df.empty:
            st.info("Nenhum alerta registrado.")
        else:
            if "unidade_id" in df.columns:
                df = df[df["unidade_id"].astype(str).map(norm) == norm(unidade_id)]

            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    with tab_checkins:
        df = data["CHECKINS"]

        if df.empty:
            st.info("Nenhum check-in registrado.")
        else:
            if "unidade_id" in df.columns:
                df = df[df["unidade_id"].astype(str).map(norm) == norm(unidade_id)]

            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    with tab_respostas:
        df = data["RESPOSTAS_CHECKLIST"]

        if df.empty:
            st.info("Nenhuma resposta registrada.")
        else:
            if "unidade_id" in df.columns:
                df = df[df["unidade_id"].astype(str).map(norm) == norm(unidade_id)]

            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    with tab_unidades:
        df = data["UNIDADES"]

        if df.empty:
            st.info("Nenhuma unidade cadastrada.")
        else:
            st.dataframe(df, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# MAIN APP
# ============================================================

def render_app():
    user = st.session_state["user"]

    st.sidebar.markdown("## LIIVV")
    st.sidebar.write(f"**{user['nome']}**")
    st.sidebar.caption(f"Perfil: {user.get('perfil', '')}")

    vinculo = select_vinculo(user)

    st.sidebar.divider()

    pages = ["Iniciar trabalho", "Checklist", "Check-in", "Dashboard", "Admin"]

    if "page" not in st.session_state:
        st.session_state["page"] = "Iniciar trabalho"

    page = st.sidebar.radio(
        "Menu",
        pages,
        index=pages.index(st.session_state["page"])
        if st.session_state["page"] in pages
        else 0,
    )

    st.session_state["page"] = page

    st.sidebar.divider()

    if st.sidebar.button("Sair", use_container_width=True):
        logout()

    if page == "Iniciar trabalho":
        render_start_work(user, vinculo)
    elif page == "Checklist":
        render_checklist(user, vinculo)
    elif page == "Check-in":
        render_checkin(user, vinculo)
    elif page == "Dashboard":
        render_dashboard(user, vinculo)
    elif page == "Admin":
        render_admin(user, vinculo)


def main():
    if not st.session_state.get("logged"):
        render_login()
    else:
        render_app()


if __name__ == "__main__":
    main()
