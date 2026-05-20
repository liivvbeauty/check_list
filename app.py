import math
import time
import uuid
import unicodedata
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo

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
    layout="centered",
    initial_sidebar_state="collapsed",
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

    .box {
        background: white;
        border: 1px solid #D7CFC3;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }

    .title {
        color: #0E2A47;
        font-size: 28px;
        font-weight: 900;
        margin-bottom: 4px;
    }

    .subtitle {
        color: #6B7785;
        font-size: 14px;
        margin-bottom: 14px;
    }

    .task {
        background: #F7F4EF;
        border: 1px solid #D7CFC3;
        border-radius: 14px;
        padding: 14px;
        margin-bottom: 12px;
    }

    .task-ok {
        background: #DCFCE7;
        border: 1px solid #86EFAC;
    }

    .task-nok {
        background: #FEE2E2;
        border: 1px solid #FCA5A5;
    }

    .task-na {
        background: #E0E7FF;
        border: 1px solid #A5B4FC;
    }

    .task-title {
        font-weight: 800;
        color: #111827;
        font-size: 16px;
    }

    .task-detail {
        color: #4B5563;
        font-size: 13px;
        margin-top: 5px;
    }

    .status-pill {
        display: inline-block;
        margin-top: 8px;
        padding: 4px 8px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 800;
        background: #111827;
        color: white;
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

def retryable(fn, tries=5, base_sleep=0.7):
    last_error = None

    for i in range(tries):
        try:
            return fn()
        except APIError as exc:
            last_error = exc
            msg = str(exc)

            is_quota = (
                "429" in msg
                or "Quota exceeded" in msg
                or "RESOURCE_EXHAUSTED" in msg
            )

            if not is_quota and i >= 1:
                raise

            time.sleep(base_sleep * (2 ** i))

    raise last_error


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
        st.error("Não foi possível abrir a planilha. Verifique o SPREADSHEET_ID e o compartilhamento com a service account.")
        st.exception(exc)
        st.stop()


@st.cache_data(ttl=30)
def read_sheet(sheet_name: str) -> pd.DataFrame:
    sh = get_spreadsheet()

    try:
        ws = retryable(lambda: sh.worksheet(sheet_name))
        values = retryable(lambda: ws.get_all_records())
        return pd.DataFrame(values)
    except Exception:
        return pd.DataFrame()


def append_row(sheet_name: str, row: list):
    sh = get_spreadsheet()
    ws = retryable(lambda: sh.worksheet(sheet_name))
    retryable(lambda: ws.append_row(row, value_input_option="USER_ENTERED"))
    st.cache_data.clear()


def update_cell_by_key(sheet_name: str, key_col: str, key_value: str, target_col: str, value):
    df = read_sheet(sheet_name)

    if df.empty or key_col not in df.columns or target_col not in df.columns:
        return False

    match = df[df[key_col].astype(str).map(norm) == norm(key_value)]

    if match.empty:
        return False

    row_index = int(match.index[0]) + 2
    col_index = list(df.columns).index(target_col) + 1

    sh = get_spreadsheet()
    ws = retryable(lambda: sh.worksheet(sheet_name))
    retryable(lambda: ws.update_cell(row_index, col_index, value))
    st.cache_data.clear()

    return True


# ============================================================
# HELPERS
# ============================================================

def now_sp():
    return datetime.now(TZ)


def date_str(dt=None):
    return (dt or now_sp()).strftime("%Y-%m-%d")


def time_str(dt=None):
    return (dt or now_sp()).strftime("%H:%M:%S")


def datetime_str(dt=None):
    return (dt or now_sp()).strftime("%Y-%m-%d %H:%M:%S")


def weekday_pt(d: date):
    names = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    return names[d.weekday()]


def strip_accents(text):
    text = str(text or "")
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def norm(value):
    return strip_accents(value).strip().lower()


def id_new(prefix):
    return f"{prefix}-{uuid.uuid4().hex[:10].upper()}"


def to_float(value, default=None):
    try:
        if value is None or value == "":
            return default
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def haversine_meters(lat1, lon1, lat2, lon2):
    radius = 6371000

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


def unidade_aplica(valor_unidade, unidade_id):
    return norm(valor_unidade) in ["todas", norm(unidade_id)]


def get_param(parametros: pd.DataFrame, unidade_id: str, parametro: str, default):
    if parametros.empty:
        return default

    if not {"unidade_id", "parametro", "valor"}.issubset(set(parametros.columns)):
        return default

    df = parametros.copy()
    df["unidade_norm"] = df["unidade_id"].astype(str).map(norm)
    df["param_norm"] = df["parametro"].astype(str).map(norm)

    specific = df[
        (df["unidade_norm"] == norm(unidade_id))
        & (df["param_norm"] == norm(parametro))
    ]

    if not specific.empty and str(specific.iloc[0]["valor"]).strip():
        return specific.iloc[0]["valor"]

    general = df[
        (df["unidade_norm"] == "todas")
        & (df["param_norm"] == norm(parametro))
    ]

    if not general.empty and str(general.iloc[0]["valor"]).strip():
        return general.iloc[0]["valor"]

    return default


def current_checkpoint():
    h = now_sp().hour

    if h < 12:
        return "CP0700", "07:00"

    if h < 18:
        return "CP1400", "14:00"

    return "CP2000", "20:00"


# ============================================================
# LOAD DATA
# ============================================================

def load_data():
    return {
        "UNIDADES": read_sheet("UNIDADES"),
        "POSICOES": read_sheet("POSICOES"),
        "PESSOAS": read_sheet("PESSOAS"),
        "PESSOA_UNIDADE_POSICAO": read_sheet("PESSOA_UNIDADE_POSICAO"),
        "USUARIOS_APP": read_sheet("USUARIOS_APP"),
        "CHECKLIST_GERAL_PADRAO": read_sheet("CHECKLIST_GERAL_PADRAO"),
        "CHECKLIST_POSICAO_PADRAO": read_sheet("CHECKLIST_POSICAO_PADRAO"),
        "AGENDA_ATENDIMENTOS": read_sheet("AGENDA_ATENDIMENTOS"),
        "CHECKINS": read_sheet("CHECKINS"),
        "RESPOSTAS_CHECKLIST": read_sheet("RESPOSTAS_CHECKLIST"),
        "ALERTAS": read_sheet("ALERTAS"),
        "PARAMETROS": read_sheet("PARAMETROS"),
    }


def require_cols(df, cols, sheet):
    missing = [c for c in cols if c not in df.columns]

    if missing:
        st.error(f"A aba {sheet} está sem as colunas: {', '.join(missing)}")
        st.stop()


# ============================================================
# AUTH
# ============================================================

def authenticate(login, senha):
    data = load_data()

    usuarios = data["USUARIOS_APP"]
    pessoas = data["PESSOAS"]
    vinculos = data["PESSOA_UNIDADE_POSICAO"]
    unidades = data["UNIDADES"]
    posicoes = data["POSICOES"]

    require_cols(usuarios, ["usuario_id", "unidade_id", "pessoa_id", "login", "senha", "perfil", "ativa"], "USUARIOS_APP")
    require_cols(pessoas, ["pessoa_id", "nome", "ativa"], "PESSOAS")
    require_cols(vinculos, ["vinculo_id", "unidade_id", "pessoa_id", "posicao_id", "ativa"], "PESSOA_UNIDADE_POSICAO")
    require_cols(unidades, ["unidade_id", "unidade_nome", "ativa"], "UNIDADES")
    require_cols(posicoes, ["posicao_id", "posicao_nome", "ativa"], "POSICOES")

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
    ]

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

        vinculos_out.append(
            {
                "vinculo_id": str(v["vinculo_id"]),
                "unidade_id": unidade_id,
                "unidade_nome": str(unidade.iloc[0].get("unidade_nome", unidade_id)),
                "posicao_id": posicao_id,
                "posicao_nome": str(posicao.iloc[0].get("posicao_nome", posicao_id)) if not posicao.empty else posicao_id,
            }
        )

    if not vinculos_out:
        return None, "Usuário sem unidade ativa."

    update_cell_by_key("USUARIOS_APP", "usuario_id", user_row["usuario_id"], "ultimo_login", datetime_str())

    return {
        "usuario_id": str(user_row["usuario_id"]),
        "pessoa_id": pessoa_id,
        "nome": str(pessoa_row["nome"]),
        "perfil": str(user_row.get("perfil", "operacao")),
        "vinculos": vinculos_out,
    }, None


def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


# ============================================================
# LOOKUPS
# ============================================================

def get_unidade(data, unidade_id):
    df = data["UNIDADES"]
    if df.empty:
        return None

    match = df[df["unidade_id"].astype(str).map(norm) == norm(unidade_id)]
    return None if match.empty else match.iloc[0].to_dict()


def get_pessoa(data, pessoa_id):
    df = data["PESSOAS"]
    if df.empty:
        return None

    match = df[df["pessoa_id"].astype(str).map(norm) == norm(pessoa_id)]
    return None if match.empty else match.iloc[0].to_dict()


def get_posicao(data, posicao_id):
    df = data["POSICOES"]
    if df.empty:
        return None

    match = df[df["posicao_id"].astype(str).map(norm) == norm(posicao_id)]
    return None if match.empty else match.iloc[0].to_dict()


def get_vinculo(data, pessoa_id, unidade_id):
    df = data["PESSOA_UNIDADE_POSICAO"]
    if df.empty:
        return None

    match = df[
        (df["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (df["unidade_id"].astype(str).map(norm) == norm(unidade_id))
        & (df["ativa"].astype(str).map(norm) == "sim")
    ]

    return None if match.empty else match.iloc[0].to_dict()


# ============================================================
# CHECK-IN
# ============================================================

def has_checkin_today(user, vinculo):
    df = read_sheet("CHECKINS")

    if df.empty:
        return False

    if not {"unidade_id", "pessoa_id", "data"}.issubset(set(df.columns)):
        return False

    today = date_str()

    match = df[
        (df["unidade_id"].astype(str).map(norm) == norm(vinculo["unidade_id"]))
        & (df["pessoa_id"].astype(str).map(norm) == norm(user["pessoa_id"]))
        & (df["data"].astype(str).str[:10] == today)
    ]

    return not match.empty


def create_alert(data, unidade_id, tipo, severidade, pessoa_id, nome_pessoa, unidade_nome, agenda_id, mensagem):
    unidade = get_unidade(data, unidade_id)
    email = DEFAULT_ALERT_EMAIL

    if unidade:
        email = unidade.get("email_alerta") or DEFAULT_ALERT_EMAIL

    append_row(
        "ALERTAS",
        [
            unidade_id,
            id_new("ALT"),
            datetime_str(),
            tipo,
            severidade,
            pessoa_id,
            nome_pessoa,
            unidade_nome,
            agenda_id or "",
            mensagem,
            "PENDENTE",
            email,
            "",
            "",
        ],
    )


def evaluate_checkin_time(data, pessoa_id, unidade_id):
    agenda = data["AGENDA_ATENDIMENTOS"]
    parametros = data["PARAMETROS"]

    if agenda.empty:
        return "SEM_AGENDA", "", "Não há atendimento cadastrado para hoje."

    if not {"unidade_id", "agenda_id", "data", "horario_inicio", "pessoa_id", "status"}.issubset(set(agenda.columns)):
        return "SEM_AGENDA", "", "Agenda incompleta."

    minutes = int(
        to_float(
            get_param(parametros, unidade_id, "CHECKIN_ANTECEDENCIA_MINUTOS", DEFAULT_CHECKIN_MINUTES),
            DEFAULT_CHECKIN_MINUTES,
        )
    )

    today = date_str()

    df = agenda[
        (agenda["unidade_id"].astype(str).map(norm) == norm(unidade_id))
        & (agenda["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (agenda["data"].astype(str).str[:10] == today)
        & (agenda["status"].astype(str).map(norm) != "cancelado")
    ].copy()

    if df.empty:
        return "SEM_AGENDA", "", "Não há atendimento cadastrado para hoje."

    df = df.sort_values("horario_inicio")
    first = df.iloc[0].to_dict()
    horario_inicio = str(first["horario_inicio"])[:5]

    try:
        start = datetime.strptime(f"{today} {horario_inicio}", "%Y-%m-%d %H:%M").replace(tzinfo=TZ)
    except Exception:
        return "ERRO_AGENDA", str(first.get("agenda_id", "")), "Horário de agenda inválido."

    limit = start - timedelta(minutes=minutes)

    if now_sp() <= limit:
        return "OK", str(first.get("agenda_id", "")), "Check-in dentro do prazo."

    return "ATRASADO", str(first.get("agenda_id", "")), f"Check-in realizado após o limite de {minutes} minutos antes do primeiro atendimento."


def register_checkin(user, vinculo, lat, lon):
    data = load_data()

    pessoa = get_pessoa(data, user["pessoa_id"])
    unidade = get_unidade(data, vinculo["unidade_id"])

    if not pessoa:
        raise ValueError("Pessoa não encontrada.")

    if not unidade:
        raise ValueError("Unidade não encontrada.")

    unidade_lat = to_float(unidade.get("latitude"))
    unidade_lon = to_float(unidade.get("longitude"))

    if unidade_lat is None or unidade_lon is None:
        raise ValueError("A unidade não possui latitude e longitude cadastradas.")

    pessoa_lat = to_float(lat)
    pessoa_lon = to_float(lon)

    if pessoa_lat is None or pessoa_lon is None:
        raise ValueError("Não foi possível obter a localização do telefone.")

    raio = to_float(unidade.get("raio_permitido_metros"), DEFAULT_DISTANCE_METERS)
    distancia = haversine_meters(pessoa_lat, pessoa_lon, unidade_lat, unidade_lon)

    status_distancia = "OK" if distancia <= raio else "FORA_DO_RAIO"
    status_horario, agenda_id, msg_horario = evaluate_checkin_time(data, user["pessoa_id"], vinculo["unidade_id"])

    alerta = "não"

    if status_distancia != "OK":
        alerta = "sim"
        create_alert(
            data,
            vinculo["unidade_id"],
            "CHECKIN_FORA_DO_RAIO",
            "ALTA",
            user["pessoa_id"],
            user["nome"],
            vinculo["unidade_nome"],
            agenda_id,
            f"Check-in fora do raio permitido. Distância: {distancia:.1f}m. Limite: {raio:.1f}m.",
        )

    if status_horario == "ATRASADO":
        alerta = "sim"
        create_alert(
            data,
            vinculo["unidade_id"],
            "CHECKIN_ATRASADO",
            "MEDIA",
            user["pessoa_id"],
            user["nome"],
            vinculo["unidade_nome"],
            agenda_id,
            msg_horario,
        )

    append_row(
        "CHECKINS",
        [
            vinculo["unidade_id"],
            id_new("CHK"),
            datetime_str(),
            date_str(),
            time_str(),
            user["pessoa_id"],
            user["nome"],
            vinculo["unidade_nome"],
            vinculo["posicao_id"],
            vinculo["posicao_nome"],
            pessoa_lat,
            pessoa_lon,
            unidade_lat,
            unidade_lon,
            round(distancia, 1),
            raio,
            status_distancia,
            status_horario,
            alerta,
            agenda_id,
            msg_horario,
        ],
    )

    return {
        "distancia": distancia,
        "raio": raio,
        "status_distancia": status_distancia,
        "status_horario": status_horario,
        "alerta": alerta,
        "mensagem": msg_horario,
    }


# ============================================================
# CHECKLIST
# ============================================================

def response_status(value):
    s = norm(value).replace(" ", "_")

    if s == "ok":
        return "OK"

    if s in ["nao_ok", "não_ok"]:
        return "NOK"

    if s in ["nao_aplicavel", "não_aplicavel", "n/a"]:
        return "NA"

    return "PENDENTE"


def latest_responses(day):
    df = read_sheet("RESPOSTAS_CHECKLIST")

    if df.empty:
        return {}

    required = {"unidade_id", "data", "checkpoint_id", "tipo_checklist", "item_id", "pessoa_id", "resposta", "timestamp"}

    if not required.issubset(set(df.columns)):
        return {}

    df = df[df["data"].astype(str).str[:10] == day].copy()

    if df.empty:
        return {}

    df["ts"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["ts"]).sort_values("ts")

    latest = df.groupby(
        ["unidade_id", "checkpoint_id", "tipo_checklist", "item_id", "pessoa_id"],
        as_index=False,
    ).tail(1)

    out = {}

    for _, row in latest.iterrows():
        key = (
            norm(row["unidade_id"]),
            norm(row["checkpoint_id"]),
            norm(row["tipo_checklist"]),
            norm(row["item_id"]),
            norm(row["pessoa_id"]),
        )
        out[key] = row.to_dict()

    return out


def get_items(data, user, vinculo, checkpoint_id, tipo):
    if tipo == "geral":
        df = data["CHECKLIST_GERAL_PADRAO"]

        if df.empty:
            return pd.DataFrame()

        return df[
            df["unidade_id"].apply(lambda x: unidade_aplica(x, vinculo["unidade_id"]))
            & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint_id))
            & (df["ativo"].astype(str).map(norm) == "sim")
        ].copy()

    df = data["CHECKLIST_POSICAO_PADRAO"]

    if df.empty:
        return pd.DataFrame()

    return df[
        df["unidade_id"].apply(lambda x: unidade_aplica(x, vinculo["unidade_id"]))
        & (df["posicao_id"].astype(str).map(norm) == norm(vinculo["posicao_id"]))
        & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint_id))
        & (df["ativo"].astype(str).map(norm) == "sim")
    ].copy()


def get_item_name(data, item_id, unidade_id):
    for sheet in ["CHECKLIST_GERAL_PADRAO", "CHECKLIST_POSICAO_PADRAO"]:
        df = data[sheet]

        if df.empty:
            continue

        match = df[
            (df["item_padrao_id"].astype(str).map(norm) == norm(item_id))
            & df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
        ]

        if not match.empty:
            return str(match.iloc[0].get("item", item_id))

    return item_id


def save_response(user, vinculo, checkpoint_id, tipo, item_id, resposta):
    data = load_data()
    item_name = get_item_name(data, item_id, vinculo["unidade_id"])

    status = "CONFORME"
    alerta = "não"

    if response_status(resposta) == "NOK":
        status = "NÃO CONFORME"
        alerta = "sim"
        create_alert(
            data,
            vinculo["unidade_id"],
            "CHECKLIST_NAO_CONFORME",
            "MEDIA",
            user["pessoa_id"],
            user["nome"],
            vinculo["unidade_nome"],
            "",
            f"Item não conforme: {item_name}",
        )

    if response_status(resposta) == "NA":
        status = "NÃO APLICÁVEL"

    append_row(
        "RESPOSTAS_CHECKLIST",
        [
            vinculo["unidade_id"],
            id_new("RSP"),
            datetime_str(),
            date_str(),
            checkpoint_id,
            tipo,
            user["pessoa_id"],
            user["nome"],
            vinculo["unidade_nome"],
            vinculo["posicao_id"],
            vinculo["posicao_nome"],
            item_id,
            item_name,
            resposta,
            "",
            "",
            status,
            alerta,
        ],
    )


def progress_summary(user, vinculo):
    data = load_data()
    day = date_str()
    responses = latest_responses(day)

    total = ok = nok = na = pending = 0

    for checkpoint_id in ["CP0700", "CP1400", "CP2000"]:
        for tipo in ["posicao", "geral"]:
            items = get_items(data, user, vinculo, checkpoint_id, tipo)

            for _, item in items.iterrows():
                item_id = str(item.get("item_padrao_id", ""))
                key = (
                    norm(vinculo["unidade_id"]),
                    norm(checkpoint_id),
                    norm(tipo),
                    norm(item_id),
                    norm(user["pessoa_id"]),
                )

                status = response_status(responses.get(key, {}).get("resposta", ""))

                total += 1

                if status == "OK":
                    ok += 1
                elif status == "NOK":
                    nok += 1
                elif status == "NA":
                    na += 1
                else:
                    pending += 1

    pct = round((ok / total) * 100) if total else 0

    return {
        "total": total,
        "ok": ok,
        "nok": nok,
        "na": na,
        "pending": pending,
        "pct": pct,
    }


# ============================================================
# UI
# ============================================================

def render_login():
    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.markdown('<div class="title">LIIVV Checklist</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Entre para iniciar seu trabalho.</div>', unsafe_allow_html=True)

    with st.form("login"):
        login = st.text_input("Login")
        senha = st.text_input("Senha", type="password")
        submitted = st.form_submit_button("Entrar", use_container_width=True)

    if submitted:
        user, error = authenticate(login, senha)

        if error:
            st.error(error)
        else:
            st.session_state["logged"] = True
            st.session_state["user"] = user
            st.session_state["page"] = "home"
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def get_selected_vinculo(user):
    vinculos = user.get("vinculos", [])

    if not vinculos:
        st.error("Usuário sem unidade ativa.")
        st.stop()

    if len(vinculos) == 1:
        return vinculos[0]

    label_map = {
        f"{v['unidade_nome']} | {v['posicao_nome']}": v
        for v in vinculos
    }

    label = st.selectbox("Escolha a unidade", list(label_map.keys()))
    return label_map[label]


def nav_button(label, page, primary=False):
    if st.button(label, type="primary" if primary else "secondary", use_container_width=True):
        st.session_state["page"] = page
        st.rerun()


def render_header(user, vinculo):
    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.markdown(f'<div class="title">Olá, {user["nome"]}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtitle">{vinculo["unidade_nome"]} | {vinculo["posicao_nome"]}</div>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)


def render_home(user, vinculo):
    render_header(user, vinculo)

    checked = has_checkin_today(user, vinculo)
    cp_id, cp_time = current_checkpoint()
    summary = progress_summary(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)

    if checked:
        st.success("Check-in realizado hoje.")
    else:
        st.warning("Faça o check-in antes de iniciar as atividades.")

    st.write(f"**Checkpoint sugerido:** {cp_time}")
    st.write(f"**Progresso do dia:** {summary['pct']}%")
    st.progress(summary["pct"] / 100 if summary["pct"] else 0)

    st.markdown("</div>", unsafe_allow_html=True)

    if not checked:
        nav_button("Fazer check-in", "checkin", primary=True)
    else:
        nav_button("Abrir checklist", "checklist", primary=True)

    nav_button("Ver resumo", "dashboard")

    if norm(user.get("perfil")) == "admin":
        nav_button("Admin", "admin")

    if st.button("Sair", use_container_width=True):
        logout()


def render_checkin(user, vinculo):
    render_header(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Check-in")

    if has_checkin_today(user, vinculo):
        st.success("Você já realizou check-in hoje.")

    st.write("Toque no botão abaixo e permita o acesso à localização do telefone.")

    if get_geolocation is None:
        st.error("Erro: o componente de localização não está disponível no app.")
        st.markdown("</div>", unsafe_allow_html=True)
        nav_button("Voltar", "home")
        return

    location = None

    try:
        location = get_geolocation()
    except Exception:
        location = None

    if st.button("Registrar localização", type="primary", use_container_width=True):
        if not location or not isinstance(location, dict) or not location.get("coords"):
            st.error("Erro ao obter localização. Permita o acesso ao GPS do telefone e tente novamente.")
        else:
            coords = location["coords"]
            lat = coords.get("latitude")
            lon = coords.get("longitude")

            if lat is None or lon is None:
                st.error("Erro ao obter localização. Tente novamente pelo telefone.")
            else:
                try:
                    result = register_checkin(user, vinculo, lat, lon)

                    if result["alerta"] == "sim":
                        st.warning("Check-in registrado com alerta.")
                    else:
                        st.success("Check-in registrado com sucesso.")

                    st.write(f"Distância calculada: **{result['distancia']:.1f} m**")
                    st.write(f"Status: **{result['status_distancia']}**")
                    st.caption(result.get("mensagem", ""))

                    st.session_state["page"] = "checklist"
                    st.rerun()

                except Exception as exc:
                    st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)
    nav_button("Voltar", "home")


def render_checklist(user, vinculo):
    render_header(user, vinculo)

    data = load_data()
    cp_current, _ = current_checkpoint()

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Checklist")

    cp_label = st.selectbox(
        "Horário",
        {
            "07:00": "CP0700",
            "14:00": "CP1400",
            "20:00": "CP2000",
        }.keys(),
        index=["CP0700", "CP1400", "CP2000"].index(cp_current),
    )

    cp_map = {
        "07:00": "CP0700",
        "14:00": "CP1400",
        "20:00": "CP2000",
    }

    checkpoint_id = cp_map[cp_label]

    tipo_label = st.selectbox("Lista", ["Minha função", "Geral da unidade"])
    tipo = "posicao" if tipo_label == "Minha função" else "geral"

    items = get_items(data, user, vinculo, checkpoint_id, tipo)
    responses = latest_responses(date_str())

    if items.empty:
        st.info("Nenhuma atividade encontrada.")
        st.markdown("</div>", unsafe_allow_html=True)
        nav_button("Voltar", "home")
        return

    st.markdown("</div>", unsafe_allow_html=True)

    for _, item in items.iterrows():
        item_id = str(item.get("item_padrao_id", ""))
        item_title = str(item.get("item", ""))
        item_detail = str(item.get("detalhe", ""))

        key = (
            norm(vinculo["unidade_id"]),
            norm(checkpoint_id),
            norm(tipo),
            norm(item_id),
            norm(user["pessoa_id"]),
        )

        status = response_status(responses.get(key, {}).get("resposta", ""))

        css_class = "task"
        label = "Pendente"

        if status == "OK":
            css_class = "task task-ok"
            label = "OK"
        elif status == "NOK":
            css_class = "task task-nok"
            label = "Não OK"
        elif status == "NA":
            css_class = "task task-na"
            label = "N/A"

        st.markdown(
            f"""
            <div class="{css_class}">
                <div class="task-title">{item_title}</div>
                <div class="task-detail">{item_detail}</div>
                <span class="status-pill">{label}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button("OK", key=f"ok_{checkpoint_id}_{tipo}_{item_id}", use_container_width=True):
                save_response(user, vinculo, checkpoint_id, tipo, item_id, "OK")
                st.rerun()

        with c2:
            if st.button("Não OK", key=f"nok_{checkpoint_id}_{tipo}_{item_id}", use_container_width=True):
                save_response(user, vinculo, checkpoint_id, tipo, item_id, "NÃO OK")
                st.rerun()

        with c3:
            if st.button("N/A", key=f"na_{checkpoint_id}_{tipo}_{item_id}", use_container_width=True):
                save_response(user, vinculo, checkpoint_id, tipo, item_id, "NÃO APLICÁVEL")
                st.rerun()

    nav_button("Voltar", "home")


def render_dashboard(user, vinculo):
    render_header(user, vinculo)

    summary = progress_summary(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Resumo do dia")
    st.write(f"{weekday_pt(now_sp().date())}, {date_str()}")

    st.metric("Progresso", f"{summary['pct']}%")
    st.progress(summary["pct"] / 100 if summary["pct"] else 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("OK", summary["ok"])
    c2.metric("Não OK", summary["nok"])
    c3.metric("N/A", summary["na"])
    c4.metric("Pendentes", summary["pending"])

    st.markdown("</div>", unsafe_allow_html=True)

    nav_button("Voltar", "home")


def render_admin(user, vinculo):
    if norm(user.get("perfil")) != "admin":
        st.error("Área restrita.")
        nav_button("Voltar", "home")
        return

    render_header(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Admin")

    option = st.selectbox("Visualizar", ["Alertas", "Check-ins", "Respostas"])

    sheet_map = {
        "Alertas": "ALERTAS",
        "Check-ins": "CHECKINS",
        "Respostas": "RESPOSTAS_CHECKLIST",
    }

    df = read_sheet(sheet_map[option])

    if df.empty:
        st.info("Sem registros.")
    else:
        if "unidade_id" in df.columns:
            df = df[df["unidade_id"].astype(str).map(norm) == norm(vinculo["unidade_id"])]

        if "timestamp" in df.columns:
            df = df.sort_values("timestamp", ascending=False)

        st.dataframe(df, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    nav_button("Voltar", "home")


def render_app():
    user = st.session_state["user"]
    vinculo = get_selected_vinculo(user)

    page = st.session_state.get("page", "home")

    if page == "home":
        render_home(user, vinculo)
    elif page == "checkin":
        render_checkin(user, vinculo)
    elif page == "checklist":
        render_checklist(user, vinculo)
    elif page == "dashboard":
        render_dashboard(user, vinculo)
    elif page == "admin":
        render_admin(user, vinculo)
    else:
        st.session_state["page"] = "home"
        st.rerun()


def main():
    if not st.session_state.get("logged"):
        render_login()
    else:
        render_app()


if __name__ == "__main__":
    main()
