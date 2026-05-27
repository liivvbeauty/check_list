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


TZ = ZoneInfo("America/Sao_Paulo")
SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
APP_TITLE = st.secrets.get("APP_TITLE", "LIIVV Checklist")
DEFAULT_DISTANCE_METERS = 200
DEFAULT_CHECKIN_MINUTES = 30
DEFAULT_ALERT_EMAIL = "operacao@liivv.com.br"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.stApp { background: #EFE7DD; }
[data-testid="stHeader"] { background: rgba(239,231,221,0.95); }
.block-container { padding-top: 2rem; max-width: 1180px; }

.box {
    background: white;
    border-radius: 18px;
    padding: 22px;
    border: 1px solid #D7CFC3;
    margin-bottom: 16px;
}

.title {
    font-size: 32px;
    font-weight: 900;
    color: #0E2A47;
}

.subtitle {
    color: #6B7280;
    margin-top: 6px;
}

.card {
    background: #0E2A47;
    border-radius: 22px;
    padding: 22px;
    color: white;
    min-height: 210px;
    margin-bottom: 18px;
}

.card-green { background: #166534; }
.card-yellow { background: #92400E; }
.card-red { background: #7F1D1D; }

.metric-big {
    font-size: 54px;
    font-weight: 900;
    line-height: 1;
    margin-top: 12px;
}

.metric-label {
    font-size: 16px;
    font-weight: 800;
    opacity: 0.95;
}

.metric-detail {
    margin-top: 16px;
    line-height: 1.8;
    font-size: 14px;
}

.task-box {
    background: white;
    border-radius: 18px;
    border: 1px solid #D7CFC3;
    padding: 18px;
    margin-bottom: 14px;
}

.task-ok { background: #DCFCE7; border-color: #86EFAC; }
.task-nok { background: #FEE2E2; border-color: #FCA5A5; }
.task-na { background: #E0E7FF; border-color: #A5B4FC; }

.task-title {
    font-size: 17px;
    font-weight: 800;
    color: #111827;
}

.task-detail {
    font-size: 13px;
    color: #6B7280;
    margin-top: 6px;
}

.login-box {
    max-width: 520px;
    margin: auto;
    margin-top: 70px;
}

.stButton button {
    border-radius: 12px;
    font-weight: 700;
    height: 46px;
}
</style>
""", unsafe_allow_html=True)


def now_sp():
    return datetime.now(TZ)


def date_str():
    return now_sp().strftime("%Y-%m-%d")


def time_str():
    return now_sp().strftime("%H:%M:%S")


def datetime_str():
    return now_sp().strftime("%Y-%m-%d %H:%M:%S")


def weekday_pt(d: date):
    return ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][d.weekday()]


def strip_accents(text):
    text = str(text or "")
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def norm(value):
    return strip_accents(value).strip().lower()


def to_float(value, default=None):
    try:
        if value is None or value == "":
            return default
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def new_id(prefix):
    return f"{prefix}-{uuid.uuid4().hex[:10].upper()}"


def response_status(value):
    s = norm(value).replace("_", " ")

    if s == "ok":
        return "OK"
    if s in ["nao ok", "não ok", "nok"]:
        return "NOK"
    if s in ["n/a", "na", "nao aplicavel", "não aplicável", "nao aplicável"]:
        return "NA"

    return "PENDENTE"


def retryable(fn, tries=5):
    last_error = None

    for i in range(tries):
        try:
            return fn()
        except APIError as exc:
            last_error = exc
            time.sleep(0.8 * (2 ** i))

    raise last_error


@st.cache_resource
def get_client():
    if "google_service_account" in st.secrets:
        creds_dict = dict(st.secrets["google_service_account"])
    elif "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    else:
        st.error("Secrets sem [google_service_account] ou [gcp_service_account].")
        st.stop()

    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    if not SPREADSHEET_ID:
        st.error("SPREADSHEET_ID não configurado.")
        st.stop()

    return retryable(lambda: get_client().open_by_key(SPREADSHEET_ID))


@st.cache_data(ttl=30)
def read_sheet(sheet_name):
    try:
        sh = get_spreadsheet()
        ws = retryable(lambda: sh.worksheet(sheet_name))
        values = retryable(lambda: ws.get_all_records())
        return pd.DataFrame(values)
    except Exception:
        return pd.DataFrame()


def append_row(sheet_name, row):
    sh = get_spreadsheet()
    ws = retryable(lambda: sh.worksheet(sheet_name))
    retryable(lambda: ws.append_row(row, value_input_option="USER_ENTERED"))
    st.cache_data.clear()


def update_cell_by_key(sheet_name, key_col, key_value, target_col, value):
    df = read_sheet(sheet_name)

    if df.empty or key_col not in df.columns or target_col not in df.columns:
        return False

    match = df[df[key_col].astype(str).map(norm) == norm(key_value)]

    if match.empty:
        return False

    row_number = int(match.index[0]) + 2
    col_number = list(df.columns).index(target_col) + 1

    sh = get_spreadsheet()
    ws = retryable(lambda: sh.worksheet(sheet_name))
    retryable(lambda: ws.update_cell(row_number, col_number, value))
    st.cache_data.clear()

    return True


def load_data():
    return {
        "UNIDADES": read_sheet("UNIDADES"),
        "POSICOES": read_sheet("POSICOES"),
        "PESSOAS": read_sheet("PESSOAS"),
        "PESSOA_UNIDADE_POSICAO": read_sheet("PESSOA_UNIDADE_POSICAO"),
        "USUARIOS_APP": read_sheet("USUARIOS_APP"),
        "CHECKLIST_GERAL_PADRAO": read_sheet("CHECKLIST_GERAL_PADRAO"),
        "CHECKLIST_POSICAO_PADRAO": read_sheet("CHECKLIST_POSICAO_PADRAO"),
        "RESPOSTAS_CHECKLIST": read_sheet("RESPOSTAS_CHECKLIST"),
        "CHECKINS": read_sheet("CHECKINS"),
        "ALERTAS": read_sheet("ALERTAS"),
        "AGENDA_ATENDIMENTOS": read_sheet("AGENDA_ATENDIMENTOS"),
        "PARAMETROS": read_sheet("PARAMETROS"),
    }


def find_row(df, col, value):
    if df.empty or col not in df.columns:
        return None

    match = df[df[col].astype(str).map(norm) == norm(value)]
    return None if match.empty else match.iloc[0].to_dict()


def unidade_aplica(valor_unidade, unidade_id):
    return norm(valor_unidade) in ["todas", norm(unidade_id)]


def haversine_meters(lat1, lon1, lat2, lon2):
    r = 6371000
    p1 = math.radians(float(lat1))
    p2 = math.radians(float(lat2))
    dp = math.radians(float(lat2) - float(lat1))
    dl = math.radians(float(lon2) - float(lon1))

    a = (
        math.sin(dp / 2) ** 2
        + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    )

    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return r * c


def get_param(data, unidade_id, param_name, default):
    df = data["PARAMETROS"]

    if df.empty or not {"unidade_id", "parametro", "valor"}.issubset(set(df.columns)):
        return default

    specific = df[
        (df["unidade_id"].astype(str).map(norm) == norm(unidade_id))
        & (df["parametro"].astype(str).map(norm) == norm(param_name))
    ]

    if not specific.empty and str(specific.iloc[0].get("valor", "")).strip():
        return specific.iloc[0].get("valor")

    general = df[
        (df["unidade_id"].astype(str).map(norm) == "todas")
        & (df["parametro"].astype(str).map(norm) == norm(param_name))
    ]

    if not general.empty and str(general.iloc[0].get("valor", "")).strip():
        return general.iloc[0].get("valor")

    return default


def current_checkpoint():
    h = now_sp().hour

    if h < 12:
        return "CP0700", "07:00"
    if h < 18:
        return "CP1400", "14:00"

    return "CP2000", "20:00"


def authenticate(login, senha):
    data = load_data()

    usuarios = data["USUARIOS_APP"]
    pessoas = data["PESSOAS"]
    unidades = data["UNIDADES"]
    posicoes = data["POSICOES"]
    vinculos = data["PESSOA_UNIDADE_POSICAO"]

    if usuarios.empty:
        return None, "Aba USUARIOS_APP vazia."

    user_match = usuarios[
        (usuarios["login"].astype(str).map(norm) == norm(login))
        & (usuarios["senha"].astype(str) == str(senha))
        & (usuarios["ativa"].astype(str).map(norm) == "sim")
    ]

    if user_match.empty:
        return None, "Login ou senha inválidos."

    user_row = user_match.iloc[0].to_dict()

    pessoa_id = str(user_row.get("pessoa_id", ""))
    unidade_id_user = str(user_row.get("unidade_id", ""))

    pessoa = find_row(pessoas, "pessoa_id", pessoa_id) or {}
    nome = (
        pessoa.get("nome")
        or user_row.get("nome")
        or user_row.get("pessoa_nome")
        or user_row.get("usuario_nome")
        or "Colaboradora"
    )

    vinculos_df = vinculos[
        (vinculos["pessoa_id"].astype(str).map(norm) == norm(pessoa_id))
        & (vinculos["ativa"].astype(str).map(norm) == "sim")
    ] if not vinculos.empty else pd.DataFrame()

    if unidade_id_user and not vinculos_df.empty:
        vinculos_df = vinculos_df[
            vinculos_df["unidade_id"].astype(str).map(norm) == norm(unidade_id_user)
        ]

    vinculos_out = []

    if not vinculos_df.empty:
        for _, v in vinculos_df.iterrows():
            unidade_id = str(v.get("unidade_id", ""))
            posicao_id = str(v.get("posicao_id", ""))

            unidade = find_row(unidades, "unidade_id", unidade_id) or {}
            posicao = find_row(posicoes, "posicao_id", posicao_id) or {}

            vinculos_out.append({
                "unidade_id": unidade_id,
                "unidade_nome": unidade.get("unidade_nome") or unidade_id,
                "posicao_id": posicao_id,
                "posicao_nome": posicao.get("posicao_nome") or posicao_id,
            })

    if not vinculos_out:
        unidade = find_row(unidades, "unidade_id", unidade_id_user) or {}
        posicao_id = str(user_row.get("posicao_id", ""))
        posicao = find_row(posicoes, "posicao_id", posicao_id) or {}

        vinculos_out.append({
            "unidade_id": unidade_id_user,
            "unidade_nome": user_row.get("unidade_nome") or unidade.get("unidade_nome") or unidade_id_user,
            "posicao_id": posicao_id,
            "posicao_nome": user_row.get("posicao_nome") or posicao.get("posicao_nome") or posicao_id,
        })

    update_cell_by_key(
        "USUARIOS_APP",
        "usuario_id",
        user_row.get("usuario_id", ""),
        "ultimo_login",
        datetime_str(),
    )

    return {
        "usuario_id": str(user_row.get("usuario_id", "")),
        "pessoa_id": pessoa_id,
        "nome": str(nome),
        "perfil": str(user_row.get("perfil", "operacao")),
        "vinculos": vinculos_out,
    }, None


def selected_vinculo(user):
    vinculos = user.get("vinculos", [])

    if not vinculos:
        st.error("Usuário sem unidade ativa.")
        st.stop()

    if len(vinculos) == 1:
        return vinculos[0]

    options = {
        f"{v['unidade_nome']} | {v['posicao_nome']}": v
        for v in vinculos
    }

    label = st.selectbox("Unidade", list(options.keys()))
    return options[label]


def has_checkin_today(user, vinculo):
    df = read_sheet("CHECKINS")

    if df.empty or not {"unidade_id", "pessoa_id", "data"}.issubset(set(df.columns)):
        return False

    match = df[
        (df["unidade_id"].astype(str).map(norm) == norm(vinculo["unidade_id"]))
        & (df["pessoa_id"].astype(str).map(norm) == norm(user["pessoa_id"]))
        & (df["data"].astype(str).str[:10] == date_str())
    ]

    return not match.empty


def create_alert(data, unidade_id, tipo, severidade, pessoa_id, nome, unidade_nome, agenda_id, mensagem):
    unidade = find_row(data["UNIDADES"], "unidade_id", unidade_id) or {}
    email = unidade.get("email_alerta") or DEFAULT_ALERT_EMAIL

    append_row("ALERTAS", [
        unidade_id,
        new_id("ALT"),
        datetime_str(),
        tipo,
        severidade,
        pessoa_id,
        nome,
        unidade_nome,
        agenda_id or "",
        mensagem,
        "PENDENTE",
        email,
        "",
        "",
    ])


def evaluate_checkin_time(data, user, vinculo):
    agenda = data["AGENDA_ATENDIMENTOS"]

    if agenda.empty:
        return "SEM_AGENDA", "", "Sem agenda cadastrada para hoje."

    required = {"unidade_id", "agenda_id", "data", "horario_inicio", "pessoa_id", "status"}

    if not required.issubset(set(agenda.columns)):
        return "SEM_AGENDA", "", "Agenda incompleta."

    minutes = int(to_float(
        get_param(data, vinculo["unidade_id"], "CHECKIN_ANTECEDENCIA_MINUTOS", DEFAULT_CHECKIN_MINUTES),
        DEFAULT_CHECKIN_MINUTES
    ))

    today = date_str()

    df = agenda[
        (agenda["unidade_id"].astype(str).map(norm) == norm(vinculo["unidade_id"]))
        & (agenda["pessoa_id"].astype(str).map(norm) == norm(user["pessoa_id"]))
        & (agenda["data"].astype(str).str[:10] == today)
        & (agenda["status"].astype(str).map(norm) != "cancelado")
    ].copy()

    if df.empty:
        return "SEM_AGENDA", "", "Sem agenda cadastrada para hoje."

    df = df.sort_values("horario_inicio")
    first = df.iloc[0].to_dict()
    agenda_id = str(first.get("agenda_id", ""))

    try:
        start = datetime.strptime(
            f"{today} {str(first['horario_inicio'])[:5]}",
            "%Y-%m-%d %H:%M"
        ).replace(tzinfo=TZ)
    except Exception:
        return "ERRO_AGENDA", agenda_id, "Horário de agenda inválido."

    limit = start - timedelta(minutes=minutes)

    if now_sp() <= limit:
        return "OK", agenda_id, "Check-in dentro do prazo."

    return "ATRASADO", agenda_id, f"Check-in após o limite de {minutes} minutos antes do primeiro atendimento."


def register_checkin(user, vinculo, lat, lon):
    data = load_data()
    unidade = find_row(data["UNIDADES"], "unidade_id", vinculo["unidade_id"]) or {}

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
    status_horario, agenda_id, msg = evaluate_checkin_time(data, user, vinculo)

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
            f"Check-in fora do raio. Distância: {distancia:.1f}m. Limite: {raio:.1f}m.",
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
            msg,
        )

    append_row("CHECKINS", [
        vinculo["unidade_id"],
        new_id("CHK"),
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
        msg,
    ])

    return {
        "distancia": distancia,
        "raio": raio,
        "status_distancia": status_distancia,
        "status_horario": status_horario,
        "alerta": alerta,
        "mensagem": msg,
    }


def latest_responses():
    df = read_sheet("RESPOSTAS_CHECKLIST")

    if df.empty:
        return {}

    required = {"unidade_id", "data", "checkpoint_id", "tipo_checklist", "item_id", "pessoa_id", "resposta", "timestamp"}

    if not required.issubset(set(df.columns)):
        return {}

    df = df[df["data"].astype(str).str[:10] == date_str()].copy()

    if df.empty:
        return {}

    df["ts"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["ts"]).sort_values("ts")

    latest = df.groupby(
        ["unidade_id", "checkpoint_id", "tipo_checklist", "item_id", "pessoa_id"],
        as_index=False
    ).tail(1)

    out = {}

    for _, r in latest.iterrows():
        key = (
            norm(r["unidade_id"]),
            norm(r["checkpoint_id"]),
            norm(r["tipo_checklist"]),
            norm(r["item_id"]),
            norm(r["pessoa_id"]),
        )
        out[key] = r.to_dict()

    return out


def get_items(user, vinculo, tipo, checkpoint):
    data = load_data()

    if tipo == "geral":
        df = data["CHECKLIST_GERAL_PADRAO"]

        if df.empty:
            return pd.DataFrame()

        return df[
            df["unidade_id"].apply(lambda x: unidade_aplica(x, vinculo["unidade_id"]))
            & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
            & (df["ativo"].astype(str).map(norm) == "sim")
        ].copy()

    df = data["CHECKLIST_POSICAO_PADRAO"]

    if df.empty:
        return pd.DataFrame()

    return df[
        df["unidade_id"].apply(lambda x: unidade_aplica(x, vinculo["unidade_id"]))
        & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
        & (df["posicao_id"].astype(str).map(norm) == norm(vinculo["posicao_id"]))
        & (df["ativo"].astype(str).map(norm) == "sim")
    ].copy()


def get_item_name(item_id, unidade_id):
    data = load_data()

    for sheet in ["CHECKLIST_GERAL_PADRAO", "CHECKLIST_POSICAO_PADRAO"]:
        df = data[sheet]

        if df.empty:
            continue

        found = df[
            (df["item_padrao_id"].astype(str).map(norm) == norm(item_id))
            & df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
        ]

        if not found.empty:
            return str(found.iloc[0].get("item", item_id))

    return item_id


def save_response(user, vinculo, checkpoint, tipo, item_id, resposta):
    data = load_data()
    item_name = get_item_name(item_id, vinculo["unidade_id"])

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

    append_row("RESPOSTAS_CHECKLIST", [
        vinculo["unidade_id"],
        new_id("RSP"),
        datetime_str(),
        date_str(),
        checkpoint,
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
    ])


def progress_unit(unidade_id):
    data = load_data()
    responses = latest_responses()

    total = ok = nok = na = pending = 0

    for checkpoint in ["CP0700", "CP1400", "CP2000"]:
        for tipo, sheet in [
            ("geral", "CHECKLIST_GERAL_PADRAO"),
            ("posicao", "CHECKLIST_POSICAO_PADRAO"),
        ]:
            df = data[sheet]

            if df.empty:
                continue

            df = df[
                df["unidade_id"].apply(lambda x: unidade_aplica(x, unidade_id))
                & (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
                & (df["ativo"].astype(str).map(norm) == "sim")
            ]

            for _, item in df.iterrows():
                total += 1
                item_id = str(item.get("item_padrao_id", ""))

                found_status = "PENDENTE"

                for key, value in responses.items():
                    if (
                        key[0] == norm(unidade_id)
                        and key[1] == norm(checkpoint)
                        and key[2] == norm(tipo)
                        and key[3] == norm(item_id)
                    ):
                        found_status = response_status(value.get("resposta", ""))

                if found_status == "OK":
                    ok += 1
                elif found_status == "NOK":
                    nok += 1
                elif found_status == "NA":
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


def progress_user(user, vinculo):
    responses = latest_responses()
    total = ok = nok = na = pending = 0

    for checkpoint in ["CP0700", "CP1400", "CP2000"]:
        for tipo in ["posicao", "geral"]:
            items = get_items(user, vinculo, tipo, checkpoint)

            for _, item in items.iterrows():
                total += 1
                item_id = str(item.get("item_padrao_id", ""))
                key = (
                    norm(vinculo["unidade_id"]),
                    norm(checkpoint),
                    norm(tipo),
                    norm(item_id),
                    norm(user["pessoa_id"]),
                )

                status = response_status(responses.get(key, {}).get("resposta", ""))

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


def render_login():
    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    st.markdown("""
    <div class="box">
        <div class="title">LIIVV Checklist</div>
        <div class="subtitle">Entre para iniciar suas atividades.</div>
    """, unsafe_allow_html=True)

    with st.form("login"):
        login = st.text_input("Login")
        senha = st.text_input("Senha", type="password")
        submit = st.form_submit_button("Entrar", use_container_width=True)

    if submit:
        user, error = authenticate(login, senha)

        if error:
            st.error(error)
        else:
            st.session_state["logged"] = True
            st.session_state["user"] = user
            st.session_state["page"] = "home"
            st.rerun()

    st.markdown("</div></div>", unsafe_allow_html=True)


def render_header(user, vinculo):
    st.markdown(f"""
    <div class="box">
        <div class="title">Olá, {user.get("nome", "Colaboradora")}</div>
        <div class="subtitle">{vinculo.get("unidade_nome", "")} • {vinculo.get("posicao_nome", "")}</div>
    </div>
    """, unsafe_allow_html=True)


def render_home(user, vinculo):
    render_header(user, vinculo)

    checked = has_checkin_today(user, vinculo)
    cp_id, cp_label = current_checkpoint()
    summary = progress_user(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)

    if checked:
        st.success("Check-in realizado hoje.")
    else:
        st.warning("Faça o check-in antes de iniciar as atividades.")

    st.write(f"**Horário sugerido:** {cp_label}")
    st.write(f"**Progresso do dia:** {summary['pct']}%")
    st.progress(summary["pct"] / 100 if summary["pct"] else 0)

    st.markdown("</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Fazer check-in", type="primary", use_container_width=True):
            st.session_state["page"] = "checkin"
            st.rerun()

    with c2:
        if st.button("Abrir checklist", use_container_width=True):
            st.session_state["page"] = "checklist"
            st.rerun()

    if norm(user.get("perfil")) == "admin":
        if st.button("Dashboard operacional", use_container_width=True):
            st.session_state["page"] = "dashboard"
            st.rerun()

    if st.button("Sair", use_container_width=True):
        st.session_state.clear()
        st.rerun()


def render_checkin(user, vinculo):
    render_header(user, vinculo)

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Check-in")
    st.write("Toque no botão abaixo e permita o acesso à localização do telefone.")

    if get_geolocation is None:
        st.error("Erro: o componente de localização não está disponível.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    try:
        location = get_geolocation()
    except Exception:
        location = None

    if st.button("Registrar localização", type="primary", use_container_width=True):
        if not location or not isinstance(location, dict) or not location.get("coords"):
            st.error("Erro ao obter GPS do telefone. Permita a localização e tente novamente.")
        else:
            coords = location["coords"]
            lat = coords.get("latitude")
            lon = coords.get("longitude")

            try:
                result = register_checkin(user, vinculo, lat, lon)

                if result["alerta"] == "sim":
                    st.warning("Check-in registrado com alerta.")
                else:
                    st.success("Check-in registrado com sucesso.")

                st.write(f"Distância calculada: **{result['distancia']:.1f} m**")
                st.write(f"Status: **{result['status_distancia']}**")
                st.caption(result["mensagem"])

            except Exception as exc:
                st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("Voltar", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()


def render_checklist(user, vinculo):
    render_header(user, vinculo)

    cp_current, _ = current_checkpoint()

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Checklist")

    cp_map = {
        "07:00": "CP0700",
        "14:00": "CP1400",
        "20:00": "CP2000",
    }

    labels = list(cp_map.keys())
    default_index = list(cp_map.values()).index(cp_current)

    cp_label = st.selectbox("Horário", labels, index=default_index)
    checkpoint = cp_map[cp_label]

    tipo_label = st.selectbox("Lista", ["Minha função", "Geral da unidade"])
    tipo = "posicao" if tipo_label == "Minha função" else "geral"

    items = get_items(user, vinculo, tipo, checkpoint)
    responses = latest_responses()

    st.markdown("</div>", unsafe_allow_html=True)

    if items.empty:
        st.info("Nenhuma atividade encontrada.")
    else:
        for _, item in items.iterrows():
            item_id = str(item.get("item_padrao_id", ""))
            item_title = str(item.get("item", item_id))
            item_detail = str(item.get("detalhe", ""))

            key = (
                norm(vinculo["unidade_id"]),
                norm(checkpoint),
                norm(tipo),
                norm(item_id),
                norm(user["pessoa_id"]),
            )

            status = response_status(responses.get(key, {}).get("resposta", ""))

            css = "task-box"
            if status == "OK":
                css += " task-ok"
            elif status == "NOK":
                css += " task-nok"
            elif status == "NA":
                css += " task-na"

            st.markdown(f"""
            <div class="{css}">
                <div class="task-title">{item_title}</div>
                <div class="task-detail">{item_detail}</div>
                <div style="margin-top:8px;">Status: <b>{status}</b></div>
            </div>
            """, unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)

            with c1:
                if st.button("OK", key=f"ok_{checkpoint}_{tipo}_{item_id}", use_container_width=True):
                    save_response(user, vinculo, checkpoint, tipo, item_id, "OK")
                    st.rerun()

            with c2:
                if st.button("Não OK", key=f"nok_{checkpoint}_{tipo}_{item_id}", use_container_width=True):
                    save_response(user, vinculo, checkpoint, tipo, item_id, "NÃO OK")
                    st.rerun()

            with c3:
                if st.button("N/A", key=f"na_{checkpoint}_{tipo}_{item_id}", use_container_width=True):
                    save_response(user, vinculo, checkpoint, tipo, item_id, "N/A")
                    st.rerun()

    if st.button("Voltar", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()


def render_dashboard(user):
    if norm(user.get("perfil")) != "admin":
        st.error("Área restrita.")
        return

    data = load_data()
    unidades = data["UNIDADES"]

    st.markdown("""
    <div class="box">
        <div class="title">Dashboard operacional</div>
        <div class="subtitle">Status consolidado por unidade</div>
    </div>
    """, unsafe_allow_html=True)

    if unidades.empty:
        st.info("Nenhuma unidade cadastrada.")
    else:
        active = unidades[unidades["ativa"].astype(str).map(norm) == "sim"]
        cols = st.columns(3)
        idx = 0

        for _, unidade in active.iterrows():
            unidade_id = str(unidade.get("unidade_id", ""))
            unidade_nome = str(unidade.get("unidade_nome", unidade_id))
            s = progress_unit(unidade_id)

            card_class = "card"
            if s["pct"] >= 90:
                card_class += " card-green"
            elif s["pct"] >= 60:
                card_class += " card-yellow"
            elif s["pct"] < 40:
                card_class += " card-red"

            with cols[idx % 3]:
                st.markdown(f"""
                <div class="{card_class}">
                    <div class="metric-label">{unidade_nome}</div>
                    <div class="metric-big">{s['pct']}%</div>
                    <div class="metric-detail">
                        ✅ OK {s['ok']}<br>
                        ❌ Não OK {s['nok']}<br>
                        ⏳ Pendentes {s['pending']}<br>
                        ℹ️ N/A {s['na']}
                    </div>
                </div>
                """, unsafe_allow_html=True)

            idx += 1

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.subheader("Registros")

    tab1, tab2, tab3 = st.tabs(["Alertas", "Check-ins", "Respostas"])

    with tab1:
        df = read_sheet("ALERTAS")
        if df.empty:
            st.info("Sem alertas.")
        else:
            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    with tab2:
        df = read_sheet("CHECKINS")
        if df.empty:
            st.info("Sem check-ins.")
        else:
            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    with tab3:
        df = read_sheet("RESPOSTAS_CHECKLIST")
        if df.empty:
            st.info("Sem respostas.")
        else:
            st.dataframe(df.sort_values("timestamp", ascending=False), use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("Voltar", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()


def render_app():
    user = st.session_state["user"]
    vinculo = selected_vinculo(user)
    page = st.session_state.get("page", "home")

    if page == "home":
        render_home(user, vinculo)
    elif page == "checkin":
        render_checkin(user, vinculo)
    elif page == "checklist":
        render_checklist(user, vinculo)
    elif page == "dashboard":
        render_dashboard(user)
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
