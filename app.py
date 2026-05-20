import math
import uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

try:
    from streamlit_js_eval import get_geolocation
except Exception:
    get_geolocation = None


# ============================================================
# CONFIG
# ============================================================

TZ = ZoneInfo("America/Sao_Paulo")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

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
        background: rgba(239, 231, 221, 0.92);
    }
    .liivv-card {
        background: #ffffff;
        border: 1px solid #D7CFC3;
        border-radius: 18px;
        padding: 22px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        margin-bottom: 16px;
    }
    .liivv-title {
        color: #0E2A47;
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 4px;
    }
    .liivv-subtitle {
        color: #6B7785;
        font-size: 14px;
        margin-bottom: 18px;
    }
    .small-muted {
        color: #6B7785;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# GOOGLE SHEETS
# ============================================================

@st.cache_resource
def get_client():
    creds_dict = dict(st.secrets["google_service_account"])
    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    if not SPREADSHEET_ID:
        st.error("SPREADSHEET_ID não configurado nos secrets.")
        st.stop()

    try:
        return get_client().open_by_key(SPREADSHEET_ID)
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
        ws = sh.worksheet(sheet_name)
    except Exception:
        return pd.DataFrame()

    rows = ws.get_all_records()
    return pd.DataFrame(rows)


def get_worksheet(sheet_name: str):
    sh = get_spreadsheet()
    return sh.worksheet(sheet_name)


def append_row(sheet_name: str, row: list):
    ws = get_worksheet(sheet_name)
    ws.append_row(row, value_input_option="USER_ENTERED")
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
    ws.update_cell(row_index, col_index, value)
    st.cache_data.clear()
    return True


# ============================================================
# HELPERS
# ============================================================

def now_sp() -> datetime:
    return datetime.now(TZ)


def date_str(dt: datetime | None = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%Y-%m-%d")


def time_str(dt: datetime | None = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%H:%M:%S")


def datetime_str(dt: datetime | None = None) -> str:
    dt = dt or now_sp()
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def norm(value) -> str:
    return str(value or "").strip().lower()


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


def unidade_aplica(valor_unidade: str, unidade_id: str) -> bool:
    return norm(valor_unidade) in {"todas", norm(unidade_id)}


# ============================================================
# DATA LOAD
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

        if not unidade.empty and norm(unidade.iloc[0].get("ativa", "")) != "sim":
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

    update_cell_by_key("USUARIOS_APP", "usuario_id", user_row["usuario_id"], "ultimo_login", datetime_str())

    session = {
        "usuario_id": str(user_row["usuario_id"]),
        "pessoa_id": pessoa_id,
        "nome": str(pessoa_row["nome"]),
        "perfil": str(user_row.get("perfil", "operacao")),
        "vinculos": vinculos_out,
    }

    return session, None


def logout():
    for key in ["user", "logged"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ============================================================
# BUSINESS LOGIC
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


def get_checklist(data: dict, pessoa_id: str, unidade_id: str, checkpoint_id: str, tipo: str) -> pd.DataFrame:
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
    observacao: str,
    evidencia_url: str,
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

    if norm(resposta) in {"não ok", "nao ok"}:
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
# UI
# ============================================================

def render_login():
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.markdown('<div class="liivv-title">LIIVV Checklist</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="liivv-subtitle">Acesse para registrar check-in e checklist operacional.</div>',
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
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def select_vinculo(user: dict):
    vinculos = user.get("vinculos", [])
    if not vinculos:
        st.error("Usuário sem vínculo ativo.")
        st.stop()

    labels = [f"{v['unidade_nome']} | {v['posicao_nome']}" for v in vinculos]
    selected_label = st.selectbox("Unidade / posição", labels)
    selected_index = labels.index(selected_label)
    return vinculos[selected_index]


def render_checkin(user: dict, vinculo: dict):
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Check-in com geolocalização")

    st.write(f"**Pessoa:** {user['nome']}")
    st.write(f"**Unidade:** {vinculo['unidade_nome']}")
    st.write(f"**Posição:** {vinculo['posicao_nome']}")

    st.caption(
        "O navegador precisa permitir acesso à localização. "
        "Em prédios corporativos, o GPS pode oscilar. Ajuste o raio na aba UNIDADES se necessário."
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

    with st.expander("Localização capturada", expanded=False):
        st.write(location if location else "Localização ainda não capturada ou bloqueada.")

    manual = st.checkbox("Informar latitude e longitude manualmente", value=not bool(default_lat and default_lon))

    if manual:
        lat = st.number_input("Latitude atual", value=0.0, format="%.8f")
        lon = st.number_input("Longitude atual", value=0.0, format="%.8f")
    else:
        lat = default_lat
        lon = default_lon
        st.success(f"Localização capturada: {lat}, {lon}")

    if st.button("Registrar check-in", use_container_width=True):
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

                st.write(f"**Distância:** {result['distancia_metros']:.1f} m")
                st.write(f"**Raio permitido:** {result['raio_permitido_metros']} m")
                st.write(f"**Status distância:** {result['status_distancia']}")
                st.write(f"**Status horário:** {result['status_horario']}")
                st.caption(result.get("mensagem", ""))

        except Exception as exc:
            st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)


def render_checklist(user: dict, vinculo: dict):
    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Checklist")

    checkpoint_label = st.selectbox(
        "Checkpoint",
        [
            "CP0700 | 07:00 - Abertura",
            "CP1400 | 14:00 - Meio do dia",
            "CP2000 | 20:00 - Fechamento",
        ],
    )
    checkpoint_id = checkpoint_label.split("|")[0].strip()

    tipo_label = st.radio(
        "Tipo de checklist",
        ["Por posição", "Geral da unidade"],
        horizontal=True,
    )
    tipo = "posicao" if tipo_label == "Por posição" else "geral"

    data = load_data()
    checklist = get_checklist(data, user["pessoa_id"], vinculo["unidade_id"], checkpoint_id, tipo)

    if checklist.empty:
        st.info("Nenhum item encontrado para esta combinação.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    st.write(f"**Itens encontrados:** {len(checklist)}")

    for _, item in checklist.iterrows():
        item_id = str(item.get("item_padrao_id", ""))
        item_nome = str(item.get("item", ""))
        detalhe = str(item.get("detalhe", ""))
        obrigatorio = str(item.get("obrigatorio", ""))

        with st.expander(f"{item_nome}", expanded=False):
            st.caption(detalhe)
            st.caption(f"Obrigatório: {obrigatorio}")

            resposta = st.radio(
                "Resposta",
                ["OK", "NÃO OK", "NÃO APLICÁVEL"],
                key=f"resp_{checkpoint_id}_{tipo}_{item_id}",
                horizontal=True,
            )

            observacao = st.text_area(
                "Observação",
                key=f"obs_{checkpoint_id}_{tipo}_{item_id}",
            )

            evidencia_url = st.text_input(
                "URL de evidência, se houver",
                key=f"evid_{checkpoint_id}_{tipo}_{item_id}",
            )

            if st.button("Enviar este item", key=f"btn_{checkpoint_id}_{tipo}_{item_id}", use_container_width=True):
                try:
                    result = register_checklist_response(
                        unidade_id=vinculo["unidade_id"],
                        pessoa_id=user["pessoa_id"],
                        checkpoint_id=checkpoint_id,
                        tipo_checklist=tipo,
                        item_id=item_id,
                        resposta=resposta,
                        observacao=observacao,
                        evidencia_url=evidencia_url,
                    )

                    if result["alerta_gerado"] == "sim":
                        st.warning(f"Resposta registrada com alerta. Status: {result['status']}")
                    else:
                        st.success(f"Resposta registrada. Status: {result['status']}")

                except Exception as exc:
                    st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)


def render_admin(user: dict):
    if norm(user.get("perfil")) != "admin":
        return

    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.subheader("Painel admin")

    data = load_data()

    tabs = st.tabs(["Alertas", "Check-ins", "Respostas", "Unidades"])

    with tabs[0]:
        alertas = data["ALERTAS"]
        if alertas.empty:
            st.info("Nenhum alerta registrado.")
        else:
            st.dataframe(alertas.sort_values("timestamp", ascending=False), use_container_width=True)

    with tabs[1]:
        checkins = data["CHECKINS"]
        if checkins.empty:
            st.info("Nenhum check-in registrado.")
        else:
            st.dataframe(checkins.sort_values("timestamp", ascending=False), use_container_width=True)

    with tabs[2]:
        respostas = data["RESPOSTAS_CHECKLIST"]
        if respostas.empty:
            st.info("Nenhuma resposta registrada.")
        else:
            st.dataframe(respostas.sort_values("timestamp", ascending=False), use_container_width=True)

    with tabs[3]:
        unidades = data["UNIDADES"]
        if unidades.empty:
            st.info("Nenhuma unidade cadastrada.")
        else:
            st.dataframe(unidades, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)


def render_app():
    user = st.session_state["user"]

    st.sidebar.write(f"**{user['nome']}**")
    st.sidebar.caption(f"Perfil: {user.get('perfil', '')}")

    if st.sidebar.button("Sair"):
        logout()

    st.markdown('<div class="liivv-card">', unsafe_allow_html=True)
    st.markdown('<div class="liivv-title">LIIVV Checklist</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="liivv-subtitle">Check-in e controle operacional por unidade.</div>',
        unsafe_allow_html=True,
    )
    vinculo = select_vinculo(user)
    st.markdown("</div>", unsafe_allow_html=True)

    tab_checkin, tab_checklist, tab_admin = st.tabs(["Check-in", "Checklist", "Admin"])

    with tab_checkin:
        render_checkin(user, vinculo)

    with tab_checklist:
        render_checklist(user, vinculo)

    with tab_admin:
        render_admin(user)


# ============================================================
# MAIN
# ============================================================

def main():
    if not st.session_state.get("logged"):
        render_login()
    else:
        render_app()


if __name__ == "__main__":
    main()