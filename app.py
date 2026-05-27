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
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>

.stApp {
    background: #EFE7DD;
}

[data-testid="stHeader"] {
    background: rgba(239,231,221,0.95);
}

.block-container {
    padding-top: 2rem;
}

.box {
    background: white;
    border-radius: 18px;
    padding: 22px;
    border: 1px solid #D7CFC3;
    margin-bottom: 16px;
}

.title {
    font-size: 34px;
    font-weight: 900;
    color: #0E2A47;
}

.subtitle {
    color: #6B7280;
    margin-top: 6px;
}

.card {
    background: #0E2A47;
    border-radius: 24px;
    padding: 22px;
    color: white;
    min-height: 210px;
    margin-bottom: 18px;
}

.card-green {
    background: #166534;
}

.card-yellow {
    background: #92400E;
}

.card-red {
    background: #7F1D1D;
}

.metric-big {
    font-size: 56px;
    font-weight: 900;
    line-height: 1;
}

.metric-label {
    font-size: 14px;
    opacity: 0.9;
}

.task-box {
    background: white;
    border-radius: 18px;
    border: 1px solid #D7CFC3;
    padding: 18px;
    margin-bottom: 14px;
}

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
    margin-top: 80px;
}

.stButton button {
    border-radius: 12px;
    font-weight: 700;
    height: 46px;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================

def now_sp():
    return datetime.now(TZ)


def date_str():
    return now_sp().strftime("%Y-%m-%d")


def datetime_str():
    return now_sp().strftime("%Y-%m-%d %H:%M:%S")


def strip_accents(text):
    text = str(text or "")
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def norm(value):
    return strip_accents(value).strip().lower()


def response_status(value):
    s = norm(value)

    if s == "ok":
        return "OK"

    if s in ["nao ok", "não ok", "nao_ok"]:
        return "NOK"

    if s in ["n/a", "nao aplicavel", "não aplicável"]:
        return "NA"

    return "PENDENTE"


def retryable(fn, tries=5):
    error = None

    for i in range(tries):
        try:
            return fn()
        except APIError as e:
            error = e
            time.sleep(0.8 * (2 ** i))

    raise error


# ============================================================
# GOOGLE SHEETS
# ============================================================

@st.cache_resource
def get_client():

    creds_dict = dict(st.secrets["google_service_account"])

    credentials = Credentials.from_service_account_info(
        creds_dict,
        scopes=SCOPES,
    )

    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():

    return retryable(
        lambda: get_client().open_by_key(SPREADSHEET_ID)
    )


@st.cache_data(ttl=30)
def read_sheet(sheet_name):

    sh = get_spreadsheet()

    try:
        ws = retryable(lambda: sh.worksheet(sheet_name))
        values = retryable(lambda: ws.get_all_records())

        return pd.DataFrame(values)

    except Exception:
        return pd.DataFrame()


def append_row(sheet_name, row):

    sh = get_spreadsheet()
    ws = retryable(lambda: sh.worksheet(sheet_name))

    retryable(
        lambda: ws.append_row(
            row,
            value_input_option="USER_ENTERED",
        )
    )

    st.cache_data.clear()


# ============================================================
# LOAD
# ============================================================

def load_data():

    return {
        "UNIDADES": read_sheet("UNIDADES"),
        "USUARIOS_APP": read_sheet("USUARIOS_APP"),
        "CHECKLIST_GERAL_PADRAO": read_sheet("CHECKLIST_GERAL_PADRAO"),
        "CHECKLIST_POSICAO_PADRAO": read_sheet("CHECKLIST_POSICAO_PADRAO"),
        "RESPOSTAS_CHECKLIST": read_sheet("RESPOSTAS_CHECKLIST"),
        "CHECKINS": read_sheet("CHECKINS"),
    }


# ============================================================
# AUTH
# ============================================================

def authenticate(login, senha):

    data = load_data()

    usuarios = data["USUARIOS_APP"]

    if usuarios.empty:
        return None

    user = usuarios[
        (usuarios["login"].astype(str).map(norm) == norm(login))
        &
        (usuarios["senha"].astype(str) == str(senha))
        &
        (usuarios["ativa"].astype(str).map(norm) == "sim")
    ]

    if user.empty:
        return None

    return user.iloc[0].to_dict()


# ============================================================
# CHECKLIST
# ============================================================

def get_items(user, tipo, checkpoint):

    data = load_data()

    if tipo == "geral":

        df = data["CHECKLIST_GERAL_PADRAO"]

        if df.empty:
            return pd.DataFrame()

        return df[
            (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
            &
            (df["ativo"].astype(str).map(norm) == "sim")
        ]

    df = data["CHECKLIST_POSICAO_PADRAO"]

    if df.empty:
        return pd.DataFrame()

    return df[
        (df["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
        &
        (df["posicao_id"].astype(str).map(norm) == norm(user["posicao_id"]))
        &
        (df["ativo"].astype(str).map(norm) == "sim")
    ]


def latest_responses():

    df = read_sheet("RESPOSTAS_CHECKLIST")

    if df.empty:
        return {}

    df = df[
        df["data"].astype(str).str[:10] == date_str()
    ]

    out = {}

    for _, row in df.iterrows():

        key = (
            norm(row["unidade_id"]),
            norm(row["checkpoint_id"]),
            norm(row["tipo_checklist"]),
            norm(row["item_id"]),
            norm(row["pessoa_id"]),
        )

        out[key] = row.to_dict()

    return out


def save_response(user, checkpoint, tipo, item_id, resposta):

    append_row(
        "RESPOSTAS_CHECKLIST",
        [
            user["unidade_id"],
            str(uuid.uuid4()),
            datetime_str(),
            date_str(),
            checkpoint,
            tipo,
            user["pessoa_id"],
            user["nome"],
            user["unidade_nome"],
            user["posicao_id"],
            user["posicao_nome"],
            item_id,
            item_id,
            resposta,
            "",
            "",
            resposta,
            "não",
        ],
    )


# ============================================================
# DASHBOARD
# ============================================================

def progress_summary_unit(unidade_id):

    data = load_data()

    responses = latest_responses()

    total = 0
    ok = 0
    nok = 0
    na = 0
    pending = 0

    for checkpoint in ["CP0700", "CP1400", "CP2000"]:

        for tipo in ["geral", "posicao"]:

            if tipo == "geral":
                items = data["CHECKLIST_GERAL_PADRAO"]
            else:
                items = data["CHECKLIST_POSICAO_PADRAO"]

            if items.empty:
                continue

            items = items[
                (items["checkpoint_id"].astype(str).map(norm) == norm(checkpoint))
                &
                (items["ativo"].astype(str).map(norm) == "sim")
            ]

            for _, item in items.iterrows():

                total += 1

                item_id = str(item["item_padrao_id"])

                found = False

                for key, value in responses.items():

                    if (
                        key[0] == norm(unidade_id)
                        and key[1] == norm(checkpoint)
                        and key[2] == norm(tipo)
                        and key[3] == norm(item_id)
                    ):

                        found = True

                        status = response_status(
                            value.get("resposta", "")
                        )

                        if status == "OK":
                            ok += 1
                        elif status == "NOK":
                            nok += 1
                        elif status == "NA":
                            na += 1
                        else:
                            pending += 1

                if not found:
                    pending += 1

    pct = round((ok / total) * 100) if total else 0

    return {
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

    st.markdown('<div class="login-box">', unsafe_allow_html=True)

    st.markdown("""
    <div class="box">
        <div class="title">LIIVV Checklist</div>
        <div class="subtitle">
            Faça login para iniciar as atividades.
        </div>
    """, unsafe_allow_html=True)

    with st.form("login"):

        login = st.text_input("Login")
        senha = st.text_input("Senha", type="password")

        submit = st.form_submit_button(
            "Entrar",
            use_container_width=True,
        )

    if submit:

        user = authenticate(login, senha)

        if not user:
            st.error("Login inválido.")
        else:
            st.session_state["logged"] = True
            st.session_state["user"] = user
            st.session_state["page"] = "home"

            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_home(user):

    st.markdown(f"""
    <div class="box">
        <div class="title">
            Olá, {user["nome"]}
        </div>

        <div class="subtitle">
            {user["unidade_nome"]} • {user["posicao_nome"]}
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "Fazer Check-in",
            use_container_width=True,
            type="primary",
        ):
            st.session_state["page"] = "checkin"
            st.rerun()

    with col2:

        if st.button(
            "Abrir Checklist",
            use_container_width=True,
        ):
            st.session_state["page"] = "checklist"
            st.rerun()

    if norm(user["perfil"]) == "admin":

        if st.button(
            "Dashboard operacional",
            use_container_width=True,
        ):
            st.session_state["page"] = "dashboard"
            st.rerun()

    if st.button(
        "Sair",
        use_container_width=True,
    ):
        st.session_state.clear()
        st.rerun()


def render_checkin(user):

    st.markdown("""
    <div class="box">
        <div class="title">
            Check-in
        </div>
    """, unsafe_allow_html=True)

    if get_geolocation is None:

        st.error(
            "Erro ao acessar localização."
        )

        return

    location = get_geolocation()

    if st.button(
        "Registrar localização",
        type="primary",
        use_container_width=True,
    ):

        if not location:
            st.error(
                "Erro ao obter GPS do telefone."
            )
        else:

            append_row(
                "CHECKINS",
                [
                    user["unidade_id"],
                    str(uuid.uuid4()),
                    datetime_str(),
                    date_str(),
                    now_sp().strftime("%H:%M:%S"),
                    user["pessoa_id"],
                    user["nome"],
                    user["unidade_nome"],
                    user["posicao_id"],
                    user["posicao_nome"],
                    location["coords"]["latitude"],
                    location["coords"]["longitude"],
                ],
            )

            st.success("Check-in realizado.")

    if st.button("Voltar"):
        st.session_state["page"] = "home"
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_checklist(user):

    st.markdown("""
    <div class="box">
        <div class="title">
            Checklist
        </div>
    """, unsafe_allow_html=True)

    checkpoint = st.selectbox(
        "Horário",
        {
            "07:00": "CP0700",
            "14:00": "CP1400",
            "20:00": "CP2000",
        }.keys()
    )

    checkpoint_id = {
        "07:00": "CP0700",
        "14:00": "CP1400",
        "20:00": "CP2000",
    }[checkpoint]

    tipo_label = st.selectbox(
        "Tipo",
        ["Minha função", "Geral"]
    )

    tipo = "posicao"

    if tipo_label == "Geral":
        tipo = "geral"

    items = get_items(
        user,
        tipo,
        checkpoint_id,
    )

    responses = latest_responses()

    st.markdown("</div>", unsafe_allow_html=True)

    for _, item in items.iterrows():

        item_id = str(item["item_padrao_id"])

        key = (
            norm(user["unidade_id"]),
            norm(checkpoint_id),
            norm(tipo),
            norm(item_id),
            norm(user["pessoa_id"]),
        )

        current = responses.get(key, {})

        status = response_status(
            current.get("resposta", "")
        )

        st.markdown(f"""
        <div class="task-box">
            <div class="task-title">
                {item["item"]}
            </div>

            <div class="task-detail">
                {item.get("detalhe", "")}
            </div>

            <div style="margin-top:8px;">
                Status atual:
                <b>{status}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button(
                "OK",
                key=f"ok_{item_id}",
                use_container_width=True,
            ):
                save_response(
                    user,
                    checkpoint_id,
                    tipo,
                    item_id,
                    "OK",
                )
                st.rerun()

        with c2:
            if st.button(
                "Não OK",
                key=f"nok_{item_id}",
                use_container_width=True,
            ):
                save_response(
                    user,
                    checkpoint_id,
                    tipo,
                    item_id,
                    "NÃO OK",
                )
                st.rerun()

        with c3:
            if st.button(
                "N/A",
                key=f"na_{item_id}",
                use_container_width=True,
            ):
                save_response(
                    user,
                    checkpoint_id,
                    tipo,
                    item_id,
                    "N/A",
                )
                st.rerun()

    if st.button(
        "Voltar",
        use_container_width=True,
    ):
        st.session_state["page"] = "home"
        st.rerun()


def render_dashboard():

    data = load_data()

    st.markdown("""
    <div class="box">
        <div class="title">
            Dashboard operacional
        </div>
    </div>
    """, unsafe_allow_html=True)

    unidades = data["UNIDADES"]

    cols = st.columns(3)

    index = 0

    for _, unidade in unidades.iterrows():

        if norm(unidade["ativa"]) != "sim":
            continue

        unidade_id = unidade["unidade_id"]
        unidade_nome = unidade["unidade_nome"]

        summary = progress_summary_unit(
            unidade_id
        )

        pct = summary["pct"]

        card_class = "card"

        if pct >= 90:
            card_class += " card-green"
        elif pct >= 60:
            card_class += " card-yellow"
        elif pct < 40:
            card_class += " card-red"

        with cols[index % 3]:

            st.markdown(f"""
            <div class="{card_class}">

                <div class="metric-label">
                    {unidade_nome}
                </div>

                <div class="metric-big">
                    {pct}%
                </div>

                <div style="margin-top:16px; line-height:1.9;">

                    ✅ OK {summary['ok']}<br>
                    ❌ Não OK {summary['nok']}<br>
                    ⏳ Pendentes {summary['pending']}<br>
                    ℹ️ N/A {summary['na']}

                </div>

            </div>
            """, unsafe_allow_html=True)

        index += 1

    if st.button(
        "Voltar",
        use_container_width=True,
    ):
        st.session_state["page"] = "home"
        st.rerun()


# ============================================================
# MAIN
# ============================================================

def render_app():

    user = st.session_state["user"]

    page = st.session_state.get(
        "page",
        "home",
    )

    if page == "home":
        render_home(user)

    elif page == "checkin":
        render_checkin(user)

    elif page == "checklist":
        render_checklist(user)

    elif page == "dashboard":
        render_dashboard()


def main():

    if not st.session_state.get("logged"):
        render_login()
    else:
        render_app()


if __name__ == "__main__":
    main()
