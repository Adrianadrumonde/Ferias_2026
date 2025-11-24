import streamlit as st
import pandas as pd
from datetime import date, timedelta
import os
import plotly.express as px
import holidays
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv


# Carrega variáveis do .env
load_dotenv()

SMTP_SERVER = "mail.cesab.pt"
SMTP_PORT = 465  # SSL
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
DESTINO_EMAIL = "a.drumonde@cesab.pt"


# =========================
# CONFIGURAÇÃO INICIAL
# =========================
st.set_page_config(page_title="Gestão de Férias", page_icon="🏖️", layout="centered")
st.title("🏖️ Sistema de Solicitação de Férias")

ARQUIVO_CSV = "ferias.csv"

# =========================
# FERIADOS PORTUGAL + MEALHADA
# =========================
feriados_pt = holidays.Portugal()

# Mealhada 2026 e 2027
feriados_pt[date(2026, 5, 14)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2027, 5, 6)] = "Feriado Municipal da Mealhada"


# =========================
# FUNÇÃO DIAS ÚTEIS
# =========================
def dias_uteis(inicio, fim):
    """Calcula dias úteis ignorando fins de semana e feriados."""
    dias = 0
    atual = inicio
    while atual <= fim:
        if atual.weekday() < 5 and atual not in feriados_pt:
            dias += 1
        atual += timedelta(days=1)
    return dias


# =========================
# SENHAS
# =========================
SENHA_FUNCIONARIO = "ferias2025"
SENHA_RH = "rh123"

if "autenticado_func" not in st.session_state:
    st.session_state.autenticado_func = False

if "autenticado_rh" not in st.session_state:
    st.session_state.autenticado_rh = False


# =========================
# GUARDAR SOLICITAÇÕES
# =========================
def salvar_solicitacao(nome, periodos):
    registros = []
    for p in periodos:
        registros.append({
            "Nome": nome,
            "Período": p["Período"],
            "Data de Início": p["Data de Início"],
            "Data de Término": p["Data de Término"],
            "Dias Úteis": p["Dias Úteis"],
            "Observações": p["Observações"]
        })

    novo = pd.DataFrame(registros)

    if os.path.exists(ARQUIVO_CSV):
        antigo = pd.read_csv(ARQUIVO_CSV)
        df_final = pd.concat([antigo, novo], ignore_index=True)
    else:
        df_final = novo

    df_final.to_csv(ARQUIVO_CSV, index=False)


# =========================
# FUNÇÃO PARA ENVIAR EMAIL COM ANEXO
# =========================
def enviar_email_com_anexo(nome, df_periodos):
    try:
        # Preparar email
        subject = f"Solicitação de Férias - {nome}"
        body = "Segue em anexo a solicitação de férias."

        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = DESTINO_EMAIL
        msg['Subject'] = subject

        msg.attach(MIMEText(body, "plain"))

        # Converte DataFrame para CSV em bytes
        csv_bytes = df_periodos.to_csv(index=False).encode("utf-8")
        part = MIMEApplication(csv_bytes, Name=f"solicitacao_{nome.replace(' ', '_')}.csv")
        part['Content-Disposition'] = f'attachment; filename="solicitacao_{nome.replace(" ", "_")}.csv"'
        msg.attach(part)

        # Envia e-mail
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, DESTINO_EMAIL, msg.as_string())

        return True
    except Exception as e:
        # Para debug, mostre erro no Streamlit
        st.error(f"Erro ao enviar email: {e}")
        return False


# =========================
# MENU LATERAL
# =========================
aba = st.sidebar.radio("📂 Menu", ["📅 Solicitar Férias", "📊 Visualizar Solicitações"])


# =========================
# ABA 1 – FORMULÁRIO
# =========================
if aba == "📅 Solicitar Férias":

    if not st.session_state.autenticado_func:
        st.header("🔐 Acesso ao Formulário")
        senha = st.text_input("Código de acesso:", type="password")
        if st.button("Entrar"):
            if senha == SENHA_FUNCIONARIO:
                st.session_state.autenticado_func = True
                st.success("Acesso autorizado!")
            else:
                st.error("Código incorreto.")
        st.stop()

    st.header("📅 Solicitação de Férias")
    nome = st.text_input("Nome do funcionário")

    periodos = []

    for i in range(1, 5):
        with st.expander(f"Período {i}", expanded=(i == 1)):
            incluir = st.checkbox(f"Incluir Período {i}", value=(i == 1))

            if incluir:
                inicio = st.date_input(f"Data de início {i}", date.today(), key=f"inicio_{i}")
                fim = st.date_input(f"Data de término {i}", date.today(), key=f"fim_{i}")
                obs = st.text_area(f"Observações (opcional) {i}", key=f"obs_{i}")

                if fim >= inicio:
                    n_dias = dias_uteis(inicio, fim)
                    st.info(f"📅 Dias úteis: **{n_dias}**")
                else:
                    st.warning("Data final deve ser posterior à inicial.")
                    n_dias = 0

                periodos.append({
                    "Período": i,
                    "Data de Início": inicio,
                    "Data de Término": fim,
                    "Dias Úteis": n_dias,
                    "Observações": obs
                })

    if st.button("📤 Enviar Solicitação"):
        if not nome:
            st.error("O nome é obrigatório.")
        elif not periodos:
            st.error("Nenhum período selecionado.")
        else:
            salvar_solicitacao(nome, periodos)
            st.success("Solicitação enviada com sucesso!")
            st.balloons()

            # download individual
            df_download = pd.DataFrame(periodos)
            csv_bytes = df_download.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Baixar cópia (CSV)",
                data=csv_bytes,
                file_name=f"solicitacao_{nome.replace(' ', '_')}.csv",
                mime="text/csv"
            )
            # Envia email automático
            df_periodos = pd.DataFrame(periodos)
            if enviar_email_com_anexo(nome, df_periodos):
                st.success("📧 Email enviado para o RH com sucesso!")

# =========================
# ABA 2 – RH
# =========================
elif aba == "📊 Visualizar Solicitações":

    if not st.session_state.autenticado_rh:
        st.header("🔐 Área do RH")
        senha = st.text_input("Senha RH:", type="password")
        if st.button("Entrar RH"):
            if senha == SENHA_RH:
                st.session_state.autenticado_rh = True
                st.success("Acesso autorizado!")
            else:
                st.error("Senha incorreta.")
        st.stop()

    st.header("📊 Solicitações Registradas")

    if not os.path.exists(ARQUIVO_CSV):
        st.info("Nenhuma solicitação encontrada.")
        st.stop()

    df = pd.read_csv(ARQUIVO_CSV)
    df["Data de Início"] = pd.to_datetime(df["Data de Início"])
    df["Data de Término"] = pd.to_datetime(df["Data de Término"])

    #nomes = ["(Todos)"] + sorted(df["Nome"].unique())
    #filtro = st.selectbox("Filtrar funcionário:", nomes)

    #if filtro != "(Todos)":
        #df = df[df["Nome"] == filtro]
    
    nomes = sorted(df["Nome"].unique())
    filtros = st.multiselect(
       "Filtrar funcionário(s):",
     nomes
    )

    if filtros:
     df = df[df["Nome"].isin(filtros)]

    st.dataframe(df, use_container_width=True)

    # ----------------------
    # GRÁFICO DE GANTT
    # ----------------------
    st.subheader("📅 Gráfico de Gantt – Períodos de Férias")
    fig = px.timeline(
        df,
        x_start="Data de Início",
        x_end="Data de Término",
        y="Nome",
        color="Período",
        hover_data=["Dias Úteis", "Observações"]
    )
    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(fig, use_container_width=True)

    # download geral
    csv_full = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Baixar CSV Completo",
        data=csv_full,
        file_name="solicitacoes_ferias.csv",
        mime="text/csv"
    )
