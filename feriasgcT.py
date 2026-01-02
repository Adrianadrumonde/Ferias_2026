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
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials

st.title("Teste Google Sheets")

scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope
)
client = gspread.authorize(creds)

sheet = client.open_by_key(st.secrets["sheets"]["sheet_id"]).sheet1

st.success("Ligação OK!")

if st.button("Escrever linha de teste"):
    sheet.append_row(["Teste", "Funciona!"])
    st.success("Linha escrita com sucesso!")


SMTP_SERVER = "mail.cesab.pt"
SMTP_PORT = 465  # SSL
SMTP_USER = st.secrets["user"]
SMTP_PASS = st.secrets["pass"]
DESTINO_EMAIL = "a.drumonde@cesab.pt"

# =========================
# LISTA DE FUNCIONÁRIOS
# =========================
FUNCIONARIOS = ["Carla Sério","Adriana Drumonde","Maria Paulino","Elsa Barracho","Sandra Paulo","João Pereira",
                "Armanda Fernandes","Andreia Mendes","Sarah Silva","Brenda Santos","M.ª do Céu Martins",
                "Ana Joaquina","André Barandas","Maksym Martens ","Jaqueline Reis","Alexandra Rajado","Diogo Reis","Liliana Nisa",
                "Sandra Pinheiro","Mónica Cerveira","Cláudia Bernardes","Beatriz Martinho","Eliari Silva",
                "Marta Pedroso","Bruno Albuquerque","Tiago Daniel","Vítor Antunes","Óscar Soares","Rúben Rosa", "Catarina Torres",
                "André Martins", "Rafael Vivas", "Telmo Menoita", "Edgar Martins", "Bruno Santos",
                "Renato Alves",  "Fábio Pego", "Pedro Robalo ", "Tomas Fernandes", "Tiago Costa", "Gabriel Pinto",
                ]
FUNCIONARIOS = sorted(FUNCIONARIOS)

# =========================
# CONFIGURAÇÃO INICIAL
# =========================
st.set_page_config(page_title="Gestão de Férias", page_icon="🏖️", layout="centered")
st.title("🏖️ Sistema de Solicitação de Férias e Banco de Horas")

ARQUIVO_CSV = "ferias.csv"

# =========================
# FERIADOS PORTUGAL + MEALHADA
# =========================
feriados_pt = holidays.country_holidays("PT")
#feriados_pt = holidays.Portugal()

# Mealhada 2026 e 2027
feriados_pt[date(2026, 5, 14)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2027, 5, 20)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2028, 5, 28)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2029, 5, 30)] = "Feriado Municipal da Mealhada"

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
aba = st.sidebar.radio("📂 Menu", ["📅 Solicitar Férias", "📊 Visualizar Solicitações", "⏱️ Banco de Horas"])

# =========================
# ABA 1 – FORMULÁRIO
# =========================
if aba == "📅 Solicitar Férias":

    if not st.session_state.get("autenticado_func", False):
        st.header("🔐 Acesso ao Formulário")
        senha = st.text_input("Código de acesso:", type="password")
        if st.button("Entrar"):
            if senha.strip().lower() == SENHA_FUNCIONARIO.lower():
                st.session_state.autenticado_func = True
                st.success("Acesso autorizado!")
            else:
                st.error("Código incorreto.")
        # Se depois do clique ainda não estiver autenticado, interrompe aqui
        if not st.session_state.get("autenticado_func", False):
            st.stop()

    st.header("📅 Solicitação de Férias")
    nome = st.selectbox("Nome do funcionário", FUNCIONARIOS)

    periodos = []

    for i in range(1, 6):
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
    # =========================
    # CONTADOR TOTAL ABA 1
    # =========================
    total_dias = sum(p["Dias Úteis"] for p in periodos if p["Dias Úteis"] > 0)
    st.subheader(f"📘 Total de dias úteis solicitados: **{total_dias}**")

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
            # Adicionar nome do funcionário ao CSV do email
            df_periodos.insert(0, "Nome do funcionário", nome)
            if enviar_email_com_anexo(nome, df_periodos):
                st.success("📧 Email enviado para o RH com sucesso!")

# =========================
# ABA 2 – RH
# =========================
elif aba == "📊 Visualizar Solicitações":

    if not st.session_state.get("autenticado_rh", False):
        st.header("🔐 Área do RH")
        senha = st.text_input("Senha RH:", type="password")
        if st.button("Entrar RH"):
            if senha.strip().lower() == SENHA_RH.lower():
                st.session_state.autenticado_rh = True
                st.success("Acesso autorizado!")
            else:
                st.error("Senha incorreta.")
        # Se depois do clique ainda não estiver autenticado, interrompe aqui
        if not st.session_state.get("autenticado_rh", False):
            st.stop()

    st.header("📊 Solicitações Registradas")

    if not os.path.exists(ARQUIVO_CSV):
        st.info("Nenhuma solicitação encontrada.")
        st.stop()

    df = pd.read_csv(ARQUIVO_CSV)
    df["Data de Início"] = pd.to_datetime(df["Data de Início"])
    df["Data de Término"] = pd.to_datetime(df["Data de Término"])

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

    # Preparar um DataFrame de plotagem: se período tem apenas um dia (ou fim <= inicio),
    # garantimos que Data de Término_plot > Data de Início para que a barra seja visível.
    df_plot = df.copy()
    # já convertido acima, mas garantimos novamente caso
    df_plot["Data de Início"] = pd.to_datetime(df_plot["Data de Início"])
    df_plot["Data de Término"] = pd.to_datetime(df_plot["Data de Término"])
    df_plot["Data de Término_plot"] = df_plot["Data de Término"]
    # Se a data de término for igual ou anterior à de início, ajusta para início + 1 dia (só para plot)
    df_plot.loc[df_plot["Data de Término_plot"] <= df_plot["Data de Início"], "Data de Término_plot"] = df_plot["Data de Início"] + pd.Timedelta(days=1)
    # Converter Período para string para cores discretas e legíveis
    df_plot["Período"] = df_plot["Período"].astype(str)

    fig = px.timeline(
        df_plot,
        x_start="Data de Início",
        x_end="Data de Término_plot",
        y="Nome",
        color="Período",
        hover_data=["Dias Úteis", "Observações", "Data de Término"]
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

# =========================
# ABA 3 – BH / BANCO DE HORAS
# =========================
elif aba == "⏱️ Banco de Horas":

    if not st.session_state.get("autenticado_func", False):
        st.header("🔐 Acesso ao Formulário BH")
        senha = st.text_input("Código de acesso (BH):", type="password")
        if st.button("Entrar BH"):
            if senha.strip().lower() == SENHA_FUNCIONARIO.lower():
                st.session_state.autenticado_func = True
                st.success("Acesso autorizado!")
            else:
                st.error("Código incorreto.")
        # Se depois do clique ainda não estiver autenticado, interrompe aqui
        if not st.session_state.get("autenticado_func", False):
            st.stop()

    st.header("⏱️ Solicitação BH - Banco de Horas")
    nome = st.selectbox("Nome do funcionário", FUNCIONARIOS)

    st.markdown("Pode submeter até 3 solicitações de BH. Cada solicitação corresponde a um único dia e deve selecionar a parte do dia (manhã/tarde).")

    registros_bh = []
    # Permitimos até 3 solicitações
    for i in range(1, 4):
        with st.expander(f"Solicitação {i}", expanded=(i == 1)):
            incluir = st.checkbox(f"Incluir Solicitação {i}", value=False, key=f"bh_incluir_{i}")
            if incluir:
                data_bh = st.date_input(f"Data (Solicitação {i})", date.today(), key=f"bh_data_{i}")
                # Dois botões (checboxes) para manhã e tarde; pelo menos um deve ser selecionado
                parte_manha = st.checkbox("Manhã", value=False, key=f"bh_manha_{i}")
                parte_tarde = st.checkbox("Tarde", value=False, key=f"bh_tarde_{i}")
                obs_bh = st.text_area(f"Observações (opcional) {i}", key=f"bh_obs_{i}")

                registros_bh.append({
                    "Período": i,
                    "Data": data_bh,
                    "Manhã": parte_manha,
                    "Tarde": parte_tarde,
                    "Observações": obs_bh
                })

    if st.button("📤 Enviar Solicitações BH"):
        if not nome:
            st.error("O nome é obrigatório.")
        else:
            # Validar pelo menos uma solicitação e validações internas
            if not registros_bh:
                st.error("Nenhuma solicitação selecionada. Marque pelo menos uma 'Incluir Solicitação'.")
            else:
                erros = []
                registros_validos = []
                for r in registros_bh:
                    # Validação: pelo menos manhã ou tarde selecionada
                    if not (r["Manhã"] or r["Tarde"]):
                        erros.append(f"Na Solicitação {r['Período']} deve selecionar pelo menos 'Manhã' ou 'Tarde'.")
                    else:
                        # Construir campo 'Parte' com valores legíveis
                        partes = []
                        if r["Manhã"]:
                            partes.append("Manhã")
                        if r["Tarde"]:
                            partes.append("Tarde")
                        parte_str = ",".join(partes)
                        registros_validos.append({
                            "Período": r["Período"],
                            "Data": r["Data"],
                            "Parte": parte_str,
                            "Observações": r["Observações"]
                        })
                if erros:
                    for e in erros:
                        st.error(e)
                else:
                    # Criar DataFrame para download e envio
                    df_bh = pd.DataFrame(registros_validos)
                    # Inserir nome do funcionário na primeira coluna
                    df_bh.insert(0, "Nome do funcionário", nome)
                    csv_bytes = df_bh.to_csv(index=False).encode("utf-8")
                    st.success("Solicitações BH preparadas com sucesso!")
                    st.download_button(
                        "📥 Baixar cópia (CSV) - BH",
                        data=csv_bytes,
                        file_name=f"solicitacao_bh_{nome.replace(' ', '_')}.csv",
                        mime="text/csv"
                    )
                    # Envia email automático com anexo para o RH
                    if enviar_email_com_anexo(nome, df_bh):
                        st.success("📧 Email com solicitações BH enviado para o RH com sucesso!")
                    st.balloons()
