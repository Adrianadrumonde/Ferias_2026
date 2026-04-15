# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import date, timedelta,datetime
import os
import plotly.express as px
import holidays
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2.service_account import Credentials
import gspread
import base64
import io
from google.auth.transport.requests import Request
import requests
from io import BytesIO



# FLAG para evitar envio de email repetido
if "email_enviado" not in st.session_state:
    st.session_state.email_enviado = False
if "email_formulario_enviado" not in st.session_state:
    st.session_state.email_formulario_enviado = False
# =========================
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

creds = Credentials.from_service_account_info(
    st.secrets.gcp_service_account,
    scopes=scope
)
client = gspread.authorize(creds)

spreadsheet = client.open_by_key(st.secrets["sheet_id"])
sheet = spreadsheet.worksheet("solicitacoes")

# Read employees and sections from Google Sheet
func_sheet = spreadsheet.worksheet("Funcionários e Secções")
data = func_sheet.get_all_values()
# Process data: assume column A = name, B = section, skip rows with empty name
func_data = [row for row in data if len(row) >= 2 and row[0].strip()]
FUNCIONARIOS = [""] + sorted([row[0].strip() for row in func_data])
MAPA_SECCOES = {row[0].strip(): row[1].strip() for row in func_data if row[1].strip()}

def guardar_no_sheets(linhas):
    """
    linhas = lista de listas (cada lista é uma linha do Sheet)
    """
    sheet.append_rows(linhas, value_input_option="USER_ENTERED")

SMTP_SERVER = "mail.cesab.pt"
SMTP_PORT = 465  # SSL
SMTP_USER = st.secrets["user"]
SMTP_PASS = st.secrets["pass"]
DESTINO_EMAIL = "s.paulo@cesab.pt"

# =========================
# LISTA DE FUNCIONÁRIOS e MAPA_SECCOES são carregados do Google Sheets
# =========================

MAPA_EMAIL_SECCAO = {
    "GAT": "j.pereira@cesab.pt",
    "GESTÃO E SEC": "j.pereira@cesab.pt",
    "LOG.": "j.pereira@cesab.pt",
    "Apoio Lab.": "j.pereira@cesab.pt",
    "Laboratório": "j.pereira@cesab.pt;laboratorio@cesab.pt",
    "Colheitas": "a.drumonde@cesab.pt",
}
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
SENHA_FUNCIONARIO = st.secrets["SENHA_FUNCIONARIO"]
SENHA_RH = st.secrets["SENHA_RH"]

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
        # descobrir a secção pelo nome
        seccao = MAPA_SECCOES.get(nome, None)
        # descobrir email da secção
        email_seccao = MAPA_EMAIL_SECCAO.get(seccao)
        # Preparar email
        subject = f"Solicitação de Férias_BH - {nome}"
        body = "Segue em anexo a solicitação de férias_BH."

        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = DESTINO_EMAIL
        msg['Subject'] = subject
        # CC apenas se existir email para a secção
        #if email_seccao:
         #   msg['Cc'] = email_seccao
          #  destinatarios = [DESTINO_EMAIL, email_seccao]

        if email_seccao:
            lista_cc = [e.strip() for e in email_seccao.split(";")]
            msg['Cc'] = ", ".join(lista_cc)
            destinatarios = [DESTINO_EMAIL] + lista_cc
        else:
            destinatarios = [DESTINO_EMAIL]


        
        msg.attach(MIMEText(body, "plain"))

        #  CONVERTER DATAFRAME PARA EXCEL (EM MEMÓRIA)
        buffer = BytesIO()
        df_periodos.to_excel(
            buffer, 
            index=False, 
            sheet_name="Solicitação"
        )
        buffer.seek(0)
        part = MIMEApplication(
            buffer.read(),
              Name=f"solicitacao_{nome.replace(' ', '_')}.xlsx"
        )
        part['Content-Disposition'] = (
            f'attachment; filename="solicitacao_{nome.replace(" ", "_")}.xlsx"'
        )
        msg.attach(part)


        # Envia e-mail
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(SMTP_USER, SMTP_PASS)
           # server.sendmail(SMTP_USER, DESTINO_EMAIL, msg.as_string())
            server.sendmail(SMTP_USER, destinatarios, msg.as_string())
        return True
    except Exception as e:
        # Para debug, mostre erro no Streamlit
        st.error(f"Erro ao enviar email: {e}")
        return False

# =========================
# MENU LATERAL
# =========================
aba = st.sidebar.radio("📂 Menu", ["📊 Visualizar Solicitações", "📅 Solicitar Férias", "⏱️ Banco de Horas", "Férias aprovadas"])

# =========================
# ABA 1 – FORMULÁRIO
# =========================
if aba == "📅 Solicitar Férias":

    if "autenticado_func" not in st.session_state:
        st.session_state.autenticado_func = False
    if not st.session_state.autenticado_func:
        st.header("🔐 Acesso ao Formulário")
        senha_func = st.text_input("Código de acesso:", type="password", key="senha_func")
        if st.button("Entrar"):
            if senha_func == SENHA_FUNCIONARIO:
                st.session_state.autenticado_func = True
                st.success("Acesso autorizado!")
                st.rerun()
            else:
                st.error("Código incorreto.")
            # Se depois do clique ainda não estiver autenticado, interrompe aqui
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
            #Salvar_solicitacao(nome, periodos)
            linhas = []
            for p in periodos:
                linhas.append([
                    nome, 
                    "FERIAS",
                    p["Período"],
                    p["Data de Início"].isoformat(),
                    p["Data de Término"].isoformat(),
                    p["Dias Úteis"],
                    "",
                    p["Observações"],
                    datetime.now().isoformat()
                ])
            guardar_no_sheets(linhas)


            st.success("Solicitação enviada com sucesso!")
            st.balloons()

            # download individual
            df_download = pd.DataFrame(periodos)
            csv_bytes = df_download.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
       
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
            if not st.session_state.email_formulario_enviado:
                if enviar_email_com_anexo(nome, df_periodos):
                    st.session_state.email_formulario_enviado = True
                    st.success("📧 Email enviado para o RH com sucesso!")

# =========================
# ABA 2 – RH
# =========================
elif aba == "📊 Visualizar Solicitações":

    if "autenticado_rh" not in st.session_state:
        st.session_state.autenticado_rh = False
    if not st.session_state.autenticado_rh:
        st.header("🔐 Área do RH")
        senha_rh = st.text_input("Senha RH:", type="password", key="senha_rh")

        if st.button("Entrar RH"):
            if senha_rh == SENHA_RH:
                st.session_state.autenticado_rh = True
                st.success("Acesso autorizado!")
                st.rerun()
            else:
                st.error("Senha incorreta.")    
        st.stop()
    
    st.header("📊 Solicitações Registradas")
#carregar os dados uma única vez
    dados = sheet.get_all_records()
    df = pd.DataFrame(dados)

    if df.empty:
        st.info("Nenhuma solicitação encontrada no Google Sheets.")
        st.stop()
 # ---- CRIAR COLUNA SECÇÃO ----
    df["Secção"] = df["Nome"].map(MAPA_SECCOES).fillna("Sem Secção")

# Conversão de datas
    if "Data de Início" in df.columns:
        df["Data_Inicio"] = pd.to_datetime(df["Data_Inicio"])
    if "Data de Fim" in df.columns:
        df["Data_Fim"] = pd.to_datetime(df["Data_de_Fim"])
    if "Dias_Úteis" in df.columns:
        df["Dias_Úteis"] = pd.to_numeric(df["Dias_Úteis"], errors="coerce")
 
 # Filtro por secção
    seccoes = sorted(df["Secção"].unique())
    filtro_seccao = st.multiselect("Filtrar secção:", seccoes)
    if filtro_seccao:
        df = df[df["Secção"].isin(filtro_seccao)]
 #Filtro por funcionário   
    nomes = sorted(df["Nome"].unique())
    filtros = st.multiselect(
       "Filtrar funcionário(s):",nomes)
    if filtros:
     df = df[df["Nome"].isin(filtros)]
    
    # Remover Observações apenas da visualização
    df_vis = df.drop(
    columns=["Observações", "Timestamp", "Secção"],
    errors="ignore")
   
    st.dataframe(df_vis, use_container_width=True)

   
# ----------------------
    # GRÁFICO DE GANTT
    # ----------------------
    st.subheader("📅 Períodos de Férias e Banco de Horas (BH) Solicitados")
   
    df_gantt = df.copy()

    # Converter datas
    df_gantt["Data_Inicio"] = pd.to_datetime(df_gantt["Data_Inicio"])
    df_gantt["Data_Fim"] = pd.to_datetime(df_gantt["Data_Fim"])

    # Criar coluna de fim para o gráfico
    df_gantt["Data_Fim_plot"] = df_gantt["Data_Fim"]

    # Tratar BH (Banco de Horas)
    mask_bh = df_gantt["Tipo"] == "BH"

    def calcular_fim_bh(row):
        horas = 0
        if row["Parte"]:
            if "Manhã" in row["Parte"]:
                horas += 4
            if "Tarde" in row["Parte"]:
                horas += 4
        if horas == 0:
            horas = 4  # fallback seguro
        return row["Data_Inicio"] + pd.Timedelta(hours=horas)
    df_gantt.loc[mask_bh, "Data_Fim_plot"] = df_gantt[mask_bh].apply(
        calcular_fim_bh, axis=1
    )
    # Garantir duração mínima para férias
    df_gantt.loc[
        df_gantt["Data_Fim_plot"] <= df_gantt["Data_Inicio"],
        "Data_Fim_plot"
    ] = df_gantt["Data_Inicio"] + pd.Timedelta(days=1)

    fig = px.timeline(
        df_gantt,
        x_start="Data_Inicio",
        x_end="Data_Fim_plot",
        y="Nome",
        color="Tipo",  # FERIAS vs BH
        color_discrete_map={
            "FERIAS": "blue",
            "BH": "gray"
        },
        hover_data=["Tipo", "Parte", "Dias_Úteis"]
    )

    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(fig, use_container_width=True)

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
                    csv_bytes = df_bh.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.success("Solicitações BH preparadas com sucesso!")
                    st.download_button(
                        "📥 Baixar cópia (CSV) - BH",
                        data=csv_bytes,
                        file_name=f"solicitacao_bh_{nome.replace(' ', '_')}.csv",
                        mime="text/csv"
                    )

                    #  GUARDAR BH NO GOOGLE SHEETS# 
                    linhas = []
                    for r in registros_validos:
                        linhas.append([
                            nome,
                            "BH",
                            r["Período"],
                            r["Data"].isoformat(),
                            r["Data"].isoformat(),
                            None,
                            r["Parte"],
                            r["Observações"],
                            datetime.now().isoformat()
                        ])
                    guardar_no_sheets(linhas)

                    # Envia email automático com anexo para o RH (UMA ÚNICA VEZ)
                    if not st.session_state.email_enviado:
                        if enviar_email_com_anexo(nome, df_bh):
                            st.success("📧 Email com solicitações BH enviado para o RH com sucesso!")
                            st.session_state.email_enviado = True
                    st.balloons()


# =========================
# ABA 4 – FÉRIAS APROVADAS (SOMENTE LEITURA)
# =========================
elif aba == "Férias aprovadas":
   
    # =========================
    # AUTENTICAÇÃO RH (SOMENTE LEITURA)
    # =========================
    if "autenticado_ferias_aprovadas" not in st.session_state:
        st.session_state.autenticado_ferias_aprovadas = False

    if not st.session_state.autenticado_ferias_aprovadas:
        st.header("🔐 Área restrita – Férias aprovadas")
        senha = st.text_input("Senha RH:", type="password", key="senha_ferias_aprovadas")

        if st.button("Entrar"):
            if senha == SENHA_RH:
                st.session_state.autenticado_ferias_aprovadas = True
                st.success("Acesso autorizado")
                st.rerun()
            else:
                st.error("Senha incorreta")

        st.stop()

    st.title("Férias aprovadas")

    # Definir worksheet
    sheet_ferias = spreadsheet.worksheet("Férias_aprovadas")
    gid = sheet_ferias.id  # id da aba "Férias_aprovadas"
    sheet_id = st.secrets["sheet_id"]
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export"
        f"?format=xlsx&gid={gid}"
    )
    # Fazer download do Excel usando o token do Google
    creds.refresh(Request())
    token = creds.token

    response = requests.get(
        export_url,
        headers={"Authorization": f"Bearer {token}"}
    )
    if response.status_code == 200:
        excel_bytes = io.BytesIO(response.content)
        st.download_button(
            label="📥 Baixar Férias Aprovadas (Excel)",
            data=excel_bytes,
            file_name="ferias_aprovadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Erro ao baixar o arquivo Excel das Férias Aprovadas.")
