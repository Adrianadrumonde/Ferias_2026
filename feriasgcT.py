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
import streamlit as st
from google.oauth2.service_account import Credentials
import gspread
import base64

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


def guardar_no_sheets(linhas):
    """
    linhas = lista de listas (cada lista é uma linha do Sheet)
    """
    sheet.append_rows(linhas, value_input_option="USER_ENTERED")

SMTP_SERVER = "mail.cesab.pt"
SMTP_PORT = 465  # SSL
SMTP_USER = st.secrets["user"]
SMTP_PASS = st.secrets["pass"]
DESTINO_EMAIL = "adrianadrumonde@sapo.pt"

# =========================
# LISTA DE FUNCIONÁRIOS
# =========================
FUNCIONARIOS = ["","Carla Sério","Adriana Drumonde","Maria Paulino","Elsa Barracho","Sandra Paulo","João Pereira",
                "Armanda Fernandes","Andreia Mendes","Sarah Silva","Brenda Santos","M.ª do Céu Martins",
                "Ana Joaquina","André Barandas","Maksym Martens ","Jaqueline Reis","Alexandra Rajado","Diogo Reis","Liliana Nisa",
                "Sandra Pinheiro","Mónica Cerveira","Cláudia Bernardes","Beatriz Martinho","Eliari Silva",
                "Marta Pedroso","Bruno Albuquerque","Tiago Daniel","Vítor Antunes","Óscar Soares","Rúben Rosa", "Catarina Torres",
                "André Martins", "Rafael Vivas", "Telmo Menoita", "Edgar Martins", "Bruno Santos",
                "Renato Alves",  "Fábio Pego", "Tomas Fernandes", "Tiago Costa", "Gabriel Pinto", "Carina Gonçalves",
                ]
FUNCIONARIOS = sorted(FUNCIONARIOS)

MAPA_SECCOES = {
    "Adriana Drumonde": "GAT",
    "Carla Sério": "GAT",
    "Elsa Barracho": "GESTÃO E SEC",
    "Sandra Paulo": "GESTÃO E SEC",
    "João Pereira": "GESTÃO E SEC",
    "Maria Paulino": "GESTÃO E SEC",

    "Andreia Mendes": "LOG.",
    "Sarah Silva": "LOG.",
    "Armanda Fernandes": "LOG.",

    "M.ª do Céu Martins": "Apoio Lab.",
    "Ana Joaquina": "Apoio Lab.",
    "André Barandas": "Apoio Lab.",
    "Brenda Santos": "Apoio Lab.",

    "Alexandra Rajado": "Laboratório",
    "Diogo Reis": "Laboratório",
    "Liliana Nisa": "Laboratório",
    "Sandra Pinheiro": "Laboratório",
    "Mónica Cerveira": "Laboratório",
    "Cláudia Bernardes": "Laboratório",
    "Beatriz Martinho": "Laboratório",
    "Eliari Silva": "Laboratório",
    "Marta Pedroso": "Laboratório",
    "Bruno Albuquerque": "Laboratório",
    "Jaqueline Reis": "Laboratório",
    "Carina Gonçalves": "Laboratório",

    "Vítor Antunes": "Colheitas",
    "Óscar Soares": "Colheitas",
    "Rúben Rosa": "Colheitas",
    "Catarina Torres": "Colheitas",
    "André Martins": "Colheitas",
    "Rafael Vivas": "Colheitas",
    "Telmo Menoita": "Colheitas",
    "Edgar Martins": "Colheitas",
    "Bruno Santos": "Colheitas",
    "Renato Alves": "Colheitas",
    "Fábio Pego": "Colheitas",
    "Tiago Daniel": "Colheitas",
    "Gabriel Pinto": "Colheitas",
    "Tomás Fernandes": "Colheitas",
    "Tiago Costa": "Colheitas",
}
MAPA_EMAIL_SECCAO = {
    "GAT": "a.drumonde@cesab.pt",
    "GESTÃO E SEC": "j.pereira@cesab.pt",
    "LOG.": "j.pereira@cesab.pt",
    "Apoio Lab.": "j.pereira@cesab.pt",
    "Laboratório": "laboratorio@cesab.pt",
    "Colheitas": "g.tecnico@cesab.pt",
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
        # descobrir a secção pelo nome
        seccao = MAPA_SECCOES.get(nome, None)
        # descobrir email da secção
        email_seccao = MAPA_EMAIL_SECCAO.get(seccao)
        # Preparar email
        subject = f"Solicitação de Férias_BH - {nome}"
        body = "Segue em anexo a solicitação de férias."

        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = DESTINO_EMAIL
        msg['Subject'] = subject
        # CC apenas se existir email para a secção
        if email_seccao:
            msg['Cc'] = email_seccao
            destinatarios = [DESTINO_EMAIL, email_seccao]
        else:
            destinatarios = [DESTINO_EMAIL]
        
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
        df["Data_Fim"] = pd.to_datetime(df["Data_de_im"])
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

    st.dataframe(df, use_container_width=True)

   
# ----------------------
    # GRÁFICO DE GANTT
    # ----------------------
    st.subheader("📅 Gráfico de Gantt – Períodos de Férias")
    # Filtrar apenas férias
    df_ferias = df[df["Tipo"] == "FERIAS"].copy()

    if df_ferias.empty:
        st.info("Não existem solicitações de férias para mostrar.")
    else:
        #Converter datas (sabemos que estas colunas existem nas férias)
        df_ferias["Data_Inicio"] = pd.to_datetime(df_ferias["Data_Inicio"])
        df_ferias["Data_Fim"] = pd.to_datetime(df_ferias["Data_Fim"])
        # Garantir duração mínima para o gráfico
        df_ferias["Data_Fim_plot"] = df_ferias["Data_Fim"]
        df_ferias.loc[
            df_ferias["Data_Fim_plot"] <= df_ferias["Data_Inicio"],
            "Data_Fim_plot"
        ] = df_ferias["Data_Inicio"] + pd.Timedelta(days=1)
        
        # Período como texto (cores)
        df_ferias["Período"] = df_ferias["Período"].astype(str)

        fig = px.timeline(
            df_ferias,
            x_start="Data_Inicio",
            x_end="Data_Fim_plot",
            y="Nome",
            color="Período",
            hover_data=["Dias_Úteis", "Observações", "Data_Fim"]
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

    st.header("📁 Férias aprovadas (apenas leitura)")
    try:
        aprov_sheet = spreadsheet.worksheet("Férias_aprovadas")
        # Construir link de visualização para a folha específica (usa gid do worksheet)
        gid = getattr(aprov_sheet, "id", None)
        sheet_key = st.secrets.get("sheet_id")
        if sheet_key and gid is not None:
            # Tentar exportar apenas a folha específica como PDF usando as credenciais do serviço
            try:
                from google.auth.transport.requests import AuthorizedSession

                authed_session = AuthorizedSession(creds)
                export_url = f"https://docs.google.com/spreadsheets/d/{sheet_key}/export"
                params = {
                    "exportFormat": "pdf",
                    "format": "pdf",
                    "gid": str(gid),
                    # request landscape A4 and prefer fitting to sheet page breaks/width
                    "portrait": "false",
                    "size": "A4",
                    "scale": "1", 
                    "fitw": "false",
                    "gridlines": "false",
                    "printtitle": "false",
                    "sheetnames": "false",
                    "fzr": "true",
                }
                headers = {"Accept": "application/pdf"}
                resp = authed_session.get(export_url, params=params, headers=headers, allow_redirects=True)

                # Obter content-type de forma segura antes de usar
                content_type = resp.headers.get("content-type", "")

                # Verificar se a resposta é um PDF válido antes de embutir
                if resp.status_code == 200 and "pdf" in content_type.lower() and len(resp.content) > 1000:
                    pdf_bytes = resp.content
                    # Verificar header do PDF
                    is_pdf = pdf_bytes.startswith(b"%PDF")
                    st.info(f"PDF header valid: {is_pdf}")

                    # Detectar número de colunas congeladas (se disponível nas propriedades do worksheet)
                    try:
                        frozen_cols = int(getattr(aprov_sheet, '_properties', {}).get('gridProperties', {}).get('frozenColumnCount', 0) or 0)
                    except Exception:
                        frozen_cols = 0
                    st.info(f"Colunas congeladas detectadas na folha: {frozen_cols}")

                    # Oferecer opção de pós-processamento para repetir colunas congeladas
                    repetir = st.checkbox("Repetir colunas congeladas em cada página (pós-processamento)", value=False)
                    if repetir and frozen_cols == 0:
                        st.warning("Nenhuma coluna congelada detectada; o pós-processamento não fará sentido.")
                        repetir = False
                    if repetir:
                        try:
                            import fitz  # pymupdf
                        except Exception:
                            st.error("Biblioteca 'pymupdf' não encontrada. Instale com: pip install pymupdf")
                            repetir = False

                    # Antes do download: se o PDF vier numa única página muito alta,
                    # dividir em páginas A4 (paisagem) para preservar quebras/legibilidade.
                    try:
                        import fitz
                        src_tmp = fitz.open(stream=pdf_bytes, filetype="pdf")
                        page_count_tmp = src_tmp.page_count
                    except Exception:
                        src_tmp = None
                        page_count_tmp = None

                    def split_long_pdf_to_a4(pdf_bytes_in, orientation_landscape=True, dpi=150):
                        import fitz, math
                        src = fitz.open(stream=pdf_bytes_in, filetype="pdf")
                        # A4 sizes in points (portrait: width=595.276, height=841.89)
                        if orientation_landscape:
                            target_h = 595.276
                        else:
                            target_h = 841.89

                        if src.page_count > 1:
                            return pdf_bytes_in

                        p = src[0]
                        r = p.rect
                        # If page height is already approximately target, return
                        if r.height <= target_h + 2:
                            return pdf_bytes_in

                        out = fitz.open()
                        n_slices = math.ceil(r.height / target_h)
                        for i in range(n_slices):
                            top = i * target_h
                            bottom = min((i + 1) * target_h, r.height)
                            clip = fitz.Rect(0, top, r.width, bottom)
                            pix = p.get_pixmap(clip=clip, dpi=dpi)
                            img_bytes = pix.tobytes("png")
                            newp = out.new_page(width=r.width, height=(bottom - top))
                            newp.insert_image(fitz.Rect(0, 0, r.width, (bottom - top)), stream=img_bytes)

                        return out.tobytes()

                    # Tentar dividir se for apenas uma página longa
                    if page_count_tmp == 1:
                        try:
                            pdf_bytes = split_long_pdf_to_a4(pdf_bytes, orientation_landscape=True)
                            st.info("PDF dividido em várias páginas A4 quando necessário.")
                        except Exception as e:
                            st.warning(f"Não foi possível dividir o PDF em páginas A4: {e}")

                    # Oferecer botão de download (funciona mesmo que o embed seja bloqueado)
                    if repetir:
                        try:
                            def repeat_left_columns(pdf_bytes_in, frozen_cols):
                                import fitz

                                src = fitz.open(stream=pdf_bytes_in, filetype="pdf")
                                out = fitz.open()

                                # Determine left width from first page
                                first = src[0]
                                r = first.rect
                                # estimate left width: proportion of page width per frozen column
                                left_w = max(100, min(r.width * 0.5, r.width * (0.08 * max(1, frozen_cols))))

                                left_rect = fitz.Rect(0, 0, left_w, r.height)
                                left_pix = first.get_pixmap(clip=left_rect, dpi=150)
                                left_png = left_pix.tobytes("png")

                                for p in src:
                                    r = p.rect
                                    full_pix = p.get_pixmap(dpi=150)
                                    full_png = full_pix.tobytes("png")

                                    newp = out.new_page(width=r.width, height=r.height)
                                    newp.insert_image(fitz.Rect(0, 0, r.width, r.height), stream=full_png)
                                    # overlay left frozen area
                                    newp.insert_image(fitz.Rect(0, 0, left_w, r.height), stream=left_png)

                                return out.tobytes()

                            processed = repeat_left_columns(pdf_bytes, frozen_cols)
                            pdf_bytes = processed
                            st.success("Pós-processamento concluído: colunas repetidas em cada página.")
                        except Exception as e:
                            st.error(f"Erro no pós-processamento: {e}")

                    st.download_button("📥 Baixar PDF - Férias Aprovadas", data=pdf_bytes, file_name="ferias_aprovadas.pdf", mime="application/pdf")

                    # Tentar embutir (pode ser bloqueado pelo browser)
                    try:
                        b64 = base64.b64encode(pdf_bytes).decode("utf-8")
                        pdf_display = f"<iframe src=\"data:application/pdf;base64,{b64}\" width=\"100%\" height=800></iframe>"
                        st.components.v1.html(pdf_display, height=820)
                    except Exception:
                        st.warning("Embed falhou use o botão de download acima.")
                else:
                    # Não mostrar link de visualização para evitar pedidos de acesso de utilizadores
                    st.error("Export não retornou um PDF válido (ver detalhes abaixo). Não será mostrado o link de visualização para evitar pedidos de acesso automáticos.")
                    # Mostrar headers e um excerto do corpo para diagnóstico
                    try:
                        snippet = resp.content[:2000].decode("utf-8", errors="replace")
                    except Exception:
                        snippet = str(resp.content[:2000])
                    st.code(snippet)
                    st.write(dict(resp.headers))
            except Exception as e:
                # Mostrar o erro na UI para depuração
                st.error(f"Erro na exportação PDF (AuthorizedSession): {e}")
                st.info("Nota: não foi mostrado o link de visualização automática por motivos de segurança.")
        else:
            st.warning("Não foi possível construir o link de visualização (chave da sheet ou gid em falta).")

        # Sem fallback de download: apenas link e tentativa de iframe para visualização

    except Exception as e:
        st.error(f"Erro ao carregar a folha 'Férias_aprovadas': {e}")
