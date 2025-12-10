import streamlit as st
import pandas as pd
from datetime import date, timedelta
import os
import plotly.express as px
import holidays

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
feriados_pt[date(2027, 5, 20)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2028, 5, 28)] = "Feriado Municipal da Mealhada"
feriados_pt[date(2029, 5, 30)] = "Feriado Municipal da Mealhada"


# =========================
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
    #df["Data de Início"] = pd.to_datetime(df["Data de Início"])
    #df["Data de Término"] = pd.to_datetime(df["Data de Término"])
    
    df["Data de Início"] = pd.to_datetime(df["Data de Início"])
    df["Data de Término"] = pd.to_datetime(df["Data de Término"])

    # Criar coluna Ano com base na data de início
    df["Ano"] = df["Data de Início"].dt.year

    # Calcular total de dias por pessoa/ano
    totais = df.groupby(["Nome", "Ano"])["Dias Úteis"].sum().reset_index()
    totais.rename(columns={"Dias Úteis": "Total dias/Ano"}, inplace=True)

    # Inserir no dataframe principal
    df = df.merge(totais, on=["Nome", "Ano"], how="left")
    
    #####
    nomes = sorted(df["Nome"].unique())
    filtros = st.multiselect(
       "Filtrar funcionário(s):",
     nomes
    )

    if filtros:
     df = df[df["Nome"].isin(filtros)]

    st.dataframe(df, use_container_width=True)


    # =========================
    # CONTADOR TOTAL ABA 2
    # =========================
    #total_dias_filtrado = df["Dias Úteis"].sum()
    #st.subheader(f"📘 Total de dias úteis (filtrados): **{total_dias_filtrado}**")


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

