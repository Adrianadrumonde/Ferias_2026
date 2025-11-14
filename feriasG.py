import streamlit as st
import pandas as pd
from datetime import date, timedelta
import os
import plotly.express as px  # para o gráfico de Gantt
import holidays

# =========================
# CONFIGURAÇÃO INICIAL
# =========================
st.set_page_config(page_title="Gestão de Férias", page_icon="🏖️", layout="centered")
st.title("🏖️ Sistema de Solicitação de Férias")

# Nome do arquivo CSV
ARQUIVO_CSV = "ferias.csv"

# Feriados nacionais de Portugal
feriados_pt = holidays.Portugal()

# --- FERIADOS MUNICIPAIS DA MEALHADA ---
# 2026
feriados_pt.append({
    date(2026, 5, 14): "Feriado Municipal da Mealhada"
})

# 2027
feriados_pt.append({
    date(2027, 5, 6): "Feriado Municipal da Mealhada"
})

# =========================
# CONFIGURAÇÃO DE SENHAS
# =========================
SENHA_FUNCIONARIO = "ferias2025"  # senha para acessar o formulário
SENHA_RH = "rh123"                # senha para acessar o painel do RH

# Inicializa estados de autenticação
if "autenticado_func" not in st.session_state:
    st.session_state.autenticado_func = False

if "autenticado_rh" not in st.session_state:
    st.session_state.autenticado_rh = False

# =========================
# FUNÇÃO PARA SALVAR DADOS
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

    novo_registro = pd.DataFrame(registros)

    if os.path.exists(ARQUIVO_CSV):
        df_existente = pd.read_csv(ARQUIVO_CSV)
        df_atualizado = pd.concat([df_existente, novo_registro], ignore_index=True)
    else:
        df_atualizado = novo_registro

    df_atualizado.to_csv(ARQUIVO_CSV, index=False)

# =========================
# FUNÇÃO PARA CALCULAR DIAS ÚTEIS
# =========================
    """Calcula número de dias úteis (segunda a sexta) entre duas datas. Ignora os feriados"""
def dias_uteis(inicio, fim):
    dias = 0
    atual = inicio
    while atual <= fim:
        # weekday() => 0=2ª feira ... 4=6ª feira
        if atual.weekday() < 5 and atual not in feriados_pt:
            dias += 1
        atual += timedelta(days=1)
    return dias
# =========================
# INTERFACE DE NAVEGAÇÃO
# =========================
aba = st.sidebar.radio("📂 Menu", ["📅 Solicitar Férias", "📊 Visualizar Solicitações"])

# =========================
# ABA 1 - SOLICITAR FÉRIAS
# =========================
if aba == "📅 Solicitar Férias":
    if not st.session_state.autenticado_func:
        st.header("🔐 Acesso ao Formulário")
        senha = st.text_input("Digite o código de acesso:", type="password")
        if st.button("Entrar", key="entrar_func"):
            if senha == SENHA_FUNCIONARIO:
                st.session_state.autenticado_func = True
                st.success("✅ Acesso autorizado! Você pode preencher o formulário.")
            else:
                st.error("❌ Código incorreto.")
        st.stop()

    st.header("📅 Solicitação de Férias")
    st.markdown("Preencha abaixo os períodos desejados. É possível informar **até 4 períodos**.")

    nome = st.text_input("Nome do funcionário")

    periodos = []
    for i in range(1, 5):
        with st.expander(f"Período {i}", expanded=(i == 1)):
            incluir = st.checkbox(f"Incluir Período {i}", value=(i == 1))
            if incluir:
                data_inicio = st.date_input(f"Data de início {i}", date.today(), key=f"inicio_{i}")
                data_fim = st.date_input(f"Data de término {i}", date.today(), key=f"fim_{i}")
                observacoes = st.text_area(f"Observações (opcional) - Período {i}", key=f"obs_{i}")

                # Calcula dias úteis
                if data_fim >= data_inicio:
                    n_dias = dias_uteis(data_inicio, data_fim)
                    st.info(f"🧮 **{n_dias} dias úteis** de férias neste período.")
                else:
                    st.warning("⚠️ A data de término deve ser posterior à data de início.")
                    n_dias = 0

                periodos.append({
                    "Período": i,
                    "Data de Início": data_inicio,
                    "Data de Término": data_fim,
                    "Dias Úteis": n_dias,
                    "Observações": observacoes
                })

    if st.button("📤 Enviar Solicitação"):
        if not nome:
            st.error("⚠️ O campo 'Nome' é obrigatório.")
        elif not periodos:
            st.error("⚠️ Informe pelo menos um período.")
        else:
            dados_validos = True
            for p in periodos:
                if p["Data de Término"] < p["Data de Início"]:
                    st.error(f"⚠️ Data final deve ser posterior à inicial (Período {p['Período']}).")
                    dados_validos = False
                    break

            if dados_validos:
                salvar_solicitacao(nome, periodos)
                st.success(f"✅ Solicitação registrada com sucesso para {nome}!")
                st.balloons()

                # Gera CSV individual para download imediato
                df_download = pd.DataFrame(periodos)
                csv_download = df_download.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Baixar cópia da solicitação (CSV)",
                    data=csv_download,
                    file_name=f"solicitacao_{nome.replace(' ', '_')}.csv",
                    mime="text/csv"
                )

# =========================
# ABA 2 - VISUALIZAR SOLICITAÇÕES (RH)
# =========================
elif aba == "📊 Visualizar Solicitações":
    if not st.session_state.autenticado_rh:
        st.header("🔐 Área Restrita do RH")
        senha = st.text_input("Digite a senha de acesso:", type="password")
        if st.button("Entrar", key="entrar_rh"):
            if senha == SENHA_RH:
                st.session_state.autenticado_rh = True
                st.success("✅ Acesso autorizado! Bem-vindo, RH.")
            else:
                st.error("❌ Senha incorreta.")
        st.stop()

    st.header("📊 Painel de Solicitações de Férias")

    if os.path.exists(ARQUIVO_CSV):
        df = pd.read_csv(ARQUIVO_CSV)

        # Garante que as datas são do tipo datetime
        df["Data de Início"] = pd.to_datetime(df["Data de Início"])
        df["Data de Término"] = pd.to_datetime(df["Data de Término"])

        # Filtro por funcionário
        nomes = df["Nome"].unique().tolist()
        nome_filtro = st.selectbox("Filtrar por funcionário:", ["(Todos)"] + nomes)

        if nome_filtro != "(Todos)":
            df = df[df["Nome"] == nome_filtro]

        st.dataframe(df, use_container_width=True)

        # -------------------------
        # 📊 GRÁFICO DE GANTT
        # -------------------------
        st.subheader("📅 Gráfico de Gantt - Períodos de Férias")

        fig = px.timeline(
            df,
            x_start="Data de Início",
            x_end="Data de Término",
            y="Nome",
            color="Período",
            hover_data=["Dias Úteis", "Observações"],
            title="Distribuição de Férias por Funcionário"
        )
        fig.update_yaxes(autorange="reversed")  # Gantt padrão
        fig.update_layout(
            xaxis_title="Data",
            yaxis_title="Funcionário",
            legend_title="Período",
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)

        # Download CSV geral
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar todas as solicitações (CSV)",
            data=csv,
            file_name="solicitacoes_ferias.csv",
            mime="text/csv"
        )

    else:
        st.info("Nenhuma solicitação registrada até o momento.")
