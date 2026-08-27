import io
import random
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Discador Automático", page_icon="📞", layout="wide")

def normalizar_telefone(valor):
    if pd.isna(valor):
        return None
    numero = re.sub(r"\D", "", str(valor))
    if len(numero) in (10, 11):
        numero = "55" + numero
    if len(numero) < 12 or len(numero) > 13:
        return None
    return "+" + numero

def ler_planilha(arquivo):
    planilhas = pd.read_excel(arquivo, sheet_name=None, header=0)
    candidatos = []

    for nome, df in planilhas.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if not df.empty:
            candidatos.append((nome, df))

    if not candidatos:
        raise ValueError("Não encontrei nenhuma linha com dados no arquivo.")

    return candidatos

def preparar_dados(df):
    if len(df.columns) < 2:
        raise ValueError("A planilha selecionada precisa ter pelo menos duas colunas.")

    dados = df.iloc[:, :2].copy()
    dados.columns = ["Telefone", "Mensagem"]
    dados = dados.dropna(how="all")

    dados["Telefone"] = dados["Telefone"].astype(str).str.strip()
    dados["Mensagem"] = dados["Mensagem"].fillna("").astype(str).str.strip()

    dados = dados[
        (dados["Telefone"] != "") &
        (dados["Telefone"].str.lower() != "nan")
    ].copy()

    if dados.empty:
        raise ValueError("Não encontrei telefones válidos para processar.")

    dados["Telefone_Normalizado"] = dados["Telefone"].apply(normalizar_telefone)
    dados["Status"] = dados["Telefone_Normalizado"].apply(
        lambda x: "Aguardando" if x else "Número inválido"
    )
    dados["Detalhe"] = dados["Telefone_Normalizado"].apply(
        lambda x: "Pronto para processamento" if x else "Telefone não reconhecido"
    )

    return dados

def executar_simulacao(df, progresso, status_texto, tabela):
    resultados = df.copy()
    validos = resultados.index[resultados["Status"] == "Aguardando"].tolist()

    if not validos:
        status_texto.warning("Não há números válidos para processar.")
        return resultados

    for posicao, indice in enumerate(validos, start=1):
        telefone = resultados.at[indice, "Telefone_Normalizado"]
        mensagem = resultados.at[indice, "Mensagem"]

        status_texto.info(
            f"📞 Processando {posicao} de {len(validos)}: {telefone}"
        )

        time.sleep(0.5)

        resultado = random.choices(
            ["Sucesso", "Não atendeu", "Ocupado", "Falha"],
            weights=[55, 25, 10, 10],
            k=1
        )[0]

        detalhes = {
            "Sucesso": f"Contato atendido. Mensagem processada: {mensagem[:60]}",
            "Não atendeu": "Chamada sem atendimento",
            "Ocupado": "Linha ocupada",
            "Falha": "Falha na tentativa de chamada"
        }

        resultados.at[indice, "Status"] = resultado
        resultados.at[indice, "Detalhe"] = detalhes[resultado]
        progresso.progress(posicao / len(validos))
        tabela.dataframe(resultados, use_container_width=True, hide_index=True)

    status_texto.success("Campanha simulada concluída!")
    return resultados

def gerar_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    buffer.seek(0)
    return buffer

st.title("📞 Discador Automático")
st.caption("Versão MVP • Importação de Excel • Validação • Simulação de contatos")

arquivo = st.file_uploader(
    "Carregue sua planilha Excel",
    type=["xlsx", "xls"]
)

if arquivo:
    try:
        if st.session_state.get("arquivo_nome") != arquivo.name:
            planilhas = ler_planilha(arquivo)
            st.session_state.planilhas = planilhas
            st.session_state.arquivo_nome = arquivo.name
            st.session_state.processado = False

        nomes = [nome for nome, _ in st.session_state.planilhas]

        nome_planilha = st.selectbox(
            "Selecione a aba da planilha que contém os telefones",
            nomes
        )

        df_original = dict(st.session_state.planilhas)[nome_planilha]
        df = preparar_dados(df_original)

        if "dados" not in st.session_state or st.session_state.get("aba_atual") != nome_planilha:
            st.session_state.dados = df
            st.session_state.aba_atual = nome_planilha
            st.session_state.processado = False

        dados = st.session_state.dados

        st.subheader("Pré visualização da campanha")
        tabela = st.empty()
        tabela.dataframe(dados, use_container_width=True, hide_index=True)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total", len(dados))
        c2.metric("Válidos", int(dados["Telefone_Normalizado"].notna().sum()))
        c3.metric("Inválidos", int(dados["Telefone_Normalizado"].isna().sum()))

        if not st.session_state.processado:
            if st.button("🚀 Iniciar campanha simulada", use_container_width=True):
                progresso = st.progress(0)
                status_texto = st.empty()
                st.session_state.dados = executar_simulacao(
                    dados, progresso, status_texto, tabela
                )
                st.session_state.processado = True
                st.rerun()
        else:
            resultado = st.session_state.dados

            st.subheader("Resultado da campanha")
            resumo = resultado["Status"].value_counts()

            colunas = st.columns(5)
            for coluna, nome in zip(
                colunas,
                ["Sucesso", "Não atendeu", "Ocupado", "Falha", "Número inválido"]
            ):
                coluna.metric(nome, int(resumo.get(nome, 0)))

            st.dataframe(resultado, use_container_width=True, hide_index=True)

            excel = gerar_excel(resultado)
            nome_arquivo = (
                f"resultado_discador_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )

            st.download_button(
                "📥 Baixar relatório Excel",
                data=excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as erro:
        st.error(f"Erro ao processar a planilha: {erro}")
else:
    st.info(
        "A planilha deve conter pelo menos duas colunas: "
        "Telefone e Mensagem."
    )
