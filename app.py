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

def processar_planilha(arquivo):
    df = pd.read_excel(arquivo)
    if len(df.columns) < 2:
        raise ValueError("A planilha precisa ter pelo menos duas colunas: Telefone e Mensagem.")
    df = df.iloc[:, :2].copy()
    df.columns = ["Telefone", "Mensagem"]
    df["Telefone_Normalizado"] = df["Telefone"].apply(normalizar_telefone)
    df["Status"] = df["Telefone_Normalizado"].apply(lambda x: "Aguardando" if x else "Número inválido")
    df["Detalhe"] = df["Telefone_Normalizado"].apply(lambda x: "Pronto para processamento" if x else "Telefone não reconhecido")
    return df

def executar_simulacao(df, progresso, status_texto):
    resultados = df.copy()
    validos = resultados.index[resultados["Status"] == "Aguardando"].tolist()

    if not validos:
        status_texto.warning("Não há números válidos para processar.")
        return resultados

    for posicao, indice in enumerate(validos, start=1):
        telefone = resultados.at[indice, "Telefone_Normalizado"]
        status_texto.info(f"📞 Processando {posicao} de {len(validos)}: {telefone}")
        time.sleep(0.3)

        resultado = random.choices(
            ["Sucesso", "Não atendeu", "Ocupado", "Falha"],
            weights=[55, 25, 10, 10],
            k=1,
        )[0]

        detalhes = {
            "Sucesso": "Contato atendido e mensagem simulada",
            "Não atendeu": "Chamada sem atendimento",
            "Ocupado": "Linha ocupada",
            "Falha": "Falha na tentativa de chamada",
        }

        resultados.at[indice, "Status"] = resultado
        resultados.at[indice, "Detalhe"] = detalhes[resultado]
        progresso.progress(posicao / len(validos))

    status_texto.success("Processamento concluído!")
    return resultados

def gerar_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    buffer.seek(0)
    return buffer

st.title("📞 Discador Automático")
st.caption("MVP com importação de Excel, validação e simulação de contatos")

arquivo = st.file_uploader("Carregue sua planilha Excel", type=["xlsx", "xls"])

if arquivo:
    if "dados" not in st.session_state or st.session_state.get("arquivo_nome") != arquivo.name:
        st.session_state.dados = processar_planilha(arquivo)
        st.session_state.arquivo_nome = arquivo.name
        st.session_state.processado = False

    df = st.session_state.dados
    st.subheader("Pré visualização da campanha")
    st.dataframe(df, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total", len(df))
    c2.metric("Válidos", int(df["Telefone_Normalizado"].notna().sum()))
    c3.metric("Inválidos", int(df["Telefone_Normalizado"].isna().sum()))

    if not st.session_state.processado:
        if st.button("🚀 Iniciar campanha simulada", use_container_width=True):
            progresso = st.progress(0)
            status_texto = st.empty()
            st.session_state.dados = executar_simulacao(df, progresso, status_texto)
            st.session_state.processado = True
            st.rerun()
    else:
        resultado = st.session_state.dados
        st.subheader("Resultado da campanha")
        resumo = resultado["Status"].value_counts()
        colunas = st.columns(4)
        for coluna, nome in zip(colunas, ["Sucesso", "Não atendeu", "Ocupado", "Falha"]):
            coluna.metric(nome, int(resumo.get(nome, 0)))

        st.dataframe(resultado, use_container_width=True, hide_index=True)

        excel = gerar_excel(resultado)
        nome = f"resultado_discador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "📥 Baixar relatório Excel",
            data=excel,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.info("A planilha deve ter duas colunas: Telefone e Mensagem.")
