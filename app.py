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

def encontrar_dados(arquivo_bytes):
    xls = pd.ExcelFile(io.BytesIO(arquivo_bytes))
    candidatos = []

    for aba in xls.sheet_names:
        bruto = pd.read_excel(
            io.BytesIO(arquivo_bytes),
            sheet_name=aba,
            header=None,
            dtype=str
        )

        bruto = bruto.dropna(how="all").dropna(axis=1, how="all")

        if bruto.empty:
            continue

        melhor = None

        for inicio in range(min(len(bruto), 20)):
            bloco = bruto.iloc[inicio:].copy()
            bloco = bloco.dropna(how="all")

            if len(bloco.columns) >= 2 and len(bloco) >= 2:
                linhas_com_dados = bloco.iloc[:, 0].notna().sum()
                if linhas_com_dados > 0:
                    score = len(bloco)
                    if melhor is None or score > melhor[0]:
                        melhor = (score, inicio, bloco)

        if melhor:
            candidatos.append((aba, melhor[1], melhor[2]))

    return candidatos

def preparar_dados(bruto, inicio):
    dados = bruto.iloc[inicio:].copy()

    if dados.empty or len(dados.columns) < 2:
        raise ValueError("Não foi possível localizar as duas primeiras colunas da planilha.")

    dados = dados.iloc[:, :2].copy()
    dados.columns = ["Telefone", "Mensagem"]

    dados["Telefone"] = dados["Telefone"].fillna("").astype(str).str.strip()
    dados["Mensagem"] = dados["Mensagem"].fillna("").astype(str).str.strip()

    primeira = dados.iloc[0]["Telefone"].lower()
    if "telefone" in primeira or "celular" in primeira or "numero" in primeira or "número" in primeira:
        dados = dados.iloc[1:].copy()

    dados = dados[
        (dados["Telefone"] != "") &
        (dados["Telefone"].str.lower() != "nan")
    ].copy()

    if dados.empty:
        raise ValueError("Encontrei a planilha, mas não localizei telefones com dados.")

    dados["Telefone_Normalizado"] = dados["Telefone"].apply(normalizar_telefone)
    dados["Status"] = dados["Telefone_Normalizado"].apply(
        lambda x: "Aguardando" if x else "Número inválido"
    )
    dados["Detalhe"] = dados["Telefone_Normalizado"].apply(
        lambda x: "Pronto para processamento" if x else "Formato de telefone inválido"
    )

    return dados.reset_index(drop=True)

def executar_simulacao(df, progresso, status_texto, tabela):
    resultados = df.copy()
    validos = resultados.index[resultados["Status"] == "Aguardando"].tolist()

    if not validos:
        status_texto.warning("Não existem números válidos para processar.")
        return resultados

    for posicao, indice in enumerate(validos, start=1):
        telefone = resultados.at[indice, "Telefone_Normalizado"]
        status_texto.info(f"📞 Processando {posicao} de {len(validos)}: {telefone}")
        time.sleep(0.35)

        resultado = random.choices(
            ["Sucesso", "Não atendeu", "Ocupado", "Falha"],
            weights=[55, 25, 10, 10],
            k=1
        )[0]

        detalhes = {
            "Sucesso": "Contato atendido e mensagem processada",
            "Não atendeu": "Tentativa sem atendimento",
            "Ocupado": "Linha ocupada",
            "Falha": "Falha na tentativa"
        }

        resultados.at[indice, "Status"] = resultado
        resultados.at[indice, "Detalhe"] = detalhes[resultado]

        progresso.progress(posicao / len(validos))
        tabela.dataframe(resultados, use_container_width=True, hide_index=True)

    status_texto.success("Campanha concluída!")
    return resultados

def gerar_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    buffer.seek(0)
    return buffer

st.title("📞 Discador Automático")
st.caption("Importe sua planilha, valide os contatos e acompanhe a campanha")

arquivo = st.file_uploader("Carregue sua planilha Excel", type=["xlsx", "xls"])

if arquivo:
    try:
        arquivo_bytes = arquivo.getvalue()

        if st.session_state.get("arquivo_nome") != arquivo.name:
            candidatos = encontrar_dados(arquivo_bytes)

            if not candidatos:
                raise ValueError("Não encontrei nenhuma aba com dados. Verifique se o arquivo possui uma planilha preenchida.")

            st.session_state.candidatos = candidatos
            st.session_state.arquivo_nome = arquivo.name
            st.session_state.processado = False
            st.session_state.dados = None

        nomes = [
            f"{aba} • dados encontrados"
            for aba, _, _ in st.session_state.candidatos
        ]

        escolha = st.selectbox(
            "Selecione a aba que contém os contatos",
            range(len(nomes)),
            format_func=lambda i: nomes[i]
        )

        aba, inicio, bruto = st.session_state.candidatos[escolha]

        if (
            st.session_state.dados is None
            or st.session_state.get("aba_atual") != aba
        ):
            st.session_state.dados = preparar_dados(bruto, 0)
            st.session_state.aba_atual = aba
            st.session_state.processado = False

        dados = st.session_state.dados

        st.success(f"Planilha encontrada: {aba}")

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
            resumo = resultado["Status"].value_counts()

            st.subheader("Resultado da campanha")

            cols = st.columns(5)
            for coluna, nome in zip(
                cols,
                ["Sucesso", "Não atendeu", "Ocupado", "Falha", "Número inválido"]
            ):
                coluna.metric(nome, int(resumo.get(nome, 0)))

            st.dataframe(resultado, use_container_width=True, hide_index=True)

            excel = gerar_excel(resultado)

            st.download_button(
                "📥 Baixar relatório Excel",
                data=excel,
                file_name=f"resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as erro:
        st.error(f"Erro ao processar a planilha: {erro}")

else:
    st.info("A planilha pode ter títulos em qualquer linha. O sistema tentará localizar os dados automaticamente.")
