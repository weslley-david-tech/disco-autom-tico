import io
import html
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st
from twilio.rest import Client

st.set_page_config(page_title="Discador Automático", page_icon="📞", layout="wide")

def normalizar_telefone(valor):
    if pd.isna(valor):
        return None
    numero = re.sub(r"\D", "", str(valor))
    if len(numero) in (10, 11):
        numero = "55" + numero
    if len(numero) not in (12, 13):
        return None
    return "+" + numero

def carregar_planilha(arquivo_bytes):
    xls = pd.ExcelFile(io.BytesIO(arquivo_bytes))
    abas = {}
    for aba in xls.sheet_names:
        df = pd.read_excel(
            io.BytesIO(arquivo_bytes),
            sheet_name=aba,
            header=None,
            dtype=object
        )
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if not df.empty:
            abas[aba] = df.reset_index(drop=True)
    if not abas:
        raise ValueError("Não encontrei dados em nenhuma aba do arquivo.")
    return abas

def preparar_dados(df):
    if df.shape[1] < 2:
        raise ValueError("A planilha precisa ter telefone na coluna A e mensagem na coluna B.")

    dados = df.iloc[:, :2].copy()
    dados.columns = ["Telefone", "Mensagem"]
    dados = dados.dropna(subset=["Telefone"]).copy()

    dados["Telefone"] = dados["Telefone"].astype(str).str.strip()
    dados["Mensagem"] = dados["Mensagem"].fillna("").astype(str).str.strip()

    if not dados.empty:
        primeira = dados.iloc[0]["Telefone"].lower()
        if any(x in primeira for x in ["telefone", "celular", "numero", "número"]):
            dados = dados.iloc[1:].copy()

    dados = dados[
        (dados["Telefone"] != "") &
        (dados["Telefone"].str.lower() != "nan")
    ].copy()

    if dados.empty:
        raise ValueError("Não encontrei telefones para processar.")

    dados["Telefone_Normalizado"] = dados["Telefone"].apply(normalizar_telefone)
    dados["Status"] = dados["Telefone_Normalizado"].apply(
        lambda x: "Aguardando" if x else "Número inválido"
    )
    dados["Detalhe"] = dados["Telefone_Normalizado"].apply(
        lambda x: "" if x else "Formato de telefone inválido"
    )
    dados["Call_SID"] = ""
    dados["Data_Hora"] = ""

    return dados.reset_index(drop=True)

def gerar_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    buffer.seek(0)
    return buffer

def criar_twiml(mensagem):
    texto = html.escape(str(mensagem))
    return f"""<Response>
    <Say language="pt-BR">{texto}</Say>
    <Pause length="1"/>
    </Response>"""

def status_final(client, sid, timeout=45):
    inicio = time.time()
    ultimo = "queued"
    while time.time() - inicio < timeout:
        chamada = client.calls(sid).fetch()
        ultimo = chamada.status
        if ultimo in ["completed", "busy", "failed", "no-answer", "canceled"]:
            return ultimo
        time.sleep(2)
    return ultimo

st.title("📞 Discador Automático")
st.caption("Chamadas reais com Twilio e relatório em Excel")

try:
    account_sid = st.secrets["TWILIO_ACCOUNT_SID"]
    auth_token = st.secrets["TWILIO_AUTH_TOKEN"]
    twilio_number = st.secrets["TWILIO_PHONE_NUMBER"]
    client = Client(account_sid, auth_token)
except Exception:
    st.error("Credenciais do Twilio não encontradas nos Secrets do Streamlit.")
    st.stop()

st.info(
    "Na conta Trial, o Twilio normalmente permite chamadas apenas para números verificados. "
    "Use inicialmente o seu próprio telefone para testar."
)

arquivo = st.file_uploader("Carregue sua planilha Excel", type=["xlsx", "xls"])

if arquivo is not None:
    try:
        arquivo_id = (arquivo.name, arquivo.size)

        if st.session_state.get("arquivo_id") != arquivo_id:
            st.session_state.abas = carregar_planilha(arquivo.getvalue())
            st.session_state.arquivo_id = arquivo_id
            st.session_state.aba_atual = None
            st.session_state.dados = None
            st.session_state.campanha_iniciada = False

        aba = st.selectbox("Selecione a aba com os contatos", list(st.session_state.abas.keys()))

        if st.session_state.aba_atual != aba:
            st.session_state.dados = preparar_dados(st.session_state.abas[aba])
            st.session_state.aba_atual = aba

        dados = st.session_state.dados

        st.subheader("Pré visualização")
        tabela = st.empty()
        tabela.dataframe(dados, use_container_width=True, hide_index=True)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total", len(dados))
        c2.metric("Válidos", int(dados["Telefone_Normalizado"].notna().sum()))
        c3.metric("Inválidos", int(dados["Telefone_Normalizado"].isna().sum()))

        confirmacao = st.checkbox(
            "Confirmo que tenho autorização para realizar contato com estes números."
        )

        if st.button("📞 Iniciar ligações reais", use_container_width=True, disabled=not confirmacao):
            resultado = dados.copy()
            validos = resultado.index[resultado["Status"] == "Aguardando"].tolist()

            if not validos:
                st.warning("Não existem números válidos.")
            else:
                progresso = st.progress(0)
                status_box = st.empty()

                for posicao, indice in enumerate(validos, start=1):
                    telefone = resultado.at[indice, "Telefone_Normalizado"]
                    mensagem = resultado.at[indice, "Mensagem"]

                    try:
                        status_box.info(
                            f"📞 Ligando para {telefone} | {posicao} de {len(validos)}"
                        )

                        chamada = client.calls.create(
                            to=telefone,
                            from_=twilio_number,
                            twiml=criar_twiml(mensagem)
                        )

                        resultado.at[indice, "Call_SID"] = chamada.sid
                        resultado.at[indice, "Status"] = "Ligação iniciada"
                        resultado.at[indice, "Detalhe"] = "Aguardando resultado da chamada"
                        resultado.at[indice, "Data_Hora"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        tabela.dataframe(resultado, use_container_width=True, hide_index=True)

                        final = status_final(client, chamada.sid)

                        mapa = {
                            "completed": ("Concluída", "A chamada foi concluída"),
                            "busy": ("Ocupado", "O número estava ocupado"),
                            "failed": ("Falha", "A chamada falhou"),
                            "no-answer": ("Sem resposta", "A chamada não foi atendida"),
                            "canceled": ("Cancelada", "A chamada foi cancelada")
                        }

                        if final in mapa:
                            resultado.at[indice, "Status"] = mapa[final][0]
                            resultado.at[indice, "Detalhe"] = mapa[final][1]
                        else:
                            resultado.at[indice, "Status"] = "Em andamento"
                            resultado.at[indice, "Detalhe"] = f"Status atual: {final}"

                    except Exception as e:
                        resultado.at[indice, "Status"] = "Erro"
                        resultado.at[indice, "Detalhe"] = str(e)

                    progresso.progress(posicao / len(validos))
                    tabela.dataframe(resultado, use_container_width=True, hide_index=True)

                st.session_state.dados = resultado
                status_box.success("Campanha finalizada.")

        resultado = st.session_state.dados

        st.subheader("Resultado")
        st.dataframe(resultado, use_container_width=True, hide_index=True)

        st.download_button(
            "📥 Baixar relatório Excel",
            data=gerar_excel(resultado),
            file_name=f"resultado_discador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as erro:
        st.error(f"Erro ao processar a planilha: {erro}")
else:
    st.info("Coluna A: telefone | Coluna B: mensagem")
