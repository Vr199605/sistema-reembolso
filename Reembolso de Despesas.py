import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
import io
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from streamlit_gsheets import GSheetsConnection
import time

# ReportLab para o layout do PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema de Reembolso", layout="wide", page_icon="💰")

# Criar pastas necessárias
for pasta in ["comprovantes_servidor", "documentos"]:
    if not os.path.exists(pasta):
        os.makedirs(pasta)

# --- CONEXÃO COM GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"Erro na conexão com Planilha: {e}")

# --- FUNÇÃO GERAR PDF ---
def gerar_pdf(nome, data_sol, dados_tabela, total):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph(f"Relatório de Reembolso de Despesas", styles['Title']))
    elements.append(Spacer(1, 24))
    
    elements.append(Paragraph(f"<b>Colaborador:</b> {nome}", styles['Normal']))
    elements.append(Paragraph(f"<b>Data da Solicitação:</b> {data_sol}", styles['Normal']))
    elements.append(Spacer(1, 20))

    data = [["Data", "Categoria", "Motivo", "Valor (R$)"]]
    for item in dados_tabela:
        data.append([item['Data'], item['Categoria'], item['Motivo'], f"{item['Valor Total']:.2f}"])
    
    t = Table(data, colWidths=[80, 150, 220, 85])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ]))
    elements.append(t)
    
    elements.append(Spacer(1, 25))
    elements.append(Paragraph(f"<b>VALOR TOTAL A REEMBOLSAR: R$ {total:.2f}</b>", styles['Heading2']))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- FUNÇÕES DE E-MAIL ---
def enviar_email_com_pdf(destinatario, assunto, corpo, pdf_buffer=None, caminhos_anexos=None):
    seu_email = "victormoreiraicnv@gmail.com"
    senha_app = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))
    
    if pdf_buffer:
        part = MIMEApplication(pdf_buffer.read(), Name="reembolso_aprovado.pdf")
        part['Content-Disposition'] = 'attachment; filename="reembolso_aprovado.pdf"'
        msg.attach(part)

    if caminhos_anexos:
        for caminho in caminhos_anexos:
            if os.path.exists(caminho):
                with open(caminho, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(caminho))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho)}"'
                    msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(seu_email, senha_app)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def reset_campos():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# --- INTERFACE ---
aba_guia, aba_solicitacao, aba_aprovacao = st.tabs(["📖 Guia", "📋 Solicitação", "🔐 Aprovação"])

with aba_guia:
    st.header("📖 Guia de Reembolso")
    st.markdown("""
    1. Preencha seu nome e data.
    2. Selecione as categorias (pode escolher várias).
    3. Informe o motivo e o valor (ou KM).
    4. **Obrigatório:** Anexe as fotos dos comprovantes.
    5. O pagamento ocorre em **D+5** após a aprovação.
    """)

with aba_solicitacao:
    st.title("🚀 Nova Solicitação")
    col1, col2 = st.columns(2)
    nome = col1.text_input("Nome completo", key="n_user")
    data_sol = col2.date_input("Data de solicitação", format="DD/MM/YYYY")

    cats = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em qtde)", "KM (em qtde)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]
    selecionadas = st.multiselect("Categorias:", cats)
    
    dados_envio = []
    if selecionadas:
        for cat in selecionadas:
            with st.expander(f"Dados de {cat}", expanded=True):
                c1, c2, c3 = st.columns([2, 2, 4])
                d_item = c1.date_input("Data despesa", key=f"d_{cat}")
                if "KM" in cat:
                    km = c2.number_input("Qtde KM", min_value=0.0, key=f"v_{cat}")
                    v_item = km * 1.37
                    c2.info(f"R$ {v_item:.2f}")
                else:
                    v_item = c2.number_input("Valor R$", min_value=0.0, key=f"v_{cat}")
                motivo = c3.text_input("Motivo *", key=f"m_{cat}")
                dados_envio.append({"Data": d_item.strftime('%d/%m/%Y'), "Categoria": cat, "Valor Total": float(v_item or 0), "Motivo": motivo})

        arq = st.file_uploader("Anexar Comprovantes (Imagens ou PDF) *", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

        if st.button("Enviar Solicitação"):
            if not nome or not arq or any(not d["Motivo"] for d in dados_envio):
                st.error("Preencha todos os campos e anexe os comprovantes.")
            else:
                with st.spinner("Enviando..."):
                    try:
                        # Salvar arquivos localmente
                        salvos = []
                        for f in arq:
                            caminho = os.path.join("comprovantes_servidor", f"{nome}_{f.name}")
                            with open(caminho, "wb") as b: b.write(f.getbuffer())
                            salvos.append(caminho)

                        # Preparar DataFrame e enviar para o Sheets
                        df_novo = pd.DataFrame(dados_envio)
                        df_novo['Colaborador'] = nome
                        df_novo['Data Solicitacao'] = data_sol.strftime('%d/%m/%Y')
                        df_novo['Caminhos_Anexos'] = "|".join(salvos)

                        existing = conn.read(worksheet="Pendentes")
                        # Limpeza para evitar Erro 500 (garantir que tipos batem)
                        updated_df = pd.concat([existing, df_novo], ignore_index=True).fillna("")
                        conn.update(worksheet="Pendentes", data=updated_df)
                        
                        enviar_email_com_pdf("gabriel.coelho@globusseguros.com.br", f"Nova Solicitação: {nome}", f"Verifique o sistema.")
                        st.success("Solicitação enviada com sucesso!")
                        time.sleep(2)
                        reset_campos()
                    except Exception as e:
                        st.error(f"Erro ao enviar: {e}")

with aba_aprovacao:
    st.title("🔐 Área de Aprovação")
    if st.text_input("Senha de Acesso", type="password") == "globus2026":
        try:
            df_pend = conn.read(worksheet="Pendentes", ttl=0)
            if not df_pend.empty:
                colab = st.selectbox("Selecione o Colaborador:", df_pend['Colaborador'].unique())
                dados_c = df_pend[df_pend['Colaborador'] == colab]
                
                # Exibir anexos
                anexos_str = str(dados_c.iloc[0]['Caminhos_Anexos'])
                lista_anexos = anexos_str.split("|") if anexos_str else []
                
                st.write("---")
                cols_anexo = st.columns(max(len(lista_anexos), 1))
                for idx, path in enumerate(lista_anexos):
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            cols_anexo[idx].download_button(f"📄 Ver Anexo {idx+1}", f, file_name=os.path.basename(path))

                # Ajustes Finais
                st.subheader("Conferência de Valores")
                ajustados = []
                for i, row in dados_c.iterrows():
                    c1, c2, c3 = st.columns([3, 2, 4])
                    c1.markdown(f"**{row['Categoria']}**")
                    v_adj = c2.number_input("Valor", value=float(row['Valor Total']), key=f"adj_v_{i}")
                    m_adj = c3.text_input("Motivo", value=row['Motivo'], key=f"adj_m_{i}")
                    ajustados.append({"Data": row['Data'], "Categoria": row['Categoria'], "Valor Total": v_adj, "Motivo": m_adj})
                
                total = sum(d['Valor Total'] for d in ajustados)
                st.metric("Total a Pagar", f"R$ {total:.2f}")

                if st.button("✅ Aprovar e Notificar"):
                    df_final = pd.DataFrame(ajustados)
                    df_final['Colaborador'] = colab
                    df_final['Data Solicitacao'] = dados_c.iloc[0]['Data Solicitacao']
                    
                    # Salvar no histórico e remover das pendências
                    hist = conn.read(worksheet="Reembolsos")
                    conn.update(worksheet="Reembolsos", data=pd.concat([hist, df_final], ignore_index=True).fillna(""))
                    conn.update(worksheet="Pendentes", data=df_pend[df_pend['Colaborador'] != colab].fillna(""))
                    
                    # Gerar e enviar PDF
                    pdf = gerar_pdf(colab, df_final['Data Solicitacao'].iloc[0], ajustados, total)
                    enviar_email_com_pdf("gabriel.coelho@globusseguros.com.br", f"APROVADO - {colab}", "Relatório gerado automaticamente.", pdf, lista_anexos)
                    
                    st.success("Aprovação concluída!")
                    time.sleep(2)
                    st.rerun()
            else:
                st.info("Não há solicitações pendentes.")
        except Exception as e:
            st.warning("Aguardando dados da planilha...")
