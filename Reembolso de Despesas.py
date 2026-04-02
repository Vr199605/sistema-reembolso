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

# ReportLab para o layout
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# Configuração da página
st.set_page_config(page_title="Sistema de Reembolso", layout="wide")

# Criar pasta para anexos se não existir
if not os.path.exists("temp_anexos"):
    os.makedirs("temp_anexos")

# --- CONEXÃO COM GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"Erro na conexão com Planilha: {e}")

# --- FUNÇÃO GERAR PDF (REPORTLAB) ---
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
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), True),
    ]))
    elements.append(t)
    
    elements.append(Spacer(1, 25))
    elements.append(Paragraph(f"<b>VALOR TOTAL A REEMBOLSAR: R$ {total:.2f}</b>", styles['Heading2']))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- FUNÇÕES DE E-MAIL (MODIFICADA PARA MÚLTIPLOS ANEXOS) ---
def enviar_email_com_pdf(destinatario, assunto, corpo, pdf_buffer=None, arquivos_extras=None):
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

    # Anexa os comprovantes originais
    if arquivos_extras:
        for caminho in arquivos_extras:
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

# --- FUNÇÃO DE RESET ---
def reset_campos():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# --- INTERFACE ---
aba_guia, aba_solicitacao, aba_aprovacao = st.tabs(["📖 Guia Passo a Passo", "📋 Solicitação", "🔐 Aprovação (Gabriel Coelho)"])

with aba_guia:
    st.header("📖 Guia de Preenchimento de Reembolso")
    st.info("Siga os passos abaixo para garantir que sua solicitação seja processada sem erros.")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 1️⃣ Identificação\nPreencha seu nome e data.\n### 2️⃣ Seleção de Categorias\n### 3️⃣ Detalhamento por Item")
    with col2:
        st.markdown("### 4️⃣ Comprovantes\n### 5️⃣ Envio\n### 6️⃣ Data de Pagamento")
    st.markdown("---")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    try:
        with open(caminho_manual, "rb") as f:
            st.download_button(label="📥 Baixar Manual de Reembolso (PDF)", data=f, file_name="manual_reembolso.pdf", mime="application/pdf")
    except: pass

with aba_solicitacao:
    st.title("🚀 Solicitação de Reembolso de Despesas")
    st.markdown("---")
    col_perfil1, col_perfil2 = st.columns(2)
    with col_perfil1: nome = st.text_input("Nome completo", key="nome_user")
    with col_perfil2: data_solicitacao = st.date_input("Data de solicitação", format="DD/MM/YYYY", key="data_sol")

    categorias_disponiveis = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em qtde)", "KM (em qtde)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]
    selecionadas = st.multiselect("Selecione as categorias:", categorias_disponiveis)
    dados_despesas = []

    if selecionadas:
        for cat in selecionadas:
            with st.container():
                c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
                c1.markdown(f"**{cat}**")
                d_desp = c2.date_input(f"Data", format="DD/MM/YYYY", key=f"d_{cat}")
                if "KM (em qtde)" in cat:
                    q_km = c3.number_input("Qtde KM", min_value=0.0, step=0.1, value=None, key=f"v_{cat}")
                    v_fin = (q_km * 1.37) if q_km else 0.0
                else:
                    v_fin = c3.number_input("Valor R$", min_value=0.0, step=0.01, value=None, key=f"v_{cat}")
                    v_fin = v_fin if v_fin else 0.0
                mot = c4.text_input("Motivo *", key=f"m_{cat}")
                dados_despesas.append({"Data": d_desp.strftime('%d/%m/%Y'), "Categoria": cat, "Valor Total": v_fin, "Motivo": mot})
        
        st.subheader("Anexar Comprovantes")
        arq = st.file_uploader("Upload (Obrigatório) *", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

        col_btn1, col_btn2 = st.columns([2, 8])
        with col_btn1:
            if st.button("Enviar Solicitação", use_container_width=True):
                if any(not d["Motivo"].strip() for d in dados_despesas) or not arq or not nome:
                    st.error("Preencha todos os campos.")
                else:
                    try:
                        # Salva arquivos fisicamente para o Gabriel acessar
                        nomes_arquivos = []
                        for f in arq:
                            caminho_f = os.path.join("temp_anexos", f"{nome}_{f.name}")
                            with open(caminho_f, "wb") as buffer_f:
                                buffer_f.write(f.getbuffer())
                            nomes_arquivos.append(caminho_f)

                        df_p = pd.DataFrame(dados_despesas)
                        df_p['Colaborador'] = nome
                        df_p['Data Solicitacao'] = data_solicitacao.strftime('%d/%m/%Y')
                        df_p['Anexos'] = ",".join(nomes_arquivos) # Salva o caminho na planilha

                        existing = conn.read(worksheet="Pendentes")
                        conn.update(worksheet="Pendentes", data=pd.concat([existing, df_p], ignore_index=True))
                        
                        enviar_email_com_pdf("victormoreiraicnv@gmail.com", f"Solicitação: {nome}", f"Nova solicitação de {nome}. Verifique a aba de Aprovação.")
                        st.success("Enviado!")
                        time.sleep(2)
                        reset_campos()
                    except Exception as e: st.error(f"Erro: {e}")
        with col_btn2:
            if st.button("🗑️ Limpar Tudo"): reset_campos()

with aba_aprovacao:
    st.title("🔐 Área de Verificação")
    if st.text_input("Senha", type="password") == "12345":
        try:
            df_p = conn.read(worksheet="Pendentes", ttl=0)
            if not df_p.empty:
                colab_sel = st.selectbox("Selecione para analisar:", df_p['Colaborador'].unique())
                dados_f = df_p[df_p['Colaborador'] == colab_sel]
                
                # MOSTRAR ANEXOS PARA O GABRIEL
                st.subheader("📁 Comprovantes Enviados")
                lista_anexos = dados_f.iloc[0]['Anexos'].split(",") if 'Anexos' in dados_f.columns else []
                if lista_anexos:
                    cols_anexos = st.columns(len(lista_anexos))
                    for idx, path in enumerate(lista_anexos):
                        if os.path.exists(path):
                            with open(path, "rb") as file_anexo:
                                cols_anexos[idx].download_button(label=f"Baixar Anexo {idx+1}", data=file_anexo, file_name=os.path.basename(path))
                
                dados_ajustados = []
                for i, row in dados_f.iterrows():
                    with st.container():
                        c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
                        c1.markdown(f"**{row['Categoria']}**")
                        adj_data = c2.text_input("Data", value=row['Data'], key=f"adj_d_{i}")
                        adj_val = c3.number_input("Valor R$", value=float(row['Valor Total']), key=f"adj_v_{i}")
                        adj_mot = c4.text_input("Motivo", value=row['Motivo'], key=f"adj_m_{i}")
                        dados_ajustados.append({"Data": adj_data, "Categoria": row['Categoria'], "Valor Total": adj_val, "Motivo": adj_mot})
                
                total_adj = sum(d["Valor Total"] for d in dados_ajustados)
                st.metric("Total Final", f"R$ {total_adj:.2f}")

                if st.button("✅ Aprovar e Enviar Tudo por E-mail"):
                    # Salva Oficial
                    df_fin = pd.DataFrame(dados_ajustados)
                    df_fin['Colaborador'] = colab_sel
                    df_fin['Data Solicitacao'] = dados_f.iloc[0]['Data Solicitacao']
                    existing_of = conn.read(worksheet="Reembolsos")
                    conn.update(worksheet="Reembolsos", data=pd.concat([existing_of, df_fin], ignore_index=True))
                    
                    # Remove Pendente
                    conn.update(worksheet="Pendentes", data=df_p[df_p['Colaborador'] != colab_sel])
                    
                    # Envia PDF + Anexos Originais
                    pdf = gerar_pdf(colab_sel, df_fin['Data Solicitacao'].iloc[0], dados_ajustados, total_adj)
                    enviar_email_com_pdf("victormoreiraicnv@gmail.com", f"APROVADO - {colab_sel}", "Relatório e comprovantes em anexo.", pdf, lista_anexos)
                    st.success("Tudo enviado com sucesso!")
                    time.sleep(2)
                    st.rerun()
            else: st.info("Sem pendências.")
        except: st.info("Aguardando solicitações...")
