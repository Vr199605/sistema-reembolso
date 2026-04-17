import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
import io
import os
import re
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

# --- BLOCO PARA OCULTAR ÍCONES E MENU ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# Criar pasta para anexos no servidor
if not os.path.exists("comprovantes_servidor"):
    os.makedirs("comprovantes_servidor")

# --- CONEXÃO COM GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"Erro na conexão com Planilha: {e}")

# --- FUNÇÃO PARA BUSCA INTELIGENTE ---
@st.cache_data(ttl=600)
def carregar_base_funcionarios():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQWoetqtPPgSLJu3bBzYNo8Avaa3DGCsenQ1yrtzYrdU48J9-SzK8gkHkCyAk6L1fJkPyCgFxKdO9Se/pub?output=csv"
    try:
        return pd.read_csv(url)
    except:
        return pd.DataFrame()

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
            caminho_norm = os.path.normpath(caminho.replace("\\", "/"))
            if os.path.exists(caminho_norm):
                with open(caminho_norm, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(caminho_norm))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_norm)}"'
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
aba_guia, aba_solicitacao, aba_aprovacao = st.tabs(["📖 Guia Passo a Passo", "📋 Solicitação", "🔐 Aprovação (Gabriel Coelho)"])

with aba_guia:
    st.header("📖 Guia de Preenchimento de Reembolso")
    st.info("Siga os passos abaixo para garantir que sua solicitação seja processada sem erros.")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        ### 1️⃣ Identificação
        Preencha seu **nome completo** e a **data atual** da solicitação. 
        ### 2️⃣ Seleção de Categorias
        Escolha as categorias que compõem sua despesa.
        ### 3️⃣ Detalhamento por Item
        Para cada categoria, informe a data, o valor/quantidade e o motivo.
        """)
    with col2:
        st.markdown("""
        ### 4️⃣ Comprovantes
        O upload de arquivos é **obrigatório**. 
        ### 5️⃣ Envio
        Clique em **Enviar Solicitação**.
        ### 6️⃣ Data de Pagamento
        Pagamento em **D+5** após a análise.
        """)
    
    st.markdown("---")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    try:
        if not os.path.exists("documentos"): os.makedirs("documentos")
        with open(caminho_manual, "rb") as f:
            st.download_button(label="📥 Baixar Manual de Reembolso (PDF)", data=f, file_name="manual_reembolso.pdf", mime="application/pdf")
    except: pass

with aba_solicitacao:
    st.title("🚀 Solicitação de Reembolso de Despesas")
    st.markdown("---")
    df_base = carregar_base_funcionarios()
    lista_nomes = sorted(df_base['Nome do Funcionário'].dropna().unique().tolist()) if not df_base.empty else []

    col_perfil1, col_perfil2 = st.columns(2)
    with col_perfil1: 
        nome = st.selectbox("Selecione seu Nome Completo", options=[""] + lista_nomes, key="nome_user")
    with col_perfil2: 
        data_solicitacao = st.date_input("Data de solicitação", format="DD/MM/YYYY", key="data_sol")

    centro_custo, setor, departamento = "", "", ""
    if nome != "":
        dados_func = df_base[df_base['Nome do Funcionário'] == nome].iloc[0]
        centro_custo = dados_func['Centro de Custo']
        setor = dados_func['SETOR']
        departamento = dados_func['DEPARTAMENTO']
        st.success(f"📌 **Dados Identificados:** Setor: {setor} | CC: {centro_custo}")

    categorias_disponiveis = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em qtde)", "KM (em qtde)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]
    
    if 'lista_categorias' not in st.session_state:
        st.session_state.lista_categorias = []

    col_cat, col_add = st.columns([8, 2])
    cat_selecionada = col_cat.selectbox("Escolha uma categoria para adicionar:", [""] + categorias_disponiveis)
    if col_add.button("➕ Adicionar"):
        if cat_selecionada:
            st.session_state.lista_categorias.append(cat_selecionada)

    dados_despesas = []
    if st.session_state.lista_categorias:
        for idx, cat in enumerate(st.session_state.lista_categorias):
            with st.container():
                c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 4, 1])
                c1.markdown(f"**{cat}**")
                d_desp = c2.date_input(f"Data", format="DD/MM/YYYY", key=f"d_{cat}_{idx}")
                if "KM (em qtde)" in cat:
                    q_km = c3.number_input("Qtde KM", min_value=0.0, step=0.1, value=None, key=f"v_{cat}_{idx}")
                    v_fin = (float(q_km) * 1.37) if q_km else 0.0
                    if q_km: c3.info(f"R$ {v_fin:.2f}")
                else:
                    v_fin = c3.number_input("Valor R$", min_value=0.0, step=0.01, value=None, key=f"v_{cat}_{idx}")
                    v_fin = v_fin if v_fin else 0.0
                mot = c4.text_input("Motivo *", key=f"m_{cat}_{idx}")
                if c5.button("🗑️", key=f"del_{idx}"):
                    st.session_state.lista_categorias.pop(idx)
                    st.rerun()
                dados_despesas.append({"Data": d_desp.strftime('%d/%m/%Y'), "Categoria": cat, "Valor Total": float(v_fin), "Motivo": mot})
        
        st.markdown("---")
        total_solicitacao = sum(d["Valor Total"] for d in dados_despesas)
        st.subheader(f"💰 Total da Solicitação: R$ {total_solicitacao:.2f}")

        arq = st.file_uploader("Upload de Comprovantes (Obrigatório) *", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
        if st.button("Enviar Solicitação", use_container_width=True):
            if any(not d["Motivo"].strip() for d in dados_despesas) or not arq or nome == "":
                st.error("Preencha todos os campos e anexe os comprovantes!")
            else:
                st.session_state.confirmar_envio = True

        if st.session_state.get('confirmar_envio'):
            if st.button("Confirmar Envio"):
                try:
                    caminhos_salvos = []
                    for f in arq:
                        caminho = os.path.join("comprovantes_servidor", f"{nome}_{f.name}")
                        with open(caminho, "wb") as b: b.write(f.getbuffer())
                        caminhos_salvos.append(caminho)

                    df_p = pd.DataFrame(dados_despesas)
                    df_p['Colaborador'] = nome
                    df_p['Data Solicitacao'] = data_solicitacao.strftime('%d/%m/%Y')
                    df_p['Caminhos_Anexos'] = "|".join(caminhos_salvos)
                    df_p['SETOR'] = setor
                    df_p['DEPARTAMENTO'] = departamento
                    df_p['Centro de Custo'] = centro_custo

                    existing = conn.read(worksheet="Pendentes").astype(str)
                    combined = pd.concat([existing, df_p.astype(str)], ignore_index=True).replace("nan", "")
                    conn.update(worksheet="Pendentes", data=combined)
                    
                    enviar_email_com_pdf("gabriel.coelho@globusseguros.com.br", f"Solicitação: {nome}", "Nova solicitação disponível.", caminhos_anexos=caminhos_salvos)
                    st.success("Enviado com sucesso!")
                    time.sleep(2)
                    reset_campos()
                except Exception as e: st.error(f"Erro: {e}")

with aba_aprovacao:
    st.title("🔐 Área de Verificação")
    if st.text_input("Senha", type="password") == "globus2026":
        try:
            df_pend = conn.read(worksheet="Pendentes", ttl=0)
            if not df_pend.empty:
                colab_sel = st.selectbox("Escolha o colaborador:", df_pend['Colaborador'].unique())
                dados_f = df_pend[df_pend['Colaborador'] == colab_sel]
                
                string_anexos = str(dados_f.iloc[0]['Caminhos_Anexos'])
                lista_anexos = [p.strip() for p in string_anexos.split("|") if p.strip() and p.strip() != "nan"]
                
                if lista_anexos:
                    for i, p in enumerate(lista_anexos):
                        if os.path.exists(p):
                            with open(p, "rb") as f_down:
                                st.download_button(f"📄 Anexo {i+1}", f_down, os.path.basename(p), key=f"dl_{i}")
                
                dados_ajustados = []
                for i, row in dados_f.iterrows():
                    with st.container():
                        c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
                        c1.markdown(f"**{row['Categoria']}**")
                        adj_data = c2.text_input("Data", value=row['Data'], key=f"adj_d_{i}")
                        
                        # --- CORREÇÃO DO VALOR (SIMPLIFICADO) ---
                        try:
                            # Converte o valor da planilha para float, tratando vírgula se houver
                            val_limpo = float(str(row['Valor Total']).replace(',', '.'))
                        except:
                            val_limpo = 0.0

                        adj_val = c3.number_input("Valor R$", value=val_limpo, format="%.2f", key=f"adj_v_{i}")
                        adj_mot = c4.text_input("Motivo", value=row['Motivo'], key=f"adj_m_{i}")
                        dados_ajustados.append({"Data": adj_data, "Categoria": row['Categoria'], "Valor Total": float(adj_val), "Motivo": adj_mot})
                
                total_adj = sum(d["Valor Total"] for d in dados_ajustados)
                st.metric("Total Final", f"R$ {total_adj:.2f}")

                if st.button("✅ Aprovar"):
                    df_fin = pd.DataFrame(dados_ajustados)
                    df_fin['Colaborador'] = colab_sel
                    df_fin['Data Solicitacao'] = dados_f.iloc[0]['Data Solicitacao']
                    ex_of = conn.read(worksheet="Reembolsos").astype(str)
                    conn.update(worksheet="Reembolsos", data=pd.concat([ex_of, df_fin.astype(str)], ignore_index=True))
                    conn.update(worksheet="Pendentes", data=df_pend[df_pend['Colaborador'] != colab_sel].astype(str))
                    
                    pdf = gerar_pdf(colab_sel, df_fin['Data Solicitacao'].iloc[0], dados_ajustados, total_adj)
                    enviar_email_com_pdf("gabriel.coelho@globusseguros.com.br", f"APROVADO - {colab_sel}", "Relatório em anexo.", pdf, lista_anexos)
                    st.success("Aprovado!")
                    time.sleep(2)
                    st.rerun()
                
                if st.button("❌ Reprovar"):
                    conn.update(worksheet="Pendentes", data=df_pend[df_pend['Colaborador'] != colab_sel].astype(str))
                    st.error("Removido.")
                    time.sleep(1)
                    st.rerun()
        except Exception as e: st.info("Sem pendências.")
