# laudo_app_v2.py - Versão corrigida
import streamlit as st
import re
from datetime import datetime
import docx
from docx import Document
from docx.shared import Pt
import io
import traceback

# ========== CONSTANTES ==========
TIPOS_MATERIAL_BASE = {
    "v": "vegetal dessecado",
    "po": "pulverizado", 
    "pd": "petrificado",
    "r": "resinoso"
}

TIPOS_EMBALAGEM_BASE = {
    "e": "microtubo do tipo eppendorf",
    "z": "embalagem do tipo zip",
    "a": "papel alumínio",
    "pl": "plástico",
    "pa": "papel"
}

# ========== CONFIGURAÇÃO DA PÁGINA ==========
st.set_page_config(
    page_title="Gerador de Laudos",
    page_icon="🔍",
    layout="wide"
)

# ========== CSS ==========
st.markdown("""
<style>
    :root {
        --primary: #1DA1F2;
        --background: #0E1117;
        --card: #1E293B;
        --text: #F8FAFC;
        --border: #334155;
    }
    
    .main {
        background-color: var(--background);
        color: var(--text);
    }
    
    .stTextInput>div>div>input,
    .stNumberInput>div>div>input,
    .stSelectbox>div>div>select {
        background-color: var(--card) !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
    }
    
    .stButton>button {
        background-color: var(--primary) !important;
        color: white !important;
        border-radius: 8px !important;
    }
    
    h1, h2, h3 {
        color: var(--primary) !important;
    }
</style>
""", unsafe_allow_html=True)

# ========== CABEÇALHO ==========
st.markdown("""
<div style="background: linear-gradient(90deg, #1E40AF, #1E3A8A);
            padding: 1.5rem;
            border-radius: 0.5rem;
            margin-bottom: 1.5rem;">
    <h1 style="color:white; text-align:center; margin:0;">GERADOR DE LAUDOS PERICIAIS</h1>
    <p style="color:#E0F2FE; text-align:center; margin:0.5rem 0 0;">
        Sistema Oficial
    </p>
</div>
""", unsafe_allow_html=True)

# ========== FUNÇÕES ==========
def gerar_documento(itens, lacre_num):
    doc = Document()
    doc.add_heading('LAUDO PERICIAL', 0)
    
    # Seção de materiais
    p = doc.add_paragraph()
    p.add_run("2 MATERIAL RECEBIDO PARA EXAME").bold = True
    
    for idx, item in enumerate(itens, start=1):
        desc = f"2.{idx} {item['quantidade']} porções de {TIPOS_MATERIAL_BASE[item['tipo_material']]}"  # CORREÇÃO AQUI
        doc.add_paragraph(desc)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# ========== FORMULÁRIO ==========
with st.form("form_laudo"):
    lacre_num = st.text_input("Número do Lacre", placeholder="LC-2023-XXXXX")
    num_itens = st.number_input("Número de Itens", min_value=1, value=1)
    
    itens = []
    for i in range(num_itens):
        with st.expander(f"Item {i+1}"):
            cols = st.columns([1, 2, 2])
            with cols[0]:
                qtd = st.number_input("Quantidade", key=f"q{i}_qtd", min_value=1, value=1)
            with cols[1]:
                tipo_mat = st.selectbox(
                    "Tipo Material",
                    options=list(TIPOS_MATERIAL_BASE.keys()),
                    format_func=lambda x: TIPOS_MATERIAL_BASE[x],
                    key=f"q{i}_mat"
                )
            with cols[2]:
                tipo_emb = st.selectbox(
                    "Embalagem",
                    options=list(TIPOS_EMBALAGEM_BASE.keys()),
                    format_func=lambda x: TIPOS_EMBALAGEM_BASE[x],
                    key=f"q{i}_emb"
                )
            
            itens.append({
                "quantidade": qtd,
                "tipo_material": tipo_mat,
                "tipo_embalagem_base": tipo_emb
            })
    
    if st.form_submit_button("Gerar Laudo"):
        if not lacre_num:
            st.error("Informe o número do lacre!")
        else:
            with st.spinner("Gerando documento..."):
                try:
                    doc_bytes = gerar_documento(itens, lacre_num)
                    st.success("Laudo gerado com sucesso!")
                    st.download_button(
                        label="⬇️ Baixar Laudo",
                        data=doc_bytes.getvalue(),
                        file_name="laudo_pericial.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Erro: {str(e)}")
                    st.text(traceback.format_exc())
