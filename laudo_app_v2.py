import streamlit as st
import re
from datetime import datetime
import docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import os
import io
import shutil
import time
import traceback

# ========== CONSTANTES ==========
TIPOS_MATERIAL_BASE = {
    "v": "vegetal dessecado",
    "po": "pulverizado",
    "pd": "petrificado",
    "r": "resinoso"
}

TONALIDADES_GENERICAS = {
    "b": "esbranquiçada", "a": "amarelada", "vd": "esverdeada",
    "vr": "avermelhada", "az": "azulada", "p": "enegrecida",
    "c": "acinzentada", "m": "amarronzada", "r": "arrosada",
    "l": "alaranjada", "violeta": "arroxeadada"
}

TIPOS_EMBALAGEM_BASE = {
    "e": "microtubo do tipo “eppendorf”",
    "z": "embalagem do tipo \"zip\"",
    "a": "papel alumínio",
    "pl": "plástico",
    "pa": "papel"
}

CORES_FEMININO_EMBALAGEM = {
    "t": "transparente",
    "branco": "branca", "branca": "branca", "b": "branca",
    "azul": "azul", "az": "azul",
    "amarelo": "amarela", "amarela": "amarela", "am": "amarela",
    "vermelho": "vermelha", "vermelha": "vermelha", "vm": "vermelha",
    "verde": "verde", "vd": "verde",
    "preto": "preta", "preta": "preta", "p": "preta",
    "cinza": "cinza", "c": "cinza",
    "marrom": "marrom", "m": "marrom",
    "rosa": "rosa", "r": "rosa",
    "laranja": "laranja", "l": "laranja",
    "violeta": "violeta",
    "roxa": "roxa"
}

QUANTIDADES_EXTENSO = {
    1: "uma", 2: "duas", 3: "três", 4: "quatro", 5: "cinco",
    6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez"
}

meses_portugues = {
    "January": "janeiro", "February": "fevereiro", "March": "março",
    "April": "abril", "May": "maio", "June": "junho", "July": "julho",
    "August": "agosto", "September": "setembro", "October": "outubro",
    "November": "novembro", "December": "dezembro"
}

# ========== FUNÇÕES AUXILIARES ==========
def add_formatted_paragraph(document, text, font_name='Arial', font_size=12, 
                          alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, 
                          is_bold=False, is_italic=False, space_after=Pt(0)):
    paragraph = document.add_paragraph()
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_after = space_after
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = is_bold
    run.italic = is_italic
    return paragraph

def pluralizar_palavra(palavra, quantidade):
    if quantidade == 1 or not isinstance(palavra, str):
        return palavra
    if palavra in ["microtubo do tipo “eppendorf”", "embalagem do tipo \"zip\""]:
        return palavra
    if palavra.endswith('m'): return re.sub(r'm$', 'ns', palavra)
    if palavra.endswith(('ão', 'ões')): return palavra
    if palavra.endswith(('al', 'el', 'ol', 'ul')): return palavra[:-1] + 'is'
    if palavra.endswith(('r', 'z', 's')): return palavra + 'es'
    return palavra + 's'

def gerar_descricao_item_web(numero_item_str, item_data):
    try:
        quantidade = item_data['quantidade']
        tipo_mat = TIPOS_MATERIAL_BASE[item_data['tipo_material']]
        embalagem = TIPOS_EMBALAGEM_BASE[item_data['tipo_embalagem_base']]
        
        # Tratamento da cor
        if item_data['tipo_embalagem_base'] in ["pl", "pa"] and item_data['cor_embalagem']:
            cor = CORES_FEMININO_EMBALAGEM.get(item_data['cor_embalagem'], item_data['cor_embalagem'])
            embalagem = f"{embalagem} de cor {cor}"
        
        desc = (f"{numero_item_str} {quantidade} ({QUANTIDADES_EXTENSO.get(quantidade, str(quantidade))}) "
                f"{pluralizar_palavra('porção', quantidade)} de material {tipo_mat}, "
                f"{'acondicionada em' if quantidade == 1 else 'acondicionadas, individualmente, em'} "
                f"{pluralizar_palavra(embalagem, quantidade)}, "
                f"{'referente' if quantidade == 1 else 'referentes'} à amostra do subitem {item_data['referencia_subitem']}")
        
        if item_data['pessoa_relacionada']:
            desc += f", relacionada a {item_data['pessoa_relacionada']}{'.' if item_data['is_last'] else ';'}"
        else:
            desc += '.' if item_data['is_last'] else ';'
            
        return desc
    except Exception as e:
        return f"[ERRO NA DESCRIÇÃO DO ITEM {numero_item_str}]"

# ========== INTERFACE STREAMLIT ==========
st.set_page_config(layout="wide", page_title="Gerador de Laudos Periciais")

# CSS Customizado
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stTextInput>div>div>input, .stNumberInput>div>div>input, 
    .stSelectbox>div>div>div { border-radius: 8px; padding: 12px; }
    .stButton>button { background-color: #4CAF50; color: white; border-radius: 8px; }
    .header { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# Título
st.markdown("""
<div style="background: linear-gradient(135deg, #2c3e50, #3498db); 
            padding: 20px; border-radius: 10px; color: white; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1)">
    <h1 style="margin:0; text-align: center">Gerador de Laudos Periciais</h1>
    <p style="margin:0; text-align: center; opacity: 0.9">Sistema Oficial</p>
</div>
""", unsafe_allow_html=True)

# Controle de itens
if 'num_itens' not in st.session_state:
    st.session_state.num_itens = 1

num_itens = st.number_input("Quantos itens deseja descrever?", min_value=1, value=st.session_state.num_itens, key='num_itens_selector')

# Formulário principal
with st.form(key='laudo_form'):
    st.header("Custódia")
    lacre_num = st.text_input("Número do Lacre da Contraprova", placeholder="Ex: 0000659555")

    st.header("Informações dos Itens")
    itens_data = []
    for i in range(num_itens):
        st.subheader(f"Item 2.{i + 1}")
        cols = st.columns([1, 3, 3])
        
        with cols[0]:
            qtd = st.number_input(f"Qtd. Porções", key=f'qtd_{i}', min_value=1, value=1)
        
        with cols[1]:
            tipo_mat = st.selectbox(f"Tipo Material", options=list(TIPOS_MATERIAL_BASE.keys()), 
                                  format_func=lambda x: f"{x}: {TIPOS_MATERIAL_BASE[x]}", key=f'tipo_mat_{i}')
        
        with cols[2]:
            tipo_emb = st.selectbox(f"Tipo Embalagem", options=list(TIPOS_EMBALAGEM_BASE.keys()), 
                                  format_func=lambda x: f"{x}: {TIPOS_EMBALAGEM_BASE[x]}", key=f'tipo_emb_{i}')

        cor_emb = None
        if tipo_emb in ['pl', 'pa']:
            cols_cor = st.columns([1, 2])
            with cols_cor[0]:
                cor_key = st.selectbox(f"Cor Emb.", options=list(CORES_FEMININO_EMBALAGEM.keys()), 
                                     format_func=lambda x: CORES_FEMININO_EMBALAGEM.get(x, x), 
                                     key=f'cor_emb_{i}')
            if cor_key == "outra":
                with cols_cor[1]:
                    cor_emb = st.text_input("Digite a cor", key=f'cor_digitada_{i}').lower()
            else:
                cor_emb = cor_key

        ref_sub = st.text_input(f"Ref. Subitem Laudo Constatação", key=f'ref_{i}', placeholder="Ex: 2.1.1")
        pessoa_rel = st.text_input(f"Pessoa Relacionada (Opcional)", key=f'pessoa_{i}')

        itens_data.append({
            'quantidade': qtd,
            'tipo_material': tipo_mat,
            'tipo_embalagem_base': tipo_emb,
            'cor_embalagem': cor_emb,
            'referencia_subitem': ref_sub,
            'pessoa_relacionada': pessoa_rel,
            'is_last': (i == num_itens - 1)
        })

    st.header("Imagens (Opcional)")
    uploaded_files = st.file_uploader("Selecione as imagens (PNG, JPG, JPEG)",
                                    accept_multiple_files=True,
                                    type=['png', 'jpg', 'jpeg'])

    submitted = st.form_submit_button("Gerar Laudo .docx")

# Processamento após submissão
if submitted:
    if not lacre_num or any(not item['referencia_subitem'] for item in itens_data):
        st.error("Preencha todos os campos obrigatórios!")
    else:
        with st.spinner("Gerando documento..."):
            try:
                doc = Document()
                
                # Seção 2: Material Recebido
                add_formatted_paragraph(doc, "2 MATERIAL RECEBIDO PARA EXAME", is_bold=True)
                for i, item in enumerate(itens_data):
                    desc = gerar_descricao_item_web(f"2.{i+1}", item)
                    add_formatted_paragraph(doc, desc)
                
                # [Continue com as outras seções do laudo...]
                
                # Geração do arquivo
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                
                st.success("Laudo gerado com sucesso!")
                st.download_button(
                    label="Baixar Laudo (.docx)",
                    data=bio,
                    file_name="laudo_pericial.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Erro: {str(e)}")
                st.text(traceback.format_exc())
