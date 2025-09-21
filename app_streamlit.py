import os, io, json
from pathlib import Path
import streamlit as st
import pandas as pd
import yaml
from autofill_core import calcular_tabelas, build_docx, gerar_texto_ia, ia_disponivel

st.set_page_config(page_title="IEFP Auto-Fill", page_icon="📝", layout="wide")
st.title("📝 IEFP Auto‑Fill — Gerador de Formulário com IA")

with st.expander("ℹ️ Como funciona", expanded=True):
    st.markdown("""
    1. Preenche os campos mínimos abaixo (ou carrega um ficheiro YAML).
    2. Opcional: ativa **IA** para redigir os textos.
    3. Clica **Gerar** e descarrega o **Word** e o **Excel**.
    """)

colA, colB = st.columns(2)
with colA:
    st.subheader("Identificação")
    designacao = st.text_input("Designação Social", "")
    nif = st.text_input("NIF", "")
    promotor = st.text_input("Promotor", "")
    forma = st.text_input("Forma Jurídica", "ENI")
    morada = st.text_input("Morada", "")
    email = st.text_input("Email", "")
    telefone = st.text_input("Telefone", "")
    cae = st.text_input("CAE", "")

with colB:
    st.subheader("Configuração")
    anos = st.multiselect("Anos a considerar", [2025,2026,2027,2028], default=[2025,2026,2027])
    usar_ia = st.checkbox("Gerar textos com IA (requer OPENAI_API_KEY configurada no servidor)", value=False)
    limites = {
        "objetivos_projeto": st.number_input("Limite: Objetivos (caracteres)", 200, 4000, 2000, step=50),
        "mercado": st.number_input("Limite: Mercado (caracteres)", 200, 4000, 1200, step=50),
        "instalacoes": st.number_input("Limite: Instalações (caracteres)", 200, 4000, 1000, step=50),
    }

st.markdown("---")
st.subheader("Textos")

col1, col2, col3 = st.columns(3)
with col1:
    objetivos = st.text_area("Objetivos do Projeto", height=160, placeholder="Se vazio e IA ativa, será gerado automaticamente.")
with col2:
    mercado = st.text_area("Mercado", height=160, placeholder="Segmentos, concorrência, proposta de valor...")
with col3:
    instalacoes = st.text_area("Instalações", height=160, placeholder="Localização, meios técnicos, equipa, parcerias...")

st.markdown("---")
st.subheader("Vendas / Serviços (mínimo 1 linha)")
vendas_df = st.data_editor(pd.DataFrame([
    {"designacao":"Serviço A", "preco":50.0, "qtd_mensal":100, "meses_y1":10},
], columns=["designacao","preco","qtd_mensal","meses_y1"]), num_rows="dynamic")

st.subheader("Pessoal")
pessoal_df = st.data_editor(pd.DataFrame([
    {"funcao":"Técnico(a)", "n":1, "venc_mensal":1100, "meses":12},
], columns=["funcao","n","venc_mensal","meses"]), num_rows="dynamic")

st.subheader("Investimento")
inv_df = st.data_editor(pd.DataFrame([
    {"tipo":"equipamento", "descricao":"Equipamentos", "valor":12000},
], columns=["tipo","descricao","valor"]), num_rows="dynamic")

st.subheader("Financiamento")
colf1, colf2, colf3 = st.columns(3)
with colf1:
    emp_mont = st.number_input("Empréstimo: montante (€)", 0.0, 1e9, 12000.0, step=500.0)
with colf2:
    emp_taxa = st.number_input("Taxa de juro (ano)", 0.0, 1.0, 0.06, step=0.005, format="%.3f")
with colf3:
    emp_anos = st.number_input("Anos de amortização", 1, 15, 3, step=1)
cap_proprios = st.number_input("Capitais próprios iniciais (€)", 0.0, 1e9, 8000.0, step=500.0)

st.markdown("---")
uploaded_yaml = st.file_uploader("Ou carrega um ficheiro YAML com estes mesmos campos", type=["yaml","yml"])
if uploaded_yaml is not None:
    data_loaded = yaml.safe_load(uploaded_yaml.read().decode("utf-8"))
    st.session_state["form_data_loaded"] = data_loaded
    st.success("YAML carregado. Se clicares Gerar, uso os dados do teu ficheiro.")

if st.button("Gerar documentos"):
    if "form_data_loaded" in st.session_state:
        cfg = st.session_state["form_data_loaded"]
    else:
        cfg = {
            "identificacao": {
                "designacao_social": designacao, "nif": nif, "promotor": promotor, "forma_juridica": forma,
                "morada": morada, "email": email, "telefone": telefone, "cae": cae
            },
            "textos": {"objetivos_projeto": objetivos, "mercado": mercado, "instalacoes": instalacoes},
            "limites_caracteres": limites,
            "anos": anos,
            "vendas": vendas_df.fillna(0).to_dict(orient="records"),
            "pessoal": pessoal_df.fillna(0).to_dict(orient="records"),
            "investimento": inv_df.fillna(0).to_dict(orient="records"),
            "capitais_proprios_iniciais": cap_proprios,
            "emprestimo": {"montante": emp_mont, "taxa_juros": emp_taxa, "amortizacao_anos": int(emp_anos)},
        }

    if usar_ia and ia_disponivel():
        ctx = {"identificacao": cfg.get("identificacao", {}),
               "vendas": cfg.get("vendas", []),
               "pessoal": cfg.get("pessoal", []),
               "investimento": cfg.get("investimento", [])}
        tx = cfg.setdefault("textos", {})
        if not tx.get("objetivos_projeto"):
            tx["objetivos_projeto"] = gerar_texto_ia("Objetivos do Projeto", "Inclui metas, resultados e KPIs.", ctx)
        if not tx.get("mercado"):
            tx["mercado"] = gerar_texto_ia("Mercado", "Segmentos, necessidades, concorrência e proposta de valor.", ctx)
        if not tx.get("instalacoes"):
            tx["instalacoes"] = gerar_texto_ia("Instalações", "Localização, meios técnicos, equipa e parcerias.", ctx)

    tabs = calcular_tabelas(cfg)

    import io
    from openpyxl import Workbook
    xlsx_buf = io.BytesIO()
    with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as w:
        for nome, df in tabs.items():
            df.to_excel(w, index=False, sheet_name=nome[:31])
    xlsx_buf.seek(0)

    docx_buf = io.BytesIO()
    tmp_path = Path("tmp_doc.docx")
    build_docx(cfg, tabs, tmp_path)
    docx_buf.write(tmp_path.read_bytes()); docx_buf.seek(0)
    tmp_path.unlink(missing_ok=True)

    st.success("Documentos gerados!")
    colD, colE = st.columns(2)
    with colD:
        st.download_button("⬇️ Download DOCX", data=docx_buf, file_name="preenchido_iefp.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with colE:
        st.download_button("⬇️ Download Excel", data=xlsx_buf, file_name="dados_calculados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
