# IEFP Auto‑Fill — Streamlit

Gerador de **formulários IEFP** em **Word + Excel**, a partir de um formulário web simples — com opção de **IA** para redigir textos.

## ▶️ Correr localmente
```bash
pip install -r requirements.txt
streamlit run app_streamlit.py
```
Abrir: http://localhost:8501

### IA opcional
Defina a variável de ambiente **OPENAI_API_KEY** para gerar textos automáticos (Objetivos, Mercado, Instalações).

- Windows (PowerShell):
```powershell
setx OPENAI_API_KEY "sk-..."
```
- macOS/Linux (bash):
```bash
export OPENAI_API_KEY="sk-..."
```

## ☁️ Deploy no Streamlit Community Cloud
1. Faça **fork** ou carregue este código para um repositório GitHub.
2. Em https://share.streamlit.io → **New app** → selecione o repo e `app_streamlit.py`.
3. (Opcional) Adicione `OPENAI_API_KEY` em *Settings → Secrets* para textos por IA.

## Estrutura
```
iefp-autofill-streamlit/
├─ app_streamlit.py      # App web
├─ autofill_core.py      # Cálculos + geração DOCX
├─ requirements.txt
├─ .streamlit/config.toml
└─ README.md
```

## RGPD
- O app não persiste dados. Downloads são gerados em memória. 
- Use **Secrets** para chaves/API. Não publique dados sensíveis.
