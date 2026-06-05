# SMQ_RS — Sistema de Monitoramento de Qualidade RS

Dashboard interativo para monitoramento de RNCs (Registros de Não Conformidade).

---

## 📦 Estrutura do projeto

```
smq_rs/
├── app.py                          # Aplicação Streamlit
├── dashboard_rnc.html              # Dashboard completo (Chart.js)
├── requirements.txt                # Dependências Python
├── .streamlit/
│   └── secrets.toml.example       # Modelo de configuração
├── data/
│   └── .gitkeep                   # Pasta versionada (dados não commitados)
└── README.md
```

---

## 🚀 Instalação local

```bash
git clone https://github.com/seu-usuario/smq_rs.git
cd smq_rs
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Acesse em `http://localhost:8501`

---

## ☁️ Deploy no Streamlit Cloud (com persistência real)

### Passo 1 — Subir para o GitHub

```bash
git init
git add .
git commit -m "SMQ_RS v1.2"
git remote add origin https://github.com/seu-usuario/smq_rs.git
git push -u origin main
```

### Passo 2 — Criar Personal Access Token no GitHub

1. Acesse: https://github.com/settings/tokens
2. **Generate new token (classic)**
3. Nome: `SMQ_RS`
4. Escopo: marque apenas **`repo`** (Contents read & write)
5. Copie o token gerado (`ghp_...`)

### Passo 3 — Deploy no Streamlit Cloud

1. Acesse https://share.streamlit.io → **New app**
2. Selecione o repositório e o arquivo `app.py`
3. Clique em **Advanced settings → Secrets** e cole:

```toml
[github]
token  = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
repo   = "seu-usuario/smq_rs"
branch = "main"
```

4. Clique em **Deploy**

### Como funciona a persistência

```
Upload de planilha
      │
      ▼
Converte XLSX → JSON
      │
      ├─► Salva em data/ (cache local)
      │
      └─► PUT /repos/{repo}/contents/data/dados_salvos.json
                    (GitHub API — persiste entre deploys)

Reinício do app
      │
      ▼
GET /repos/{repo}/contents/data/dados_salvos.json
      │                                          │
   Encontrou ──────────────────────────────────► Carrega dados
      │
   Não encontrou → usa cache local → usa dados originais
```

Os dados ficam salvos **no próprio repositório GitHub** como arquivos JSON. Isso garante que qualquer reinício, redeploy ou novo usuário sempre carregue a versão mais recente da planilha.

---

## 📊 Abas do dashboard

| Aba | Descrição |
|-----|-----------|
| **Visão Geral** | Totais completos · RNCs em atraso · Gráficos de pizza |
| **Por Período** | KPIs + gráficos com todos os filtros |
| **Comparação** | Dois intervalos de meses lado a lado |
| **Por Responsável** | Ranking de analistas + top motivos |
| **Registros** | Tabela completa com múltipla seleção |

---

## 🔧 Tecnologias

- **Python** — Streamlit 1.35+, Pandas, OpenPyXL
- **JavaScript** — Chart.js 4.4.1, XLSX.js
- **Persistência** — GitHub API (REST v3)

---

## ❓ Troubleshooting

**Dados não persistem após redeploy**
→ Verifique se o token está configurado em Secrets e tem escopo `repo`.

**Erro 401/403 na sidebar**
→ Token expirado ou sem permissão. Gere um novo em github.com/settings/tokens.

**Planilha não lida corretamente**
→ Confirme que as colunas seguem o modelo da planilha de referência (Consultas_RNC_APP.xlsx).
