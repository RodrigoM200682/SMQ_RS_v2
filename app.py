"""
SMQ_RS v5.0 — Persistência via st.session_state + arquivo local
"""

import streamlit as st
import pandas as pd
import json, re, base64, pickle
from datetime import datetime
from pathlib import Path
from io import BytesIO

BASE_DIR  = Path(__file__).parent
HTML_FILE = BASE_DIR / "dashboard_rnc.html"
DATA_DIR  = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

# Arquivo de dados persistente — salvo pelo próprio Python
DADOS_FILE = DATA_DIR / "dados.pkl"

# ── Página ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SMQ_RS",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown("""
<style>
  #MainMenu,footer,header{visibility:hidden;}
  .block-container{padding:0!important;max-width:100%!important;}
  .stApp{background:#0f1117;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PERSISTÊNCIA — pickle (binário, mais robusto que JSON para dados grandes)
# ══════════════════════════════════════════════════════════════════════════════

def salvar_disco(json_str: str, ts: str, n: int) -> None:
    """Salva dados em arquivo binário no disco."""
    with open(DADOS_FILE, "wb") as f:
        pickle.dump({"json": json_str, "ts": ts, "n": n}, f)


def carregar_disco() -> tuple[str | None, str, int]:
    """Lê dados do disco. Retorna (json_str, ts, n)."""
    if DADOS_FILE.exists():
        try:
            with open(DADOS_FILE, "rb") as f:
                d = pickle.load(f)
            return d["json"], d["ts"], d["n"]
        except Exception:
            pass
    return None, "", 0


# ══════════════════════════════════════════════════════════════════════════════
# GITHUB — backup extra de persistência (opcional mas recomendado)
# ══════════════════════════════════════════════════════════════════════════════

def gh_cfg():
    try:
        return (
            st.secrets["github"]["token"].strip(),
            st.secrets["github"]["repo"].strip(),
            st.secrets["github"].get("branch", "main").strip(),
        )
    except Exception:
        return None, None, None


def gh_ok() -> bool:
    t, r, _ = gh_cfg()
    return bool(t and r)


def gh_salvar(json_str: str, ts: str, n: int) -> tuple[bool, str]:
    """Salva JSON no GitHub como backup."""
    import urllib.request, urllib.error
    t, r, b = gh_cfg()
    if not t:
        return False, "não configurado"

    path    = "data/dados.json"
    payload = json.dumps({"ts": ts, "n": n, "dados": json.loads(json_str)},
                         ensure_ascii=False)
    content = base64.b64encode(payload.encode()).decode()

    # Buscar SHA atual
    sha = None
    url = f"https://api.github.com/repos/{r}/contents/{path}?ref={b}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {t}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "SMQ_RS",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            sha = json.loads(resp.read()).get("sha")
    except Exception:
        pass

    # Salvar (criar ou atualizar)
    body = {"message": f"SMQ_RS dados {ts}", "content": content, "branch": b}
    if sha:
        body["sha"] = sha

    url2 = f"https://api.github.com/repos/{r}/contents/{path}"
    req2 = urllib.request.Request(
        url2,
        data=json.dumps(body).encode(),
        method="PUT",
        headers={
            "Authorization": f"Bearer {t}",
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
            "User-Agent": "SMQ_RS",
        },
    )
    try:
        with urllib.request.urlopen(req2, timeout=20) as resp:
            ok = resp.status in (200, 201)
            return ok, "salvo" if ok else f"status {resp.status}"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}"
    except Exception as e:
        return False, str(e)


def gh_ler() -> tuple[str | None, str, int]:
    """Lê dados do GitHub. Retorna (json_str, ts, n)."""
    import urllib.request, urllib.error
    t, r, b = gh_cfg()
    if not t:
        return None, "", 0
    try:
        url = f"https://api.github.com/repos/{r}/contents/data/dados.json?ref={b}"
        req = urllib.request.Request(url, headers={
            "Authorization": f"Bearer {t}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "SMQ_RS",
        })
        with urllib.request.urlopen(req, timeout=15) as resp:
            body    = json.loads(resp.read())
            payload = json.loads(
                base64.b64decode(body["content"].replace("\n", "")).decode()
            )
            return (
                json.dumps(payload["dados"], ensure_ascii=False),
                payload.get("ts", "—"),
                int(payload.get("n", 0)),
            )
    except Exception:
        return None, "", 0


# ══════════════════════════════════════════════════════════════════════════════
# XLSX → JSON
# ══════════════════════════════════════════════════════════════════════════════

def xlsx_para_json(file_bytes: bytes) -> tuple[str | None, int, str]:
    try:
        df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]

        def col(palavras):
            for p in palavras:
                for c in df.columns:
                    if p.lower() in c.lower():
                        return c
            return None

        c_cod = col(["código","codigo"])
        c_tit = col(["título","titulo"])
        c_st  = col(["status"])
        c_sit = col(["situação","situacao"])
        c_dt  = col(["emissão","emissao","data"])
        c_rsp = col(["responsável","responsavel"])
        c_cli = col(["cliente"])
        c_rca = col(["análise de causa","analise de causa"])
        c_mot = col(["motivo"])
        c_qtd = col(["quantidade"])
        c_trn = col(["turno"])

        if not c_cod:
            return None, 0, f"Coluna Código não encontrada. Colunas: {list(df.columns)}"

        registros = []
        for _, row in df.iterrows():
            cod = str(row.get(c_cod, "")).strip()
            if not cod or cod == "nan":
                continue

            dt_str, ano, mes = "", None, None
            dv = row.get(c_dt) if c_dt else None
            if pd.notna(dv):
                try:
                    dt = pd.Timestamp(dv)
                    dt_str = dt.strftime("%Y-%m-%d")
                    ano, mes = int(dt.year), int(dt.month)
                except Exception:
                    pass

            trn = str(row.get(c_trn, "")) if c_trn else ""
            if "1" in trn and "2" in trn and "3" in trn:
                turno = "Múltiplos Turnos"
            elif "1" in trn: turno = "1° Turno"
            elif "2" in trn: turno = "2° Turno"
            elif "3" in trn: turno = "3° Turno"
            else:             turno = "Não Informado"

            def safe(c):
                if not c: return ""
                v = row.get(c)
                return str(v).strip() if pd.notna(v) and str(v) != "nan" else ""

            registros.append({
                "codigo": cod,             "titulo": safe(c_tit),
                "status": safe(c_st),      "situacao": safe(c_sit),
                "data":   dt_str,          "ano": ano, "mes": mes,
                "responsavel":       safe(c_rsp),
                "cliente":           safe(c_cli),
                "responsavel_causa": safe(c_rca),
                "motivo": safe(c_mot),     "qtd": safe(c_qtd),
                "turno":  turno,
            })

        if not registros:
            return None, 0, "Nenhum registro encontrado."
        return json.dumps(registros, ensure_ascii=False), len(registros), ""
    except Exception as e:
        return None, 0, f"Erro: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════════════════

_RE = re.compile(r"const RAW_DATA = \[.*?\];", re.DOTALL)

def montar_html(j, ts, n):
    html = HTML_FILE.read_text(encoding="utf-8")
    if j:
        html = _RE.sub(f"const RAW_DATA = {j};", html)
        html = re.sub(r"\d+ registros carregados", f"{n} registros carregados", html)
        html = re.sub(
            r"(Base original[^<\"]*|Atualizado em [^<\"]*)",
            f"Atualizado em {ts}", html,
        )
    return html


def render_html(html, height=980):
    b64 = base64.b64encode(html.encode("utf-8")).decode("ascii")
    st.markdown(
        f'<iframe src="data:text/html;base64,{b64}" '
        f'width="100%" height="{height}px" frameborder="0" '
        f'style="border:none;display:block;" '
        f'sandbox="allow-scripts allow-same-origin allow-forms '
        f'allow-popups allow-downloads"></iframe>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# INICIALIZAÇÃO — carrega dados na ordem: GitHub → disco → original
# ══════════════════════════════════════════════════════════════════════════════

if "dados" not in st.session_state:
    j, ts, n, origem = None, "", 0, "original"

    # 1. Tentar GitHub (fonte mais atualizada)
    if gh_ok():
        j, ts, n = gh_ler()
        if j:
            origem = "github"
            salvar_disco(j, ts, n)   # atualiza cache local

    # 2. Fallback: arquivo local no disco
    if not j:
        j, ts, n = carregar_disco()
        if j:
            origem = "disco"

    st.session_state.update({
        "dados": j, "ts": ts, "n": n, "origem": origem
    })


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔵 SMQ_RS")
    st.caption("Sistema de Monitoramento de Qualidade")
    st.divider()

    # Status dos dados
    origem = st.session_state.get("origem", "original")
    icons  = {"github": "☁️", "disco": "💾", "original": "📋"}
    labels = {
        "github":   "GitHub — persistência ativa",
        "disco":    "Disco local — GitHub não acessado",
        "original": "Dados originais do sistema",
    }
    if origem == "github":
        st.success(f"☁️ {labels[origem]}")
    elif origem == "disco":
        st.warning(f"💾 {labels[origem]}")
    else:
        st.info(f"📋 {labels[origem]}")

    if st.session_state.get("ts"):
        st.caption(f"🕐 {st.session_state['ts']}")
        st.caption(f"📋 {st.session_state['n']:,} registros")

    st.divider()

    # Status GitHub
    if gh_ok():
        _, repo, branch = gh_cfg()
        st.success(f"☁️ GitHub: `{repo}` · `{branch}`")
    else:
        with st.expander("⚙️ Configurar persistência (GitHub)"):
            st.code("""[github]
token  = "ghp_xxxxxxxxxxxxxxxxxxxx"
repo   = "seu-usuario/smq_rs"
branch = "main"
""", language="toml")
            st.caption(
                "Cole em: Streamlit Cloud → Settings → Secrets  \n"
                "Token em: github.com/settings/tokens → escopo **repo**"
            )

    st.divider()

    # Upload
    st.markdown("### 📂 Atualizar Planilha")
    arquivo = st.file_uploader("Arquivo .xlsx", type=["xlsx","xls"])

    if arquivo:
        if st.button("⬆️ Salvar Planilha", type="primary", use_container_width=True):
            log = []

            # Passo 1 — ler planilha
            with st.spinner("Lendo planilha..."):
                j, n, erro = xlsx_para_json(arquivo.read())

            if not j:
                st.error(f"❌ {erro}")
                st.stop()

            ts = datetime.now().strftime("%d/%m/%Y às %H:%M")
            log.append(f"✅ {n:,} registros lidos")

            # Passo 2 — salvar no disco
            try:
                salvar_disco(j, ts, n)
                log.append("✅ Disco: salvo")
            except Exception as e:
                log.append(f"❌ Disco: {e}")

            # Passo 3 — salvar no GitHub
            if gh_ok():
                with st.spinner("Salvando no GitHub..."):
                    ok, det = gh_salvar(j, ts, n)
                log.append(f"{'✅' if ok else '❌'} GitHub: {det}")
                nova_origem = "github" if ok else "disco"
            else:
                log.append("ℹ️ GitHub não configurado")
                nova_origem = "disco"

            # Atualizar estado
            st.session_state.update({
                "dados": j, "ts": ts, "n": n,
                "origem": nova_origem,
            })

            # Exibir log
            st.divider()
            for linha in log:
                if "✅" in linha:   st.success(linha)
                elif "❌" in linha: st.error(linha)
                else:               st.caption(linha)

            st.rerun()

    st.divider()

    if DADOS_FILE.exists():
        if st.button("🗑️ Limpar dados salvos", type="secondary"):
            DADOS_FILE.unlink(missing_ok=True)
            st.session_state.update({
                "dados": None, "ts": "", "n": 0, "origem": "original"
            })
            st.rerun()

    st.caption("SMQ_RS v5.0")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

render_html(montar_html(
    st.session_state.get("dados"),
    st.session_state.get("ts", ""),
    st.session_state.get("n", 0),
))
