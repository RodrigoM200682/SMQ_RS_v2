"""
SMQ_RS v4.0
Persistência simples: salva a planilha como arquivo no repositório GitHub.
A cada reinício, lê o arquivo salvo automaticamente.

Secrets (Streamlit Cloud → Settings → Secrets):
  [github]
  token  = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxx"
  repo   = "seu-usuario/smq_rs"
  branch = "main"
"""

import streamlit as st
import pandas as pd
import json, re, base64, urllib.request, urllib.error
from datetime import datetime
from pathlib import Path
from io import BytesIO

BASE_DIR  = Path(__file__).parent
HTML_FILE = BASE_DIR / "dashboard_rnc.html"
DATA_DIR  = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

# Cache local (sobrevive dentro da mesma sessão do servidor)
CACHE_JSON = DATA_DIR / "cache.json"
CACHE_META = DATA_DIR / "cache_meta.json"

# Caminho do arquivo no repositório GitHub
GH_PATH = "data/planilha.json"   # salva o JSON convertido, não o xlsx

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
# GITHUB — ler e salvar arquivo no repositório
# ══════════════════════════════════════════════════════════════════════════════

def gh_cfg():
    try:
        t = st.secrets["github"]["token"].strip()
        r = st.secrets["github"]["repo"].strip()
        b = st.secrets["github"].get("branch", "main").strip()
        return t, r, b
    except Exception:
        return None, None, None


def gh_ok() -> bool:
    t, r, _ = gh_cfg()
    return bool(t and r)


def _gh_req(method, path, token, repo, branch, data=None):
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    if method == "GET":
        url += f"?ref={branch}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "SMQ_RS/4.0",
        "Content-Type": "application/json",
    }
    body = json.dumps(data).encode() if data else None
    req  = urllib.request.Request(url, data=body, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status, json.loads(r.read())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read())
        except Exception:
            return e.code, {}
    except Exception as e:
        return 0, {"error": str(e)}


def gh_ler() -> tuple[str | None, str | None]:
    """Lê arquivo do GitHub. Retorna (conteudo_str, sha) ou (None, None)."""
    t, r, b = gh_cfg()
    if not t:
        return None, None
    status, body = _gh_req("GET", GH_PATH, t, r, b)
    if status == 200:
        try:
            conteudo = base64.b64decode(
                body["content"].replace("\n", "")
            ).decode("utf-8")
            return conteudo, body.get("sha")
        except Exception:
            return None, None
    return None, None


def gh_salvar(conteudo: str, mensagem: str) -> tuple[bool, str]:
    """Salva arquivo no GitHub. Retorna (sucesso, detalhe)."""
    t, r, b = gh_cfg()
    if not t:
        return False, "GitHub não configurado"

    _, sha = gh_ler()   # busca SHA para update

    payload = {
        "message": mensagem,
        "content": base64.b64encode(conteudo.encode("utf-8")).decode("ascii"),
        "branch":  b,
    }
    if sha:
        payload["sha"] = sha

    status, body = _gh_req("PUT", GH_PATH, t, r, b, payload)
    if status in (200, 201):
        commit = body.get("commit", {}).get("sha", "")[:7]
        return True, f"commit {commit}"
    else:
        msg = body.get("message", str(body))
        return False, f"HTTP {status}: {msg}"


# ══════════════════════════════════════════════════════════════════════════════
# CACHE LOCAL
# ══════════════════════════════════════════════════════════════════════════════

def salvar_cache(json_str: str, ts: str, n: int):
    CACHE_JSON.write_text(json_str, encoding="utf-8")
    CACHE_META.write_text(
        json.dumps({"timestamp": ts, "n_records": n}),
        encoding="utf-8",
    )

def ler_cache() -> tuple[str | None, str, int]:
    if CACHE_JSON.exists() and CACHE_META.exists():
        try:
            j    = CACHE_JSON.read_text(encoding="utf-8")
            meta = json.loads(CACHE_META.read_text(encoding="utf-8"))
            return j, meta.get("timestamp", "—"), int(meta.get("n_records", 0))
        except Exception:
            pass
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
            return None, 0, (
                f"Coluna 'Código' não encontrada.\n"
                f"Colunas detectadas: {', '.join(df.columns.tolist())}"
            )

        registros = []
        for _, row in df.iterrows():
            cod = str(row.get(c_cod, "")).strip()
            if not cod or cod == "nan":
                continue

            dt_str, ano, mes = "", None, None
            dv = row.get(c_dt) if c_dt else None
            if pd.notna(dv):
                try:
                    dt     = pd.Timestamp(dv)
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
        return None, 0, f"Erro ao processar: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# CARREGAR DADOS: GitHub → cache → original
# ══════════════════════════════════════════════════════════════════════════════

def carregar_dados() -> tuple[str | None, str, int, str]:
    """Retorna (json_str, ts, n, origem)."""

    # 1. GitHub (fonte de verdade)
    if gh_ok():
        conteudo, _ = gh_ler()
        if conteudo:
            try:
                payload = json.loads(conteudo)
                j  = json.dumps(payload["dados"], ensure_ascii=False)
                ts = payload.get("timestamp", "—")
                n  = int(payload.get("n_records", 0))
                salvar_cache(j, ts, n)   # atualiza cache local
                return j, ts, n, "github"
            except Exception:
                pass

    # 2. Cache local (última sessão bem-sucedida)
    j, ts, n = ler_cache()
    if j:
        return j, ts, n, "cache"

    # 3. Dados originais embutidos no HTML
    return None, "", 0, "original"


def salvar_dados(j: str, ts: str, n: int) -> list[str]:
    """Salva cache + GitHub. Retorna log."""
    log = []

    # Cache local
    try:
        salvar_cache(j, ts, n)
        log.append("✅ Cache local: salvo")
    except Exception as e:
        log.append(f"⚠️ Cache local: {e}")

    # GitHub
    if gh_ok():
        payload  = json.dumps(
            {"timestamp": ts, "n_records": n, "dados": json.loads(j)},
            ensure_ascii=False
        )
        ok, det = gh_salvar(payload, f"SMQ_RS: {n} registros — {ts}")
        if ok:
            log.append(f"✅ GitHub: salvo ({det})")
        else:
            log.append(f"❌ GitHub: {det}")
    else:
        log.append("ℹ️ GitHub não configurado — dados apenas no cache local")

    return log


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
# INICIALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════════

if "dados" not in st.session_state:
    with st.spinner("Carregando dados..."):
        j, ts, n, origem = carregar_dados()
    st.session_state.update({
        "dados": j, "ts": ts, "n": n,
        "origem": origem, "log": [],
    })


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔵 SMQ_RS")
    st.caption("Sistema de Monitoramento de Qualidade")
    st.divider()

    # Status
    origem = st.session_state.get("origem", "original")
    if origem == "github":
        st.success("✅ Dados carregados do GitHub")
    elif origem == "cache":
        st.warning("💾 Cache local (GitHub offline ou não configurado)")
    else:
        st.info("📋 Dados originais do sistema")

    if st.session_state.get("ts"):
        st.caption(f"🕐 {st.session_state['ts']}")
        st.caption(f"📋 {st.session_state['n']:,} registros")

    st.divider()

    # Status GitHub
    if gh_ok():
        _, repo, branch = gh_cfg()
        st.success(f"☁️ **GitHub:** `{repo}` · `{branch}`")
    else:
        st.warning("⚠️ GitHub não configurado")
        with st.expander("Como configurar"):
            st.code("""
# Streamlit Cloud → Settings → Secrets

[github]
token  = "ghp_xxxxxxxxxxxxxxxxxxxx"
repo   = "seu-usuario/smq_rs"
branch = "main"
""", language="toml")
            st.caption(
                "Token: github.com/settings/tokens "
                "→ Generate new token (classic) → escopo **repo**"
            )

    st.divider()

    # Upload de planilha
    st.markdown("### 📂 Atualizar Planilha")
    arquivo = st.file_uploader("Selecione o arquivo .xlsx", type=["xlsx","xls"])

    if arquivo:
        if st.button("⬆️ Salvar Planilha", type="primary", use_container_width=True):
            # Ler
            with st.spinner("Lendo planilha..."):
                j, n, erro = xlsx_para_json(arquivo.read())

            if not j:
                st.error(f"❌ {erro}")
            else:
                ts = datetime.now().strftime("%d/%m/%Y às %H:%M")

                # Salvar
                with st.spinner("Salvando..."):
                    log = salvar_dados(j, ts, n)

                # Exibir resultado
                for linha in log:
                    if "✅" in linha:   st.success(linha)
                    elif "❌" in linha: st.error(linha)
                    else:               st.caption(linha)

                # Atualizar sessão
                nova_origem = "github" if any("✅ GitHub" in l for l in log) else "cache"
                st.session_state.update({
                    "dados": j, "ts": ts, "n": n,
                    "origem": nova_origem, "log": log,
                })
                st.rerun()

    st.divider()

    if CACHE_JSON.exists():
        if st.button("🗑️ Limpar dados salvos", type="secondary"):
            CACHE_JSON.unlink(missing_ok=True)
            CACHE_META.unlink(missing_ok=True)
            st.session_state.update({
                "dados": None, "ts": "", "n": 0,
                "origem": "original", "log": [],
            })
            st.rerun()

    st.caption("SMQ_RS v4.0 · GitHub + Chart.js")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

render_html(montar_html(
    st.session_state.get("dados"),
    st.session_state.get("ts", ""),
    st.session_state.get("n", 0),
))
