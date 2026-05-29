"""
SMQ_RS v3.0
Lê a planilha .xlsx diretamente do Google Drive.
Ao iniciar, busca o arquivo e atualiza os dados automaticamente.

Secrets (Streamlit Cloud → Settings → Secrets):
  [gcp]
  file_id     = "ID_DO_ARQUIVO_XLSX_NO_DRIVE"
  credentials = '''{ JSON da conta de serviço }'''
"""

import streamlit as st
import pandas as pd
import json, re, base64, io
from datetime import datetime
from pathlib import Path
from io import BytesIO

BASE_DIR  = Path(__file__).parent
HTML_FILE = BASE_DIR / "dashboard_rnc.html"
DATA_DIR  = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
CACHE_JSON = DATA_DIR / "cache.json"
CACHE_META = DATA_DIR / "cache_meta.json"

# ── Página ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SMQ_RS — Monitoramento de Qualidade",
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
# GOOGLE DRIVE — baixar planilha .xlsx pelo file_id
# ══════════════════════════════════════════════════════════════════════════════

def drive_configurado() -> bool:
    try:
        _ = st.secrets["gcp"]["file_id"]
        _ = st.secrets["gcp"]["credentials"]
        return True
    except Exception:
        return False


def drive_baixar_xlsx() -> tuple[bytes | None, str]:
    """
    Baixa o arquivo .xlsx do Google Drive.
    Retorna (bytes_do_arquivo, mensagem_erro).
    """
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        # Autenticar
        raw   = st.secrets["gcp"]["credentials"]
        info  = json.loads(raw) if isinstance(raw, str) else dict(raw)
        creds = Credentials.from_service_account_info(
            info,
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        svc     = build("drive", "v3", credentials=creds, cache_discovery=False)
        file_id = st.secrets["gcp"]["file_id"].strip()

        # Baixar arquivo
        buf  = io.BytesIO()
        req  = svc.files().get_media(fileId=file_id)
        dl   = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()

        return buf.getvalue(), ""

    except Exception as e:
        return None, str(e)


# ══════════════════════════════════════════════════════════════════════════════
# XLSX → JSON (converte a planilha para o formato do dashboard)
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
                f"Colunas encontradas: {', '.join(df.columns.tolist())}"
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
            return None, 0, "Nenhum registro encontrado na planilha."

        return json.dumps(registros, ensure_ascii=False), len(registros), ""

    except Exception as e:
        return None, 0, f"Erro ao processar planilha: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# CACHE LOCAL — evita baixar o Drive toda vez que o usuário muda de aba
# ══════════════════════════════════════════════════════════════════════════════

def salvar_cache(json_str: str, ts: str, n: int) -> None:
    CACHE_JSON.write_text(json_str, encoding="utf-8")
    CACHE_META.write_text(
        json.dumps({"timestamp": ts, "n_records": n}),
        encoding="utf-8",
    )

def carregar_cache() -> tuple[str | None, str, int]:
    if CACHE_JSON.exists() and CACHE_META.exists():
        try:
            j    = CACHE_JSON.read_text(encoding="utf-8")
            meta = json.loads(CACHE_META.read_text(encoding="utf-8"))
            return j, meta.get("timestamp", "—"), int(meta.get("n_records", 0))
        except Exception:
            pass
    return None, "", 0


# ══════════════════════════════════════════════════════════════════════════════
# CARREGAR DADOS — Drive → cache → original
# ══════════════════════════════════════════════════════════════════════════════

def carregar_dados() -> tuple[str | None, str, int, str, str]:
    """
    Retorna (json_str, timestamp, n_registros, origem, erro).
    origem: 'drive' | 'cache' | 'original'
    """
    # 1. Tentar Google Drive
    if drive_configurado():
        xlsx_bytes, erro_dl = drive_baixar_xlsx()
        if xlsx_bytes:
            j, n, erro_conv = xlsx_para_json(xlsx_bytes)
            if j:
                ts = datetime.now().strftime("%d/%m/%Y às %H:%M")
                salvar_cache(j, ts, n)
                return j, ts, n, "drive", ""
            else:
                # Drive OK mas conversão falhou → tentar cache
                pass
        # Drive falhou → tentar cache
        j, ts, n = carregar_cache()
        if j:
            return j, ts, n, "cache", f"Drive indisponível ({erro_dl}), usando cache"
        return None, "", 0, "original", f"Drive: {erro_dl}"

    # 2. Sem Drive configurado → cache local
    j, ts, n = carregar_cache()
    if j:
        return j, ts, n, "cache", ""

    # 3. Dados originais embutidos
    return None, "", 0, "original", ""


# ══════════════════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════════════════

_RE = re.compile(r"const RAW_DATA = \[.*?\];", re.DOTALL)

def montar_html(json_str, ts, n):
    html = HTML_FILE.read_text(encoding="utf-8")
    if json_str:
        html = _RE.sub(f"const RAW_DATA = {json_str};", html)
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
    with st.spinner("Carregando dados do Google Drive..."):
        j, ts, n, origem, aviso = carregar_dados()
    st.session_state.update({
        "dados": j, "ts": ts, "n": n,
        "origem": origem, "aviso": aviso,
    })


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔵 SMQ_RS")
    st.caption("Sistema de Monitoramento de Qualidade")
    st.divider()

    # ── Status atual ──────────────────────────────────────────────────────────
    origem = st.session_state.get("origem", "original")
    aviso  = st.session_state.get("aviso", "")

    if origem == "drive":
        st.success("☁️ **Google Drive** — dados atualizados")
    elif origem == "cache":
        st.warning("💾 **Cache local** — Drive não acessado")
    else:
        st.info("📋 Dados originais do sistema")

    if aviso:
        st.caption(f"⚠️ {aviso}")

    if st.session_state.get("ts"):
        st.caption(f"🕐 {st.session_state['ts']}")
        st.caption(f"📋 {st.session_state['n']:,} registros")

    st.divider()

    # ── Configuração Google Drive ─────────────────────────────────────────────
    if drive_configurado():
        st.success("☁️ **Drive configurado**")

        # Botão para forçar atualização manual
        if st.button("🔄 Atualizar do Drive agora", use_container_width=True):
            with st.spinner("Baixando planilha do Drive..."):
                j, ts, n, origem, aviso = carregar_dados()
            if j and origem == "drive":
                st.session_state.update({
                    "dados": j, "ts": ts, "n": n,
                    "origem": origem, "aviso": aviso,
                })
                st.success(f"✅ {n:,} registros carregados!\n\n🕐 {ts}")
                st.rerun()
            else:
                st.error(f"❌ {aviso or 'Falha ao baixar do Drive'}")
    else:
        with st.expander("⚙️ Configurar Google Drive", expanded=True):
            st.markdown("""
**Como configurar em 3 passos:**

**1. Criar conta de serviço:**
- console.cloud.google.com
- APIs → ativar **Google Drive API**
- Credenciais → Conta de serviço → baixar JSON

**2. Compartilhar a planilha:**
- Abra o `.xlsx` no Google Drive
- Compartilhar → cole o `client_email` do JSON → Editor
- Copie o **ID do arquivo** da URL:
  `drive.google.com/file/d/`**`ID`**`/view`

**3. Streamlit → Settings → Secrets:**
```toml
[gcp]
file_id     = "ID_DO_ARQUIVO"
credentials = '''
{ cole o JSON completo aqui }
'''
```
""")

    st.divider()

    # ── Upload manual (alternativa sem Drive) ─────────────────────────────────
    with st.expander("📂 Upload manual de planilha"):
        st.caption("Use se não tiver o Google Drive configurado")
        arquivo = st.file_uploader(
            "Selecione o .xlsx",
            type=["xlsx", "xls"],
            key="manual_upload",
        )
        if arquivo:
            if st.button("⬆️ Carregar", type="primary", use_container_width=True):
                with st.spinner("Processando..."):
                    j, n, erro = xlsx_para_json(arquivo.read())
                if not j:
                    st.error(f"❌ {erro}")
                else:
                    ts = datetime.now().strftime("%d/%m/%Y às %H:%M")
                    salvar_cache(j, ts, n)
                    st.session_state.update({
                        "dados": j, "ts": ts, "n": n,
                        "origem": "cache", "aviso": "",
                    })
                    st.success(f"✅ {n:,} registros carregados!")
                    st.rerun()

    st.divider()
    st.caption("SMQ_RS v3.0 · Google Drive + Chart.js")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

render_html(montar_html(
    st.session_state.get("dados"),
    st.session_state.get("ts", ""),
    st.session_state.get("n", 0),
))
