from fastapi import FastAPI, Form, UploadFile, File, Request, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional, List
from datetime import datetime, date
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # Adicionado estilos
from openpyxl.utils import get_column_letter                           # Adicionado utilitário
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
import sqlite3
import os
import re
import uuid
import qrcode
import base64
import html
import secrets
import json
import urllib.parse

app = FastAPI()

DATABASE = "database.db"
UPLOAD_DIR = "static/imagens"

os.makedirs(UPLOAD_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

RACAS = ["Branca", "Preta", "Parda", "Amarela", "Indígena"]
ANOS = ["6º ano", "7º ano", "8º ano", "9º ano"]

# Tipos de questão suportados. Cada um vira um fluxo de cadastro/resposta diferente.
TIPOS_QUESTAO = {
    "multipla_escolha": {"label": "Múltipla escolha (A/B/C/D)", "icone": "🔘"},
    "discursiva":       {"label": "Discursiva (resposta livre)", "icone": "📝"},
    "vf":               {"label": "Verdadeiro ou Falso (afirmações)", "icone": "✓✗"},
    "associacao":       {"label": "Associação de colunas", "icone": "↔"},
}

# Limites pra cartão impresso (mantém legibilidade)
VF_MAX_AFIRMACOES = 5      # até 5 afirmações por questão V/F
ASSOC_MAX_PARES = 5         # até 5×5 (5 itens × 5 letras) na associação

# === Autenticação ===
# Variáveis de ambiente esperadas em produção. Em dev, defaults permitem testar.
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
SESSION_SECRET_KEY = os.environ.get("SESSION_SECRET_KEY", "dev-key-CHANGE-IN-PRODUCTION-" + secrets.token_hex(8))
BASE_URL = os.environ.get("BASE_URL", "http://localhost:8000")
ALLOWED_EMAIL_DOMAIN = os.environ.get("ALLOWED_EMAIL_DOMAIN", "smevr.com.br")
# Modo dev: se não tem credenciais OAuth, libera login fake só com email
DEV_MODE = (os.environ.get("DEV_MODE", "1") == "1") and not GOOGLE_CLIENT_ID
SESSION_COOKIE = "corretor_session"
SESSION_MAX_AGE = 60 * 60 * 24 * 30  # 30 dias

_session_serializer = URLSafeTimedSerializer(SESSION_SECRET_KEY, salt="session-v1")


def _pode_editar_questao(prof: Optional[dict], questao_criador_id: Optional[int]) -> bool:
    """Autor da questão OU admin podem editar. Questões legadas (sem dono) só admin edita."""
    if not prof:
        return False
    if prof.get("is_admin"):
        return True
    if questao_criador_id is None:
        return False
    return prof["id"] == questao_criador_id


def _sanitizar_html_enunciado(html: str) -> str:
    """Permite apenas tags básicas de formatação no enunciado. Remove scripts, iframes, handlers JS.
    Tags permitidas: strong/b, em/i, u, br, p, div (só com style text-align), span (só com style text-align), ul, ol, li, blockquote.
    Atributos permitidos: apenas style com text-align."""
    import re as _re
    if not html:
        return ""
    # Remove tags perigosas completas (com conteúdo)
    html = _re.sub(r'<(script|style|iframe|object|embed|form|input|button|textarea|select|link|meta)\b[^>]*>.*?</\1>',
                   '', html, flags=_re.IGNORECASE | _re.DOTALL)
    html = _re.sub(r'<(script|style|iframe|object|embed|form|input|button|textarea|select|link|meta)\b[^>]*/?>',
                   '', html, flags=_re.IGNORECASE)
    # Remove atributos on* (onclick, onerror, etc.) e javascript: em href/src
    html = _re.sub(r'\son[a-z]+\s*=\s*"[^"]*"', '', html, flags=_re.IGNORECASE)
    html = _re.sub(r"\son[a-z]+\s*=\s*'[^']*'", '', html, flags=_re.IGNORECASE)
    html = _re.sub(r'\son[a-z]+\s*=\s*[^\s>]+', '', html, flags=_re.IGNORECASE)
    html = _re.sub(r'(href|src)\s*=\s*["\']?\s*javascript:[^"\'>\s]*["\']?', '', html, flags=_re.IGNORECASE)
    # Whitelist de tags - remove qualquer tag que não esteja na lista
    permitidas = {"strong", "b", "em", "i", "u", "br", "p", "div", "span", "ul", "ol", "li", "blockquote"}
    def _filtrar_tag(m):
        tag_full = m.group(0)
        tag_name = m.group(1).lower()
        if tag_name not in permitidas:
            return ""
        # Para div/span/p, mantém só style com text-align
        if tag_name in ("div", "span", "p"):
            ta_match = _re.search(r'style\s*=\s*["\']([^"\']*text-align\s*:\s*(left|center|right|justify)[^"\']*)["\']', tag_full, _re.IGNORECASE)
            if ta_match:
                align_val = ta_match.group(2).lower()
                return f'<{tag_name} style="text-align:{align_val};">' if not tag_full.startswith("</") else f"</{tag_name}>"
        # Outras tags: sem atributos
        if tag_full.startswith("</"):
            return f"</{tag_name}>"
        return f"<{tag_name}>"
    html = _re.sub(r'</?([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>', _filtrar_tag, html)
    return html.strip()


def _editor_enunciado_html(name: str = "enunciado", valor_inicial: str = "", required: bool = True,
                            label: str = "Enunciado", compact: bool = False, min_height: int = 120,
                            placeholder: str = "", detectar_alternativas: bool = False) -> str:
    """Editor WYSIWYG com toolbar EMBAIXO do conteúdo (estilo Slack/Discord).
    - compact=True mostra só B / I / U / limpar (pra campos curtos como alternativas).
    - placeholder aparece DENTRO da caixa quando vazia, some ao digitar.
    - detectar_alternativas=True: ao colar texto com "A) ... B) ... C) ... D) ...",
      oferece extrair as alternativas pros campos alt_a/alt_b/alt_c/alt_d automaticamente.
    O HTML editado é sincronizado num <textarea hidden> que vai no submit."""
    import html as _html
    valor_escapado_textarea = _html.escape(valor_inicial or "")
    req_attr = " required" if required else ""

    # Toolbar: botões variam conforme compact
    btn_style = "padding:3px 7px; background:transparent; border:1px solid var(--border); border-radius:3px; cursor:pointer; font-family:inherit; font-size:12px; color:inherit;"
    bot_basicos = (
        f'<button type="button" data-cmd="bold" title="Negrito (Ctrl+B)" style="{btn_style} font-weight:700; min-width:26px;">B</button>'
        f'<button type="button" data-cmd="italic" title="Itálico (Ctrl+I)" style="{btn_style} font-style:italic; min-width:26px;">I</button>'
        f'<button type="button" data-cmd="underline" title="Sublinhado (Ctrl+U)" style="{btn_style} text-decoration:underline; min-width:26px;">U</button>'
    )
    sep = '<span style="border-left:1px solid var(--border); margin:0 2px;"></span>'
    bot_extra = (
        f'<button type="button" data-cmd="justifyLeft" title="Alinhar à esquerda" style="{btn_style}">⇤</button>'
        f'<button type="button" data-cmd="justifyCenter" title="Centralizar" style="{btn_style}">⇔</button>'
        f'<button type="button" data-cmd="justifyRight" title="Alinhar à direita" style="{btn_style}">⇥</button>'
        f'{sep}'
        f'<button type="button" data-cmd="insertUnorderedList" title="Lista" style="{btn_style}">• Lista</button>'
        f'<button type="button" data-cmd="formatBlock" data-arg="blockquote" title="Citação" style="{btn_style}">❝ Citação</button>'
        f'{sep}'
    )
    bot_limpar = f'<button type="button" data-cmd="removeFormat" title="Limpar formatação" style="{btn_style} color:var(--text-muted);">⌫ limpar</button>'

    toolbar_buttons = bot_basicos + sep + bot_limpar if compact else bot_basicos + sep + bot_extra + bot_limpar

    placeholder_attr = f' data-placeholder="{_html.escape(placeholder, quote=True)}"' if placeholder else ""

    return f"""
        <style>
            .editor-content[data-placeholder]:empty::before {{
                content: attr(data-placeholder);
                color: var(--text-muted);
                opacity: 0.7;
                pointer-events: none;
                font-style: italic;
            }}
            .ed-wrap:focus-within {{ box-shadow: 0 0 0 2px rgba(59,130,246,0.3); border-color: var(--accent); }}
            .editor-content blockquote {{ margin: 8px 0; padding: 6px 14px; border-left: 3px solid var(--border); color: var(--text-muted); font-style: italic; }}
            .editor-content ul {{ margin: 6px 0 6px 22px; }}
        </style>
        <label style="display:block; margin:8px 0;">{label}
            <div class="ed-wrap" style="border:1px solid var(--border); border-radius:5px; background:var(--bg); overflow:hidden;">
                <div class="editor-content" contenteditable="true" data-target="{name}"{placeholder_attr} style="min-height:{min_height}px; padding:10px 12px; outline:none; font-family:inherit; font-size:14px; line-height:1.5;">{valor_inicial}</div>
                <div class="editor-toolbar" style="display:flex; gap:3px; flex-wrap:wrap; align-items:center; padding:5px 7px; background:var(--bg-subtle); border-top:1px solid var(--border);">
                    {toolbar_buttons}
                </div>
            </div>
            <textarea name="{name}" id="{name}_hidden" style="display:none;"{req_attr}>{valor_escapado_textarea}</textarea>
        </label>
        <script>
        (function() {{
            const editor = document.querySelector('.editor-content[data-target="{name}"]');
            const hidden = document.getElementById('{name}_hidden');
            if (!editor || !hidden) return;
            function sync() {{ hidden.value = editor.innerHTML; }}
            editor.addEventListener('input', sync);
            editor.addEventListener('blur', sync);
            const form = editor.closest('form');
            if (form) form.addEventListener('submit', sync);

            // Placeholder: mostra quando vazio (via CSS :empty já cobre em alguns browsers; aqui garantimos)
            const ph = editor.getAttribute('data-placeholder') || '';
            function refreshPlaceholder() {{
                const isEmpty = editor.innerHTML.trim() === '' || editor.innerHTML.trim() === '<br>';
                if (isEmpty && ph && !editor.hasAttribute('data-ph-shown')) {{
                    editor.setAttribute('data-ph-shown', '1');
                    editor.style.position = 'relative';
                }}
                if (!isEmpty) editor.removeAttribute('data-ph-shown');
            }}
            editor.addEventListener('input', refreshPlaceholder);
            refreshPlaceholder();

            const toolbar = editor.parentNode.querySelector('.editor-toolbar');
            if (toolbar) {{
                toolbar.querySelectorAll('button[data-cmd]').forEach(btn => {{
                    btn.addEventListener('click', e => {{
                        e.preventDefault();
                        const cmd = btn.getAttribute('data-cmd');
                        const arg = btn.getAttribute('data-arg') || null;
                        editor.focus();
                        try {{ document.execCommand(cmd, false, arg); }} catch(err) {{}}
                        sync();
                        refreshPlaceholder();
                    }});
                }});
            }}

            {"" if not detectar_alternativas else """
            // ===== Detector de alternativas ao colar =====
            function detectarAlternativas(texto) {
                texto = texto.replace(/\\r\\n/g, '\\n').replace(/\\u00A0/g, ' ').trim();
                // Padrão: início do texto OU quebra de linha, espaços, (opcional, letra A-D, )/./:, espaço
                const padrao = /(?:^|\\n)[ \\t]*\\(?([A-Da-d])[\\)\\.\\:][ \\t]+/g;
                const matches = [...texto.matchAll(padrao)];
                let idxA = -1, idxB = -1, idxC = -1, idxD = -1;
                for (const m of matches) {
                    const letra = m[1].toUpperCase();
                    const pos = m.index;
                    if (letra === 'A' && idxA === -1) idxA = pos;
                    else if (letra === 'B' && idxB === -1 && idxA !== -1 && pos > idxA) idxB = pos;
                    else if (letra === 'C' && idxC === -1 && idxB !== -1 && pos > idxB) idxC = pos;
                    else if (letra === 'D' && idxD === -1 && idxC !== -1 && pos > idxC) idxD = pos;
                }
                if (idxA === -1 || idxB === -1 || idxC === -1 || idxD === -1) return null;
                const enunciado = texto.slice(0, idxA).trim();
                function extrair(start, end) {
                    return texto.slice(start, end).replace(/^\\n?[ \\t]*\\(?[A-Da-d][\\)\\.\\:][ \\t]+/, '').trim();
                }
                return {
                    enunciado,
                    alternativas: [
                        extrair(idxA, idxB),
                        extrair(idxB, idxC),
                        extrair(idxC, idxD),
                        extrair(idxD, texto.length),
                    ]
                };
            }

            editor.addEventListener('paste', (e) => {
                const cb = e.clipboardData || window.clipboardData;
                if (!cb) return;
                const texto = cb.getData('text/plain') || '';
                if (!texto) return;
                const r = detectarAlternativas(texto);
                if (!r) return;  // cola normal
                e.preventDefault();
                const trunc = s => (s.length > 50 ? s.slice(0, 50) + '...' : s);
                const msg = 'Detectei 4 alternativas no que você colou. Aplicar automaticamente?\\n\\n'
                          + 'Enunciado: ' + trunc(r.enunciado) + '\\n'
                          + 'A) ' + trunc(r.alternativas[0]) + '\\n'
                          + 'B) ' + trunc(r.alternativas[1]) + '\\n'
                          + 'C) ' + trunc(r.alternativas[2]) + '\\n'
                          + 'D) ' + trunc(r.alternativas[3]) + '\\n\\n'
                          + 'Atenção: substitui o conteúdo atual dos 5 campos.';
                if (!confirm(msg)) {
                    document.execCommand('insertText', false, texto);
                    return;
                }
                // Aplica: enunciado limpo + 4 alternativas
                editor.innerHTML = r.enunciado.replace(/\\n/g, '<br>');
                hidden.value = editor.innerHTML;
                refreshPlaceholder();
                ['a','b','c','d'].forEach((letra, i) => {
                    const altEd = document.querySelector('.editor-content[data-target="alt_' + letra + '"]');
                    const altHid = document.getElementById('alt_' + letra + '_hidden');
                    if (altEd && altHid) {
                        altEd.innerHTML = r.alternativas[i].replace(/\\n/g, '<br>');
                        altHid.value = altEd.innerHTML;
                        altEd.removeAttribute('data-ph-shown');
                    }
                });
            });
            """}
        }})();
        </script>
    """
    """Autor da questão OU admin podem editar. Questões legadas (sem dono) só admin edita."""
    if not prof:
        return False
    if prof["is_admin"]:
        return True
    if questao_criador_id is None:
        return False
    return prof["id"] == questao_criador_id


def _require_admin_or_403(request: Request) -> HTMLResponse:
    """Retorna None se admin, ou HTMLResponse 403 se não. Helper p/ rotas internas."""
    prof = get_current_professor(request)
    if not prof or not prof["is_admin"]:
        return HTMLResponse(render_page(
            "Acesso restrito",
            '<div class="page-header"><h1>🔒 Acesso restrito</h1></div>'
            '<div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:16px; border-radius:6px;">'
            '<p>Apenas o administrador da escola pode criar, editar ou excluir <strong>turmas e estudantes</strong>.</p>'
            '<p>Se você precisa de uma turma cadastrada, fale com o administrador.</p>'
            '</div>'
            '<div class="page-actions" style="margin-top:16px;"><a href="/turmas" class="btn">← Ver turmas</a></div>',
            active="turmas"
        ), status_code=403)
    return None


# ContextVar pra propagar prof logado pra render_page sem mudar 50 assinaturas
import contextvars
_current_prof_ctx: contextvars.ContextVar[Optional[dict]] = contextvars.ContextVar("current_prof", default=None)


def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS disciplinas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE);
        CREATE TABLE IF NOT EXISTS questoes (id INTEGER PRIMARY KEY AUTOINCREMENT, disciplina_id INTEGER NOT NULL, enunciado TEXT NOT NULL, FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id));
        CREATE TABLE IF NOT EXISTS alternativas (id INTEGER PRIMARY KEY AUTOINCREMENT, questao_id INTEGER NOT NULL, letra TEXT NOT NULL, texto TEXT NOT NULL, correta INTEGER NOT NULL DEFAULT 0, FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS textos_apoio (id INTEGER PRIMARY KEY AUTOINCREMENT, questao_id INTEGER NOT NULL, conteudo TEXT NOT NULL, fonte TEXT, ordem INTEGER NOT NULL DEFAULT 0, FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS imagens (id INTEGER PRIMARY KEY AUTOINCREMENT, questao_id INTEGER NOT NULL, caminho TEXT NOT NULL, legenda TEXT, fonte TEXT, ordem INTEGER NOT NULL DEFAULT 0, FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS provas (id INTEGER PRIMARY KEY AUTOINCREMENT, titulo TEXT NOT NULL, descricao TEXT, criada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS prova_questoes (id INTEGER PRIMARY KEY AUTOINCREMENT, prova_id INTEGER NOT NULL, questao_id INTEGER NOT NULL, ordem INTEGER NOT NULL DEFAULT 0, FOREIGN KEY (prova_id) REFERENCES provas(id) ON DELETE CASCADE, FOREIGN KEY (questao_id) REFERENCES questoes(id));
        CREATE TABLE IF NOT EXISTS habilidades_bncc (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT NOT NULL UNIQUE, descricao TEXT);
        CREATE TABLE IF NOT EXISTS questao_habilidades (id INTEGER PRIMARY KEY AUTOINCREMENT, questao_id INTEGER NOT NULL, habilidade_id INTEGER NOT NULL, UNIQUE(questao_id, habilidade_id), FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE, FOREIGN KEY (habilidade_id) REFERENCES habilidades_bncc(id));
        CREATE TABLE IF NOT EXISTS turmas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, ano_letivo INTEGER NOT NULL, criada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS alunos (id INTEGER PRIMARY KEY AUTOINCREMENT, turma_id INTEGER NOT NULL, nome TEXT NOT NULL, numero INTEGER, codigo_unico TEXT NOT NULL UNIQUE, raca TEXT, email TEXT, data_nascimento TEXT, FOREIGN KEY (turma_id) REFERENCES turmas(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS aplicacoes (id INTEGER PRIMARY KEY AUTOINCREMENT, prova_id INTEGER NOT NULL, turma_id INTEGER NOT NULL, modo TEXT NOT NULL DEFAULT 'online', titulo TEXT, aberta INTEGER NOT NULL DEFAULT 1, criada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (prova_id) REFERENCES provas(id), FOREIGN KEY (turma_id) REFERENCES turmas(id));
        CREATE TABLE IF NOT EXISTS respostas (id INTEGER PRIMARY KEY AUTOINCREMENT, aplicacao_id INTEGER NOT NULL, aluno_id INTEGER NOT NULL, questao_id INTEGER NOT NULL, alternativa_letra TEXT, respondida_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(aplicacao_id, aluno_id, questao_id), FOREIGN KEY (aplicacao_id) REFERENCES aplicacoes(id) ON DELETE CASCADE, FOREIGN KEY (aluno_id) REFERENCES alunos(id), FOREIGN KEY (questao_id) REFERENCES questoes(id));
        CREATE TABLE IF NOT EXISTS entregas (id INTEGER PRIMARY KEY AUTOINCREMENT, aplicacao_id INTEGER NOT NULL, aluno_id INTEGER NOT NULL, finalizada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(aplicacao_id, aluno_id), FOREIGN KEY (aplicacao_id) REFERENCES aplicacoes(id) ON DELETE CASCADE, FOREIGN KEY (aluno_id) REFERENCES alunos(id));
        CREATE TABLE IF NOT EXISTS professores (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT NOT NULL UNIQUE, nome TEXT NOT NULL, foto_url TEXT, is_admin INTEGER NOT NULL DEFAULT 0, criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP, ultimo_acesso TIMESTAMP);
    """)
    cols = {row[1] for row in conn.execute("PRAGMA table_info(alunos)").fetchall()}
    if "raca" not in cols:
        conn.execute("ALTER TABLE alunos ADD COLUMN raca TEXT")
    if "email" not in cols:
        conn.execute("ALTER TABLE alunos ADD COLUMN email TEXT")
    if "data_nascimento" not in cols:
        conn.execute("ALTER TABLE alunos ADD COLUMN data_nascimento TEXT")
    cols_q = {row[1] for row in conn.execute("PRAGMA table_info(questoes)").fetchall()}
    if "ano" not in cols_q:
        conn.execute("ALTER TABLE questoes ADD COLUMN ano TEXT")
    if "criada_por_professor_id" not in cols_q:
        conn.execute("ALTER TABLE questoes ADD COLUMN criada_por_professor_id INTEGER")
    if "tipo" not in cols_q:
        conn.execute("ALTER TABLE questoes ADD COLUMN tipo TEXT DEFAULT 'multipla_escolha'")
        # Questões antigas viram múltipla escolha (que era o único tipo até agora)
        conn.execute("UPDATE questoes SET tipo = 'multipla_escolha' WHERE tipo IS NULL")

    # Migração: respostas ganham coluna pra V/F e Associação (JSON)
    cols_resp = {row[1] for row in conn.execute("PRAGMA table_info(respostas)").fetchall()}
    if "dados_extra" not in cols_resp:
        conn.execute("ALTER TABLE respostas ADD COLUMN dados_extra TEXT")

    # Tabelas pra V ou F
    conn.execute("""CREATE TABLE IF NOT EXISTS vf_afirmacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        questao_id INTEGER NOT NULL,
        ordem INTEGER NOT NULL,
        texto TEXT NOT NULL,
        gabarito TEXT NOT NULL CHECK(gabarito IN ('V','F')),
        FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE
    )""")

    # Tabelas pra Associação de colunas
    conn.execute("""CREATE TABLE IF NOT EXISTS assoc_itens_a (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        questao_id INTEGER NOT NULL,
        ordem INTEGER NOT NULL,
        texto TEXT NOT NULL,
        gabarito_letra TEXT NOT NULL,
        FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS assoc_itens_b (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        questao_id INTEGER NOT NULL,
        letra TEXT NOT NULL,
        texto TEXT NOT NULL,
        FOREIGN KEY (questao_id) REFERENCES questoes(id) ON DELETE CASCADE
    )""")
    # Migrations multi-prof: criada_por_professor_id em provas e aplicacoes
    cols_p = {row[1] for row in conn.execute("PRAGMA table_info(provas)").fetchall()}
    if "criada_por_professor_id" not in cols_p:
        conn.execute("ALTER TABLE provas ADD COLUMN criada_por_professor_id INTEGER")
    if "status_revisao" not in cols_p:
        conn.execute("ALTER TABLE provas ADD COLUMN status_revisao TEXT NOT NULL DEFAULT 'rascunho'")
    if "obs_gestao" not in cols_p:
        conn.execute("ALTER TABLE provas ADD COLUMN obs_gestao TEXT")
    if "revisado_por_id" not in cols_p:
        conn.execute("ALTER TABLE provas ADD COLUMN revisado_por_id INTEGER")
    if "revisado_em" not in cols_p:
        conn.execute("ALTER TABLE provas ADD COLUMN revisado_em TIMESTAMP")
    cols_a = {row[1] for row in conn.execute("PRAGMA table_info(aplicacoes)").fetchall()}
    if "criada_por_professor_id" not in cols_a:
        conn.execute("ALTER TABLE aplicacoes ADD COLUMN criada_por_professor_id INTEGER")
    # Perfil gestor em professores
    cols_prof = {row[1] for row in conn.execute("PRAGMA table_info(professores)").fetchall()}
    if "is_gestor" not in cols_prof:
        conn.execute("ALTER TABLE professores ADD COLUMN is_gestor INTEGER NOT NULL DEFAULT 0")
    conn.commit()
    conn.close()


init_db()


# ==========================================
#  AUTENTICAÇÃO MULTI-PROFESSOR (D1)
# ==========================================

def _criar_sessao(professor_id: int, email: str) -> str:
    return _session_serializer.dumps({"professor_id": professor_id, "email": email})


def _ler_sessao(token: str) -> Optional[dict]:
    if not token:
        return None
    try:
        return _session_serializer.loads(token, max_age=SESSION_MAX_AGE)
    except (BadSignature, SignatureExpired):
        return None


def get_current_professor(request: Request) -> Optional[dict]:
    """Retorna dict com id/email/nome/is_admin do professor logado, ou None."""
    token = request.cookies.get(SESSION_COOKIE)
    payload = _ler_sessao(token)
    if not payload:
        return None
    conn = get_db()
    prof = conn.execute("SELECT * FROM professores WHERE id = ?", (payload["professor_id"],)).fetchone()
    conn.close()
    if not prof:
        return None
    return {
        "id": prof["id"], "email": prof["email"], "nome": prof["nome"],
        "foto_url": prof["foto_url"], "is_admin": bool(prof["is_admin"]), "is_gestor": bool(prof["is_gestor"] if "is_gestor" in prof.keys() else 0),
    }


# Rotas públicas (sem login)
PUBLIC_PATHS = {"/login", "/auth/google", "/auth/google/callback", "/auth/dev-login", "/logout"}
PUBLIC_PREFIXES = ("/static/", "/responder/")


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path
    if path in PUBLIC_PATHS or any(path.startswith(p) for p in PUBLIC_PREFIXES):
        return await call_next(request)
    prof = get_current_professor(request)
    if not prof:
        from_url = path + ("?" + request.url.query if request.url.query else "")
        return RedirectResponse(f"/login?next={urllib.parse.quote(from_url)}", status_code=303)
    request.state.professor = prof
    token = _current_prof_ctx.set(prof)
    try:
        return await call_next(request)
    finally:
        _current_prof_ctx.reset(token)


def _upsert_professor(email: str, nome: str, foto_url: Optional[str] = None) -> dict:
    """Cria ou atualiza professor. O PRIMEIRO professor cadastrado vira admin automaticamente."""
    conn = get_db()
    existing = conn.execute("SELECT * FROM professores WHERE email = ?", (email,)).fetchone()
    if existing:
        conn.execute("UPDATE professores SET nome = ?, foto_url = ?, ultimo_acesso = CURRENT_TIMESTAMP WHERE id = ?",
                     (nome, foto_url, existing["id"]))
        prof_id = existing["id"]
        is_admin = bool(existing["is_admin"])
    else:
        total = conn.execute("SELECT COUNT(*) AS c FROM professores").fetchone()["c"]
        is_admin_val = 1 if total == 0 else 0
        c = conn.execute(
            "INSERT INTO professores (email, nome, foto_url, is_admin, ultimo_acesso) VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)",
            (email, nome, foto_url, is_admin_val)
        )
        prof_id = c.lastrowid
        is_admin = bool(is_admin_val)
        # Se é o primeiro professor (admin), herda dados legados sem dono
        if is_admin_val == 1:
            conn.execute("UPDATE provas SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
            conn.execute("UPDATE aplicacoes SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
            conn.execute("UPDATE questoes SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
    conn.commit()
    conn.close()
    return {"id": prof_id, "email": email, "nome": nome, "is_admin": is_admin}


@app.get("/login", response_class=HTMLResponse)
def pagina_login(request: Request, next: str = "/", erro: str = ""):
    prof = get_current_professor(request)
    if prof:
        return RedirectResponse(next or "/", status_code=303)

    import html as _html
    erro_html = f'<div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:12px; border-radius:6px; margin-bottom:16px;">{_html.escape(erro)}</div>' if erro else ""

    if DEV_MODE:
        botao_login = f"""
            <div style="background:var(--orange-bg); color:var(--orange); border:1px solid var(--orange); padding:12px; border-radius:6px; margin-bottom:16px; font-size:13px;">
                ⚙ <strong>Modo de desenvolvimento</strong> ativo (sem credenciais OAuth Google).
                Em produção, este botão será substituído por "Entrar com Google".
            </div>
            <form action="/auth/dev-login" method="post">
                <input type="hidden" name="next" value="{_html.escape(next, quote=True)}">
                <label>Email institucional<input type="email" name="email" required placeholder="seu.nome@{ALLOWED_EMAIL_DOMAIN}" autofocus></label>
                <label>Nome<input type="text" name="nome" required placeholder="Seu nome"></label>
                <button type="submit" class="btn btn-primary" style="width:100%; margin-top:10px;">Entrar (dev)</button>
            </form>
        """
    else:
        botao_login = f"""
            <a href="/auth/google?next={urllib.parse.quote(next)}" class="btn btn-primary" style="display:flex; align-items:center; justify-content:center; gap:10px; width:100%; padding:12px; font-size:15px;">
                <svg width="20" height="20" viewBox="0 0 24 24"><path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg>
                Entrar com Google
            </a>
            <p class="muted-line" style="font-size:12px; text-align:center; margin-top:14px;">Use sua conta institucional <strong>@{ALLOWED_EMAIL_DOMAIN}</strong>. Acesso de outros domínios será recusado.</p>
        """

    content = f"""
    <div style="max-width:420px; margin:60px auto; padding:30px; background:var(--bg); border:1px solid var(--border); border-radius:8px;">
        <div style="text-align:center; margin-bottom:18px;">
            <img src="/static/imagens/logo_walmir.png" alt="E.M. Walmir de Freitas Monteiro" style="max-width:200px; height:auto; display:block; margin:0 auto;">
        </div>
        <h1 style="margin:0 0 6px 0; text-align:center; font-size:22px;">Sistema Pedagógico</h1>
        <p class="muted-line" style="margin:0 0 24px 0; text-align:center;">E.M. Walmir de Freitas Monteiro</p>
        {erro_html}
        {botao_login}
    </div>
    """
    return render_page("Entrar", content, active="", standalone=True)


@app.get("/auth/google")
def auth_google_redirect(next: str = "/"):
    if DEV_MODE:
        return RedirectResponse(f"/login?next={urllib.parse.quote(next)}", status_code=303)
    state = _session_serializer.dumps({"next": next})
    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "response_type": "code",
        "scope": "openid email profile",
        "redirect_uri": f"{BASE_URL}/auth/google/callback",
        "state": state,
        "hd": ALLOWED_EMAIL_DOMAIN,
        "access_type": "online",
    }
    url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params)
    return RedirectResponse(url, status_code=303)


@app.get("/auth/google/callback")
async def auth_google_callback(code: str = "", state: str = "", error: str = ""):
    if error:
        return RedirectResponse(f"/login?erro={urllib.parse.quote('Login cancelado: ' + error)}", status_code=303)
    if not code:
        return RedirectResponse("/login?erro=Código%20de%20autorização%20ausente", status_code=303)

    try:
        state_data = _session_serializer.loads(state, max_age=600)
        next_url = state_data.get("next", "/")
    except Exception:
        next_url = "/"

    import httpx
    async with httpx.AsyncClient(timeout=15.0) as client:
        token_resp = await client.post("https://oauth2.googleapis.com/token", data={
            "client_id": GOOGLE_CLIENT_ID,
            "client_secret": GOOGLE_CLIENT_SECRET,
            "code": code,
            "grant_type": "authorization_code",
            "redirect_uri": f"{BASE_URL}/auth/google/callback",
        })
        if token_resp.status_code != 200:
            return RedirectResponse(f"/login?erro={urllib.parse.quote('Falha no token: ' + token_resp.text[:80])}", status_code=303)
        tokens = token_resp.json()
        userinfo_resp = await client.get(
            "https://www.googleapis.com/oauth2/v3/userinfo",
            headers={"Authorization": f"Bearer {tokens['access_token']}"}
        )
        if userinfo_resp.status_code != 200:
            return RedirectResponse("/login?erro=Falha%20ao%20obter%20userinfo", status_code=303)
        userinfo = userinfo_resp.json()

    email = (userinfo.get("email") or "").lower().strip()
    if not email:
        return RedirectResponse("/login?erro=Email%20ausente", status_code=303)
    if not email.endswith("@" + ALLOWED_EMAIL_DOMAIN):
        return RedirectResponse(
            f"/login?erro={urllib.parse.quote(f'Apenas contas @{ALLOWED_EMAIL_DOMAIN} são aceitas. Você entrou com {email}.')}",
            status_code=303
        )

    nome = userinfo.get("name") or email.split("@")[0]
    foto = userinfo.get("picture")
    prof = _upsert_professor(email, nome, foto)

    token = _criar_sessao(prof["id"], prof["email"])
    response = RedirectResponse(next_url or "/", status_code=303)
    response.set_cookie(
        key=SESSION_COOKIE, value=token,
        max_age=SESSION_MAX_AGE, httponly=True, samesite="lax",
        secure=BASE_URL.startswith("https://"),
    )
    return response


@app.post("/auth/dev-login")
def auth_dev_login(email: str = Form(...), nome: str = Form(...), next: str = Form("/")):
    if not DEV_MODE:
        return RedirectResponse("/login?erro=Modo%20dev%20desabilitado", status_code=303)
    email = email.lower().strip()
    if not email.endswith("@" + ALLOWED_EMAIL_DOMAIN):
        return RedirectResponse(
            f"/login?erro={urllib.parse.quote(f'Apenas contas @{ALLOWED_EMAIL_DOMAIN}')}",
            status_code=303
        )
    prof = _upsert_professor(email, nome.strip())
    token = _criar_sessao(prof["id"], prof["email"])
    response = RedirectResponse(next or "/", status_code=303)
    response.set_cookie(
        key=SESSION_COOKIE, value=token,
        max_age=SESSION_MAX_AGE, httponly=True, samesite="lax", secure=False,
    )
    return response


@app.post("/logout")
@app.get("/logout")
def logout():
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(SESSION_COOKIE)
    return response


MATHJAX = """
<script>
window.MathJax = { tex: { inlineMath: [['$', '$']], displayMath: [['$$', '$$']] }, svg: { fontCache: 'global' } };
</script>
<script async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
"""

INTER_FONT = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;500;600;700;800&display=swap" rel="stylesheet">'

def _css_version():
    """Hash curto do app.css pra cache-busting automático. Muda sempre que o CSS muda."""
    import hashlib
    try:
        with open(os.path.join("static", "css", "app.css"), "rb") as f:
            return hashlib.md5(f.read()).hexdigest()[:8]
    except Exception:
        return "0"

CSS_VERSION = _css_version()
CSS_LINK = f'<link rel="stylesheet" href="/static/css/app.css?v={CSS_VERSION}">'

# Script de tema (claro/escuro) — aplicado em todas as páginas via render_page.
# Lê preferência do localStorage e aplica antes do render pra evitar flash.
THEME_BOOT_SCRIPT = """<script>
(function(){
  try {
    var saved = localStorage.getItem('walmir-theme') || 'light';
    document.documentElement.setAttribute('data-theme', saved);
    var sb = localStorage.getItem('walmir-sidebar') || 'expanded';
    document.documentElement.setAttribute('data-sidebar', sb);
  } catch(e) {
    document.documentElement.setAttribute('data-theme', 'light');
    document.documentElement.setAttribute('data-sidebar', 'expanded');
  }
})();
function _walmirToggleTheme() {
  var html = document.documentElement;
  var cur = html.getAttribute('data-theme') || 'light';
  var next = cur === 'light' ? 'dark' : 'light';
  html.setAttribute('data-theme', next);
  try { localStorage.setItem('walmir-theme', next); } catch(e) {}
  document.querySelectorAll('[data-theme-toggle]').forEach(function(btn){
    btn.innerHTML = next === 'dark' ? '☀️ Tema claro' : '🌙 Tema escuro';
  });
}
function _walmirToggleSidebar() {
  var html = document.documentElement;
  var cur = html.getAttribute('data-sidebar') || 'expanded';
  var next = cur === 'expanded' ? 'collapsed' : 'expanded';
  html.setAttribute('data-sidebar', next);
  try { localStorage.setItem('walmir-sidebar', next); } catch(e) {}
}
</script>"""


def render_page(title: str, content: str, active: str = "", head_extra: str = "", standalone: bool = False, professor: Optional[dict] = None) -> str:
    """standalone=True omite sidebar (usado em /login).
    professor: dict do usuário logado pra exibir no rodapé do sidebar. Se None,
    o middleware pode ter colocado em request.state — mas como render_page não
    recebe request, em rotas internas o caller passa explicitamente, OU usamos
    o helper render_page_for(request, ...)."""
    def nav_class(name):
        return ' class="active"' if active == name else ''

    if standalone:
        return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="color-scheme" content="light dark">
    <title>{title} · Sistema Pedagógico do Walmir</title>
    {INTER_FONT}
    {THEME_BOOT_SCRIPT}
    {CSS_LINK}
    {head_extra}
</head>
<body>
    {content}
</body>
</html>"""

    # Rodapé do sidebar com info do prof + toggle de tema
    if professor is None:
        professor = _current_prof_ctx.get()
    user_block = ""
    if professor:
        admin_badge = ' <span style="background:var(--purple); color:white; font-size:9px; padding:1px 5px; border-radius:3px; vertical-align:middle;">ADMIN</span>' if professor.get("is_admin") else ""
        user_block = f"""
            <div class="sidebar-user-footer" style="margin-top:auto; padding:12px; border-top:1px solid var(--border); font-size:12px;">
                <div class="sidebar-user-info">
                    <div style="font-weight:600;">{professor.get("nome", "")}{admin_badge}</div>
                    <div style="color:var(--text-muted); font-size:11px; margin-top:2px; word-break:break-all;">{professor.get("email", "")}</div>
                </div>
                <button data-theme-toggle data-icon="🌙" class="theme-toggle" onclick="_walmirToggleTheme()">🌙 Tema escuro</button>
                <a href="/logout" style="display:inline-block; margin-top:8px; font-size:11px; color:var(--text-muted);">Sair</a>
            </div>
        """

    # Sidebar dinâmico: itens de admin escondidos para professores comuns
    is_admin_view = bool(professor and professor.get("is_admin"))
    # Helper pra montar item da nav: ícone emoji + label (escondida quando collapsed) + data-name (tooltip)
    def nav_item(href, key, icon, label):
        return (
            f'<a href="{href}" data-name="{label}"{nav_class(key)}>'
            f'<span class="nav-icon">{icon}</span>'
            f'<span class="nav-label">{label}</span>'
            f'</a>'
        )

    link_disciplinas = nav_item("/disciplinas", "disciplinas", "📚", "Disciplinas") if is_admin_view else ''
    link_habilidades = nav_item("/habilidades", "habilidades", "🎯", "Habilidades BNCC") if is_admin_view else ''
    link_turmas = nav_item("/turmas", "turmas", "👥", "Turmas") if is_admin_view else ''

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="color-scheme" content="light dark">
    <title>{title} · Sistema Pedagógico do Walmir</title>
    {INTER_FONT}
    {THEME_BOOT_SCRIPT}
    {CSS_LINK}
    {head_extra}
</head>
<body>
    <div class="app">
        <aside class="sidebar" style="display:flex; flex-direction:column;">
            <button class="sidebar-toggle" onclick="_walmirToggleSidebar()" type="button" title="Recolher/expandir menu" aria-label="Recolher menu">
                <span class="sidebar-toggle-icon">≡</span>
            </button>
            <div class="sidebar-brand" style="text-align:center; padding:8px 6px 4px;">
                <img src="/static/imagens/logo_walmir.png" class="sidebar-logo-full" alt="Walmir" style="max-width:100%; height:auto; max-height:80px; display:block; margin:0 auto;">
                <div class="sidebar-logo-mini" aria-hidden="true">W</div>
                <div class="sidebar-brand-text" style="font-size:11px; color:var(--text-muted); margin-top:6px; font-weight:600; letter-spacing:0.3px;">Sistema Pedagógico</div>
            </div>
            <nav>
                {nav_item("/", "home", "🏠", "Início")}
                <div class="sidebar-section">Banco</div>
                {link_disciplinas}
                {link_habilidades}
                {nav_item("/questoes", "questoes", "✏️", "Cadastrar questão")}
                <div class="sidebar-section">Avaliações</div>
                {nav_item("/provas", "provas", "📝", "Cadastrar atividade")}
                {link_turmas}
                {nav_item("/aplicacoes", "aplicacoes", "📤", "Aplicar atividade")}
                {nav_item("/minhas-aplicacoes", "minhas-aplicacoes", "📋", "Minhas aplicações")}
                {nav_item("/painel-gestao", "painel-gestao", "\U0001f3db\ufe0f", "Painel de gestão") if (professor and (professor.get("is_admin") or professor.get("is_gestor"))) else ""}
                {nav_item("/admin/usuarios", "admin-usuarios", "\U0001f465", "Usuários") if (professor and professor.get("is_admin")) else ""}
            </nav>
            {user_block}
        </aside>
        <main class="main">
            {content}
        </main>
    </div>
</body>
</html>"""


def gerar_codigo_aluno(conn):
    for _ in range(10):
        codigo = uuid.uuid4().hex[:8].upper()
        if not conn.execute("SELECT id FROM alunos WHERE codigo_unico = ?", (codigo,)).fetchone():
            return codigo
    raise RuntimeError("Não foi possível gerar código único após 10 tentativas")

def qr_data_uri(texto):
    """Gera um QR Code e retorna como data URI base64 (pra embutir em <img src>)."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(texto)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"

def get_base_url(request):
    """Constrói a URL base correta, considerando proxy reverso (Codespaces, produção, etc.)."""
    host = request.headers.get("x-forwarded-host") or request.headers.get("host", "localhost")
    proto = request.headers.get("x-forwarded-proto", "http")
    return f"{proto}://{host}"

def format_data_br(iso_str):
    if not iso_str:
        return ""
    try:
        return datetime.fromisoformat(iso_str).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return iso_str


def render_questao_card(conn, q, numero=None, mostrar_acoes=False, compact=False, pode_editar=True, autor_nome=None):
    """
    pode_editar: se False, esconde botões Editar/Excluir mesmo com mostrar_acoes=True.
    autor_nome: se passado, exibe badge 'Por: <nome>' (usado quando admin lista questões alheias).
    """
    textos = conn.execute("SELECT conteudo, fonte FROM textos_apoio WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
    imagens = conn.execute("SELECT caminho, legenda, fonte FROM imagens WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
    alts = conn.execute("SELECT letra, texto, correta FROM alternativas WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()
    habilidades = conn.execute("SELECT h.codigo FROM questao_habilidades qh JOIN habilidades_bncc h ON h.id = qh.habilidade_id WHERE qh.questao_id = ? ORDER BY h.codigo", (q["id"],)).fetchall()
    ano_q = q["ano"] if "ano" in q.keys() and q["ano"] else None

    textos_html = ""
    for t in textos:
        fonte_html = f'<footer>Fonte: {t["fonte"]}</footer>' if t["fonte"] else ""
        textos_html += f'<blockquote>{t["conteudo"]}{fonte_html}</blockquote>'

    imagens_html = ""
    for img in imagens:
        legenda_html = f'<figcaption>{img["legenda"]}</figcaption>' if img["legenda"] else ""
        fonte_html = f'<figcaption><small>Fonte: {img["fonte"]}</small></figcaption>' if img["fonte"] else ""
        imagens_html += f'<figure><img src="/{img["caminho"]}" alt="">{legenda_html}{fonte_html}</figure>'

    tipo_q = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
    tipo_info = TIPOS_QUESTAO.get(tipo_q, TIPOS_QUESTAO["multipla_escolha"])

    alts_html = ""
    if tipo_q == "multipla_escolha":
        for a in alts:
            cls = ' class="correct"' if a["correta"] else ''
            marca = ' ✓' if a["correta"] else ''
            alts_html += f'<li{cls}><strong>{a["letra"]})</strong> {a["texto"]}{marca}</li>'
    elif tipo_q == "discursiva":
        alts_html = '<li style="list-style:none; padding:8px 12px; background:var(--bg-subtle); border-left:3px solid var(--accent); color:var(--text-muted); font-style:italic;">📝 Questão discursiva — resposta livre (correção manual)</li>'
    elif tipo_q == "vf":
        afirmacoes = conn.execute("SELECT ordem, texto, gabarito FROM vf_afirmacoes WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
        items = ""
        for af in afirmacoes:
            cor = "var(--green)" if af["gabarito"] == "V" else "var(--red)"
            items += (
                f'<li style="list-style:none; padding:6px 10px; background:var(--bg-subtle); margin-bottom:4px; border-radius:4px; border-left:3px solid {cor};">'
                f'<strong style="color:{cor};">({af["gabarito"]})</strong> {af["texto"]}'
                f'</li>'
            )
        alts_html = items or '<li style="list-style:none; padding:8px; color:var(--text-muted); font-style:italic;">(sem afirmações cadastradas)</li>'
    elif tipo_q == "associacao":
        itens_a = conn.execute("SELECT ordem, texto, gabarito_letra FROM assoc_itens_a WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
        itens_b = conn.execute("SELECT letra, texto FROM assoc_itens_b WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()
        ca_html = "".join(
            f'<li style="margin-bottom:4px;"><strong>{a["ordem"]+1}.</strong> {a["texto"]} '
            f'<span style="font-size:11px; color:var(--green);">→ resposta: ({a["gabarito_letra"]})</span></li>'
            for a in itens_a
        )
        cb_html = "".join(
            f'<li style="margin-bottom:4px;"><strong>({b["letra"]})</strong> {b["texto"]}</li>'
            for b in itens_b
        )
        alts_html = (
            f'<li style="list-style:none; padding:8px; background:var(--bg-subtle); border-radius:4px;">'
            f'<div style="display:grid; grid-template-columns:1fr 1fr; gap:14px;">'
            f'<div><strong style="font-size:12px; text-transform:uppercase; color:var(--text-muted);">Coluna A</strong><ul style="margin:6px 0 0 18px; padding:0;">{ca_html}</ul></div>'
            f'<div><strong style="font-size:12px; text-transform:uppercase; color:var(--text-muted);">Coluna B</strong><ul style="margin:6px 0 0 18px; padding:0;">{cb_html}</ul></div>'
            f'</div></li>'
        )

    habilidades_html = ""
    if habilidades:
        badges = "".join(f'<span class="badge">{h["codigo"]}</span>' for h in habilidades)
        habilidades_html = f'<div class="habilidades-row">{badges}</div>'

    # Badge de tipo (sempre visível)
    cores_tipo = {
        "multipla_escolha": ("var(--accent-bg)", "var(--accent)"),
        "discursiva":       ("var(--orange-bg)", "var(--orange)"),
        "vf":               ("var(--green-bg)", "var(--green)"),
        "associacao":       ("var(--purple-bg)", "var(--purple)"),
    }
    cor_tipo_bg, cor_tipo_fg = cores_tipo.get(tipo_q, cores_tipo["multipla_escolha"])
    tipo_badge = f' · <span class="badge" style="background:{cor_tipo_bg}; color:{cor_tipo_fg}; font-size:10px;">{tipo_info["icone"]} {tipo_info["label"]}</span>'

    cabecalho = f'Questão {numero} · {q["disciplina_nome"]}' if numero else q["disciplina_nome"]
    ano_badge = f' · <span style="color:var(--text-muted); font-weight:400;">{ano_q}</span>' if ano_q else ""
    autor_badge_inline = f' · <span class="badge" style="background:var(--purple-bg); color:var(--purple); font-size:10px;">Por: {autor_nome}</span>' if autor_nome else ""

    acoes_html = ""
    if mostrar_acoes and pode_editar:
        acoes_html = (
            f'<div class="page-actions" style="margin-top:16px; padding-top:12px; border-top:1px solid var(--border);">'
            f'<a href="/questoes/{q["id"]}/editar" class="btn">Editar</a>'
            f'<form action="/questoes/{q["id"]}/deletar" method="post" style="margin:0;" '
            f'onsubmit="return confirm(\'Excluir esta questão? Se ela for usada em alguma prova, a exclusão será bloqueada.\');">'
            f'<button type="submit" class="btn" style="background:var(--red); color:white; border-color:var(--red);">Excluir</button>'
            f'</form>'
            f'</div>'
        )
    elif mostrar_acoes and not pode_editar:
        acoes_html = (
            f'<div style="margin-top:12px; padding-top:10px; border-top:1px solid var(--border); font-size:11px; color:var(--text-muted);">'
            f'🔒 Questão de outro professor — você pode usá-la em suas provas, mas só o autor ou o administrador podem editá-la.'
            f'</div>'
        )

    if compact:
        # IMPORTANTE: strip de tags HTML antes do slice. Cortar HTML em 160 chars pode
        # deixar uma tag aberta (ex: <p style="..."> sem </p>), quebrando o layout.
        preview_text = re.sub(r'<[^>]+>', '', q["enunciado"])
        preview = html.escape(preview_text[:160]) + ("..." if len(preview_text) > 160 else "")
        habs_inline = ""
        if habilidades:
            habs_inline = " " + "".join(f'<span class="badge" style="font-size:10px;">{h["codigo"]}</span>' for h in habilidades)
        return (
            f'<div class="question" style="margin-bottom:8px; padding:12px 16px;">'
            f'<div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">'
            f'<div style="flex:1; min-width:0;">'
            f'<div class="question-header" style="margin:0;">Q{q["id"]} · {q["disciplina_nome"]}{ano_badge}{tipo_badge}{autor_badge_inline}{habs_inline}</div>'
            f'<div style="margin-top:6px; color:var(--text); font-size:14px; line-height:1.5;">{preview}</div>'
            f'</div>'
            f'<button type="button" onclick="toggleQuestao({q["id"]})" id="q-toggle-{q["id"]}" '
            f'style="background:none; border:1px solid var(--border); border-radius:4px; padding:4px 10px; color:var(--text-muted); cursor:pointer; font-size:11px; white-space:nowrap; font-family:inherit;">'
            f'Ver completa ▾</button>'
            f'</div>'
            f'<div id="q-detalhes-{q["id"]}" style="display:none; margin-top:14px; padding-top:14px; border-top:1px solid var(--border);">'
            f'{textos_html}{imagens_html}'
            f'<div class="enunciado">{q["enunciado"]}</div>'
            f'<ul class="alternativas">{alts_html}</ul>'
            f'{habilidades_html}{acoes_html}'
            f'</div>'
            f'</div>'
        )

    return f'<div class="question"><div class="question-header">{cabecalho}{ano_badge}{tipo_badge}</div>{textos_html}{imagens_html}<div class="enunciado">{q["enunciado"]}</div><ul class="alternativas">{alts_html}</ul>{habilidades_html}{acoes_html}</div>'


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    prof = get_current_professor(request)
    nome_prof = prof["nome"].split()[0] if prof else "professor(a)"
    prof_id = prof["id"] if prof else 0

    conn = get_db()

    # === ACERVO DA ESCOLA (global) ===
    total_questoes = conn.execute("SELECT COUNT(*) AS c FROM questoes").fetchone()["c"]
    total_disciplinas = conn.execute("SELECT COUNT(*) AS c FROM disciplinas").fetchone()["c"]
    total_habilidades = conn.execute("SELECT COUNT(*) AS c FROM habilidades_bncc").fetchone()["c"]
    total_turmas = conn.execute("SELECT COUNT(*) AS c FROM turmas").fetchone()["c"]
    total_alunos = conn.execute("SELECT COUNT(*) AS c FROM alunos").fetchone()["c"]

    # Questões por ano de escolaridade (6º a 9º)
    questoes_por_ano = {}
    for ano in ANOS:
        n = conn.execute("SELECT COUNT(*) AS c FROM questoes WHERE ano = ?", (ano,)).fetchone()["c"]
        questoes_por_ano[ano] = n
    n_sem_ano = conn.execute("SELECT COUNT(*) AS c FROM questoes WHERE ano IS NULL OR ano = ''").fetchone()["c"]

    # === SEU PAINEL (do prof; pra ADMIN, é o painel DA ESCOLA inteira) ===
    is_admin = bool(prof and prof["is_admin"])
    if is_admin:
        # Admin: contadores globais (vê tudo)
        minhas_provas = conn.execute("SELECT COUNT(*) AS c FROM provas").fetchone()["c"]
        minhas_aplicacoes_abertas = conn.execute(
            "SELECT COUNT(*) AS c FROM aplicacoes WHERE aberta = 1"
        ).fetchone()["c"]
        minhas_aplicacoes_encerradas = conn.execute(
            "SELECT COUNT(*) AS c FROM aplicacoes WHERE aberta = 0"
        ).fetchone()["c"]
        # Últimas 3 aplicações da ESCOLA (qualquer prof)
        minhas_ultimas = conn.execute("""
            SELECT a.id, a.modo, a.aberta,
                   COALESCE(a.titulo, p.titulo) AS titulo,
                   t.nome AS turma_nome, t.ano_letivo,
                   prof.nome AS criador_nome,
                   (SELECT COUNT(*) FROM entregas e WHERE e.aplicacao_id = a.id) AS n_entregas,
                   (SELECT COUNT(*) FROM alunos al WHERE al.turma_id = a.turma_id) AS n_alunos
            FROM aplicacoes a
            JOIN provas p ON p.id = a.prova_id
            JOIN turmas t ON t.id = a.turma_id
            LEFT JOIN professores prof ON prof.id = a.criada_por_professor_id
            ORDER BY a.id DESC LIMIT 3
        """).fetchall()
        # Contagem de professores ativos (com login pelo menos uma vez)
        total_profs = conn.execute("SELECT COUNT(*) AS c FROM professores").fetchone()["c"]
    else:
        # Prof comum: só os dele
        minhas_provas = conn.execute(
            "SELECT COUNT(*) AS c FROM provas WHERE criada_por_professor_id = ?", (prof_id,)
        ).fetchone()["c"]
        minhas_aplicacoes_abertas = conn.execute(
            "SELECT COUNT(*) AS c FROM aplicacoes WHERE criada_por_professor_id = ? AND aberta = 1", (prof_id,)
        ).fetchone()["c"]
        minhas_aplicacoes_encerradas = conn.execute(
            "SELECT COUNT(*) AS c FROM aplicacoes WHERE criada_por_professor_id = ? AND aberta = 0", (prof_id,)
        ).fetchone()["c"]
        minhas_ultimas = conn.execute("""
            SELECT a.id, a.modo, a.aberta,
                   COALESCE(a.titulo, p.titulo) AS titulo,
                   t.nome AS turma_nome, t.ano_letivo,
                   NULL AS criador_nome,
                   (SELECT COUNT(*) FROM entregas e WHERE e.aplicacao_id = a.id) AS n_entregas,
                   (SELECT COUNT(*) FROM alunos al WHERE al.turma_id = a.turma_id) AS n_alunos
            FROM aplicacoes a
            JOIN provas p ON p.id = a.prova_id
            JOIN turmas t ON t.id = a.turma_id
            WHERE a.criada_por_professor_id = ?
            ORDER BY a.id DESC LIMIT 3
        """, (prof_id,)).fetchall()
        total_profs = 0  # não exibido pra prof comum
    conn.close()

    # ----- HTML -----
    # Bloco "Acervo da Escola"
    qpa_cards = "".join(
        f'<div style="text-align:center; padding:10px; background:var(--bg-subtle); border-radius:6px;">'
        f'<div style="font-size:11px; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px;">{ano}</div>'
        f'<div style="font-size:22px; font-weight:600; margin-top:4px;">{questoes_por_ano[ano]}</div>'
        f'</div>'
        for ano in ANOS
    )
    if n_sem_ano > 0:
        qpa_cards += (
            f'<div style="text-align:center; padding:10px; background:var(--bg-subtle); border-radius:6px;">'
            f'<div style="font-size:11px; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px;">Sem ano</div>'
            f'<div style="font-size:22px; font-weight:600; margin-top:4px;">{n_sem_ano}</div>'
            f'</div>'
        )

    acervo_html = f"""
        <h2 style="margin-top:24px; font-size:15px; text-transform:uppercase; letter-spacing:1px; color:var(--text-muted);">📚 Acervo da Escola</h2>
        <div style="display:flex; align-items:baseline; gap:12px; margin:8px 0 12px 0;">
            <span style="font-size:36px; font-weight:700;">{total_questoes}</span>
            <span style="font-size:14px; color:var(--text-muted);">questões disponíveis no banco coletivo</span>
        </div>
        <div style="display:grid; grid-template-columns: repeat({len(ANOS) + (1 if n_sem_ano > 0 else 0)}, 1fr); gap:8px; margin-bottom:14px;">
            {qpa_cards}
        </div>
        <p style="font-size:12px; color:var(--text-muted); margin:0 0 18px 0;">
            <strong>{total_disciplinas}</strong> disciplinas · <strong>{total_habilidades}</strong> habilidades BNCC · <strong>{total_turmas}</strong> turmas · <strong>{total_alunos}</strong> alunos cadastrados
        </p>
    """

    # Bloco "Seu Painel" — labels adaptados pra admin/prof
    label_provas = "Provas / tarefas (escola)" if is_admin else "Provas / tarefas criadas"
    label_abertas = "Aplicações abertas (escola)" if is_admin else "Aplicações abertas"
    label_encerradas = "Aplicações encerradas (escola)" if is_admin else "Aplicações encerradas"
    painel_metrics = f"""
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:10px; margin:8px 0 14px 0;">
            <div class="status-card">
                <div class="status-card-label">{label_provas}</div>
                <div class="status-card-value">{minhas_provas}</div>
            </div>
            <div class="status-card status-card-success">
                <div class="status-card-label">{label_abertas}</div>
                <div class="status-card-value">{minhas_aplicacoes_abertas}</div>
            </div>
            <div class="status-card">
                <div class="status-card-label">{label_encerradas}</div>
                <div class="status-card-value">{minhas_aplicacoes_encerradas}</div>
            </div>
        </div>
    """

    # Últimas aplicações
    if minhas_ultimas:
        linhas = ""
        for u in minhas_ultimas:
            status_dot = '<span style="color:var(--green);">●</span>' if u["aberta"] else '<span style="color:var(--text-muted);">○</span>'
            modo_label = "online" if u["modo"] == "online" else "impressa"
            pct = (u["n_entregas"] / u["n_alunos"] * 100) if u["n_alunos"] > 0 else 0
            autor_inline = f' · <span style="color:var(--purple);">por {u["criador_nome"] or "—"}</span>' if is_admin else ""
            linhas += (
                f'<a href="/aplicacoes/{u["id"]}" style="display:flex; justify-content:space-between; align-items:center; padding:10px 12px; border:1px solid var(--border); border-radius:6px; margin-bottom:6px; text-decoration:none; color:inherit;">'
                f'<div style="min-width:0; flex:1;">{status_dot} <strong>{u["titulo"]}</strong> <span style="font-size:12px; color:var(--text-muted);">· {u["turma_nome"]} · {modo_label}{autor_inline}</span></div>'
                f'<div style="font-size:12px; color:var(--text-muted); flex-shrink:0;">{u["n_entregas"]}/{u["n_alunos"]} entregas ({pct:.0f}%)</div>'
                f'</a>'
            )
        label_ultimas = "Últimas aplicações da escola:" if is_admin else "Últimas aplicações criadas por você:"
        ultimas_html = f"""
            <p style="font-size:12px; color:var(--text-muted); margin:14px 0 6px 0;">{label_ultimas}</p>
            {linhas}
        """
    else:
        ultimas_html = ""

    titulo_painel = "🏫 Painel da Escola" if is_admin else "👤 Seu Painel"
    painel_html = f"""
        <h2 style="margin-top:24px; font-size:15px; text-transform:uppercase; letter-spacing:1px; color:var(--text-muted);">{titulo_painel}</h2>
        {painel_metrics}
        {ultimas_html}
    """

    # Atalhos de ação rápida
    acoes_html = f"""
        <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:24px; padding-top:18px; border-top:1px solid var(--border);">
            <a href="/questoes/nova" class="btn btn-primary">+ Nova questão</a>
            <a href="/provas/nova" class="btn">+ Nova prova/tarefa</a>
            <a href="/aplicacoes/nova" class="btn">+ Nova aplicação</a>
            <a href="/questoes" class="btn">Ver banco de questões</a>
        </div>
    """

    content = f"""
        <div class="page-header">
            <h1 style="margin-bottom:4px;">Olá, {nome_prof} 👋</h1>
            <p class="subtitle" style="margin-top:0;">Veja seu panorama atualizado.</p>
        </div>
        {acervo_html}
        {painel_html}
        {acoes_html}
    """
    return render_page("Início", content, active="home")


@app.get("/disciplinas", response_class=HTMLResponse)
def listar_disciplinas(request: Request):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    disciplinas = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()
    conn.close()
    if disciplinas:
        linhas = "".join(f"<li>{d['nome']}</li>" for d in disciplinas)
        lista_html = f'<ul class="clean">{linhas}</ul>'
    else:
        lista_html = '<div class="empty">Nenhuma disciplina cadastrada ainda.</div>'
    content = f'<div class="page-header"><h1>Disciplinas</h1><div class="page-actions"><a href="/disciplinas/nova" class="btn btn-primary">+ Nova disciplina</a></div></div>{lista_html}'
    return render_page("Disciplinas", content, active="disciplinas")


@app.get("/disciplinas/nova", response_class=HTMLResponse)
def form_nova_disciplina():
    content = '<div class="page-header"><h1>Nova disciplina</h1></div><form action="/disciplinas/nova" method="post"><label>Nome<input type="text" name="nome" required autofocus></label><div class="page-actions"><button type="submit" class="btn btn-primary">Cadastrar</button><a href="/disciplinas" class="btn">Cancelar</a></div></form>'
    return render_page("Nova disciplina", content, active="disciplinas")


@app.post("/disciplinas/nova")
def criar_disciplina(nome: str = Form(...)):
    conn = get_db()
    try:
        conn.execute("INSERT INTO disciplinas (nome) VALUES (?)", (nome.strip(),))
        conn.commit()
    except sqlite3.IntegrityError:
        pass
    finally:
        conn.close()
    return RedirectResponse("/disciplinas", status_code=303)


# ═══════════════════════════════════════════════════════════════
# IMPORTADOR DE BANCO DE QUESTÕES (JSON em batch)
# ═══════════════════════════════════════════════════════════════

@app.get("/admin/importar-questoes", response_class=HTMLResponse)
def form_importar_questoes(request: Request):
    """Tela admin: upload de JSON com questões pra importar em batch."""
    _r = _require_admin_or_403(request)
    if _r is not None: return _r

    content = """
        <div class="page-header">
            <h1>📥 Importar banco de questões</h1>
            <p class="subtitle">Carregue um arquivo <code>.json</code> com questões estruturadas. Útil para popular o banco com provas oficiais (OBMEP, SAEB, OBA, etc.).</p>
        </div>

        <div class="card" style="background:var(--accent-bg); border-left:3px solid var(--accent);">
            <h3 style="margin-top:0;">📋 Formato esperado</h3>
            <p style="font-size:13px;">O arquivo deve conter um objeto JSON com a chave <code>questoes</code> contendo uma lista. Cada questão precisa de:</p>
            <ul style="font-size:13px; margin:8px 0 0 20px;">
                <li><code>disciplina</code> — nome (será criada se não existir)</li>
                <li><code>ano</code> — ex: <code>"6º ano"</code>, <code>"7º ano"</code>, etc.</li>
                <li><code>tipo</code> — por enquanto só <code>"multipla_escolha"</code></li>
                <li><code>enunciado</code> — texto da questão (HTML permitido)</li>
                <li><code>fonte</code> (opcional) — ex: <code>"OBMEP 2019, Nível 1, Q5"</code></li>
                <li><code>habilidade_bncc</code> (opcional) — código (ex: <code>"EF06MA15"</code>)</li>
                <li><code>alternativas</code> — lista com <code>{letra, texto, correta}</code></li>
            </ul>
            <p style="font-size:12px; color:var(--text-muted); margin-top:10px;">
                Veja o arquivo de exemplo: <code>banco_inicial_obmep.json</code> distribuído junto com o sistema.
            </p>
        </div>

        <form method="post" action="/admin/importar-questoes/preview" enctype="multipart/form-data" style="margin-top:18px;">
            <label>
                Arquivo JSON
                <input type="file" name="arquivo" accept=".json,application/json" required>
            </label>
            <div style="margin-top:14px;">
                <button type="submit" class="btn btn-primary">Visualizar preview →</button>
                <a href="/" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return HTMLResponse(render_page("Importar questões", content, active=""))


@app.post("/admin/importar-questoes/preview", response_class=HTMLResponse)
async def preview_importar_questoes(request: Request):
    """Recebe JSON, valida estrutura e mostra preview antes de confirmar."""
    _r = _require_admin_or_403(request)
    if _r is not None: return _r

    form = await request.form()
    arquivo = form.get("arquivo")
    if not arquivo:
        return HTMLResponse(render_page("Erro", '<div class="card" style="background:var(--red-bg); color:var(--red);">Nenhum arquivo enviado.</div><a href="/admin/importar-questoes" class="btn">← Voltar</a>'))

    import json as _json
    try:
        conteudo = await arquivo.read()
        dados = _json.loads(conteudo.decode("utf-8"))
    except Exception as e:
        return HTMLResponse(render_page("Erro", f'<div class="card" style="background:var(--red-bg); color:var(--red);">Erro ao ler JSON: {html.escape(str(e))}</div><a href="/admin/importar-questoes" class="btn">← Voltar</a>'))

    questoes_raw = dados.get("questoes", [])
    if not isinstance(questoes_raw, list) or not questoes_raw:
        return HTMLResponse(render_page("Erro", '<div class="card" style="background:var(--red-bg); color:var(--red);">JSON inválido: chave <code>questoes</code> ausente ou vazia.</div><a href="/admin/importar-questoes" class="btn">← Voltar</a>'))

    # Validação de cada questão
    erros = []
    questoes_validas = []
    for i, q in enumerate(questoes_raw, start=1):
        problemas = []
        if not q.get("disciplina"): problemas.append("disciplina ausente")
        if not q.get("enunciado"): problemas.append("enunciado ausente")
        tipo = q.get("tipo", "multipla_escolha")
        if tipo != "multipla_escolha": problemas.append(f"tipo '{tipo}' não suportado por importação ainda")
        alts = q.get("alternativas", [])
        if not isinstance(alts, list) or len(alts) < 2:
            problemas.append("alternativas insuficientes (mínimo 2)")
        else:
            n_corretas = sum(1 for a in alts if a.get("correta"))
            if n_corretas != 1:
                problemas.append(f"deve ter exatamente 1 alternativa correta ({n_corretas} marcadas)")
        if problemas:
            erros.append(f"<li>Questão #{i}: {', '.join(problemas)}</li>")
        else:
            questoes_validas.append(q)

    # Armazena o JSON na sessão pra confirmar depois (codifica em base64 pra ficar na URL)
    import base64
    payload_b64 = base64.urlsafe_b64encode(_json.dumps({"questoes": questoes_validas}).encode("utf-8")).decode("ascii")

    # Tabela de preview
    rows = ""
    for i, q in enumerate(questoes_validas[:50], start=1):  # mostra até 50
        alt_corr = next((a["letra"] for a in q["alternativas"] if a.get("correta")), "?")
        enun_curto = re.sub(r'<[^>]+>', '', q["enunciado"])[:120]
        fonte = q.get("fonte", "—")
        rows += f"""
            <tr>
                <td>{i}</td>
                <td>{html.escape(q.get('disciplina', '—'))}</td>
                <td>{html.escape(q.get('ano', '—'))}</td>
                <td style="max-width:400px;">{html.escape(enun_curto)}{"..." if len(q["enunciado"]) > 120 else ""}</td>
                <td><span class="badge-success badge">✓ {alt_corr}</span></td>
                <td style="font-size:11px; color:var(--text-muted);">{html.escape(fonte)}</td>
            </tr>
        """
    if len(questoes_validas) > 50:
        rows += f'<tr><td colspan="6" style="text-align:center; color:var(--text-muted); padding:14px;">… e mais {len(questoes_validas) - 50} questões válidas (não exibidas pra economizar espaço)</td></tr>'

    erros_html = ""
    if erros:
        erros_html = f"""
            <div class="card" style="background:var(--orange-bg); border-left:3px solid var(--orange); margin-top:14px;">
                <strong style="color:var(--orange);">⚠️ {len(erros)} questão(ões) com problemas (serão ignoradas):</strong>
                <ul style="margin-top:6px; font-size:13px;">{"".join(erros)}</ul>
            </div>
        """

    content = f"""
        <div class="page-header">
            <h1>👀 Preview da importação</h1>
            <p class="subtitle">{len(questoes_validas)} questões prontas pra importar · {len(erros)} com erros</p>
        </div>

        {erros_html}

        <div style="overflow-x:auto; margin-top:14px;">
            <table>
                <thead>
                    <tr><th>#</th><th>Disciplina</th><th>Ano</th><th>Enunciado (resumo)</th><th>Gabarito</th><th>Fonte</th></tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
        </div>

        <form method="post" action="/admin/importar-questoes/confirmar" style="margin-top:18px;">
            <input type="hidden" name="payload" value="{payload_b64}">
            <button type="submit" class="btn btn-primary" {"disabled" if not questoes_validas else ""}>
                ✓ Confirmar importação de {len(questoes_validas)} questões
            </button>
            <a href="/admin/importar-questoes" class="btn">← Voltar</a>
        </form>
    """
    return HTMLResponse(render_page("Preview da importação", content, active=""))


@app.post("/admin/importar-questoes/confirmar", response_class=HTMLResponse)
async def confirmar_importar_questoes(request: Request):
    """Efetiva a importação: cria disciplinas/habilidades novas se necessário, insere questões + alternativas."""
    _r = _require_admin_or_403(request)
    if _r is not None: return _r

    prof = get_current_professor(request)
    form = await request.form()
    import json as _json, base64
    try:
        payload_b64 = form.get("payload", "")
        dados = _json.loads(base64.urlsafe_b64decode(payload_b64.encode("ascii")).decode("utf-8"))
        questoes = dados.get("questoes", [])
    except Exception as e:
        return HTMLResponse(render_page("Erro", f'<div class="card" style="background:var(--red-bg); color:var(--red);">Erro ao decodificar payload: {html.escape(str(e))}</div>'))

    conn = get_db()
    importadas = 0
    bncc_criadas = 0
    disciplinas_criadas = 0

    try:
        for q in questoes:
            # 1. Disciplina (cria se não existir)
            disc_nome = q["disciplina"].strip()
            row = conn.execute("SELECT id FROM disciplinas WHERE LOWER(nome) = LOWER(?)", (disc_nome,)).fetchone()
            if row:
                disc_id = row["id"]
            else:
                c = conn.execute("INSERT INTO disciplinas (nome) VALUES (?)", (disc_nome,))
                disc_id = c.lastrowid
                disciplinas_criadas += 1

            # 2. Enunciado com fonte apêndice
            enunciado = q["enunciado"]
            fonte = q.get("fonte", "").strip()
            if fonte:
                enunciado = enunciado + f'<p style="font-size:11px; color:var(--text-muted); margin-top:10px; font-style:italic;">📚 Fonte: {html.escape(fonte)}</p>'

            # 3. Insere questão
            c = conn.execute(
                "INSERT INTO questoes (disciplina_id, enunciado, ano, tipo, criada_por_professor_id) VALUES (?, ?, ?, ?, ?)",
                (disc_id, enunciado, q.get("ano", ""), q.get("tipo", "multipla_escolha"), prof["id"])
            )
            qid = c.lastrowid

            # 4. Alternativas
            for a in q["alternativas"]:
                conn.execute(
                    "INSERT INTO alternativas (questao_id, letra, texto, correta) VALUES (?, ?, ?, ?)",
                    (qid, a["letra"].upper(), a["texto"], 1 if a.get("correta") else 0)
                )

            # 5. Habilidade BNCC (vincula se existe; cria fantasma se não existir)
            bncc = q.get("habilidade_bncc", "").strip()
            if bncc:
                row = conn.execute("SELECT id FROM habilidades_bncc WHERE codigo = ?", (bncc,)).fetchone()
                if not row:
                    c2 = conn.execute("INSERT INTO habilidades_bncc (codigo, descricao) VALUES (?, ?)", (bncc, "(importada — sem descrição)"))
                    h_id = c2.lastrowid
                    bncc_criadas += 1
                else:
                    h_id = row["id"]
                conn.execute("INSERT INTO questao_habilidades (questao_id, habilidade_id) VALUES (?, ?)", (qid, h_id))

            importadas += 1

        conn.commit()
    finally:
        conn.close()

    content = f"""
        <div class="page-header">
            <h1>✅ Importação concluída</h1>
        </div>
        <div class="card" style="background:var(--green-bg); border-left:3px solid var(--green);">
            <h3 style="margin-top:0; color:var(--green);">🎉 {importadas} questão(ões) cadastrada(s) com sucesso!</h3>
            <ul style="margin-top:8px; font-size:13px;">
                <li>{importadas} questões inseridas no banco coletivo</li>
                {"<li>" + str(disciplinas_criadas) + " disciplina(s) criada(s) automaticamente</li>" if disciplinas_criadas else ""}
                {"<li>" + str(bncc_criadas) + " código(s) BNCC novo(s) cadastrado(s)</li>" if bncc_criadas else ""}
            </ul>
        </div>
        <div class="page-actions" style="margin-top:18px;">
            <a href="/questoes" class="btn btn-primary">Ver banco de questões →</a>
            <a href="/admin/importar-questoes" class="btn">Importar mais</a>
        </div>
    """
    return HTMLResponse(render_page("Importação concluída", content, active=""))


@app.get("/habilidades", response_class=HTMLResponse)
def listar_habilidades(request: Request):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    habs = conn.execute("SELECT h.id, h.codigo, h.descricao, COUNT(qh.id) AS uso FROM habilidades_bncc h LEFT JOIN questao_habilidades qh ON qh.habilidade_id = h.id GROUP BY h.id ORDER BY h.codigo").fetchall()
    total = len(habs)
    com_desc = sum(1 for h in habs if (h["descricao"] or "").strip())
    sem_desc = total - com_desc
    conn.close()

    acoes_html = (
        f'<div class="page-actions">'
        f'<a href="/habilidades/importar" class="btn btn-primary">📥 Importar BNCC (Excel/CSV)</a>'
        f'</div>'
    )

    metricas_html = ""
    if total > 0:
        metricas_html = f"""
        <div class="metric-grid">
            <div class="metric"><div class="metric-label">Total cadastradas</div><div class="metric-value">{total}</div></div>
            <div class="metric"><div class="metric-label">Com descrição</div><div class="metric-value">{com_desc}</div></div>
            <div class="metric"><div class="metric-label">Sem descrição</div><div class="metric-value">{sem_desc}</div></div>
        </div>"""

    if habs:
        items = ""
        for h in habs:
            desc = h["descricao"] or '<em style="color:var(--text-subtle)">sem descrição</em>'
            items += f'<a href="/habilidades/{h["id"]}/editar" class="card card-link"><div class="card-title"><span class="badge">{h["codigo"]}</span></div><div class="card-meta">{desc}</div><div class="card-meta">{h["uso"]} questões usam essa habilidade</div></a>'
        body = items
    else:
        body = '<div class="empty"><p>Nenhuma habilidade cadastrada ainda.</p><p style="font-size:13px;">Use o botão <strong>Importar BNCC</strong> acima para subir a planilha oficial do MEC (1.408 habilidades do Ensino Fundamental), ou cadastre códigos digitando-os no campo BNCC ao criar questões.</p></div>'
    content = f'<div class="page-header"><h1>Habilidades BNCC</h1><p class="subtitle">Catálogo de códigos da BNCC. Clique numa habilidade para editar a descrição.</p>{acoes_html}</div>{metricas_html}{body}'
    return render_page("Habilidades BNCC", content, active="habilidades")


@app.get("/habilidades/{id}/editar", response_class=HTMLResponse)
def form_editar_habilidade(id: int):
    conn = get_db()
    h = conn.execute("SELECT * FROM habilidades_bncc WHERE id = ?", (id,)).fetchone()
    conn.close()
    if not h:
        return RedirectResponse("/habilidades", status_code=303)
    content = f'<div class="page-header"><h1><span class="badge">{h["codigo"]}</span></h1><p class="subtitle">Editar descrição.</p></div><form action="/habilidades/{id}/editar" method="post"><label>Descrição<textarea name="descricao" rows="4">{h["descricao"] or ""}</textarea></label><div class="page-actions"><button type="submit" class="btn btn-primary">Salvar</button><a href="/habilidades" class="btn">Cancelar</a></div></form>'
    return render_page(f"Editar {h['codigo']}", content, active="habilidades")


@app.post("/habilidades/{id}/editar")
def atualizar_habilidade(id: int, descricao: str = Form("")):
    conn = get_db()
    conn.execute("UPDATE habilidades_bncc SET descricao = ? WHERE id = ?", (descricao.strip() or None, id))
    conn.commit()
    conn.close()
    return RedirectResponse("/habilidades", status_code=303)


@app.get("/questoes", response_class=HTMLResponse)
def listar_questoes(request: Request, disciplina: Optional[str] = None, ano: Optional[str] = None, bncc: Optional[str] = None, q: Optional[str] = None):
    disciplina_id: Optional[int] = int(disciplina) if (disciplina and disciplina.strip().isdigit()) else None
    prof = get_current_professor(request)
    is_admin = bool(prof and prof["is_admin"])
    conn = get_db()

    # Montar query com filtros aplicados — agora trazendo também o autor
    sql = """
        SELECT DISTINCT q.id, q.enunciado, q.ano, q.criada_por_professor_id, q.tipo,
               d.id AS disciplina_id, d.nome AS disciplina_nome,
               aut.nome AS autor_nome
        FROM questoes q
        JOIN disciplinas d ON d.id = q.disciplina_id
        LEFT JOIN professores aut ON aut.id = q.criada_por_professor_id
        LEFT JOIN questao_habilidades qh ON qh.questao_id = q.id
        LEFT JOIN habilidades_bncc h ON h.id = qh.habilidade_id
        WHERE 1=1
    """
    params = []
    if disciplina_id:
        sql += " AND d.id = ?"
        params.append(disciplina_id)
    if ano:
        sql += " AND q.ano = ?"
        params.append(ano)
    if bncc and bncc.strip():
        sql += " AND h.codigo LIKE ?"
        params.append(f"%{bncc.strip().upper()}%")
    if q and q.strip():
        sql += " AND q.enunciado LIKE ?"
        params.append(f"%{q.strip()}%")
    sql += " ORDER BY d.nome, q.id DESC"

    questoes = conn.execute(sql, params).fetchall()
    disciplinas = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()
    total_geral = conn.execute("SELECT COUNT(*) AS c FROM questoes").fetchone()["c"]

    # Matriz disciplina × ano com contagens — apenas disciplinas que têm pelo menos 1 questão
    matriz_rows = conn.execute("""
        SELECT d.id AS disc_id, d.nome AS disc_nome, q.ano, COUNT(q.id) AS qtd
        FROM questoes q JOIN disciplinas d ON d.id = q.disciplina_id
        GROUP BY d.id, q.ano
        ORDER BY d.nome COLLATE NOCASE, q.ano IS NULL, q.ano
    """).fetchall()

    # Organiza por disciplina
    from collections import defaultdict
    matriz_por_disc = defaultdict(list)
    nomes_disc = {}
    for r in matriz_rows:
        matriz_por_disc[r["disc_id"]].append((r["ano"] or "", r["qtd"]))
        nomes_disc[r["disc_id"]] = r["disc_nome"]

    matriz_html = ""
    if matriz_por_disc:
        import urllib.parse as _urlp
        linhas = []
        for disc_id, anos_qtds in matriz_por_disc.items():
            badges_linha = []
            for ano_v, qtd in anos_qtds:
                rotulo = ano_v if ano_v else "Sem ano"
                # Constrói query string preservando o filtro
                qs = _urlp.urlencode({"disciplina": disc_id, "ano": ano_v} if ano_v else {"disciplina": disc_id})
                # Destaca badge se o filtro atual bate
                ativo = (disciplina_id == disc_id and ((ano_v and ano == ano_v) or (not ano_v and not ano)))
                cor_bg = "var(--accent)" if ativo else "var(--bg)"
                cor_fg = "white" if ativo else "var(--text)"
                borda = "var(--accent)" if ativo else "var(--border)"
                badges_linha.append(
                    f'<a href="/questoes?{qs}" class="badge" style="background:{cor_bg}; color:{cor_fg}; '
                    f'border:1px solid {borda}; text-decoration:none; padding:3px 9px; font-size:11px;">'
                    f'{rotulo}: {qtd}</a>'
                )
            linhas.append(
                f'<div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">'
                f'<strong style="min-width:140px; font-size:13px;">{nomes_disc[disc_id]}</strong>'
                f'{"".join(badges_linha)}'
                f'</div>'
            )
        matriz_html = (
            f'<div style="background:var(--bg-subtle); padding:14px 16px; border-radius:8px; margin-bottom:14px;">'
            f'<div style="font-size:11px; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-muted); margin-bottom:10px;">Visão geral · clique para filtrar</div>'
            f'<div style="display:flex; flex-direction:column; gap:8px;">{"".join(linhas)}</div>'
            f'</div>'
        )

    disciplinas_opts = '<option value="">Todas</option>' + "".join(
        f'<option value="{d["id"]}"{(" selected" if disciplina_id == d["id"] else "")}>{d["nome"]}</option>'
        for d in disciplinas
    )
    anos_opts = '<option value="">Todos</option>' + "".join(
        f'<option value="{a}"{(" selected" if ano == a else "")}>{a}</option>'
        for a in ANOS
    )

    filtros_html = (
        f'<form action="/questoes" method="get" '
        f'style="background:var(--bg-subtle); padding:14px 16px; border-radius:8px; margin-bottom:18px;">'
        f'<div style="display:grid; grid-template-columns: 1.5fr 1.2fr 1.2fr 2fr auto auto; gap:10px; align-items:end;">'
        f'<label style="margin:0;">Disciplina<select name="disciplina">{disciplinas_opts}</select></label>'
        f'<label style="margin:0;">Ano<select name="ano">{anos_opts}</select></label>'
        f'<label style="margin:0;">Código BNCC<input type="text" name="bncc" placeholder="EF06MA" value="{bncc or ""}"></label>'
        f'<label style="margin:0;">Buscar no enunciado<input type="text" name="q" placeholder="palavra-chave" value="{q or ""}"></label>'
        f'<button type="submit" class="btn btn-primary" style="margin:0;">Filtrar</button>'
        f'<a href="/questoes" class="btn" style="margin:0;">Limpar</a>'
        f'</div></form>'
    )

    if questoes:
        cards_list = []
        for qx in questoes:
            pode_ed = _pode_editar_questao(prof, qx["criada_por_professor_id"])
            # Badge "Por: X" só pra admin (e quando o autor é diferente do admin logado)
            mostrar_autor = is_admin and qx["autor_nome"] and qx["criada_por_professor_id"] != prof["id"]
            autor_nome_card = qx["autor_nome"] if mostrar_autor else None
            cards_list.append(render_questao_card(
                conn, qx, mostrar_acoes=True, compact=True,
                pode_editar=pode_ed, autor_nome=autor_nome_card
            ))
        cards = "".join(cards_list)
    else:
        cards = '<div class="empty">Nenhuma questão encontrada com os filtros selecionados.</div>'
    conn.close()

    tem_filtro = bool(disciplina or ano or bncc or q)
    subtitle = f'{len(questoes)} de {total_geral} questão(ões)' if tem_filtro else f'{total_geral} questão(ões) cadastradas'

    toggle_js = """
    <script>
    function toggleQuestao(id) {
        const detalhes = document.getElementById('q-detalhes-' + id);
        const btn = document.getElementById('q-toggle-' + id);
        if (detalhes.style.display === 'none' || !detalhes.style.display) {
            detalhes.style.display = 'block';
            btn.textContent = 'Recolher ▴';
            if (window.MathJax && MathJax.typesetPromise) { MathJax.typesetPromise([detalhes]); }
        } else {
            detalhes.style.display = 'none';
            btn.textContent = 'Ver completa ▾';
        }
    }
    </script>
    """

    content = (
        f'<div class="page-header"><h1>Banco de questões</h1>'
        f'<p class="subtitle">{subtitle}</p>'
        f'<div class="page-actions"><a href="/questoes/nova" class="btn btn-primary">+ Nova questão</a></div></div>'
        f'{matriz_html}{filtros_html}{cards}{toggle_js}'
    )
    return render_page("Questões", content, active="questoes", head_extra=MATHJAX)


@app.get("/questoes/nova", response_class=HTMLResponse)
def form_nova_questao_passo1():
    """Passo 1: seleciona disciplina, ano e habilidades BNCC antes do cadastro completo."""
    conn = get_db()
    disciplinas = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()
    habs_existentes = conn.execute("SELECT codigo FROM habilidades_bncc ORDER BY codigo").fetchall()
    conn.close()
    if not disciplinas:
        return render_page("Nova questão", '<div class="page-header"><h1>Nova questão</h1></div><div class="empty"><p>Você precisa cadastrar pelo menos uma disciplina antes de criar questões.</p><a href="/disciplinas/nova" class="btn btn-primary">Cadastrar disciplina</a></div>', active="questoes")

    options = "".join(f'<option value="{d["id"]}">{d["nome"]}</option>' for d in disciplinas)
    anos_options = '<option value="">— Não definido —</option>' + "".join(f'<option value="{a}">{a}</option>' for a in ANOS)

    total_habs = len(habs_existentes)
    link_catalogo = (
        f'<p class="muted-line" style="font-size:11px;">'
        f'💡 {total_habs} habilidade(s) cadastrada(s) no catálogo. '
        f'<a href="/habilidades" target="_blank" style="color:var(--text-muted);">Consultar lista completa</a>'
        f'</p>'
    ) if total_habs > 0 else '<p class="muted-line" style="font-size:11px;">Nenhuma habilidade cadastrada ainda. <a href="/habilidades/importar" target="_blank">Importar BNCC oficial</a>.</p>'

    js_preview = """
    <script>
    (function() {
        const ta = document.querySelector('textarea[name="habilidades_codigos"]');
        const discSel = document.querySelector('select[name="disciplina_id"]');
        if (!ta) return;

        // === Painel de validação dos códigos digitados ===
        const preview = document.createElement('div');
        preview.id = 'bncc-preview';
        preview.style.cssText = 'margin-top:6px; font-size:12px; line-height:1.5;';
        ta.parentNode.appendChild(preview);

        async function validar() {
            const codigos = ta.value.split(/[,\\n]/).map(c => c.trim().toUpperCase()).filter(c => c);
            if (codigos.length === 0) { preview.innerHTML = ''; return; }
            try {
                const resp = await fetch('/habilidades/buscar?codigos=' + encodeURIComponent(codigos.join(',')));
                const data = await resp.json();
                let html = '';
                for (const c of codigos) {
                    if (data[c]) {
                        html += '<div style="padding:4px 8px; background:var(--green-bg); border-left:3px solid var(--green); margin-bottom:3px; color:var(--text);"><strong style="color:var(--green);">' + c + '</strong>: ' + data[c].replace(/</g, '&lt;') + '</div>';
                    } else {
                        html += '<div style="padding:4px 8px; background:var(--red-bg); border-left:3px solid var(--red); margin-bottom:3px; color:var(--red);"><strong style="color:var(--red);">' + c + '</strong>: ⚠ código não encontrado no catálogo (será criado sem descrição)</div>';
                    }
                }
                preview.innerHTML = html;
            } catch (e) { preview.innerHTML = ''; }
        }
        ta.addEventListener('blur', validar);
        ta.addEventListener('input', () => { if (ta._t) clearTimeout(ta._t); ta._t = setTimeout(validar, 600); });

        // === Busca por palavra/conceito ===
        const buscaWrap = document.createElement('div');
        buscaWrap.style.cssText = 'margin-top:14px; padding:12px; background:var(--bg-subtle); border-radius:6px;';
        buscaWrap.innerHTML = '<label style="margin:0; font-size:12px;">🔍 Não sabe o código? Busque por palavra ou conceito<input type="search" id="bncc-busca" placeholder="ex: fração, Constituição, fotossíntese" style="margin-top:4px;"></label><div id="bncc-resultados" style="margin-top:8px; font-size:12px;"></div>';
        ta.parentNode.appendChild(buscaWrap);

        const inputBusca = buscaWrap.querySelector('#bncc-busca');
        const divRes = buscaWrap.querySelector('#bncc-resultados');

        async function buscarPorPalavra() {
            const q = inputBusca.value.trim();
            if (q.length < 2) { divRes.innerHTML = ''; return; }
            const disc = discSel ? discSel.value : '';
            const url = '/habilidades/buscar?q=' + encodeURIComponent(q) + (disc ? '&disciplina_id=' + disc : '');
            try {
                const resp = await fetch(url);
                const data = await resp.json();
                const results = data.results || [];
                if (results.length === 0) {
                    divRes.innerHTML = '<div style="color:var(--text-muted); padding:8px 0;">Nenhum resultado para "' + q.replace(/</g, '&lt;') + '"' + (disc ? ' na disciplina selecionada' : '') + '.</div>';
                    return;
                }
                const escopo = disc ? ' (filtrado pela disciplina)' : ' (todas as disciplinas)';
                let html = '<div style="color:var(--text-muted); padding:4px 0;">' + results.length + ' habilidade(s) encontrada(s)' + escopo + ' — clique para adicionar:</div>';
                for (const r of results) {
                    html += '<div data-codigo="' + r.codigo + '" style="padding:6px 8px; border:1px solid var(--border); border-radius:4px; margin-bottom:4px; cursor:pointer; background:var(--bg); color:var(--text);" onmouseover="this.style.background=\\'var(--accent-bg)\\'" onmouseout="this.style.background=\\'var(--bg)\\'"><strong style="color:var(--accent);">' + r.codigo + '</strong> · ' + r.descricao.replace(/</g, '&lt;') + '</div>';
                }
                divRes.innerHTML = html;
            } catch (e) { divRes.innerHTML = ''; }
        }
        inputBusca.addEventListener('input', () => { if (inputBusca._t) clearTimeout(inputBusca._t); inputBusca._t = setTimeout(buscarPorPalavra, 400); });
        if (discSel) discSel.addEventListener('change', buscarPorPalavra);

        divRes.addEventListener('click', (e) => {
            const item = e.target.closest('[data-codigo]');
            if (!item) return;
            const codigo = item.dataset.codigo;
            const cur = ta.value.trim();
            const codigos = cur ? cur.split(/[,\\n]/).map(c => c.trim().toUpperCase()).filter(c => c) : [];
            if (codigos.includes(codigo)) return;
            codigos.push(codigo);
            ta.value = codigos.join(', ');
            validar();
        });
    })();
    </script>
    """

    tipo_options = "".join(
        f'<option value="{k}">{v["icone"]} {v["label"]}</option>'
        for k, v in TIPOS_QUESTAO.items()
    )

    content = f"""
        <div class="page-header">
            <h1>Nova questão</h1>
            <p class="subtitle">Passo 1 de 2 — defina o tipo, disciplina, ano e habilidades.</p>
        </div>
        <form action="/questoes/nova/passo2" method="post">
            <label>Tipo de questão<select name="tipo" required>{tipo_options}</select></label>
            <div style="display:grid; grid-template-columns: 2fr 1fr; gap:12px;">
                <label>Disciplina<select name="disciplina_id" required>{options}</select></label>
                <label>Ano de escolaridade<select name="ano">{anos_options}</select></label>
            </div>
            <label>Habilidades BNCC (opcional, separadas por vírgula ou uma por linha)<textarea name="habilidades_codigos" rows="2" placeholder="EF09MA09, EF09MA10"></textarea></label>
            {link_catalogo}
            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Próximo: cadastrar conteúdo →</button>
                <a href="/questoes" class="btn">Cancelar</a>
            </div>
        </form>
        {js_preview}
    """
    return render_page("Nova questão · Passo 1", content, active="questoes")


@app.post("/questoes/nova/passo2", response_class=HTMLResponse)
def form_nova_questao_passo2(
    disciplina_id: int = Form(...),
    ano: str = Form(""),
    habilidades_codigos: str = Form(""),
    tipo: str = Form("multipla_escolha"),
):
    """Passo 2: cadastramento do conteúdo da questão. Dados do Passo 1 ficam em hidden fields."""
    if tipo not in TIPOS_QUESTAO:
        tipo = "multipla_escolha"
    conn = get_db()
    disciplina = conn.execute("SELECT * FROM disciplinas WHERE id = ?", (disciplina_id,)).fetchone()
    conn.close()
    if not disciplina:
        return RedirectResponse("/questoes/nova", status_code=303)

    ano_label = ano if ano else "— não definido —"
    tipo_info = TIPOS_QUESTAO[tipo]

    # Badges informativas das habilidades digitadas no passo 1
    codigos_clean = [c.strip().upper() for c in habilidades_codigos.replace("\n", ",").split(",") if c.strip()]
    badges_bncc = "".join(f'<span class="badge">{c}</span>' for c in codigos_clean) if codigos_clean else '<span class="muted-line">— sem BNCC —</span>'

    # Bloco específico do tipo da questão
    fieldset_alternativas = ""
    enunciado_detecta_alts = (tipo == "multipla_escolha")
    if tipo == "multipla_escolha":
        alternativas_html = ""
        for letra in ["A", "B", "C", "D"]:
            required_radio = ' required' if letra == "A" else ''
            editor_alt = _editor_enunciado_html(
                name=f"alt_{letra.lower()}", valor_inicial="", required=True,
                label="", compact=True, min_height=42,
                placeholder=f"Texto da alternativa {letra}"
            )
            alternativas_html += (
                f'<div style="display:grid; grid-template-columns:auto 1fr; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<label style="margin:8px 0 0 0; display:flex; align-items:center; gap:8px; white-space:nowrap;">'
                f'<input type="radio" name="correta" value="{letra}"{required_radio} style="width:auto; margin:0;"> <strong>{letra})</strong>'
                f'</label>'
                f'<div style="margin:0;">{editor_alt}</div>'
                f'</div>'
            )
        fieldset_alternativas = f"""
            <fieldset>
                <legend>Alternativas — marque o radio da correta</legend>
                {alternativas_html}
            </fieldset>
        """
    elif tipo == "discursiva":
        # Discursiva: sem alternativas. Aviso visual.
        fieldset_alternativas = """
            <div style="background:var(--accent-bg); color:var(--accent); border:1px solid var(--accent); padding:14px 16px; border-radius:6px; margin:12px 0;">
                <strong>📝 Questão discursiva</strong><br>
                <span style="font-size:13px;">O aluno responderá em texto livre. No modo impresso, será reservado espaço para resposta manuscrita. A correção é manual — feita por você fora do sistema.</span>
            </div>
        """
    elif tipo == "vf":
        # V ou F: até 5 afirmações, cada uma com radio V/F
        afirms_html = ""
        for i in range(VF_MAX_AFIRMACOES):
            editor_afirm = _editor_enunciado_html(
                name=f"vf_afirm_{i}_texto", valor_inicial="", required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Afirmação {i+1} (deixe em branco se não usar)"
            )
            afirms_html += (
                f'<div style="display:grid; grid-template-columns:1fr auto; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<div style="margin:0;"><strong style="font-size:13px;">Afirmação {i+1}</strong>{editor_afirm}</div>'
                f'<div style="display:flex; gap:10px; align-items:center; padding-top:24px; white-space:nowrap;">'
                f'<label style="margin:0; font-size:13px;"><input type="radio" name="vf_afirm_{i}_gabarito" value="V" style="width:auto; margin:0 4px 0 0;">V</label>'
                f'<label style="margin:0; font-size:13px;"><input type="radio" name="vf_afirm_{i}_gabarito" value="F" style="width:auto; margin:0 4px 0 0;">F</label>'
                f'</div></div>'
            )
        fieldset_alternativas = f"""
            <fieldset>
                <legend>Afirmações — marque V ou F para cada (até {VF_MAX_AFIRMACOES})</legend>
                <p class="muted-line" style="font-size:12px; margin:0 0 10px 0;">Deixe em branco as afirmações que não usar (mínimo 2 afirmações preenchidas).</p>
                {afirms_html}
            </fieldset>
        """
    elif tipo == "associacao":
        # Associação: 2 colunas de até 5 itens; coluna A tem texto + qual letra da B é a resposta correta
        col_a_html = ""
        for i in range(ASSOC_MAX_PARES):
            editor_a = _editor_enunciado_html(
                name=f"assoc_a_{i}_texto", valor_inicial="", required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Item {i+1} da coluna A (em branco se não usar)"
            )
            # Select pra escolher qual letra da B é o gabarito
            letras_options = '<option value="">—</option>' + "".join(
                f'<option value="{chr(97+j)}">{chr(97+j)}</option>' for j in range(ASSOC_MAX_PARES)
            )
            col_a_html += (
                f'<div style="display:grid; grid-template-columns:auto 1fr auto; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<strong style="padding-top:20px;">{i+1}.</strong>'
                f'<div style="margin:0;">{editor_a}</div>'
                f'<label style="margin:0; padding-top:14px; font-size:12px; white-space:nowrap;">Resposta: '
                f'<select name="assoc_a_{i}_gabarito" style="width:auto; display:inline-block; margin-left:4px;">{letras_options}</select>'
                f'</label></div>'
            )
        col_b_html = ""
        for j in range(ASSOC_MAX_PARES):
            letra_b = chr(97+j)
            editor_b = _editor_enunciado_html(
                name=f"assoc_b_{letra_b}_texto", valor_inicial="", required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Item ({letra_b}) da coluna B (em branco se não usar)"
            )
            col_b_html += (
                f'<div style="display:grid; grid-template-columns:auto 1fr; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<strong style="padding-top:20px;">({letra_b})</strong>'
                f'<div style="margin:0;">{editor_b}</div>'
                f'</div>'
            )
        fieldset_alternativas = f"""
            <fieldset>
                <legend>Coluna A — itens (1, 2, 3...) com gabarito da resposta</legend>
                <p class="muted-line" style="font-size:12px; margin:0 0 10px 0;">Para cada item da coluna A, indique qual letra da coluna B é a resposta correta. Mínimo 2 pares preenchidos.</p>
                {col_a_html}
            </fieldset>
            <fieldset>
                <legend>Coluna B — opções de associação (a, b, c...)</legend>
                {col_b_html}
            </fieldset>
        """

    # Hidden fields carregam dados do passo 1; valores escapados
    import html as _html
    h_disc = _html.escape(str(disciplina_id), quote=True)
    h_ano = _html.escape(ano, quote=True)
    h_habs = _html.escape(habilidades_codigos, quote=True)
    h_tipo = _html.escape(tipo, quote=True)

    content = f"""
        <div class="page-header">
            <h1>Nova questão</h1>
            <p class="subtitle">Passo 2 de 2 — conteúdo da questão.</p>
        </div>

        <div style="background:var(--bg-subtle); border:1px solid var(--border); border-radius:8px; padding:12px 16px; margin-bottom:18px; display:flex; align-items:center; gap:16px; flex-wrap:wrap;">
            <div><strong>Tipo:</strong> {tipo_info['icone']} {tipo_info['label']}</div>
            <div><strong>Disciplina:</strong> {disciplina["nome"]}</div>
            <div><strong>Ano:</strong> {ano_label}</div>
            <div><strong>BNCC:</strong> {badges_bncc}</div>
            <a href="/questoes/nova" style="margin-left:auto; font-size:13px; color:var(--text-muted);">← Voltar e alterar</a>
        </div>

        <div class="tip"><strong>Dica:</strong> use <code>$fórmula$</code> para fórmulas inline ou <code>$$fórmula$$</code> para centralizadas.</div>

        <form action="/questoes/criar" method="post" enctype="multipart/form-data">
            <input type="hidden" name="disciplina_id" value="{h_disc}">
            <input type="hidden" name="ano" value="{h_ano}">
            <input type="hidden" name="habilidades_codigos" value="{h_habs}">
            <input type="hidden" name="tipo" value="{h_tipo}">

            <fieldset>
                <legend>Textos de apoio (opcionais)</legend>
                {_editor_enunciado_html(name="texto1_conteudo", valor_inicial="", required=False, label="Texto 1 — conteúdo", min_height=80, placeholder="Cole ou digite aqui o texto de apoio (opcional)")}
                <label>Texto 1 — fonte<input type="text" name="texto1_fonte" placeholder="Autor, obra, ano"></label>
                {_editor_enunciado_html(name="texto2_conteudo", valor_inicial="", required=False, label="Texto 2 — conteúdo", min_height=80, placeholder="Segundo texto de apoio (opcional)")}
                <label>Texto 2 — fonte<input type="text" name="texto2_fonte" placeholder="Autor, obra, ano"></label>
            </fieldset>

            <fieldset>
                <legend>Imagens (opcionais)</legend>
                <label>Imagem 1<input type="file" name="imagem1" accept="image/*"></label>
                <label>Legenda da imagem 1<input type="text" name="imagem1_legenda"></label>
                <label>Fonte da imagem 1<input type="text" name="imagem1_fonte"></label>
                <label>Imagem 2<input type="file" name="imagem2" accept="image/*"></label>
                <label>Legenda da imagem 2<input type="text" name="imagem2_legenda"></label>
                <label>Fonte da imagem 2<input type="text" name="imagem2_fonte"></label>
            </fieldset>

            {_editor_enunciado_html(name="enunciado", valor_inicial="", required=True, label="Enunciado", placeholder="Digite o enunciado da questão. Use a barra abaixo para formatar.", detectar_alternativas=enunciado_detecta_alts)}

            {fieldset_alternativas}

            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Cadastrar questão</button>
                <a href="/questoes/nova" class="btn">← Voltar</a>
            </div>
        </form>
    """
    return render_page("Nova questão · Passo 2", content, active="questoes", head_extra=MATHJAX)


@app.post("/questoes/criar")
async def criar_questao(
    request: Request,
    disciplina_id: int = Form(...), enunciado: str = Form(...),
    tipo: str = Form("multipla_escolha"),
    alt_a: str = Form(""), alt_b: str = Form(""), alt_c: str = Form(""), alt_d: str = Form(""),
    correta: str = Form(""), habilidades_codigos: str = Form(""),
    ano: str = Form(""),
    texto1_conteudo: str = Form(""), texto1_fonte: str = Form(""),
    texto2_conteudo: str = Form(""), texto2_fonte: str = Form(""),
    imagem1: Optional[UploadFile] = File(None), imagem1_legenda: str = Form(""), imagem1_fonte: str = Form(""),
    imagem2: Optional[UploadFile] = File(None), imagem2_legenda: str = Form(""), imagem2_fonte: str = Form(""),
):
    if tipo not in TIPOS_QUESTAO:
        tipo = "multipla_escolha"
    # Form recebe dinamicamente os campos de V/F e Associação; pega tudo via request
    form_extra = await request.form()
    prof = get_current_professor(request)
    prof_id = prof["id"] if prof else None
    conn = get_db()
    cursor = conn.execute(
        "INSERT INTO questoes (disciplina_id, enunciado, ano, criada_por_professor_id, tipo) VALUES (?, ?, ?, ?, ?)",
        (disciplina_id, _sanitizar_html_enunciado(enunciado), ano.strip() or None, prof_id, tipo)
    )
    questao_id = cursor.lastrowid

    for ordem, (conteudo, fonte) in enumerate([(texto1_conteudo, texto1_fonte), (texto2_conteudo, texto2_fonte)]):
        conteudo_sanit = _sanitizar_html_enunciado(conteudo)
        if conteudo_sanit:
            conn.execute("INSERT INTO textos_apoio (questao_id, conteudo, fonte, ordem) VALUES (?, ?, ?, ?)", (questao_id, conteudo_sanit, fonte.strip() or None, ordem))

    for ordem, (img, legenda, fonte) in enumerate([(imagem1, imagem1_legenda, imagem1_fonte), (imagem2, imagem2_legenda, imagem2_fonte)]):
        if img and img.filename:
            ext = os.path.splitext(img.filename)[1].lower()
            unique_name = f"{uuid.uuid4().hex}{ext}"
            file_path = os.path.join(UPLOAD_DIR, unique_name)
            content_bytes = await img.read()
            with open(file_path, "wb") as f:
                f.write(content_bytes)
            conn.execute("INSERT INTO imagens (questao_id, caminho, legenda, fonte, ordem) VALUES (?, ?, ?, ?, ?)", (questao_id, f"static/imagens/{unique_name}", legenda.strip() or None, fonte.strip() or None, ordem))

    # Conteúdo específico do tipo
    if tipo == "multipla_escolha":
        for letra, texto in [("A", alt_a), ("B", alt_b), ("C", alt_c), ("D", alt_d)]:
            conn.execute("INSERT INTO alternativas (questao_id, letra, texto, correta) VALUES (?, ?, ?, ?)", (questao_id, letra, _sanitizar_html_enunciado(texto), 1 if letra == correta else 0))
    elif tipo == "vf":
        ordem_real = 0
        for i in range(VF_MAX_AFIRMACOES):
            texto_afirm = _sanitizar_html_enunciado(str(form_extra.get(f"vf_afirm_{i}_texto", "")))
            gabarito = str(form_extra.get(f"vf_afirm_{i}_gabarito", "")).strip().upper()
            if texto_afirm and gabarito in ("V", "F"):
                conn.execute("INSERT INTO vf_afirmacoes (questao_id, ordem, texto, gabarito) VALUES (?, ?, ?, ?)",
                             (questao_id, ordem_real, texto_afirm, gabarito))
                ordem_real += 1
    elif tipo == "associacao":
        # Coluna A (com gabarito)
        ordem_real = 0
        for i in range(ASSOC_MAX_PARES):
            texto_a = _sanitizar_html_enunciado(str(form_extra.get(f"assoc_a_{i}_texto", "")))
            gabarito = str(form_extra.get(f"assoc_a_{i}_gabarito", "")).strip().lower()
            if texto_a and gabarito:
                conn.execute("INSERT INTO assoc_itens_a (questao_id, ordem, texto, gabarito_letra) VALUES (?, ?, ?, ?)",
                             (questao_id, ordem_real, texto_a, gabarito))
                ordem_real += 1
        # Coluna B (opções)
        for j in range(ASSOC_MAX_PARES):
            letra_b = chr(97+j)
            texto_b = _sanitizar_html_enunciado(str(form_extra.get(f"assoc_b_{letra_b}_texto", "")))
            if texto_b:
                conn.execute("INSERT INTO assoc_itens_b (questao_id, letra, texto) VALUES (?, ?, ?)",
                             (questao_id, letra_b, texto_b))

    for parte in habilidades_codigos.replace("\n", ",").split(","):
        codigo = parte.strip().upper()
        if not codigo: continue
        existing = conn.execute("SELECT id FROM habilidades_bncc WHERE codigo = ?", (codigo,)).fetchone()
        habilidade_id = existing["id"] if existing else conn.execute("INSERT INTO habilidades_bncc (codigo) VALUES (?)", (codigo,)).lastrowid
        try:
            conn.execute("INSERT INTO questao_habilidades (questao_id, habilidade_id) VALUES (?, ?)", (questao_id, habilidade_id))
        except sqlite3.IntegrityError:
            pass

    conn.commit()
    conn.close()
    return RedirectResponse("/questoes", status_code=303)


# ==========================================
#  ROTAS DE PROVAS (ATUALIZADAS TAREFA A2)
# ==========================================

@app.get("/provas", response_class=HTMLResponse)
def listar_provas(request: Request, disciplina: Optional[str] = None, ano: Optional[str] = None, q: Optional[str] = None):
    disciplina_id: Optional[int] = int(disciplina) if (disciplina and disciplina.strip().isdigit()) else None
    prof = get_current_professor(request)
    is_admin = prof and prof["is_admin"]
    conn = get_db()

    # Filtros: admin vê tudo da escola; prof comum só as próprias
    where_extras = []
    params = []
    if not is_admin:
        where_extras.append("(p.criada_por_professor_id = ? OR p.criada_por_professor_id IS NULL)")
        params.append(prof["id"])
    if q and q.strip():
        where_extras.append("p.titulo LIKE ?")
        params.append(f"%{q.strip()}%")
    if disciplina_id:
        where_extras.append("""EXISTS (
            SELECT 1 FROM prova_questoes pq2
            JOIN questoes q2 ON q2.id = pq2.questao_id
            WHERE pq2.prova_id = p.id AND q2.disciplina_id = ?
        )""")
        params.append(disciplina_id)
    if ano and ano.strip():
        where_extras.append("""EXISTS (
            SELECT 1 FROM prova_questoes pq3
            JOIN questoes q3 ON q3.id = pq3.questao_id
            WHERE pq3.prova_id = p.id AND q3.ano = ?
        )""")
        params.append(ano)
    where_clause = " WHERE " + " AND ".join(where_extras) if where_extras else ""

    sql = f"""
        SELECT p.id, p.titulo, p.descricao, p.criada_por_professor_id,
               prof.nome AS criador_nome,
               (SELECT COUNT(*) FROM prova_questoes WHERE prova_id = p.id) AS qtd_questoes
        FROM provas p
        LEFT JOIN professores prof ON prof.id = p.criada_por_professor_id
        {where_clause}
        ORDER BY p.id DESC
    """
    provas = conn.execute(sql, params).fetchall()

    # Tags (disciplinas + anos) de cada prova — query única pra todas
    tags_map = {}
    if provas:
        prova_ids = [p["id"] for p in provas]
        placeholders = ",".join("?" * len(prova_ids))
        tags_rows = conn.execute(f"""
            SELECT pq.prova_id, d.nome AS disc_nome, q.ano
            FROM prova_questoes pq
            JOIN questoes q ON q.id = pq.questao_id
            JOIN disciplinas d ON d.id = q.disciplina_id
            WHERE pq.prova_id IN ({placeholders})
        """, prova_ids).fetchall()
        for r in tags_rows:
            tm = tags_map.setdefault(r["prova_id"], {"disciplinas": set(), "anos": set()})
            tm["disciplinas"].add(r["disc_nome"])
            if r["ano"]:
                tm["anos"].add(r["ano"])

    # Aplicações por prova (pra mostrar no card)
    apl_count = {row["prova_id"]: row["c"] for row in conn.execute(
        "SELECT prova_id, COUNT(*) AS c FROM aplicacoes GROUP BY prova_id"
    ).fetchall()}

    disciplinas_lista = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()
    total_geral = conn.execute("SELECT COUNT(*) AS c FROM provas").fetchone()["c"]
    conn.close()

    # Filtros
    disciplinas_opts = '<option value="">Todas</option>' + "".join(
        f'<option value="{d["id"]}"{(" selected" if disciplina_id == d["id"] else "")}>{d["nome"]}</option>'
        for d in disciplinas_lista
    )
    anos_opts = '<option value="">Todos</option>' + "".join(
        f'<option value="{a}"{(" selected" if ano == a else "")}>{a}</option>'
        for a in ANOS
    )
    filtros_html = (
        f'<form action="/provas" method="get" '
        f'style="background:var(--bg-subtle); padding:14px 16px; border-radius:8px; margin-bottom:18px;">'
        f'<div style="display:grid; grid-template-columns: 2fr 1.2fr 1.2fr auto auto; gap:10px; align-items:end;">'
        f'<label style="margin:0;">Buscar por título<input type="text" name="q" placeholder="palavra do título" value="{q or ""}"></label>'
        f'<label style="margin:0;">Disciplina<select name="disciplina">{disciplinas_opts}</select></label>'
        f'<label style="margin:0;">Ano<select name="ano">{anos_opts}</select></label>'
        f'<button type="submit" class="btn btn-primary" style="margin:0;">Filtrar</button>'
        f'<a href="/provas" class="btn" style="margin:0;">Limpar</a>'
        f'</div></form>'
    )

    # Cards
    if provas:
        cards = ""
        for p in provas:
            tm = tags_map.get(p["id"], {"disciplinas": set(), "anos": set()})
            disc_tags = "".join(f'<span class="badge" style="background:var(--accent-bg); color:var(--accent);">{d}</span>' for d in sorted(tm["disciplinas"]))
            ano_tags = "".join(f'<span class="badge">{a}</span>' for a in sorted(tm["anos"]))
            desc = f'<div style="font-size:13px; color:var(--text-muted); margin-top:4px;">{p["descricao"]}</div>' if p["descricao"] else ""
            n_apl = apl_count.get(p["id"], 0)
            apl_badge = f'<span class="badge" style="background:var(--orange-bg); color:var(--orange);">{n_apl} aplicação{"" if n_apl == 1 else "ões"}</span>' if n_apl else ""

            # Badge "Por: <nome>" só pra admin (pra ele saber de quem é cada prova)
            autor_badge = ""
            if is_admin:
                nome_autor = p["criador_nome"] if p["criador_nome"] else "—"
                autor_badge = f'<span class="badge" style="background:var(--purple-bg); color:var(--purple);">Por: {nome_autor}</span>'

            cards += f"""
            <div style="background:var(--bg); border:1px solid var(--border); border-radius:8px; padding:14px 18px; margin-bottom:10px;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:14px;">
                    <div style="flex:1; min-width:0;">
                        <div style="font-weight:600; font-size:16px;">
                            <a href="/provas/{p["id"]}" style="color:inherit; text-decoration:none;">{p["titulo"]}</a>
                        </div>
                        {desc}
                        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:8px; align-items:center;">
                            <span class="badge">{p["qtd_questoes"]} questões</span>
                            {disc_tags}{ano_tags}{apl_badge}{autor_badge}
                        </div>
                    </div>
                    <div style="display:flex; gap:6px; flex-shrink:0;">
                        <a href="/provas/{p["id"]}" class="btn" style="padding:4px 10px; font-size:12px;">Abrir</a>
                        <a href="/provas/{p["id"]}/editar" class="btn" style="padding:4px 10px; font-size:12px;">Editar</a>
                        <form action="/provas/{p["id"]}/deletar" method="post" style="margin:0;" onsubmit="return confirm('Excluir esta prova | tarefa? Se ela tiver aplicações, a exclusão será bloqueada.');">
                            <button type="submit" class="btn" style="padding:4px 10px; font-size:12px; background:var(--red); color:white; border-color:var(--red);">Excluir</button>
                        </form>
                    </div>
                </div>
            </div>
            """
    else:
        cards = '<div class="empty">Nenhuma prova | tarefa encontrada com esses filtros.</div>'

    tem_filtro = bool(disciplina or ano or q)
    subtitle = f'{len(provas)} de {total_geral} prova(s) | tarefa(s)' if tem_filtro else f'{total_geral} prova(s) | tarefa(s) cadastrada(s)'

    content = f"""
        <div class="page-header">
            <h1>Provas | Tarefas</h1>
            <p class="subtitle">{subtitle}</p>
            <div class="page-actions"><a href="/provas/nova" class="btn btn-primary">+ Nova Prova | Tarefa</a></div>
        </div>
        {filtros_html}
        {cards}
    """
    return render_page("Provas | Tarefas", content, active="provas")


def _render_picker_questoes(conn, selected_ids=None):
    """Widget de seleção de questões com filtros, duas colunas e reordenação.
    Usado tanto em criar quanto editar prova. JS serializa IDs em string CSV no campo 'questoes_serializadas'."""
    if selected_ids is None:
        selected_ids = []
    import json

    questoes_db = conn.execute("""
        SELECT q.id, q.enunciado, q.ano, q.tipo, d.nome AS disciplina_nome
        FROM questoes q JOIN disciplinas d ON d.id = q.disciplina_id
        ORDER BY d.nome, q.id
    """).fetchall()

    bncc_map = {}
    for row in conn.execute("""
        SELECT qh.questao_id, h.codigo
        FROM questao_habilidades qh JOIN habilidades_bncc h ON h.id = qh.habilidade_id
        ORDER BY h.codigo
    """).fetchall():
        bncc_map.setdefault(row["questao_id"], []).append(row["codigo"])

    disciplinas = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()

    questoes_payload = [
        {
            "id": q["id"],
            "disciplina": q["disciplina_nome"],
            "ano": q["ano"] if q["ano"] else "",
            "enunciado": q["enunciado"],
            "preview": q["enunciado"][:120] + ("..." if len(q["enunciado"]) > 120 else ""),
            "bnccs": bncc_map.get(q["id"], []),
        }
        for q in questoes_db
    ]
    questoes_json = json.dumps(questoes_payload, ensure_ascii=False)
    selected_json = json.dumps(list(selected_ids))

    disciplinas_opts = '<option value="">Todas</option>' + "".join(
        f'<option value="{d["nome"]}">{d["nome"]}</option>' for d in disciplinas
    )
    anos_opts = '<option value="">Todos</option>' + "".join(f'<option value="{a}">{a}</option>' for a in ANOS)

    template = r'''
<input type="hidden" name="questoes_serializadas" id="questoes_serializadas" value="">

<div style="display:grid; grid-template-columns: 1.4fr 1fr; gap:20px; align-items:flex-start;">
    <div>
        <h3 style="margin-top:0;">Questões disponíveis</h3>
        <div style="background:var(--bg-subtle); padding:12px; border-radius:6px; margin-bottom:12px;">
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">
                <label style="margin:0;">Disciplina<select id="filtro-disciplina">__DISC_OPTS__</select></label>
                <label style="margin:0;">Ano<select id="filtro-ano">__ANOS_OPTS__</select></label>
            </div>
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px; margin-top:8px;">
                <label style="margin:0;">BNCC<input type="text" id="filtro-bncc" placeholder="EF06MA"></label>
                <label style="margin:0;">Buscar<input type="text" id="filtro-q" placeholder="palavra-chave no enunciado"></label>
            </div>
        </div>
        <div id="picker-disponiveis" style="max-height:600px; overflow-y:auto; border:1px solid var(--border); border-radius:6px; padding:8px;"></div>
    </div>
    <div style="position:sticky; top:20px;">
        <h3 style="margin-top:0;">Selecionadas (<span id="picker-counter">0</span>)</h3>
        <div id="picker-selecionadas" style="max-height:600px; overflow-y:auto; border:1px solid var(--border); border-radius:6px; padding:8px;"></div>
    </div>
</div>

<script>
const TODAS_QUESTOES = __QUESTOES_JSON__;
let selecionadas = __SELECTED_JSON__;

function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function renderPicker() {
    const disc = document.getElementById('filtro-disciplina').value;
    const ano = document.getElementById('filtro-ano').value;
    const bncc = document.getElementById('filtro-bncc').value.trim().toUpperCase();
    const q = document.getElementById('filtro-q').value.trim().toLowerCase();

    const filtradas = TODAS_QUESTOES.filter(function(quest) {
        if (disc && quest.disciplina !== disc) return false;
        if (ano && quest.ano !== ano) return false;
        if (bncc && !quest.bnccs.some(function(c){ return c.includes(bncc); })) return false;
        if (q && !quest.enunciado.toLowerCase().includes(q)) return false;
        return true;
    });

    const dispDiv = document.getElementById('picker-disponiveis');
    if (filtradas.length === 0) {
        dispDiv.innerHTML = '<p style="color:var(--text-muted); padding:12px;">Nenhuma questão com esses filtros.</p>';
    } else {
        dispDiv.innerHTML = filtradas.map(function(quest) {
            const isSelected = selecionadas.includes(quest.id);
            const badgeAno = quest.ano ? ' · ' + escapeHtml(quest.ano) : '';
            const badgesBncc = quest.bnccs.map(function(c){ return '<span class="badge" style="font-size:10px;">' + escapeHtml(c) + '</span>'; }).join(' ');
            if (isSelected) {
                return '<div style="padding:8px 10px; margin-bottom:6px; background:var(--bg-muted); border-radius:4px; opacity:0.6;">' +
                    '<div style="font-size:11px; color:var(--text-muted);">Q' + quest.id + ' · ' + escapeHtml(quest.disciplina) + badgeAno + ' ' + badgesBncc + '</div>' +
                    '<div style="font-size:13px; margin-top:4px;">' + escapeHtml(quest.preview) + '</div>' +
                    '<div style="font-size:11px; color:var(--text-muted); margin-top:6px;">✓ Já adicionada</div>' +
                    '</div>';
            }
            return '<div style="padding:8px 10px; margin-bottom:6px; background:var(--bg); border:1px solid var(--border); border-radius:4px; display:flex; gap:8px; align-items:flex-start;">' +
                '<div style="flex:1; min-width:0;">' +
                '<div style="font-size:11px; color:var(--text-muted);">Q' + quest.id + ' · ' + escapeHtml(quest.disciplina) + badgeAno + ' ' + badgesBncc + '</div>' +
                '<div style="font-size:13px; margin-top:4px;">' + escapeHtml(quest.preview) + '</div>' +
                '</div>' +
                '<button type="button" onclick="adicionar(' + quest.id + ')" class="btn" style="padding:4px 10px; font-size:12px; white-space:nowrap;">+ Adicionar</button>' +
                '</div>';
        }).join('');
    }

    const selDiv = document.getElementById('picker-selecionadas');
    document.getElementById('picker-counter').textContent = selecionadas.length;
    if (selecionadas.length === 0) {
        selDiv.innerHTML = '<p style="color:var(--text-muted); padding:12px; font-size:13px;">Nenhuma questão selecionada ainda. Use o painel à esquerda para adicionar.</p>';
    } else {
        const byId = {};
        TODAS_QUESTOES.forEach(function(q){ byId[q.id] = q; });
        selDiv.innerHTML = selecionadas.map(function(qid, idx) {
            const quest = byId[qid];
            if (!quest) return '';
            const badgeAno = quest.ano ? ' · ' + escapeHtml(quest.ano) : '';
            const upDisabled = idx === 0 ? 'disabled style="opacity:0.3;"' : '';
            const downDisabled = idx === selecionadas.length - 1 ? 'disabled style="opacity:0.3;"' : '';
            return '<div style="padding:8px 10px; margin-bottom:6px; background:var(--bg); border:1px solid var(--border); border-radius:4px; display:flex; gap:8px; align-items:flex-start;">' +
                '<div style="flex:0 0 26px; font-weight:600;">' + (idx + 1) + '.</div>' +
                '<div style="flex:1; min-width:0;">' +
                '<div style="font-size:11px; color:var(--text-muted);">Q' + quest.id + ' · ' + escapeHtml(quest.disciplina) + badgeAno + '</div>' +
                '<div style="font-size:12px; margin-top:2px;">' + escapeHtml(quest.preview.slice(0, 80)) + (quest.preview.length > 80 ? '...' : '') + '</div>' +
                '</div>' +
                '<div style="display:flex; flex-direction:column; gap:2px;">' +
                '<button type="button" onclick="mover(' + idx + ', -1)" ' + upDisabled + ' class="btn" style="padding:0 6px; font-size:11px;">▴</button>' +
                '<button type="button" onclick="mover(' + idx + ', 1)" ' + downDisabled + ' class="btn" style="padding:0 6px; font-size:11px;">▾</button>' +
                '</div>' +
                '<button type="button" onclick="remover(' + quest.id + ')" class="btn" style="padding:4px 8px; font-size:11px; color:var(--red);">✕</button>' +
                '</div>';
        }).join('');
    }

    document.getElementById('questoes_serializadas').value = selecionadas.join(',');
}

function adicionar(id) { if (!selecionadas.includes(id)) selecionadas.push(id); renderPicker(); }
function remover(id) { selecionadas = selecionadas.filter(function(x){ return x !== id; }); renderPicker(); }
function mover(idx, delta) {
    const newIdx = idx + delta;
    if (newIdx < 0 || newIdx >= selecionadas.length) return;
    const tmp = selecionadas[idx];
    selecionadas[idx] = selecionadas[newIdx];
    selecionadas[newIdx] = tmp;
    renderPicker();
}

document.getElementById('filtro-disciplina').addEventListener('change', renderPicker);
document.getElementById('filtro-ano').addEventListener('change', renderPicker);
document.getElementById('filtro-bncc').addEventListener('input', renderPicker);
document.getElementById('filtro-q').addEventListener('input', renderPicker);

renderPicker();
</script>
'''
    return (template
        .replace("__QUESTOES_JSON__", questoes_json)
        .replace("__SELECTED_JSON__", selected_json)
        .replace("__DISC_OPTS__", disciplinas_opts)
        .replace("__ANOS_OPTS__", anos_opts))


@app.get("/provas/nova", response_class=HTMLResponse)
def form_nova_prova():
    conn = get_db()
    n_questoes = conn.execute("SELECT COUNT(*) AS c FROM questoes").fetchone()["c"]
    if n_questoes == 0:
        conn.close()
        return render_page("Nova prova", '<div class="page-header"><h1>Nova prova</h1></div><div class="empty"><p>Você precisa cadastrar questões antes de montar uma prova.</p><a href="/questoes/nova" class="btn btn-primary">Cadastrar questão</a></div>', active="provas")
    picker = _render_picker_questoes(conn, selected_ids=[])
    conn.close()
    content = (
        '<div class="page-header"><h1>Nova prova</h1></div>'
        '<form action="/provas/nova" method="post">'
        '<label>Título<input type="text" name="titulo" required placeholder="Ex: Prova de Matemática — 1º Bimestre — 9º Ano"></label>'
        '<label>Descrição (opcional)<textarea name="descricao" rows="2"></textarea></label>'
        f'{picker}'
        '<div class="page-actions"><button type="submit" class="btn btn-primary">Criar prova</button><a href="/provas" class="btn">Cancelar</a></div>'
        '</form>'
    )
    return render_page("Nova prova", content, active="provas", head_extra=MATHJAX)


@app.post("/provas/nova")
def criar_prova(request: Request, titulo: str = Form(...), descricao: str = Form(""), questoes_serializadas: str = Form("")):
    prof = get_current_professor(request)
    if not prof:
        return RedirectResponse("/login", status_code=303)
    ids_str = [x.strip() for x in questoes_serializadas.split(",") if x.strip()]
    questao_ids = []
    for s in ids_str:
        try:
            questao_ids.append(int(s))
        except ValueError:
            pass
    if not questao_ids:
        return RedirectResponse("/provas/nova", status_code=303)
    conn = get_db()
    cursor = conn.execute("INSERT INTO provas (titulo, descricao, criada_por_professor_id) VALUES (?, ?, ?)",
                          (titulo.strip(), descricao.strip() or None, prof["id"]))
    prova_id = cursor.lastrowid
    for ordem, qid in enumerate(questao_ids):
        conn.execute("INSERT INTO prova_questoes (prova_id, questao_id, ordem) VALUES (?, ?, ?)", (prova_id, qid, ordem))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/provas/{prova_id}", status_code=303)


@app.get("/provas/{prova_id}", response_class=HTMLResponse)
def ver_prova(prova_id: int):
    conn = get_db()
    prova = conn.execute("SELECT * FROM provas WHERE id = ?", (prova_id,)).fetchone()
    if not prova:
        conn.close()
        return HTMLResponse(render_page("Não encontrada", '<h1>Prova | tarefa não encontrada</h1><p><a href="/provas">← Voltar</a></p>', active="provas"), status_code=404)
    questoes = conn.execute("SELECT q.id, q.enunciado, q.ano, d.nome AS disciplina_nome FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id JOIN disciplinas d ON d.id = q.disciplina_id WHERE pq.prova_id = ? ORDER BY pq.ordem", (prova_id,)).fetchall()
    n_aplicacoes = conn.execute("SELECT COUNT(*) AS c FROM aplicacoes WHERE prova_id = ?", (prova_id,)).fetchone()["c"]
    questoes_html = "".join(render_questao_card(conn, q, numero=idx) for idx, q in enumerate(questoes, start=1))
    conn.close()
    desc_html = f'<p class="subtitle">{prova["descricao