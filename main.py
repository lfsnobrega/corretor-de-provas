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


def _redimensionar_imagem(data: bytes, max_width: int = 800) -> bytes:
    """Redimensiona imagem para no máximo max_width px de largura, convertendo para JPEG."""
    try:
        from PIL import Image as _PilImage
        import io as _io
        img = _PilImage.open(_io.BytesIO(data))
        if img.mode in ("RGBA", "P"): img = img.convert("RGB")
        w, h = img.size
        if w > max_width:
            img = img.resize((max_width, int(h * max_width / w)), _PilImage.LANCZOS)
        buf = _io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        return buf.getvalue()
    except Exception:
        return data


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
    bot_fracao = f'<button type="button" class="btn-insert-frac" title="Inserir fração" style="{btn_style}">½ fração</button>'

    toolbar_buttons = bot_basicos + sep + bot_fracao + sep + bot_limpar if compact else bot_basicos + sep + bot_extra + bot_fracao + sep + bot_limpar

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
                const btnFrac = toolbar.querySelector('.btn-insert-frac');
                if (btnFrac) {{
                    btnFrac.addEventListener('click', e => {{
                        e.preventDefault();
                        editor.focus();
                        const num = prompt('Numerador da fração:');
                        if (num === null) return;
                        const den = prompt('Denominador da fração:');
                        if (den === null) return;
                        document.execCommand('insertHTML', false, '$\\frac{' + num + '}{' + den + '}$');
                        sync();
                        if (window.MathJax) MathJax.typesetPromise([editor]);
                    }});
                }}
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
    # Migrations multi-prof + gestão
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
    cols_prof = {row[1] for row in conn.execute("PRAGMA table_info(professores)").fetchall()}
    if "is_gestor" not in cols_prof:
        conn.execute("ALTER TABLE professores ADD COLUMN is_gestor INTEGER NOT NULL DEFAULT 0")
    if "status" not in cols_prof:
        conn.execute("ALTER TABLE professores ADD COLUMN status TEXT NOT NULL DEFAULT 'ativo'")
        conn.execute("UPDATE professores SET status = 'ativo' WHERE status IS NULL OR status = ''")
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
        "foto_url": prof["foto_url"], "is_admin": bool(prof["is_admin"]), "is_gestor": bool(prof["is_gestor"] if "is_gestor" in prof.keys() else 0), "status": (prof["status"] if "status" in prof.keys() else "ativo"),
    }


# Rotas públicas (sem login)
PUBLIC_PATHS = {"/login", "/auth/google", "/auth/google/callback", "/auth/dev-login", "/logout", "/acesso-pendente", "/acesso-bloqueado"}
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
    status_prof = prof.get("status", "ativo")
    if status_prof == "pendente" and path != "/acesso-pendente":
        return RedirectResponse("/acesso-pendente", status_code=303)
    if status_prof == "bloqueado" and path != "/acesso-bloqueado":
        return RedirectResponse("/acesso-bloqueado", status_code=303)
    request.state.professor = prof
    token = _current_prof_ctx.set(prof)
    try:
        return await call_next(request)
    finally:
        _current_prof_ctx.reset(token)


def _upsert_professor(email: str, nome: str, foto_url: Optional[str] = None) -> dict:
    """Cria ou atualiza professor. Primeiro = admin ativo. Demais = pendente até aprovação."""
    conn = get_db()
    existing = conn.execute("SELECT * FROM professores WHERE email = ?", (email,)).fetchone()
    if existing:
        conn.execute("UPDATE professores SET nome = ?, foto_url = ?, ultimo_acesso = CURRENT_TIMESTAMP WHERE id = ?",
                     (nome, foto_url, existing["id"]))
        prof_id = existing["id"]
        is_admin = bool(existing["is_admin"])
        is_gestor = bool(existing["is_gestor"] if "is_gestor" in existing.keys() else 0)
        status = existing["status"] if "status" in existing.keys() else "ativo"
    else:
        total = conn.execute("SELECT COUNT(*) AS c FROM professores").fetchone()["c"]
        is_admin_val = 1 if total == 0 else 0
        status_val = "ativo" if is_admin_val == 1 else "pendente"
        c = conn.execute(
            "INSERT INTO professores (email, nome, foto_url, is_admin, is_gestor, status, ultimo_acesso) VALUES (?, ?, ?, ?, 0, ?, CURRENT_TIMESTAMP)",
            (email, nome, foto_url, is_admin_val, status_val)
        )
        prof_id = c.lastrowid
        is_admin = bool(is_admin_val)
        is_gestor = False
        status = status_val
        if is_admin_val == 1:
            conn.execute("UPDATE provas SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
            conn.execute("UPDATE aplicacoes SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
            conn.execute("UPDATE questoes SET criada_por_professor_id = ? WHERE criada_por_professor_id IS NULL", (prof_id,))
    conn.commit()
    conn.close()
    return {"id": prof_id, "email": email, "nome": nome, "is_admin": is_admin, "is_gestor": is_gestor, "status": status}


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
window.MathJax = { tex: { inlineMath: [['$', '$']], displayMath: [['$$', '$$']], processEscapes: true }, svg: { fontCache: 'global' }, options: { skipHtmlTags: ['script','noscript','style','textarea','pre','code'], ignoreHtmlClass: 'editor-content|ed-wrap|editor-toolbar' } };
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
                {nav_item("/painel-gestao", "painel-gestao", "🏛️", "Painel de gestão") if (professor and (professor.get("is_admin") or professor.get("is_gestor"))) else ""}
                {nav_item("/admin/usuarios", "admin-usuarios", "👥", "Usuários") if (professor and professor.get("is_admin")) else ""}
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


def _preview_enunciado(enunciado: str, max_chars: int = 160) -> str:
    """Gera texto limpo para preview: remove tags HTML, entidades e notação MathJax."""
    import re as _re, html as _html
    # Remove blocos MathJax: \(...\) e \[...\] e $...$ e $$...$$
    texto = _re.sub(r'\\\(.*?\\\)', '[fórmula]', enunciado, flags=_re.DOTALL)
    texto = _re.sub(r'\\\[.*?\\\]', '[fórmula]', texto, flags=_re.DOTALL)
    texto = _re.sub(r'\$\$.*?\$\$', '[fórmula]', texto, flags=_re.DOTALL)
    texto = _re.sub(r'\$[^$\n]+\$', '[fórmula]', texto)
    # Remove tags HTML
    texto = _re.sub(r'<[^>]+>', '', texto)
    # Decodifica entidades HTML (&nbsp; → espaço, &amp; → &, etc.)
    texto = _html.unescape(texto)
    # Normaliza espaços
    texto = ' '.join(texto.split())
    return _html.escape(texto[:max_chars]) + ("..." if len(texto) > max_chars else "")


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
        preview = _preview_enunciado(q["enunciado"], max_chars=160)
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

    js_preview = '\n    <script>\n    (function() {\n        var container = document.getElementById(\'bncc-container\');\n        var hiddenInput = document.getElementById(\'bncc-hidden\');\n        var searchInput = document.getElementById(\'bncc-search\');\n        var chipsDiv = document.getElementById(\'bncc-chips\');\n        var resultsDiv = document.getElementById(\'bncc-results\');\n        var discSel = document.querySelector(\'select[name="disciplina_id"]\');\n        if (!container || !hiddenInput || !searchInput) return;\n        var selecionados = [];\n        function renderChips() {\n            chipsDiv.innerHTML = \'\';\n            selecionados.forEach(function(cod) {\n                var chip = document.createElement(\'span\');\n                chip.style.cssText = \'display:inline-flex;align-items:center;gap:4px;background:var(--accent-bg);color:var(--accent);border:1px solid var(--accent-border);border-radius:4px;padding:2px 8px;font-size:12px;font-weight:600;\';\n                chip.innerHTML = cod + \' <button type="button" style="background:none;border:none;cursor:pointer;color:var(--accent);font-size:14px;padding:0;line-height:1;" title="Remover">\\xd7</button>\';\n                chip.querySelector(\'button\').addEventListener(\'click\', function() {\n                    selecionados = selecionados.filter(function(c){return c!==cod;});\n                    renderChips();\n                });\n                chipsDiv.appendChild(chip);\n            });\n            hiddenInput.value = selecionados.join(\', \');\n        }\n        function adicionar(cod) {\n            cod = cod.trim().toUpperCase();\n            if (!cod || selecionados.indexOf(cod) >= 0) return;\n            selecionados.push(cod); renderChips(); resultsDiv.innerHTML = \'\'; searchInput.value = \'\';\n        }\n        function buscar() {\n            var q = searchInput.value.trim();\n            if (q.length < 2) { resultsDiv.innerHTML = \'\'; return; }\n            var disc = discSel ? discSel.value : \'\';\n            var pareceCode = /^[A-Za-z]{2}\\d{2}[A-Za-z]{2}\\d{2}/.test(q);\n            var url = pareceCode ? \'/habilidades/buscar?codigos=\' + encodeURIComponent(q.toUpperCase())\n                : \'/habilidades/buscar?q=\' + encodeURIComponent(q) + (disc ? \'&disciplina_id=\' + disc : \'\');\n            fetch(url).then(function(r){return r.json();}).then(function(data) {\n                var results = [];\n                if (pareceCode) { Object.keys(data).forEach(function(k){if(k!==\'results\') results.push({codigo:k,descricao:data[k]});}); }\n                else { results = data.results || []; }\n                if (results.length === 0) {\n                    if (pareceCode) {\n                        resultsDiv.innerHTML = \'<div style="padding:6px 8px;font-size:12px;color:var(--text-muted);">Código não encontrado. <button type="button" style="background:none;border:none;color:var(--accent);cursor:pointer;font-size:12px;padding:0;text-decoration:underline;">Adicionar mesmo assim</button></div>\';\n                        resultsDiv.querySelector(\'button\').addEventListener(\'click\', function(){adicionar(q);});\n                    } else {\n                        resultsDiv.innerHTML = \'<div style="padding:6px 8px;font-size:12px;color:var(--text-muted);">Nenhum resultado.</div>\';\n                    }\n                    return;\n                }\n                var html = \'<div style="color:var(--text-muted);font-size:11px;padding:4px 2px;">\' + results.length + \' habilidade(s) \\u2014 clique para adicionar:</div>\';\n                results.forEach(function(r) {\n                    html += \'<div data-cod="\' + r.codigo + \'" style="padding:6px 8px;border:1px solid var(--border);border-radius:4px;margin-bottom:3px;cursor:pointer;background:var(--card);font-size:12px;" onmouseover="this.style.background=\\\'var(--accent-bg)\\\'" onmouseout="this.style.background=\\\'var(--card)\\\'"><strong style="color:var(--accent);">\' + r.codigo + \'</strong> \\xb7 \' + (r.descricao||\'\').replace(/</g,\'&lt;\') + \'</div>\';\n                });\n                resultsDiv.innerHTML = html;\n            }).catch(function(){resultsDiv.innerHTML=\'\';});\n        }\n        var _t;\n        searchInput.addEventListener(\'input\', function(){clearTimeout(_t); _t=setTimeout(buscar,350);});\n        searchInput.addEventListener(\'keydown\', function(e){if(e.key===\'Enter\'){e.preventDefault();buscar();}});\n        if (discSel) discSel.addEventListener(\'change\', buscar);\n        resultsDiv.addEventListener(\'click\', function(e){\n            var item = e.target.closest(\'[data-cod]\');\n            if (item) adicionar(item.dataset.cod);\n        });\n        var init = hiddenInput.value.trim();\n        if (init) {\n            init.split(/[,\\n]/).map(function(x){return x.trim().toUpperCase();}).filter(Boolean).forEach(function(c){\n                if(selecionados.indexOf(c)<0) selecionados.push(c);\n            });\n            renderChips();\n        }\n    })();\n    </script>\n'

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
            <div id="bncc-container" style="margin:10px 0;">
                <label style="margin-bottom:6px;">Habilidades BNCC <span style="font-weight:400; color:var(--text-muted); font-size:12px;">(opcional)</span></label>
                <input type="hidden" name="habilidades_codigos" id="bncc-hidden">
                <div id="bncc-chips" style="display:flex; flex-wrap:wrap; gap:6px; min-height:24px; margin-bottom:8px;"></div>
                <input type="search" id="bncc-search" placeholder="Digite o código (EF09MA09) ou palavra-chave (fração, célula...)" style="margin:0;">
                <div id="bncc-results" style="margin-top:6px;"></div>
            </div>
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

            <style>
                .coll-sec{border:1px solid var(--border);border-radius:8px;margin-bottom:12px;overflow:hidden;font-style:normal;}
                .coll-hdr{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;background:var(--bg-subtle);cursor:pointer;user-select:none;font-size:12px;font-weight:600;color:var(--text-muted);letter-spacing:0.05em;text-transform:uppercase;font-style:normal;}
                .coll-hdr:hover{background:var(--border);}
                .coll-arrow{font-size:11px;transition:transform 0.2s;font-style:normal;}
                .coll-body{padding:14px;display:none;font-style:normal;}
                .coll-sec.open .coll-body{display:block;}
                .coll-sec.open .coll-arrow{transform:rotate(180deg);}
            </style>
            <script>function toggleColl(el){{el.closest('.coll-sec').classList.toggle('open');}}</script>

            <div class="coll-sec">
                <div class="coll-hdr" onclick="toggleColl(this)">
                    <span>📝 Textos de apoio (opcionais)</span><span class="coll-arrow">▼</span>
                </div>
                <div class="coll-body">
                    {_editor_enunciado_html(name="texto1_conteudo", valor_inicial="", required=False, label="Texto 1 — conteúdo", min_height=80, placeholder="Cole ou digite aqui o texto de apoio (opcional)")}
                    <label>Texto 1 — fonte<input type="text" name="texto1_fonte" placeholder="Autor, obra, ano"></label>
                    {_editor_enunciado_html(name="texto2_conteudo", valor_inicial="", required=False, label="Texto 2 — conteúdo", min_height=80, placeholder="Segundo texto de apoio (opcional)")}
                    <label>Texto 2 — fonte<input type="text" name="texto2_fonte" placeholder="Autor, obra, ano"></label>
                </div>
            </div>

            <div class="coll-sec">
                <div class="coll-hdr" onclick="toggleColl(this)">
                    <span>🖼️ Imagens (opcionais)</span><span class="coll-arrow">▼</span>
                </div>
                <div class="coll-body">
                    <label>Imagem 1<input type="file" name="imagem1" accept="image/*"></label>
                    <label>Legenda da imagem 1<input type="text" name="imagem1_legenda"></label>
                    <label>Fonte da imagem 1<input type="text" name="imagem1_fonte"></label>
                    <label>Imagem 2<input type="file" name="imagem2" accept="image/*"></label>
                    <label>Legenda da imagem 2<input type="text" name="imagem2_legenda"></label>
                    <label>Fonte da imagem 2<input type="text" name="imagem2_fonte"></label>
                </div>
            </div>

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
            content_bytes = await img.read()
            content_bytes = _redimensionar_imagem(content_bytes, max_width=800)
            unique_name = f"{uuid.uuid4().hex}.jpg"
            file_path = os.path.join(UPLOAD_DIR, unique_name)
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
            "preview": _preview_enunciado(q["enunciado"], max_chars=120),
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
    desc_html = f'<p class="subtitle">{prova["descricao"]}</p>' if prova["descricao"] else ""
    comparativo_btn = ""
    if n_aplicacoes > 0:
        label = "Comparativo entre turmas" if n_aplicacoes >= 2 else "Ver análises pedagógicas"
        comparativo_btn = f'<a href="/provas/{prova_id}/comparativo" class="btn">📊 {label} ({n_aplicacoes} aplicação{"" if n_aplicacoes == 1 else "ões"})</a>'
    prof_ctx = _current_prof_ctx.get()
    status_rev = prova["status_revisao"] if "status_revisao" in prova.keys() else "rascunho"
    eh_dono = prof_ctx and (prova["criada_por_professor_id"] == prof_ctx["id"] or prof_ctx.get("is_admin"))
    eh_gestor_ou_admin = prof_ctx and (prof_ctx.get("is_admin") or prof_ctx.get("is_gestor"))
    status_badge_html = _status_badge_html(status_rev)
    obs_html = ""
    if "obs_gestao" in prova.keys() and prova["obs_gestao"]:
        obs_html = f'<div style="background:var(--orange-bg); border-left:3px solid var(--orange); padding:8px 12px; border-radius:6px; margin-top:8px; font-size:13px;"><strong>Obs. da gestão:</strong> {prova["obs_gestao"]}</div>'
    submeter_btn = ""
    if eh_dono and status_rev in ("rascunho", "devolvida"):
        submeter_btn = f'<form method="post" action="/provas/{prova_id}/submeter" style="margin:0;" onsubmit="return confirm(\'Submeter para revisão da gestão?\')"><button type="submit" class="btn btn-primary" style="background:var(--orange); border-color:var(--orange);">📤 Submeter para revisão</button></form>'
    imprimir_btn = f'<a href="/provas/{prova_id}/imprimir" class="btn btn-primary" target="_blank">🖨️ Imprimir prova</a>' if (status_rev == "aprovada" or eh_gestor_ou_admin) else ""
    acoes_html = f'<div class="page-actions" style="display:flex; gap:8px; flex-wrap:wrap; align-items:center;">{imprimir_btn}{comparativo_btn}{submeter_btn}{status_badge_html}{obs_html}</div>'
    content = f'<div class="page-header"><h1>{prova["titulo"]}</h1><p class="subtitle">{len(questoes)} questões</p>{desc_html}{acoes_html}</div><hr>{questoes_html}'
    return render_page(prova["titulo"], content, active="provas", head_extra=MATHJAX)


@app.get("/provas/{id}/editar", response_class=HTMLResponse)
def form_editar_prova(id: int):
    conn = get_db()
    prova = conn.execute("SELECT * FROM provas WHERE id = ?", (id,)).fetchone()
    if not prova:
        conn.close()
        return RedirectResponse("/provas", status_code=303)
    
    todas_questoes = conn.execute("""
        SELECT q.id, q.enunciado, q.ano, q.tipo, d.nome AS disciplina_nome 
        FROM questoes q 
        JOIN disciplinas d ON d.id = q.disciplina_id 
        ORDER BY d.nome, q.id
    """).fetchall()
    
    selecionadas = [r["questao_id"] for r in conn.execute("SELECT questao_id FROM prova_questoes WHERE prova_id = ?", (id,)).fetchall()]
    conn.close()
    
    questoes_html = ""
    for q in todas_questoes:
        checked = " checked" if q["id"] in selecionadas else ""
        resumo_enunciado = q["enunciado"][:110] + "..." if len(q["enunciado"]) > 110 else q["enunciado"]
        questoes_html += f"""
        <div style="margin-bottom:10px; display:flex; align-items:flex-start; gap:10px;">
            <input type="checkbox" name="questoes_ids" value="{q["id"]}"{checked} id="q_{q["id"]}" style="width:auto; margin-top:4px;">
            <label for="q_{q["id"]}" style="font-weight:normal; margin:0; cursor:pointer;">
                <span class="badge" style="margin-right:4px;">{q["disciplina_nome"]}</span> 
                <strong>(ID #{q["id"]})</strong> - {resumo_enunciado}
            </label>
        </div>
        """

    content = f"""
    <div class="page-header"><h1>Editar Prova: {prova["titulo"]}</h1></div>
    <form action="/provas/{id}/editar" method="post">
        <label>Título da Prova
            <input type="text" name="titulo" value="{prova["titulo"]}" required>
        </label>
        
        <fieldset style="margin-top:20px;">
            <legend>Selecione as Questões Integrantes</legend>
            {questoes_html if questoes_html else '<p class="empty">Nenhuma questão encontrada para vincular.</p>'}
        </fieldset>
        
        <div class="page-actions" style="margin-top:20px;">
            <button type="submit" class="btn btn-primary">Salvar Alterações</button>
            <a href="/provas" class="btn">Cancelar</a>
        </div>
    </form>
    """
    return render_page("Editar Prova", content, active="provas")


@app.post("/provas/{id}/editar")
def atualizar_prova(id: int, titulo: str = Form(...), questoes_ids: List[int] = Form([])):
    conn = get_db()
    conn.execute("UPDATE provas SET titulo = ? WHERE id = ?", (titulo.strip(), id))
    conn.execute("DELETE FROM prova_questoes WHERE prova_id = ?", (id,))
    
    for idx, q_id in enumerate(questoes_ids):
        conn.execute("INSERT INTO prova_questoes (prova_id, questao_id, ordem) VALUES (?, ?, ?)", (id, q_id, idx))
        
    conn.commit()
    conn.close()
    return RedirectResponse("/provas", status_code=303)


@app.post("/provas/{id}/deletar", response_class=HTMLResponse)
def deletar_prova(id: int):
    conn = get_db()
    uso_ativo = conn.execute("SELECT id FROM aplicacoes WHERE prova_id = ?", (id,)).fetchone()
    if uso_ativo:
        conn.close()
        content = """
        <div style="border: 1px solid var(--red); background: var(--red-bg); padding: 20px; border-radius: 6px; margin-top:20px; color:var(--red);">
            <h3 style="color:var(--red); margin-top:0;">Operação Impedida</h3>
            <p>Não é possível deletar esta prova | tarefa porque ela possui <strong>Aplicações</strong> em andamento ou histórico de notas associado a turmas.</p>
            <p>Se deseja realmente excluí-la, remova primeiro as respectivas aplicações na aba de "Aplicações".</p>
            <a href="/provas" class="btn" style="margin-top:10px;">Voltar para Provas</a>
        </div>
        """
        return render_page("Erro ao Deletar Prova", content, active="provas")
        
    conn.execute("DELETE FROM prova_questoes WHERE prova_id = ?", (id,))
    conn.execute("DELETE FROM provas WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/provas", status_code=303)


# ==========================================
#  ROTAS DE TURMAS
# ==========================================

@app.get("/turmas", response_class=HTMLResponse)
def listar_turmas(request: Request):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    prof = get_current_professor(request)
    conn = get_db()
    turmas = conn.execute("SELECT t.id, t.nome, t.ano_letivo, COUNT(a.id) AS total_alunos FROM turmas t LEFT JOIN alunos a ON a.turma_id = t.id GROUP BY t.id ORDER BY t.ano_letivo DESC, t.nome").fetchall()
    conn.close()
    if turmas:
        cards = "".join(f'<a href="/turmas/{t["id"]}" class="card card-link"><div class="card-title">{t["nome"]}</div><div class="card-meta">Ano letivo {t["ano_letivo"]} · {t["total_alunos"]} alunos</div></a>' for t in turmas)
    else:
        cards = '<div class="empty">Nenhuma turma cadastrada ainda.</div>'
    botoes_admin = (
        '<div class="page-actions"><a href="/turmas/nova" class="btn btn-primary">+ Nova turma</a><a href="/turmas/importar" class="btn">Importar planilha</a></div>'
        if prof and prof["is_admin"] else
        '<p class="muted-line" style="font-size:13px;">As turmas são gerenciadas pelo administrador da escola.</p>'
    )
    content = f'<div class="page-header"><h1>Turmas</h1>{botoes_admin}</div>{cards}'
    return render_page("Turmas", content, active="turmas")


@app.get("/turmas/nova", response_class=HTMLResponse)
def form_nova_turma(request: Request):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    content = '<div class="page-header"><h1>Nova turma</h1></div><form action="/turmas/nova" method="post"><label>Nome<input type="text" name="nome" required placeholder="Ex: 9º Ano A" autofocus></label><label>Ano letivo<input type="number" name="ano_letivo" required value="2026" min="2020" max="2099"></label><div class="page-actions"><button type="submit" class="btn btn-primary">Cadastrar</button><a href="/turmas" class="btn">Cancelar</a></div></form>'
    return render_page("Nova turma", content, active="turmas")


@app.post("/turmas/nova")
def criar_turma(request: Request, nome: str = Form(...), ano_letivo: int = Form(...)):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    cursor = conn.execute("INSERT INTO turmas (nome, ano_letivo) VALUES (?, ?)", (nome.strip(), ano_letivo))
    turma_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return RedirectResponse(f"/turmas/{turma_id}", status_code=303)

@app.get("/turmas/template")
def baixar_template_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Alunos"
    headers = ["Turma", "Ano Letivo", "Número", "Nome", "Raça", "E-mail", "Data de Nascimento"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.append(["9º Ano A", 2026, 1, "Maria da Silva", "Parda", "maria@escola.com", "2010-03-15"])
    ws.append(["9º Ano A", 2026, 2, "João Santos", "Branca", "joao@escola.com", "2010-07-22"])
    ws.append(["9º Ano B", 2026, 1, "Ana Pereira", "", "", ""])

    larguras = [16, 12, 10, 28, 12, 26, 22]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=template_alunos.xlsx"},
    )


@app.get("/turmas/importar", response_class=HTMLResponse)
def form_importar_excel(request: Request):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    content = """
        <div class="page-header">
            <h1>Importar planilha</h1>
            <p class="subtitle">Cadastre várias turmas e alunos de uma vez subindo um arquivo Excel.</p>
        </div>
        <div class="tip">
            <strong>Como funciona:</strong> a planilha deve ter as colunas <code>Turma</code>, <code>Ano Letivo</code>, <code>Número</code>, <code>Nome</code>, <code>Raça</code>, <code>E-mail</code>, e <code>Data de Nascimento</code> (na primeira linha como cabeçalho). Cada linha seguinte é um aluno. Turmas que ainda não existem são criadas automaticamente. Alunos com nome já cadastrado na mesma turma são pulados (evita duplicação se você importar a planilha duas vezes).
        </div>

        <h2>1. Baixar template</h2>
        <p>Se você ainda não tem a planilha, baixa um modelo pronto com a estrutura certa:</p>
        <p><a href="/turmas/template" class="btn">Baixar template Excel</a></p>

        <h2>2. Subir planilha preenchida</h2>
        <form action="/turmas/importar" method="post" enctype="multipart/form-data">
            <label>
                Arquivo .xlsx
                <input type="file" name="arquivo" accept=".xlsx" required>
            </label>
            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Importar</button>
                <a href="/turmas" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return render_page("Importar planilha", content, active="turmas")


@app.post("/turmas/importar", response_class=HTMLResponse)
async def importar_excel(request: Request, arquivo: UploadFile = File(...)):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    if not arquivo.filename.lower().endswith(".xlsx"):
        content = """
            <div class="page-header"><h1>Erro na importação</h1></div>
            <div class="tip">O arquivo precisa ser .xlsx (Excel moderno).</div>
            <p><a href="/turmas/importar" class="btn">Voltar</a></p>
        """
        return HTMLResponse(render_page("Erro", content, active="turmas"))

    content_bytes = await arquivo.read()
    try:
        wb = load_workbook(BytesIO(content_bytes), read_only=True, data_only=True)
        ws = wb.active
    except Exception as e:
        content = f"""
            <div class="page-header"><h1>Erro ao ler a planilha</h1></div>
            <div class="tip">{str(e)}</div>
            <p><a href="/turmas/importar" class="btn">Voltar</a></p>
        """
        return HTMLResponse(render_page("Erro", content, active="turmas"))

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        content = """
            <div class="page-header"><h1>Planilha vazia</h1></div>
            <p>A planilha não contém dados além do cabeçalho.</p>
            <p><a href="/turmas/importar" class="btn">Voltar</a></p>
        """
        return HTMLResponse(render_page("Vazia", content, active="turmas"))

    conn = get_db()
    turmas_criadas = 0
    alunos_criados = 0
    alunos_pulados = 0
    avisos = []

    for row_num, row in enumerate(rows[1:], start=2):
        if not row or all(c is None or (isinstance(c, str) and not c.strip()) for c in row):
            continue

        row = list(row) + [None] * (7 - len(row))
        turma_nome, ano_letivo_raw, numero_raw, nome, raca, email, data_nasc_raw = row[:7]

        if not turma_nome or not str(turma_nome).strip():
            avisos.append(f"Linha {row_num}: turma vazia, ignorada")
            continue
        if not nome or not str(nome).strip():
            avisos.append(f"Linha {row_num}: nome vazio, ignorada")
            continue

        try:
            ano_letivo = int(ano_letivo_raw) if ano_letivo_raw else 2026
        except (TypeError, ValueError):
            avisos.append(f"Linha {row_num}: ano letivo inválido ({ano_letivo_raw}), ignorada")
            continue

        turma_nome_clean = str(turma_nome).strip()
        nome_clean = str(nome).strip()

        turma = conn.execute(
            "SELECT id FROM turmas WHERE nome = ? AND ano_letivo = ?",
            (turma_nome_clean, ano_letivo),
        ).fetchone()

        if turma:
            turma_id = turma["id"]
        else:
            cursor = conn.execute(
                "INSERT INTO turmas (nome, ano_letivo) VALUES (?, ?)",
                (turma_nome_clean, ano_letivo),
            )
            turma_id = cursor.lastrowid
            turmas_criadas += 1

        existing = conn.execute(
            "SELECT id FROM alunos WHERE turma_id = ? AND LOWER(nome) = LOWER(?)",
            (turma_id, nome_clean),
        ).fetchone()
        if existing:
            alunos_pulados += 1
            continue

        try:
            numero = int(numero_raw) if numero_raw else None
        except (TypeError, ValueError):
            numero = None

        raca_clean = str(raca).strip() if raca else None
        email_clean = str(email).strip() if email else None

        data_nasc_str = None
        if data_nasc_raw:
            if isinstance(data_nasc_raw, (date, datetime)):
                data_nasc_str = data_nasc_raw.isoformat()[:10]
            else:
                s = str(data_nasc_raw).strip()
                for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]:
                    try:
                        data_nasc_str = datetime.strptime(s, fmt).date().isoformat()
                        break
                    except ValueError:
                        continue
                if not data_nasc_str:
                    avisos.append(f"Linha {row_num}: data '{s}' não reconhecida, gravado vazio")

        codigo = gerar_codigo_aluno(conn)
        conn.execute(
            "INSERT INTO alunos (turma_id, nome, numero, codigo_unico, raca, email, data_nascimento) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (turma_id, nome_clean, numero, codigo, raca_clean, email_clean, data_nasc_str),
        )
        alunos_criados += 1

    conn.commit()
    conn.close()

    avisos_html = ""
    if avisos:
        items = "".join(f"<li>{a}</li>" for a in avisos)
        avisos_html = f'<h2>Avisos</h2><ul class="clean">{items}</ul>'

    content = f"""
        <div class="page-header">
            <h1>Importação concluída</h1>
            <p class="subtitle">Resumo do que foi processado.</p>
        </div>
        <div class="metric-grid">
            <div class="metric"><div class="metric-label">Turmas criadas</div><div class="metric-value">{turmas_criadas}</div></div>
            <div class="metric"><div class="metric-label">Alunos criados</div><div class="metric-value">{alunos_criados}</div></div>
            <div class="metric"><div class="metric-label">Alunos pulados</div><div class="metric-value">{alunos_pulados}</div></div>
            <div class="metric"><div class="metric-label">Avisos</div><div class="metric-value">{len(avisos)}</div></div>
        </div>
        {avisos_html}
        <div class="page-actions">
            <a href="/turmas" class="btn btn-primary">Ver turmas</a>
            <a href="/turmas/importar" class="btn">Importar outra</a>
        </div>
    """
    return HTMLResponse(render_page("Importação concluída", content, active="turmas"))

@app.get("/turmas/{turma_id}", response_class=HTMLResponse)
def ver_turma(request: Request, turma_id: int):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    prof = get_current_professor(request)
    is_admin = prof and prof["is_admin"]
    conn = get_db()
    turma = conn.execute("SELECT * FROM turmas WHERE id = ?", (turma_id,)).fetchone()
    if not turma:
        conn.close()
        return HTMLResponse(render_page("Não encontrada", '<h1>Turma não encontrada</h1><p><a href="/turmas">← Voltar</a></p>', active="turmas"), status_code=404)
    alunos = conn.execute("SELECT * FROM alunos WHERE turma_id = ? ORDER BY numero, nome", (turma_id,)).fetchall()
    proximo_numero = conn.execute("SELECT COALESCE(MAX(numero), 0) + 1 AS n FROM alunos WHERE turma_id = ?", (turma_id,)).fetchone()["n"]
    conn.close()

    if alunos:
        alunos_html = ""
        for a in alunos:
            num = a["numero"] if a["numero"] else "—"
            extras = []
            if a["raca"]: extras.append(a["raca"])
            if a["email"]: extras.append(a["email"])
            if a["data_nascimento"]: extras.append(f'nasc. {format_data_br(a["data_nascimento"])}')
            extra_line = f'<div style="font-size:12px; color:var(--text-muted); margin-top:2px;">{" · ".join(extras)}</div>' if extras else ""
            if is_admin:
                nome_escapado = a["nome"].replace("'", "\\'")
                acoes = (
                    f'<div style="font-size:11px; margin-top:6px;">'
                    f'<a href="/alunos/{a["id"]}/editar" style="color:var(--text-muted);">Editar</a>'
                    f'<span style="color:var(--text-subtle);"> · </span>'
                    f'<a href="/alunos/{a["id"]}/transferir" style="color:var(--text-muted);">Transferir</a>'
                    f'<span style="color:var(--text-subtle);"> · </span>'
                    f'<form action="/alunos/{a["id"]}/deletar" method="post" style="display:inline; margin:0;" '
                    f"onsubmit=\"return confirm('Excluir {nome_escapado}? Se o aluno tiver entregas registradas, você poderá forçar a exclusão na próxima tela.');\">"
                    f'<button type="submit" style="background:none; border:none; padding:0; color:var(--red); cursor:pointer; font-size:inherit; font-family:inherit;">Excluir</button>'
                    f'</form>'
                    f'</div>'
                )
            else:
                acoes = ""
            alunos_html += f'<div class="student-row"><div class="numero">{num}</div><div>{a["nome"]}{extra_line}{acoes}</div><div class="codigo">{a["codigo_unico"]}</div></div>'
    else:
        alunos_html = '<div class="empty">Nenhum aluno cadastrado nesta turma ainda.</div>'

    racas_options = '<option value="">Não informada</option>' + "".join(f'<option value="{r}">{r}</option>' for r in RACAS)

    if is_admin:
        excluir_turma_btn = (
            f'<div class="page-actions"><form action="/turmas/{turma_id}/deletar" method="post" style="margin:0;" '
            f"onsubmit=\"return confirm('Excluir esta turma?\\n\\nIsso removerá: alunos, aplicações desta turma, respostas e entregas associadas.') && "
            f"confirm('TEM CERTEZA? Esta ação é IRREVERSÍVEL e não pode ser desfeita.');\">"
            f'<button type="submit" class="btn" style="background:var(--red); color:white; border-color:var(--red);">🗑️ Excluir turma</button>'
            f'</form></div>'
        )
        form_adicionar = f"""
            <h2>Adicionar aluno</h2>
            <form action="/turmas/{turma_id}/alunos" method="post">
                <div style="display:grid; grid-template-columns: 100px 1fr; gap:12px;">
                    <label>Número<input type="number" name="numero" value="{proximo_numero}" min="1"></label>
                    <label>Nome<input type="text" name="nome" required placeholder="Nome completo"></label>
                </div>
                <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:12px;">
                    <label>Raça<select name="raca">{racas_options}</select></label>
                    <label>E-mail<input type="email" name="email" placeholder="aluno@email.com"></label>
                    <label>Data de nascimento<input type="date" name="data_nascimento"></label>
                </div>
                <div class="page-actions">
                    <button type="submit" class="btn btn-primary">Adicionar</button>
                </div>
            </form>
        """
    else:
        excluir_turma_btn = ""
        form_adicionar = '<p class="muted-line" style="font-size:13px; margin-top:18px;">Apenas o administrador pode adicionar/editar/excluir alunos.</p>'

    content = f"""
        <div class="page-header">
            <h1>{turma["nome"]}</h1>
            <p class="subtitle">Ano letivo {turma["ano_letivo"]} · {len(alunos)} alunos</p>
            {excluir_turma_btn}
        </div>

        <h2>Alunos</h2>
        {alunos_html}

        {form_adicionar}
    """
    return render_page(f"Turma {turma['nome']}", content, active="turmas")


@app.post("/turmas/{turma_id}/alunos")
def adicionar_aluno(request: Request, 
    turma_id: int,
    nome: str = Form(...),
    numero: Optional[int] = Form(None),
    raca: str = Form(""),
    email: str = Form(""),
    data_nascimento: str = Form(""),
):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    codigo = gerar_codigo_aluno(conn)
    conn.execute(
        "INSERT INTO alunos (turma_id, nome, numero, codigo_unico, raca, email, data_nascimento) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (turma_id, nome.strip(), numero, codigo, raca.strip() or None, email.strip() or None, data_nascimento.strip() or None)
    )
    conn.commit()
    conn.close()
    return RedirectResponse(f"/turmas/{turma_id}", status_code=303)


# ==========================================
#  ROTAS DE APLICAÇÕES (ATUALIZADAS TAREFA A2)
# ==========================================


# ==========================================
#  ACESSO PENDENTE / BLOQUEADO
# ==========================================

@app.get("/acesso-pendente", response_class=HTMLResponse)
def acesso_pendente(request: Request):
    prof = get_current_professor(request)
    nome = prof["nome"] if prof else "Professor"
    body = f"""
        <div style="max-width:480px; margin:80px auto; text-align:center; padding:0 20px;">
            <div style="font-size:56px; margin-bottom:16px;">⏳</div>
            <h1 style="font-size:22px; margin-bottom:8px;">Acesso aguardando aprovação</h1>
            <p style="color:var(--text-muted); margin-bottom:24px;">
                Olá, <strong>{nome}</strong>! Seu cadastro foi recebido e está aguardando aprovação da gestão escolar.
            </p>
            <a href="/logout" class="btn" style="margin-top:24px;">Sair</a>
        </div>"""
    return HTMLResponse(render_page("Acesso pendente", body, active=""))


@app.get("/acesso-bloqueado", response_class=HTMLResponse)
def acesso_bloqueado(request: Request):
    prof = get_current_professor(request)
    nome = prof["nome"] if prof else "Professor"
    body = f"""
        <div style="max-width:480px; margin:80px auto; text-align:center; padding:0 20px;">
            <div style="font-size:56px; margin-bottom:16px;">🚫</div>
            <h1 style="font-size:22px; margin-bottom:8px;">Acesso bloqueado</h1>
            <p style="color:var(--text-muted);">Olá, <strong>{nome}</strong>. Seu acesso foi bloqueado. Entre em contato com o administrador.</p>
            <a href="/logout" class="btn" style="margin-top:24px;">Sair</a>
        </div>"""
    return HTMLResponse(render_page("Acesso bloqueado", body, active=""))


# ==========================================
#  PAINEL DE GESTÃO DE PROVAS
# ==========================================

STATUS_REVISAO_LABEL = {
    "rascunho":   ("✏️", "Rascunho",  "var(--text-muted)", "var(--bg-subtle)"),
    "submetida":  ("📤", "Submetida", "var(--orange)",     "var(--orange-bg)"),
    "aprovada":   ("✅", "Aprovada",  "var(--green)",      "var(--green-bg)"),
    "devolvida":  ("↩️", "Devolvida", "var(--red)",        "var(--red-bg)"),
}

def _status_badge_html(status: str) -> str:
    icon, label, color, bg = STATUS_REVISAO_LABEL.get(status, ("❓", status, "var(--text-muted)", "var(--bg-subtle)"))
    return f'<span style="background:{bg}; color:{color}; border-radius:6px; padding:2px 10px; font-size:12px; font-weight:600;">{icon} {label}</span>'


@app.post("/provas/{prova_id}/submeter")
def submeter_prova(prova_id: int):
    prof = _current_prof_ctx.get()
    if not prof:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    prova = conn.execute("SELECT * FROM provas WHERE id = ?", (prova_id,)).fetchone()
    if prova and (prova["criada_por_professor_id"] == prof["id"] or prof.get("is_admin")):
        status_rev = prova["status_revisao"] if "status_revisao" in prova.keys() else "rascunho"
        if status_rev in ("rascunho", "devolvida"):
            conn.execute("UPDATE provas SET status_revisao = 'submetida', obs_gestao = NULL WHERE id = ?", (prova_id,))
            conn.commit()
    conn.close()
    return RedirectResponse(f"/provas/{prova_id}", status_code=303)


@app.post("/provas/{prova_id}/aprovar")
def aprovar_prova(prova_id: int, obs: str = Form("")):
    prof = _current_prof_ctx.get()
    if not prof or not (prof.get("is_admin") or prof.get("is_gestor")):
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    conn.execute("UPDATE provas SET status_revisao = 'aprovada', obs_gestao = ?, revisado_por_id = ?, revisado_em = CURRENT_TIMESTAMP WHERE id = ?",
        (obs.strip() or None, prof["id"], prova_id))
    conn.commit(); conn.close()
    return RedirectResponse("/painel-gestao", status_code=303)


@app.post("/provas/{prova_id}/devolver")
def devolver_prova(prova_id: int, obs: str = Form(...)):
    prof = _current_prof_ctx.get()
    if not prof or not (prof.get("is_admin") or prof.get("is_gestor")):
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    conn.execute("UPDATE provas SET status_revisao = 'devolvida', obs_gestao = ?, revisado_por_id = ?, revisado_em = CURRENT_TIMESTAMP WHERE id = ?",
        (obs.strip(), prof["id"], prova_id))
    conn.commit(); conn.close()
    return RedirectResponse("/painel-gestao", status_code=303)


@app.get("/painel-gestao", response_class=HTMLResponse)
def painel_gestao(request: Request, status: Optional[str] = "submetida", prof_id: Optional[int] = None):
    prof = _current_prof_ctx.get()
    if not prof or not (prof.get("is_admin") or prof.get("is_gestor")):
        return HTMLResponse(render_page("Acesso negado", '<div class="empty">Sem permissão.</div>', active="painel-gestao"), status_code=403)
    conn = get_db()
    where = []; params = []
    status_atual = status or "submetida"
    if status_atual != "todas":
        where.append("p.status_revisao = ?"); params.append(status_atual)
    if prof_id:
        where.append("p.criada_por_professor_id = ?"); params.append(prof_id)
    wc = ("WHERE " + " AND ".join(where)) if where else ""
    provas = conn.execute(f"""
        SELECT p.id, p.titulo, p.status_revisao, p.obs_gestao, p.criada_em, p.revisado_em,
               pr.nome AS criador_nome, rv.nome AS revisor_nome,
               (SELECT COUNT(*) FROM prova_questoes WHERE prova_id = p.id) AS n_questoes
        FROM provas p
        LEFT JOIN professores pr ON pr.id = p.criada_por_professor_id
        LEFT JOIN professores rv ON rv.id = p.revisado_por_id
        {wc} ORDER BY CASE p.status_revisao WHEN 'submetida' THEN 0 WHEN 'devolvida' THEN 1 WHEN 'aprovada' THEN 2 ELSE 3 END, p.id DESC
    """, params).fetchall()
    professores_lista = conn.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall()
    contadores = {r["status_revisao"]: r["c"] for r in conn.execute("SELECT status_revisao, COUNT(*) AS c FROM provas GROUP BY status_revisao").fetchall()}
    conn.close()

    def _cnt(s): return contadores.get(s, 0)
    tabs_data = [
        ("submetida", "📤 Aguardando (" + str(_cnt("submetida")) + ")", "var(--orange)"),
        ("devolvida",  "↩️ Devolvidas (" + str(_cnt("devolvida")) + ")", "var(--red)"),
        ("aprovada",   "✅ Aprovadas (" + str(_cnt("aprovada")) + ")", "var(--green)"),
        ("rascunho",   "✏️ Rascunhos (" + str(_cnt("rascunho")) + ")", "var(--text-muted)"),
        ("todas",      "📋 Todas", "var(--accent)"),
    ]
    tabs_html = '<div style="display:flex; gap:0; border-bottom:2px solid var(--border); margin-bottom:18px; flex-wrap:wrap;">'
    for key, label, color in tabs_data:
        ativo = status_atual == key
        tabs_html += f'<a href="/painel-gestao?status={key}" style="padding:9px 16px; font-size:13px; font-weight:600; text-decoration:none; border-bottom:3px solid {"var(--accent)" if ativo else "transparent"}; color:{"var(--accent)" if ativo else "var(--text-muted)"}; margin-bottom:-2px; white-space:nowrap;">{label}</a>'
    tabs_html += '</div>'

    prof_opts = '<option value="">Todos os professores</option>' + "".join(f'<option value="{p["id"]}"{" selected" if prof_id == p["id"] else ""}>{p["nome"]}</option>' for p in professores_lista)
    filtros_html = f"""<form method="get" action="/painel-gestao" style="background:var(--bg-subtle); padding:10px 16px; border-radius:8px; margin-bottom:18px;">
        <input type="hidden" name="status" value="{status_atual}">
        <div style="display:flex; gap:10px; align-items:flex-end; flex-wrap:wrap;">
            <label style="margin:0;">Professor<select name="prof_id">{prof_opts}</select></label>
            <button type="submit" class="btn btn-primary" style="margin:0;">Filtrar</button>
            <a href="/painel-gestao?status={status_atual}" class="btn" style="margin:0;">Limpar</a>
        </div></form>"""

    cards_html = ""
    for p in provas:
        badge = _status_badge_html(p["status_revisao"])
        obs_html = f'<div style="margin-top:6px; background:var(--orange-bg); border-left:3px solid var(--orange); padding:6px 10px; border-radius:4px; font-size:12px;"><strong>Obs:</strong> {p["obs_gestao"]}</div>' if p["obs_gestao"] else ""
        revisor = f'<span style="font-size:11px; color:var(--text-muted);">Revisado por {p["revisor_nome"]} em {(p["revisado_em"] or "")[:10]}</span>' if p["revisor_nome"] else ""
        acoes = f'<a href="/provas/{p["id"]}" class="btn" style="padding:5px 10px; font-size:12px;">👁️ Ver</a><a href="/provas/{p["id"]}/imprimir" class="btn" style="padding:5px 10px; font-size:12px;" target="_blank">🖨️ PDF</a>'
        if p["status_revisao"] == "submetida":
            acoes += f'''<form method="post" action="/provas/{p["id"]}/aprovar" style="margin:0; display:inline-flex; gap:4px; align-items:center;">
                <input type="text" name="obs" placeholder="Obs. opcional..." style="width:160px; padding:4px 8px; font-size:12px; margin:0;">
                <button type="submit" class="btn" style="padding:5px 10px; font-size:12px; color:var(--green); border-color:var(--green);">✅ Aprovar</button>
            </form>
            <form method="post" action="/provas/{p["id"]}/devolver" style="margin:0; display:inline-flex; gap:4px; align-items:center;">
                <input type="text" name="obs" placeholder="Motivo..." required style="width:160px; padding:4px 8px; font-size:12px; margin:0;">
                <button type="submit" class="btn" style="padding:5px 10px; font-size:12px; color:var(--red); border-color:var(--red);">↩️ Devolver</button>
            </form>'''
        elif p["status_revisao"] == "aprovada":
            acoes += f'<a href="/provas/{p["id"]}/editar" class="btn" style="padding:5px 10px; font-size:12px;">✏️ Editar</a>'
        cards_html += f"""<div style="border:1px solid var(--border); border-radius:10px; padding:14px 18px; margin-bottom:12px; background:var(--card);">
            <div style="display:flex; align-items:flex-start; justify-content:space-between; gap:12px; flex-wrap:wrap;">
                <div style="flex:1; min-width:200px;">
                    <div style="font-weight:600; font-size:14px;">{p["titulo"]}</div>
                    <div style="font-size:12px; color:var(--text-muted); margin-top:3px;">{p["n_questoes"]} questões · por {p["criador_nome"] or "—"} · {(p["criada_em"] or "")[:10]}</div>
                    {obs_html}<div style="margin-top:4px;">{revisor}</div>
                </div>
                <div style="display:flex; flex-direction:column; gap:6px; align-items:flex-end;">
                    {badge}
                    <div style="display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end;">{acoes}</div>
                </div>
            </div></div>"""
    if not cards_html:
        cards_html = '<div class="empty">Nenhuma prova encontrada.</div>'

    content_html = f"""<div class="page-header"><h1>🏛️ Painel de gestão</h1><p class="subtitle">Revisão e aprovação de provas para impressão</p></div>
        {tabs_html}{filtros_html}{cards_html}"""
    return render_page("Painel de gestão", content_html, active="painel-gestao")


# ==========================================
#  MINHAS APLICAÇÕES
# ==========================================

@app.get("/minhas-aplicacoes", response_class=HTMLResponse)
def minhas_aplicacoes(request: Request, aba: Optional[str] = "abertas", turma: Optional[str] = None, prof_id: Optional[int] = None):
    prof = _current_prof_ctx.get()
    if not prof:
        return RedirectResponse("/login", status_code=303)
    is_admin = bool(prof.get("is_admin"))
    conn = get_db()
    turma_id = int(turma) if turma and turma.isdigit() else None
    aberta_val = 1 if aba == "abertas" else 0
    where = ["a.aberta = ?"]; params_q = [aberta_val]
    if not is_admin:
        where.append("(a.criada_por_professor_id = ? OR a.criada_por_professor_id IS NULL)"); params_q.append(prof["id"])
    elif prof_id:
        where.append("a.criada_por_professor_id = ?"); params_q.append(prof_id)
    if turma_id:
        where.append("a.turma_id = ?"); params_q.append(turma_id)
    wc = "WHERE " + " AND ".join(where)
    aplicacoes = conn.execute(f"""SELECT a.id, a.titulo, a.modo, a.aberta, a.criada_em, p.titulo AS prova_titulo,
               t.nome AS turma_nome, t.ano_letivo, pr.nome AS criador_nome,
               (SELECT COUNT(*) FROM entregas WHERE aplicacao_id = a.id) AS qtd_entregas,
               (SELECT COUNT(*) FROM alunos WHERE turma_id = t.id) AS qtd_alunos
        FROM aplicacoes a JOIN provas p ON p.id = a.prova_id JOIN turmas t ON t.id = a.turma_id
        LEFT JOIN professores pr ON pr.id = a.criada_por_professor_id {wc} ORDER BY a.id DESC""", params_q).fetchall()
    turmas_lista = conn.execute("SELECT * FROM turmas ORDER BY ano_letivo DESC, nome").fetchall()
    professores_lista = conn.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall() if is_admin else []
    bf = "" if is_admin else f" AND (criada_por_professor_id = {prof['id']} OR criada_por_professor_id IS NULL)"
    t_ab = conn.execute(f"SELECT COUNT(*) AS c FROM aplicacoes WHERE aberta = 1{bf}").fetchone()["c"]
    t_enc = conn.execute(f"SELECT COUNT(*) AS c FROM aplicacoes WHERE aberta = 0{bf}").fetchone()["c"]
    conn.close()

    tabs_html = f"""<div style="display:flex; gap:0; border-bottom:2px solid var(--border); margin-bottom:18px;">
        <a href="/minhas-aplicacoes?aba=abertas" style="padding:10px 20px; font-weight:600; font-size:14px; text-decoration:none; border-bottom:3px solid {"var(--accent)" if aba=="abertas" else "transparent"}; color:{"var(--accent)" if aba=="abertas" else "var(--text-muted)"}; margin-bottom:-2px;">
           🟢 Abertas <span style="background:var(--green-bg); color:var(--green); border-radius:10px; padding:1px 7px; font-size:12px; margin-left:4px;">{t_ab}</span></a>
        <a href="/minhas-aplicacoes?aba=encerradas" style="padding:10px 20px; font-weight:600; font-size:14px; text-decoration:none; border-bottom:3px solid {"var(--accent)" if aba=="encerradas" else "transparent"}; color:{"var(--accent)" if aba=="encerradas" else "var(--text-muted)"}; margin-bottom:-2px;">
           🔒 Encerradas <span style="background:var(--bg-subtle); color:var(--text-muted); border-radius:10px; padding:1px 7px; font-size:12px; margin-left:4px;">{t_enc}</span></a>
    </div>"""

    turmas_opts = '<option value="">Todas as turmas</option>' + "".join(f'<option value="{t["id"]}"{" selected" if turma_id==t["id"] else ""}>{t["nome"]} ({t["ano_letivo"]})</option>' for t in turmas_lista)
    filtro_prof = ""
    if is_admin:
        po = '<option value="">Todos</option>' + "".join(f'<option value="{p["id"]}"{" selected" if prof_id==p["id"] else ""}>{p["nome"]}</option>' for p in professores_lista)
        filtro_prof = f'<label style="margin:0;">Professor<select name="prof_id">{po}</select></label>'
    filtros_html = f"""<form method="get" action="/minhas-aplicacoes" style="background:var(--bg-subtle); padding:12px 16px; border-radius:8px; margin-bottom:18px;">
        <input type="hidden" name="aba" value="{aba}">
        <div style="display:flex; gap:10px; align-items:flex-end; flex-wrap:wrap;">
            <label style="margin:0;">Turma<select name="turma">{turmas_opts}</select></label>{filtro_prof}
            <button type="submit" class="btn btn-primary" style="margin:0;">Filtrar</button>
            <a href="/minhas-aplicacoes?aba={aba}" class="btn" style="margin:0;">Limpar</a>
        </div></form>"""

    cards = ""
    for a in aplicacoes:
        titulo_apl = a["titulo"] or a["prova_titulo"]
        modo_icon = "📱" if a["modo"] == "online" else "📄"
        sb = ('<span style="background:var(--green-bg); color:var(--green); border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600;">Aberta</span>' if a["aberta"] else '<span style="background:var(--bg-subtle); color:var(--text-muted); border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600;">Encerrada</span>')
        prog = f'{a["qtd_entregas"]}/{a["qtd_alunos"]}' if a["qtd_alunos"] else "—"
        criador = f'<span style="font-size:11px; color:var(--text-muted);">por {a["criador_nome"]}</span>' if is_admin and a["criador_nome"] else ""
        cards += f"""<div style="border:1px solid var(--border); border-radius:10px; padding:14px 18px; margin-bottom:10px; background:var(--card);">
            <div style="display:flex; align-items:center; gap:10px; justify-content:space-between; flex-wrap:wrap;">
                <div><div style="font-weight:600; font-size:14px;">{modo_icon} {titulo_apl}</div>
                    <div style="font-size:12px; color:var(--text-muted); margin-top:3px;">{a["turma_nome"]} ({a["ano_letivo"]}) {criador}</div></div>
                <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">{sb}
                    <span style="font-size:12px; color:var(--text-muted);">Entregas: {prog}</span>
                    <a href="/aplicacoes/{a['id']}" class="btn" style="padding:5px 12px; font-size:12px;">Ver →</a>
                    <a href="/aplicacoes/{a['id']}/analise" class="btn" style="padding:5px 12px; font-size:12px;">📈</a>
                </div></div></div>"""
    if not cards:
        cards = f'<div class="empty">Nenhuma aplicação {"aberta" if aba=="abertas" else "encerrada"} encontrada.</div>'

    body = f"""<div class="page-header"><h1>📋 Minhas aplicações</h1><p class="subtitle">{"Visão geral da escola" if is_admin else "Suas atividades aplicadas"}</p></div>
        {tabs_html}{filtros_html}{cards}"""
    return render_page("Minhas aplicações", body, active="minhas-aplicacoes")


# ==========================================
#  ADMIN: GERENCIAMENTO DE USUÁRIOS
# ==========================================

@app.get("/admin/usuarios", response_class=HTMLResponse)
def admin_usuarios(request: Request):
    prof = _current_prof_ctx.get()
    if not prof or not prof.get("is_admin"):
        return HTMLResponse(render_page("Acesso negado", '<div class="empty">Apenas administradores.</div>', active=""), status_code=403)
    conn = get_db()
    usuarios = conn.execute(
        "SELECT id, email, nome, is_admin, is_gestor, status, criado_em, ultimo_acesso FROM professores ORDER BY CASE COALESCE(status,'ativo') WHEN 'pendente' THEN 0 WHEN 'ativo' THEN 1 ELSE 2 END, nome"
    ).fetchall()
    conn.close()
    pendentes = [u for u in usuarios if (u["status"] if "status" in u.keys() else "ativo") == "pendente"]
    alerta = f'''<div style="background:var(--orange-bg); border:1px solid var(--orange); border-radius:8px; padding:12px 16px; margin-bottom:18px; display:flex; align-items:center; gap:10px;">
        <span style="font-size:20px;">⏳</span><span><strong>{len(pendentes)} professor{"es" if len(pendentes)>1 else ""} aguardando aprovação.</strong></span></div>''' if pendentes else ""
    rows = ""
    for u in usuarios:
        us = u["status"] if "status" in u.keys() else "ativo"
        perfil = []
        if u["is_admin"]: perfil.append('<span style="background:#7c3aed; color:white; font-size:10px; padding:1px 6px; border-radius:3px;">ADMIN</span>')
        if u["is_gestor"]: perfil.append('<span style="background:var(--accent); color:white; font-size:10px; padding:1px 6px; border-radius:3px;">GESTOR</span>')
        if us == "pendente": perfil.append('<span style="background:var(--orange-bg); color:var(--orange); font-size:10px; padding:1px 6px; border-radius:3px; border:1px solid var(--orange);">PENDENTE</span>')
        elif us == "bloqueado": perfil.append('<span style="background:var(--red-bg); color:var(--red); font-size:10px; padding:1px 6px; border-radius:3px;">BLOQUEADO</span>')
        elif not u["is_admin"] and not u["is_gestor"]: perfil.append('<span style="background:var(--bg-subtle); color:var(--text-muted); font-size:10px; padding:1px 6px; border-radius:3px;">PROFESSOR</span>')
        acesso = (u["ultimo_acesso"] or "")[:16].replace("T", " ")
        is_eu = u["id"] == prof["id"]
        acoes = ""
        if not is_eu:
            if us == "pendente":
                acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/aprovar" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px; color:var(--green); border-color:var(--green);">✅ Aprovar</button></form>'
                acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/bloquear" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px; color:var(--red); border-color:var(--red);">🚫 Bloquear</button></form>'
            else:
                lg = "Remover gestor" if u["is_gestor"] else "Tornar gestor"
                acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/toggle-gestor" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px;">{lg}</button></form>'
                if not u["is_admin"]:
                    acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/toggle-admin" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px; color:var(--red); border-color:var(--red);">Admin</button></form>'
                if us == "ativo":
                    acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/bloquear" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px; color:var(--red); border-color:var(--red);">🚫</button></form>'
                elif us == "bloqueado":
                    acoes += f'<form method="post" action="/admin/usuarios/{u["id"]}/aprovar" style="margin:0; display:inline;"><button type="submit" class="btn" style="padding:4px 10px; font-size:11px; color:var(--green); border-color:var(--green);">✅ Desbloquear</button></form>'
        rb = ' style="background:var(--orange-bg);"' if us == "pendente" else ""
        rows += f'''<tr{rb}><td style="padding:10px 8px;">{u["nome"]}{"&nbsp;<em style=\'font-size:11px; color:var(--text-muted);\'>( você)</em>" if is_eu else ""}</td>
            <td style="padding:10px 8px; font-size:12px; color:var(--text-muted);">{u["email"]}</td>
            <td style="padding:10px 8px;">{" ".join(perfil)}</td>
            <td style="padding:10px 8px; font-size:12px; color:var(--text-muted);">{acesso}</td>
            <td style="padding:10px 8px;">{acoes}</td></tr>'''
    body = f"""<div class="page-header"><h1>👤 Gerenciamento de usuários</h1><p class="subtitle">Gerencie perfis de acesso</p></div>
        {alerta}
        <div class="tip" style="margin-bottom:18px;"><strong>Perfis:</strong> Admin — acesso total. Gestor — aprova/devolve provas. Professor — acesso padrão. Novos usuários entram como <strong>Pendente</strong>.</div>
        <table style="width:100%; border-collapse:collapse; background:var(--card); border-radius:10px; overflow:hidden; border:1px solid var(--border);">
            <thead><tr style="background:var(--bg-subtle); font-size:12px; text-transform:uppercase; color:var(--text-muted);">
                <th style="padding:10px 8px; text-align:left;">Nome</th><th style="padding:10px 8px; text-align:left;">E-mail</th>
                <th style="padding:10px 8px; text-align:left;">Perfil</th><th style="padding:10px 8px; text-align:left;">Último acesso</th>
                <th style="padding:10px 8px; text-align:left;">Ações</th></tr></thead>
            <tbody>{rows}</tbody></table>"""
    return render_page("Usuários", body, active="")


@app.post("/admin/usuarios/{usuario_id}/toggle-gestor")
def toggle_gestor(usuario_id: int):
    prof = _current_prof_ctx.get()
    if not prof or not prof.get("is_admin"): return RedirectResponse("/login", status_code=303)
    conn = get_db()
    atual = conn.execute("SELECT is_gestor FROM professores WHERE id = ?", (usuario_id,)).fetchone()
    if atual:
        conn.execute("UPDATE professores SET is_gestor = ? WHERE id = ?", (0 if atual["is_gestor"] else 1, usuario_id)); conn.commit()
    conn.close()
    return RedirectResponse("/admin/usuarios", status_code=303)


@app.post("/admin/usuarios/{usuario_id}/toggle-admin")
def toggle_admin(usuario_id: int):
    prof = _current_prof_ctx.get()
    if not prof or not prof.get("is_admin"): return RedirectResponse("/login", status_code=303)
    conn = get_db()
    atual = conn.execute("SELECT is_admin FROM professores WHERE id = ?", (usuario_id,)).fetchone()
    if atual:
        conn.execute("UPDATE professores SET is_admin = ? WHERE id = ?", (0 if atual["is_admin"] else 1, usuario_id)); conn.commit()
    conn.close()
    return RedirectResponse("/admin/usuarios", status_code=303)


@app.post("/admin/usuarios/{usuario_id}/aprovar")
def aprovar_usuario(usuario_id: int):
    prof = _current_prof_ctx.get()
    if not prof or not prof.get("is_admin"): return RedirectResponse("/login", status_code=303)
    conn = get_db()
    conn.execute("UPDATE professores SET status = 'ativo' WHERE id = ?", (usuario_id,)); conn.commit(); conn.close()
    return RedirectResponse("/admin/usuarios", status_code=303)


@app.post("/admin/usuarios/{usuario_id}/bloquear")
def bloquear_usuario(usuario_id: int):
    prof = _current_prof_ctx.get()
    if not prof or not prof.get("is_admin"): return RedirectResponse("/login", status_code=303)
    conn = get_db()
    conn.execute("UPDATE professores SET status = 'bloqueado' WHERE id = ?", (usuario_id,)); conn.commit(); conn.close()
    return RedirectResponse("/admin/usuarios", status_code=303)


@app.get("/aplicacoes", response_class=HTMLResponse)
def listar_aplicacoes(
    request: Request,
    turma: Optional[str] = None,
    modo: Optional[str] = None,
    status: Optional[str] = None,
    q: Optional[str] = None,
):
    # Aceita string vazia ("Todas") sem dar erro de parsing
    turma_id: Optional[int] = None
    if turma and turma.strip().isdigit():
        turma_id = int(turma)
    prof = get_current_professor(request)
    is_admin = prof and prof["is_admin"]
    conn = get_db()

    # Filtros: admin vê tudo; prof comum só as próprias
    where_extras = []
    params = []
    if not is_admin:
        where_extras.append("(a.criada_por_professor_id = ? OR a.criada_por_professor_id IS NULL)")
        params.append(prof["id"])
    if q and q.strip():
        where_extras.append("(a.titulo LIKE ? OR p.titulo LIKE ?)")
        params.append(f"%{q.strip()}%")
        params.append(f"%{q.strip()}%")
    if turma_id:
        where_extras.append("a.turma_id = ?")
        params.append(turma_id)
    if modo and modo in ("online", "impressa"):
        where_extras.append("a.modo = ?")
        params.append(modo)
    if status == "aberta":
        where_extras.append("a.aberta = 1")
    elif status == "encerrada":
        where_extras.append("a.aberta = 0")
    where_clause = " WHERE " + " AND ".join(where_extras) if where_extras else ""

    sql = f"""
        SELECT a.id, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo,
               a.criada_em, a.aberta, a.modo, a.titulo, a.criada_por_professor_id,
               prof.nome AS criador_nome,
               (SELECT COUNT(*) FROM entregas WHERE aplicacao_id = a.id) AS qtd_entregas,
               (SELECT COUNT(*) FROM alunos WHERE turma_id = t.id) AS qtd_alunos
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        LEFT JOIN professores prof ON prof.id = a.criada_por_professor_id
        {where_clause}
        ORDER BY a.id DESC
    """
    aplicacoes = conn.execute(sql, params).fetchall()

    turmas_lista = conn.execute("SELECT * FROM turmas ORDER BY ano_letivo DESC, nome").fetchall()
    if is_admin:
        total_geral = conn.execute("SELECT COUNT(*) AS c FROM aplicacoes").fetchone()["c"]
    else:
        total_geral = conn.execute(
            "SELECT COUNT(*) AS c FROM aplicacoes WHERE (criada_por_professor_id = ? OR criada_por_professor_id IS NULL)",
            (prof["id"],)
        ).fetchone()["c"]
    conn.close()

    # Filtros
    turmas_opts = '<option value="">Todas</option>' + "".join(
        f'<option value="{t["id"]}"{(" selected" if turma_id == t["id"] else "")}>{t["nome"]} ({t["ano_letivo"]})</option>'
        for t in turmas_lista
    )
    modos_opts = (
        '<option value="">Todos</option>'
        f'<option value="online"{(" selected" if modo == "online" else "")}>📱 Online</option>'
        f'<option value="impressa"{(" selected" if modo == "impressa" else "")}>📄 Impressa</option>'
    )
    status_opts = (
        '<option value="">Todos</option>'
        f'<option value="aberta"{(" selected" if status == "aberta" else "")}>Aberta</option>'
        f'<option value="encerrada"{(" selected" if status == "encerrada" else "")}>Encerrada</option>'
    )

    filtros_html = (
        f'<form action="/aplicacoes" method="get" '
        f'style="background:var(--bg-subtle); padding:14px 16px; border-radius:8px; margin-bottom:18px;">'
        f'<div style="display:grid; grid-template-columns: 2fr 1.3fr 1fr 1fr auto auto; gap:10px; align-items:end;">'
        f'<label style="margin:0;">Buscar por título<input type="text" name="q" placeholder="palavra do título" value="{q or ""}"></label>'
        f'<label style="margin:0;">Turma<select name="turma">{turmas_opts}</select></label>'
        f'<label style="margin:0;">Modo<select name="modo">{modos_opts}</select></label>'
        f'<label style="margin:0;">Status<select name="status">{status_opts}</select></label>'
        f'<button type="submit" class="btn btn-primary" style="margin:0;">Filtrar</button>'
        f'<a href="/aplicacoes" class="btn" style="margin:0;">Limpar</a>'
        f'</div></form>'
    )

    # Cards
    if aplicacoes:
        cards = ""
        for a in aplicacoes:
            titulo_apl = a["titulo"] if a["titulo"] else a["prova_titulo"]
            subtitulo = f'<div style="font-size:13px; color:var(--text-muted); margin-top:2px;">{a["prova_titulo"]}</div>' if a["titulo"] and a["titulo"] != a["prova_titulo"] else ""

            modo_badge = (
                '<span class="badge" style="background:var(--accent-bg); color:var(--accent);">📱 Online</span>'
                if a["modo"] == "online"
                else '<span class="badge" style="background:var(--orange-bg); color:var(--orange);">📄 Impressa</span>'
            )
            status_badge = (
                '<span class="badge" style="background:var(--green-bg); color:var(--green);">Aberta</span>'
                if a["aberta"]
                else '<span class="badge" style="background:var(--bg-muted); color:var(--text-muted);">Encerrada</span>'
            )
            turma_badge = f'<span class="badge">{a["turma_nome"]} ({a["ano_letivo"]})</span>'

            n_e = a["qtd_entregas"] or 0
            n_a = a["qtd_alunos"] or 0
            progresso_color = "var(--green)" if n_e == n_a and n_a > 0 else ("var(--orange)" if n_e > 0 else "var(--text-muted)")
            progresso_badge = f'<span class="badge" style="color:{progresso_color};">{n_e}/{n_a} entregas</span>'

            data_str = f'<span style="font-size:12px; color:var(--text-muted);">{format_data_br(a["criada_em"])}</span>' if a["criada_em"] else ""

            # Badge "Por: <nome>" só pra admin
            autor_badge = ""
            if is_admin:
                nome_autor = a["criador_nome"] if a["criador_nome"] else "—"
                autor_badge = f'<span class="badge" style="background:var(--purple-bg); color:var(--purple);">Por: {nome_autor}</span>'

            cards += f"""
            <div style="background:var(--bg); border:1px solid var(--border); border-radius:8px; padding:14px 18px; margin-bottom:10px;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:14px;">
                    <div style="flex:1; min-width:0;">
                        <div style="font-weight:600; font-size:16px;">
                            <a href="/aplicacoes/{a["id"]}" style="color:inherit; text-decoration:none;">{titulo_apl}</a>
                        </div>
                        {subtitulo}
                        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:8px; align-items:center;">
                            {turma_badge}{modo_badge}{status_badge}{progresso_badge}{autor_badge}{data_str}
                        </div>
                    </div>
                    <div style="display:flex; gap:6px; flex-shrink:0; flex-wrap:wrap; justify-content:flex-end;">
                        <a href="/aplicacoes/{a["id"]}" class="btn btn-primary" style="padding:4px 10px; font-size:12px;">Abrir</a>
                        <form action="/aplicacoes/{a["id"]}/deletar" method="post" style="margin:0;" onsubmit="return confirm('Excluir esta aplicação? Os registros de entregas, respostas e notas dos alunos serão apagados permanentemente. Esta ação não pode ser desfeita.');">
                            <button type="submit" class="btn" style="padding:4px 10px; font-size:12px; background:var(--red); color:white; border-color:var(--red);">Excluir</button>
                        </form>
                    </div>
                </div>
            </div>
            """
    else:
        cards = '<div class="empty">Nenhuma aplicação encontrada com esses filtros.</div>' if (turma_id or modo or status or q) else '<div class="empty">Nenhuma aplicação gerada ainda.</div>'

    tem_filtro = bool(turma_id or modo or status or q)
    subtitle = f'{len(aplicacoes)} de {total_geral} aplicação(ões)' if tem_filtro else f'{total_geral} aplicação(ões) cadastrada(s)'

    content = f"""
        <div class="page-header">
            <h1>Aplicações</h1>
            <p class="subtitle">{subtitle}</p>
            <div class="page-actions"><a href="/aplicacoes/nova" class="btn btn-primary">+ Nova Aplicação</a></div>
        </div>
        {filtros_html}
        {cards}
    """
    return render_page("Aplicações", content, active="aplicacoes")


@app.get("/aplicacoes/nova", response_class=HTMLResponse)
def form_nova_aplicacao():
    conn = get_db()
    provas = conn.execute("SELECT id, titulo FROM provas ORDER BY criada_em DESC").fetchall()
    turmas = conn.execute("SELECT id, nome, ano_letivo FROM turmas ORDER BY ano_letivo DESC, nome").fetchall()
    conn.close()

    if not provas or not turmas:
        falta = []
        if not provas:
            falta.append('<p><a href="/provas/nova" class="btn">Criar prova</a></p>')
        if not turmas:
            falta.append('<p><a href="/turmas/nova" class="btn">Criar turma</a></p>')
        content = f"""
            <div class="page-header"><h1>Nova aplicação</h1></div>
            <div class="empty">
                <p>Você precisa de pelo menos uma prova e uma turma cadastradas.</p>
                {"".join(falta)}
            </div>
        """
        return render_page("Nova aplicação", content, active="aplicacoes")

    provas_options = "".join(f'<option value="{p["id"]}">{p["titulo"]}</option>' for p in provas)
    turmas_options = "".join(f'<option value="{t["id"]}">{t["nome"]} ({t["ano_letivo"]})</option>' for t in turmas)

    content = f"""
        <div class="page-header"><h1>Nova aplicação</h1></div>
        <form action="/aplicacoes/nova" method="post">
            <label>Prova<select name="prova_id" required>{provas_options}</select></label>
            <label>Turma<select name="turma_id" required>{turmas_options}</select></label>

            <fieldset>
                <legend>Modo de aplicação</legend>
                <label style="font-weight:normal; display:flex; align-items:flex-start; gap:8px; margin-bottom:12px;">
                    <input type="radio" name="modo" value="online" required checked style="width:auto; margin-top:4px;">
                    <span><strong>Online</strong><br><small>Cada aluno recebe um link único para responder pelo celular ou computador.</small></span>
                </label>
                <label style="font-weight:normal; display:flex; align-items:flex-start; gap:8px;">
                    <input type="radio" name="modo" value="impressa" style="width:auto; margin-top:4px;">
                    <span><strong>Impressa</strong><br><small>Prova e cartão resposta serão impressos. Correção por foto do cartão (OMR — virá na Fase 3).</small></span>
                </label>
            </fieldset>

            <label>Título da aplicação (opcional)<input type="text" name="titulo" placeholder="Ex: 1º Bimestre — 9º A"></label>

            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Criar aplicação</button>
                <a href="/aplicacoes" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return render_page("Nova aplicação", content, active="aplicacoes")


@app.post("/aplicacoes/nova")
def criar_aplicacao(request: Request, prova_id: int = Form(...), turma_id: int = Form(...), modo: str = Form(...), titulo: str = Form("")):
    prof = get_current_professor(request)
    if not prof:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    cursor = conn.execute(
        "INSERT INTO aplicacoes (prova_id, turma_id, modo, titulo, criada_por_professor_id) VALUES (?, ?, ?, ?, ?)",
        (prova_id, turma_id, modo, titulo.strip() or None, prof["id"])
    )
    aplicacao_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return RedirectResponse(f"/aplicacoes/{aplicacao_id}", status_code=303)


@app.post("/aplicacoes/{id}/deletar")
def deletar_aplicacao(id: int):
    conn = get_db()
    conn.execute("DELETE FROM respostas WHERE aplicacao_id = ?", (id,))
    conn.execute("DELETE FROM entregas WHERE aplicacao_id = ?", (id,))
    conn.execute("DELETE FROM aplicacoes WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/aplicacoes", status_code=303)


@app.get("/aplicacoes/{aplicacao_id}", response_class=HTMLResponse)
def ver_aplicacao(aplicacao_id: int, request: Request):
    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()

    if not apl:
        conn.close()
        return HTMLResponse(render_page("Não encontrada", '<h1>Aplicação não encontrada</h1><p><a href="/aplicacoes">← Voltar</a></p>', active="aplicacoes"), status_code=404)

    alunos_atuais = conn.execute("SELECT *, NULL AS transferido_para FROM alunos WHERE turma_id = ? ORDER BY numero, nome", (apl["turma_id"],)).fetchall()
    fantasmas = conn.execute("""
        SELECT a.*, t.nome AS transferido_para FROM alunos a
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.turma_id != ?
          AND (a.id IN (SELECT DISTINCT aluno_id FROM entregas WHERE aplicacao_id = ?)
            OR a.id IN (SELECT DISTINCT aluno_id FROM respostas WHERE aplicacao_id = ?))
        ORDER BY a.nome
    """, (apl["turma_id"], aplicacao_id, aplicacao_id)).fetchall()
    alunos = list(alunos_atuais) + list(fantasmas)
    entregas = {row["aluno_id"]: row["finalizada_em"] for row in conn.execute("SELECT aluno_id, finalizada_em FROM entregas WHERE aplicacao_id = ?", (aplicacao_id,)).fetchall()}
    total_questoes = conn.execute("SELECT COUNT(*) AS c FROM prova_questoes WHERE prova_id = ?", (apl["prova_id"],)).fetchone()["c"]

    alunos_data = {}
    notas = []
    for a in alunos:
        if a["id"] in entregas:
            score, _ = _calcular_nota(conn, aplicacao_id, a["id"])
            alunos_data[a["id"]] = {"entregue": True, "score": score}
            notas.append(score)
        else:
            alunos_data[a["id"]] = {"entregue": False}

    media_nota = sum(notas) / len(notas) if notas else 0
    conn.close()

    base_url = get_base_url(request)
    titulo = apl["titulo"] or f'{apl["prova_titulo"]} — {apl["turma_nome"]}'
    modo_label = "Online" if apl["modo"] == "online" else "Impressa"

    acoes_btn = '<div class="page-actions" style="display:flex; gap:8px; flex-wrap:wrap;">'
    if apl["modo"] == "online":
        acoes_btn += f'<a href="/aplicacoes/{aplicacao_id}/cartoes" class="btn btn-primary" target="_blank">Folha com QR Codes</a>'
    else:
        acoes_btn += f'<a href="/aplicacoes/{aplicacao_id}/cartao-resposta" class="btn btn-primary">📄 Cartões Resposta (PDF)</a>'
        acoes_btn += f'<a href="/aplicacoes/{aplicacao_id}/escanear" class="btn btn-primary" style="background:var(--green); border-color:var(--green);">📷 Escanear cartão</a>'
        acoes_btn += f'<a href="/provas/{apl["prova_id"]}/imprimir" class="btn" target="_blank">🖨️ Imprimir prova</a>'
    acoes_btn += f'<a href="/aplicacoes/{aplicacao_id}/analise" class="btn">📈 Análise pedagógica</a>'
    acoes_btn += f'<a href="/aplicacoes/{aplicacao_id}/exportar" class="btn">📊 Exportar Planilha Excel</a>'
    if apl["aberta"]:
        acoes_btn += f'<form method="post" action="/aplicacoes/{aplicacao_id}/encerrar" style="margin:0;" onsubmit="return confirm(\'Encerrar esta aplicação?\')"><button type="submit" class="btn" style="color:var(--red); border-color:var(--red);">🔒 Encerrar aplicação</button></form>'
    else:
        acoes_btn += f'<form method="post" action="/aplicacoes/{aplicacao_id}/reabrir" style="margin:0;"><button type="submit" class="btn" style="color:var(--green); border-color:var(--green);">🔓 Reabrir aplicação</button></form>'
    acoes_btn += '</div>'

    metrics_html = ""
    if alunos:
        n_atuais = len(alunos_atuais)
        pendentes_atuais = sum(1 for a in alunos_atuais if a["id"] not in entregas)
        metrics_html = f"""
        <div class="metric-grid">
            <div class="metric"><div class="metric-label">Alunos da turma</div><div class="metric-value">{n_atuais}</div></div>
            <div class="metric"><div class="metric-label">Entregaram</div><div class="metric-value">{len(entregas)}</div></div>
            <div class="metric"><div class="metric-label">Pendentes</div><div class="metric-value">{pendentes_atuais}</div></div>
            <div class="metric"><div class="metric-label">Média</div><div class="metric-value">{media_nota:.1f}/{total_questoes}</div></div>
        </div>"""

    if alunos:
        rows = ""
        for a in alunos:
            num = a["numero"] if a["numero"] else "—"
            data = alunos_data[a["id"]]
            transferido_badge = ""
            if a["transferido_para"]:
                transferido_badge = f' <span class="badge" style="background:var(--orange-bg); color:var(--orange); font-size:10px; vertical-align:middle;">→ {a["transferido_para"]}</span>'
            if data["entregue"]:
                status_html = f'<div style="margin-top:4px;"><span class="badge" style="background:var(--success-bg, var(--green-bg)); color:var(--success-text, var(--green));">Entregue · {data["score"]}/{total_questoes}</span></div>'
            elif apl["modo"] == "online" and not a["transferido_para"]:
                url = f'{base_url}/responder/{a["codigo_unico"]}/{aplicacao_id}'
                status_html = f'<input type="text" value="{url}" readonly style="font-family:\'SF Mono\',monospace; font-size:11px; width:100%; margin:4px 0 0; padding:6px 8px;" onclick="this.select()">'
            elif a["transferido_para"]:
                status_html = '<div style="margin-top:4px;"><span class="badge" style="background:var(--bg-muted); color:var(--text-muted);">Respostas parciais (não finalizou antes de transferir)</span></div>'
            else:
                status_html = '<div style="margin-top:4px;"><span class="badge">Aguardando leitura OMR</span></div>'

            nome_link = f'<a href="/aplicacoes/{aplicacao_id}/aluno/{a["id"]}" style="color:inherit; text-decoration:none; font-weight:600;">{a["nome"]}</a>{transferido_badge}'
            rows += f'<div class="student-row"><div class="numero">{num}</div><div>{nome_link}{status_html}</div><div class="codigo">{a["codigo_unico"]}</div></div>'
        lista_html = rows
    else:
        lista_html = '<div class="empty">Esta turma não tem alunos cadastrados.</div>'

    content = f"""
        <div class="page-header">
            <h1>{titulo}</h1>
            <p class="subtitle">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]}) · Modo {modo_label}</p>
            {acoes_btn}
        </div>
        {metrics_html}
        <h2>Alunos</h2>
        {lista_html}
        <p style="margin-top:24px;"><a href="/aplicacoes" class="btn">← Voltar</a></p>
    """
    return render_page(titulo, content, active="aplicacoes")


@app.post("/aplicacoes/{aplicacao_id}/encerrar")
def encerrar_aplicacao(aplicacao_id: int):
    conn = get_db()
    conn.execute("UPDATE aplicacoes SET aberta = 0 WHERE id = ?", (aplicacao_id,))
    conn.commit(); conn.close()
    return RedirectResponse(f"/aplicacoes/{aplicacao_id}", status_code=303)


@app.post("/aplicacoes/{aplicacao_id}/reabrir")
def reabrir_aplicacao(aplicacao_id: int):
    conn = get_db()
    conn.execute("UPDATE aplicacoes SET aberta = 1 WHERE id = ?", (aplicacao_id,))
    conn.commit(); conn.close()
    return RedirectResponse(f"/aplicacoes/{aplicacao_id}", status_code=303)


@app.get("/responder/{codigo}/{aplicacao_id}", response_class=HTMLResponse)
def pagina_resposta_aluno(codigo: str, aplicacao_id: int):
    conn = get_db()
    aluno = conn.execute("SELECT * FROM alunos WHERE codigo_unico = ?", (codigo,)).fetchone()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()

    if not aluno or not apl or aluno["turma_id"] != apl["turma_id"]:
        conn.close()
        return HTMLResponse(_pagina_simples("Link inválido", "<p>Esse link não corresponde a uma aplicação válida. Confira com seu professor.</p>"))

    entrega = conn.execute("SELECT * FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno["id"])).fetchone()
    if entrega:
        score, total = _calcular_nota(conn, aplicacao_id, aluno["id"])
        conn.close()
        return HTMLResponse(_pagina_simples(
            "Prova entregue",
            f"""
            <p>Olá, <strong>{aluno["nome"]}</strong>! Sua prova foi entregue em {entrega["finalizada_em"]}.</p>
            <p style="font-size:32px; font-weight:600; margin: 24px 0;">Nota: {score}/{total}</p>
            <p>Você não pode mais alterar suas respostas.</p>
            """
        ))

    questoes = conn.execute("""
        SELECT q.id, q.enunciado, d.nome AS disciplina_nome
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        JOIN disciplinas d ON d.id = q.disciplina_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()

    respostas_dadas = {}
    for r in conn.execute("SELECT questao_id, alternativa_letra FROM respostas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno["id"])).fetchall():
        respostas_dadas[r["questao_id"]] = r["alternativa_letra"]

    questoes_html = ""
    for idx, q in enumerate(questoes, start=1):
        textos = conn.execute("SELECT conteudo, fonte FROM textos_apoio WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
        imagens = conn.execute("SELECT caminho, legenda FROM imagens WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
        alts = conn.execute("SELECT letra, texto FROM alternativas WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()

        textos_html = ""
        for t in textos:
            fonte_html = f'<footer>Fonte: {t["fonte"]}</footer>' if t["fonte"] else ""
            textos_html += f'<blockquote>{t["conteudo"]}{fonte_html}</blockquote>'

        imagens_html = ""
        for img in imagens:
            legenda_html = f'<figcaption>{img["legenda"]}</figcaption>' if img["legenda"] else ""
            imagens_html += f'<figure><img src="/{img["caminho"]}" alt="">{legenda_html}</figure>'

        marcada = respostas_dadas.get(q["id"], "")
        alts_html = ""
        for a in alts:
            checked = ' checked' if a["letra"] == marcada else ''
            alts_html += f'<label style="display:flex; gap:10px; padding:10px 12px; border:1px solid var(--border); border-radius:6px; margin-bottom:6px; cursor:pointer; align-items:flex-start;"><input type="radio" name="q_{q["id"]}" value="{a["letra"]}"{checked} style="width:auto; margin-top:3px;"><span><strong>{a["letra"]})</strong> {a["texto"]}</span></label>'

        bncc_pref = _bncc_prefix(conn, q["id"])
        questoes_html += f'<div class="question"><div class="question-header">Questão {idx} · {q["disciplina_nome"]}</div>{textos_html}{imagens_html}<div class="enunciado">{bncc_pref}{q["enunciado"]}</div>{alts_html}</div>'

    conn.close()

    content = f"""
        <div class="page-header">
            <h1>{apl["prova_titulo"]}</h1>
            <p class="subtitle">Aluno: <strong>{aluno["nome"]}</strong> · Turma: {apl["turma_nome"]}</p>
        </div>
        <div class="tip">Marque a alternativa correta em cada questão. Você pode salvar o progresso e voltar depois usando o mesmo link. Quando terminar, clique em <strong>Finalizar e entregar</strong>.</div>
        <form action="/responder/{codigo}/{aplicacao_id}" method="post">
            {questoes_html}
            <div class="page-actions" style="margin-top:24px;">
                <button type="submit" name="acao" value="salvar" class="btn">Salvar progresso</button>
                <button type="submit" name="acao" value="finalizar" class="btn btn-primary" onclick="return confirm('Tem certeza? Após finalizar, você não pode mais alterar as respostas.');">Finalizar e entregar</button>
            </div>
        </form>
    """
    return HTMLResponse(_pagina_aluno(apl["prova_titulo"], content))

@app.get("/aplicacoes/{aplicacao_id}/cartoes", response_class=HTMLResponse)
def cartoes_qr(aplicacao_id: int, request: Request):
    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()

    if not apl:
        conn.close()
        return HTMLResponse("<h1>Aplicação não encontrada</h1>", status_code=404)

    alunos = conn.execute("SELECT * FROM alunos WHERE turma_id = ? ORDER BY numero, nome", (apl["turma_id"],)).fetchall()
    conn.close()

    base_url = get_base_url(request)
    titulo = apl["titulo"] or f'{apl["prova_titulo"]} — {apl["turma_nome"]}'

    cards_html = ""
    for a in alunos:
        url = f'{base_url}/responder/{a["codigo_unico"]}/{aplicacao_id}'
        qr_src = qr_data_uri(url)
        num = a["numero"] if a["numero"] else "—"
        cards_html += f"""
        <div class="qr-card">
            <div class="qr-name">{a["nome"]}</div>
            <div class="qr-meta">Nº {num} · <code>{a["codigo_unico"]}</code></div>
            <img src="{qr_src}" alt="QR Code">
            <div class="qr-instr">Aponte a câmera do celular</div>
        </div>"""

    return HTMLResponse(f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <title>Cartões QR — {titulo}</title>
    {INTER_FONT}
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Sora', -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 24px;
            background: var(--bg-subtle);
            color: var(--text);
        }}
        .container {{ max-width: 21cm; margin: 0 auto; background: white; padding: 32px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        h1 {{ font-size: 22px; margin: 0 0 4px; }}
        .meta {{ color: var(--text-muted); font-size: 14px; margin: 0 0 24px; }}
        .actions {{ margin-bottom: 24px; display: flex; gap: 8px; }}
        .btn {{ display: inline-block; padding: 8px 16px; border-radius: 6px; border: 1px solid #d4d4d8; background: white; color: var(--text); text-decoration: none; font-size: 14px; font-weight: 500; font-family: inherit; cursor: pointer; }}
        .btn-primary {{ background: var(--accent); border-color: var(--accent); color: white; }}
        .qr-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }}
        .qr-card {{ border: 1px dashed var(--text-muted); border-radius: 6px; padding: 12px; text-align: center; page-break-inside: avoid; break-inside: avoid; }}
        .qr-name {{ font-size: 13px; font-weight: 600; margin-bottom: 4px; }}
        .qr-meta {{ font-size: 11px; color: var(--text-muted); margin-bottom: 8px; }}
        .qr-meta code {{ font-family: 'SF Mono', Monaco, monospace; background: #f4f4f5; padding: 1px 4px; border-radius: 3px; }}
        .qr-card img {{ width: 100%; max-width: 160px; height: auto; display: block; margin: 0 auto; }}
        .qr-instr {{ font-size: 10px; color: var(--text-muted); margin-top: 6px; font-style: italic; }}
        @media print {{
            @page {{ size: A4; margin: 1cm; }}
            body {{ background: white; padding: 0; }}
            .container {{ box-shadow: none; padding: 0; border-radius: 0; }}
            .no-print {{ display: none !important; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{titulo}</h1>
        <p class="meta">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]}) · {len(alunos)} alunos</p>
        <div class="actions no-print">
            <button onclick="window.print()" class="btn btn-primary">Imprimir / Salvar como PDF</button>
            <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar</a>
        </div>
        <div class="qr-grid">{cards_html}</div>
    </div>
</body>
</html>""")

@app.post("/responder/{codigo}/{aplicacao_id}")
async def salvar_respostas(codigo: str, aplicacao_id: int, request: Request):
    form = await request.form()
    acao = form.get("acao", "salvar")

    conn = get_db()
    aluno = conn.execute("SELECT * FROM alunos WHERE codigo_unico = ?", (codigo,)).fetchone()
    apl = conn.execute("SELECT * FROM aplicacoes WHERE id = ?", (aplicacao_id,)).fetchone()

    if not aluno or not apl or aluno["turma_id"] != apl["turma_id"]:
        conn.close()
        return HTMLResponse(_pagina_simples("Link inválido", "<p>Link inválido.</p>"))

    entrega = conn.execute("SELECT * FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno["id"])).fetchone()
    if entrega:
        conn.close()
        return RedirectResponse(f"/responder/{codigo}/{aplicacao_id}", status_code=303)

    for key, value in form.items():
        if key.startswith("q_") and value:
            questao_id = int(key[2:])
            conn.execute("""
                INSERT INTO respostas (aplicacao_id, aluno_id, questao_id, alternativa_letra)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(aplicacao_id, aluno_id, questao_id)
                DO UPDATE SET alternativa_letra = excluded.alternativa_letra, respondida_em = CURRENT_TIMESTAMP
            """, (aplicacao_id, aluno["id"], questao_id, value))

    if acao == "finalizar":
        conn.execute("INSERT INTO entregas (aplicacao_id, aluno_id) VALUES (?, ?)", (aplicacao_id, aluno["id"]))

    conn.commit()
    conn.close()
    return RedirectResponse(f"/responder/{codigo}/{aplicacao_id}", status_code=303)


def _bncc_prefix(conn, questao_id):
    """Retorna string com códigos BNCC formatados como '(COD1, COD2) ' ou '' se não tem.
    Usado para prefixar o enunciado em telas de impressão e resposta online."""
    rows = conn.execute("""
        SELECT h.codigo FROM questao_habilidades qh
        JOIN habilidades_bncc h ON h.id = qh.habilidade_id
        WHERE qh.questao_id = ? ORDER BY h.codigo
    """, (questao_id,)).fetchall()
    if not rows:
        return ""
    codigos = ", ".join(r["codigo"] for r in rows)
    return f"<strong>({codigos})</strong> "


def _gravar_resposta_questao(conn, aplicacao_id, aluno_id, questao_id, tipo, valor):
    """Grava uma resposta no formato correto pro tipo da questão.
    valor:
      - multipla_escolha: string "A"/"B"/"C"/"D"/None
      - vf: dict {"0":"V", "1":"F", ...}
      - associacao: dict {"0":"b", "1":"a", ...}
      - discursiva: ignora (correção manual fora do sistema)"""
    import json as _json
    if tipo == "multipla_escolha":
        if valor and valor in ("A", "B", "C", "D"):
            conn.execute(
                "INSERT INTO respostas (aplicacao_id, aluno_id, questao_id, alternativa_letra) VALUES (?, ?, ?, ?)",
                (aplicacao_id, aluno_id, questao_id, valor)
            )
    elif tipo in ("vf", "associacao"):
        if isinstance(valor, dict) and any(v for v in valor.values() if v):
            conn.execute(
                "INSERT INTO respostas (aplicacao_id, aluno_id, questao_id, dados_extra) VALUES (?, ?, ?, ?)",
                (aplicacao_id, aluno_id, questao_id, _json.dumps(valor))
            )
    # discursiva: nada


# ═══════════════════════════════════════════════════════════════
# FAIXAS DE PROFICIÊNCIA SAEB (escala 0-10)
# ═══════════════════════════════════════════════════════════════
FAIXAS_SAEB = [
    {"nome": "Insuficiente", "min": 0.0,  "max": 5.0,  "cor": "var(--red)",    "cor_bg": "var(--red-bg)",    "cor_border": "var(--red-border)",    "emoji": "🔴", "hex": "#dc2626"},
    {"nome": "Básico",       "min": 5.0,  "max": 6.6,  "cor": "var(--orange)", "cor_bg": "var(--orange-bg)", "cor_border": "var(--orange-border)", "emoji": "🟡", "hex": "#ea580c"},
    {"nome": "Adequado",     "min": 6.6,  "max": 8.0,  "cor": "var(--accent)", "cor_bg": "var(--accent-bg)", "cor_border": "var(--accent-border)", "emoji": "🔵", "hex": "#0284c7"},
    {"nome": "Avançado",     "min": 8.0,  "max": 10.01,"cor": "var(--green)",  "cor_bg": "var(--green-bg)",  "cor_border": "var(--green-border)",  "emoji": "🟢", "hex": "#16a34a"},
]

def _faixa_saeb(nota_10):
    """Retorna o dict da faixa SAEB pra uma nota de 0 a 10."""
    for f in FAIXAS_SAEB:
        if f["min"] <= nota_10 < f["max"]:
            return f
    return FAIXAS_SAEB[-1] if nota_10 >= 8.0 else FAIXAS_SAEB[0]


def _calcular_nota_objetiva(conn, aplicacao_id, aluno_id):
    """Variante de _calcular_nota que IGNORA discursivas no total.
    Retorna (acertos, total_objetivas, nota_10). Usado pra análises SAEB."""
    import json as _json
    apl = conn.execute("SELECT prova_id FROM aplicacoes WHERE id = ?", (aplicacao_id,)).fetchone()
    if not apl:
        return (0, 0, 0.0)
    questoes = conn.execute("""
        SELECT q.id, q.tipo FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        WHERE pq.prova_id = ? ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()
    total_obj = 0
    acertos = 0
    for q in questoes:
        tipo = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
        if tipo == "discursiva":
            continue  # ignora discursivas no cálculo de proficiência
        total_obj += 1
        resp = conn.execute(
            "SELECT alternativa_letra, dados_extra FROM respostas WHERE aplicacao_id = ? AND aluno_id = ? AND questao_id = ?",
            (aplicacao_id, aluno_id, q["id"])
        ).fetchone()
        if not resp:
            continue
        if tipo == "multipla_escolha":
            ok = conn.execute(
                "SELECT 1 FROM alternativas WHERE questao_id = ? AND letra = ? AND correta = 1",
                (q["id"], resp["alternativa_letra"])
            ).fetchone()
            if ok:
                acertos += 1
        elif tipo == "vf":
            try:
                marcadas = _json.loads(resp["dados_extra"] or "{}")
            except Exception:
                marcadas = {}
            gabaritos = {str(a["ordem"]): a["gabarito"] for a in conn.execute(
                "SELECT ordem, gabarito FROM vf_afirmacoes WHERE questao_id = ?", (q["id"],)
            ).fetchall()}
            if gabaritos and all(marcadas.get(k) == v for k, v in gabaritos.items()):
                acertos += 1
        elif tipo == "associacao":
            try:
                marcadas = _json.loads(resp["dados_extra"] or "{}")
            except Exception:
                marcadas = {}
            gabaritos = {str(a["ordem"]): a["gabarito_letra"] for a in conn.execute(
                "SELECT ordem, gabarito_letra FROM assoc_itens_a WHERE questao_id = ?", (q["id"],)
            ).fetchall()}
            if gabaritos and all(marcadas.get(k) == v for k, v in gabaritos.items()):
                acertos += 1
    nota_10 = (acertos / total_obj * 10.0) if total_obj > 0 else 0.0
    return (acertos, total_obj, nota_10)


def _calcular_nota(conn, aplicacao_id, aluno_id):
    """Calcula (acertos, total) considerando todos os tipos.
    Regra: múltipla escolha = 0/1 (compara letra). V/F e Associação = 0/1 só se TODAS as
    afirmações/pares estiverem corretas. Discursiva entra no total mas conta 0 (correção manual)."""
    import json as _json
    apl = conn.execute("SELECT prova_id FROM aplicacoes WHERE id = ?", (aplicacao_id,)).fetchone()
    if not apl:
        return (0, 0)
    questoes = conn.execute("""
        SELECT q.id, q.tipo FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        WHERE pq.prova_id = ? ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()
    total = len(questoes)
    acertos = 0
    for q in questoes:
        tipo = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
        resp = conn.execute(
            "SELECT alternativa_letra, dados_extra FROM respostas WHERE aplicacao_id = ? AND aluno_id = ? AND questao_id = ?",
            (aplicacao_id, aluno_id, q["id"])
        ).fetchone()
        if not resp:
            continue
        if tipo == "multipla_escolha":
            ok = conn.execute(
                "SELECT 1 FROM alternativas WHERE questao_id = ? AND letra = ? AND correta = 1",
                (q["id"], resp["alternativa_letra"])
            ).fetchone()
            if ok:
                acertos += 1
        elif tipo == "vf":
            try:
                marcadas = _json.loads(resp["dados_extra"] or "{}")
            except Exception:
                marcadas = {}
            gabaritos = {str(a["ordem"]): a["gabarito"] for a in conn.execute(
                "SELECT ordem, gabarito FROM vf_afirmacoes WHERE questao_id = ?", (q["id"],)
            ).fetchall()}
            if gabaritos and all(marcadas.get(k) == v for k, v in gabaritos.items()):
                acertos += 1
        elif tipo == "associacao":
            try:
                marcadas = _json.loads(resp["dados_extra"] or "{}")
            except Exception:
                marcadas = {}
            gabaritos = {str(a["ordem"]): a["gabarito_letra"] for a in conn.execute(
                "SELECT ordem, gabarito_letra FROM assoc_itens_a WHERE questao_id = ?", (q["id"],)
            ).fetchall()}
            if gabaritos and all(marcadas.get(k) == v for k, v in gabaritos.items()):
                acertos += 1
        # discursiva: não soma acerto automático
    return (acertos, total)


def _pagina_simples(titulo, corpo_html):
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="color-scheme" content="light dark">
    <title>{titulo}</title>
    {INTER_FONT}
    {CSS_LINK}
</head>
<body>
    <main style="max-width: 600px; margin: 40px auto; padding: 0 20px;">
        <h1>{titulo}</h1>
        {corpo_html}
    </main>
</body>
</html>"""


def _pagina_aluno(titulo, conteudo_html):
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="color-scheme" content="light dark">
    <title>{titulo}</title>
    {INTER_FONT}
    {CSS_LINK}
    {MATHJAX}
</head>
<body>
    <main style="max-width: 800px; margin: 40px auto; padding: 0 20px;">
        {conteudo_html}
    </main>
</body>
</html>"""

@app.get("/aplicacoes/{aplicacao_id}/aluno/{aluno_id}", response_class=HTMLResponse)
def ver_respostas_aluno(aplicacao_id: int, aluno_id: int):
    conn = get_db()
    aluno = conn.execute("SELECT a.*, t.nome AS turma_nome FROM alunos a JOIN turmas t ON t.id = a.turma_id WHERE a.id = ?", (aluno_id,)).fetchone()
    apl = conn.execute("SELECT a.*, p.titulo AS prova_titulo FROM aplicacoes a JOIN provas p ON p.id = a.prova_id WHERE a.id = ?", (aplicacao_id,)).fetchone()

    if not aluno or not apl:
        conn.close()
        return HTMLResponse(render_page("Não encontrado", '<h1>Não encontrado</h1>', active="aplicacoes"), status_code=404)

    entrega = conn.execute("SELECT * FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id)).fetchone()
    questoes = conn.execute("""
        SELECT q.id, q.enunciado, d.nome AS disciplina_nome
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        JOIN disciplinas d ON d.id = q.disciplina_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()

    respostas = {r["questao_id"]: r["alternativa_letra"] for r in conn.execute("SELECT questao_id, alternativa_letra FROM respostas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id)).fetchall()}
    score, total = _calcular_nota(conn, aplicacao_id, aluno_id)

    questoes_html = ""
    for idx, q in enumerate(questoes, start=1):
        alts = conn.execute("SELECT letra, texto, correta FROM alternativas WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()
        marcada = respostas.get(q["id"])

        alts_html = ""
        for a in alts:
            estilo = "padding:10px 12px; border:1px solid var(--border); border-radius:6px; margin-bottom:6px;"
            label = ""
            if a["letra"] == marcada and a["correta"]:
                estilo += " background:var(--success-bg, var(--green-bg)); border-color:var(--success-text, var(--green));"
                label = ' <span style="color:var(--success-text, var(--green)); font-weight:600;">✓ Marcada (correta)</span>'
            elif a["letra"] == marcada and not a["correta"]:
                estilo += " background:var(--red-bg); border-color:var(--red);"
                label = ' <span style="color:var(--danger-text, var(--red)); font-weight:600;">✗ Marcada (incorreta)</span>'
            elif a["correta"]:
                estilo += " border-color:var(--success-text, var(--green)); border-style:dashed;"
                label = ' <span style="color:var(--success-text, var(--green)); font-weight:600;">← Resposta correta</span>'

            alts_html += f'<div style="{estilo}"><strong>{a["letra"]})</strong> {a["texto"]}{label}</div>'

        status_questao = ""
        if marcada is None:
            status_questao = '<div style="color:var(--text-muted); font-style:italic; margin-bottom:8px;">⚠ Aluno não respondeu esta questão</div>'

        questoes_html += f'<div class="question"><div class="question-header">Questão {idx} · {q["disciplina_nome"]}</div>{status_questao}<div class="enunciado">{q["enunciado"]}</div>{alts_html}</div>'

    conn.close()

    if entrega:
        status_badge = f'<span class="badge" style="background:var(--success-bg, var(--green-bg)); color:var(--success-text, var(--green));">Entregue · {score}/{total}</span>'
        status_data = f' · Entregue em {entrega["finalizada_em"]}'
    elif respostas:
        status_badge = f'<span class="badge">Em andamento · {score}/{total} (parcial)</span>'
        status_data = ""
    else:
        status_badge = '<span class="badge">Sem respostas</span>'
        status_data = ""

    content = f"""
        <div class="page-header">
            <h1>{aluno["nome"]}</h1>
            <p class="subtitle">{apl["prova_titulo"]} · Turma: {aluno["turma_nome"]} · Código: <code>{aluno["codigo_unico"]}</code>{status_data}</p>
            <div class="page-actions">{status_badge}</div>
        </div>
        {questoes_html}
        <p style="margin-top:24px;"><a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar pra aplicação</a></p>
    """
    return render_page(aluno["nome"], content, active="aplicacoes")

@app.get("/aplicacoes/{aplicacao_id}/exportar")
def exportar_resultados_excel(aplicacao_id: int):
    conn = get_db()
    
    # 1. Procurar os dados principais da aplicação, prova e turma
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo, t.id AS turma_id
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()

    if not apl:
        conn.close()
        return HTMLResponse("<h1>Aplicação não encontrada</h1>", status_code=404)

    # 2. Obter a lista de questões da prova (na ordem correta)
    questoes = conn.execute("""
        SELECT q.id, pq.ordem
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()
    
    total_questoes = len(questoes)

    # 3. Obter todos os alunos da turma
    alunos = conn.execute("SELECT * FROM alunos WHERE turma_id = ? ORDER BY numero, nome", (apl["turma_id"],)).fetchall()
    
    # 4. Obter o mapeamento de entregas finalizadas
    entregas = {row["aluno_id"]: row["finalizada_em"] for row in conn.execute("SELECT aluno_id, finalizada_em FROM entregas WHERE aplicacao_id = ?", (aplicacao_id,)).fetchall()}

    # Inicializar o workbook do openpyxl
    wb = Workbook()
    
    # Estilos profissionais (Cores e Fontes)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Azul Escuro
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Azul Claro
    white_bold_font = Font(name="Inter", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Inter", size=11, bold=True)
    regular_font = Font(name="Inter", size=11)
    title_font = Font(name="Inter", size=16, bold=True, color="1F4E78")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    # ABA 1: RESUMO GERAL
    ws_resumo = wb.active
    ws_resumo.title = "Resumo Geral"
    ws_resumo.views.sheetView[0].showGridLines = True
    
    ws_resumo.append(["Relatório de Desempenho da Avaliação"])
    ws_resumo["A1"].font = title_font
    ws_resumo.append([])
    
    dados_cabecalho = [
        ("Título da Aplicação:", apl["titulo"] or f'{apl["prova_titulo"]} — {apl["turma_nome"]}'),
        ("Prova Aplicada:", apl["prova_titulo"]),
        ("Turma / Ano Letivo:", f'{apl["turma_nome"]} ({apl["ano_letivo"]})'),
        ("Modo de Aplicação:", "Online" if apl["modo"] == "online" else "Impressa"),
        ("Data de Exportação:", datetime.now().strftime("%d/%m/%Y %H:%M"))
    ]
    
    for label, valor in dados_cabecalho:
        ws_resumo.append([label, valor])
        curr_row = ws_resumo.max_row
        ws_resumo[f"A{curr_row}"].font = bold_font
        ws_resumo[f"B{curr_row}"].font = regular_font
        
    ws_resumo.append([])
    
    # Calcular Métricas
    notas = []
    qtd_entregue = 0
    for a in alunos:
        if a["id"] in entregas:
            score, _ = _calcular_nota(conn, aplicacao_id, a["id"])
            notas.append(score)
            qtd_entregue += 1
            
    total_alunos = len(alunos)
    media_nota = sum(notas) / len(notas) if notas else 0
    porcentagem_media = (media_nota / total_questoes * 100) if total_questoes > 0 else 0

    ws_resumo.append(["Métricas Estatísticas da Turma"])
    ws_resumo.cell(row=ws_resumo.max_row, column=1).font = Font(name="Inter", size=13, bold=True, color="1F4E78")
    ws_resumo.append([])
    
    metricas = [
        ("Total de Alunos Inscritos", total_alunos),
        ("Provas Entregues", qtd_entregue),
        ("Aplicações Pendentes", total_alunos - qtd_entregue),
        ("Média de Acertos", f"{media_nota:.1f} / {total_questoes} ({porcentagem_media:.1f}%)")
    ]
    
    for label, valor in metricas:
        ws_resumo.append([label, valor])
        curr_row = ws_resumo.max_row
        ws_resumo[f"A{curr_row}"].font = bold_font
        ws_resumo[f"B{curr_row}"].font = regular_font
        ws_resumo[f"A{curr_row}"].fill = section_fill
        ws_resumo[f"A{curr_row}"].border = thin_border
        ws_resumo[f"B{curr_row}"].border = thin_border
        
    ws_resumo.column_dimensions['A'].width = 30
    ws_resumo.column_dimensions['B'].width = 45

    # ABA 2: NOTAS E RESPOSTAS DETALHADAS
    ws_detalhes = wb.create_sheet(title="Notas e Respostas")
    ws_detalhes.views.sheetView[0].showGridLines = True
    
    headers = ["Nº", "Nome do Aluno", "Código Único", "Status", "Acertos", "Total", "% Aproveitamento"]
    for idx, _ in enumerate(questoes, start=1):
        headers.append(f"Q{idx}")
        
    ws_detalhes.append(headers)
    
    for col_idx, cell in enumerate(ws_detalhes[1], start=1):
        cell.fill = header_fill
        cell.font = white_bold_font
        cell.alignment = center_align if col_idx != 2 else left_align
        cell.border = thin_border

    for row_idx, a in enumerate(alunos, start=2):
        num = a["numero"] if a["numero"] else "—"
        status = "Entregue" if a["id"] in entregas else "Pendente"
        
        if a["id"] in entregas:
            score, _ = _calcular_nota(conn, aplicacao_id, a["id"])
            porcentagem = (score / total_questoes) * 100 if total_questoes > 0 else 0
            porcentagem_str = f"{porcentagem:.1f}%"
        else:
            score = "—"
            porcentagem_str = "—"
            
        row_data = [num, a["nome"], a["codigo_unico"], status, score, total_questoes, porcentagem_str]
        
        respostas_aluno = {r["questao_id"]: r["alternativa_letra"] for r in conn.execute(
            "SELECT questao_id, alternativa_letra FROM respostas WHERE aplicacao_id = ? AND aluno_id = ?", 
            (aplicacao_id, a["id"])
        ).fetchall()}
        
        for q in questoes:
            resp = respostas_aluno.get(q["id"])
            if resp:
                is_correct = conn.execute(
                    "SELECT correta FROM alternativas WHERE questao_id = ? AND letra = ?", 
                    (q["id"], resp)
                ).fetchone()
                
                marca = "✓" if (is_correct and is_correct["correta"] == 1) else "✗"
                row_data.append(f"{resp} ({marca})")
            else:
                row_data.append("—")
                
        ws_detalhes.append(row_data)
        
        for col_idx, cell in enumerate(ws_detalhes[row_idx], start=1):
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = left_align if col_idx == 2 else center_align
            
            if col_idx > 7 and "(" in str(cell.value):
                if "✓" in str(cell.value):
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Verde Claro
                else:
                    cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Vermelho Claro

    for col in ws_detalhes.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_detalhes.column_dimensions[col_letter].width = max(max_len + 3, 10)

    conn.close()

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    safe_title = (apl["titulo"] or f"resultados_aplicacao_{aplicacao_id}").lower().replace(" ", "_")
    filename = f"resultados_{safe_title}.xlsx"
    
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )    

# ==========================================
#  EDITAR E EXCLUIR QUESTÕES
# ==========================================

@app.get("/questoes/{id}/editar", response_class=HTMLResponse)
def form_editar_questao(id: int, request: Request):
    prof = get_current_professor(request)
    conn = get_db()
    q = conn.execute("SELECT * FROM questoes WHERE id = ?", (id,)).fetchone()
    if not q:
        conn.close()
        return RedirectResponse("/questoes", status_code=303)
    if not _pode_editar_questao(prof, q["criada_por_professor_id"]):
        conn.close()
        return HTMLResponse(render_page(
            "Sem permissão",
            '<div class="page-header"><h1>🔒 Sem permissão</h1></div>'
            '<div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:16px; border-radius:6px;">'
            '<p>Essa questão foi criada por outro professor. Apenas <strong>o autor da questão</strong> ou o <strong>administrador</strong> podem editá-la.</p>'
            '<p>Você pode <strong>visualizar</strong> a questão e <strong>usá-la em suas próprias provas</strong> normalmente.</p>'
            '</div>'
            '<div class="page-actions" style="margin-top:14px;"><a href="/questoes" class="btn">← Voltar ao banco</a></div>',
            active="questoes"
        ), status_code=403)

    disciplinas = conn.execute("SELECT * FROM disciplinas ORDER BY nome").fetchall()
    alts = conn.execute("SELECT letra, texto, correta FROM alternativas WHERE questao_id = ? ORDER BY letra", (id,)).fetchall()
    vf_afirms = conn.execute("SELECT ordem, texto, gabarito FROM vf_afirmacoes WHERE questao_id = ? ORDER BY ordem", (id,)).fetchall()
    assoc_a = conn.execute("SELECT ordem, texto, gabarito_letra FROM assoc_itens_a WHERE questao_id = ? ORDER BY ordem", (id,)).fetchall()
    assoc_b = conn.execute("SELECT letra, texto FROM assoc_itens_b WHERE questao_id = ? ORDER BY letra", (id,)).fetchall()
    textos = conn.execute("SELECT id, conteudo, fonte FROM textos_apoio WHERE questao_id = ? ORDER BY ordem", (id,)).fetchall()
    imagens = conn.execute("SELECT id, caminho, legenda, fonte FROM imagens WHERE questao_id = ? ORDER BY ordem", (id,)).fetchall()
    habilidades = conn.execute("SELECT h.codigo FROM questao_habilidades qh JOIN habilidades_bncc h ON h.id = qh.habilidade_id WHERE qh.questao_id = ? ORDER BY h.codigo", (id,)).fetchall()
    habs_existentes = conn.execute("SELECT codigo FROM habilidades_bncc ORDER BY codigo").fetchall()
    conn.close()

    options = "".join(
        f'<option value="{d["id"]}"{(" selected" if d["id"] == q["disciplina_id"] else "")}>{d["nome"]}</option>'
        for d in disciplinas
    )
    ano_atual = q["ano"] if "ano" in q.keys() else None
    anos_options = '<option value="">— Não definido —</option>' + "".join(
        f'<option value="{a}"{(" selected" if ano_atual == a else "")}>{a}</option>' for a in ANOS
    )

    habs_preset = ", ".join(h["codigo"] for h in habilidades)

    total_habs = len(habs_existentes)
    link_catalogo = (
        f'<p class="muted-line" style="font-size:11px;">'
        f'💡 {total_habs} habilidade(s) no catálogo. '
        f'<a href="/habilidades" target="_blank" style="color:var(--text-muted);">Consultar lista</a>'
        f'</p>'
    ) if total_habs > 0 else ''

    alts_by_letra = {a["letra"]: a for a in alts}
    alternativas_html = ""
    for letra in ["A", "B", "C", "D"]:
        a = alts_by_letra.get(letra)
        valor_alt = a["texto"] if a else ""
        checked = ' checked' if a and a["correta"] else ''
        required_radio = ' required' if letra == "A" else ''
        editor_alt = _editor_enunciado_html(
            name=f"alt_{letra.lower()}", valor_inicial=valor_alt, required=True,
            label="", compact=True, min_height=42,
            placeholder=f"Texto da alternativa {letra}"
        )
        alternativas_html += (
            f'<div style="display:grid; grid-template-columns:auto 1fr; gap:12px; align-items:flex-start; margin-bottom:10px;">'
            f'<label style="margin:8px 0 0 0; display:flex; align-items:center; gap:8px; white-space:nowrap;">'
            f'<input type="radio" name="correta" value="{letra}"{required_radio}{checked} style="width:auto; margin:0;"> <strong>{letra})</strong>'
            f'</label>'
            f'<div style="margin:0;">{editor_alt}</div>'
            f'</div>'
        )

    textos_existentes_html = ""
    if textos:
        items = ""
        for t in textos:
            fonte_part = f' <small style="color:var(--text-muted);">({t["fonte"]})</small>' if t["fonte"] else ""
            items += (
                f'<div style="display:flex; gap:12px; align-items:flex-start; padding:10px; border:1px solid var(--border); border-radius:6px; margin-bottom:6px;">'
                f'<div style="flex:1;">{t["conteudo"]}{fonte_part}</div>'
                f'<form action="/textos_apoio/{t["id"]}/deletar" method="post" style="margin:0;" '
                f'onsubmit="return confirm(\'Remover este texto de apoio?\');">'
                f'<button type="submit" class="btn" style="padding:4px 10px; font-size:12px;">Remover</button>'
                f'</form></div>'
            )
        textos_existentes_html = f'<h3>Textos de apoio existentes</h3>{items}'

    imagens_existentes_html = ""
    if imagens:
        items = ""
        for img in imagens:
            legenda_html = f'<div style="font-size:12px; color:var(--text-muted);">{img["legenda"]}</div>' if img["legenda"] else ""
            fonte_html = f'<div style="font-size:11px; color:var(--text-subtle);">Fonte: {img["fonte"]}</div>' if img["fonte"] else ""
            items += (
                f'<div style="display:flex; gap:12px; align-items:flex-start; padding:10px; border:1px solid var(--border); border-radius:6px; margin-bottom:6px;">'
                f'<img src="/{img["caminho"]}" alt="" style="max-width:120px; max-height:120px; border-radius:4px;">'
                f'<div style="flex:1;">{legenda_html}{fonte_html}</div>'
                f'<form action="/imagens/{img["id"]}/deletar" method="post" style="margin:0;" '
                f'onsubmit="return confirm(\'Remover esta imagem? O arquivo será apagado do servidor.\');">'
                f'<button type="submit" class="btn" style="padding:4px 10px; font-size:12px;">Remover</button>'
                f'</form></div>'
            )
        imagens_existentes_html = f'<h3>Imagens existentes</h3>{items}'

    enunciado_safe = q["enunciado"]
    tipo_q = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
    if tipo_q not in TIPOS_QUESTAO:
        tipo_q = "multipla_escolha"
    tipo_info = TIPOS_QUESTAO[tipo_q]

    # Bloco de alternativas / afirmações / pares — conforme o tipo
    fieldset_alts = ""
    if tipo_q == "multipla_escolha":
        fieldset_alts = f"""
            <fieldset>
                <legend>Alternativas — marque o radio da correta</legend>
                {alternativas_html}
            </fieldset>
        """
    elif tipo_q == "discursiva":
        fieldset_alts = """
            <div style="background:var(--accent-bg); color:var(--accent); border:1px solid var(--accent); padding:14px 16px; border-radius:6px; margin:12px 0;">
                <strong>📝 Questão discursiva</strong> — resposta livre, correção manual.
            </div>
        """
    elif tipo_q == "vf":
        afirms_dict = {af["ordem"]: af for af in vf_afirms}
        afirms_html_edit = ""
        for i in range(VF_MAX_AFIRMACOES):
            af = afirms_dict.get(i)
            valor = af["texto"] if af else ""
            gab = af["gabarito"] if af else None
            editor = _editor_enunciado_html(
                name=f"vf_afirm_{i}_texto", valor_inicial=valor, required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Afirmação {i+1} (deixe em branco se não usar)"
            )
            ck_v = " checked" if gab == "V" else ""
            ck_f = " checked" if gab == "F" else ""
            afirms_html_edit += (
                f'<div style="display:grid; grid-template-columns:1fr auto; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<div style="margin:0;"><strong style="font-size:13px;">Afirmação {i+1}</strong>{editor}</div>'
                f'<div style="display:flex; gap:10px; align-items:center; padding-top:24px; white-space:nowrap;">'
                f'<label style="margin:0; font-size:13px;"><input type="radio" name="vf_afirm_{i}_gabarito" value="V"{ck_v} style="width:auto; margin:0 4px 0 0;">V</label>'
                f'<label style="margin:0; font-size:13px;"><input type="radio" name="vf_afirm_{i}_gabarito" value="F"{ck_f} style="width:auto; margin:0 4px 0 0;">F</label>'
                f'</div></div>'
            )
        fieldset_alts = f"""
            <fieldset>
                <legend>Afirmações — marque V ou F (até {VF_MAX_AFIRMACOES})</legend>
                {afirms_html_edit}
            </fieldset>
        """
    elif tipo_q == "associacao":
        assoc_a_dict = {a["ordem"]: a for a in assoc_a}
        assoc_b_dict = {b["letra"]: b for b in assoc_b}
        col_a_html_edit = ""
        for i in range(ASSOC_MAX_PARES):
            item_a = assoc_a_dict.get(i)
            val_a = item_a["texto"] if item_a else ""
            gab_a = item_a["gabarito_letra"] if item_a else ""
            editor_a = _editor_enunciado_html(
                name=f"assoc_a_{i}_texto", valor_inicial=val_a, required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Item {i+1} da coluna A"
            )
            letras_opts = '<option value="">—</option>' + "".join(
                f'<option value="{chr(97+j)}"{" selected" if gab_a == chr(97+j) else ""}>{chr(97+j)}</option>'
                for j in range(ASSOC_MAX_PARES)
            )
            col_a_html_edit += (
                f'<div style="display:grid; grid-template-columns:auto 1fr auto; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<strong style="padding-top:20px;">{i+1}.</strong>'
                f'<div style="margin:0;">{editor_a}</div>'
                f'<label style="margin:0; padding-top:14px; font-size:12px; white-space:nowrap;">Resposta: '
                f'<select name="assoc_a_{i}_gabarito" style="width:auto; display:inline-block; margin-left:4px;">{letras_opts}</select>'
                f'</label></div>'
            )
        col_b_html_edit = ""
        for j in range(ASSOC_MAX_PARES):
            letra_b = chr(97+j)
            item_b = assoc_b_dict.get(letra_b)
            val_b = item_b["texto"] if item_b else ""
            editor_b = _editor_enunciado_html(
                name=f"assoc_b_{letra_b}_texto", valor_inicial=val_b, required=False,
                label="", compact=True, min_height=42,
                placeholder=f"Item ({letra_b}) da coluna B"
            )
            col_b_html_edit += (
                f'<div style="display:grid; grid-template-columns:auto 1fr; gap:12px; align-items:flex-start; margin-bottom:10px;">'
                f'<strong style="padding-top:20px;">({letra_b})</strong>'
                f'<div style="margin:0;">{editor_b}</div>'
                f'</div>'
            )
        fieldset_alts = f"""
            <fieldset>
                <legend>Coluna A — itens (1, 2...) com gabarito</legend>
                {col_a_html_edit}
            </fieldset>
            <fieldset>
                <legend>Coluna B — opções (a, b...)</legend>
                {col_b_html_edit}
            </fieldset>
        """

    content = f"""
        <div class="page-header"><h1>Editar questão</h1>
            <p class="subtitle">{tipo_info['icone']} {tipo_info['label']} <span style="color:var(--text-muted); font-size:12px;">(o tipo não pode ser alterado depois da criação)</span></p>
        </div>
        <div class="tip"><strong>Dica:</strong> use <code>$fórmula$</code> para fórmulas inline ou <code>$$fórmula$$</code> para centralizadas. Textos e imagens existentes podem ser removidos individualmente acima; os campos "Adicionar novos" só inserem novos itens.</div>

        {textos_existentes_html}
        {imagens_existentes_html}

        <form action="/questoes/{id}/editar" method="post" enctype="multipart/form-data">
            <input type="hidden" name="tipo" value="{tipo_q}">
            <div style="display:grid; grid-template-columns: 2fr 1fr; gap:12px;">
                <label>Disciplina<select name="disciplina_id" required>{options}</select></label>
                <label>Ano de escolaridade<select name="ano">{anos_options}</select></label>
            </div>
            <div id="bncc-container" style="margin:10px 0;">
                <label style="margin-bottom:6px;">Habilidades BNCC <span style="font-weight:400; color:var(--text-muted); font-size:12px;">(opcional)</span></label>
                <input type="hidden" name="habilidades_codigos" id="bncc-hidden" value="{habs_preset}">
                <div id="bncc-chips" style="display:flex; flex-wrap:wrap; gap:6px; min-height:24px; margin-bottom:8px;"></div>
                <input type="search" id="bncc-search" placeholder="Digite o código (EF09MA09) ou palavra-chave (fração, célula...)" style="margin:0;">
                <div id="bncc-results" style="margin-top:6px;"></div>
            </div>
            {link_catalogo}

            <fieldset>
                <legend>Adicionar novos textos de apoio (opcional)</legend>
                {_editor_enunciado_html(name="texto1_conteudo", valor_inicial="", required=False, label="Texto novo — conteúdo", min_height=80, placeholder="Cole ou digite o texto de apoio")}
                <label>Texto novo — fonte<input type="text" name="texto1_fonte" placeholder="Autor, obra, ano"></label>
                {_editor_enunciado_html(name="texto2_conteudo", valor_inicial="", required=False, label="Outro texto novo — conteúdo", min_height=80, placeholder="Segundo texto de apoio (opcional)")}
                <label>Outro texto novo — fonte<input type="text" name="texto2_fonte"></label>
            </fieldset>

            <fieldset>
                <legend>Adicionar novas imagens (opcional)</legend>
                <label>Imagem nova<input type="file" name="imagem1" accept="image/*"></label>
                <label>Legenda<input type="text" name="imagem1_legenda"></label>
                <label>Fonte<input type="text" name="imagem1_fonte"></label>
                <label>Outra imagem nova<input type="file" name="imagem2" accept="image/*"></label>
                <label>Legenda<input type="text" name="imagem2_legenda"></label>
                <label>Fonte<input type="text" name="imagem2_fonte"></label>
            </fieldset>

            {_editor_enunciado_html(name="enunciado", valor_inicial=enunciado_safe or "", required=True, label="Enunciado", placeholder="Digite o enunciado da questão.", detectar_alternativas=(tipo_q == "multipla_escolha"))}

            {fieldset_alts}

            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Salvar alterações</button>
                <a href="/questoes" class="btn">Cancelar</a>
            </div>
        </form>
        <script>
        (function() {{
            const ta = document.querySelector('textarea[name="habilidades_codigos"]');
            const discSel = document.querySelector('select[name="disciplina_id"]');
            if (!ta) return;

            // === Painel de validação ===
            const preview = document.createElement('div');
            preview.id = 'bncc-preview';
            preview.style.cssText = 'margin-top:6px; font-size:12px; line-height:1.5;';
            ta.parentNode.appendChild(preview);

            async function validar() {{
                const codigos = ta.value.split(/[,\\n]/).map(c => c.trim().toUpperCase()).filter(c => c);
                if (codigos.length === 0) {{ preview.innerHTML = ''; return; }}
                try {{
                    const resp = await fetch('/habilidades/buscar?codigos=' + encodeURIComponent(codigos.join(',')));
                    const data = await resp.json();
                    let html = '';
                    for (const c of codigos) {{
                        if (data[c]) {{
                            html += '<div style="padding:4px 8px; background:var(--green-bg); border-left:3px solid var(--green); margin-bottom:3px; color:var(--text);"><strong style="color:var(--green);">' + c + '</strong>: ' + data[c].replace(/</g, '&lt;') + '</div>';
                        }} else {{
                            html += '<div style="padding:4px 8px; background:var(--red-bg); border-left:3px solid var(--red); margin-bottom:3px; color:var(--red);"><strong style="color:var(--red);">' + c + '</strong>: ⚠ código não encontrado no catálogo</div>';
                        }}
                    }}
                    preview.innerHTML = html;
                }} catch (e) {{ preview.innerHTML = ''; }}
            }}
            ta.addEventListener('blur', validar);
            ta.addEventListener('input', () => {{ if (ta._t) clearTimeout(ta._t); ta._t = setTimeout(validar, 600); }});
            if (ta.value.trim()) validar();

            // === Busca por palavra/conceito ===
            const buscaWrap = document.createElement('div');
            buscaWrap.style.cssText = 'margin-top:14px; padding:12px; background:var(--bg-subtle); border-radius:6px;';
            buscaWrap.innerHTML = '<label style="margin:0; font-size:12px;">🔍 Não sabe o código? Busque por palavra ou conceito<input type="search" id="bncc-busca" placeholder="ex: fração, Constituição, fotossíntese" style="margin-top:4px;"></label><div id="bncc-resultados" style="margin-top:8px; font-size:12px;"></div>';
            ta.parentNode.appendChild(buscaWrap);

            const inputBusca = buscaWrap.querySelector('#bncc-busca');
            const divRes = buscaWrap.querySelector('#bncc-resultados');

            async function buscarPorPalavra() {{
                const q = inputBusca.value.trim();
                if (q.length < 2) {{ divRes.innerHTML = ''; return; }}
                const disc = discSel ? discSel.value : '';
                const url = '/habilidades/buscar?q=' + encodeURIComponent(q) + (disc ? '&disciplina_id=' + disc : '');
                try {{
                    const resp = await fetch(url);
                    const data = await resp.json();
                    const results = data.results || [];
                    if (results.length === 0) {{
                        divRes.innerHTML = '<div style="color:var(--text-muted); padding:8px 0;">Nenhum resultado para "' + q.replace(/</g, '&lt;') + '"' + (disc ? ' na disciplina selecionada' : '') + '.</div>';
                        return;
                    }}
                    const escopo = disc ? ' (filtrado pela disciplina)' : ' (todas as disciplinas)';
                    let html = '<div style="color:var(--text-muted); padding:4px 0;">' + results.length + ' habilidade(s) encontrada(s)' + escopo + ' — clique para adicionar:</div>';
                    for (const r of results) {{
                        html += '<div data-codigo="' + r.codigo + '" style="padding:6px 8px; border:1px solid var(--border); border-radius:4px; margin-bottom:4px; cursor:pointer; background:var(--bg); color:var(--text);" onmouseover="this.style.background=\\'var(--accent-bg)\\'" onmouseout="this.style.background=\\'var(--bg)\\'"><strong style="color:var(--accent);">' + r.codigo + '</strong> · ' + r.descricao.replace(/</g, '&lt;') + '</div>';
                    }}
                    divRes.innerHTML = html;
                }} catch (e) {{ divRes.innerHTML = ''; }}
            }}
            inputBusca.addEventListener('input', () => {{ if (inputBusca._t) clearTimeout(inputBusca._t); inputBusca._t = setTimeout(buscarPorPalavra, 400); }});
            if (discSel) discSel.addEventListener('change', buscarPorPalavra);

            divRes.addEventListener('click', (e) => {{
                const item = e.target.closest('[data-codigo]');
                if (!item) return;
                const codigo = item.dataset.codigo;
                const cur = ta.value.trim();
                const codigos = cur ? cur.split(/[,\\n]/).map(c => c.trim().toUpperCase()).filter(c => c) : [];
                if (codigos.includes(codigo)) return;
                codigos.push(codigo);
                ta.value = codigos.join(', ');
                validar();
            }});
        }})();
        </script>
    """
    return render_page("Editar questão", content, active="questoes", head_extra=MATHJAX)


@app.post("/questoes/{id}/editar", response_class=HTMLResponse)
async def atualizar_questao(
    id: int,
    request: Request,
    disciplina_id: int = Form(...),
    enunciado: str = Form(...),
    tipo: str = Form("multipla_escolha"),
    alt_a: str = Form(""), alt_b: str = Form(""), alt_c: str = Form(""), alt_d: str = Form(""),
    correta: str = Form(""),
    habilidades_codigos: str = Form(""),
    ano: str = Form(""),
    texto1_conteudo: str = Form(""), texto1_fonte: str = Form(""),
    texto2_conteudo: str = Form(""), texto2_fonte: str = Form(""),
    imagem1: Optional[UploadFile] = File(None), imagem1_legenda: str = Form(""), imagem1_fonte: str = Form(""),
    imagem2: Optional[UploadFile] = File(None), imagem2_legenda: str = Form(""), imagem2_fonte: str = Form(""),
):
    prof = get_current_professor(request)
    conn = get_db()
    q_existente = conn.execute("SELECT criada_por_professor_id, tipo FROM questoes WHERE id = ?", (id,)).fetchone()
    if not q_existente:
        conn.close()
        return RedirectResponse("/questoes", status_code=303)
    if not _pode_editar_questao(prof, q_existente["criada_por_professor_id"]):
        conn.close()
        return HTMLResponse(render_page(
            "Sem permissão",
            '<div class="page-header"><h1>🔒 Sem permissão</h1></div>'
            '<div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:16px; border-radius:6px;">'
            '<p>Apenas o autor da questão ou o administrador podem editá-la.</p></div>'
            '<div class="page-actions" style="margin-top:14px;"><a href="/questoes" class="btn">← Voltar</a></div>',
            active="questoes"
        ), status_code=403)
    # Tipo NÃO pode mudar depois de criada (preserva integridade de respostas históricas)
    tipo_atual = q_existente["tipo"] if "tipo" in q_existente.keys() and q_existente["tipo"] else "multipla_escolha"
    conn.execute("UPDATE questoes SET disciplina_id = ?, enunciado = ?, ano = ? WHERE id = ?", (disciplina_id, _sanitizar_html_enunciado(enunciado), ano.strip() or None, id))

    # Recarrega o form pra pegar campos dinâmicos
    form_extra = await request.form()

    if tipo_atual == "multipla_escolha":
        conn.execute("DELETE FROM alternativas WHERE questao_id = ?", (id,))
        for letra, texto in [("A", alt_a), ("B", alt_b), ("C", alt_c), ("D", alt_d)]:
            conn.execute("INSERT INTO alternativas (questao_id, letra, texto, correta) VALUES (?, ?, ?, ?)",
                         (id, letra, _sanitizar_html_enunciado(texto), 1 if letra == correta else 0))
    elif tipo_atual == "vf":
        conn.execute("DELETE FROM vf_afirmacoes WHERE questao_id = ?", (id,))
        ordem_real = 0
        for i in range(VF_MAX_AFIRMACOES):
            texto_afirm = _sanitizar_html_enunciado(str(form_extra.get(f"vf_afirm_{i}_texto", "")))
            gabarito = str(form_extra.get(f"vf_afirm_{i}_gabarito", "")).strip().upper()
            if texto_afirm and gabarito in ("V", "F"):
                conn.execute("INSERT INTO vf_afirmacoes (questao_id, ordem, texto, gabarito) VALUES (?, ?, ?, ?)",
                             (id, ordem_real, texto_afirm, gabarito))
                ordem_real += 1
    elif tipo_atual == "associacao":
        conn.execute("DELETE FROM assoc_itens_a WHERE questao_id = ?", (id,))
        conn.execute("DELETE FROM assoc_itens_b WHERE questao_id = ?", (id,))
        ordem_real = 0
        for i in range(ASSOC_MAX_PARES):
            texto_a = _sanitizar_html_enunciado(str(form_extra.get(f"assoc_a_{i}_texto", "")))
            gabarito = str(form_extra.get(f"assoc_a_{i}_gabarito", "")).strip().lower()
            if texto_a and gabarito:
                conn.execute("INSERT INTO assoc_itens_a (questao_id, ordem, texto, gabarito_letra) VALUES (?, ?, ?, ?)",
                             (id, ordem_real, texto_a, gabarito))
                ordem_real += 1
        for j in range(ASSOC_MAX_PARES):
            letra_b = chr(97+j)
            texto_b = _sanitizar_html_enunciado(str(form_extra.get(f"assoc_b_{letra_b}_texto", "")))
            if texto_b:
                conn.execute("INSERT INTO assoc_itens_b (questao_id, letra, texto) VALUES (?, ?, ?)",
                             (id, letra_b, texto_b))

    conn.execute("DELETE FROM questao_habilidades WHERE questao_id = ?", (id,))
    for parte in habilidades_codigos.replace("\n", ",").split(","):
        codigo = parte.strip().upper()
        if not codigo:
            continue
        existing = conn.execute("SELECT id FROM habilidades_bncc WHERE codigo = ?", (codigo,)).fetchone()
        habilidade_id = existing["id"] if existing else conn.execute("INSERT INTO habilidades_bncc (codigo) VALUES (?)", (codigo,)).lastrowid
        try:
            conn.execute("INSERT INTO questao_habilidades (questao_id, habilidade_id) VALUES (?, ?)", (id, habilidade_id))
        except sqlite3.IntegrityError:
            pass

    proximo_ordem_texto = conn.execute("SELECT COALESCE(MAX(ordem), -1) + 1 AS n FROM textos_apoio WHERE questao_id = ?", (id,)).fetchone()["n"]
    for offset, (conteudo, fonte) in enumerate([(texto1_conteudo, texto1_fonte), (texto2_conteudo, texto2_fonte)]):
        conteudo_sanit = _sanitizar_html_enunciado(conteudo)
        if conteudo_sanit:
            conn.execute("INSERT INTO textos_apoio (questao_id, conteudo, fonte, ordem) VALUES (?, ?, ?, ?)",
                         (id, conteudo_sanit, fonte.strip() or None, proximo_ordem_texto + offset))

    proximo_ordem_img = conn.execute("SELECT COALESCE(MAX(ordem), -1) + 1 AS n FROM imagens WHERE questao_id = ?", (id,)).fetchone()["n"]
    for offset, (img, legenda, fonte) in enumerate([(imagem1, imagem1_legenda, imagem1_fonte), (imagem2, imagem2_legenda, imagem2_fonte)]):
        if img and img.filename:
            content_bytes = await img.read()
            content_bytes = _redimensionar_imagem(content_bytes, max_width=800)
            unique_name = f"{uuid.uuid4().hex}.jpg"
            file_path = os.path.join(UPLOAD_DIR, unique_name)
            with open(file_path, "wb") as f:
                f.write(content_bytes)
            conn.execute("INSERT INTO imagens (questao_id, caminho, legenda, fonte, ordem) VALUES (?, ?, ?, ?, ?)",
                         (id, f"static/imagens/{unique_name}", legenda.strip() or None, fonte.strip() or None, proximo_ordem_img + offset))

    conn.commit()
    conn.close()
    return RedirectResponse("/questoes", status_code=303)


@app.post("/questoes/{id}/deletar", response_class=HTMLResponse)
def deletar_questao(id: int, request: Request):
    prof = get_current_professor(request)
    conn = get_db()
    q_existente = conn.execute("SELECT criada_por_professor_id FROM questoes WHERE id = ?", (id,)).fetchone()
    if not q_existente:
        conn.close()
        return RedirectResponse("/questoes", status_code=303)
    if not _pode_editar_questao(prof, q_existente["criada_por_professor_id"]):
        conn.close()
        return HTMLResponse(render_page(
            "Sem permissão",
            '<div class="page-header"><h1>🔒 Sem permissão</h1></div>'
            '<div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:16px; border-radius:6px;">'
            '<p>Apenas o autor da questão ou o administrador podem excluí-la.</p></div>'
            '<div class="page-actions" style="margin-top:14px;"><a href="/questoes" class="btn">← Voltar</a></div>',
            active="questoes"
        ), status_code=403)
    em_uso = conn.execute("SELECT COUNT(*) AS c FROM prova_questoes WHERE questao_id = ?", (id,)).fetchone()["c"]
    if em_uso > 0:
        conn.close()
        content = """
        <div style="border: 1px solid var(--red); background: var(--red-bg); padding: 20px; border-radius: 6px; margin-top:20px; color:var(--red);">
            <h3 style="color:var(--red); margin-top:0;">Operação Impedida</h3>
            <p>Não é possível excluir esta questão porque ela está sendo usada em uma ou mais <strong>provas</strong>.</p>
            <p>Se deseja realmente excluí-la, remova-a primeiro das provas que a usam (na tela de edição de prova).</p>
            <a href="/questoes" class="btn" style="margin-top:10px;">Voltar para Questões</a>
        </div>
        """
        return render_page("Erro ao Excluir Questão", content, active="questoes")

    imagens = conn.execute("SELECT caminho FROM imagens WHERE questao_id = ?", (id,)).fetchall()
    for img in imagens:
        try:
            if os.path.exists(img["caminho"]):
                os.remove(img["caminho"])
        except Exception:
            pass

    conn.execute("DELETE FROM alternativas WHERE questao_id = ?", (id,))
    conn.execute("DELETE FROM textos_apoio WHERE questao_id = ?", (id,))
    conn.execute("DELETE FROM imagens WHERE questao_id = ?", (id,))
    conn.execute("DELETE FROM questao_habilidades WHERE questao_id = ?", (id,))
    conn.execute("DELETE FROM questoes WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/questoes", status_code=303)


@app.post("/textos_apoio/{id}/deletar")
def deletar_texto_apoio(id: int):
    conn = get_db()
    row = conn.execute("SELECT questao_id FROM textos_apoio WHERE id = ?", (id,)).fetchone()
    if not row:
        conn.close()
        return RedirectResponse("/questoes", status_code=303)
    questao_id = row["questao_id"]
    conn.execute("DELETE FROM textos_apoio WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/questoes/{questao_id}/editar", status_code=303)


@app.post("/imagens/{id}/deletar")
def deletar_imagem(id: int):
    conn = get_db()
    row = conn.execute("SELECT questao_id, caminho FROM imagens WHERE id = ?", (id,)).fetchone()
    if not row:
        conn.close()
        return RedirectResponse("/questoes", status_code=303)
    questao_id = row["questao_id"]
    caminho = row["caminho"]

    try:
        if os.path.exists(caminho):
            os.remove(caminho)
    except Exception:
        pass

    conn.execute("DELETE FROM imagens WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/questoes/{questao_id}/editar", status_code=303)


# ==========================================
#  ANÁLISES PEDAGÓGICAS (FASE B)
# ==========================================

def _estatisticas_questao(conn, aplicacao_id, questao_id, alunos_entregues):
    """Retorna estatísticas de uma questão entre os alunos que entregaram.
    alunos_entregues: set/list de aluno_ids que finalizaram a aplicação."""
    if not alunos_entregues:
        return {"total": 0, "acertos": 0, "distribuicao": {"A": 0, "B": 0, "C": 0, "D": 0},
                "em_branco": 0, "correta_letra": None, "pct_acerto": 0.0}

    correta_row = conn.execute(
        "SELECT letra FROM alternativas WHERE questao_id = ? AND correta = 1",
        (questao_id,)
    ).fetchone()
    correta_letra = correta_row["letra"] if correta_row else None

    placeholders = ",".join("?" * len(alunos_entregues))
    respostas = conn.execute(
        f"SELECT alternativa_letra FROM respostas WHERE aplicacao_id = ? AND questao_id = ? AND aluno_id IN ({placeholders})",
        (aplicacao_id, questao_id, *list(alunos_entregues))
    ).fetchall()

    distribuicao = {"A": 0, "B": 0, "C": 0, "D": 0}
    acertos = 0
    respondidas = 0
    for r in respostas:
        letra = r["alternativa_letra"]
        if letra in distribuicao:
            distribuicao[letra] += 1
            respondidas += 1
        if letra == correta_letra:
            acertos += 1

    total = len(alunos_entregues)
    em_branco = total - respondidas
    pct = (acertos / total) * 100 if total > 0 else 0.0

    return {
        "total": total, "acertos": acertos, "distribuicao": distribuicao,
        "em_branco": em_branco, "correta_letra": correta_letra, "pct_acerto": pct,
    }


def _cor_por_pct(pct):
    """Retorna cor de fundo conforme % de acerto (verde > 70%, amarelo 40-70%, vermelho < 40%)."""
    if pct >= 70:
        return "var(--green)"  # verde
    elif pct >= 40:
        return "var(--orange)"  # amarelo/laranja
    else:
        return "var(--red)"  # vermelho


def _barra_html(pct, largura_total=200):
    """Gera HTML de barra horizontal com cor adaptativa."""
    cor = _cor_por_pct(pct)
    return (
        f'<div style="display:inline-flex; align-items:center; gap:8px; width:{largura_total + 60}px;">'
        f'<div style="background:var(--bg-muted); border-radius:4px; height:18px; width:{largura_total}px; overflow:hidden; position:relative;">'
        f'<div style="background:{cor}; height:100%; width:{pct}%; transition:width 0.3s;"></div>'
        f'</div>'
        f'<span style="font-weight:600; font-size:13px; min-width:50px;">{pct:.1f}%</span>'
        f'</div>'
    )


@app.get("/aplicacoes/{aplicacao_id}/analise", response_class=HTMLResponse)
def analise_aplicacao(aplicacao_id: int):
    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a
        JOIN provas p ON p.id = a.prova_id
        JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()

    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    questoes = conn.execute("""
        SELECT q.id, q.enunciado, d.nome AS disciplina_nome
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        JOIN disciplinas d ON d.id = q.disciplina_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (apl["prova_id"],)).fetchall()
    total_questoes = len(questoes)

    entregas = conn.execute("SELECT aluno_id FROM entregas WHERE aplicacao_id = ?", (aplicacao_id,)).fetchall()
    alunos_entregues = [e["aluno_id"] for e in entregas]
    total_entregas = len(alunos_entregues)

    if total_entregas == 0:
        conn.close()
        content = f"""
            <div class="page-header">
                <h1>Análise pedagógica</h1>
                <p class="subtitle">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]})</p>
            </div>
            <div class="empty">
                <p>Nenhum aluno finalizou a prova ainda. A análise pedagógica fica disponível após pelo menos uma entrega.</p>
                <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar para a aplicação</a>
            </div>
        """
        return render_page("Análise pedagógica", content, active="aplicacoes")

    notas_alunos = []
    for aluno_id in alunos_entregues:
        score, total = _calcular_nota(conn, aplicacao_id, aluno_id)
        # Variante objetiva (sem discursivas) pra calcular nota 10 + faixa SAEB
        acertos_obj, total_obj, nota_10 = _calcular_nota_objetiva(conn, aplicacao_id, aluno_id)
        faixa = _faixa_saeb(nota_10)
        notas_alunos.append({
            "aluno_id": aluno_id, "score": score, "total": total,
            "acertos_obj": acertos_obj, "total_obj": total_obj,
            "nota_10": nota_10, "faixa": faixa,
        })

    media_acertos = sum(n["score"] for n in notas_alunos) / total_entregas
    media_pct = (media_acertos / total_questoes) * 100 if total_questoes > 0 else 0
    maior_nota = max(n["score"] for n in notas_alunos)
    menor_nota = min(n["score"] for n in notas_alunos)
    # Média na escala 0-10 (só objetivas)
    media_10 = sum(n["nota_10"] for n in notas_alunos) / total_entregas if total_entregas else 0
    faixa_media = _faixa_saeb(media_10)
    # Distribuição nas 4 faixas
    dist_faixas = {f["nome"]: 0 for f in FAIXAS_SAEB}
    for n in notas_alunos:
        dist_faixas[n["faixa"]["nome"]] += 1

    questoes_stats = []
    for idx, q in enumerate(questoes, start=1):
        stats = _estatisticas_questao(conn, aplicacao_id, q["id"], alunos_entregues)
        stats["questao_id"] = q["id"]
        stats["numero"] = idx
        preview = _preview_enunciado(q["enunciado"], max_chars=120)
        if len(q["enunciado"]) > 120:
            preview += "..."
        stats["enunciado_preview"] = preview
        stats["disciplina_nome"] = q["disciplina_nome"]
        questoes_stats.append(stats)

    questoes_ordenadas = sorted(questoes_stats, key=lambda x: x["pct_acerto"])
    mais_dificeis = questoes_ordenadas[:3]
    mais_faceis = list(reversed(questoes_ordenadas[-3:]))

    habilidades = conn.execute("""
        SELECT DISTINCT h.id, h.codigo, h.descricao
        FROM habilidades_bncc h
        JOIN questao_habilidades qh ON qh.habilidade_id = h.id
        JOIN prova_questoes pq ON pq.questao_id = qh.questao_id
        WHERE pq.prova_id = ?
        ORDER BY h.codigo
    """, (apl["prova_id"],)).fetchall()

    habilidades_stats = []
    for h in habilidades:
        questoes_da_hab = conn.execute("""
            SELECT q.id FROM questao_habilidades qh
            JOIN questoes q ON q.id = qh.questao_id
            JOIN prova_questoes pq ON pq.questao_id = q.id
            WHERE qh.habilidade_id = ? AND pq.prova_id = ?
        """, (h["id"], apl["prova_id"])).fetchall()
        acertos_h = 0
        oport_h = 0
        for q in questoes_da_hab:
            s = _estatisticas_questao(conn, aplicacao_id, q["id"], alunos_entregues)
            acertos_h += s["acertos"]
            oport_h += s["total"]
        pct = (acertos_h / oport_h) * 100 if oport_h > 0 else 0
        habilidades_stats.append({
            "codigo": h["codigo"],
            "descricao": h["descricao"] or "—",
            "n_questoes": len(questoes_da_hab),
            "pct_acerto": pct,
            "acertos": acertos_h,
            "total": oport_h,
        })

    alunos_info = {a["id"]: a for a in conn.execute("SELECT id, nome, numero FROM alunos WHERE turma_id = ?", (apl["turma_id"],)).fetchall()}

    ranking = []
    for n in notas_alunos:
        aluno = alunos_info.get(n["aluno_id"])
        if aluno:
            pct = (n["score"] / n["total"] * 100) if n["total"] > 0 else 0
            ranking.append({
                "nome": aluno["nome"],
                "numero": aluno["numero"],
                "aluno_id": n["aluno_id"],
                "score": n["score"],
                "total": n["total"],
                "pct": pct,
                "nota_10": n["nota_10"],
                "faixa": n["faixa"],
            })
    ranking.sort(key=lambda x: x["nota_10"], reverse=True)

    # Alertas automáticos
    alunos_alerta = [r for r in ranking if r["faixa"]["nome"] == "Insuficiente"]
    alunos_destaque = [r for r in ranking if r["faixa"]["nome"] == "Avançado"]

    conn.close()

    # ═══ BLOCO SAEB: KPI principal (média da turma + faixa) ═══
    saeb_kpi_html = f"""
        <div class="card" style="background:{faixa_media['cor_bg']}; border-left:4px solid {faixa_media['cor']}; padding:18px 20px; margin-bottom:18px;">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:14px;">
                <div>
                    <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; font-weight:700; color:{faixa_media['cor']};">Proficiência média da turma</div>
                    <div style="font-size:36px; font-weight:800; color:{faixa_media['cor']}; line-height:1.1; margin-top:4px;">{media_10:.1f} <small style="font-size:18px; opacity:0.7;">/10</small></div>
                    <div style="font-size:14px; color:var(--text); margin-top:6px; font-weight:600;">{faixa_media['emoji']} {faixa_media['nome']}</div>
                </div>
                <div style="font-size:12px; color:var(--text-muted); max-width:300px;">
                    Calculado sobre {total_entregas} aluno(s) entregue(s) considerando apenas questões objetivas (MC, V/F, Associação). Discursivas são corrigidas manualmente.
                </div>
            </div>
        </div>
    """

    # ═══ DISTRIBUIÇÃO NAS 4 FAIXAS (4 cards + gráfico) ═══
    cards_faixas = ""
    for f in FAIXAS_SAEB:
        qtd = dist_faixas[f["nome"]]
        pct_faixa = (qtd / total_entregas * 100) if total_entregas else 0
        cards_faixas += f"""
            <div class="card" style="background:{f['cor_bg']}; border-color:{f['cor_border']}; padding:14px 16px; text-align:center;">
                <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; font-weight:700; color:{f['cor']};">{f['emoji']} {f['nome']}</div>
                <div style="font-size:28px; font-weight:800; color:{f['cor']}; line-height:1.1; margin-top:6px;">{qtd}</div>
                <div style="font-size:12px; color:var(--text-muted); margin-top:4px;">{pct_faixa:.0f}% da turma</div>
            </div>
        """
    distribuicao_html = f"""
        <h2 style="margin-top:24px;">📊 Distribuição por faixa de proficiência</h2>
        <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; margin-bottom:18px;">
            {cards_faixas}
        </div>
        <div class="card" style="padding:16px;">
            <canvas id="chartDistFaixas" style="max-height:280px;"></canvas>
        </div>
    """

    # ═══ GRÁFICO de % de acerto por questão ═══
    questoes_labels_js = "[" + ", ".join(f'"Q{q["numero"]}"' for q in questoes_stats) + "]"
    questoes_valores_js = "[" + ", ".join(f'{q["pct_acerto"]:.1f}' for q in questoes_stats) + "]"
    # Cores: vermelho < 50%, laranja 50-65%, azul 66-79%, verde >= 80%
    questoes_cores_js = "[" + ", ".join(
        f'"{_faixa_saeb(q["pct_acerto"]/10)["hex"]}"' for q in questoes_stats
    ) + "]"
    chart_questoes_html = f"""
        <h2 style="margin-top:32px;">📈 % de acerto por questão</h2>
        <div class="card" style="padding:16px;">
            <canvas id="chartQuestoes" style="max-height:320px;"></canvas>
        </div>
    """

    # ═══ ALERTAS automáticos ═══
    alertas_html = ""
    if alunos_alerta or alunos_destaque:
        alerta_list = ""
        if alunos_alerta:
            items = "".join(
                f'<li style="padding:6px 0; display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border);">'
                f'<span><strong>{a["nome"]}</strong>{" · Nº " + str(a["numero"]) if a["numero"] else ""}</span>'
                f'<span style="font-weight:700; color:var(--red);">{a["nota_10"]:.1f}/10</span>'
                f'</li>'
                for a in alunos_alerta
            )
            alerta_list = f"""
                <div class="card" style="background:var(--red-bg); border-left:4px solid var(--red); padding:14px 16px;">
                    <div style="font-size:13px; font-weight:700; color:var(--red); margin-bottom:8px;">🔴 Atenção necessária — {len(alunos_alerta)} aluno(s) na faixa Insuficiente</div>
                    <div style="font-size:11px; color:var(--text-muted); margin-bottom:10px;">Recomende reforço, retomada de conteúdo ou acompanhamento individual.</div>
                    <ul style="list-style:none; padding:0; margin:0;">{items}</ul>
                </div>
            """
        destaque_list = ""
        if alunos_destaque:
            items = "".join(
                f'<li style="padding:6px 0; display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border);">'
                f'<span><strong>{a["nome"]}</strong>{" · Nº " + str(a["numero"]) if a["numero"] else ""}</span>'
                f'<span style="font-weight:700; color:var(--green);">{a["nota_10"]:.1f}/10</span>'
                f'</li>'
                for a in alunos_destaque
            )
            destaque_list = f"""
                <div class="card" style="background:var(--green-bg); border-left:4px solid var(--green); padding:14px 16px;">
                    <div style="font-size:13px; font-weight:700; color:var(--green); margin-bottom:8px;">🟢 Destaque positivo — {len(alunos_destaque)} aluno(s) na faixa Avançado</div>
                    <div style="font-size:11px; color:var(--text-muted); margin-bottom:10px;">Vale parabenizar e considerar atividades de aprofundamento.</div>
                    <ul style="list-style:none; padding:0; margin:0;">{items}</ul>
                </div>
            """
        alertas_html = f"""
            <h2 style="margin-top:32px;">🎯 Alertas automáticos</h2>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:14px;">
                {alerta_list or '<div></div>'}
                {destaque_list or '<div></div>'}
            </div>
        """

    # ═══ Script Chart.js (renderiza os 2 gráficos) ═══
    dist_data_js = "[" + ", ".join(str(dist_faixas[f["nome"]]) for f in FAIXAS_SAEB) + "]"
    dist_cores_js = "[" + ", ".join(f'"{f["hex"]}"' for f in FAIXAS_SAEB) + "]"
    dist_labels_js = "[" + ", ".join(f'"{f["emoji"]} {f["nome"]}"' for f in FAIXAS_SAEB) + "]"

    charts_script = f"""
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js" onerror="window._chartFail=true"></script>
    <script>
    (function() {{
      // Fallback visual se Chart.js não carregou (ex: rede sem acesso ao CDN)
      if (typeof Chart === 'undefined') {{
        document.querySelectorAll('canvas[id^=chart]').forEach(function(c) {{
          var msg = document.createElement('div');
          msg.style.cssText = 'padding:20px; text-align:center; color:var(--text-muted); font-size:12px; font-style:italic;';
          msg.innerHTML = '📡 Gráfico indisponível (Chart.js não carregou — verifique conexão).';
          c.parentNode.replaceChild(msg, c);
        }});
        return;
      }}

      // Cores que se adaptam ao tema (texto e grid)
      function temaCores() {{
        var dark = document.documentElement.getAttribute('data-theme') === 'dark';
        return {{
          texto: dark ? '#e2eaf5' : '#1e293b',
          grid: dark ? '#1e3050' : '#e2e8f0',
          tooltipBg: dark ? '#172341' : '#ffffff',
          tooltipBorder: dark ? '#1e3050' : '#e2e8f0',
        }};
      }}

      var cores = temaCores();

      // ━━ Gráfico 1: distribuição nas 4 faixas ━━
      var elDist = document.getElementById('chartDistFaixas');
      if (elDist) {{
        new Chart(elDist, {{
          type: 'bar',
          data: {{
            labels: {dist_labels_js},
            datasets: [{{
              label: 'Alunos',
              data: {dist_data_js},
              backgroundColor: {dist_cores_js},
              borderRadius: 8,
              borderSkipped: false,
            }}]
          }},
          options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
              legend: {{ display: false }},
              tooltip: {{
                backgroundColor: cores.tooltipBg,
                titleColor: cores.texto,
                bodyColor: cores.texto,
                borderColor: cores.tooltipBorder,
                borderWidth: 1,
              }}
            }},
            scales: {{
              x: {{ ticks: {{ color: cores.texto, font: {{ size: 12, weight: 600 }} }}, grid: {{ display: false }} }},
              y: {{ beginAtZero: true, ticks: {{ color: cores.texto, precision: 0 }}, grid: {{ color: cores.grid }} }}
            }}
          }}
        }});
      }}

      // ━━ Gráfico 2: % de acerto por questão ━━
      var elQ = document.getElementById('chartQuestoes');
      if (elQ) {{
        new Chart(elQ, {{
          type: 'bar',
          data: {{
            labels: {questoes_labels_js},
            datasets: [{{
              label: '% de acerto',
              data: {questoes_valores_js},
              backgroundColor: {questoes_cores_js},
              borderRadius: 6,
              borderSkipped: false,
            }}]
          }},
          options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
              legend: {{ display: false }},
              tooltip: {{
                backgroundColor: cores.tooltipBg,
                titleColor: cores.texto,
                bodyColor: cores.texto,
                borderColor: cores.tooltipBorder,
                borderWidth: 1,
                callbacks: {{
                  label: function(ctx) {{ return ctx.parsed.y.toFixed(1) + '% de acerto'; }}
                }}
              }}
            }},
            scales: {{
              x: {{ ticks: {{ color: cores.texto }}, grid: {{ display: false }} }},
              y: {{ beginAtZero: true, max: 100, ticks: {{ color: cores.texto, callback: v => v + '%' }}, grid: {{ color: cores.grid }} }}
            }}
          }}
        }});
      }}
    }})();
    </script>
    """

    metrics_html = f"""
        <div class="metric-grid">
            <div class="metric"><div class="metric-label">Alunos com entrega</div><div class="metric-value">{total_entregas}</div></div>
            <div class="metric"><div class="metric-label">Média da turma</div><div class="metric-value">{media_acertos:.1f}<small style="font-size:14px; color:var(--text-muted);">/{total_questoes}</small></div><div class="card-meta">{media_pct:.1f}%</div></div>
            <div class="metric"><div class="metric-label">Maior nota</div><div class="metric-value">{maior_nota}/{total_questoes}</div></div>
            <div class="metric"><div class="metric-label">Menor nota</div><div class="metric-value">{menor_nota}/{total_questoes}</div></div>
        </div>
    """

    destaques_html = ""
    if mais_dificeis:
        items_dif = "".join(
            f'<li style="padding:6px 0;">Q{q["numero"]} ({q["disciplina_nome"]}) — {_barra_html(q["pct_acerto"], 120)}</li>'
            for q in mais_dificeis
        )
        items_fac = "".join(
            f'<li style="padding:6px 0;">Q{q["numero"]} ({q["disciplina_nome"]}) — {_barra_html(q["pct_acerto"], 120)}</li>'
            for q in mais_faceis
        )
        destaques_html = f"""
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin:24px 0;">
            <div class="card">
                <div class="card-title" style="color:var(--red);">Questões mais difíceis</div>
                <ul style="list-style:none; padding:0; margin:12px 0 0;">{items_dif}</ul>
            </div>
            <div class="card">
                <div class="card-title" style="color:var(--green);">Questões mais fáceis</div>
                <ul style="list-style:none; padding:0; margin:12px 0 0;">{items_fac}</ul>
            </div>
        </div>
        """

    questoes_detalhe_html = ""
    for q in questoes_stats:
        dist_html = ""
        for letra in ["A", "B", "C", "D"]:
            count = q["distribuicao"][letra]
            pct_letra = (count / q["total"] * 100) if q["total"] > 0 else 0
            destaque = ""
            if letra == q["correta_letra"]:
                destaque = ' style="background:var(--green-bg); color:var(--green); font-weight:600; padding:2px 6px; border-radius:4px;"'
            dist_html += f'<span{destaque}>{letra}: {count} ({pct_letra:.0f}%)</span>'
            if letra != "D":
                dist_html += '<span style="color:var(--text-subtle); margin:0 6px;">·</span>'
        em_branco = q["em_branco"]
        em_branco_html = f' · <span style="color:var(--text-muted);">Em branco: {em_branco}</span>' if em_branco > 0 else ""
        questoes_detalhe_html += f"""
        <div class="card">
            <div class="card-meta">Questão {q["numero"]} · {q["disciplina_nome"]}</div>
            <div style="margin:8px 0 12px;">{q["enunciado_preview"]}</div>
            {_barra_html(q["pct_acerto"], 300)}
            <div style="margin-top:12px; font-size:13px;">{dist_html}{em_branco_html}</div>
            <div style="margin-top:8px;"><a href="/aplicacoes/{aplicacao_id}" style="font-size:12px; color:var(--text-muted);">Detalhes da aplicação →</a></div>
        </div>
        """

    habilidades_detalhe_html = ""
    if habilidades_stats:
        habs_html = ""
        for h in habilidades_stats:
            habs_html += f"""
            <div class="card">
                <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
                    <span class="badge">{h["codigo"]}</span>
                    <small style="color:var(--text-muted);">{h["n_questoes"]} questão{"" if h["n_questoes"] == 1 else "ões"} na prova</small>
                </div>
                <div style="font-size:13px; margin-bottom:8px; color:var(--text-muted);">{h["descricao"]}</div>
                {_barra_html(h["pct_acerto"], 300)}
                <div style="margin-top:6px; font-size:12px; color:var(--text-muted);">{h["acertos"]} acertos em {h["total"]} oportunidades</div>
            </div>
            """
        habilidades_detalhe_html = f'<h2>Desempenho por habilidade BNCC</h2>{habs_html}'
    else:
        habilidades_detalhe_html = '<h2>Habilidades BNCC</h2><div class="empty">Nenhuma das questões desta prova tem habilidades BNCC cadastradas. Vincule códigos BNCC nas questões para ver esta análise.</div>'

    ranking_html = ""
    for pos, r in enumerate(ranking, start=1):
        num = r["numero"] if r["numero"] else "—"
        f = r["faixa"]
        # Recupera total objetivo do aluno (vem de notas_alunos)
        n_dict = next((x for x in notas_alunos if x["aluno_id"] == r["aluno_id"]), {})
        total_obj_aluno = n_dict.get("total_obj", 0)
        acertos_obj_aluno = n_dict.get("acertos_obj", 0)
        ranking_html += f"""
        <div class="student-row">
            <div class="numero">{pos}º</div>
            <div>
                <a href="/aplicacoes/{aplicacao_id}/aluno/{r["aluno_id"]}" style="color:inherit; text-decoration:none; font-weight:600;">{r["nome"]}</a>
                <div style="font-size:12px; color:var(--text-muted);">Nº {num} · <span style="color:{f['cor']}; font-weight:600;">{f['emoji']} {f['nome']}</span></div>
            </div>
            <div style="text-align:right;">
                <div style="font-size:15px; font-weight:700; color:{f['cor']};">{r["nota_10"]:.1f}/10</div>
                <div style="font-size:11px; color:var(--text-muted);">{acertos_obj_aluno}/{total_obj_aluno} acertos</div>
            </div>
        </div>
        """

    content = f"""
        <div class="page-header">
            <h1>📈 Análise pedagógica</h1>
            <p class="subtitle">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]})</p>
            <div class="page-actions">
                <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar para a aplicação</a>
                <a href="/aplicacoes/{aplicacao_id}/exportar" class="btn">📊 Exportar Excel</a>
            </div>
        </div>

        {saeb_kpi_html}
        {metrics_html}

        {distribuicao_html}
        {alertas_html}

        {chart_questoes_html}
        {destaques_html}

        <h2 style="margin-top:32px;">🏆 Ranking de alunos</h2>
        {ranking_html}

        <h2 style="margin-top:32px;">Desempenho detalhado por questão</h2>
        {questoes_detalhe_html}

        <div style="margin-top:32px;"></div>
        {habilidades_detalhe_html}

        {charts_script}
    """
    return render_page("Análise pedagógica", content, active="aplicacoes", head_extra=MATHJAX)


@app.get("/provas/{prova_id}/comparativo", response_class=HTMLResponse)
def comparativo_prova(prova_id: int):
    conn = get_db()
    prova = conn.execute("SELECT * FROM provas WHERE id = ?", (prova_id,)).fetchone()
    if not prova:
        conn.close()
        return RedirectResponse("/provas", status_code=303)

    aplicacoes = conn.execute("""
        SELECT a.id, a.titulo, a.modo, a.criada_em, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a JOIN turmas t ON t.id = a.turma_id
        WHERE a.prova_id = ?
        ORDER BY a.criada_em DESC
    """, (prova_id,)).fetchall()

    if not aplicacoes:
        conn.close()
        content = f"""
            <div class="page-header">
                <h1>Comparativo de aplicações</h1>
                <p class="subtitle">{prova["titulo"]}</p>
            </div>
            <div class="empty">
                <p>Esta prova ainda não foi aplicada em nenhuma turma.</p>
                <a href="/aplicacoes/nova" class="btn btn-primary">Criar aplicação</a>
            </div>
        """
        return render_page("Comparativo", content, active="provas")

    questoes = conn.execute("""
        SELECT q.id, q.enunciado, d.nome AS disciplina_nome
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        JOIN disciplinas d ON d.id = q.disciplina_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (prova_id,)).fetchall()
    total_questoes = len(questoes)

    habilidades = conn.execute("""
        SELECT DISTINCT h.id, h.codigo
        FROM habilidades_bncc h
        JOIN questao_habilidades qh ON qh.habilidade_id = h.id
        JOIN prova_questoes pq ON pq.questao_id = qh.questao_id
        WHERE pq.prova_id = ?
        ORDER BY h.codigo
    """, (prova_id,)).fetchall()

    aplicacoes_dados = []
    for apl in aplicacoes:
        entregas = conn.execute("SELECT aluno_id FROM entregas WHERE aplicacao_id = ?", (apl["id"],)).fetchall()
        alunos_entregues = [e["aluno_id"] for e in entregas]
        n_entregas = len(alunos_entregues)
        total_alunos = conn.execute("SELECT COUNT(*) AS c FROM alunos a JOIN aplicacoes ap ON ap.turma_id = a.turma_id WHERE ap.id = ?", (apl["id"],)).fetchone()["c"]

        if n_entregas > 0:
            notas = []
            for aluno_id in alunos_entregues:
                score, _ = _calcular_nota(conn, apl["id"], aluno_id)
                notas.append(score)
            media = sum(notas) / n_entregas
            media_pct = (media / total_questoes * 100) if total_questoes > 0 else 0
        else:
            media = 0
            media_pct = 0

        questoes_stats = {}
        for q in questoes:
            s = _estatisticas_questao(conn, apl["id"], q["id"], alunos_entregues)
            questoes_stats[q["id"]] = s

        habilidades_stats = {}
        for h in habilidades:
            qs_hab = conn.execute("""
                SELECT q.id FROM questao_habilidades qh
                JOIN questoes q ON q.id = qh.questao_id
                JOIN prova_questoes pq ON pq.questao_id = q.id
                WHERE qh.habilidade_id = ? AND pq.prova_id = ?
            """, (h["id"], prova_id)).fetchall()
            ac = 0
            op = 0
            for q in qs_hab:
                s = _estatisticas_questao(conn, apl["id"], q["id"], alunos_entregues)
                ac += s["acertos"]
                op += s["total"]
            habilidades_stats[h["id"]] = (ac / op * 100) if op > 0 else 0

        aplicacoes_dados.append({
            "id": apl["id"],
            "titulo": apl["titulo"] or f'{apl["turma_nome"]} ({apl["ano_letivo"]})',
            "turma_nome": apl["turma_nome"],
            "ano_letivo": apl["ano_letivo"],
            "modo": apl["modo"],
            "n_entregas": n_entregas,
            "total_alunos": total_alunos,
            "media": media,
            "media_pct": media_pct,
            "questoes_stats": questoes_stats,
            "habilidades_stats": habilidades_stats,
        })

    conn.close()

    cards_html = ""
    for ad in aplicacoes_dados:
        pct_entrega = (ad["n_entregas"] / ad["total_alunos"] * 100) if ad["total_alunos"] > 0 else 0
        cards_html += f"""
        <div class="card">
            <div class="card-title">
                <a href="/aplicacoes/{ad["id"]}" style="color:inherit; text-decoration:none;">{ad["titulo"]}</a>
            </div>
            <div class="card-meta">{ad["turma_nome"]} · {ad["ano_letivo"]} · Modo {ad["modo"]}</div>
            <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:12px; margin-top:12px;">
                <div>
                    <div class="metric-label">Entregas</div>
                    <div style="font-size:18px; font-weight:600;">{ad["n_entregas"]}/{ad["total_alunos"]}</div>
                    <div class="card-meta">{pct_entrega:.0f}%</div>
                </div>
                <div>
                    <div class="metric-label">Média</div>
                    <div style="font-size:18px; font-weight:600;">{ad["media"]:.1f}<small style="color:var(--text-muted); font-size:12px;">/{total_questoes}</small></div>
                    <div class="card-meta">{ad["media_pct"]:.0f}%</div>
                </div>
                <div style="display:flex; align-items:flex-end; justify-content:flex-end;">
                    <a href="/aplicacoes/{ad["id"]}/analise" class="btn" style="font-size:12px; padding:4px 10px;">Análise detalhada</a>
                </div>
            </div>
        </div>
        """

    questoes_tabela_html = ""
    if total_questoes > 0:
        col_aplicacoes = "".join(f'<th style="text-align:center; padding:8px; font-size:12px; min-width:110px;">{ad["turma_nome"]}</th>' for ad in aplicacoes_dados)
        rows = ""
        for idx, q in enumerate(questoes, start=1):
            preview = _preview_enunciado(q["enunciado"], max_chars=60)
            if len(q["enunciado"]) > 60:
                preview += "..."
            cells = ""
            for ad in aplicacoes_dados:
                s = ad["questoes_stats"].get(q["id"], {"pct_acerto": 0, "total": 0})
                if s["total"] == 0:
                    cells += '<td style="text-align:center; padding:8px; color:var(--text-subtle);">—</td>'
                else:
                    cor = _cor_por_pct(s["pct_acerto"])
                    cells += f'<td style="text-align:center; padding:8px; color:{cor}; font-weight:600;">{s["pct_acerto"]:.0f}%</td>'
            rows += f"""
            <tr style="border-top:1px solid var(--border);">
                <td style="padding:8px; font-size:13px;"><strong>Q{idx}</strong> <span style="color:var(--text-muted);">({q["disciplina_nome"]})</span><br><small style="color:var(--text-subtle);">{preview}</small></td>
                {cells}
            </tr>
            """
        questoes_tabela_html = f"""
        <h2 style="margin-top:32px;">Acertos por questão</h2>
        <p class="muted-line">Percentual de alunos que acertaram cada questão, em cada turma. Verde ≥ 70%, amarelo 40-70%, vermelho &lt; 40%.</p>
        <div style="overflow-x:auto;">
        <table style="width:100%; border-collapse:collapse; margin-top:12px;">
            <thead>
                <tr style="background:var(--bg-subtle);">
                    <th style="text-align:left; padding:8px;">Questão</th>
                    {col_aplicacoes}
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        </div>
        """

    habilidades_tabela_html = ""
    if habilidades:
        col_aplicacoes = "".join(f'<th style="text-align:center; padding:8px; font-size:12px; min-width:110px;">{ad["turma_nome"]}</th>' for ad in aplicacoes_dados)
        rows = ""
        for h in habilidades:
            cells = ""
            for ad in aplicacoes_dados:
                pct = ad["habilidades_stats"].get(h["id"], 0)
                if ad["n_entregas"] == 0:
                    cells += '<td style="text-align:center; padding:8px; color:var(--text-subtle);">—</td>'
                else:
                    cor = _cor_por_pct(pct)
                    cells += f'<td style="text-align:center; padding:8px; color:{cor}; font-weight:600;">{pct:.0f}%</td>'
            rows += f"""
            <tr style="border-top:1px solid var(--border);">
                <td style="padding:8px;"><span class="badge">{h["codigo"]}</span></td>
                {cells}
            </tr>
            """
        habilidades_tabela_html = f"""
        <h2 style="margin-top:32px;">Acertos por habilidade BNCC</h2>
        <p class="muted-line">Percentual médio de acertos nas questões vinculadas a cada habilidade.</p>
        <div style="overflow-x:auto;">
        <table style="width:100%; border-collapse:collapse; margin-top:12px;">
            <thead>
                <tr style="background:var(--bg-subtle);">
                    <th style="text-align:left; padding:8px;">Habilidade</th>
                    {col_aplicacoes}
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        </div>
        """

    titulo_pagina = "Comparativo entre turmas" if len(aplicacoes_dados) >= 2 else "Análises da prova"
    subtitulo = f'{prova["titulo"]} · {len(aplicacoes_dados)} aplicação{"" if len(aplicacoes_dados) == 1 else "ões"}'

    content = f"""
        <div class="page-header">
            <h1>{titulo_pagina}</h1>
            <p class="subtitle">{subtitulo}</p>
            <div class="page-actions">
                <a href="/provas/{prova_id}" class="btn">← Voltar para a prova</a>
            </div>
        </div>

        <h2>Aplicações desta prova</h2>
        {cards_html}

        {questoes_tabela_html}
        {habilidades_tabela_html}
    """
    return render_page(titulo_pagina, content, active="provas")


# ==========================================
#  FASE C1: PDF DA PROVA E CARTÃO RESPOSTA
# ==========================================

@app.get("/provas/{prova_id}/imprimir", response_class=HTMLResponse)
def imprimir_prova(prova_id: int):
    """Versão da prova otimizada pra impressão (sem gabarito visível)."""
    conn = get_db()
    prova = conn.execute("SELECT * FROM provas WHERE id = ?", (prova_id,)).fetchone()
    if not prova:
        conn.close()
        return RedirectResponse("/provas", status_code=303)

    questoes = conn.execute("""
        SELECT q.id, q.enunciado, q.tipo, d.nome AS disciplina_nome
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        JOIN disciplinas d ON d.id = q.disciplina_id
        WHERE pq.prova_id = ? ORDER BY pq.ordem
    """, (prova_id,)).fetchall()

    questoes_html = ""
    for idx, q in enumerate(questoes, start=1):
        tipo_q = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
        textos = conn.execute("SELECT conteudo, fonte FROM textos_apoio WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
        imagens = conn.execute("SELECT caminho, legenda, fonte FROM imagens WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()

        textos_html = ""
        for t in textos:
            fonte_html = f'<footer>Fonte: {t["fonte"]}</footer>' if t["fonte"] else ""
            textos_html += f'<blockquote>{t["conteudo"]}{fonte_html}</blockquote>'

        imagens_html = ""
        for img in imagens:
            legenda_html = f'<figcaption>{img["legenda"]}</figcaption>' if img["legenda"] else ""
            fonte_html = f'<figcaption><small>Fonte: {img["fonte"]}</small></figcaption>' if img["fonte"] else ""
            imagens_html += f'<figure><img src="/{img["caminho"]}" alt="">{legenda_html}{fonte_html}</figure>'

        # Conteúdo específico do tipo
        if tipo_q == "multipla_escolha":
            alts = conn.execute("SELECT letra, texto FROM alternativas WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()
            corpo_resposta = "<ul class=\"q-alts\">" + "".join(f'<li><strong>{a["letra"]})</strong> {a["texto"]}</li>' for a in alts) + "</ul>"
        elif tipo_q == "discursiva":
            # Linhas em branco pra resposta manuscrita
            linhas = "".join('<div style="border-bottom: 1px solid #999; height: 22px; margin-bottom: 4px;"></div>' for _ in range(6))
            corpo_resposta = f'<div style="margin-top:10px;"><strong style="font-size:11px; color:#555;">Resposta:</strong>{linhas}</div>'
        elif tipo_q == "vf":
            afirms = conn.execute("SELECT ordem, texto FROM vf_afirmacoes WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
            linhas = ""
            for af in afirms:
                num_sub = f"{idx}.{af['ordem']+1}"
                linhas += (
                    f'<div style="display:grid; grid-template-columns:auto 1fr auto; gap:10px; align-items:center; margin-bottom:6px; padding:4px 0; border-bottom:1px dotted #ccc;">'
                    f'<strong style="min-width:32px;">{num_sub}</strong>'
                    f'<span>{af["texto"]}</span>'
                    f'<span style="font-size:11px; color:#555; white-space:nowrap;">( ) V&nbsp;&nbsp;( ) F</span>'
                    f'</div>'
                )
            corpo_resposta = f'<div style="margin-top:8px;"><strong style="font-size:11px; color:#555;">Julgue cada afirmação:</strong>{linhas}</div>'
        elif tipo_q == "associacao":
            itens_a = conn.execute("SELECT ordem, texto FROM assoc_itens_a WHERE questao_id = ? ORDER BY ordem", (q["id"],)).fetchall()
            itens_b = conn.execute("SELECT letra, texto FROM assoc_itens_b WHERE questao_id = ? ORDER BY letra", (q["id"],)).fetchall()
            ca = "".join(f'<li><strong>{a["ordem"]+1}.</strong> {a["texto"]}</li>' for a in itens_a)
            cb = "".join(f'<li><strong>({b["letra"]})</strong> {b["texto"]}</li>' for b in itens_b)
            # Linhas pra o aluno escrever as respostas (1→ , 2→ , ...)
            respostas = " ".join(
                f'<span style="border-bottom:1px solid #555; display:inline-block; min-width:30px;">&nbsp;</span> ({a["ordem"]+1})'
                for a in itens_a
            )
            corpo_resposta = (
                f'<div style="display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-top:8px;">'
                f'<div><strong style="font-size:11px; color:#555;">Coluna A</strong><ul style="margin:6px 0 0 20px;">{ca}</ul></div>'
                f'<div><strong style="font-size:11px; color:#555;">Coluna B</strong><ul style="margin:6px 0 0 20px;">{cb}</ul></div>'
                f'</div>'
                f'<div style="margin-top:10px; padding:6px 10px; background:#f9f9f9; border:1px dashed #aaa; border-radius:4px; font-size:12px;">'
                f'<strong>Resposta — associe cada item da A à letra da B:</strong> {respostas}</div>'
            )
        else:
            corpo_resposta = ""

        bncc_pref = _bncc_prefix(conn, q["id"])
        questoes_html += f"""
        <div class="q-print">
            <div class="q-head">Questão {idx} · {q['disciplina_nome']}</div>
            {textos_html}{imagens_html}
            <div class="q-enunciado">{bncc_pref}{q['enunciado']}</div>
            {corpo_resposta}
        </div>
        """

    conn.close()

    desc_html = f'<p style="margin:6px 0 0; color:#555;">{prova["descricao"]}</p>' if prova["descricao"] else ""

    return HTMLResponse(f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <title>{prova["titulo"]} — Versão impressa</title>
    {INTER_FONT}
    {MATHJAX}
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Sora', -apple-system, sans-serif;
            margin: 0 auto;
            padding: 24px;
            background: white;
            color: black;
            max-width: 21cm;
        }}
        .actions {{ margin-bottom: 24px; display: flex; gap: 8px; }}
        .btn {{ display: inline-block; padding: 8px 16px; background: #2563eb; color: white; text-decoration: none; border-radius: 6px; font-size: 14px; border: none; cursor: pointer; font-family: inherit; font-weight: 500; }}
        .btn-secondary {{ background: #6b7280; }}
        .inst-header {{ display: flex; gap: 14px; align-items: center; padding-bottom: 12px; border-bottom: 1px solid #888; margin-bottom: 20px; }}
        .inst-header img {{ width: 60px; height: auto; flex-shrink: 0; }}
        .inst-header .inst-text {{ font-size: 11px; color: #333; line-height: 1.5; }}
        .inst-header .inst-text .escola {{ font-weight: 700; font-size: 13px; color: #000; }}
        .header {{ border-bottom: 2px solid #000; padding-bottom: 14px; margin-bottom: 22px; }}
        .header h1 {{ margin: 0; font-size: 22px; }}
        .student-info {{ display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 16px; margin-top: 16px; font-size: 11px; color:#555; }}
        .student-info > div {{ border-bottom: 1px solid #555; padding-bottom: 8px; min-height: 28px; }}
        .q-print {{ margin-bottom: 22px; page-break-inside: avoid; }}
        .q-head {{ font-size: 11px; color: #555; margin-bottom: 6px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        blockquote {{ border-left: 3px solid #aaa; padding: 8px 14px; margin: 10px 0; color: #333; font-style: italic; background: #fafafa; }}
        blockquote footer {{ font-size: 10px; margin-top: 6px; font-style: normal; }}
        figure {{ margin: 10px 0; }}
        figure img {{ max-width: 100%; max-height: 250px; }}
        figcaption {{ font-size: 10px; color: #666; margin-top: 2px; }}
        .q-enunciado {{ margin: 8px 0 10px; line-height: 1.5; }}
        .q-alts {{ list-style: none; padding-left: 8px; margin: 8px 0; }}
        .q-alts li {{ padding: 4px 0; line-height: 1.5; }}
        @media print {{
            @page {{ size: A4; margin: 1.5cm; }}
            body {{ padding: 0; max-width: 100%; }}
            .no-print {{ display: none !important; }}
        }}
    </style>
</head>
<body>
    <div class="actions no-print">
        <button onclick="window.print()" class="btn">🖨️ Imprimir / Salvar PDF</button>
        <a href="/provas/{prova_id}" class="btn btn-secondary">← Voltar</a>
    </div>
    <div class="inst-header">
        <img src="/static/imagens/brasao_vr.png" alt="Brasão Volta Redonda">
        <div class="inst-text">
            <div>Estado do Rio de Janeiro</div>
            <div>Prefeitura de Volta Redonda</div>
            <div>Secretaria Municipal de Educação</div>
            <div class="escola">E.M. WALMIR DE FREITAS MONTEIRO</div>
        </div>
    </div>
    <div class="header">
        <h1>{prova["titulo"]}</h1>
        {desc_html}
        <div class="student-info">
            <div><small>Nome:</small></div>
            <div><small>Turma:</small></div>
            <div><small>Data:</small></div>
        </div>
    </div>
    {questoes_html}
</body>
</html>""")


def _calcular_layout_cartao(questoes_info):
    """Calcula coordenadas das bolhas pro cartão de respostas.
    Recebe lista de dicts: [{id, num, tipo, vf_count, assoc_a_count, assoc_b_count}, ...]
    Retorna lista de blocos: [{num, tipo, top_y_mm, height_mm, bubbles: [{label, x_mm, y_mm}], header: str}]
    Coordenadas em mm. Origem (0,0) é o canto superior esquerdo da folha A4 (210x297mm).
    Usado tanto pelo gerador de PDF quanto pelo leitor OMR — garantia de sincronia."""
    # Dimensões A4
    PAGE_W = 210
    PAGE_H = 297
    MARGEM = 10
    MARKER = 8

    # Área útil: abaixo do cabeçalho (~55mm pra título/aluno/QR)
    AREA_TOP_Y = 75      # começa abaixo do header
    AREA_BOTTOM_Y = PAGE_H - MARGEM - MARKER - 3   # acima dos marcadores inferiores
    BLOCK_W_MM = 80
    BLOCK_X_INIT = 25    # x do início da coluna 1
    NUM_OFFSET_X = 18    # onde fica o número
    FIRST_BUBBLE_X = 28  # onde começa a primeira bolha

    ROW_H = 8           # altura padrão de linha
    BUBBLE_SPACING = 9  # distância entre bolhas
    HEADER_H = 6        # cabeçalho de cada bloco (tipo da questão)

    blocos = []
    cur_y = AREA_TOP_Y
    cur_col = 0
    cur_x = BLOCK_X_INIT + cur_col * BLOCK_W_MM

    for info in questoes_info:
        tipo = info["tipo"]
        if tipo == "discursiva":
            continue  # discursiva não vai pro cartão (correção manual)

        # Calcular altura do bloco
        if tipo == "multipla_escolha":
            n_linhas = 1
        elif tipo == "vf":
            n_linhas = info.get("vf_count", 0)
            if n_linhas == 0:
                continue
        elif tipo == "associacao":
            n_linhas = info.get("assoc_a_count", 0)
            if n_linhas == 0:
                continue
        else:
            continue

        bloco_h = HEADER_H + n_linhas * ROW_H + 2  # +2 pequena folga

        # Cabe na coluna atual?
        if cur_y + bloco_h > AREA_BOTTOM_Y:
            cur_col += 1
            if cur_col >= 2:
                # extrapolou 2 colunas — quebrar pra próxima página
                # (por enquanto não suportamos; aviso silencioso e cortamos)
                break
            cur_x = BLOCK_X_INIT + cur_col * BLOCK_W_MM
            cur_y = AREA_TOP_Y

        bubbles = []
        labels_header = ""

        if tipo == "multipla_escolha":
            labels_header = "A   B   C   D"
            for i, letra in enumerate(["A", "B", "C", "D"]):
                bx = cur_x + FIRST_BUBBLE_X + i * BUBBLE_SPACING
                by = cur_y + HEADER_H + 4
                bubbles.append({"label": letra, "x_mm": bx, "y_mm": by, "afirm": None, "item": None})

        elif tipo == "vf":
            labels_header = "V   F"
            for k in range(n_linhas):
                by = cur_y + HEADER_H + 4 + k * ROW_H
                for i, vf in enumerate(["V", "F"]):
                    bx = cur_x + FIRST_BUBBLE_X + i * BUBBLE_SPACING
                    bubbles.append({"label": vf, "x_mm": bx, "y_mm": by, "afirm": k, "item": None})

        elif tipo == "associacao":
            n_letras = info.get("assoc_b_count", 0)
            letras = [chr(97 + i) for i in range(n_letras)]
            labels_header = "   ".join(letras)
            for k in range(n_linhas):
                by = cur_y + HEADER_H + 4 + k * ROW_H
                for i, letra in enumerate(letras):
                    bx = cur_x + FIRST_BUBBLE_X + i * BUBBLE_SPACING
                    bubbles.append({"label": letra, "x_mm": bx, "y_mm": by, "afirm": None, "item": k})

        blocos.append({
            "num": info["num"],
            "questao_id": info["id"],
            "tipo": tipo,
            "top_y_mm": cur_y,
            "x_mm": cur_x,
            "height_mm": bloco_h,
            "header": labels_header,
            "bubbles": bubbles,
            "n_linhas": n_linhas,
        })

        cur_y += bloco_h + 2  # +2 espaço entre blocos

    return blocos


def _coletar_info_questoes_cartao(conn, prova_id):
    """Coleta info necessária pra layout do cartão a partir da prova."""
    questoes = conn.execute("""
        SELECT q.id, q.tipo
        FROM prova_questoes pq
        JOIN questoes q ON q.id = pq.questao_id
        WHERE pq.prova_id = ?
        ORDER BY pq.ordem
    """, (prova_id,)).fetchall()
    infos = []
    for idx, q in enumerate(questoes, start=1):
        tipo = q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha"
        info = {"id": q["id"], "num": idx, "tipo": tipo, "vf_count": 0, "assoc_a_count": 0, "assoc_b_count": 0}
        if tipo == "vf":
            info["vf_count"] = conn.execute("SELECT COUNT(*) AS c FROM vf_afirmacoes WHERE questao_id = ?", (q["id"],)).fetchone()["c"]
        elif tipo == "associacao":
            info["assoc_a_count"] = conn.execute("SELECT COUNT(*) AS c FROM assoc_itens_a WHERE questao_id = ?", (q["id"],)).fetchone()["c"]
            info["assoc_b_count"] = conn.execute("SELECT COUNT(*) AS c FROM assoc_itens_b WHERE questao_id = ?", (q["id"],)).fetchone()["c"]
        infos.append(info)
    return infos


def _gerar_cartao_resposta_pdf(apl, alunos, questoes_info):
    """Gera PDF com cartões resposta padronizados (um por aluno) para OMR posterior.
    questoes_info: lista de dicts retornada por _coletar_info_questoes_cartao()."""
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    blocos = _calcular_layout_cartao(questoes_info)
    n_questoes_no_cartao = len(blocos)
    bubble_radius = 2.2 * mm

    for aluno in alunos:
        # 4 marcadores de canto pra OMR (quadrados pretos)
        marker_size = 8 * mm
        margin = 10 * mm
        c.setFillColorRGB(0, 0, 0)
        c.rect(margin, height - margin - marker_size, marker_size, marker_size, fill=1, stroke=0)
        c.rect(width - margin - marker_size, height - margin - marker_size, marker_size, marker_size, fill=1, stroke=0)
        c.rect(margin, margin, marker_size, marker_size, fill=1, stroke=0)
        c.rect(width - margin - marker_size, margin, marker_size, marker_size, fill=1, stroke=0)

        # Cabeçalho de texto
        c.setFont("Helvetica-Bold", 14)
        titulo_str = apl["prova_titulo"][:60]
        c.drawString(30*mm, height - 25*mm, titulo_str)
        c.setFont("Helvetica", 10)
        c.drawString(30*mm, height - 32*mm, f"Turma: {apl['turma_nome']} ({apl['ano_letivo']})")
        c.setFont("Helvetica-Bold", 11)
        c.drawString(30*mm, height - 40*mm, f"Aluno: {aluno['nome']}")
        c.setFont("Helvetica", 9)
        num_str = f"Nº {aluno['numero']} · " if aluno["numero"] else ""
        c.drawString(30*mm, height - 46*mm, f"{num_str}Código: {aluno['codigo_unico']}")

        # QR Code: codifica aluno_id e aplicacao_id
        qr_data = f"CR:{aluno['id']}:{apl['id']}"
        qr_obj = qrcode.QRCode(box_size=10, border=1, error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr_obj.add_data(qr_data)
        qr_obj.make(fit=True)
        qr_img = qr_obj.make_image(fill_color="black", back_color="white")
        qr_buf = BytesIO()
        qr_img.save(qr_buf, format="PNG")
        qr_buf.seek(0)
        c.drawImage(ImageReader(qr_buf), width - 50*mm, height - 50*mm, width=30*mm, height=30*mm)

        # Instruções
        c.setFont("Helvetica", 8)
        c.drawString(30*mm, height - 55*mm, "Preencha com caneta preta. Pinte toda a bolha. Não use corretivo nem rasure.")

        # Renderiza blocos
        for blk in blocos:
            tipo = blk["tipo"]
            # Cabeçalho do bloco: "Q3 (V/F)"
            label_tipo = {"multipla_escolha": "", "vf": "(V/F)", "associacao": "(Assoc.)"}.get(tipo, "")
            c.setFont("Helvetica-Bold", 9)
            # Posição em PDF: y é "altura - y_mm" (PDF tem origem em baixo)
            head_y_pdf = height - blk["top_y_mm"] * mm
            c.drawString(blk["x_mm"] * mm, head_y_pdf, f"Q{blk['num']} {label_tipo}")

            # Numerações internas: V/F mostra "Q.1, Q.2..." e Associação mostra "1., 2., ..."
            if tipo == "vf":
                c.setFont("Helvetica", 8)
                for k in range(blk["n_linhas"]):
                    by_mm = blk["top_y_mm"] + 6 + 4 + k * 8
                    c.drawRightString((blk["x_mm"] + 16) * mm, (height - by_mm * mm) - 1,
                                       f"{blk['num']}.{k+1}")
            elif tipo == "associacao":
                c.setFont("Helvetica", 8)
                for k in range(blk["n_linhas"]):
                    by_mm = blk["top_y_mm"] + 6 + 4 + k * 8
                    c.drawRightString((blk["x_mm"] + 16) * mm, (height - by_mm * mm) - 1,
                                       f"{k+1}.")
            # Múltipla escolha: numeração "Q1" no cabeçalho do bloco já basta (sem duplicar com "1.")

            # Cabeçalho de letras: centralizadas EXATAMENTE acima de cada bolha (uma por coluna).
            # Pega primeira ocorrência de cada x_mm distinto (= primeira linha de bolhas) preservando ordem.
            seen_x = set()
            col_labels = []
            for b in blk["bubbles"]:
                if b["x_mm"] not in seen_x:
                    seen_x.add(b["x_mm"])
                    col_labels.append(b)
            c.setFont("Helvetica-Bold", 9)
            c.setFillColorRGB(0.35, 0.35, 0.35)
            for b in col_labels:
                bx_pdf = b["x_mm"] * mm
                label_y_pdf = height - (b["y_mm"] - 4.5) * mm
                c.drawCentredString(bx_pdf, label_y_pdf, b["label"])
            c.setFillColorRGB(0, 0, 0)

            # Bolhas
            for b in blk["bubbles"]:
                # Converte mm → PDF point (origem do PDF é canto inf esquerdo)
                bx_pdf = b["x_mm"] * mm
                by_pdf = height - b["y_mm"] * mm
                c.circle(bx_pdf, by_pdf, bubble_radius, stroke=1, fill=0)

        # Rodapé
        c.setFont("Helvetica", 7)
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.drawString(margin + marker_size + 4*mm, margin + 2*mm,
                     f"Cartão resposta · Aluno {aluno['id']} · Aplicação {apl['id']} · {n_questoes_no_cartao} questões")
        c.setFillColorRGB(0, 0, 0)

        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer


@app.get("/aplicacoes/{aplicacao_id}/cartao-resposta")
def cartao_resposta_pdf(aplicacao_id: int):
    """Gera PDF com cartões resposta de todos os alunos da turma."""
    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a JOIN provas p ON p.id = a.prova_id JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()
    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    alunos = conn.execute("SELECT * FROM alunos WHERE turma_id = ? ORDER BY numero, nome", (apl["turma_id"],)).fetchall()
    questoes = conn.execute("SELECT q.id FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id WHERE pq.prova_id = ? ORDER BY pq.ordem", (apl["prova_id"],)).fetchall()

    if not alunos:
        conn.close()
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Esta turma não tem alunos cadastrados.</p><a href="/aplicacoes/' + str(aplicacao_id) + '" class="btn">← Voltar</a></div>', active="aplicacoes"))
    if not questoes:
        conn.close()
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Esta prova não tem questões.</p><a href="/aplicacoes/' + str(aplicacao_id) + '" class="btn">← Voltar</a></div>', active="aplicacoes"))

    apl_dict = dict(apl)
    apl_dict["id"] = aplicacao_id
    questoes_info = _coletar_info_questoes_cartao(conn, apl["prova_id"])
    conn.close()
    buffer = _gerar_cartao_resposta_pdf(apl_dict, alunos, questoes_info)

    base_name = (apl["titulo"] or apl["prova_titulo"]).lower().replace(" ", "_")
    # Remove acentos e caracteres não-ASCII (Content-Disposition exige ASCII puro)
    import unicodedata as _ud
    base_name = _ud.normalize('NFKD', base_name).encode('ascii', 'ignore').decode('ascii')
    safe = "".join(c for c in base_name if c.isalnum() or c in "_-")[:40] or f"cartoes"
    filename = f"cartoes_resposta_{safe}_{aplicacao_id}.pdf"

    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


# ==========================================
#  FASE C2: OMR — LEITURA AUTOMÁTICA DE CARTÕES
# ==========================================

def _decode_image_universal(image_bytes, filename=""):
    """Decodifica imagem em bytes → numpy array BGR (cv2). Suporta JPG, PNG, HEIC.
    Retorna (img_np, erro_str). Se erro_str não for None, img_np é None.
    HEIC: tenta usar pillow_heif se instalado; se não, sugere conversão."""
    import cv2
    import numpy as np

    # Detectar HEIC pelos magic bytes (ftypheic, ftypheix, ftypmif1 nas posições 4-12)
    is_heic = False
    if len(image_bytes) > 12:
        marker = image_bytes[4:12]
        if b"ftyp" in marker and (b"heic" in marker or b"heix" in marker or b"mif1" in marker or b"heif" in marker):
            is_heic = True
    if not is_heic and filename and filename.lower().endswith((".heic", ".heif")):
        is_heic = True

    if is_heic:
        try:
            from pillow_heif import register_heif_opener
            from PIL import Image
            import io
            register_heif_opener()
            pil_img = Image.open(io.BytesIO(image_bytes))
            if pil_img.mode != "RGB":
                pil_img = pil_img.convert("RGB")
            arr = np.array(pil_img)
            img = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
            return img, None
        except ImportError:
            return None, ("Foto em formato HEIC (iPhone). Para o sistema ler HEIC, "
                          "rode no terminal do Codespaces: <code>pip install pillow-heif --break-system-packages</code> "
                          "e reinicie o servidor. Como alternativa imediata: no iPhone, vá em "
                          "<em>Ajustes → Câmera → Formatos → Mais Compatível</em> (passa a tirar em JPG).")
        except Exception as e:
            return None, f"Erro ao decodificar HEIC: {e}"

    # Caminho padrão JPG/PNG via cv2
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None:
        return None, "Formato de imagem não reconhecido. Use JPG, PNG ou HEIC."
    return img, None


def _processar_cartao_resposta(image_bytes, n_questoes_esperado, filename="", threshold_modo="normal", questoes_info=None):
    """Processa imagem de cartão resposta preenchido:
    - Detecta marcadores de canto
    - Corrige perspectiva
    - Lê QR Code (CR:aluno_id:aplicacao_id)
    - Detecta bolhas marcadas
    Retorna dict com success, aluno_id, aplicacao_id, answers, warnings, preview_base64.
    """
    import cv2
    import numpy as np
    import base64

    # Decode image (suporta JPG/PNG/HEIC)
    img, erro = _decode_image_universal(image_bytes, filename)
    if img is None:
        return {"success": False, "error": erro or "Erro ao abrir imagem."}

    h, w = img.shape[:2]
    if h < 400 or w < 300:
        return {"success": False, "error": f"Imagem muito pequena ({w}×{h}px). Use uma foto com pelo menos 400×300px."}

    # === STEP 1: Detect corner markers ===
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    candidates = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < 100:
            continue
        x, y, ww, hh = cv2.boundingRect(cnt)
        if ww == 0 or hh == 0:
            continue
        aspect = ww / hh
        if not (0.7 < aspect < 1.3):
            continue
        rect_area = ww * hh
        if area / rect_area < 0.7:
            continue
        if ww < 15 or ww > w * 0.15:
            continue
        cx, cy = x + ww/2, y + hh/2
        candidates.append((cx, cy, area))

    if len(candidates) < 4:
        return {"success": False, "error": f"Detectados apenas {len(candidates)} candidatos a marcador (esperados 4 ou mais). Verifique se a foto mostra a folha inteira, está bem iluminada e nítida."}

    # Identificar o marcador mais próximo de cada canto da imagem.
    # Os QR finder patterns ficam todos juntos perto do QR; nossos marcadores ficam
    # nos cantos extremos da folha. Por isso, pra cada canto da imagem, pegar o
    # candidato mais próximo (dentro de uma zona razoável) é mais robusto que pegar
    # os 4 maiores.
    max_corner_dist = (w * w + h * h) ** 0.5 * 0.30  # diagonal × 30%

    def closest_in_zone(corner_x, corner_y):
        best = None
        best_d = float("inf")
        for cx, cy, a in candidates:
            d = ((cx - corner_x) ** 2 + (cy - corner_y) ** 2) ** 0.5
            if d > max_corner_dist:
                continue
            if d < best_d:
                best = (cx, cy)
                best_d = d
        return best

    tl = closest_in_zone(0, 0)
    tr = closest_in_zone(w, 0)
    bl = closest_in_zone(0, h)
    br = closest_in_zone(w, h)

    if not all([tl, tr, bl, br]):
        found = sum(1 for m in [tl, tr, bl, br] if m)
        return {"success": False, "error": f"Não foi possível identificar os 4 marcadores de canto da folha. Encontrei {found} de 4. Verifique se a foto inclui os 4 cantos da folha, com boa iluminação."}

    # Garantir que os 4 marcadores são distintos (não o mesmo candidato pego 2x)
    pontos_set = {tl, tr, bl, br}
    if len(pontos_set) < 4:
        return {"success": False, "error": "Marcadores de canto sobrepostos detectados. A folha pode estar parcialmente fora do enquadramento."}

    # === STEP 2: Perspective transform to canonical A4 ===
    # Canonical A4 at ~144 DPI: 1191 x 1684 px
    canon_w, canon_h = 1191, 1684
    # Markers are at 10mm (margin) + 4mm (half of 8mm marker) = 14mm from edges
    margin_canon = int(14 / 210 * canon_w)  # ~79
    src_pts = np.float32([tl, tr, bl, br])
    dst_pts = np.float32([
        [margin_canon, margin_canon],
        [canon_w - margin_canon, margin_canon],
        [margin_canon, canon_h - margin_canon],
        [canon_w - margin_canon, canon_h - margin_canon],
    ])
    M = cv2.getPerspectiveTransform(src_pts, dst_pts)
    warped = cv2.warpPerspective(img, M, (canon_w, canon_h))
    warped_gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)

    # === STEP 3: Read QR code ===
    qr_detector = cv2.QRCodeDetector()
    qr_data = ""
    try:
        result = qr_detector.detectAndDecode(warped)
        qr_data = result[0] if result else ""
    except Exception:
        qr_data = ""

    if not qr_data:
        # Fallback: try in original image
        try:
            result = qr_detector.detectAndDecode(img)
            qr_data = result[0] if result else ""
        except Exception:
            qr_data = ""

    if not qr_data or not qr_data.startswith("CR:"):
        return {"success": False, "error": f"Não foi possível ler o QR Code do cartão. Tente uma foto mais nítida ou com mais resolução. (Dado lido: '{qr_data[:30]}')"}

    try:
        parts = qr_data.split(":")
        if len(parts) != 3:
            raise ValueError("formato inesperado")
        aluno_id = int(parts[1])
        aplicacao_id_qr = int(parts[2])
    except (ValueError, AttributeError) as e:
        return {"success": False, "error": f"QR Code com formato inválido: '{qr_data}' ({e})"}

    # === STEP 4: Sample bubbles ===
    def mm_to_px_x(mm_x):
        return int(mm_x / 210 * canon_w)

    def mm_to_px_y(mm_y_from_bottom):
        return int((297 - mm_y_from_bottom) / 297 * canon_h)

    # Importante: mm_to_px_y aceita coord medida a partir do topo da página (igual layout do cartão).
    # Se o sistema usar "mm_y_from_bottom" em coord, converter por subtração.
    def mm_top_to_px_y(mm_y_from_top):
        return int(mm_y_from_top / 297 * canon_h)

    # === NOVO: layout dinâmico baseado em questoes_info se foi passado ===
    # Fallback: legado, layout fixo A/B/C/D pra n_questoes_esperado
    bubble_positions = []  # (q_num, letra/label, x_px, y_px, afirm, item, tipo)

    if questoes_info is not None and len(questoes_info) > 0:
        # Usa o mesmo cálculo do gerador pra ter mesmas coordenadas
        blocos = _calcular_layout_cartao(questoes_info)
        bubble_radius_mm = 2.2
        bubble_radius_px = int(bubble_radius_mm / 210 * canon_w)
        for blk in blocos:
            for b in blk["bubbles"]:
                x_px = mm_to_px_x(b["x_mm"])
                y_px = mm_top_to_px_y(b["y_mm"])
                bubble_positions.append({
                    "q": blk["num"], "tipo": blk["tipo"],
                    "label": b["label"], "afirm": b["afirm"], "item": b["item"],
                    "x": x_px, "y": y_px
                })
    else:
        # Modo legado: layout antigo (1 ou 2 colunas, A/B/C/D só)
        n_cols = 1 if n_questoes_esperado <= 25 else 2
        questions_per_col = (n_questoes_esperado + n_cols - 1) // n_cols
        block_x_mm_start = 55 if n_cols == 1 else 25
        block_width_mm = 80
        bubble_radius_mm = 2.5
        bubble_radius_px = int(bubble_radius_mm / 210 * canon_w)
        col_letter_spacing_mm = 12
        start_y_mm = 222
        row_height_mm = 8

        for col_idx in range(n_cols):
            block_x_mm = block_x_mm_start + col_idx * block_width_mm
            start_q = col_idx * questions_per_col
            end_q = min(start_q + questions_per_col, n_questoes_esperado)
            for offset in range(end_q - start_q):
                q_num = start_q + offset + 1
                y_mm = start_y_mm - offset * row_height_mm
                y_px = mm_to_px_y(y_mm)
                for i, letra in enumerate(["A", "B", "C", "D"]):
                    x_mm = block_x_mm + 22 + i * col_letter_spacing_mm
                    x_px = mm_to_px_x(x_mm)
                    bubble_positions.append({
                        "q": q_num, "tipo": "multipla_escolha",
                        "label": letra, "afirm": None, "item": None,
                        "x": x_px, "y": y_px
                    })

    # Sample darkness for each bubble
    sample_radius = max(3, bubble_radius_px - 2)
    bubble_data = []
    for bp in bubble_positions:
        x, y = bp["x"], bp["y"]
        y1 = max(0, y - sample_radius)
        y2 = min(canon_h, y + sample_radius)
        x1 = max(0, x - sample_radius)
        x2 = min(canon_w, x + sample_radius)
        region = warped_gray[y1:y2, x1:x2]
        if region.size == 0:
            continue
        mean = float(region.mean())
        bubble_data.append({**bp, "mean": mean, "marked": False})

    # Agrupa por (questao, afirmação/item) — cada grupo decide independentemente
    # Pra multipla_escolha: agrupar só por q
    # Pra vf: agrupar por (q, afirm) — 2 bolhas por grupo (V/F)
    # Pra associacao: agrupar por (q, item) — N bolhas por grupo
    grupos = {}
    for b in bubble_data:
        if b["tipo"] == "multipla_escolha":
            key = (b["q"], None, None)
        elif b["tipo"] == "vf":
            key = (b["q"], b["afirm"], None)
        elif b["tipo"] == "associacao":
            key = (b["q"], None, b["item"])
        else:
            continue
        grupos.setdefault(key, []).append(b)

    # Thresholds
    if threshold_modo == "permissivo":
        DARK_THRESHOLD = 140
        AMBIGUOUS_THRESHOLD = 165
        LIGHT_THRESHOLD = 200
    else:
        DARK_THRESHOLD = 110
        AMBIGUOUS_THRESHOLD = 140
        LIGHT_THRESHOLD = 180

    # answers estruturada por tipo:
    # - multipla_escolha: { q_num: "A"/None }
    # - vf: { q_num: {"0": "V", "1": "F", ...} }
    # - associacao: { q_num: {"0": "b", "1": "a", ...} }
    answers = {}
    warnings = []
    for (q_num, afirm, item), bubbles in sorted(grupos.items(), key=lambda x: (x[0][0], x[0][1] if x[0][1] is not None else -1, x[0][2] if x[0][2] is not None else -1)):
        bubbles.sort(key=lambda b: b["mean"])
        darkest = bubbles[0]
        second = bubbles[1] if len(bubbles) > 1 else None
        tipo = darkest["tipo"]

        # Decide marcação deste grupo
        if darkest["mean"] > LIGHT_THRESHOLD:
            marcado = None
        elif darkest["mean"] < DARK_THRESHOLD:
            marcado = darkest["label"]
            darkest["marked"] = True
            if second and second["mean"] < AMBIGUOUS_THRESHOLD:
                ctx = f"Q{q_num}"
                if afirm is not None: ctx += f".{afirm+1}"
                if item is not None: ctx += f" item {item+1}"
                warnings.append(f"{ctx}: marcação ambígua entre {darkest['label']} e {second['label']} (confira)")
        else:
            marcado = darkest["label"]
            darkest["marked"] = True
            ctx = f"Q{q_num}"
            if afirm is not None: ctx += f".{afirm+1}"
            if item is not None: ctx += f" item {item+1}"
            warnings.append(f"{ctx}: marca fraca em {darkest['label']} (confira)")

        # Salva conforme tipo
        if tipo == "multipla_escolha":
            answers[q_num] = marcado
        elif tipo == "vf":
            if q_num not in answers or not isinstance(answers[q_num], dict):
                answers[q_num] = {}
            answers[q_num][str(afirm)] = marcado
        elif tipo == "associacao":
            if q_num not in answers or not isinstance(answers[q_num], dict):
                answers[q_num] = {}
            answers[q_num][str(item)] = marcado

    # === STEP 5: Build preview with overlays ===
    preview = warped.copy()
    for b in bubble_data:
        if b["marked"]:
            cv2.circle(preview, (b["x"], b["y"]), bubble_radius_px + 1, (0, 200, 0), 3)
        else:
            cv2.circle(preview, (b["x"], b["y"]), bubble_radius_px, (180, 180, 180), 1)

    # Resize for HTML (max 800px wide)
    if canon_w > 800:
        scale = 800 / canon_w
        new_w, new_h = int(canon_w * scale), int(canon_h * scale)
        preview = cv2.resize(preview, (new_w, new_h), interpolation=cv2.INTER_AREA)

    _, encoded = cv2.imencode('.jpg', preview, [cv2.IMWRITE_JPEG_QUALITY, 70])
    preview_b64 = base64.b64encode(encoded.tobytes()).decode()

    return {
        "success": True,
        "aluno_id": aluno_id,
        "aplicacao_id_qr": aplicacao_id_qr,
        "qr_data": qr_data,
        "answers": answers,
        "warnings": warnings,
        "preview_base64": preview_b64,
    }


@app.get("/aplicacoes/{aplicacao_id}/escanear", response_class=HTMLResponse)
def form_escanear(aplicacao_id: int):
    """Formulário pra upload de foto do cartão resposta."""
    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a JOIN provas p ON p.id = a.prova_id JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()
    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)
    conn.close()

    content = f"""
        <div class="page-header">
            <h1>📷 Escanear cartão resposta</h1>
            <p class="subtitle">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]})</p>
            <div class="page-actions"><a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar</a></div>
        </div>

        <div class="tip">
            <strong>Dicas pra uma boa leitura:</strong>
            <ul style="margin:8px 0 0 18px;">
                <li>Tire a foto com boa luz, sem sombras sobre a folha</li>
                <li>Mantenha o celular paralelo à folha (sem inclinar)</li>
                <li>Inclua os 4 marcadores pretos dos cantos no enquadramento</li>
                <li>O QR Code precisa estar legível (sem reflexo nem desfoque)</li>
                <li><strong>HEIC do iPhone agora é suportado</strong> (se o sistema acusar, instala <code>pillow-heif</code>)</li>
            </ul>
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:18px; margin-top:24px;">

            <form id="form-single" action="/aplicacoes/{aplicacao_id}/escanear" method="post" enctype="multipart/form-data" style="background:var(--bg-subtle); padding:18px; border-radius:8px;">
                <h3 style="margin-top:0;">📷 Um cartão por vez</h3>
                <p class="muted-line" style="font-size:13px;">Recomendado pra correção ao vivo, durante a aplicação.</p>
                <label>Foto<input type="file" name="foto" accept="image/*" capture="environment" required></label>
                <p class="muted-line" style="font-size:11px;">No celular abre a câmera direto.</p>
                <button type="submit" class="btn btn-primary" style="width:100%;">Processar 1 foto</button>
            </form>

            <form id="form-lote" action="/aplicacoes/{aplicacao_id}/escanear-lote" method="post" enctype="multipart/form-data" style="background:var(--bg-subtle); padding:18px; border-radius:8px;">
                <h3 style="margin-top:0;">📁 Lote (várias de uma vez)</h3>
                <p class="muted-line" style="font-size:13px;">Recomendado quando você já tem todas as fotos prontas (galeria).</p>
                <label>Fotos ou PDF<input type="file" name="fotos" accept="image/*,.pdf" multiple required></label>
                <p class="muted-line" style="font-size:11px;">Selecione imagens (JPEG/HEIC) <strong>ou</strong> um PDF com várias páginas.</p>
                <button type="submit" class="btn btn-primary" style="width:100%;">Processar lote</button>
            </form>

            <script>
            (function() {{
                function travar(form) {{
                    form.addEventListener('submit', function() {{
                        var btn = form.querySelector('button[type="submit"]');
                        if (btn) {{
                            btn.disabled = true;
                            btn.textContent = '⏳ Processando…';
                            btn.style.opacity = '0.7';
                        }}
                    }});
                }}
                var fs = document.getElementById('form-single');
                var fl = document.getElementById('form-lote');
                if (fs) travar(fs);
                if (fl) travar(fl);
            }})();
            </script>

        </div>
    """
    return render_page("Escanear cartão", content, active="aplicacoes")


@app.post("/aplicacoes/{aplicacao_id}/escanear", response_class=HTMLResponse)
async def processar_escaneamento(aplicacao_id: int, foto: UploadFile = File(...)):
    """Recebe foto, processa OMR e mostra tela de revisão antes de salvar."""
    image_bytes = await foto.read()
    if not image_bytes:
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Arquivo vazio.</p><a href="javascript:history.back()" class="btn">← Voltar</a></div>', active="aplicacoes"))

    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome
        FROM aplicacoes a JOIN provas p ON p.id = a.prova_id JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()
    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    questoes = conn.execute(
        "SELECT q.id FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id WHERE pq.prova_id = ? ORDER BY pq.ordem",
        (apl["prova_id"],)
    ).fetchall()
    n_questoes = len(questoes)
    questoes_info = _coletar_info_questoes_cartao(conn, apl["prova_id"])

    result = _processar_cartao_resposta(image_bytes, n_questoes, filename=foto.filename or "", questoes_info=questoes_info)

    if not result["success"]:
        conn.close()
        content = f"""
            <div class="page-header">
                <h1>❌ Erro na leitura do cartão</h1>
            </div>
            <div style="border:1px solid var(--red); background:var(--red-bg); padding:16px; border-radius:6px; margin:16px 0; color:var(--red);">
                <strong>Problema:</strong> {result.get("error", "Erro desconhecido")}
            </div>
            <p>Tente novamente com uma foto mais nítida, com melhor iluminação ou de um ângulo mais frontal.</p>
            <div class="page-actions">
                <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn btn-primary">📷 Tentar outra foto</a>
                <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar para a aplicação</a>
            </div>
        """
        return render_page("Erro no escaneamento", content, active="aplicacoes")

    # Validar que o aluno pertence à turma desta aplicação
    aluno = conn.execute(
        "SELECT * FROM alunos WHERE id = ? AND turma_id = ?",
        (result["aluno_id"], apl["turma_id"])
    ).fetchone()

    if not aluno:
        conn.close()
        content = f"""
            <div class="page-header"><h1>⚠️ Cartão de outra turma</h1></div>
            <div style="border:1px solid var(--red); background:var(--red-bg); padding:16px; border-radius:6px; color:var(--red);">
                <p>O QR Code deste cartão aponta para o aluno <code>{result["aluno_id"]}</code>, que não pertence à turma <strong>{apl["turma_nome"]}</strong> desta aplicação.</p>
                <p>Verifique se você está na aplicação certa antes de escanear.</p>
            </div>
            <div class="page-actions">
                <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">📷 Tentar outra foto</a>
                <a href="/aplicacoes" class="btn">Lista de aplicações</a>
            </div>
        """
        return render_page("Cartão de outra turma", content, active="aplicacoes")

    if result["aplicacao_id_qr"] != aplicacao_id:
        conn.close()
        content = f"""
            <div class="page-header"><h1>⚠️ Cartão de outra aplicação</h1></div>
            <div style="border:1px solid var(--red); background:var(--red-bg); padding:16px; border-radius:6px; color:var(--red);">
                <p>Este cartão foi gerado para a aplicação <code>{result["aplicacao_id_qr"]}</code>, mas você está na aplicação <code>{aplicacao_id}</code>.</p>
            </div>
            <div class="page-actions">
                <a href="/aplicacoes/{result["aplicacao_id_qr"]}/escanear" class="btn btn-primary">Ir para aplicação {result["aplicacao_id_qr"]}</a>
                <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">Tentar outra foto</a>
            </div>
        """
        return render_page("Cartão de outra aplicação", content, active="aplicacoes")

    # Verificar se já existe entrega pra esse aluno (override?)
    ja_entregue = conn.execute(
        "SELECT finalizada_em FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?",
        (aplicacao_id, result["aluno_id"])
    ).fetchone()
    conn.close()

    answers = result["answers"]

    # Build tabela editável de respostas
    rows_html = ""
    for q_num in range(1, n_questoes + 1):
        detected = answers.get(q_num)
        cells = ""
        for letra in ["A", "B", "C", "D"]:
            checked = " checked" if detected == letra else ""
            cells += f'<td style="text-align:center;"><input type="radio" name="q_{q_num}" value="{letra}"{checked}></td>'
        # Em branco option
        em_branco_checked = " checked" if detected is None else ""
        cells += f'<td style="text-align:center; background:var(--bg-muted);"><input type="radio" name="q_{q_num}" value=""{em_branco_checked}></td>'
        marca_status = f"Detectado: <strong>{detected}</strong>" if detected else '<span style="color:var(--text-muted);">Em branco</span>'
        rows_html += f'<tr><td style="padding:6px 8px;"><strong>Q{q_num}</strong></td>{cells}<td style="font-size:11px; color:var(--text-muted); padding:0 8px;">{marca_status}</td></tr>'

    avisos_html = ""
    if result["warnings"]:
        items = "".join(f"<li>{w}</li>" for w in result["warnings"])
        avisos_html = f'<div style="border:1px solid var(--orange); background:var(--orange-bg); padding:12px; border-radius:6px; margin:16px 0; color:var(--orange);"><strong>⚠️ Avisos da leitura:</strong><ul style="margin:6px 0 0 18px;">{items}</ul></div>'

    override_aviso = ""
    if ja_entregue:
        override_aviso = f'<div style="border:1px solid var(--orange); background:var(--orange-bg); padding:12px; border-radius:6px; margin:16px 0; color:var(--orange);"><strong>⚠️ Atenção:</strong> este aluno já tem entrega registrada ({ja_entregue["finalizada_em"]}). Confirmar irá <strong>sobrescrever</strong> as respostas anteriores.</div>'

    content = f"""
        <div class="page-header">
            <h1>Revisão da leitura</h1>
            <p class="subtitle">{apl["prova_titulo"]} · Aluno: <strong>{aluno["nome"]}</strong> (Nº {aluno["numero"] or "—"}, Código {aluno["codigo_unico"]})</p>
        </div>

        {avisos_html}
        {override_aviso}

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:24px; align-items:flex-start;">
            <div>
                <h2 style="margin-top:0;">Imagem processada</h2>
                <p class="muted-line">Bolhas detectadas como marcadas estão em verde. Confira se está correto antes de confirmar.</p>
                <img src="data:image/jpeg;base64,{result["preview_base64"]}" style="max-width:100%; border:1px solid var(--border); border-radius:6px;">
            </div>

            <div>
                <h2 style="margin-top:0;">Respostas detectadas</h2>
                <p class="muted-line">Você pode corrigir qualquer marcação antes de salvar.</p>
                <form action="/aplicacoes/{aplicacao_id}/escanear/confirmar" method="post">
                    <input type="hidden" name="aluno_id" value="{result["aluno_id"]}">
                    <table style="width:100%; border-collapse:collapse; font-size:13px;">
                        <thead>
                            <tr style="background:var(--bg-subtle);">
                                <th style="padding:6px;">Q</th>
                                <th style="padding:6px;">A</th>
                                <th style="padding:6px;">B</th>
                                <th style="padding:6px;">C</th>
                                <th style="padding:6px;">D</th>
                                <th style="padding:6px;">∅</th>
                                <th style="padding:6px;">Detectado</th>
                            </tr>
                        </thead>
                        <tbody>{rows_html}</tbody>
                    </table>
                    <div class="page-actions" style="margin-top:16px;">
                        <button type="submit" class="btn btn-primary">✓ Confirmar e salvar</button>
                        <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">📷 Tentar outra foto</a>
                    </div>
                </form>
            </div>
        </div>
    """
    return render_page("Revisão", content, active="aplicacoes")


@app.post("/aplicacoes/{aplicacao_id}/escanear/confirmar", response_class=HTMLResponse)
async def confirmar_escaneamento(aplicacao_id: int, request: Request, aluno_id: int = Form(...)):
    """Salva as respostas confirmadas pelo professor após revisão."""
    form = await request.form()

    conn = get_db()
    apl = conn.execute("SELECT * FROM aplicacoes WHERE id = ?", (aplicacao_id,)).fetchone()
    aluno = conn.execute("SELECT * FROM alunos WHERE id = ? AND turma_id = ?", (aluno_id, apl["turma_id"])).fetchone() if apl else None

    if not apl or not aluno:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    questoes = conn.execute(
        "SELECT q.id FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id WHERE pq.prova_id = ? ORDER BY pq.ordem",
        (apl["prova_id"],)
    ).fetchall()
    questao_ids = [q["id"] for q in questoes]

    # Limpar respostas antigas deste aluno nesta aplicação (override completo)
    conn.execute("DELETE FROM respostas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id))

    # Inserir respostas novas
    for q_num, q_id in enumerate(questao_ids, start=1):
        letra = form.get(f"q_{q_num}", "").strip()
        if letra in ("A", "B", "C", "D"):
            conn.execute(
                "INSERT INTO respostas (aplicacao_id, aluno_id, questao_id, alternativa_letra) VALUES (?, ?, ?, ?)",
                (aplicacao_id, aluno_id, q_id, letra)
            )

    # Inserir ou atualizar entrega
    existing = conn.execute("SELECT id FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id)).fetchone()
    if not existing:
        conn.execute("INSERT INTO entregas (aplicacao_id, aluno_id) VALUES (?, ?)", (aplicacao_id, aluno_id))
    else:
        conn.execute("UPDATE entregas SET finalizada_em = CURRENT_TIMESTAMP WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id))

    conn.commit()

    score, total = _calcular_nota(conn, aplicacao_id, aluno_id)
    conn.close()

    content = f"""
        <div class="page-header"><h1>✅ Cartão registrado</h1></div>
        <div style="border:1px solid var(--green); background:var(--green-bg); padding:20px; border-radius:6px; margin:16px 0; color:var(--green);">
            <p style="margin:0;"><strong>Aluno:</strong> {aluno["nome"]} (Nº {aluno["numero"] or "—"})</p>
            <p style="margin:8px 0 0;"><strong>Nota:</strong> <span style="font-size:24px; font-weight:600;">{score}/{total}</span> ({(score/total*100 if total > 0 else 0):.0f}%)</p>
        </div>
        <div class="page-actions">
            <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn btn-primary">📷 Escanear próximo cartão</a>
            <a href="/aplicacoes/{aplicacao_id}/aluno/{aluno_id}" class="btn">Ver detalhe da prova deste aluno</a>
            <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar para a aplicação</a>
        </div>
    """
    return render_page("Cartão registrado", content, active="aplicacoes")


# ==========================================
#  FASE R1: GESTÃO COMPLETA DE TURMAS E ALUNOS
# ==========================================

@app.post("/turmas/{turma_id}/deletar")
def deletar_turma(request: Request, turma_id: int):
    """Cascade delete. Restrito a admin."""
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    turma = conn.execute("SELECT * FROM turmas WHERE id = ?", (turma_id,)).fetchone()
    if not turma:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)

    # Aplicações desta turma → suas respostas, entregas, e a aplicação
    aplicacoes_ids = [r["id"] for r in conn.execute("SELECT id FROM aplicacoes WHERE turma_id = ?", (turma_id,)).fetchall()]
    for apl_id in aplicacoes_ids:
        conn.execute("DELETE FROM respostas WHERE aplicacao_id = ?", (apl_id,))
        conn.execute("DELETE FROM entregas WHERE aplicacao_id = ?", (apl_id,))
        conn.execute("DELETE FROM aplicacoes WHERE id = ?", (apl_id,))

    # Alunos atualmente nesta turma → suas respostas/entregas em quaisquer aplicações + o aluno
    alunos_ids = [r["id"] for r in conn.execute("SELECT id FROM alunos WHERE turma_id = ?", (turma_id,)).fetchall()]
    for aluno_id in alunos_ids:
        conn.execute("DELETE FROM respostas WHERE aluno_id = ?", (aluno_id,))
        conn.execute("DELETE FROM entregas WHERE aluno_id = ?", (aluno_id,))
        conn.execute("DELETE FROM alunos WHERE id = ?", (aluno_id,))

    conn.execute("DELETE FROM turmas WHERE id = ?", (turma_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/turmas", status_code=303)


@app.get("/alunos/{aluno_id}/editar", response_class=HTMLResponse)
def form_editar_aluno(request: Request, aluno_id: int):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    aluno = conn.execute("SELECT a.*, t.nome AS turma_nome, t.id AS turma_id_atual FROM alunos a JOIN turmas t ON t.id = a.turma_id WHERE a.id = ?", (aluno_id,)).fetchone()
    if not aluno:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)
    conn.close()

    racas_opts = '<option value="">Não informada</option>' + "".join(
        f'<option value="{r}"{(" selected" if aluno["raca"] == r else "")}>{r}</option>' for r in RACAS
    )

    content = f"""
        <div class="page-header">
            <h1>Editar aluno</h1>
            <p class="subtitle">Turma atual: <strong>{aluno["turma_nome"]}</strong> · Código único: <code>{aluno["codigo_unico"]}</code> (imutável)</p>
        </div>
        <div class="tip">O <strong>código único</strong> não pode ser alterado — ele é usado nos QR Codes já distribuídos. Para mudar a turma, use a opção <strong>Transferir</strong>.</div>

        <form action="/alunos/{aluno_id}/editar" method="post">
            <div style="display:grid; grid-template-columns: 100px 1fr; gap:12px;">
                <label>Número<input type="number" name="numero" value="{aluno["numero"] or ''}" min="1"></label>
                <label>Nome<input type="text" name="nome" required value="{aluno["nome"]}"></label>
            </div>
            <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:12px;">
                <label>Raça<select name="raca">{racas_opts}</select></label>
                <label>E-mail<input type="email" name="email" value="{aluno["email"] or ''}"></label>
                <label>Data de nascimento<input type="date" name="data_nascimento" value="{aluno["data_nascimento"] or ''}"></label>
            </div>
            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Salvar alterações</button>
                <a href="/turmas/{aluno['turma_id_atual']}" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return render_page("Editar aluno", content, active="turmas")


@app.post("/alunos/{aluno_id}/editar")
def atualizar_aluno(
    aluno_id: int,
    nome: str = Form(...),
    numero: Optional[int] = Form(None),
    raca: str = Form(""),
    email: str = Form(""),
    data_nascimento: str = Form(""),
):
    conn = get_db()
    aluno = conn.execute("SELECT turma_id FROM alunos WHERE id = ?", (aluno_id,)).fetchone()
    if not aluno:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)
    conn.execute(
        "UPDATE alunos SET nome = ?, numero = ?, raca = ?, email = ?, data_nascimento = ? WHERE id = ?",
        (nome.strip(), numero, raca.strip() or None, email.strip() or None, data_nascimento.strip() or None, aluno_id),
    )
    conn.commit()
    turma_id = aluno["turma_id"]
    conn.close()
    return RedirectResponse(f"/turmas/{turma_id}", status_code=303)


@app.post("/alunos/{aluno_id}/deletar", response_class=HTMLResponse)
def deletar_aluno(request: Request, aluno_id: int, forcar: int = 0):
    """Por padrão mostra confirmação se há entregas. Com ?forcar=1 apaga em cascade.
    Restrito a admin (turmas/alunos são da escola)."""
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    aluno = conn.execute("SELECT turma_id, nome FROM alunos WHERE id = ?", (aluno_id,)).fetchone()
    if not aluno:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)
    turma_id = aluno["turma_id"]

    n_entregas = conn.execute("SELECT COUNT(*) AS c FROM entregas WHERE aluno_id = ?", (aluno_id,)).fetchone()["c"]

    if n_entregas > 0 and not forcar:
        # Mostra tela de confirmação com opção de forçar
        n_respostas = conn.execute("SELECT COUNT(*) AS c FROM respostas WHERE aluno_id = ?", (aluno_id,)).fetchone()["c"]
        conn.close()
        content = f"""
            <div class="page-header"><h1>⚠️ Confirmação necessária</h1></div>

            <div style="border:1px solid var(--orange); background:var(--orange-bg); padding:16px; border-radius:6px; color:var(--orange);">
                <p style="margin:0;"><strong>{aluno["nome"]}</strong> tem <strong>{n_entregas} entrega(s)</strong> e <strong>{n_respostas} resposta(s)</strong> registradas no histórico.</p>
            </div>

            <div style="margin-top:18px;">
                <h3 style="margin-bottom:8px;">Opções:</h3>

                <div style="border:1px solid var(--border); padding:14px; border-radius:6px; margin-bottom:10px;">
                    <strong>Transferir para outra turma</strong> (recomendado se ele só mudou de turma)
                    <p style="margin:6px 0 10px 0; font-size:13px; color:var(--text-muted);">Preserva todo o histórico. Você pode criar uma turma "Inativos 2026" e mover ele pra lá.</p>
                    <a href="/alunos/{aluno_id}/transferir" class="btn">→ Ir para transferência</a>
                </div>

                <div style="border:1px solid var(--red); padding:14px; border-radius:6px; background:var(--red-bg); color:var(--red);">
                    <strong>Excluir definitivamente</strong> (use quando o aluno saiu da escola e o histórico não importa mais)
                    <p style="margin:6px 0 10px 0; font-size:13px;">
                        ⚠ Esta ação <strong>apaga permanentemente</strong>:<br>
                        • O cadastro do aluno<br>
                        • Todas as {n_respostas} respostas em provas/tarefas<br>
                        • Todas as {n_entregas} entregas registradas<br>
                        • As notas calculadas dessas aplicações deixarão de existir<br>
                        <strong>Não há como recuperar.</strong>
                    </p>
                    <form action="/alunos/{aluno_id}/deletar?forcar=1" method="post" style="display:inline; margin:0;"
                          onsubmit="return confirm('CONFIRMAÇÃO FINAL\\n\\nVocê está prestes a EXCLUIR PERMANENTEMENTE o aluno {aluno["nome"]} e TODOS os seus dados ({n_entregas} entregas, {n_respostas} respostas).\\n\\nEsta ação NÃO PODE ser desfeita.\\n\\nDeseja prosseguir?');">
                        <button type="submit" class="btn" style="background:var(--red); color:white; border-color:var(--red);">
                            Sim, excluir tudo definitivamente
                        </button>
                    </form>
                </div>
            </div>

            <div class="page-actions" style="margin-top:18px;">
                <a href="/turmas/{turma_id}" class="btn">← Voltar para a turma</a>
            </div>
        """
        return render_page("Confirmar exclusão", content, active="turmas")

    # Excluir em cascade (sem entregas OU forçado pelo botão de confirmação)
    conn.execute("DELETE FROM respostas WHERE aluno_id = ?", (aluno_id,))
    conn.execute("DELETE FROM entregas WHERE aluno_id = ?", (aluno_id,))
    conn.execute("DELETE FROM alunos WHERE id = ?", (aluno_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/turmas/{turma_id}", status_code=303)


@app.get("/alunos/{aluno_id}/transferir", response_class=HTMLResponse)
def form_transferir_aluno(request: Request, aluno_id: int):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    aluno = conn.execute("SELECT a.*, t.nome AS turma_nome_atual FROM alunos a JOIN turmas t ON t.id = a.turma_id WHERE a.id = ?", (aluno_id,)).fetchone()
    if not aluno:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)
    outras_turmas = conn.execute("SELECT * FROM turmas WHERE id != ? ORDER BY ano_letivo DESC, nome", (aluno["turma_id"],)).fetchall()
    conn.close()

    if not outras_turmas:
        content = f"""
            <div class="page-header"><h1>Transferir {aluno["nome"]}</h1></div>
            <div class="empty">
                <p>Não há outra turma cadastrada para transferir. Crie uma turma de destino primeiro.</p>
                <a href="/turmas/nova" class="btn btn-primary">Criar nova turma</a>
                <a href="/turmas/{aluno['turma_id']}" class="btn">← Voltar</a>
            </div>
        """
        return render_page("Transferir aluno", content, active="turmas")

    options = "".join(
        f'<option value="{t["id"]}">{t["nome"]} ({t["ano_letivo"]})</option>' for t in outras_turmas
    )

    content = f"""
        <div class="page-header">
            <h1>Transferir aluno</h1>
            <p class="subtitle"><strong>{aluno["nome"]}</strong> · Turma atual: {aluno["turma_nome_atual"]}</p>
        </div>
        <div class="tip">
            Após a transferência:
            <ul style="margin:8px 0 0 18px;">
                <li>O aluno aparece na lista da nova turma</li>
                <li>Aplicações já feitas na turma anterior continuam mostrando suas notas e respostas</li>
                <li>O <code>código único</code> e os QR Codes já impressos continuam válidos</li>
                <li>Esta ação pode ser desfeita transferindo de volta a qualquer momento</li>
            </ul>
        </div>
        <form action="/alunos/{aluno_id}/transferir" method="post" style="margin-top:24px;">
            <label>Nova turma<select name="nova_turma_id" required>{options}</select></label>
            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Transferir</button>
                <a href="/turmas/{aluno['turma_id']}" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return render_page("Transferir aluno", content, active="turmas")


@app.post("/alunos/{aluno_id}/transferir")
def transferir_aluno(request: Request, aluno_id: int, nova_turma_id: int = Form(...)):
    _r = _require_admin_or_403(request)
    if _r is not None: return _r
    conn = get_db()
    aluno = conn.execute("SELECT turma_id FROM alunos WHERE id = ?", (aluno_id,)).fetchone()
    turma_destino = conn.execute("SELECT id FROM turmas WHERE id = ?", (nova_turma_id,)).fetchone()
    if not aluno or not turma_destino:
        conn.close()
        return RedirectResponse("/turmas", status_code=303)
    conn.execute("UPDATE alunos SET turma_id = ? WHERE id = ?", (nova_turma_id, aluno_id))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/turmas/{nova_turma_id}", status_code=303)


# ==========================================
#  FASE R3: IMPORTAÇÃO DE HABILIDADES BNCC
# ==========================================

from fastapi.responses import JSONResponse


# Mapa de nome de disciplina (em pt-BR, normalizado) → 2 letras do código BNCC.
# A posição 5-6 do código identifica o componente curricular:
# EF06MA01 → MA (Matemática), EF06LP01 → LP (Língua Portuguesa), etc.
BNCC_COMPONENTE_POR_DISCIPLINA = {
    "matematica": "MA",
    "lingua portuguesa": "LP", "portugues": "LP",
    "ciencias": "CI",
    "historia": "HI",
    "geografia": "GE",
    "arte": "AR", "artes": "AR",
    "educacao fisica": "EF", "ed fisica": "EF",
    "ingles": "LI", "lingua inglesa": "LI",
    "ensino religioso": "ER", "religiao": "ER",
    "computacao": "CO",
}


def _bncc_componente_de_disciplina(nome):
    """Recebe nome de disciplina (livre) e retorna o código BNCC do componente, ou None."""
    if not nome:
        return None
    # Normaliza: remove acentos e lowercase
    import unicodedata
    norm = unicodedata.normalize("NFD", nome).encode("ASCII", "ignore").decode("ASCII").strip().lower()
    return BNCC_COMPONENTE_POR_DISCIPLINA.get(norm)


@app.get("/habilidades/buscar")
def buscar_habilidades_json(codigos: str = "", q: str = "", disciplina_id: Optional[int] = None):
    """Endpoint JSON usado pelo JS na criação/edição de questão.
    Dois modos:
    - ?codigos=EF06MA01,EF06MA02 → retorna {codigo: descricao} pra validar códigos digitados
    - ?q=palavra&disciplina_id=N → retorna {"results": [...]} com até 30 habilidades cuja descrição
      contém a palavra. Se disciplina_id for fornecido, filtra pelo componente BNCC mapeado.
    """
    conn = get_db()

    # Modo 1: lookup direto por códigos
    if codigos.strip():
        codigos_list = [c.strip().upper() for c in codigos.split(",") if c.strip()]
        if not codigos_list:
            conn.close()
            return JSONResponse({})
        placeholders = ",".join("?" * len(codigos_list))
        rows = conn.execute(
            f"SELECT codigo, descricao FROM habilidades_bncc WHERE codigo IN ({placeholders})",
            codigos_list
        ).fetchall()
        conn.close()
        return JSONResponse({r["codigo"]: (r["descricao"] or "") for r in rows})

    # Modo 2: busca por palavra (opcionalmente filtrada por disciplina)
    if q.strip():
        sql = "SELECT codigo, descricao FROM habilidades_bncc WHERE descricao LIKE ? AND descricao IS NOT NULL"
        params = [f"%{q.strip()}%"]

        if disciplina_id:
            disc = conn.execute("SELECT nome FROM disciplinas WHERE id = ?", (disciplina_id,)).fetchone()
            if disc:
                comp = _bncc_componente_de_disciplina(disc["nome"])
                if comp:
                    # Filtra códigos com componente igual (posição 5-6 do código)
                    sql += " AND substr(codigo, 5, 2) = ?"
                    params.append(comp)

        sql += " ORDER BY codigo LIMIT 30"
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return JSONResponse({
            "results": [{"codigo": r["codigo"], "descricao": r["descricao"]} for r in rows]
        })

    conn.close()
    return JSONResponse({})


@app.get("/habilidades/importar", response_class=HTMLResponse)
def form_importar_habilidades():
    content = """
        <div class="page-header">
            <h1>Importar habilidades BNCC</h1>
            <p class="subtitle">Sobe a planilha oficial do MEC (ou um Excel/CSV próprio) e o sistema cadastra/atualiza tudo.</p>
            <div class="page-actions"><a href="/habilidades" class="btn">← Voltar</a></div>
        </div>

        <div class="tip" style="background:var(--accent-bg); color:var(--accent); border-color:var(--accent);">
            <strong>Formatos aceitos:</strong>
            <ul style="margin:8px 0 0 18px; line-height:1.6; color:var(--accent);">
                <li><strong>Planilha oficial do MEC</strong> (downloadbncc.mec.gov.br) — basta ter uma coluna chamada <code>Habilidade</code> no formato <code>(CODIGO) descrição</code>. As outras colunas (Disciplina, Ano, etc.) são ignoradas.</li>
                <li><strong>Excel/CSV personalizado</strong> — deve ter colunas <code>codigo</code> e <code>descricao</code> (nessa grafia).</li>
            </ul>
        </div>

        <div class="tip" style="background:var(--orange-bg); color:var(--orange); border-color:var(--orange);">
            <strong>Comportamento da importação:</strong>
            <ul style="margin:8px 0 0 18px; line-height:1.6; color:var(--orange);">
                <li>Códigos novos → <strong>cadastrados</strong></li>
                <li>Códigos já existentes <strong>sem descrição</strong> → descrição é <strong>preenchida</strong></li>
                <li>Códigos já existentes <strong>com descrição</strong> → mantida (não sobrescreve)</li>
                <li>Vínculos com questões já cadastradas → preservados</li>
            </ul>
        </div>

        <form action="/habilidades/importar" method="post" enctype="multipart/form-data" style="margin-top:20px;">
            <label>Arquivo<input type="file" name="arquivo" accept=".xlsx,.xls,.csv" required></label>
            <div class="page-actions">
                <button type="submit" class="btn btn-primary">Processar arquivo</button>
                <a href="/habilidades" class="btn">Cancelar</a>
            </div>
        </form>
    """
    return render_page("Importar BNCC", content, active="habilidades")


def _extrair_habilidades_de_xlsx(file_bytes):
    """Detecta automaticamente o formato da planilha e extrai pares (codigo, descricao).
    Suporta:
    1. Planilha do MEC: coluna 'Habilidade' com formato '(CODIGO) descrição'
    2. Planilha custom: colunas 'codigo' + 'descricao'
    """
    import io
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    sheet = wb[wb.sheetnames[0]]

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], "Planilha vazia."

    # Procurar a linha de cabeçalho (pode estar em row 0 ou row 1+)
    header_row_idx = None
    header_cols = {}
    for idx, row in enumerate(rows[:5]):  # busca nas 5 primeiras linhas
        cells_lower = [(str(c).strip().lower() if c is not None else "") for c in row]
        # Formato MEC: coluna "habilidade"
        if "habilidade" in cells_lower:
            header_row_idx = idx
            header_cols["habilidade"] = cells_lower.index("habilidade")
            break
        # Formato custom: "codigo" + "descricao"
        elif "codigo" in cells_lower or "código" in cells_lower:
            header_row_idx = idx
            for nome in ("codigo", "código"):
                if nome in cells_lower:
                    header_cols["codigo"] = cells_lower.index(nome)
                    break
            for nome in ("descricao", "descrição", "description"):
                if nome in cells_lower:
                    header_cols["descricao"] = cells_lower.index(nome)
                    break
            break

    if header_row_idx is None:
        return [], 'Não encontrei a coluna "Habilidade" (formato MEC) nem "codigo" (formato customizado) nas primeiras linhas. Confira o cabeçalho.'

    pad = re.compile(r'^\(([A-Z]{2}\d{2,3}[A-Z]{2,3}\d{2,3})\)\s*(.+)$', re.DOTALL)
    encontrados = []
    for row in rows[header_row_idx + 1:]:
        if "habilidade" in header_cols:
            cell = row[header_cols["habilidade"]] if header_cols["habilidade"] < len(row) else None
            if cell is None:
                continue
            txt = str(cell).strip()
            if not txt:
                continue
            m = pad.match(txt)
            if not m:
                continue
            codigo = m.group(1).strip().upper()
            desc = m.group(2).strip().replace("\n", " ")
            while "  " in desc:
                desc = desc.replace("  ", " ")
            encontrados.append((codigo, desc))
        else:
            codigo_idx = header_cols.get("codigo")
            desc_idx = header_cols.get("descricao")
            if codigo_idx is None:
                continue
            codigo_val = row[codigo_idx] if codigo_idx < len(row) else None
            if not codigo_val:
                continue
            codigo = str(codigo_val).strip().upper()
            if not re.match(r'^[A-Z]{2}\d{2,3}[A-Z]{2,3}\d{2,3}$', codigo):
                continue  # código inválido, pula
            desc = ""
            if desc_idx is not None and desc_idx < len(row) and row[desc_idx] is not None:
                desc = str(row[desc_idx]).strip()
            encontrados.append((codigo, desc))

    return encontrados, None


@app.post("/habilidades/importar", response_class=HTMLResponse)
async def processar_importacao_habilidades(arquivo: UploadFile = File(...)):
    if not arquivo or not arquivo.filename:
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Nenhum arquivo enviado.</p><a href="/habilidades/importar" class="btn">← Voltar</a></div>', active="habilidades"))

    file_bytes = await arquivo.read()
    if not file_bytes:
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Arquivo vazio.</p><a href="/habilidades/importar" class="btn">← Voltar</a></div>', active="habilidades"))

    nome = arquivo.filename.lower()
    if nome.endswith(".csv"):
        # Suporte simples a CSV: converter pra lista (codigo, descricao)
        import csv, io
        try:
            text = file_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = file_bytes.decode("latin-1")
        reader = csv.reader(io.StringIO(text))
        rows = list(reader)
        if not rows:
            return HTMLResponse(render_page("Erro", '<div class="empty"><p>CSV vazio.</p><a href="/habilidades/importar" class="btn">← Voltar</a></div>', active="habilidades"))
        header = [c.strip().lower() for c in rows[0]]
        if "codigo" not in header and "código" not in header:
            return HTMLResponse(render_page("Erro", '<div class="empty"><p>CSV precisa ter coluna "codigo".</p><a href="/habilidades/importar" class="btn">← Voltar</a></div>', active="habilidades"))
        codigo_idx = header.index("codigo") if "codigo" in header else header.index("código")
        desc_idx = -1
        for nome_col in ("descricao", "descrição", "description"):
            if nome_col in header:
                desc_idx = header.index(nome_col)
                break
        encontrados = []
        for row in rows[1:]:
            if codigo_idx >= len(row): continue
            codigo = (row[codigo_idx] or "").strip().upper()
            if not re.match(r'^[A-Z]{2}\d{2,3}[A-Z]{2,3}\d{2,3}$', codigo):
                continue
            desc = (row[desc_idx].strip() if desc_idx >= 0 and desc_idx < len(row) else "")
            encontrados.append((codigo, desc))
        erro = None
    else:
        encontrados, erro = _extrair_habilidades_de_xlsx(file_bytes)

    if erro:
        return HTMLResponse(render_page("Erro", f'<div class="empty"><p>{erro}</p><a href="/habilidades/importar" class="btn">← Tentar outro arquivo</a></div>', active="habilidades"))

    if not encontrados:
        return HTMLResponse(render_page("Sem dados", '<div class="empty"><p>O arquivo foi lido mas nenhum código BNCC válido foi encontrado. Verifique se as células no formato <code>(CODIGO) descrição</code> estão preenchidas.</p><a href="/habilidades/importar" class="btn">← Voltar</a></div>', active="habilidades"))

    # UPSERT no banco
    conn = get_db()
    novas, atualizadas, mantidas = 0, 0, 0
    for codigo, desc in encontrados:
        existing = conn.execute("SELECT id, descricao FROM habilidades_bncc WHERE codigo = ?", (codigo,)).fetchone()
        if existing:
            cur_desc = (existing["descricao"] or "").strip()
            if not cur_desc and desc:
                conn.execute("UPDATE habilidades_bncc SET descricao = ? WHERE id = ?", (desc, existing["id"]))
                atualizadas += 1
            else:
                mantidas += 1
        else:
            conn.execute("INSERT INTO habilidades_bncc (codigo, descricao) VALUES (?, ?)", (codigo, desc or None))
            novas += 1
    conn.commit()
    total_apos = conn.execute("SELECT COUNT(*) AS c FROM habilidades_bncc").fetchone()["c"]
    conn.close()

    # Amostras
    amostras = ""
    for codigo, desc in encontrados[:5]:
        amostras += f'<li><strong>{codigo}</strong>: {desc[:140]}{"..." if len(desc)>140 else ""}</li>'

    content = f"""
        <div class="page-header">
            <h1>✅ Importação concluída</h1>
            <p class="subtitle">Arquivo <code>{arquivo.filename}</code> processado com sucesso.</p>
        </div>

        <div class="metric-grid">
            <div class="metric"><div class="metric-label">Lidas do arquivo</div><div class="metric-value">{len(encontrados)}</div></div>
            <div class="metric"><div class="metric-label">Novas cadastradas</div><div class="metric-value" style="color:var(--green);">{novas}</div></div>
            <div class="metric"><div class="metric-label">Atualizadas (descrição)</div><div class="metric-value" style="color:var(--orange);">{atualizadas}</div></div>
            <div class="metric"><div class="metric-label">Já existiam (mantidas)</div><div class="metric-value" style="color:var(--text-muted);">{mantidas}</div></div>
        </div>

        <div class="tip" style="margin-top:18px;">
            Total no banco agora: <strong>{total_apos}</strong> habilidades.
        </div>

        <h3 style="margin-top:24px;">Primeiras 5 lidas (amostra):</h3>
        <ul style="line-height:1.6;">{amostras}</ul>

        <div class="page-actions" style="margin-top:18px;">
            <a href="/habilidades" class="btn btn-primary">Ver catálogo</a>
            <a href="/habilidades/importar" class="btn">Importar outro arquivo</a>
        </div>
    """
    return render_page("Importação concluída", content, active="habilidades")


# ==========================================
#  FASE C3: OMR EM LOTE
# ==========================================

def _extrair_imagens_de_arquivo(file_bytes: bytes, filename: str) -> list:
    """Recebe bytes de um arquivo (imagem ou PDF) e retorna lista de (nome_exibicao, bytes_jpeg)."""
    fname_lower = (filename or "").lower()
    if fname_lower.endswith(".pdf"):
        try:
            from pdf2image import convert_from_bytes
            import io as _io
            paginas = convert_from_bytes(file_bytes, dpi=200, fmt="jpeg")
            resultado = []
            for i, pil_img in enumerate(paginas, start=1):
                buf = _io.BytesIO()
                pil_img.save(buf, format="JPEG", quality=90)
                resultado.append((f"{filename} — pág. {i}", buf.getvalue()))
            return resultado
        except ImportError:
            return [(filename, None)]
        except Exception:
            return [(filename, None)]
    else:
        return [(filename, file_bytes)]


@app.post("/aplicacoes/{aplicacao_id}/escanear-lote", response_class=HTMLResponse)
async def processar_escaneamento_lote(aplicacao_id: int, fotos: List[UploadFile] = File(...)):
    """Recebe N fotos ou PDF multipágina, processa cada uma, mostra tela de revisão com grid de cards."""
    if not fotos:
        return HTMLResponse(render_page("Erro", '<div class="empty"><p>Nenhum arquivo enviado.</p></div>', active="aplicacoes"))

    conn = get_db()
    apl = conn.execute("""
        SELECT a.*, p.titulo AS prova_titulo, t.nome AS turma_nome, t.ano_letivo
        FROM aplicacoes a JOIN provas p ON p.id = a.prova_id JOIN turmas t ON t.id = a.turma_id
        WHERE a.id = ?
    """, (aplicacao_id,)).fetchone()
    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    questoes = conn.execute(
        "SELECT q.id FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id WHERE pq.prova_id = ? ORDER BY pq.ordem",
        (apl["prova_id"],)
    ).fetchall()
    n_questoes = len(questoes)
    questoes_info = _coletar_info_questoes_cartao(conn, apl["prova_id"])

    arquivos_expandidos = []
    for foto in fotos:
        raw = await foto.read()
        if not raw:
            arquivos_expandidos.append((foto.filename or "sem nome", None))
            continue
        arquivos_expandidos.extend(_extrair_imagens_de_arquivo(raw, foto.filename or ""))

    cards_html_parts = []
    n_ok = 0
    n_warn = 0
    n_erro = 0
    alunos_ja_no_lote = set()

    for idx, (nome_exib, image_bytes) in enumerate(arquivos_expandidos):
        if not image_bytes:
            n_erro += 1
            if (nome_exib or "").lower().endswith(".pdf"):
                cards_html_parts.append(_render_card_erro(idx, nome_exib,
                    "PDF recebido mas 'pdf2image' não está instalado. Execute: pip install pdf2image --break-system-packages"))
            else:
                cards_html_parts.append(_render_card_erro(idx, nome_exib, "Arquivo vazio."))
            continue

        result = _processar_cartao_resposta(image_bytes, n_questoes, filename=nome_exib or "", questoes_info=questoes_info)

        if not result["success"]:
            n_erro += 1
            cards_html_parts.append(_render_card_erro(idx, nome_exib, result.get("error", "Erro desconhecido")))
            continue

        aluno = conn.execute("SELECT * FROM alunos WHERE id = ? AND turma_id = ?",
                             (result["aluno_id"], apl["turma_id"])).fetchone()
        if not aluno:
            n_erro += 1
            cards_html_parts.append(_render_card_erro(
                idx, nome_exib,
                f"QR aponta para aluno {result['aluno_id']} que NÃO pertence à turma {apl['turma_nome']}.",
                preview_b64=result.get("preview_base64")
            ))
            continue

        if result["aplicacao_id_qr"] != aplicacao_id:
            n_erro += 1
            cards_html_parts.append(_render_card_erro(
                idx, nome_exib,
                f"Cartão de OUTRA aplicação (id {result['aplicacao_id_qr']}).",
                preview_b64=result.get("preview_base64")
            ))
            continue

        ja_entregue = conn.execute(
            "SELECT finalizada_em FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?",
            (aplicacao_id, result["aluno_id"])
        ).fetchone()
        duplicata_lote = result["aluno_id"] in alunos_ja_no_lote
        alunos_ja_no_lote.add(result["aluno_id"])

        if result["warnings"] or duplicata_lote:
            n_warn += 1
        else:
            n_ok += 1

        cards_html_parts.append(_render_card_revisao_lote(
            idx, nome_exib, aluno, result, n_questoes,
            ja_entregue=ja_entregue, duplicata_lote=duplicata_lote,
            questoes_info=questoes_info
        ))

    conn.close()

    if not cards_html_parts:
        return HTMLResponse(render_page("Lote vazio",
            '<div class="empty">Nenhuma foto foi processada.</div>', active="aplicacoes"))

    resumo = f"""
        <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:10px; margin-bottom:18px;">
            <div class="metric"><div class="metric-label">Cartões processados</div><div class="metric-value">{len(arquivos_expandidos)}</div></div>
            <div class="metric"><div class="metric-label">Lidas OK</div><div class="metric-value" style="color:var(--green);">{n_ok}</div></div>
            <div class="metric"><div class="metric-label">Com avisos</div><div class="metric-value" style="color:var(--orange);">{n_warn}</div></div>
            <div class="metric"><div class="metric-label">Com erro</div><div class="metric-value" style="color:var(--red);">{n_erro}</div></div>
        </div>
    """

    legenda = """
        <div class="tip" style="font-size:12px;">
            <strong>Como usar:</strong>
            Verifique os cartões marcados em amarelo/vermelho (warnings/erros). Cada card permite editar manualmente as respostas marcadas. Os cards de erro NÃO serão salvos (sem checkbox de confirmar). Ao final, clique em <strong>"Salvar todos confirmados"</strong> e o sistema grava tudo de uma vez.
        </div>
    """

    cards_html = "".join(cards_html_parts)

    content = f"""
        <div class="page-header">
            <h1>📋 Revisão do lote</h1>
            <p class="subtitle">{apl["prova_titulo"]} · {apl["turma_nome"]} ({apl["ano_letivo"]})</p>
            <div class="page-actions"><a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">← Voltar para escanear</a></div>
        </div>

        {resumo}
        {legenda}

        <form action="/aplicacoes/{aplicacao_id}/escanear-lote/confirmar" method="post">
            <input type="hidden" name="n_questoes" value="{n_questoes}">
            {cards_html}
            <div style="position:sticky; bottom:0; background:var(--bg); padding:14px; border-top:2px solid var(--border); margin-top:18px; display:flex; gap:10px; align-items:center;">
                <button type="submit" class="btn btn-primary" style="font-size:15px;">✓ Salvar todos confirmados</button>
                <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">📷 Escanear mais</a>
                <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar</a>
                <span style="margin-left:auto; font-size:12px; color:var(--text-muted);">Cards desmarcados NÃO serão salvos. Cards com erro são puramente informativos.</span>
            </div>
        </form>

        <script>
        // Expansão dos cards de revisão
        document.addEventListener('click', e => {{
            const btn = e.target.closest('[data-toggle-card]');
            if (!btn) return;
            const card = btn.closest('.lote-card');
            const body = card.querySelector('.lote-card-body');
            const open = body.style.display !== 'none';
            body.style.display = open ? 'none' : 'block';
            btn.textContent = open ? '▼ Expandir' : '▲ Recolher';
        }});
        </script>
    """
    return render_page("Revisão do lote", content, active="aplicacoes")


def _render_card_erro(idx, filename, mensagem_erro, preview_b64=None):
    """Card visual de uma foto que falhou no processamento. NÃO inclui no form (sem fields)."""
    nome_seguro = (filename or f"foto_{idx+1}").replace("<", "&lt;")
    preview_img = ""
    if preview_b64:
        preview_img = f'<img src="data:image/jpeg;base64,{preview_b64}" style="max-width:200px; max-height:150px; border:1px solid var(--border); margin-top:8px; border-radius:4px;">'
    return f"""
    <div class="lote-card" style="border:1px solid var(--red); border-radius:8px; padding:14px; margin-bottom:10px; background:var(--red-bg); color:var(--red);">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <strong style="color:var(--red);">✗ Foto {idx+1}: {nome_seguro}</strong>
                <div style="font-size:13px; color:var(--red); margin-top:4px;">{mensagem_erro}</div>
                {preview_img}
            </div>
            <span style="font-size:12px; color:var(--red); flex-shrink:0;">Não será salvo</span>
        </div>
    </div>
    """


def _render_card_revisao_lote(idx, filename, aluno, result, n_questoes, ja_entregue=None, duplicata_lote=False, questoes_info=None):
    """Card visual de uma foto lida com sucesso. Inclui form fields editáveis.
    questoes_info: lista [{id, num, tipo, vf_count, assoc_a_count, assoc_b_count}] pra renderizar conforme tipo."""
    nome_seguro = (filename or f"foto_{idx+1}").replace("<", "&lt;")
    aluno_id = result["aluno_id"]

    # Status: verde se sem warnings, amarelo se tem warning, ou amarelo se duplicata
    tem_avisos = bool(result["warnings"]) or duplicata_lote or ja_entregue
    if tem_avisos:
        border_color = "var(--orange)"
        bg = "var(--orange-bg)"
        status_icon = "⚠"
        status_color = "var(--orange)"
        body_default_display = "block"  # auto-expandido se tem aviso
        toggle_label = "▲ Recolher"
    else:
        border_color = "var(--green)"
        bg = "var(--green-bg)"
        status_icon = "✓"
        status_color = "var(--green)"
        body_default_display = "none"  # colapsado por padrão
        toggle_label = "▼ Expandir"

    # Avisos
    avisos = []
    if duplicata_lote:
        avisos.append("⚠ Foto repetida no lote: já apareceu um cartão deste aluno antes (a última marcação prevalece).")
    if ja_entregue:
        avisos.append(f"⚠ Aluno já tem entrega registrada ({ja_entregue['finalizada_em']}). Confirmar irá sobrescrever as respostas anteriores.")
    avisos.extend(result.get("warnings", []) or [])
    avisos_html = ""
    if avisos:
        items = "".join(f"<li>{w}</li>" for w in avisos)
        avisos_html = f'<ul style="margin:8px 0 0 18px; font-size:12px; color:var(--orange);">{items}</ul>'

    # Grid de respostas editáveis — adapta conforme tipo
    answers = result["answers"]
    info_by_num = {i["num"]: i for i in (questoes_info or [])}
    tabela_html = ""
    for q_num in range(1, n_questoes + 1):
        info = info_by_num.get(q_num, {"tipo": "multipla_escolha"})
        tipo_q = info.get("tipo", "multipla_escolha")
        detected = answers.get(q_num)

        if tipo_q == "multipla_escolha":
            cells = ""
            for letra in ["A", "B", "C", "D"]:
                checked = " checked" if detected == letra else ""
                cells += f'<td style="text-align:center; padding:2px;"><label style="cursor:pointer;"><input type="radio" name="card_{idx}_q_{q_num}" value="{letra}"{checked} style="width:auto; margin:0;"> {letra}</label></td>'
            em_branco_checked = " checked" if detected is None else ""
            cells += f'<td style="text-align:center; padding:2px; background:var(--bg-subtle);"><label style="cursor:pointer;"><input type="radio" name="card_{idx}_q_{q_num}" value=""{em_branco_checked} style="width:auto; margin:0;"> ∅</label></td>'
            tabela_html += f'<tr><td style="padding:3px 6px; font-weight:600;">Q{q_num}</td>{cells}</tr>'
        elif tipo_q == "vf":
            n_afirms = info.get("vf_count", 0)
            detected_dict = detected if isinstance(detected, dict) else {}
            sub_rows = ""
            for k in range(n_afirms):
                val = detected_dict.get(str(k))
                ck_v = " checked" if val == "V" else ""
                ck_f = " checked" if val == "F" else ""
                ck_n = " checked" if not val else ""
                sub_rows += (
                    f'<tr><td style="padding:3px 6px; font-weight:600; color:var(--text-muted);">Q{q_num}.{k+1}</td>'
                    f'<td colspan="5" style="padding:2px;">'
                    f'<label style="margin-right:14px;"><input type="radio" name="card_{idx}_q_{q_num}_vf_{k}" value="V"{ck_v} style="width:auto; margin:0 3px 0 0;">V</label>'
                    f'<label style="margin-right:14px;"><input type="radio" name="card_{idx}_q_{q_num}_vf_{k}" value="F"{ck_f} style="width:auto; margin:0 3px 0 0;">F</label>'
                    f'<label><input type="radio" name="card_{idx}_q_{q_num}_vf_{k}" value=""{ck_n} style="width:auto; margin:0 3px 0 0;">∅</label>'
                    f'</td></tr>'
                )
            tabela_html += sub_rows
        elif tipo_q == "associacao":
            n_a = info.get("assoc_a_count", 0)
            n_b = info.get("assoc_b_count", 0)
            detected_dict = detected if isinstance(detected, dict) else {}
            letras = [chr(97+i) for i in range(n_b)]
            sub_rows = ""
            for k in range(n_a):
                val = detected_dict.get(str(k))
                opts = ""
                for letra in letras:
                    ck = " checked" if val == letra else ""
                    opts += f'<label style="margin-right:10px;"><input type="radio" name="card_{idx}_q_{q_num}_assoc_{k}" value="{letra}"{ck} style="width:auto; margin:0 3px 0 0;">{letra}</label>'
                ck_n = " checked" if not val else ""
                opts += f'<label><input type="radio" name="card_{idx}_q_{q_num}_assoc_{k}" value=""{ck_n} style="width:auto; margin:0 3px 0 0;">∅</label>'
                sub_rows += (
                    f'<tr><td style="padding:3px 6px; font-weight:600; color:var(--text-muted);">Q{q_num}.{k+1}</td>'
                    f'<td colspan="5" style="padding:2px;">{opts}</td></tr>'
                )
            tabela_html += sub_rows
        elif tipo_q == "discursiva":
            tabela_html += (
                f'<tr><td style="padding:3px 6px; font-weight:600; color:var(--text-muted);">Q{q_num}</td>'
                f'<td colspan="5" style="padding:3px 6px; font-size:11px; color:var(--text-muted); font-style:italic;">📝 Discursiva — correção manual</td></tr>'
            )

    return f"""
    <div class="lote-card" style="border:2px solid {border_color}; border-radius:8px; padding:14px; margin-bottom:10px; background:{bg};">
        <input type="hidden" name="card_{idx}_aluno_id" value="{aluno_id}">

        <div style="display:flex; justify-content:space-between; align-items:center; gap:12px;">
            <div style="flex:1; min-width:0;">
                <label style="display:flex; align-items:center; gap:8px; cursor:pointer; margin:0; font-size:15px;">
                    <input type="checkbox" name="card_{idx}_confirmar" value="1" checked style="width:auto; margin:0;">
                    <span style="color:{status_color}; font-size:18px;">{status_icon}</span>
                    <strong>{aluno["nome"]}</strong>
                    <span style="font-size:12px; color:var(--text-muted);">· Nº {aluno["numero"] or "—"} · {aluno["codigo_unico"]} · foto: {nome_seguro}</span>
                </label>
            </div>
            <button type="button" data-toggle-card class="btn" style="padding:4px 10px; font-size:12px; flex-shrink:0;">{toggle_label}</button>
        </div>

        {avisos_html}

        <div class="lote-card-body" style="display:{body_default_display}; margin-top:12px;">
            <div style="display:grid; grid-template-columns: 1fr 1.2fr; gap:14px;">
                <div>
                    <p class="muted-line" style="font-size:11px; margin:0 0 4px 0;">Imagem processada (verde = detectado como marcado)</p>
                    <img src="data:image/jpeg;base64,{result['preview_base64']}" style="width:100%; border:1px solid var(--border); border-radius:4px;">
                </div>
                <div>
                    <p class="muted-line" style="font-size:11px; margin:0 0 4px 0;">Respostas (corrija se necessário)</p>
                    <table style="width:100%; border-collapse:collapse; font-size:12px;">
                        <thead><tr style="background:var(--bg-subtle);"><th style="padding:3px;">Q</th><th style="padding:3px;">A</th><th style="padding:3px;">B</th><th style="padding:3px;">C</th><th style="padding:3px;">D</th><th style="padding:3px;">∅</th></tr></thead>
                        <tbody>{tabela_html}</tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    """


@app.post("/aplicacoes/{aplicacao_id}/escanear-lote/confirmar", response_class=HTMLResponse)
async def confirmar_lote(aplicacao_id: int, request: Request):
    """Salva no banco todos os cards confirmados (checkbox marcado) do lote."""
    form = await request.form()
    n_questoes = int(form.get("n_questoes", 0))

    conn = get_db()
    apl = conn.execute("SELECT * FROM aplicacoes WHERE id = ?", (aplicacao_id,)).fetchone()
    if not apl:
        conn.close()
        return RedirectResponse("/aplicacoes", status_code=303)

    questoes = conn.execute(
        "SELECT q.id FROM prova_questoes pq JOIN questoes q ON q.id = pq.questao_id WHERE pq.prova_id = ? ORDER BY pq.ordem",
        (apl["prova_id"],)
    ).fetchall()
    questao_ids = [q["id"] for q in questoes]

    # Identificar cards confirmados varrendo os campos card_X_confirmar=1
    confirmados_idx = set()
    for key in form.keys():
        if key.startswith("card_") and key.endswith("_confirmar"):
            try:
                idx = int(key.split("_")[1])
                confirmados_idx.add(idx)
            except (ValueError, IndexError):
                continue

    salvos = []
    for idx in sorted(confirmados_idx):
        aluno_id_str = form.get(f"card_{idx}_aluno_id")
        if not aluno_id_str:
            continue
        try:
            aluno_id = int(aluno_id_str)
        except ValueError:
            continue

        aluno = conn.execute("SELECT * FROM alunos WHERE id = ? AND turma_id = ?",
                             (aluno_id, apl["turma_id"])).fetchone()
        if not aluno:
            continue

        # Override completo de respostas anteriores deste aluno
        conn.execute("DELETE FROM respostas WHERE aplicacao_id = ? AND aluno_id = ?", (aplicacao_id, aluno_id))

        # Coleta tipos das questões da prova
        tipos_q = {q["id"]: (q["tipo"] if "tipo" in q.keys() and q["tipo"] else "multipla_escolha")
                   for q in conn.execute("""
                       SELECT q.id, q.tipo FROM prova_questoes pq
                       JOIN questoes q ON q.id = pq.questao_id
                       WHERE pq.prova_id = ? ORDER BY pq.ordem
                   """, (apl["prova_id"],)).fetchall()}

        for q_num, q_id in enumerate(questao_ids, start=1):
            tipo_q = tipos_q.get(q_id, "multipla_escolha")
            if tipo_q == "multipla_escolha":
                letra = form.get(f"card_{idx}_q_{q_num}", "").strip()
                _gravar_resposta_questao(conn, aplicacao_id, aluno_id, q_id, tipo_q, letra or None)
            elif tipo_q == "vf":
                # form tem campos card_X_q_Y_vf_N = "V" ou "F" pra cada afirmação N
                marcadas = {}
                for k in range(VF_MAX_AFIRMACOES):
                    v = form.get(f"card_{idx}_q_{q_num}_vf_{k}", "").strip().upper()
                    if v in ("V", "F"):
                        marcadas[str(k)] = v
                if marcadas:
                    _gravar_resposta_questao(conn, aplicacao_id, aluno_id, q_id, tipo_q, marcadas)
            elif tipo_q == "associacao":
                # form tem campos card_X_q_Y_assoc_N = letra pra cada item N
                marcadas = {}
                for k in range(ASSOC_MAX_PARES):
                    v = form.get(f"card_{idx}_q_{q_num}_assoc_{k}", "").strip().lower()
                    if v:
                        marcadas[str(k)] = v
                if marcadas:
                    _gravar_resposta_questao(conn, aplicacao_id, aluno_id, q_id, tipo_q, marcadas)
            # discursiva: nada (correção manual)

        existing = conn.execute("SELECT id FROM entregas WHERE aplicacao_id = ? AND aluno_id = ?",
                                (aplicacao_id, aluno_id)).fetchone()
        if not existing:
            conn.execute("INSERT INTO entregas (aplicacao_id, aluno_id) VALUES (?, ?)", (aplicacao_id, aluno_id))
        else:
            conn.execute("UPDATE entregas SET finalizada_em = CURRENT_TIMESTAMP WHERE aplicacao_id = ? AND aluno_id = ?",
                         (aplicacao_id, aluno_id))

        score, total = _calcular_nota(conn, aplicacao_id, aluno_id)
        salvos.append({"aluno": aluno["nome"], "numero": aluno["numero"], "score": score, "total": total})

    conn.commit()
    conn.close()

    # Resumo
    if not salvos:
        content = f"""
            <div class="page-header"><h1>⚠️ Nada salvo</h1></div>
            <div class="empty">Nenhum cartão foi confirmado pra salvar. Talvez todos estejam desmarcados ou em erro.</div>
            <div class="page-actions">
                <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn btn-primary">📷 Tentar de novo</a>
                <a href="/aplicacoes/{aplicacao_id}" class="btn">← Voltar</a>
            </div>
        """
        return render_page("Nada salvo", content, active="aplicacoes")

    linhas = "".join(
        f'<tr><td>{s["aluno"]}</td><td>{s["numero"] or "—"}</td><td><strong>{s["score"]}/{s["total"]}</strong> ({(s["score"]/s["total"]*100 if s["total"]>0 else 0):.0f}%)</td></tr>'
        for s in salvos
    )

    content = f"""
        <div class="page-header"><h1>✅ {len(salvos)} cartão(ões) salvos</h1></div>
        <div style="border:1px solid var(--green); background:var(--green-bg); padding:16px; border-radius:6px; margin:16px 0; color:var(--green);">
            <p style="margin:0 0 10px 0;"><strong>Resultados:</strong></p>
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg);"><th style="padding:6px; text-align:left;">Aluno</th><th style="padding:6px;">Nº</th><th style="padding:6px;">Nota</th></tr></thead>
                <tbody>{linhas}</tbody>
            </table>
        </div>
        <div class="page-actions">
            <a href="/aplicacoes/{aplicacao_id}" class="btn btn-primary">Ver aplicação</a>
            <a href="/aplicacoes/{aplicacao_id}/escanear" class="btn">📷 Escanear mais</a>
        </div>
    """
    return render_page("Lote salvo", content, active="aplicacoes")