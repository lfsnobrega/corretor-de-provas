"""
Módulo de Horários Escolares — Sistema Pedagógico do Walmir
===========================================================

Fase 1, primeira entrega (01/11/2026): COLETA DE DISPONIBILIDADE dos professores,
substituindo o Google Forms "Disponibilidade de horário".

Este arquivo fica SEPARADO do main.py de propósito (decisão de Felipe, 01/11/2026):
assim o código do horário não aumenta o main.py, que já passa de 28 mil linhas.
O main.py só precisa de uma linha no final para ligar este módulo:

    horarios.registrar(app, get_db=get_db, render_page=render_page,
                       get_current_professor=get_current_professor)

O que tem aqui:
  - Tabelas próprias, todas com prefixo horario_ (criadas sozinhas ao iniciar).
  - /horarios/disponibilidade ............ formulário do professor (com as diretrizes
                                            da escola conferidas na hora).
  - /horarios/disponibilidade/painel ..... painel da gestão: quem respondeu, quem está
                                            pendente, disponibilidade por turno, editar,
                                            excluir, dispensar, período da coleta.
  - /horarios/disponibilidade/exportar ... planilha Excel no mesmo formato do Forms.

Lições do sistema aplicadas aqui (ver histórico do main.py):
  - Gravação sempre com INSERT ... ON CONFLICT DO UPDATE (nunca "buscar-ou-inserir").
  - Nenhuma função de "mesclar" rodando a cada reinício.
  - Exclusões feitas explicitamente no código (o SQLite desta app não faz CASCADE).
  - Rotas específicas declaradas antes das genéricas.
"""

from __future__ import annotations

import html
from datetime import datetime, date, timedelta, timezone
from io import BytesIO
from typing import Optional

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse

router = APIRouter()

# Funções emprestadas do main.py (preenchidas em registrar()). Assim este arquivo não
# precisa importar o main.py — evita importação circular.
_ctx: dict = {}


def registrar(app, get_db, render_page, get_current_professor):
    """Chamada uma única vez, no final do main.py."""
    _ctx["get_db"] = get_db
    _ctx["render_page"] = render_page
    _ctx["get_prof"] = get_current_professor
    _criar_tabelas()
    app.include_router(router)


# ==========================================
#  CONSTANTES
# ==========================================

# Brasil não tem mais horário de verão (desde 2019): UTC-3 fixo é seguro e não depende
# do fuso configurado na VM (que costuma ser UTC).
FUSO_BR = timezone(timedelta(hours=-3))

DIAS = [("seg", "Segunda"), ("ter", "Terça"), ("qua", "Quarta"), ("qui", "Quinta"), ("sex", "Sexta")]
TURNOS = [("m", "Manhã"), ("t", "Tarde")]
# Uma coluna por turno no banco: seg_m, ter_m, ..., sex_t (10 colunas). Fica fácil para
# o gerador de horário consultar depois ("quem está disponível quinta à tarde?").
CAMPOS_TURNO = [f"{d}_{t}" for t, _ in TURNOS for d, _ in DIAS]

SITUACOES = ["Matrícula", "REDA", "Dupla Regência"]
# Mesma lista do Google Forms de 2026 + Educação Digital (os professores de ED também
# precisam informar disponibilidade, e o Forms antigo não tinha essa opção).
DISCIPLINAS = ["Português", "Matemática", "História", "Ciências", "Geografia",
               "Arte", "Educação Física", "Inglês", "Educação Digital"]

TURNOS_MINIMOS = 5        # diretriz 4: só a matrícula = 5 turnos
TURNOS_MAX_DUPLA = 8      # diretriz 5: com dupla regência, até 8 turnos
TEMPOS_DUPLA_MAX = 12     # opções 0 a 12, como no Forms

# Coleta criada automaticamente na primeira vez que o módulo roda (só se ainda não
# existir nenhuma coleta). Datas combinadas com Felipe em 01/11/2026.
COLETA_INICIAL = {"ano_letivo": 2027, "data_inicio": "2026-11-03", "data_fim": "2026-11-15"}


# ==========================================
#  BANCO DE DADOS
# ==========================================

def _db():
    return _ctx["get_db"]()


def _criar_tabelas():
    colunas_turno = ",\n            ".join(f"{c} INTEGER NOT NULL DEFAULT 0" for c in CAMPOS_TURNO)
    conn = _db()
    conn.executescript(f"""
        CREATE TABLE IF NOT EXISTS horario_coletas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL UNIQUE,
            data_inicio TEXT NOT NULL,
            data_fim TEXT NOT NULL,
            encerrada INTEGER NOT NULL DEFAULT 0,
            criada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS horario_disponibilidade (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            coleta_id INTEGER NOT NULL,
            professor_id INTEGER NOT NULL,
            situacao TEXT NOT NULL,
            disciplina TEXT NOT NULL,
            dupla_interesse INTEGER NOT NULL DEFAULT 0,
            dupla_tempos INTEGER NOT NULL DEFAULT 0,
            {colunas_turno},
            observacoes TEXT,
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            atualizado_por INTEGER,
            UNIQUE(coleta_id, professor_id)
        );
        CREATE TABLE IF NOT EXISTS horario_coleta_dispensas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            coleta_id INTEGER NOT NULL,
            professor_id INTEGER NOT NULL,
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(coleta_id, professor_id)
        );
    """)
    # Primeira coleta (2027): só cria se a tabela estiver vazia. INSERT OR IGNORE garante
    # que dois processos iniciando juntos não criem duas.
    total = conn.execute("SELECT COUNT(*) AS c FROM horario_coletas").fetchone()["c"]
    if total == 0:
        conn.execute(
            "INSERT OR IGNORE INTO horario_coletas (ano_letivo, data_inicio, data_fim) VALUES (?, ?, ?)",
            (COLETA_INICIAL["ano_letivo"], COLETA_INICIAL["data_inicio"], COLETA_INICIAL["data_fim"]),
        )
    conn.commit()
    conn.close()


# ==========================================
#  AUXILIARES
# ==========================================

def _e(texto) -> str:
    """Escapa texto digitado por usuário antes de colocar no HTML."""
    return html.escape("" if texto is None else str(texto))


def _hoje() -> date:
    return datetime.now(FUSO_BR).date()


def _data_br(iso: Optional[str]) -> str:
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "—"


def _data_br_curta(iso: Optional[str]) -> str:
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d/%m")
    except (ValueError, TypeError):
        return "—"


def _ts_br(ts: Optional[str]) -> str:
    """CURRENT_TIMESTAMP do SQLite vem em UTC; mostra no horário de Brasília."""
    if not ts:
        return "—"
    try:
        dt = datetime.strptime(str(ts)[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
        return dt.astimezone(FUSO_BR).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(ts)


def _eh_gestao(prof: Optional[dict]) -> bool:
    return bool(prof and (prof.get("is_admin") or prof.get("is_gestor") or prof.get("papel") == "gestao"))


def _pode_responder(prof: Optional[dict]) -> bool:
    # Apoio Educacional e Orientação Educacional não dão aula (e o middleware do main.py
    # já os redireciona); todo o resto pode responder.
    if not prof:
        return False
    if _eh_gestao(prof):
        return True
    return prof.get("papel") not in ("apoio", "orientacao")


def _status_coleta(coleta) -> str:
    """'agendada' (ainda não começou), 'aberta' ou 'encerrada'."""
    if coleta is None:
        return "inexistente"
    if coleta["encerrada"]:
        return "encerrada"
    hoje = _hoje()
    try:
        inicio = datetime.strptime(coleta["data_inicio"], "%Y-%m-%d").date()
        fim = datetime.strptime(coleta["data_fim"], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return "encerrada"
    if hoje < inicio:
        return "agendada"
    if hoje > fim:
        return "encerrada"
    return "aberta"


def _coleta_atual(conn):
    """A coleta mais recente (maior ano letivo)."""
    return conn.execute("SELECT * FROM horario_coletas ORDER BY ano_letivo DESC LIMIT 1").fetchone()


def _coleta_por_id(conn, coleta_id) -> Optional[dict]:
    try:
        cid = int(coleta_id)
    except (TypeError, ValueError):
        return None
    return conn.execute("SELECT * FROM horario_coletas WHERE id = ?", (cid,)).fetchone()


def _turnos_marcados(resp) -> list:
    return [c for c in CAMPOS_TURNO if resp and resp[c]]


def _dias_do_turno(resp, turno: str) -> list:
    return [nome for d, nome in DIAS if resp[f"{d}_{turno}"]]


def _avaliar_diretrizes(dados: dict) -> list:
    """Confere as diretrizes da escola (texto do Forms de 2026) e devolve a lista de
    avisos. Não bloqueia sozinho: quem decide se bloqueia é quem chama."""
    marcados = [c for c in CAMPOS_TURNO if dados.get(c)]
    n = len(marcados)
    avisos = []
    if n and not any(c.startswith(("seg_", "sex_")) for c in marcados):
        avisos.append("Nenhum turno na segunda ou na sexta-feira (diretriz 1: é obrigatório ter ao menos um).")
    if n and not any(c.startswith(("ter_", "qua_", "qui_")) for c in marcados):
        avisos.append("Nenhum turno na terça, quarta ou quinta-feira (diretriz 2: o planejamento é cumprido nesses dias).")
    if n and n < TURNOS_MINIMOS:
        avisos.append(f"Apenas {n} turno(s) marcado(s) — a diretriz 4 pede {TURNOS_MINIMOS} turnos para a matrícula.")
    if dados.get("dupla_interesse") and n > TURNOS_MAX_DUPLA:
        avisos.append(f"{n} turnos marcados — com dupla regência, a diretriz 5 pede até {TURNOS_MAX_DUPLA}.")
    return avisos


def _ler_formulario(form) -> dict:
    dados = {
        "situacao": (form.get("situacao") or "").strip(),
        "disciplina": (form.get("disciplina") or "").strip(),
        "dupla_interesse": 1 if form.get("dupla_interesse") == "sim" else 0,
        "observacoes": (form.get("observacoes") or "").strip(),
        "_dupla_resposta": form.get("dupla_interesse") or "",
    }
    try:
        dados["dupla_tempos"] = int(form.get("dupla_tempos") or 0)
    except ValueError:
        dados["dupla_tempos"] = -1
    if not dados["dupla_interesse"]:
        dados["dupla_tempos"] = 0
    for c in CAMPOS_TURNO:
        dados[c] = 1 if form.get(c) else 0
    return dados


def _validar(dados: dict, exigir_justificativa: bool) -> list:
    """Erros que impedem salvar. Quando a disponibilidade foge das diretrizes, o
    professor precisa explicar o motivo em Observações (a gestão pode salvar sem isso)."""
    erros = []
    if dados["situacao"] not in SITUACOES:
        erros.append("Escolha a situação funcional.")
    if dados["disciplina"] not in DISCIPLINAS:
        erros.append("Escolha a disciplina.")
    if dados["_dupla_resposta"] not in ("sim", "nao"):
        erros.append("Responda se tem interesse em dupla regência.")
    elif dados["dupla_interesse"] and not (1 <= dados["dupla_tempos"] <= TEMPOS_DUPLA_MAX):
        erros.append("Você marcou interesse em dupla regência: escolha quantos tempos (de 1 a 12).")
    if not any(dados.get(c) for c in CAMPOS_TURNO):
        erros.append("Marque pelo menos um turno no quadro de disponibilidade.")
    elif exigir_justificativa and _avaliar_diretrizes(dados) and not dados["observacoes"]:
        erros.append("Sua disponibilidade não segue todas as diretrizes (veja os avisos em laranja). "
                     "Para enviar assim, explique o motivo em Observações gerais.")
    return erros


def _salvar(conn, coleta_id: int, professor_id: int, dados: dict, atualizado_por: int):
    colunas = ["situacao", "disciplina", "dupla_interesse", "dupla_tempos"] + CAMPOS_TURNO + ["observacoes"]
    valores = [dados[c] for c in colunas]
    lista_cols = ", ".join(colunas)
    marcadores = ", ".join("?" for _ in colunas)
    atualiza = ", ".join(f"{c} = excluded.{c}" for c in colunas)
    conn.execute(
        f"""INSERT INTO horario_disponibilidade (coleta_id, professor_id, {lista_cols}, atualizado_por)
            VALUES (?, ?, {marcadores}, ?)
            ON CONFLICT(coleta_id, professor_id) DO UPDATE SET
                {atualiza}, atualizado_por = excluded.atualizado_por, atualizado_em = CURRENT_TIMESTAMP""",
        [coleta_id, professor_id] + valores + [atualizado_por],
    )


def _pagina(request: Request, titulo: str, conteudo: str, ativo: str) -> HTMLResponse:
    return HTMLResponse(_ctx["render_page"](titulo, conteudo, active=ativo))


def _acesso_negado(request: Request, mensagem: str) -> HTMLResponse:
    conteudo = f"""
        <div class="page-header"><h1>🔒 Acesso restrito</h1></div>
        <div style="background:var(--red-bg); color:var(--red); border:1px solid var(--red); padding:16px; border-radius:6px;">{mensagem}</div>
    """
    return HTMLResponse(_ctx["render_page"]("Acesso restrito", conteudo), status_code=403)


# ==========================================
#  PEDAÇOS DE TELA REAPROVEITADOS
# ==========================================

def _diretrizes_html(coleta) -> str:
    return f"""
    <div class="card" style="padding:14px 16px; margin-bottom:16px; border-left:4px solid var(--purple, #7c3aed);">
        <div style="font-weight:700; margin-bottom:6px;">📌 Diretrizes para a disponibilidade de horário {coleta["ano_letivo"]}</div>
        <ol style="margin:0; padding-left:20px; font-size:13px; line-height:1.6;">
            <li>É obrigatório fornecer turnos disponíveis na <strong>segunda ou sexta-feira</strong>.</li>
            <li>O planejamento será cumprido obrigatoriamente <strong>terça, quarta ou quinta-feira</strong>.</li>
            <li>Preferencialmente disponibilizar o <strong>turno inteiro</strong>. Qualquer especificidade, sinalizar nas observações gerais.</li>
            <li>Todos os docentes deverão disponibilizar, em caso de somente a matrícula, <strong>5 turnos</strong> de trabalho.</li>
            <li>Caso tenha interesse em dupla regência, disponibilizar <strong>até 8 turnos</strong> de trabalho.</li>
            <li><strong>Atenção:</strong> no próximo ano haverá realocação de turmas, com maior número de estudantes no <strong>turno da tarde</strong> — ajustes na disponibilidade deverão ser realizados.</li>
        </ol>
    </div>"""


def _mini_quadro(resp) -> str:
    """Quadrinho 2×5 (manhã/tarde × dias) para listas e resumos."""
    cab = "".join(f'<th style="padding:1px 4px; font-weight:600; font-size:10px; color:var(--text-muted);">{nome[:3]}</th>' for _, nome in DIAS)
    linhas = ""
    for t, nome_t in TURNOS:
        celulas = ""
        for d, _ in DIAS:
            on = bool(resp[f"{d}_{t}"])
            cor = "background:#16a34a; color:white;" if on else "background:var(--bg-subtle, #f1f5f9); color:var(--text-muted);"
            celulas += f'<td style="{cor} text-align:center; padding:1px 4px; font-size:10px; border-radius:3px;">{"✓" if on else "·"}</td>'
        linhas += f'<tr><td style="font-size:10px; color:var(--text-muted); padding-right:4px;">{nome_t[0]}</td>{celulas}</tr>'
    return f'<table style="border-collapse:separate; border-spacing:2px;"><tr><td></td>{cab}</tr>{linhas}</table>'


def _resumo_resposta_html(resp) -> str:
    n = len(_turnos_marcados(resp))
    dupla = f'Sim — {resp["dupla_tempos"]} tempo(s)' if resp["dupla_interesse"] else "Não"
    obs = _e(resp["observacoes"]) or '<span style="color:var(--text-muted);">—</span>'
    avisos = _avaliar_diretrizes(dict(resp))
    avisos_html = ""
    if avisos:
        avisos_html = '<div style="margin-top:10px; font-size:12px; color:#b45309;">' + "".join(f"⚠️ {_e(a)}<br>" for a in avisos) + "</div>"
    return f"""
    <div style="display:flex; flex-wrap:wrap; gap:24px; align-items:flex-start;">
        <div style="font-size:13px; line-height:1.8;">
            <div><strong>Situação funcional:</strong> {_e(resp["situacao"])}</div>
            <div><strong>Disciplina:</strong> {_e(resp["disciplina"])}</div>
            <div><strong>Interesse em dupla regência:</strong> {dupla}</div>
            <div><strong>Turnos disponíveis:</strong> {n}</div>
            <div><strong>Última atualização:</strong> {_ts_br(resp["atualizado_em"])}</div>
        </div>
        <div>{_mini_quadro(resp)}</div>
    </div>
    <div style="margin-top:10px; font-size:13px;"><strong>Observações gerais:</strong> {obs}</div>
    {avisos_html}"""


_CSS_FORM = """
<style>
.hd-quadro { border-collapse: separate; border-spacing: 6px; margin: 4px 0 8px; }
.hd-quadro th { font-size: 12px; font-weight: 600; color: var(--text-muted); padding: 2px 4px; }
.hd-quadro td.hd-rot { font-size: 13px; font-weight: 600; padding-right: 6px; white-space: nowrap; }
.hd-cel label { display: block; margin: 0; cursor: pointer; position: relative; }
.hd-cel input { position: absolute; opacity: 0; width: 1px; height: 1px; }
.hd-cel span { display: flex; align-items: center; justify-content: center; width: 64px; height: 44px;
    border: 2px solid var(--border, #cbd5e1); border-radius: 8px; font-size: 18px; color: transparent;
    user-select: none; transition: background .12s, border-color .12s; }
.hd-cel label:hover span { border-color: #16a34a; }
.hd-cel input:checked + span { background: #16a34a; border-color: #16a34a; color: white; }
.hd-cel input:focus-visible + span { outline: 2px solid #2563eb; outline-offset: 3px; }
@media (max-width: 560px) { .hd-cel span { width: 46px; height: 40px; } .hd-quadro { border-spacing: 4px; } }
.hd-avisos { font-size: 13px; color: #b45309; background: #fffbeb; border: 1px solid #fcd34d;
    border-radius: 6px; padding: 8px 12px; margin: 8px 0; display: none; }
.hd-erros { background: var(--red-bg, #fef2f2); color: var(--red, #b91c1c); border: 1px solid var(--red, #b91c1c);
    padding: 10px 14px; border-radius: 6px; margin-bottom: 14px; font-size: 13px; }
.hd-campo { margin-bottom: 14px; }
.hd-campo > span { display: block; font-weight: 600; font-size: 13px; margin-bottom: 4px; }
.hd-radio { display: inline-flex; align-items: center; gap: 6px; margin-right: 18px; font-weight: 400; cursor: pointer; }
.hd-radio input { width: auto; margin: 0; }
</style>
"""

_JS_FORM = """
<script>
(function () {
  var form = document.getElementById('hd-form');
  if (!form) return;
  var MIN = %(min)d, MAX_DUPLA = %(max_dupla)d;
  var caixas = form.querySelectorAll('.hd-cel input');
  var contador = document.getElementById('hd-contador');
  var avisos = document.getElementById('hd-avisos');
  var obsMarca = document.getElementById('hd-obs-marca');
  var tempos = form.querySelector('select[name=dupla_tempos]');
  function duplaSim() { var r = form.querySelector('input[name=dupla_interesse]:checked'); return r && r.value === 'sim'; }
  function atualizar() {
    var marc = [];
    caixas.forEach(function (c) { if (c.checked) marc.push(c.name); });
    var n = marc.length;
    contador.textContent = n + (n === 1 ? ' turno marcado' : ' turnos marcados');
    var lista = [];
    if (n) {
      if (!marc.some(function (x) { return x.indexOf('seg_') === 0 || x.indexOf('sex_') === 0; }))
        lista.push('Nenhum turno na segunda ou na sexta-feira (diretriz 1).');
      if (!marc.some(function (x) { return /^(ter|qua|qui)_/.test(x); }))
        lista.push('Nenhum turno na terça, quarta ou quinta-feira (diretriz 2 — planejamento).');
      if (n < MIN) lista.push('Apenas ' + n + ' turno(s) — a diretriz 4 pede ' + MIN + ' para a matrícula.');
      if (duplaSim() && n > MAX_DUPLA) lista.push(n + ' turnos — com dupla regência, a diretriz 5 pede até ' + MAX_DUPLA + '.');
    }
    if (lista.length) {
      avisos.innerHTML = lista.map(function (a) { return '⚠️ ' + a; }).join('<br>') +
        '<br><strong>Você pode enviar mesmo assim, desde que explique o motivo em Observações gerais.</strong>';
      avisos.style.display = 'block';
      if (obsMarca) obsMarca.style.display = 'inline';
    } else {
      avisos.style.display = 'none';
      if (obsMarca) obsMarca.style.display = 'none';
    }
    if (tempos) {
      if (duplaSim()) { tempos.disabled = false; }
      else { tempos.value = '0'; tempos.disabled = true; }
    }
  }
  form.addEventListener('change', atualizar);
  atualizar();
})();
</script>
"""


def _form_html(action: str, coleta, dados: Optional[dict], erros: list, nome_alvo: str,
               modo_gestao: bool, texto_botao: str) -> str:
    d = dados or {}

    def sel(valor, atual):
        return " selected" if valor == atual else ""

    opts_sit = '<option value="">Escolha…</option>' + "".join(
        f'<option{sel(s, d.get("situacao"))}>{s}</option>' for s in SITUACOES)
    opts_disc = '<option value="">Escolha…</option>' + "".join(
        f'<option{sel(x, d.get("disciplina"))}>{x}</option>' for x in DISCIPLINAS)
    dupla_resp = d.get("_dupla_resposta") or ("sim" if d.get("dupla_interesse") else ("nao" if dados else ""))
    tempos_atual = int(d.get("dupla_tempos") or 0)
    opts_tempos = "".join(
        f'<option value="{i}"{" selected" if i == tempos_atual else ""}>{i}</option>' for i in range(0, TEMPOS_DUPLA_MAX + 1))

    cab = "".join(f"<th>{nome}</th>" for _, nome in DIAS)
    linhas = ""
    for t, nome_t in TURNOS:
        celulas = ""
        for dd, nome_d in DIAS:
            campo = f"{dd}_{t}"
            marcado = " checked" if d.get(campo) else ""
            celulas += (f'<td class="hd-cel"><label title="{nome_d} — {nome_t}">'
                        f'<input type="checkbox" name="{campo}" value="1"{marcado}><span>✓</span></label></td>')
        linhas += f'<tr><td class="hd-rot">{nome_t}</td>{celulas}</tr>'

    erros_html = ""
    if erros:
        erros_html = '<div class="hd-erros">' + "".join(f"• {_e(x)}<br>" for x in erros) + "</div>"

    aviso_gestao = ""
    if modo_gestao:
        aviso_gestao = ('<div class="tip" style="margin-bottom:14px;">Você está preenchendo como <strong>gestão</strong>. '
                        'Os avisos das diretrizes aparecem, mas não impedem salvar.</div>')

    return f"""
    {_CSS_FORM}
    {erros_html}
    {aviso_gestao}
    <form id="hd-form" method="post" action="{action}" class="card" style="padding:18px 20px;">
        <div class="hd-campo"><span>1. Nome</span>
            <div style="font-size:15px;">{_e(nome_alvo)}</div>
            <div style="font-size:11px; color:var(--text-muted);">Vem do seu cadastro no sistema.</div>
        </div>
        <label class="hd-campo"><span>2. Situação funcional *</span>
            <select name="situacao" required style="max-width:320px;">{opts_sit}</select>
        </label>
        <label class="hd-campo"><span>3. Disciplina *</span>
            <select name="disciplina" required style="max-width:320px;">{opts_disc}</select>
        </label>
        <div class="hd-campo"><span>4. Possui interesse em dupla regência? *</span>
            <label class="hd-radio"><input type="radio" name="dupla_interesse" value="sim" required{" checked" if dupla_resp == "sim" else ""}> Sim</label>
            <label class="hd-radio"><input type="radio" name="dupla_interesse" value="nao" required{" checked" if dupla_resp == "nao" else ""}> Não</label>
        </div>
        <label class="hd-campo"><span>5. Quantos tempos de dupla regência? (sem interesse, fica 0)</span>
            <select name="dupla_tempos" style="max-width:120px;">{opts_tempos}</select>
        </label>
        <div class="hd-campo"><span>6. Marque no quadro os turnos em que você está disponível para a Escola *</span>
            <table class="hd-quadro"><tr><th></th>{cab}</tr>{linhas}</table>
            <div id="hd-contador" style="font-size:13px; font-weight:600;"></div>
            <div id="hd-avisos" class="hd-avisos"></div>
        </div>
        <label class="hd-campo"><span>7. Observações gerais <span id="hd-obs-marca" style="display:none; color:#b45309;">(obrigatório explicar o motivo)</span></span>
            <textarea name="observacoes" rows="4" placeholder="Ex.: disponível na quinta só até as 16h; preferência por 6º e 9º anos…">{_e(d.get("observacoes"))}</textarea>
        </label>
        <div class="page-actions" style="margin-top:6px;">
            <button type="submit" class="btn btn-primary">{texto_botao}</button>
        </div>
    </form>
    {_JS_FORM % {"min": TURNOS_MINIMOS, "max_dupla": TURNOS_MAX_DUPLA}}
    """


# ==========================================
#  TELA DO PROFESSOR
# ==========================================

@router.get("/horarios/disponibilidade", response_class=HTMLResponse)
def disponibilidade_professor(request: Request, salvo: str = ""):
    prof = _ctx["get_prof"](request)
    if not _pode_responder(prof):
        return _acesso_negado(request, "Esta área é para os professores informarem a disponibilidade de horário.")
    conn = _db()
    coleta = _coleta_atual(conn)
    resp = None
    if coleta:
        resp = conn.execute("SELECT * FROM horario_disponibilidade WHERE coleta_id = ? AND professor_id = ?",
                            (coleta["id"], prof["id"])).fetchone()
    conn.close()
    return _render_tela_professor(request, prof, coleta, resp, dados_form=None, erros=[], salvo=(salvo == "1"))


def _render_tela_professor(request, prof, coleta, resp, dados_form, erros, salvo):
    gestao = _eh_gestao(prof)
    link_painel = ('<a href="/horarios/disponibilidade/painel" class="btn" style="margin-left:8px;">📊 Painel da gestão</a>'
                   if gestao else "")
    if not coleta:
        conteudo = f"""
            <div class="page-header"><h1>🗓️ Disponibilidade de horário</h1></div>
            <div class="empty">Nenhuma coleta de disponibilidade foi criada ainda. {link_painel}</div>"""
        return _pagina(request, "Disponibilidade de horário", conteudo, "horarios-disponibilidade")

    status = _status_coleta(coleta)
    periodo = f'de {_data_br(coleta["data_inicio"])} a {_data_br(coleta["data_fim"])}'
    cabecalho = f"""
        <div class="page-header">
            <h1>🗓️ Disponibilidade de horário {coleta["ano_letivo"]}</h1>
            <p class="subtitle">Coleta {periodo}. Suas respostas são usadas para gerar o horário escolar de {coleta["ano_letivo"]}. {link_painel}</p>
        </div>"""

    msg_salvo = ""
    if salvo:
        msg_salvo = ('<div style="background:#ecfdf5; color:#047857; border:1px solid #10b981; padding:12px 14px; '
                     'border-radius:6px; margin-bottom:14px; font-size:14px;">✅ <strong>Disponibilidade salva!</strong> '
                     f'Você pode voltar e alterar até {_data_br(coleta["data_fim"])}.</div>')

    pode_editar = status == "aberta" or (gestao and status == "agendada")

    if pode_editar:
        aviso_previa = ""
        if status == "agendada":
            aviso_previa = (f'<div class="tip" style="margin-bottom:14px;">👀 <strong>Prévia da gestão:</strong> a coleta só abre para os '
                            f'professores em {_data_br(coleta["data_inicio"])}. Você pode testar o formulário agora — '
                            'depois, exclua a resposta de teste no painel.</div>')
        if dados_form is None and resp is not None:
            dados_form = dict(resp)
        texto_botao = "💾 Atualizar minha disponibilidade" if resp else "📨 Enviar minha disponibilidade"
        ja_respondeu = ""
        if resp and not salvo:
            ja_respondeu = (f'<div class="tip" style="margin-bottom:14px;">Você já respondeu em {_ts_br(resp["atualizado_em"])}. '
                            'Se precisar, altere abaixo e clique em atualizar.</div>')
        conteudo = (cabecalho + msg_salvo + aviso_previa + ja_respondeu + _diretrizes_html(coleta) +
                    _form_html("/horarios/disponibilidade", coleta, dados_form, erros, prof["nome"], False, texto_botao))
        return _pagina(request, "Disponibilidade de horário", conteudo, "horarios-disponibilidade")

    if status == "agendada":
        corpo = (f'<div class="card" style="padding:18px;"><p style="font-size:15px; margin:0;">⏳ A coleta abre em '
                 f'<strong>{_data_br(coleta["data_inicio"])}</strong> e vai até <strong>{_data_br(coleta["data_fim"])}</strong>.</p>'
                 '<p style="font-size:13px; color:var(--text-muted); margin:8px 0 0;">Leia as diretrizes abaixo enquanto isso.</p></div>'
                 + _diretrizes_html(coleta))
    else:  # encerrada
        if resp:
            corpo = (msg_salvo + '<div class="card" style="padding:18px;"><p style="margin:0 0 12px; font-size:14px;">🔒 O prazo da coleta terminou. '
                     'Esta é a disponibilidade que você enviou (para alterar, fale com a gestão):</p>'
                     + _resumo_resposta_html(resp) + "</div>")
        else:
            corpo = ('<div class="card" style="padding:18px;"><p style="margin:0; font-size:14px;">🔒 O prazo da coleta terminou e não '
                     'encontramos resposta sua. <strong>Procure a gestão</strong> para informar sua disponibilidade.</p></div>')
    return _pagina(request, "Disponibilidade de horário", cabecalho + corpo, "horarios-disponibilidade")


@router.post("/horarios/disponibilidade", response_class=HTMLResponse)
async def salvar_disponibilidade_professor(request: Request):
    prof = _ctx["get_prof"](request)
    if not _pode_responder(prof):
        return _acesso_negado(request, "Esta área é para os professores informarem a disponibilidade de horário.")
    form = await request.form()
    conn = _db()
    coleta = _coleta_atual(conn)
    status = _status_coleta(coleta)
    gestao = _eh_gestao(prof)
    if not coleta or not (status == "aberta" or (gestao and status == "agendada")):
        conn.close()
        return RedirectResponse("/horarios/disponibilidade", status_code=303)
    dados = _ler_formulario(form)
    erros = _validar(dados, exigir_justificativa=True)
    if erros:
        resp = conn.execute("SELECT * FROM horario_disponibilidade WHERE coleta_id = ? AND professor_id = ?",
                            (coleta["id"], prof["id"])).fetchone()
        conn.close()
        return _render_tela_professor(request, prof, coleta, resp, dados_form=dados, erros=erros, salvo=False)
    _salvar(conn, coleta["id"], prof["id"], dados, atualizado_por=prof["id"])
    # Se a gestão tinha dispensado este professor, a resposta dele passa a valer.
    conn.execute("DELETE FROM horario_coleta_dispensas WHERE coleta_id = ? AND professor_id = ?", (coleta["id"], prof["id"]))
    conn.commit()
    conn.close()
    return RedirectResponse("/horarios/disponibilidade?salvo=1", status_code=303)


# ==========================================
#  PAINEL DA GESTÃO
# ==========================================

def _dados_painel(conn, coleta) -> dict:
    respostas = conn.execute("""
        SELECT hd.*, p.nome AS professor_nome, p.email AS professor_email,
               pa.nome AS atualizado_por_nome
        FROM horario_disponibilidade hd
        JOIN professores p ON p.id = hd.professor_id
        LEFT JOIN professores pa ON pa.id = hd.atualizado_por
        WHERE hd.coleta_id = ?
        ORDER BY p.nome COLLATE NOCASE
    """, (coleta["id"],)).fetchall()
    respondidos_ids = {r["professor_id"] for r in respostas}
    dispensados = conn.execute("""
        SELECT p.id, p.nome, p.email FROM horario_coleta_dispensas d
        JOIN professores p ON p.id = d.professor_id
        WHERE d.coleta_id = ? ORDER BY p.nome COLLATE NOCASE
    """, (coleta["id"],)).fetchall()
    dispensados_ids = {d["id"] for d in dispensados}
    # Quem deveria responder: todo docente ativo (papel "docente").
    docentes = conn.execute("""
        SELECT id, nome, email FROM professores
        WHERE COALESCE(status, 'ativo') = 'ativo' AND papel = 'docente'
        ORDER BY nome COLLATE NOCASE
    """).fetchall()
    pendentes = [p for p in docentes if p["id"] not in respondidos_ids and p["id"] not in dispensados_ids]
    contagem = {c: sum(1 for r in respostas if r[c]) for c in CAMPOS_TURNO}
    return {"respostas": respostas, "pendentes": pendentes, "dispensados": dispensados, "contagem": contagem}


@router.get("/horarios/disponibilidade/painel", response_class=HTMLResponse)
def painel_disponibilidade(request: Request, coleta_id: str = "", msg: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão acessa o painel de disponibilidade.")
    conn = _db()
    coletas = conn.execute("SELECT * FROM horario_coletas ORDER BY ano_letivo DESC").fetchall()
    coleta = _coleta_por_id(conn, coleta_id) if coleta_id else (coletas[0] if coletas else None)
    if not coleta:
        conn.close()
        conteudo = f"""
            <div class="page-header"><h1>📊 Painel de disponibilidade</h1></div>
            <div class="empty">Nenhuma coleta criada.</div>
            {_form_nova_coleta_html()}"""
        return _pagina(request, "Painel de disponibilidade", conteudo, "horarios-painel")
    dp = _dados_painel(conn, coleta)
    conn.close()

    cid = coleta["id"]
    status = _status_coleta(coleta)
    rotulo_status = {
        "aberta": ('<span style="background:#dcfce7; color:#166534; padding:2px 10px; border-radius:12px; font-weight:600;">● Aberta</span>'),
        "agendada": ('<span style="background:#e0e7ff; color:#3730a3; padding:2px 10px; border-radius:12px; font-weight:600;">⏳ Agendada</span>'),
        "encerrada": ('<span style="background:#f1f5f9; color:#475569; padding:2px 10px; border-radius:12px; font-weight:600;">🔒 Encerrada</span>'),
    }[status]

    respondidos = len(dp["respostas"])
    pendentes = len(dp["pendentes"])
    esperado = respondidos + pendentes
    pct = int(round(100 * respondidos / esperado)) if esperado else 0

    seletor = ""
    if len(coletas) > 1:
        opts = "".join(f'<option value="{c["id"]}"{" selected" if c["id"] == cid else ""}>Coleta {c["ano_letivo"]}</option>' for c in coletas)
        seletor = (f'<form method="get" style="display:inline-block; margin-left:10px;">'
                   f'<select name="coleta_id" onchange="this.form.submit();" style="width:auto;">{opts}</select></form>')

    msgs = {
        "periodo": "Período da coleta atualizado.",
        "encerrada": "Coleta encerrada. Os professores não conseguem mais enviar ou alterar respostas.",
        "reaberta": "Coleta reaberta.",
        "excluida": "Resposta excluída.",
        "dispensado": "Professor dispensado desta coleta.",
        "reincluido": "Professor voltou para a lista de pendentes.",
        "salva": "Disponibilidade salva.",
        "criada": "Nova coleta criada.",
        "erro_datas": "Datas inválidas: confira se o fim não é antes do início.",
        "erro_ano": "Já existe uma coleta para esse ano letivo.",
    }
    msg_html = ""
    if msg in msgs:
        erro = msg.startswith("erro")
        cor = "background:var(--red-bg,#fef2f2); color:var(--red,#b91c1c); border:1px solid var(--red,#b91c1c);" if erro else \
              "background:#ecfdf5; color:#047857; border:1px solid #10b981;"
        msg_html = f'<div style="{cor} padding:10px 14px; border-radius:6px; margin-bottom:14px; font-size:13px;">{msgs[msg]}</div>'

    # Botão encerrar / reabrir
    if coleta["encerrada"]:
        botao_estado = (f'<form method="post" action="/horarios/coletas/{cid}/reabrir" style="display:inline;">'
                        '<button class="btn" type="submit">🔓 Reabrir coleta</button></form>')
    elif status != "encerrada":
        botao_estado = (f'<form method="post" action="/horarios/coletas/{cid}/encerrar" style="display:inline;" '
                        'onsubmit="return confirm(\'Encerrar a coleta agora? Os professores não poderão mais enviar nem alterar.\');">'
                        '<button class="btn" type="submit" style="color:var(--red); border-color:var(--red);">🔒 Encerrar agora</button></form>')
    else:
        botao_estado = '<span style="font-size:12px; color:var(--text-muted);">Para reabrir, aumente a data de fim.</span>'

    cabecalho = f"""
        <div class="page-header">
            <h1>📊 Painel de disponibilidade — {coleta["ano_letivo"]} {seletor}</h1>
            <p class="subtitle">{rotulo_status} &nbsp; Coleta de {_data_br(coleta["data_inicio"])} a {_data_br(coleta["data_fim"])}</p>
        </div>
        {msg_html}
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <div style="display:flex; flex-wrap:wrap; gap:18px; align-items:center; justify-content:space-between;">
                <div style="flex:1 1 260px;">
                    <div style="font-size:28px; font-weight:800;">{respondidos} <span style="font-size:15px; font-weight:500; color:var(--text-muted);">de {esperado} responderam</span></div>
                    <div style="height:10px; background:var(--bg-subtle,#e2e8f0); border-radius:6px; overflow:hidden; margin-top:6px;">
                        <div style="height:100%; width:{pct}%; background:#16a34a;"></div>
                    </div>
                    <div style="font-size:12px; color:var(--text-muted); margin-top:4px;">{pendentes} pendente(s) · {len(dp["dispensados"])} dispensado(s)</div>
                </div>
                <div style="display:flex; flex-wrap:wrap; gap:8px;">
                    <a href="/horarios/disponibilidade/exportar?coleta_id={cid}" class="btn btn-primary">📥 Exportar Excel</a>
                    <a href="/horarios/disponibilidade" class="btn">📝 Ver o formulário</a>
                    {botao_estado}
                </div>
            </div>
            <details style="margin-top:12px;">
                <summary style="cursor:pointer; font-size:13px; font-weight:600;">✏️ Alterar período da coleta</summary>
                <form method="post" action="/horarios/coletas/{cid}/periodo" style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin-top:10px;">
                    <label style="margin:0;">Início<input type="date" name="data_inicio" value="{_e(coleta["data_inicio"])}" required></label>
                    <label style="margin:0;">Fim<input type="date" name="data_fim" value="{_e(coleta["data_fim"])}" required></label>
                    <button type="submit" class="btn">Salvar período</button>
                </form>
            </details>
        </div>"""

    # Disponibilidade por turno — mostra onde falta gente (ex.: tarde, com a realocação de turmas)
    maior = max(dp["contagem"].values()) if dp["contagem"] else 0
    cab = "".join(f'<th style="padding:6px 10px;">{nome}</th>' for _, nome in DIAS)
    linhas_turno = ""
    for t, nome_t in TURNOS:
        tds = ""
        for d, _ in DIAS:
            v = dp["contagem"][f"{d}_{t}"]
            intensidade = (v / maior) if maior else 0
            fundo = f"rgba(22,163,74,{0.10 + 0.55 * intensidade:.2f})" if v else "transparent"
            tds += f'<td style="padding:8px 10px; text-align:center; font-weight:700; font-size:15px; background:{fundo}; border-radius:4px;">{v}</td>'
        linhas_turno += f'<tr><td style="padding:6px 10px; font-weight:600;">{nome_t}</td>{tds}</tr>'
    quadro_turnos = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 4px; font-size:15px;">Professores disponíveis por turno</h2>
            <p style="margin:0 0 8px; font-size:12px; color:var(--text-muted);">Quantos dos {respondidos} que responderam marcaram cada turno.</p>
            <div style="overflow-x:auto;"><table style="border-collapse:separate; border-spacing:4px;"><tr><th></th>{cab}</tr>{linhas_turno}</table></div>
        </div>"""

    # Respondidos
    linhas_resp = ""
    for r in dp["respostas"]:
        n = len(_turnos_marcados(r))
        avisos = _avaliar_diretrizes(dict(r))
        alerta = (f'<span title="{_e(" ".join(avisos))}" style="color:#b45309; font-weight:600; cursor:help;">⚠️ {len(avisos)}</span>'
                  if avisos else '<span style="color:#16a34a;">✓</span>')
        dupla = f'Sim ({r["dupla_tempos"]})' if r["dupla_interesse"] else "Não"
        por_gestao = ""
        if r["atualizado_por"] and r["atualizado_por"] != r["professor_id"]:
            por_gestao = f'<div style="font-size:10px; color:var(--text-muted);">por {_e(r["atualizado_por_nome"] or "gestão")}</div>'
        obs_txt = (r["observacoes"] or "").strip()
        if len(obs_txt) > 110:
            # Observação longa: mostra o começo e um "ver tudo" que abre o texto inteiro,
            # para a tabela não ficar com linhas gigantes.
            obs = (f'{_e(obs_txt[:110].rstrip())}… <details style="display:inline;"><summary style="display:inline; cursor:pointer; '
                   f'color:#2563eb; font-size:11px;">ver tudo</summary><div style="margin-top:4px;">{_e(obs_txt)}</div></details>')
        else:
            obs = _e(obs_txt)
        linhas_resp += f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:8px;"><strong>{_e(r["professor_nome"])}</strong></td>
            <td style="padding:8px;">{_e(r["situacao"])}</td>
            <td style="padding:8px;">{_e(r["disciplina"])}</td>
            <td style="padding:8px; text-align:center;">{dupla}</td>
            <td style="padding:8px;">{_mini_quadro(r)}</td>
            <td style="padding:8px; text-align:center; font-weight:700;">{n}</td>
            <td style="padding:8px; text-align:center;">{alerta}</td>
            <td style="padding:8px; font-size:12px; max-width:260px;">{obs or '<span style="color:var(--text-muted);">—</span>'}</td>
            <td style="padding:8px; font-size:11px; white-space:nowrap;">{_ts_br(r["atualizado_em"])}{por_gestao}</td>
            <td style="padding:8px; white-space:nowrap;">
                <a href="/horarios/disponibilidade/professor/{r["professor_id"]}?coleta_id={cid}" class="btn" style="padding:2px 8px; font-size:11px;" title="Editar">✏️</a>
                <form method="post" action="/horarios/disponibilidade/professor/{r["professor_id"]}/excluir?coleta_id={cid}" style="display:inline;"
                      onsubmit="return confirm('Excluir esta resposta? O professor volta para a lista de pendentes.');">
                    <button type="submit" class="btn" style="padding:2px 8px; font-size:11px; color:var(--red);" title="Excluir">🗑️</button>
                </form>
            </td>
        </tr>"""
    if not linhas_resp:
        linhas_resp = '<tr><td colspan="10" style="padding:16px; text-align:center; color:var(--text-muted);">Ninguém respondeu ainda.</td></tr>'
    tabela_resp = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 8px; font-size:15px;">✅ Responderam ({respondidos})</h2>
            <p style="margin:0 0 8px; font-size:12px; color:var(--text-muted);">⚠️ = disponibilidade fora de alguma diretriz (passe o mouse para ver qual). O professor precisou justificar nas observações.</p>
            <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);">
                    <th style="padding:8px; text-align:left;">Professor</th><th style="padding:8px; text-align:left;">Situação</th>
                    <th style="padding:8px; text-align:left;">Disciplina</th><th style="padding:8px;">Dupla</th>
                    <th style="padding:8px; text-align:left;">Turnos</th><th style="padding:8px;">Total</th>
                    <th style="padding:8px;">Diretrizes</th><th style="padding:8px; text-align:left;">Observações</th>
                    <th style="padding:8px; text-align:left;">Atualizado</th><th style="padding:8px;"></th>
                </tr></thead>
                <tbody>{linhas_resp}</tbody>
            </table>
            </div>
        </div>"""

    # Pendentes
    linhas_pend = ""
    for p in dp["pendentes"]:
        linhas_pend += f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:8px;">{_e(p["nome"])}</td>
            <td style="padding:8px; font-size:12px; color:var(--text-muted);">{_e(p["email"])}</td>
            <td style="padding:8px; white-space:nowrap;">
                <a href="/horarios/disponibilidade/professor/{p["id"]}?coleta_id={cid}" class="btn" style="padding:2px 8px; font-size:11px;">📝 Preencher por ele(a)</a>
                <form method="post" action="/horarios/disponibilidade/professor/{p["id"]}/dispensar?coleta_id={cid}" style="display:inline;">
                    <button type="submit" class="btn" style="padding:2px 8px; font-size:11px;" title="Não precisa responder (ex.: não estará na escola no próximo ano)">Dispensar</button>
                </form>
            </td>
        </tr>"""
    if not linhas_pend:
        linhas_pend = '<tr><td colspan="3" style="padding:14px; text-align:center; color:var(--text-muted);">Nenhum pendente. 🎉</td></tr>'
    tabela_pend = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 4px; font-size:15px;">⏳ Pendentes ({pendentes})</h2>
            <p style="margin:0 0 8px; font-size:12px; color:var(--text-muted);">Docentes ativos no sistema que ainda não responderam. "Dispensar" tira da lista quem não precisa responder.</p>
            <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; font-size:13px;"><tbody>{linhas_pend}</tbody></table>
            </div>
        </div>"""

    tabela_disp = ""
    if dp["dispensados"]:
        linhas_disp = "".join(f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:8px;">{_e(d["nome"])}</td>
            <td style="padding:8px;"><form method="post" action="/horarios/disponibilidade/professor/{d["id"]}/reincluir?coleta_id={cid}" style="display:inline;">
                <button type="submit" class="btn" style="padding:2px 8px; font-size:11px;">↩️ Voltar para pendentes</button></form></td>
        </tr>""" for d in dp["dispensados"])
        tabela_disp = f"""
        <details class="card" style="padding:14px 16px; margin-bottom:16px;">
            <summary style="cursor:pointer; font-weight:600;">Dispensados ({len(dp["dispensados"])})</summary>
            <table style="width:100%; border-collapse:collapse; font-size:13px; margin-top:8px;"><tbody>{linhas_disp}</tbody></table>
        </details>"""

    conteudo = cabecalho + quadro_turnos + tabela_resp + tabela_pend + tabela_disp + _form_nova_coleta_html()
    return _pagina(request, "Painel de disponibilidade", conteudo, "horarios-painel")


def _form_nova_coleta_html() -> str:
    ano_sugerido = _hoje().year + 1
    return f"""
        <details class="card" style="padding:14px 16px; margin-bottom:16px;">
            <summary style="cursor:pointer; font-weight:600;">➕ Criar coleta para outro ano letivo</summary>
            <form method="post" action="/horarios/coletas/nova" style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin-top:10px;">
                <label style="margin:0;">Ano letivo<input type="number" name="ano_letivo" value="{ano_sugerido}" min="2025" max="2100" required style="width:110px;"></label>
                <label style="margin:0;">Início<input type="date" name="data_inicio" required></label>
                <label style="margin:0;">Fim<input type="date" name="data_fim" required></label>
                <button type="submit" class="btn">Criar coleta</button>
            </form>
        </details>"""


# ---------- gestão: preencher/editar a resposta de um professor ----------

@router.get("/horarios/disponibilidade/professor/{professor_id}", response_class=HTMLResponse)
def gestao_form_professor(request: Request, professor_id: int, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão pode editar a disponibilidade de outro professor.")
    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id) or _coleta_atual(conn)
    alvo = conn.execute("SELECT id, nome FROM professores WHERE id = ?", (professor_id,)).fetchone()
    if not coleta or not alvo:
        conn.close()
        return RedirectResponse("/horarios/disponibilidade/painel", status_code=303)
    resp = conn.execute("SELECT * FROM horario_disponibilidade WHERE coleta_id = ? AND professor_id = ?",
                        (coleta["id"], professor_id)).fetchone()
    conn.close()
    return _render_form_gestao(request, coleta, alvo, dict(resp) if resp else None, [])


def _render_form_gestao(request, coleta, alvo, dados, erros):
    conteudo = f"""
        <div class="page-header">
            <h1>✏️ Disponibilidade de {_e(alvo["nome"])}</h1>
            <p class="subtitle">Coleta {coleta["ano_letivo"]} · <a href="/horarios/disponibilidade/painel?coleta_id={coleta["id"]}">← voltar ao painel</a></p>
        </div>
        {_form_html(f'/horarios/disponibilidade/professor/{alvo["id"]}?coleta_id={coleta["id"]}', coleta, dados, erros,
                    alvo["nome"], True, "💾 Salvar disponibilidade")}"""
    return _pagina(request, "Editar disponibilidade", conteudo, "horarios-painel")


@router.post("/horarios/disponibilidade/professor/{professor_id}/excluir")
def gestao_excluir_resposta(request: Request, professor_id: int, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão pode excluir respostas.")
    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id)
    if coleta:
        conn.execute("DELETE FROM horario_disponibilidade WHERE coleta_id = ? AND professor_id = ?", (coleta["id"], professor_id))
        conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=excluida", status_code=303)


@router.post("/horarios/disponibilidade/professor/{professor_id}/dispensar")
def gestao_dispensar(request: Request, professor_id: int, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão pode dispensar professores da coleta.")
    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id)
    if coleta:
        conn.execute("INSERT OR IGNORE INTO horario_coleta_dispensas (coleta_id, professor_id) VALUES (?, ?)", (coleta["id"], professor_id))
        conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=dispensado", status_code=303)


@router.post("/horarios/disponibilidade/professor/{professor_id}/reincluir")
def gestao_reincluir(request: Request, professor_id: int, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão pode alterar a lista da coleta.")
    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id)
    if coleta:
        conn.execute("DELETE FROM horario_coleta_dispensas WHERE coleta_id = ? AND professor_id = ?", (coleta["id"], professor_id))
        conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=reincluido", status_code=303)


# Rota genérica de salvar fica DEPOIS das específicas (/excluir, /dispensar, /reincluir).
@router.post("/horarios/disponibilidade/professor/{professor_id}", response_class=HTMLResponse)
async def gestao_salvar_professor(request: Request, professor_id: int, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão pode editar a disponibilidade de outro professor.")
    form = await request.form()
    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id)
    alvo = conn.execute("SELECT id, nome FROM professores WHERE id = ?", (professor_id,)).fetchone()
    if not coleta or not alvo:
        conn.close()
        return RedirectResponse("/horarios/disponibilidade/painel", status_code=303)
    dados = _ler_formulario(form)
    erros = _validar(dados, exigir_justificativa=False)
    if erros:
        conn.close()
        return _render_form_gestao(request, coleta, alvo, dados, erros)
    _salvar(conn, coleta["id"], professor_id, dados, atualizado_por=prof["id"])
    conn.execute("DELETE FROM horario_coleta_dispensas WHERE coleta_id = ? AND professor_id = ?", (coleta["id"], professor_id))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta['id']}&msg=salva", status_code=303)


# ---------- gestão: período, encerrar, reabrir, nova coleta ----------

def _datas_validas(inicio: str, fim: str) -> bool:
    try:
        return datetime.strptime(inicio, "%Y-%m-%d") <= datetime.strptime(fim, "%Y-%m-%d")
    except (ValueError, TypeError):
        return False


@router.post("/horarios/coletas/nova")
async def gestao_nova_coleta(request: Request):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão cria coletas.")
    form = await request.form()
    inicio, fim = (form.get("data_inicio") or "").strip(), (form.get("data_fim") or "").strip()
    try:
        ano = int(form.get("ano_letivo") or 0)
    except ValueError:
        ano = 0
    if not _datas_validas(inicio, fim) or ano < 2000:
        return RedirectResponse("/horarios/disponibilidade/painel?msg=erro_datas", status_code=303)
    conn = _db()
    cur = conn.execute("INSERT OR IGNORE INTO horario_coletas (ano_letivo, data_inicio, data_fim) VALUES (?, ?, ?)", (ano, inicio, fim))
    conn.commit()
    criou = cur.rowcount > 0
    nova = conn.execute("SELECT id FROM horario_coletas WHERE ano_letivo = ?", (ano,)).fetchone()
    conn.close()
    if not criou:
        return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={nova['id']}&msg=erro_ano", status_code=303)
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={nova['id']}&msg=criada", status_code=303)


@router.post("/horarios/coletas/{coleta_id}/periodo")
async def gestao_periodo(request: Request, coleta_id: int):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão altera o período da coleta.")
    form = await request.form()
    inicio, fim = (form.get("data_inicio") or "").strip(), (form.get("data_fim") or "").strip()
    if not _datas_validas(inicio, fim):
        return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=erro_datas", status_code=303)
    conn = _db()
    conn.execute("UPDATE horario_coletas SET data_inicio = ?, data_fim = ? WHERE id = ?", (inicio, fim, coleta_id))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=periodo", status_code=303)


@router.post("/horarios/coletas/{coleta_id}/encerrar")
def gestao_encerrar(request: Request, coleta_id: int):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão encerra a coleta.")
    conn = _db()
    conn.execute("UPDATE horario_coletas SET encerrada = 1 WHERE id = ?", (coleta_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=encerrada", status_code=303)


@router.post("/horarios/coletas/{coleta_id}/reabrir")
def gestao_reabrir(request: Request, coleta_id: int):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão reabre a coleta.")
    conn = _db()
    conn.execute("UPDATE horario_coletas SET encerrada = 0 WHERE id = ?", (coleta_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/disponibilidade/painel?coleta_id={coleta_id}&msg=reaberta", status_code=303)


# ---------- exportação Excel ----------

@router.get("/horarios/disponibilidade/exportar")
def exportar_disponibilidade(request: Request, coleta_id: str = ""):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return _acesso_negado(request, "Apenas a gestão exporta as respostas.")
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    conn = _db()
    coleta = _coleta_por_id(conn, coleta_id) or _coleta_atual(conn)
    if not coleta:
        conn.close()
        return RedirectResponse("/horarios/disponibilidade/painel", status_code=303)
    dp = _dados_painel(conn, coleta)
    conn.close()

    wb = Workbook()
    negrito = Font(bold=True, color="FFFFFF")
    fundo = PatternFill("solid", fgColor="2563EB")

    ws = wb.active
    ws.title = "Respostas"
    # Mesmas colunas da planilha do Google Forms, mais algumas úteis.
    cab = ["Última atualização", "Nome", "E-mail", "Situação funcional", "Disciplina",
           "Possui interesse em dupla regência?", "Quantos tempos?",
           "Disponibilidade [Manhã]", "Disponibilidade [Tarde]", "Total de turnos",
           "Fora das diretrizes", "Observações gerais", "Preenchido por"]
    ws.append(cab)
    for r in dp["respostas"]:
        por = "" if (not r["atualizado_por"] or r["atualizado_por"] == r["professor_id"]) else (r["atualizado_por_nome"] or "gestão")
        ws.append([
            _ts_br(r["atualizado_em"]), r["professor_nome"], r["professor_email"], r["situacao"], r["disciplina"],
            "Sim" if r["dupla_interesse"] else "Não", r["dupla_tempos"],
            ", ".join(_dias_do_turno(r, "m")), ", ".join(_dias_do_turno(r, "t")), len(_turnos_marcados(r)),
            " | ".join(_avaliar_diretrizes(dict(r))), r["observacoes"] or "", por,
        ])
    larguras = [18, 34, 30, 18, 18, 14, 10, 34, 34, 10, 50, 60, 22]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for c in ws[1]:
        c.font, c.fill = negrito, fundo
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "C2"

    ws2 = wb.create_sheet("Por turno")
    ws2.append(["Turno"] + [nome for _, nome in DIAS])
    for t, nome_t in TURNOS:
        ws2.append([nome_t] + [dp["contagem"][f"{d}_{t}"] for d, _ in DIAS])
    for c in ws2[1]:
        c.font, c.fill = negrito, fundo

    ws3 = wb.create_sheet("Pendentes")
    ws3.append(["Nome", "E-mail"])
    for p in dp["pendentes"]:
        ws3.append([p["nome"], p["email"]])
    ws3.column_dimensions["A"].width = 36
    ws3.column_dimensions["B"].width = 34
    for c in ws3[1]:
        c.font, c.fill = negrito, fundo

    saida = BytesIO()
    wb.save(saida)
    saida.seek(0)
    nome_arq = f"disponibilidade_horario_{coleta['ano_letivo']}.xlsx"
    return StreamingResponse(
        saida, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nome_arq}"'},
    )


# ==========================================
#  PEÇAS PARA O MAIN.PY (menu e tela inicial)
# ==========================================

def itens_menu(prof: Optional[dict], nav_item) -> str:
    """Seção 'Horário Escolar' do menu lateral."""
    if not _pode_responder(prof):
        return ""
    itens = nav_item("/horarios/disponibilidade", "horarios-disponibilidade", "🗓️", "Minha disponibilidade")
    if _eh_gestao(prof):
        itens += nav_item("/horarios/disponibilidade/painel", "horarios-painel", "📊", "Painel de disponibilidade")
    return '<div class="sidebar-section">Horário Escolar</div>' + itens


def grupo_atalhos(prof: Optional[dict]):
    """Grupo de atalhos para a tela inicial (mesmo formato dos outros grupos)."""
    if not _pode_responder(prof):
        return None
    return ("Horário Escolar", "#0f766e", [
        ("🗓️", "Minha disponibilidade", "/horarios/disponibilidade", True),
        ("📊", "Painel de disponibilidade", "/horarios/disponibilidade/painel", _eh_gestao(prof)),
    ])


def card_tela_inicial(prof: Optional[dict]) -> str:
    """Faixa no topo da tela inicial enquanto a coleta está aberta. Nunca deve derrubar
    a tela inicial: qualquer erro aqui só faz a faixa não aparecer."""
    try:
        if not _pode_responder(prof):
            return ""
        conn = _db()
        coleta = _coleta_atual(conn)
        status = _status_coleta(coleta)
        gestao = _eh_gestao(prof)
        if status not in ("aberta", "agendada") or (status == "agendada" and not gestao):
            conn.close()
            return ""
        estilo = ("display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between; "
                  "padding:14px 18px; border-radius:10px; margin-bottom:16px; color:white; "
                  "background:linear-gradient(135deg,#0f766e,#0891b2);")
        if gestao:
            dp = _dados_painel(conn, coleta)
            conn.close()
            resp, pend = len(dp["respostas"]), len(dp["pendentes"])
            quando = (f'abre em {_data_br_curta(coleta["data_inicio"])}' if status == "agendada"
                      else f'prazo até {_data_br_curta(coleta["data_fim"])}')
            return (f'<div style="{estilo}"><div><div style="font-weight:700; font-size:15px;">🗓️ Disponibilidade de horário {coleta["ano_letivo"]}</div>'
                    f'<div style="font-size:13px; opacity:.9;">{resp} responderam · {pend} pendente(s) · {quando}</div></div>'
                    '<a href="/horarios/disponibilidade/painel" class="btn" style="background:white; color:#0f766e; border:none; font-weight:700;">Abrir painel →</a></div>')
        ja = conn.execute("SELECT atualizado_em FROM horario_disponibilidade WHERE coleta_id = ? AND professor_id = ?",
                          (coleta["id"], prof["id"])).fetchone()
        conn.close()
        if ja:
            return (f'<div style="{estilo} background:linear-gradient(135deg,#15803d,#16a34a);"><div>'
                    f'<div style="font-weight:700; font-size:15px;">✅ Disponibilidade de horário {coleta["ano_letivo"]} enviada</div>'
                    f'<div style="font-size:13px; opacity:.9;">Você pode alterar até {_data_br_curta(coleta["data_fim"])}.</div></div>'
                    '<a href="/horarios/disponibilidade" class="btn" style="background:white; color:#15803d; border:none; font-weight:700;">Revisar</a></div>')
        return (f'<div style="{estilo}"><div><div style="font-weight:700; font-size:15px;">🗓️ Informe sua disponibilidade de horário {coleta["ano_letivo"]}</div>'
                f'<div style="font-size:13px; opacity:.9;">Prazo até {_data_br(coleta["data_fim"])} · leva 2 minutos</div></div>'
                '<a href="/horarios/disponibilidade" class="btn" style="background:white; color:#0f766e; border:none; font-weight:700;">Responder agora →</a></div>')
    except Exception as erro:  # pragma: no cover - proteção da tela inicial
        print(f"[horarios] card da tela inicial falhou: {erro}")
        return ""