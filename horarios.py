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
  - /horarios/configuracao ............... (2ª entrega) tempos, componentes e matriz,
                                            turmas e turnos, quem dá cada aula, professor
                                            de Educação Digital por turno e relatório de
                                            viabilidade (confere se a conta fecha).
  - /horarios/geracao .................... (Fase 2) gerador automático com OR-Tools, versões,
                                            visualização por turma/professor/Laboratório,
                                            impressão e Excel. Requer: pip install ortools

Lições do sistema aplicadas aqui (ver histórico do main.py):
  - Gravação sempre com INSERT ... ON CONFLICT DO UPDATE (nunca "buscar-ou-inserir").
  - Nenhuma função de "mesclar" rodando a cada reinício.
  - Exclusões feitas explicitamente no código (o SQLite desta app não faz CASCADE).
  - Rotas específicas declaradas antes das genéricas.
"""

from __future__ import annotations

import html
import time as _time
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
    _marcar_geracoes_interrompidas()
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

        -- ===== Configuração do horário (Fase 1, 2ª entrega) — tudo por ano letivo =====
        -- Os 3 blocos geminados de cada turno (1º-2º, 3º-4º, 5º-6º tempos) e seus horários
        -- de relógio (só para impressão; o gerador trabalha com bloco 1, 2 e 3).
        CREATE TABLE IF NOT EXISTS horario_blocos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            turno TEXT NOT NULL,
            bloco INTEGER NOT NULL,
            inicio TEXT,
            fim TEXT,
            UNIQUE(ano_letivo, turno, bloco)
        );
        -- Componentes do horário e a matriz curricular (tempos por semana). O horário usa
        -- componentes que o boletim não usa (Produção de Texto, Álgebra, Geometria), por
        -- isso cada um aponta para a disciplina do boletim e para a disciplina da coleta.
        CREATE TABLE IF NOT EXISTS horario_componentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            nome TEXT NOT NULL,
            tempos_semana INTEGER NOT NULL DEFAULT 2,
            disciplina_boletim TEXT,
            area_coleta TEXT,
            ed_anos TEXT NOT NULL DEFAULT '',
            ordem INTEGER NOT NULL DEFAULT 0,
            UNIQUE(ano_letivo, nome)
        );
        CREATE TABLE IF NOT EXISTS horario_turmas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            nome TEXT NOT NULL,
            ano_escolaridade INTEGER NOT NULL,
            turno TEXT NOT NULL,
            UNIQUE(ano_letivo, nome)
        );
        -- Quem dá cada componente em cada turma.
        CREATE TABLE IF NOT EXISTS horario_atribuicoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            turma_id INTEGER NOT NULL,
            componente_id INTEGER NOT NULL,
            professor_id INTEGER NOT NULL,
            atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(turma_id, componente_id)
        );
        -- Um professor de Educação Digital por turno.
        CREATE TABLE IF NOT EXISTS horario_ed_professores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            turno TEXT NOT NULL,
            professor_id INTEGER NOT NULL,
            UNIQUE(ano_letivo, turno)
        );
        -- Cada horário gerado vira uma VERSÃO; as aulas guardam também os nomes (fotografia do
        -- momento da geração), para a versão continuar legível se algo mudar depois.
        CREATE TABLE IF NOT EXISTS horario_versoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            criada_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            criada_por INTEGER,
            status TEXT NOT NULL DEFAULT 'gerando',
            modo TEXT,
            resultado TEXT,
            limite_segundos INTEGER,
            tempo REAL,
            indicadores TEXT,
            pares_ed TEXT,
            fora TEXT,
            ed_anos TEXT,
            mensagem TEXT,
            sugestoes TEXT,
            publicada INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS horario_aulas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            versao_id INTEGER NOT NULL,
            turma_id INTEGER,
            componente_id INTEGER,
            professor_id INTEGER,
            turno TEXT NOT NULL,
            dia INTEGER NOT NULL,
            bloco INTEGER NOT NULL,
            turma_nome TEXT,
            componente_nome TEXT,
            professor_nome TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_horario_aulas_versao ON horario_aulas(versao_id);
        -- Opções gerais do horário por ano (ex.: se o planejamento ocupa blocos no horário).
        CREATE TABLE IF NOT EXISTS horario_parametros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ano_letivo INTEGER NOT NULL,
            chave TEXT NOT NULL,
            valor TEXT,
            UNIQUE(ano_letivo, chave)
        );
    """)
    # Colunas acrescentadas depois da primeira versão das tabelas (só adiciona se faltar).
    colunas_versoes = {r["name"] for r in conn.execute("PRAGMA table_info(horario_versoes)").fetchall()}
    if "sugestoes" not in colunas_versoes:
        conn.execute("ALTER TABLE horario_versoes ADD COLUMN sugestoes TEXT")
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
                    <a href="/horarios/configuracao?ano={coleta["ano_letivo"]}" class="btn">⚙️ Configuração do horário</a>
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
#  CONFIGURAÇÃO DO HORÁRIO (Fase 1, 2ª entrega — 01/11/2026)
# ==========================================
# Tudo o que o gerador vai precisar, por ano letivo:
#   tempos (3 blocos geminados por turno) · componentes e matriz · turmas e turnos ·
#   quem dá cada componente em cada turma · professor de Educação Digital por turno.
# E o relatório de viabilidade, que confere se a conta fecha ANTES de gerar.

NOME_TURNO = {"m": "Manhã", "t": "Tarde"}
BLOCOS_POR_TURNO = 3          # 6 tempos por turno = 3 blocos geminados
BLOCOS_SEMANA = 15            # 5 dias × 3 blocos
BLOCOS_PLANEJAMENTO = 2       # 4 tempos de planejamento = 2 blocos
DIAS_PLANEJAMENTO = ("ter", "qua", "qui")

# Início de cada bloco, conforme o horário de 2026 do Urânia (o fim fica em branco
# até a gestão preencher — só aparece na impressão).
BLOCOS_PADRAO = {("m", 1): "07:00", ("m", 2): "08:30", ("m", 3): "10:20",
                 ("t", 1): "13:00", ("t", 2): "14:30", ("t", 3): "16:20"}

# Matriz curricular da escola (Felipe, 01/11/2026). Colunas: nome, tempos por semana,
# disciplina no boletim, disciplina na coleta de disponibilidade, anos com Educação Digital.
MATRIZ_PADRAO = [
    ("Língua Portuguesa", 4, "Português", "Português", "6,7"),
    ("Produção de Texto", 2, "Português", "Português", ""),
    ("Álgebra", 4, "Matemática", "Matemática", "8,9"),
    ("Geometria", 2, "Matemática", "Matemática", ""),
    ("História", 4, "História", "História", ""),
    ("Geografia", 4, "Geografia", "Geografia", ""),
    ("Ciências", 4, "Ciências", "Ciências", ""),
    ("Educação Física", 2, "Ed. Física", "Educação Física", ""),
    ("Inglês", 2, "Inglês", "Inglês", ""),
    ("Arte", 2, "Arte", "Arte", ""),
]

# As três formas de dividir os anos em dois pares para a Educação Digital quinzenal.
PARES_ED = [((6, 7), (8, 9)), ((6, 8), (7, 9)), ((6, 9), (7, 8))]


def _ano_config(conn, ano: str = "") -> int:
    """Ano letivo das telas de configuração: o da URL, senão o da coleta mais recente."""
    if str(ano).strip().isdigit():
        return int(ano)
    coleta = _coleta_atual(conn)
    return coleta["ano_letivo"] if coleta else _hoje().year + 1


def _garantir_config(conn, ano: int):
    """Na primeira vez que a gestão abre a configuração de um ano, cria os blocos e a
    matriz padrão. Só cria o que ainda não existe — nunca altera nem mescla nada."""
    tem_blocos = conn.execute("SELECT COUNT(*) AS c FROM horario_blocos WHERE ano_letivo = ?", (ano,)).fetchone()["c"]
    if not tem_blocos:
        for (turno, bloco), inicio in BLOCOS_PADRAO.items():
            conn.execute("INSERT OR IGNORE INTO horario_blocos (ano_letivo, turno, bloco, inicio, fim) VALUES (?, ?, ?, ?, '')",
                         (ano, turno, bloco, inicio))
    tem_comp = conn.execute("SELECT COUNT(*) AS c FROM horario_componentes WHERE ano_letivo = ?", (ano,)).fetchone()["c"]
    if not tem_comp:
        for ordem, (nome, tempos, disc_bol, area, ed) in enumerate(MATRIZ_PADRAO, start=1):
            conn.execute("""INSERT OR IGNORE INTO horario_componentes
                            (ano_letivo, nome, tempos_semana, disciplina_boletim, area_coleta, ed_anos, ordem)
                            VALUES (?, ?, ?, ?, ?, ?, ?)""", (ano, nome, tempos, disc_bol, area, ed, ordem))
    conn.commit()


# Planejamento fora do quadro de aulas: a escola cadastra à mão, quase sempre no
# contraturno (Felipe, 01/11/2026). A opção continua na aba Tempos, desmarcada.
PARAMETROS_PADRAO = {"planejamento_no_horario": "0"}


def _parametro(conn, ano: int, chave: str) -> str:
    r = conn.execute("SELECT valor FROM horario_parametros WHERE ano_letivo = ? AND chave = ?", (ano, chave)).fetchone()
    return r["valor"] if r and r["valor"] is not None else PARAMETROS_PADRAO.get(chave, "")


def _definir_parametro(conn, ano: int, chave: str, valor: str):
    conn.execute("""INSERT INTO horario_parametros (ano_letivo, chave, valor) VALUES (?, ?, ?)
                    ON CONFLICT(ano_letivo, chave) DO UPDATE SET valor = excluded.valor""", (ano, chave, valor))


def _componentes(conn, ano: int) -> list:
    return conn.execute("SELECT * FROM horario_componentes WHERE ano_letivo = ? ORDER BY ordem, id", (ano,)).fetchall()


def _turmas(conn, ano: int) -> list:
    return conn.execute("SELECT * FROM horario_turmas WHERE ano_letivo = ? ORDER BY turno, nome", (ano,)).fetchall()


def _atribuicoes(conn, ano: int) -> dict:
    linhas = conn.execute("""
        SELECT a.turma_id, a.componente_id, a.professor_id FROM horario_atribuicoes a
        JOIN horario_turmas t ON t.id = a.turma_id WHERE t.ano_letivo = ?""", (ano,)).fetchall()
    return {(r["turma_id"], r["componente_id"]): r["professor_id"] for r in linhas}


def _ed_professores(conn, ano: int) -> dict:
    linhas = conn.execute("SELECT turno, professor_id FROM horario_ed_professores WHERE ano_letivo = ?", (ano,)).fetchall()
    return {r["turno"]: r["professor_id"] for r in linhas}


def _coleta_do_ano(conn, ano: int):
    return conn.execute("SELECT * FROM horario_coletas WHERE ano_letivo = ?", (ano,)).fetchone()


def _disponibilidades(conn, ano: int) -> dict:
    coleta = _coleta_do_ano(conn, ano)
    if not coleta:
        return {}
    linhas = conn.execute("SELECT * FROM horario_disponibilidade WHERE coleta_id = ?", (coleta["id"],)).fetchall()
    return {r["professor_id"]: r for r in linhas}


def _professores_opcoes(conn, ano: int) -> list:
    """Professores que podem receber aulas: docentes ativos, quem respondeu a coleta do
    ano e quem já tem aula atribuída. Cada um com a disciplina informada na coleta."""
    disp = _disponibilidades(conn, ano)
    ids_extra = set(disp.keys()) | set(_atribuicoes(conn, ano).values()) | set(_ed_professores(conn, ano).values())
    linhas = conn.execute("""SELECT id, nome FROM professores
                             WHERE (COALESCE(status, 'ativo') = 'ativo' AND papel = 'docente')""").fetchall()
    profs = {r["id"]: r["nome"] for r in linhas}
    if ids_extra:
        marcadores = ",".join("?" for _ in ids_extra)
        for r in conn.execute(f"SELECT id, nome FROM professores WHERE id IN ({marcadores})", list(ids_extra)).fetchall():
            profs[r["id"]] = r["nome"]
    saida = [{"id": pid, "nome": nome, "area": (disp[pid]["disciplina"] if pid in disp else "")} for pid, nome in profs.items()]
    return sorted(saida, key=lambda p: (p["nome"] or "").lower())


def _pl(n: int, singular: str, plural: str) -> str:
    return f"{n} {singular if n == 1 else plural}"


def _nomes_curtos(profs: list) -> dict:
    """Nome + último sobrenome para caber na grade (ex.: "Eriwelton Lopes Ferraz" → "Eriwelton Ferraz").
    Se dois professores ficariam com o mesmo nome curto, os dois aparecem com o nome completo."""
    curto = {}
    for p in profs:
        partes = (p["nome"] or "").split()
        curto[p["id"]] = f"{partes[0]} {partes[-1]}" if len(partes) > 2 else (p["nome"] or "")
    contagem = {}
    for v in curto.values():
        contagem[v] = contagem.get(v, 0) + 1
    return {pid: (v if contagem[v] == 1 else next(p["nome"] for p in profs if p["id"] == pid)) for pid, v in curto.items()}


def _anos_ed(componente) -> set:
    return {int(x) for x in (componente["ed_anos"] or "").split(",") if x.strip().isdigit()}


def _blocos_comp(componente) -> int:
    return int(componente["tempos_semana"] or 0) // 2


def _abas_config(ativa: str, ano: int) -> str:
    abas = [
        ("resumo", "📋 Resumo", "/horarios/configuracao"),
        ("tempos", "⏰ Tempos", "/horarios/configuracao/tempos"),
        ("componentes", "📚 Componentes e matriz", "/horarios/configuracao/componentes"),
        ("turmas", "👥 Turmas", "/horarios/configuracao/turmas"),
        ("atribuicoes", "🔗 Quem dá cada aula", "/horarios/configuracao/atribuicoes"),
        ("viabilidade", "🔎 Viabilidade", "/horarios/configuracao/viabilidade"),
        ("geracao", "🚀 Gerar horário", "/horarios/geracao"),
    ]
    links = ""
    for chave, rotulo, href in abas:
        estilo = ("background:#0f766e; color:white; border-color:#0f766e;" if chave == ativa else "")
        links += f'<a href="{href}?ano={ano}" class="btn" style="margin:0 6px 6px 0; {estilo}">{rotulo}</a>'
    return f'<div style="display:flex; flex-wrap:wrap; margin-bottom:16px;">{links}</div>'


def _cabecalho_config(titulo: str, subtitulo: str, ativa: str, ano: int, msg_html: str = "") -> str:
    return f"""
        <div class="page-header">
            <h1>{titulo} <span style="font-size:15px; color:var(--text-muted); font-weight:500;">· Horário {ano}</span></h1>
            <p class="subtitle">{subtitulo}</p>
        </div>
        {_abas_config(ativa, ano)}
        {msg_html}"""


def _msg_box(texto: str, erro: bool = False) -> str:
    if not texto:
        return ""
    cor = ("background:var(--red-bg,#fef2f2); color:var(--red,#b91c1c); border:1px solid var(--red,#b91c1c);" if erro
           else "background:#ecfdf5; color:#047857; border:1px solid #10b981;")
    return f'<div style="{cor} padding:10px 14px; border-radius:6px; margin-bottom:14px; font-size:13px;">{texto}</div>'


def _exige_gestao(request: Request):
    prof = _ctx["get_prof"](request)
    if not _eh_gestao(prof):
        return None, _acesso_negado(request, "Apenas a gestão acessa a configuração do horário.")
    return prof, None


# ---------- Resumo (página central) ----------

@router.get("/horarios/configuracao", response_class=HTMLResponse)
def config_resumo(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    comps = _componentes(conn, ano_i)
    turmas = _turmas(conn, ano_i)
    atrib = _atribuicoes(conn, ano_i)
    ed = _ed_professores(conn, ano_i)
    disp = _disponibilidades(conn, ano_i)
    analise = _analisar_viabilidade(conn, ano_i)
    conn.close()

    soma_tempos = sum(int(c["tempos_semana"] or 0) for c in comps)
    celulas = len(turmas) * len(comps)
    preenchidas = sum(1 for t in turmas for c in comps if (t["id"], c["id"]) in atrib)
    n_m = sum(1 for t in turmas if t["turno"] == "m")
    n_t = len(turmas) - n_m

    def item(ok, titulo, detalhe, href):
        icone = "✅" if ok is True else ("❌" if ok == "erro" else ("⚠️" if ok is None else "⬜"))
        return (f'<a href="{href}?ano={ano_i}" class="card" style="display:flex; gap:14px; align-items:center; padding:14px 16px; '
                f'margin-bottom:10px; text-decoration:none; color:inherit;"><span style="font-size:22px;">{icone}</span>'
                f'<span><strong>{titulo}</strong><br><span style="font-size:13px; color:var(--text-muted);">{detalhe}</span></span></a>')

    problemas = analise["n_erros"]
    lista = (
        item(len(disp) > 0 or None, "1. Disponibilidade dos professores", f"{len(disp)} resposta(s) na coleta de {ano_i}.", "/horarios/disponibilidade/painel")
        + item(True, "2. Tempos", "3 blocos geminados por turno (1º-2º, 3º-4º, 5º-6º tempos).", "/horarios/configuracao/tempos")
        + item(soma_tempos == BLOCOS_SEMANA * 2 or None, "3. Componentes e matriz",
               f"{len(comps)} componentes, {soma_tempos} tempos por semana" + ("" if soma_tempos == 30 else " — o esperado é 30"), "/horarios/configuracao/componentes")
        + item(bool(turmas) or False, "4. Turmas", f"{len(turmas)} turma(s): {n_m} de manhã, {n_t} à tarde." if turmas else "Nenhuma turma cadastrada ainda.", "/horarios/configuracao/turmas")
        + item((celulas > 0 and preenchidas == celulas and len(ed) == len({t["turno"] for t in turmas})) or (None if preenchidas else False),
               "5. Quem dá cada aula", f"{preenchidas} de {celulas} aulas com professor · Educação Digital: "
               + (", ".join(f"{NOME_TURNO[k]} ✓" for k in sorted(ed)) or "nenhum professor definido"), "/horarios/configuracao/atribuicoes")
        + item((True if problemas == 0 else "erro") if turmas else False, "6. Viabilidade",
               ("Nenhum impedimento encontrado." if problemas == 0 else f"{problemas} impedimento(s) para resolver antes de gerar.") if turmas else "Aparece depois de cadastrar turmas e aulas.",
               "/horarios/configuracao/viabilidade")
    )
    conteudo = _cabecalho_config("⚙️ Configuração do horário", "Preencha na ordem. Quando tudo estiver ✅, o horário poderá ser gerado.", "resumo", ano_i) + f"""
        {lista}
        <a href="/horarios/geracao?ano={ano_i}" class="card" style="display:flex; gap:14px; align-items:center; padding:14px 16px; margin-top:6px;
           text-decoration:none; color:white; background:linear-gradient(135deg,#0f766e,#0891b2); border:none;">
            <span style="font-size:22px;">🚀</span>
            <span><strong>7. Gerar o horário</strong><br><span style="font-size:13px; opacity:.9;">Monta o horário automaticamente e mostra o que impede, se não fechar.</span></span>
        </a>"""
    return _pagina(request, "Configuração do horário", conteudo, "horarios-config")


# ---------- Tempos ----------

@router.get("/horarios/configuracao/tempos", response_class=HTMLResponse)
def config_tempos(request: Request, ano: str = "", msg: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    blocos = {(r["turno"], r["bloco"]): r for r in conn.execute("SELECT * FROM horario_blocos WHERE ano_letivo = ?", (ano_i,)).fetchall()}
    plan_no_horario = _parametro(conn, ano_i, "planejamento_no_horario") == "1"
    conn.close()
    linhas = ""
    for turno in ("m", "t"):
        for b in range(1, BLOCOS_POR_TURNO + 1):
            r = blocos.get((turno, b))
            ini = _e(r["inicio"]) if r else ""
            fim = _e(r["fim"]) if r else ""
            linhas += f"""<tr style="border-top:1px solid var(--border);">
                <td style="padding:8px; font-weight:600;">{NOME_TURNO[turno]}</td>
                <td style="padding:8px;">{b}º bloco <span style="color:var(--text-muted); font-size:12px;">({2*b-1}º e {2*b}º tempos)</span></td>
                <td style="padding:8px;"><input type="time" name="ini_{turno}_{b}" value="{ini}" style="width:130px;"></td>
                <td style="padding:8px;"><input type="time" name="fim_{turno}_{b}" value="{fim}" style="width:130px;"></td>
            </tr>"""
    conteudo = _cabecalho_config("⏰ Tempos", "Cada turno tem 6 tempos, sempre em pares geminados: 3 blocos por dia, 15 por semana. Os horários de relógio só aparecem na impressão.",
                                 "tempos", ano_i, _msg_box("Tempos salvos." if msg == "ok" else "")) + f"""
        <form method="post" action="/horarios/configuracao/tempos?ano={ano_i}" class="card" style="padding:14px 16px;">
            <table style="border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);"><th style="padding:8px; text-align:left;">Turno</th><th style="padding:8px; text-align:left;">Bloco</th>
                <th style="padding:8px; text-align:left;">Início</th><th style="padding:8px; text-align:left;">Fim</th></tr></thead>
                <tbody>{linhas}</tbody>
            </table>
            <div style="margin-top:16px; padding-top:14px; border-top:1px solid var(--border);">
                <label style="display:flex; gap:8px; align-items:flex-start; font-size:13px; cursor:pointer;">
                    <input type="checkbox" name="planejamento_no_horario" value="1"{" checked" if plan_no_horario else ""} style="width:auto; margin-top:3px;">
                    <span><strong>O planejamento ocupa lugar no horário.</strong> Reservar, para cada professor, 4 tempos de planejamento (2 blocos)
                    dentro dos turnos que ele marcou, em terça, quarta ou quinta.<br>
                    <span style="color:var(--text-muted);">Desmarque se o planejamento é cumprido fora dos turnos de aula informados na coleta.</span></span>
                </label>
            </div>
            <div class="page-actions" style="margin-top:12px;"><button type="submit" class="btn btn-primary">💾 Salvar</button></div>
        </form>"""
    return _pagina(request, "Tempos do horário", conteudo, "horarios-config")


@router.post("/horarios/configuracao/tempos")
async def config_tempos_salvar(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    conn = _db()
    ano_i = _ano_config(conn, ano)
    for turno in ("m", "t"):
        for b in range(1, BLOCOS_POR_TURNO + 1):
            ini = (form.get(f"ini_{turno}_{b}") or "").strip()[:5]
            fim = (form.get(f"fim_{turno}_{b}") or "").strip()[:5]
            conn.execute("""INSERT INTO horario_blocos (ano_letivo, turno, bloco, inicio, fim) VALUES (?, ?, ?, ?, ?)
                            ON CONFLICT(ano_letivo, turno, bloco) DO UPDATE SET inicio = excluded.inicio, fim = excluded.fim""",
                         (ano_i, turno, b, ini, fim))
    _definir_parametro(conn, ano_i, "planejamento_no_horario", "1" if form.get("planejamento_no_horario") else "0")
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/configuracao/tempos?ano={ano_i}&msg=ok", status_code=303)


# ---------- Componentes e matriz ----------

@router.get("/horarios/configuracao/componentes", response_class=HTMLResponse)
def config_componentes(request: Request, ano: str = "", msg: str = "", erro: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    comps = _componentes(conn, ano_i)
    conn.close()
    return _render_componentes(request, ano_i, comps, msg, erro)


def _render_componentes(request, ano_i, comps, msg, erro):
    soma = sum(int(c["tempos_semana"] or 0) for c in comps)
    opts_area = lambda atual: "".join(f'<option{" selected" if a == atual else ""}>{a}</option>' for a in DISCIPLINAS if a != "Educação Digital")
    linhas = ""
    for c in comps:
        anos_ed = _anos_ed(c)
        tempos_opts = "".join(f'<option value="{t}"{" selected" if t == int(c["tempos_semana"] or 0) else ""}>{t}</option>' for t in (2, 4, 6, 8))
        caixas_ed = "".join(
            f'<label style="display:inline-flex; gap:3px; align-items:center; margin-right:8px; font-weight:400;">'
            f'<input type="checkbox" name="ed_{c["id"]}_{a}" value="1"{" checked" if a in anos_ed else ""} style="width:auto; margin:0;">{a}º</label>'
            for a in (6, 7, 8, 9))
        linhas += f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:6px;"><input type="number" name="ordem_{c["id"]}" value="{c["ordem"]}" style="width:60px;"></td>
            <td style="padding:6px;"><input type="text" name="nome_{c["id"]}" value="{_e(c["nome"])}" required style="min-width:160px;"></td>
            <td style="padding:6px;"><select name="tempos_{c["id"]}" style="width:80px;">{tempos_opts}</select></td>
            <td style="padding:6px;"><select name="area_{c["id"]}">{opts_area(c["area_coleta"])}</select></td>
            <td style="padding:6px;"><input type="text" name="bol_{c["id"]}" value="{_e(c["disciplina_boletim"])}" style="width:130px;"></td>
            <td style="padding:6px; white-space:nowrap;">{caixas_ed}</td>
            <td style="padding:6px; text-align:center;"><input type="checkbox" name="excluir_{c["id"]}" value="1" style="width:auto;" title="Excluir este componente"></td>
        </tr>"""
    aviso_soma = ("" if soma == 30 else
                  f'<div class="hd-avisos" style="display:block;">⚠️ A matriz soma {soma} tempos por semana. Cada turma tem 30 tempos (15 blocos) — confira.</div>')
    conteudo = _cabecalho_config(
        "📚 Componentes e matriz", "Componentes que aparecem no horário e quantos tempos semanais cada um tem (vale para 6º a 9º ano).",
        "componentes", ano_i, _msg_box("Matriz salva." if msg == "ok" else "") + _msg_box(_e(erro), erro=True)) + f"""
        {_CSS_FORM}
        <form method="post" action="/horarios/configuracao/componentes?ano={ano_i}" class="card" style="padding:14px 16px;">
            <div style="font-size:15px; margin-bottom:8px;">Total: <strong>{soma} tempos</strong> por semana ({soma // 2} blocos geminados)</div>
            {aviso_soma}
            <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);">
                    <th style="padding:6px; text-align:left;">Ordem</th><th style="padding:6px; text-align:left;">Componente</th>
                    <th style="padding:6px; text-align:left;">Tempos/semana</th>
                    <th style="padding:6px; text-align:left;" title="Usada para sugerir professores pela resposta da coleta">Disciplina na coleta</th>
                    <th style="padding:6px; text-align:left;">Disciplina no boletim</th>
                    <th style="padding:6px; text-align:left;" title="Anos em que este componente acontece junto com a Educação Digital (turma dividida, metade no Laboratório)">Educação Digital junto, nos anos</th>
                    <th style="padding:6px;">Excluir</th>
                </tr></thead>
                <tbody>{linhas}
                <tr style="border-top:2px solid var(--border); background:var(--bg-subtle);">
                    <td style="padding:6px; font-size:12px; color:var(--text-muted);">Novo:</td>
                    <td style="padding:6px;"><input type="text" name="novo_nome" placeholder="(opcional) nome do novo componente"></td>
                    <td style="padding:6px;"><select name="novo_tempos" style="width:80px;"><option>2</option><option>4</option><option>6</option></select></td>
                    <td style="padding:6px;"><select name="novo_area">{opts_area("")}</select></td>
                    <td style="padding:6px;"><input type="text" name="novo_bol" style="width:130px;"></td>
                    <td colspan="2"></td>
                </tr>
                </tbody>
            </table>
            </div>
            <p style="font-size:12px; color:var(--text-muted); margin:10px 0 0;">Tempos sempre em número par (aulas geminadas). Excluir um componente apaga também quem estava atribuído a ele nas turmas.</p>
            <div class="page-actions" style="margin-top:12px;"><button type="submit" class="btn btn-primary">💾 Salvar matriz</button></div>
        </form>"""
    return _pagina(request, "Componentes e matriz", conteudo, "horarios-config")


@router.post("/horarios/configuracao/componentes")
async def config_componentes_salvar(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    conn = _db()
    ano_i = _ano_config(conn, ano)
    comps = _componentes(conn, ano_i)
    nomes_vistos = set()
    erros = []
    mudancas = []
    for c in comps:
        cid = c["id"]
        if form.get(f"excluir_{cid}"):
            mudancas.append(("excluir", cid, None))
            continue
        nome = (form.get(f"nome_{cid}") or "").strip()
        try:
            tempos = int(form.get(f"tempos_{cid}") or 0)
            ordem = int(form.get(f"ordem_{cid}") or 0)
        except ValueError:
            tempos, ordem = 0, 0
        if not nome:
            erros.append("Todo componente precisa de nome.")
        elif nome.lower() in nomes_vistos:
            erros.append(f"Nome repetido: {nome}.")
        nomes_vistos.add(nome.lower())
        if tempos not in (2, 4, 6, 8):
            erros.append(f"{nome}: tempos por semana precisa ser par (2, 4, 6 ou 8).")
        anos_ed = ",".join(str(a) for a in (6, 7, 8, 9) if form.get(f"ed_{cid}_{a}"))
        mudancas.append(("atualizar", cid, (nome, tempos, (form.get(f"area_{cid}") or "").strip(),
                                           (form.get(f"bol_{cid}") or "").strip(), anos_ed, ordem)))
    novo_nome = (form.get("novo_nome") or "").strip()
    if novo_nome:
        if novo_nome.lower() in nomes_vistos:
            erros.append(f"Já existe um componente chamado {novo_nome}.")
        try:
            novo_tempos = int(form.get("novo_tempos") or 2)
        except ValueError:
            novo_tempos = 2
    if erros:
        conn.close()
        return RedirectResponse(f"/horarios/configuracao/componentes?ano={ano_i}&erro=" + _url(" ".join(dict.fromkeys(erros))), status_code=303)
    for tipo, cid, valores in mudancas:
        if tipo == "excluir":
            # Sem CASCADE nesta app: apaga as atribuições do componente explicitamente.
            conn.execute("DELETE FROM horario_atribuicoes WHERE componente_id = ?", (cid,))
            conn.execute("DELETE FROM horario_componentes WHERE id = ?", (cid,))
        else:
            nome, tempos, area, bol, anos_ed, ordem = valores
            conn.execute("""UPDATE horario_componentes SET nome = ?, tempos_semana = ?, area_coleta = ?, disciplina_boletim = ?,
                            ed_anos = ?, ordem = ? WHERE id = ?""", (nome, tempos, area, bol, anos_ed, ordem, cid))
    if novo_nome:
        prox = conn.execute("SELECT COALESCE(MAX(ordem), 0) + 1 AS o FROM horario_componentes WHERE ano_letivo = ?", (ano_i,)).fetchone()["o"]
        conn.execute("""INSERT OR IGNORE INTO horario_componentes (ano_letivo, nome, tempos_semana, disciplina_boletim, area_coleta, ed_anos, ordem)
                        VALUES (?, ?, ?, ?, ?, '', ?)""",
                     (ano_i, novo_nome, novo_tempos, (form.get("novo_bol") or "").strip(), (form.get("novo_area") or "").strip(), prox))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/configuracao/componentes?ano={ano_i}&msg=ok", status_code=303)


def _url(texto: str) -> str:
    from urllib.parse import quote
    return quote(texto)


# ---------- Turmas ----------

@router.get("/horarios/configuracao/turmas", response_class=HTMLResponse)
def config_turmas(request: Request, ano: str = "", msg: str = "", erro: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    turmas = _turmas(conn, ano_i)
    conn.close()
    linhas = ""
    for t in turmas:
        anos_opts = "".join(f'<option value="{a}"{" selected" if a == t["ano_escolaridade"] else ""}>{a}º ano</option>' for a in (6, 7, 8, 9))
        turno_opts = "".join(f'<option value="{k}"{" selected" if k == t["turno"] else ""}>{v}</option>' for k, v in NOME_TURNO.items())
        linhas += f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:6px 8px; font-weight:700; font-size:14px;">{_e(t["nome"])}</td>
            <td style="padding:6px 8px;"><select name="ano_{t["id"]}" style="width:110px;">{anos_opts}</select></td>
            <td style="padding:6px 8px;"><select name="turno_{t["id"]}" style="width:110px;">{turno_opts}</select></td>
            <td style="padding:6px 8px; text-align:center;"><input type="checkbox" name="excluir_{t["id"]}" value="1" style="width:auto;"></td>
        </tr>"""
    n_m = sum(1 for t in turmas if t["turno"] == "m")
    tabela = ""
    if turmas:
        tabela = f"""
        <form method="post" action="/horarios/configuracao/turmas/salvar?ano={ano_i}" class="card" style="padding:14px 16px; margin-bottom:16px;"
              onsubmit="if(document.querySelector('input[name^=excluir_]:checked')) return confirm('Excluir as turmas marcadas? Quem estava atribuído a elas também é apagado.'); return true;">
            <h2 style="margin:0 0 8px; font-size:15px;">Turmas de {ano_i} — {len(turmas)} ({n_m} de manhã, {len(turmas) - n_m} à tarde)</h2>
            <table style="border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);"><th style="padding:6px 8px; text-align:left;">Turma</th>
                <th style="padding:6px 8px; text-align:left;">Ano</th><th style="padding:6px 8px; text-align:left;">Turno</th><th style="padding:6px 8px;">Excluir</th></tr></thead>
                <tbody>{linhas}</tbody>
            </table>
            <div class="page-actions" style="margin-top:12px;"><button type="submit" class="btn btn-primary">💾 Salvar alterações</button></div>
        </form>"""
    msgs = {"ok": "Turmas salvas.", "add": "Turmas adicionadas."}
    conteudo = _cabecalho_config("👥 Turmas", "Turmas que terão horário neste ano letivo e o turno de cada uma. As turmas mudam de turno a cada ano — confira sempre.",
                                 "turmas", ano_i, _msg_box(_e(msgs.get(msg, msg)) if msg else "") + _msg_box(_e(erro), erro=True)) + f"""
        {tabela}
        <form method="post" action="/horarios/configuracao/turmas/adicionar?ano={ano_i}" class="card" style="padding:14px 16px;">
            <h2 style="margin:0 0 4px; font-size:15px;">➕ Adicionar turmas</h2>
            <p style="margin:0 0 10px; font-size:12px; color:var(--text-muted);">Digite os nomes separados por espaço ou vírgula (ex.: <code>601 602 701 702</code>). O ano sai do primeiro número (601 → 6º ano).</p>
            <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end;">
                <label style="margin:0; flex:1 1 320px;">Turmas<input type="text" name="nomes" required placeholder="601 602 701 702 801 802 901"></label>
                <label style="margin:0;">Turno<select name="turno" style="width:130px;"><option value="m">Manhã</option><option value="t">Tarde</option></select></label>
                <button type="submit" class="btn btn-primary">Adicionar</button>
            </div>
        </form>"""
    return _pagina(request, "Turmas do horário", conteudo, "horarios-config")


@router.post("/horarios/configuracao/turmas/adicionar")
async def config_turmas_adicionar(request: Request, ano: str = ""):
    import re as _re
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    conn = _db()
    ano_i = _ano_config(conn, ano)
    turno = form.get("turno") if form.get("turno") in NOME_TURNO else "m"
    nomes = [n.strip() for n in _re.split(r"[\s,;]+", form.get("nomes") or "") if n.strip()]
    invalidos, repetidos, novos = [], [], 0
    for nome in nomes:
        if not _re.fullmatch(r"[0-9A-Za-z]{1,10}", nome) or nome[0] not in "6789":
            invalidos.append(nome)
            continue
        cur = conn.execute("INSERT OR IGNORE INTO horario_turmas (ano_letivo, nome, ano_escolaridade, turno) VALUES (?, ?, ?, ?)",
                           (ano_i, nome, int(nome[0]), turno))
        if cur.rowcount:
            novos += 1
        else:
            repetidos.append(nome)
    conn.commit()
    conn.close()
    msg = f"{novos} turma(s) adicionada(s)."
    if repetidos:
        msg += f" Já existiam: {', '.join(repetidos)}."
    url = f"/horarios/configuracao/turmas?ano={ano_i}&msg={_url(msg)}"
    if invalidos:
        url += "&erro=" + _url(f"Não reconhecidas (o nome deve começar com 6, 7, 8 ou 9): {', '.join(invalidos)}.")
    return RedirectResponse(url, status_code=303)


@router.post("/horarios/configuracao/turmas/salvar")
async def config_turmas_salvar(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    conn = _db()
    ano_i = _ano_config(conn, ano)
    for t in _turmas(conn, ano_i):
        if form.get(f"excluir_{t['id']}"):
            conn.execute("DELETE FROM horario_atribuicoes WHERE turma_id = ?", (t["id"],))
            conn.execute("DELETE FROM horario_turmas WHERE id = ?", (t["id"],))
            continue
        try:
            ano_esc = int(form.get(f"ano_{t['id']}") or t["ano_escolaridade"])
        except ValueError:
            ano_esc = t["ano_escolaridade"]
        turno = form.get(f"turno_{t['id']}") if form.get(f"turno_{t['id']}") in NOME_TURNO else t["turno"]
        if ano_esc in (6, 7, 8, 9):
            conn.execute("UPDATE horario_turmas SET ano_escolaridade = ?, turno = ? WHERE id = ?", (ano_esc, turno, t["id"]))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/configuracao/turmas?ano={ano_i}&msg=ok", status_code=303)


# ---------- Quem dá cada aula (atribuições) ----------

_JS_ATRIB = """
<script>
(function () {
  var TODOS = %(todos)s;
  document.querySelectorAll('select.hd-prof').forEach(function (s) {
    s.addEventListener('change', function () {
      if (s.value !== '__outro__') return;
      var atual = '';
      s.innerHTML = '<option value="">—</option>' + TODOS.map(function (p) {
        return '<option value="' + p.id + '">' + p.nome + (p.area ? ' (' + p.area + ')' : '') + '</option>';
      }).join('');
      s.value = atual;
      s.focus();
    });
  });
})();
</script>
"""


def _select_prof(nome_campo: str, atual: Optional[int], sugeridos: list, todos_por_id: dict) -> str:
    """Select com os professores da disciplina primeiro; 'outro professor…' abre a lista toda."""
    ids = [p["id"] for p in sugeridos]
    if atual and atual not in ids and atual in todos_por_id:
        sugeridos = [todos_por_id[atual]] + sugeridos
    opts = '<option value="">—</option>' + "".join(
        f'<option value="{p["id"]}"{" selected" if p["id"] == atual else ""} title="{_e(p["nome"])}">{_e(p.get("curto") or p["nome"])}</option>'
        for p in sugeridos)
    opts += '<option value="__outro__">outro professor…</option>'
    cor = "" if atual else "border-color:#f59e0b;"
    return f'<select name="{nome_campo}" class="hd-prof" style="min-width:150px; font-size:12px; padding:4px; {cor}">{opts}</select>'


@router.get("/horarios/configuracao/atribuicoes", response_class=HTMLResponse)
def config_atribuicoes(request: Request, ano: str = "", msg: str = ""):
    import json as _json
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    comps = _componentes(conn, ano_i)
    turmas = _turmas(conn, ano_i)
    atrib = _atribuicoes(conn, ano_i)
    ed = _ed_professores(conn, ano_i)
    profs = _professores_opcoes(conn, ano_i)
    conn.close()
    curtos = _nomes_curtos(profs)
    for p in profs:
        p["curto"] = curtos[p["id"]]
    por_id = {p["id"]: p for p in profs}

    if not turmas:
        conteudo = _cabecalho_config("🔗 Quem dá cada aula", "", "atribuicoes", ano_i) + \
            f'<div class="empty">Cadastre as turmas primeiro. <a href="/horarios/configuracao/turmas?ano={ano_i}">Ir para Turmas →</a></div>'
        return _pagina(request, "Quem dá cada aula", conteudo, "horarios-config")

    def sugeridos(area):
        da_area = [p for p in profs if p["area"] == area]
        return da_area

    # Educação Digital: um professor por turno
    turnos_usados = [k for k in ("m", "t") if any(t["turno"] == k for t in turmas)]
    bloco_ed = "".join(
        f'<label style="margin:0 18px 0 0; display:inline-block;">{NOME_TURNO[k]}<br>{_select_prof(f"ed_{k}", ed.get(k), sugeridos("Educação Digital"), por_id)}</label>'
        for k in turnos_usados)

    tabelas = ""
    for k in turnos_usados:
        cab = "".join(f'<th style="padding:6px; text-align:left; font-size:12px; white-space:nowrap;">{_e(c["nome"])}<br>'
                      f'<span style="font-weight:400; color:var(--text-muted);">{c["tempos_semana"]} tempos'
                      f'{" · com ED " + ",".join(str(a) + "º" for a in sorted(_anos_ed(c))) if _anos_ed(c) else ""}</span></th>' for c in comps)
        linhas = ""
        for t in [t for t in turmas if t["turno"] == k]:
            celulas = "".join(
                '<td style="padding:4px;">' + _select_prof("a_%d_%d" % (t["id"], c["id"]), atrib.get((t["id"], c["id"])),
                                                           sugeridos(c["area_coleta"]), por_id) + '</td>'
                for c in comps)
            linhas += (f'<tr style="border-top:1px solid var(--border);"><td style="padding:6px; font-weight:700; position:sticky; left:0; '
                       f'background:var(--bg, #fff);">{_e(t["nome"])}</td>{celulas}</tr>')
        tabelas += f"""
            <h2 style="margin:18px 0 8px; font-size:15px;">{NOME_TURNO[k]}</h2>
            <div style="overflow-x:auto; border:1px solid var(--border); border-radius:8px;">
            <table style="border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);"><th style="padding:6px; text-align:left; position:sticky; left:0; background:var(--bg-subtle);">Turma</th>{cab}</tr></thead>
                <tbody>{linhas}</tbody>
            </table></div>"""

    todos_json = _json.dumps([{"id": p["id"], "nome": p["nome"], "area": p["area"]} for p in profs], ensure_ascii=False).replace("</", "<\\/")
    vazias = sum(1 for t in turmas for c in comps if (t["id"], c["id"]) not in atrib)
    conteudo = _cabecalho_config(
        "🔗 Quem dá cada aula", "Para cada turma, escolha o professor de cada componente. A lista mostra primeiro quem informou aquela disciplina na coleta; "
        "“outro professor…” abre a lista completa.", "atribuicoes", ano_i,
        _msg_box("Atribuições salvas." if msg == "ok" else "")) + f"""
        <form method="post" action="/horarios/configuracao/atribuicoes?ano={ano_i}">
            <div class="card" style="padding:14px 16px;">
                <h2 style="margin:0 0 4px; font-size:15px;">💻 Professor de Educação Digital por turno</h2>
                <p style="margin:0 0 10px; font-size:12px; color:var(--text-muted);">Acompanha os blocos de Língua Portuguesa (6º/7º) e Álgebra (8º/9º), com a turma dividida e metade no Laboratório — quinzenal.</p>
                {bloco_ed}
            </div>
            <p style="font-size:12px; color:var(--text-muted); margin:12px 0 0;">{vazias} aula(s) ainda sem professor (contorno laranja).</p>
            {tabelas}
            <div class="page-actions" style="margin-top:14px; position:sticky; bottom:0; background:var(--bg, #fff); padding:10px 0;">
                <button type="submit" class="btn btn-primary">💾 Salvar atribuições</button>
                <a href="/horarios/configuracao/viabilidade?ano={ano_i}" class="btn">🔎 Ver viabilidade</a>
            </div>
        </form>
        {_JS_ATRIB % {"todos": todos_json}}"""
    return _pagina(request, "Quem dá cada aula", conteudo, "horarios-config")


@router.post("/horarios/configuracao/atribuicoes")
async def config_atribuicoes_salvar(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    conn = _db()
    ano_i = _ano_config(conn, ano)
    validos = {r["id"] for r in conn.execute("SELECT id FROM professores").fetchall()}

    def pid(valor):
        v = (valor or "").strip()
        return int(v) if v.isdigit() and int(v) in validos else None

    for t in _turmas(conn, ano_i):
        for c in _componentes(conn, ano_i):
            campo = f"a_{t['id']}_{c['id']}"
            if campo not in form:
                continue
            p = pid(form.get(campo))
            if p:
                conn.execute("""INSERT INTO horario_atribuicoes (turma_id, componente_id, professor_id) VALUES (?, ?, ?)
                                ON CONFLICT(turma_id, componente_id) DO UPDATE SET professor_id = excluded.professor_id,
                                atualizado_em = CURRENT_TIMESTAMP""", (t["id"], c["id"], p))
            else:
                conn.execute("DELETE FROM horario_atribuicoes WHERE turma_id = ? AND componente_id = ?", (t["id"], c["id"]))
    for k in NOME_TURNO:
        campo = f"ed_{k}"
        if campo not in form:
            continue
        p = pid(form.get(campo))
        if p:
            conn.execute("""INSERT INTO horario_ed_professores (ano_letivo, turno, professor_id) VALUES (?, ?, ?)
                            ON CONFLICT(ano_letivo, turno) DO UPDATE SET professor_id = excluded.professor_id""", (ano_i, k, p))
        else:
            conn.execute("DELETE FROM horario_ed_professores WHERE ano_letivo = ? AND turno = ?", (ano_i, k))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/configuracao/atribuicoes?ano={ano_i}&msg=ok", status_code=303)


# ---------- Viabilidade ----------

def _analisar_viabilidade(conn, ano: int) -> dict:
    """Confere, antes de gerar, se a conta fecha. Devolve listas de problemas por área.
    'erro' = impede o horário de existir; 'aviso' = vale olhar, mas não impede."""
    comps = _componentes(conn, ano)
    turmas = _turmas(conn, ano)
    atrib = _atribuicoes(conn, ano)
    ed = _ed_professores(conn, ano)
    disp = _disponibilidades(conn, ano)
    nomes = {r["id"]: r["nome"] for r in conn.execute("SELECT id, nome FROM professores").fetchall()}
    comp_por_id = {c["id"]: c for c in comps}
    turma_por_id = {t["id"]: t for t in turmas}
    plan = _parametro(conn, ano, "planejamento_no_horario") == "1"
    blocos_plan = BLOCOS_PLANEJAMENTO if plan else 0

    def dias_disp(pid, turno):
        r = disp.get(pid)
        return [d for d, _ in DIAS if r and r[f"{d}_{turno}"]] if r else []

    gerais = []
    soma_blocos = sum(_blocos_comp(c) for c in comps)
    if comps and soma_blocos != BLOCOS_SEMANA:
        gerais.append(("erro", f"A matriz soma {soma_blocos * 2} tempos por semana; cada turma tem exatamente {BLOCOS_SEMANA * 2}."))

    # --- turmas sem professor ---
    turmas_prob = []
    for t in turmas:
        faltando = [c["nome"] for c in comps if (t["id"], c["id"]) not in atrib]
        if faltando:
            turmas_prob.append({"turma": t["nome"], "turno": t["turno"], "faltando": faltando})

    # --- Educação Digital por turno ---
    ed_turnos = []
    for k in ("m", "t"):
        do_turno = [t for t in turmas if t["turno"] == k]
        if not do_turno:
            continue
        carga_ano = {a: 0 for a in (6, 7, 8, 9)}
        for t in do_turno:
            carga_ano[t["ano_escolaridade"]] += sum(_blocos_comp(c) for c in comps if t["ano_escolaridade"] in _anos_ed(c))
        opcoes = []
        for par_a, par_b in PARES_ED:
            ca, cb = sum(carga_ano[a] for a in par_a), sum(carga_ano[a] for a in par_b)
            opcoes.append({"par_a": par_a, "par_b": par_b, "carga_a": ca, "carga_b": cb, "maior": max(ca, cb),
                           "cabe": max(ca, cb) <= BLOCOS_SEMANA})
        melhor = min(opcoes, key=lambda o: o["maior"])
        prof_ed = ed.get(k)
        problemas = []
        if not any(o["cabe"] for o in opcoes):
            problemas.append(("erro", f"Nenhuma divisão em pares cabe no Laboratório: o menor par precisa de {melhor['maior']} blocos e a semana tem {BLOCOS_SEMANA}."))
        total_ed = sum(carga_ano.values())
        if total_ed and not prof_ed:
            problemas.append(("erro", "Falta definir o professor de Educação Digital deste turno."))
        elif prof_ed:
            dias = dias_disp(prof_ed, k)
            if prof_ed not in disp:
                problemas.append(("erro", f"{nomes.get(prof_ed, '?')} (Educação Digital) não respondeu a coleta de disponibilidade."))
            elif len(dias) * BLOCOS_POR_TURNO < melhor["maior"]:
                problemas.append(("erro", f"{nomes.get(prof_ed, '?')} (Educação Digital) tem {_pl(len(dias), 'turno disponível', 'turnos disponíveis')} aqui "
                                          f"(cabem {len(dias) * BLOCOS_POR_TURNO} blocos), mas o par mais cheio precisa de {melhor['maior']}."))
        ed_turnos.append({"turno": k, "carga_ano": carga_ano, "opcoes": opcoes, "melhor": melhor,
                          "prof": prof_ed, "prof_nome": nomes.get(prof_ed, ""), "problemas": problemas})

    # --- professores: carga × disponibilidade ---
    carga = {}  # pid -> {"m": blocos, "t": blocos, "aulas": [(turma, comp, blocos, turno)]}
    for (tid, cid), pid in atrib.items():
        t, c = turma_por_id.get(tid), comp_por_id.get(cid)
        if not t or not c:
            continue
        dados = carga.setdefault(pid, {"m": 0, "t": 0, "aulas": [], "ed": {"m": 0, "t": 0}})
        b = _blocos_comp(c)
        dados[t["turno"]] += b
        dados["aulas"].append((t["nome"], c["nome"], b, t["turno"]))
    for item_ed in ed_turnos:
        if item_ed["prof"]:
            dados = carga.setdefault(item_ed["prof"], {"m": 0, "t": 0, "aulas": [], "ed": {"m": 0, "t": 0}})
            dados["ed"][item_ed["turno"]] = item_ed["melhor"]["maior"]
            dados[item_ed["turno"]] += item_ed["melhor"]["maior"]

    profs_linhas = []
    for pid, dados in carga.items():
        r = disp.get(pid)
        dm, dt = dias_disp(pid, "m"), dias_disp(pid, "t")
        cap_m, cap_t = len(dm) * BLOCOS_POR_TURNO, len(dt) * BLOCOS_POR_TURNO
        problemas = []
        if not r:
            problemas.append(("erro", "Não respondeu a coleta de disponibilidade."))
        else:
            if dados["m"] > cap_m:
                problemas.append(("erro", f"Manhã: {_pl(dados['m'], 'bloco', 'blocos')} de aula, mas só {_pl(len(dm), 'turno disponível', 'turnos disponíveis')} (cabem {cap_m})."))
            if dados["t"] > cap_t:
                problemas.append(("erro", f"Tarde: {_pl(dados['t'], 'bloco', 'blocos')} de aula, mas só {_pl(len(dt), 'turno disponível', 'turnos disponíveis')} (cabem {cap_t})."))
            total, cap = dados["m"] + dados["t"], cap_m + cap_t
            if plan and total <= cap and total + blocos_plan > cap:
                problemas.append(("erro", f"Sem espaço para os 4 tempos de planejamento: {total} blocos de aula para {cap} disponíveis."))
            if plan and not any(r[f"{d}_{t}"] for d in DIAS_PLANEJAMENTO for t in ("m", "t")):
                problemas.append(("erro", "Nenhum turno na terça, quarta ou quinta para o planejamento."))
            for turma_nome, comp_nome, b, turno in dados["aulas"]:
                dias = dm if turno == "m" else dt
                if b > len(dias) and b <= (cap_m if turno == "m" else cap_t):
                    problemas.append(("erro", f"{comp_nome} na {turma_nome} tem {b} blocos (no máximo um por dia), mas o professor só tem "
                                              f"{_pl(len(dias), 'dia disponível', 'dias disponíveis')} de {NOME_TURNO[turno].lower()}."))
            sobra = cap - total - blocos_plan
            if not problemas and sobra == 0:
                problemas.append(("aviso", "Disponibilidade exata: nenhum bloco de folga — o gerador terá pouca margem com este professor."))
        profs_linhas.append({"id": pid, "nome": nomes.get(pid, f"#{pid}"), "area": r["disciplina"] if r else "",
                             "m": dados["m"], "t": dados["t"], "dm": len(dm), "dt": len(dt), "aulas": dados["aulas"],
                             "ed": dados["ed"], "problemas": problemas})
    profs_linhas.sort(key=lambda p: (0 if any(x[0] == "erro" for x in p["problemas"]) else (1 if p["problemas"] else 2), p["nome"].lower()))

    sem_aula = sorted([nomes.get(pid, "?") for pid in disp if pid not in carga], key=str.lower)

    n_erros = (sum(1 for g in gerais if g[0] == "erro") + len(turmas_prob)
               + sum(1 for e in ed_turnos for p in e["problemas"] if p[0] == "erro")
               + sum(1 for p in profs_linhas for x in p["problemas"] if x[0] == "erro"))
    return {"gerais": gerais, "turmas": turmas_prob, "ed": ed_turnos, "professores": profs_linhas,
            "sem_aula": sem_aula, "n_erros": n_erros, "n_turmas": len(turmas), "planejamento_no_horario": plan}


@router.get("/horarios/configuracao/viabilidade", response_class=HTMLResponse)
def config_viabilidade(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    an = _analisar_viabilidade(conn, ano_i)
    conn.close()

    def lista_prob(probs):
        return "".join(
            f'<div style="font-size:12px; color:{"var(--red,#b91c1c)" if tipo == "erro" else "#b45309"};">{"❌" if tipo == "erro" else "⚠️"} {_e(txt)}</div>'
            for tipo, txt in probs)

    if not an["n_turmas"]:
        conteudo = _cabecalho_config("🔎 Viabilidade", "", "viabilidade", ano_i) + \
            f'<div class="empty">Cadastre as turmas e as atribuições primeiro. <a href="/horarios/configuracao/turmas?ano={ano_i}">Ir para Turmas →</a></div>'
        return _pagina(request, "Viabilidade", conteudo, "horarios-config")

    if an["n_erros"] == 0:
        topo = ('<div style="background:#ecfdf5; color:#047857; border:1px solid #10b981; padding:14px 16px; border-radius:8px; margin-bottom:16px; font-size:15px;">'
                '✅ <strong>Nenhum impedimento encontrado.</strong> A conta fecha: dá para seguir para a geração do horário.</div>')
    else:
        topo = (f'<div style="background:var(--red-bg,#fef2f2); color:var(--red,#b91c1c); border:1px solid var(--red,#b91c1c); padding:14px 16px; '
                f'border-radius:8px; margin-bottom:16px; font-size:15px;">❌ <strong>{an["n_erros"]} impedimento(s)</strong> — com eles, nenhum horário consegue '
                'respeitar todas as regras. Resolva ajustando atribuições ou a disponibilidade (pelo painel da coleta).</div>')
    topo += lista_prob(an["gerais"])

    # Educação Digital
    blocos_ed = ""
    for e in an["ed"]:
        linhas_op = ""
        for o in e["opcoes"]:
            eh_melhor = o is e["melhor"]
            pa = " + ".join(f"{a}º" for a in o["par_a"])
            pb = " + ".join(f"{a}º" for a in o["par_b"])
            linhas_op += (f'<tr style="border-top:1px solid var(--border); {"background:#ecfdf5;" if eh_melhor and o["cabe"] else ""}">'
                          f'<td style="padding:6px 8px;">Semana 1: <strong>{pa}</strong> · Semana 2: <strong>{pb}</strong></td>'
                          f'<td style="padding:6px 8px; text-align:center;">{o["carga_a"]} e {o["carga_b"]}</td>'
                          f'<td style="padding:6px 8px; text-align:center;">{"✅ cabe" if o["cabe"] else "❌ não cabe"}{" · melhor" if eh_melhor and o["cabe"] else ""}</td></tr>')
        cargas = " · ".join(f"{a}º: {e['carga_ano'][a]}" for a in (6, 7, 8, 9) if e["carga_ano"][a])
        blocos_ed += f"""
            <div style="flex:1 1 380px;">
                <h3 style="margin:0 0 4px; font-size:14px;">{NOME_TURNO[e["turno"]]} — professor: {_e(e["prof_nome"]) or '<span style="color:var(--red);">não definido</span>'}</h3>
                <p style="margin:0 0 6px; font-size:12px; color:var(--text-muted);">Blocos com Educação Digital por ano: {cargas or "nenhum"}</p>
                <table style="width:100%; border-collapse:collapse; font-size:12px;">
                    <thead><tr style="background:var(--bg-subtle);"><th style="padding:6px 8px; text-align:left;">Divisão em pares</th>
                    <th style="padding:6px 8px;">Blocos por par (máx. {BLOCOS_SEMANA})</th><th style="padding:6px 8px;">Laboratório</th></tr></thead>
                    <tbody>{linhas_op}</tbody>
                </table>
                <div style="margin-top:6px;">{lista_prob(e["problemas"])}</div>
            </div>"""
    card_ed = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 4px; font-size:15px;">💻 Educação Digital e Laboratório</h2>
            <p style="margin:0 0 10px; font-size:12px; color:var(--text-muted);">Quinzenal: cada par de anos usa o Laboratório numa semana, uma turma por bloco. O gerador escolhe a divisão que fecha melhor.</p>
            <div style="display:flex; flex-wrap:wrap; gap:20px;">{blocos_ed}</div>
        </div>"""

    # Turmas incompletas
    card_turmas = ""
    if an["turmas"]:
        linhas_t = "".join(f'<tr style="border-top:1px solid var(--border);"><td style="padding:6px 8px; font-weight:700;">{_e(x["turma"])}</td>'
                           f'<td style="padding:6px 8px;">{NOME_TURNO[x["turno"]]}</td><td style="padding:6px 8px; font-size:12px;">{_e(", ".join(x["faltando"]))}</td></tr>'
                           for x in an["turmas"])
        card_turmas = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 8px; font-size:15px;">❌ Turmas com aulas sem professor ({len(an["turmas"])})</h2>
            <table style="border-collapse:collapse; font-size:13px;"><thead><tr style="background:var(--bg-subtle);"><th style="padding:6px 8px; text-align:left;">Turma</th>
            <th style="padding:6px 8px; text-align:left;">Turno</th><th style="padding:6px 8px; text-align:left;">Sem professor</th></tr></thead><tbody>{linhas_t}</tbody></table>
            <div style="margin-top:10px;"><a href="/horarios/configuracao/atribuicoes?ano={ano_i}" class="btn">🔗 Completar atribuições</a></div>
        </div>"""

    # Professores
    linhas_p = ""
    for p in an["professores"]:
        status = ("❌" if any(x[0] == "erro" for x in p["problemas"]) else ("⚠️" if p["problemas"] else "✅"))
        resumo_aulas = ", ".join(f"{c} {t}" for t, c, b, tu in sorted(p["aulas"]))
        if p["ed"]["m"] or p["ed"]["t"]:
            resumo_aulas = (resumo_aulas + ", " if resumo_aulas else "") + "Educação Digital " + " e ".join(NOME_TURNO[k].lower() for k in ("m", "t") if p["ed"][k])

        def cel(blocos, dias):
            if not blocos and not dias:
                return '<span style="color:var(--text-muted);">—</span>'
            cor = "color:var(--red,#b91c1c); font-weight:700;" if blocos > dias * BLOCOS_POR_TURNO else ""
            return f'<span style="{cor}">{blocos}</span> <span style="color:var(--text-muted); font-size:11px;">/ {dias * BLOCOS_POR_TURNO}</span>'
        linhas_p += f"""<tr style="border-top:1px solid var(--border); vertical-align:top;">
            <td style="padding:6px 8px; text-align:center; font-size:16px;">{status}</td>
            <td style="padding:6px 8px;"><strong>{_e(p["nome"])}</strong><div style="font-size:11px; color:var(--text-muted);">{_e(p["area"])}</div></td>
            <td style="padding:6px 8px; text-align:center; white-space:nowrap;">{cel(p["m"], p["dm"])}</td>
            <td style="padding:6px 8px; text-align:center; white-space:nowrap;">{cel(p["t"], p["dt"])}</td>
            <td style="padding:6px 8px; font-size:12px; max-width:320px;">{_e(resumo_aulas)}{lista_prob(p["problemas"])}</td>
        </tr>"""
    n_prof_erro = sum(1 for p in an["professores"] if any(x[0] == "erro" for x in p["problemas"]))
    card_prof = f"""
        <div class="card" style="padding:14px 16px; margin-bottom:16px;">
            <h2 style="margin:0 0 4px; font-size:15px;">👩‍🏫 Professores — aulas × disponibilidade ({len(an["professores"])}, {n_prof_erro} com impedimento)</h2>
            <p style="margin:0 0 8px; font-size:12px; color:var(--text-muted);">Em cada turno: <strong>blocos de aula</strong> / blocos que cabem nos turnos disponíveis (3 por turno).
            {"Os 4 tempos de planejamento (2 blocos) precisam caber na sobra, em terça, quarta ou quinta." if an["planejamento_no_horario"] else "Planejamento fora do horário (opção na aba Tempos): não entra na conta."}</p>
            <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead><tr style="background:var(--bg-subtle);"><th style="padding:6px 8px;"></th><th style="padding:6px 8px; text-align:left;">Professor</th>
                <th style="padding:6px 8px;">Manhã</th><th style="padding:6px 8px;">Tarde</th><th style="padding:6px 8px; text-align:left;">Aulas e observações</th></tr></thead>
                <tbody>{linhas_p or '<tr><td colspan="5" style="padding:14px; text-align:center; color:var(--text-muted);">Nenhuma aula atribuída ainda.</td></tr>'}</tbody>
            </table></div>
        </div>"""

    card_sem = ""
    if an["sem_aula"]:
        card_sem = f"""
        <details class="card" style="padding:14px 16px; margin-bottom:16px;">
            <summary style="cursor:pointer; font-weight:600;">Responderam a coleta mas não têm aula atribuída ({len(an["sem_aula"])})</summary>
            <p style="font-size:13px; margin:8px 0 0;">{_e(", ".join(an["sem_aula"]))}</p>
        </details>"""

    conteudo = (_cabecalho_config("🔎 Viabilidade", "Confere se a conta fecha antes de gerar. Não monta o horário — só aponta o que certamente impediria um horário de existir.",
                                  "viabilidade", ano_i) + topo + card_turmas + card_ed + card_prof + card_sem)
    return _pagina(request, "Viabilidade do horário", conteudo, "horarios-config")


# ==========================================
#  GERADOR AUTOMÁTICO (Fase 2 — 01/11/2026)
# ==========================================
# Usa o OR-Tools (biblioteca gratuita do Google, "CP-SAT") para montar o horário inteiro.
# Como funciona, em palavras simples:
#   - Para cada turma, cada componente e cada lugar da semana (dia × bloco) existe uma
#     pergunta de sim/não: "este componente acontece aqui?".
#   - As REGRAS OBRIGATÓRIAS viram restrições que o resolvedor nunca quebra.
#   - As REGRAS DESEJÁVEIS viram "multas"; o resolvedor procura o horário com a menor soma.
# Se a disponibilidade informada não permite nenhum horário, roda de novo com a
# disponibilidade "flexível" (multa altíssima para cada aula fora dela) e mostra
# exatamente quais turnos fariam o horário fechar — é o diagnóstico.

DIAS_SEMANA = [d for d, _ in DIAS]           # seg, ter, qua, qui, sex
NOME_DIA = dict(DIAS)

# Pesos das multas (quanto maior, mais o gerador evita). Ajustáveis no futuro.
PESOS_PADRAO = {
    "janela": 10,                   # professor com bloco vago entre duas aulas no mesmo turno
    "turno_usado": 4,               # cada turno em que o professor vem à escola (incentiva turnos cheios)
}
# Componente em dias consecutivos é normal na escola: não entra como multa (Felipe, 01/11/2026).


def _ortools_disponivel() -> bool:
    try:
        from ortools.sat.python import cp_model  # noqa: F401
        return True
    except ImportError:
        return False


def _dados_geracao(conn, ano: int) -> dict:
    """Junta tudo o que o gerador precisa, num formato simples (sem objetos do banco)."""
    comps = {c["id"]: {"nome": c["nome"], "blocos": _blocos_comp(c), "ed_anos": _anos_ed(c)} for c in _componentes(conn, ano)}
    turmas = {t["id"]: {"nome": t["nome"], "ano": t["ano_escolaridade"], "turno": t["turno"]} for t in _turmas(conn, ano)}
    atrib = _atribuicoes(conn, ano)
    disp_rows = _disponibilidades(conn, ano)
    disp = {pid: {(d, k) for d in DIAS_SEMANA for k in ("m", "t") if r[f"{d}_{k}"]} for pid, r in disp_rows.items()}
    nomes = {r["id"]: r["nome"] for r in conn.execute("SELECT id, nome FROM professores").fetchall()}
    return {"ano": ano, "comps": comps, "turmas": turmas, "atrib": atrib, "ed": _ed_professores(conn, ano),
            "disp": disp, "nomes": nomes}


def _resolver_horario(dados: dict, limite_segundos: int = 60, flexivel: bool = False, pesos: Optional[dict] = None,
                      sementes: int = 0, progresso=None, tempo_sugestoes: int = 45) -> dict:
    """Monta e resolve o modelo. Devolve um dicionário com status, aulas, indicadores e — no
    modo flexível — as pendências (aulas fora da disponibilidade) e as sugestões para cada uma.
    `progresso`, se informado, é chamado durante a busca com os números ao vivo."""
    from ortools.sat.python import cp_model
    p = dict(PESOS_PADRAO, **(pesos or {}))
    comps, turmas, atrib, disp, ed = dados["comps"], dados["turmas"], dados["atrib"], dados["disp"], dados["ed"]
    D, B = range(len(DIAS_SEMANA)), range(BLOCOS_POR_TURNO)
    m = cp_model.CpModel()
    avisar = progresso or (lambda info: None)

    def disponivel(pid, d, k):
        # Professor sem resposta na coleta: tratado como disponível (a viabilidade já avisa).
        return pid not in disp or (DIAS_SEMANA[d], k) in disp[pid]

    x = {}
    for tid, t in turmas.items():
        for cid, c in comps.items():
            if c["blocos"] <= 0:
                continue
            for d in D:
                for b in B:
                    x[tid, cid, d, b] = m.NewBoolVar(f"x_{tid}_{cid}_{d}_{b}")

    # Pendências possíveis (só no modo flexível): (professor, dia, turno) -> [(variável, tid, cid, bloco)]
    grupos_fora = {}

    for tid, t in turmas.items():
        cids = [cid for cid in comps if comps[cid]["blocos"] > 0]
        for d in D:
            for b in B:
                m.AddExactlyOne(x[tid, cid, d, b] for cid in cids)          # 1 componente por bloco
        for cid in cids:
            m.Add(sum(x[tid, cid, d, b] for d in D for b in B) == comps[cid]["blocos"])   # matriz
            for d in D:
                m.Add(sum(x[tid, cid, d, b] for b in B) <= 1)               # no máximo 1 bloco por dia

    # Professor: nunca em duas turmas no mesmo bloco; só onde está disponível.
    por_prof = {}
    for (tid, cid), pid in atrib.items():
        if tid in turmas and cid in comps and comps[cid]["blocos"] > 0:
            por_prof.setdefault(pid, []).append((tid, cid))
    ocup = {}   # (pid, turno, d, b) -> lista de variáveis
    for pid, aulas in por_prof.items():
        for tid, cid in aulas:
            k = turmas[tid]["turno"]
            for d in D:
                for b in B:
                    ocup.setdefault((pid, k, d, b), []).append(x[tid, cid, d, b])
                if not disponivel(pid, d, k):
                    for b in B:
                        if flexivel:
                            grupos_fora.setdefault((pid, d, k), []).append((x[tid, cid, d, b], tid, cid, b + 1))
                        else:
                            m.Add(x[tid, cid, d, b] == 0)
    for vs in ocup.values():
        if len(vs) > 1:
            m.AddAtMostOne(vs)

    # Educação Digital quinzenal: em cada turno, escolher uma das 3 divisões em pares; dentro
    # de um par, no máximo uma turma com ED por bloco (Laboratório). Professor de ED disponível.
    escolha_par = {}
    for k in ("m", "t"):
        tids_k = [tid for tid, t in turmas.items() if t["turno"] == k]
        ed_occ = {}
        for tid in tids_k:
            cids_ed = [cid for cid, c in comps.items() if turmas[tid]["ano"] in c["ed_anos"] and c["blocos"] > 0]
            if cids_ed:
                for d in D:
                    for b in B:
                        ed_occ[tid, d, b] = sum(x[tid, cid, d, b] for cid in cids_ed)
        if not ed_occ:
            continue
        pid_ed = ed.get(k)
        if pid_ed:
            for d in D:
                if not disponivel(pid_ed, d, k):
                    for b in B:
                        for tid in tids_k:
                            if (tid, d, b) in ed_occ:
                                if flexivel:
                                    v = m.NewBoolVar(f"edfora_{tid}_{d}_{b}")
                                    m.Add(ed_occ[tid, d, b] <= v)
                                    grupos_fora.setdefault((pid_ed, d, k), []).append((v, tid, None, b + 1))
                                else:
                                    m.Add(ed_occ[tid, d, b] == 0)
        z = [m.NewBoolVar(f"par_{k}_{j}") for j in range(len(PARES_ED))]
        m.AddExactlyOne(z)
        escolha_par[k] = z
        for j, (pa, pb) in enumerate(PARES_ED):
            for par in (pa, pb):
                tids_par = [tid for tid in tids_k if turmas[tid]["ano"] in par]
                for d in D:
                    for b in B:
                        termos = [ed_occ[tid, d, b] for tid in tids_par if (tid, d, b) in ed_occ]
                        if len(termos) > 1:
                            m.Add(sum(termos) <= 1).OnlyEnforceIf(z[j])

    # ----- multas (regras desejáveis) -----
    # (Componente em dias seguidos NÃO é multado: na escola isso é normal — Felipe, 01/11/2026.)
    n_janelas, n_turnos = [], []
    for (pid, k, d, b0), _ in list(ocup.items()):
        if b0 != 0:
            continue
        o = [sum(ocup.get((pid, k, d, b), [])) for b in B]
        usado = m.NewBoolVar(f"u_{pid}_{k}_{d}")
        for b in B:
            if ocup.get((pid, k, d, b)):
                m.Add(o[b] <= usado)
        n_turnos.append(usado)
        janela = m.NewBoolVar(f"j_{pid}_{k}_{d}")
        m.Add(janela >= o[0] + o[2] - o[1] - 1)     # aula no 1º e no 3º bloco, vago no 2º
        n_janelas.append(janela)
    qualidade = p["janela"] * sum(n_janelas) + p["turno_usado"] * sum(n_turnos)

    def novo_solver(segundos):
        sv = cp_model.CpSolver()
        sv.parameters.max_time_in_seconds = max(1.0, float(segundos))
        sv.parameters.num_search_workers = 2      # a VM é pequena (2 vCPUs)
        sv.parameters.random_seed = sementes
        return sv

    todas_fora = [v for g in grupos_fora.values() for v, *_ in g]
    soma_fora = sum(todas_fora) if todas_fora else 0
    # "Este turno do professor está sendo usado fora da disponibilidade?" — um sim/não por turno,
    # usado na análise das sugestões (para testar uma alternativa de cada vez).
    ativa_grupo = {}
    for chave, itens in grupos_fora.items():
        a = m.NewBoolVar(f"g_{chave[0]}_{chave[1]}_{chave[2]}")
        for v, *_ in itens:
            m.Add(v <= a)
        ativa_grupo[chave] = a

    def grupos_ativos(valor):
        """Pendências de uma solução: {(pid, d, k): [(tid, cid, bloco), ...]}"""
        ativos = {}
        for chave, itens in grupos_fora.items():
            for v, tid, cid, b in itens:
                if valor(v):
                    ativos.setdefault(chave, []).append((tid, cid, b))
        return ativos

    class _AoVivo(cp_model.CpSolverSolutionCallback):
        def __init__(self, etapa):
            super().__init__()
            self.etapa = etapa

        def on_solution_callback(self):
            try:
                ativos = grupos_ativos(self.Value)
                avisar({"etapa": self.etapa, "pendencias": sum(len(v) for v in ativos.values()),
                        "pendencias_lista": _descrever_pendencias(ativos, dados), "sugestoes": len(ativos)})
            except Exception:
                pass

    tempo_total = 0.0
    if todas_fora:
        # Modo flexível em duas etapas: 1ª) o MENOR número possível de aulas fora da
        # disponibilidade; 2ª) com esse número travado, o horário de melhor qualidade.
        m.Minimize(soma_fora)
        sv1 = novo_solver(limite_segundos * 0.4)
        st1 = sv1.Solve(m, _AoVivo("Reduzindo as pendências…"))
        tempo_total += sv1.WallTime()
        mq = m.clone()
        dica = {}
        if st1 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            minimo = int(round(sv1.ObjectiveValue()))
            mq.Add(soma_fora <= minimo)
            dica = {v.Index(): sv1.Value(v) for v in x.values()}
            for v in x.values():
                mq.AddHint(v, dica[v.Index()])
        mq.Minimize(qualidade)
        solver = novo_solver(limite_segundos - tempo_total)
        status = solver.Solve(mq, _AoVivo("Melhorando a qualidade do horário…"))
    else:
        m.Minimize(qualidade)
        solver = novo_solver(limite_segundos)
        status = solver.Solve(m)
    tempo_total += solver.WallTime()
    nome_status = {cp_model.OPTIMAL: "otimo", cp_model.FEASIBLE: "viavel", cp_model.INFEASIBLE: "impossivel",
                   cp_model.MODEL_INVALID: "invalido"}.get(status, "sem_solucao")
    saida = {"status": nome_status, "tempo": round(tempo_total, 1), "flexivel": flexivel}
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return saida

    aulas = []
    for (tid, cid, d, b), v in x.items():
        if solver.Value(v):
            aulas.append({"turma_id": tid, "componente_id": cid, "professor_id": atrib.get((tid, cid)),
                          "dia": d, "bloco": b + 1, "turno": turmas[tid]["turno"]})
    pares = {}
    for k, z in escolha_par.items():
        j = next(i for i, v in enumerate(z) if solver.Value(v))
        pares[k] = [list(PARES_ED[j][0]), list(PARES_ED[j][1])]
    ativos = grupos_ativos(solver.Value)
    fora_lista = [{"professor_id": pid, "dia": d, "turno": k} for (pid, d, k) in ativos]

    # ----- sugestões para cada pendência (estilo Urânia) -----
    # Para cada pendência, pergunta: "se este professor NÃO puder mesmo neste turno, existe
    # UM outro turno (de qualquer professor) que, liberado sozinho, resolve — mantendo as
    # demais pendências como estão?". Cada resposta é uma alternativa independente:
    #   mesmo total de pendências → "ajudará muito"; com mais aulas fora → "ajudará um pouco".
    sugestoes = []
    if ativos:
        total = sum(len(v) for v in ativos.values())
        chaves = sorted(ativos, key=lambda c: (dados["nomes"].get(c[0], ""), c[1], c[2]))
        prazo = _time.time() + tempo_sugestoes
        dica = {v.Index(): solver.Value(v) for v in x.values()}
        for i, chave in enumerate(chaves, start=1):
            pid, d, k = chave
            item = {"n": i, "professor_id": pid, "professor_nome": dados["nomes"].get(pid, "?"), "dia": d, "turno": k,
                    "blocos": sorted({b for _, _, b in ativos[chave]}),
                    "aulas": _rotulos_aulas(ativos[chave], dados), "sem_resposta": pid not in disp,
                    "muito": [], "pouco": [], "analisada": False}
            avisar({"etapa": f"Analisando sugestões ({i} de {len(chaves)})…", "pendencias": total,
                    "pendencias_lista": _descrever_pendencias(ativos, dados), "sugestoes": len(chaves)})
            ja_testadas = []
            for _tentativa in range(3):            # até 3 alternativas por pendência
                restante = prazo - _time.time()
                if restante < 1.0:
                    break
                mt = m.clone()
                for v, *_ in grupos_fora[chave]:
                    mt.Add(v == 0)
                novas = [ativa_grupo[c] for c in grupos_fora if c not in ativos]
                if novas:
                    mt.Add(sum(novas) <= 1)          # só UM turno novo liberado
                for c in ja_testadas:
                    mt.Add(ativa_grupo[c] == 0)
                mt.Minimize(soma_fora)
                for v in x.values():
                    mt.AddHint(v, dica[v.Index()])
                svt = novo_solver(min(3.0, max(1.0, restante / max(1, len(chaves) - i + 1))))
                stt = svt.Solve(mt)
                if stt not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    break
                novos = grupos_ativos(svt.Value)
                alternativa = next((c for c in novos if c not in ativos), None)
                if alternativa is None:
                    break
                novo_total = int(round(svt.ObjectiveValue()))
                destino = "muito" if novo_total <= total else "pouco"
                pa, da, ka = alternativa
                item[destino].append({"professor_id": pa, "professor_nome": dados["nomes"].get(pa, "?"), "dia": da, "turno": ka,
                                      "blocos": sorted({b for _, _, b in novos[alternativa]})})
                ja_testadas.append(alternativa)
            item["analisada"] = True
            sugestoes.append(item)

    saida.update({
        "aulas": aulas, "pares_ed": pares, "fora": fora_lista, "sugestoes": sugestoes,
        "indicadores": {"janelas": int(sum(solver.Value(v) for v in n_janelas)),
                        "turnos_usados": int(sum(solver.Value(v) for v in n_turnos)),
                        "fora_disponibilidade": len(fora_lista),
                        "pendencias": sum(len(v) for v in ativos.values()),
                        "sugestoes": sum(1 + len(s["muito"]) + len(s["pouco"]) for s in sugestoes)},
        "objetivo": int(solver.ObjectiveValue()),
    })
    return saida


def _rotulos_aulas(itens, dados) -> list:
    """[(tid, cid, bloco)] -> ['Ciências 703', ...] (cid None = Educação Digital)."""
    rot = []
    for tid, cid, _ in itens:
        nome = dados["comps"][cid]["nome"] if cid in dados["comps"] else "Educação Digital"
        txt = f'{nome} {dados["turmas"][tid]["nome"]}'
        if txt not in rot:
            rot.append(txt)
    return rot


def _descrever_pendencias(ativos, dados) -> list:
    return [{"professor_nome": dados["nomes"].get(pid, "?"), "dia": d, "turno": k,
             "blocos": sorted({b for _, _, b in itens}), "aulas": _rotulos_aulas(itens, dados)}
            for (pid, d, k), itens in sorted(ativos.items(), key=lambda c: (dados["nomes"].get(c[0][0], ""), c[0][1]))]


# ---------- versões geradas: execução em segundo plano ----------

import threading as _threading
import json as _json
import time as _time

_TRAVA_GERACAO = _threading.Lock()     # um horário por vez (a VM é pequena)
_GERACOES = {}                          # versao_id -> {"inicio": t, "limite": s, "etapa": texto}


def _marcar_geracoes_interrompidas():
    """Se o sistema reiniciou no meio de uma geração, a versão fica marcada como erro
    (em vez de 'gerando' para sempre). Só muda o status — não mexe em nenhum outro dado."""
    conn = _db()
    conn.execute("""UPDATE horario_versoes SET status = 'erro',
                    mensagem = 'Interrompida: o sistema reiniciou durante a geração. Gere de novo.'
                    WHERE status = 'gerando'""")
    conn.commit()
    conn.close()


def _pre_requisitos_geracao(conn, ano: int) -> list:
    """O que impede de começar a gerar (diferente da viabilidade: aqui é o mínimo)."""
    erros = []
    if not _ortools_disponivel():
        erros.append("O OR-Tools não está instalado no servidor. Na VM, rode: pip install ortools (dentro da venv do sistema) e reinicie o serviço.")
    comps = _componentes(conn, ano)
    turmas = _turmas(conn, ano)
    if not turmas:
        erros.append("Nenhuma turma cadastrada.")
    soma = sum(_blocos_comp(c) for c in comps)
    if comps and soma != BLOCOS_SEMANA:
        erros.append(f"A matriz soma {soma * 2} tempos; precisa somar {BLOCOS_SEMANA * 2}.")
    atrib = _atribuicoes(conn, ano)
    faltando = sum(1 for t in turmas for c in comps if _blocos_comp(c) > 0 and (t["id"], c["id"]) not in atrib)
    if faltando:
        erros.append(f"{faltando} aula(s) ainda sem professor em “Quem dá cada aula”.")
    return erros


def _executar_geracao(versao_id: int, ano: int, limite: int):
    """Roda em segundo plano (thread). Tenta primeiro respeitando 100% da disponibilidade;
    se não existir horário assim, roda o modo flexível (diagnóstico)."""
    with _TRAVA_GERACAO:
        conn = _db()
        try:
            dados = _dados_geracao(conn, ano)
            _GERACOES[versao_id] = {"inicio": _time.time(), "limite": limite, "etapa": "Procurando um horário que respeite toda a disponibilidade…"}
            r = _resolver_horario(dados, limite_segundos=limite, flexivel=False)
            if r["status"] not in ("otimo", "viavel"):
                _GERACOES[versao_id] = {"inicio": _time.time(), "limite": limite, "flexivel": True,
                                        "etapa": "Não existe horário dentro da disponibilidade informada. Procurando o horário com o menor número de aulas fora dela…"}

                def ao_vivo(info, _vid=versao_id):
                    estado = _GERACOES.get(_vid)
                    if estado is not None:
                        estado.update(info)
                r = _resolver_horario(dados, limite_segundos=limite, flexivel=True, progresso=ao_vivo,
                                      tempo_sugestoes=max(15, min(60, limite // 2)))
            if r["status"] not in ("otimo", "viavel"):
                conn.execute("UPDATE horario_versoes SET status = 'erro', resultado = ?, mensagem = ? WHERE id = ?",
                             (r["status"], "Nenhum horário encontrado, nem com a disponibilidade flexível. Confira a viabilidade "
                                           "(matriz, turmas, Educação Digital) ou tente com mais tempo.", versao_id))
                conn.commit()
                return
            turmas = dados["turmas"]
            comps = dados["comps"]
            for a in r["aulas"]:
                conn.execute("""INSERT INTO horario_aulas (versao_id, turma_id, componente_id, professor_id, turno, dia, bloco,
                                turma_nome, componente_nome, professor_nome) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                             (versao_id, a["turma_id"], a["componente_id"], a["professor_id"], a["turno"], a["dia"], a["bloco"],
                              turmas[a["turma_id"]]["nome"], comps[a["componente_id"]]["nome"], dados["nomes"].get(a["professor_id"], "")))
            ed_info = {k: {"pares": v, "professor_id": dados["ed"].get(k), "professor_nome": dados["nomes"].get(dados["ed"].get(k), "")}
                       for k, v in r["pares_ed"].items()}
            fora = [dict(f, professor_nome=dados["nomes"].get(f["professor_id"], "")) for f in r["fora"]]
            # Quem tem aula mas não respondeu a coleta foi tratado como disponível em tudo — registrar para avisar.
            com_aula = set(dados["atrib"].values()) | {p for p in dados["ed"].values() if p}
            r["indicadores"]["sem_resposta"] = sorted((dados["nomes"].get(p, "?") for p in com_aula if p not in dados["disp"]), key=str.lower)
            conn.execute("""UPDATE horario_versoes SET status = 'pronta', modo = ?, resultado = ?, tempo = ?, indicadores = ?,
                            pares_ed = ?, fora = ?, ed_anos = ?, sugestoes = ? WHERE id = ?""",
                         ("flexivel" if r["flexivel"] else "estrito", r["status"], r["tempo"], _json.dumps(r["indicadores"]),
                          _json.dumps(ed_info), _json.dumps(fora, ensure_ascii=False),
                          _json.dumps({cid: sorted(c["ed_anos"]) for cid, c in comps.items() if c["ed_anos"]}),
                          _json.dumps(r.get("sugestoes", []), ensure_ascii=False), versao_id))
            conn.commit()
        except Exception as erro:  # registra o erro na versão em vez de sumir em silêncio
            print(f"[horarios] geração {versao_id} falhou: {erro}")
            conn.execute("UPDATE horario_versoes SET status = 'erro', mensagem = ? WHERE id = ?", (f"Erro inesperado: {erro}", versao_id))
            conn.commit()
        finally:
            _GERACOES.pop(versao_id, None)
            conn.close()


# ---------- tela: gerar e listar versões ----------

@router.get("/horarios/geracao", response_class=HTMLResponse)
def geracao_inicio(request: Request, ano: str = "", msg: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    ano_i = _ano_config(conn, ano)
    _garantir_config(conn, ano_i)
    pre = _pre_requisitos_geracao(conn, ano_i)
    analise = _analisar_viabilidade(conn, ano_i)
    versoes = conn.execute("""SELECT v.*, p.nome AS autor FROM horario_versoes v LEFT JOIN professores p ON p.id = v.criada_por
                              WHERE v.ano_letivo = ? ORDER BY v.id DESC""", (ano_i,)).fetchall()
    conn.close()

    if pre:
        bloco_gerar = ('<div class="hd-erros"><strong>Antes de gerar:</strong><br>' + "".join(f"• {_e(x)}<br>" for x in pre) + "</div>")
    else:
        aviso_viab = ""
        if analise["n_erros"]:
            aviso_viab = (f'<div class="hd-avisos" style="display:block;">⚠️ A viabilidade aponta <strong>{analise["n_erros"]} impedimento(s)</strong>. '
                          'Dá para gerar mesmo assim: o gerador vai montar o horário com o menor número possível de aulas fora da '
                          f'disponibilidade e mostrar quais são. <a href="/horarios/configuracao/viabilidade?ano={ano_i}">Ver viabilidade</a></div>')
        ocupado = _TRAVA_GERACAO.locked()
        bloco_gerar = f"""
            {aviso_viab}
            <form method="post" action="/horarios/geracao/gerar?ano={ano_i}" style="display:flex; flex-wrap:wrap; gap:12px; align-items:flex-end;">
                <label style="margin:0;">Tempo de busca
                    <select name="limite" style="width:200px;">
                        <option value="60">1 minuto (rascunho)</option>
                        <option value="120" selected>2 minutos</option>
                        <option value="300">5 minutos (melhor qualidade)</option>
                    </select>
                </label>
                <button type="submit" class="btn btn-primary" {"disabled" if ocupado else ""}>🚀 Gerar horário</button>
                {'<span style="font-size:12px; color:var(--text-muted);">Já existe uma geração em andamento — aguarde terminar.</span>' if ocupado else ""}
            </form>
            <p style="font-size:12px; color:var(--text-muted); margin:10px 0 0;">Mais tempo = horário melhor (menos janelas e turnos picados). Se a disponibilidade não
            permitir nenhum horário, o gerador roda uma segunda busca de diagnóstico e depois analisa as sugestões para cada pendência —
            o tempo total pode passar do dobro. Você acompanha as pendências ao vivo.</p>"""

    linhas = ""
    for v in versoes:
        ind = _json.loads(v["indicadores"]) if v["indicadores"] else {}
        if v["status"] == "pronta":
            situacao = ("✅ Dentro da disponibilidade" if v["modo"] == "estrito"
                        else f'⚠️ {ind.get("pendencias", ind.get("fora_disponibilidade", "?"))} pendência(s) · {ind.get("sugestoes", "?")} sugestões')
            qual = f'{ind.get("janelas", "—")} janelas · {ind.get("turnos_usados", "—")} turnos de presença'
        elif v["status"] == "gerando":
            situacao, qual = "⏳ Gerando…", ""
        else:
            situacao, qual = "❌ Erro", _e(v["mensagem"] or "")
        linhas += f"""<tr style="border-top:1px solid var(--border);">
            <td style="padding:8px; font-weight:700;">#{v["id"]}</td>
            <td style="padding:8px; font-size:12px; white-space:nowrap;">{_ts_br(v["criada_em"])}<div style="color:var(--text-muted);">{_e(v["autor"] or "")}</div></td>
            <td style="padding:8px;">{situacao}</td>
            <td style="padding:8px; font-size:12px;">{qual}</td>
            <td style="padding:8px; white-space:nowrap;">
                <a href="/horarios/geracao/versao/{v["id"]}" class="btn" style="padding:3px 10px; font-size:12px;">Abrir</a>
                <form method="post" action="/horarios/geracao/versao/{v["id"]}/excluir" style="display:inline;" onsubmit="return confirm('Excluir esta versão do horário?');">
                    <button type="submit" class="btn" style="padding:3px 8px; font-size:12px; color:var(--red);">🗑️</button></form>
            </td>
        </tr>"""
    tabela = (f"""<table style="width:100%; border-collapse:collapse; font-size:13px;">
        <thead><tr style="background:var(--bg-subtle);"><th style="padding:8px; text-align:left;">Versão</th><th style="padding:8px; text-align:left;">Gerada em</th>
        <th style="padding:8px; text-align:left;">Situação</th><th style="padding:8px; text-align:left;">Qualidade</th><th style="padding:8px;"></th></tr></thead>
        <tbody>{linhas}</tbody></table>""" if linhas else '<p style="color:var(--text-muted); font-size:13px;">Nenhuma versão gerada ainda.</p>')

    msgs = {"excluida": "Versão excluída.", "ocupado": "Já existe uma geração em andamento. Aguarde terminar."}
    conteudo = _cabecalho_config("🚀 Gerar horário", "Monta o horário inteiro automaticamente, respeitando as regras da escola. Cada geração vira uma versão — dá para comparar e escolher.",
                                 "geracao", ano_i, _msg_box(msgs.get(msg, ""), erro=(msg == "ocupado"))) + f"""
        {_CSS_FORM}
        <div class="card" style="padding:16px; margin-bottom:16px;">{bloco_gerar}</div>
        <div class="card" style="padding:16px;"><h2 style="margin:0 0 10px; font-size:15px;">Versões geradas</h2>{tabela}</div>"""
    return _pagina(request, "Gerar horário", conteudo, "horarios-geracao")


@router.post("/horarios/geracao/gerar")
async def geracao_gerar(request: Request, ano: str = ""):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    form = await request.form()
    try:
        limite = int(form.get("limite") or 120)
    except ValueError:
        limite = 120
    limite = max(10, min(limite, 600))
    conn = _db()
    ano_i = _ano_config(conn, ano)
    if _pre_requisitos_geracao(conn, ano_i):
        conn.close()
        return RedirectResponse(f"/horarios/geracao?ano={ano_i}", status_code=303)
    if _TRAVA_GERACAO.locked():
        conn.close()
        return RedirectResponse(f"/horarios/geracao?ano={ano_i}&msg=ocupado", status_code=303)
    cur = conn.execute("INSERT INTO horario_versoes (ano_letivo, criada_por, status, limite_segundos) VALUES (?, ?, 'gerando', ?)",
                       (ano_i, prof["id"], limite))
    versao_id = cur.lastrowid
    conn.commit()
    conn.close()
    _GERACOES[versao_id] = {"inicio": _time.time(), "limite": limite, "etapa": "Preparando os dados…"}
    _threading.Thread(target=_executar_geracao, args=(versao_id, ano_i, limite), daemon=True).start()
    return RedirectResponse(f"/horarios/geracao/versao/{versao_id}", status_code=303)


@router.post("/horarios/geracao/versao/{versao_id}/excluir")
def geracao_excluir(request: Request, versao_id: int):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    v = conn.execute("SELECT ano_letivo, status FROM horario_versoes WHERE id = ?", (versao_id,)).fetchone()
    if v and v["status"] != "gerando":
        conn.execute("DELETE FROM horario_aulas WHERE versao_id = ?", (versao_id,))   # sem CASCADE: apaga as aulas antes
        conn.execute("DELETE FROM horario_versoes WHERE id = ?", (versao_id,))
        conn.commit()
    conn.close()
    return RedirectResponse(f"/horarios/geracao?ano={v['ano_letivo'] if v else ''}&msg=excluida", status_code=303)



# ---------- pendências e sugestões (apresentação no estilo do Urânia) ----------

def _quando(dia: int, turno: str, blocos: list, horas: dict) -> str:
    nome = NOME_DIA[DIAS_SEMANA[dia]] + "-feira"
    bl = ", ".join(f"{b}º" + (f" ({horas[(turno, b)]})" if horas.get((turno, b)) else "") for b in blocos)
    return f'{nome} à {"manhã" if turno == "m" else "tarde"}' + (f" — {bl} bloco" + ("s" if len(blocos) > 1 else "") if blocos else "")


def _html_sugestoes(sugestoes: list, horas: dict, ao_vivo: bool = False) -> str:
    """Texto da caixa de sugestões. Cada pendência = aulas que só couberam num turno que o
    professor marcou como indisponível."""
    if not sugestoes:
        return '<p>Nenhuma pendência. 🎉</p>'
    partes = []
    for s in sugestoes:
        nome = _e(s["professor_nome"])
        motivo = "não respondeu a coleta" if s.get("sem_resposta") else "está indisponível neste turno"
        linhas = [f'<div class="hs-titulo">ANÁLISE DA PENDÊNCIA {s.get("n", "")}</div>',
                  f'<div class="hs-pend">{_e(", ".join(s.get("aulas", [])))} · {nome} — só coube em {_e(_quando(s["dia"], s["turno"], s.get("blocos", []), horas))}.</div>',
                  '<div class="hs-nivel">Ajudará muitíssimo na eliminação da pendência se....</div>',
                  f'<div class="hs-item">• {nome} deixar disponível {_e(_quando(s["dia"], s["turno"], s.get("blocos", []), horas))} ({motivo})</div>']
        if s.get("muito"):
            linhas.append('<div class="hs-nivel">Ajudará muito na eliminação da pendência se....</div>')
            linhas += [f'<div class="hs-item">• {_e(a["professor_nome"])} deixar disponível {_e(_quando(a["dia"], a["turno"], a.get("blocos", []), horas))} '
                       f'(está indisponível neste turno)</div>' for a in s["muito"]]
        if s.get("pouco"):
            linhas.append('<div class="hs-nivel">Ajudará um pouco na eliminação da pendência se....</div>')
            linhas += [f'<div class="hs-item">• {_e(a["professor_nome"])} deixar disponível {_e(_quando(a["dia"], a["turno"], a.get("blocos", []), horas))} '
                       f'(está indisponível neste turno)</div>' for a in s["pouco"]]
        if not ao_vivo:
            linhas.append('<div class="hs-nivel">Se nenhuma sugestão acima for viável, verifique se é possível....</div>')
            if any(a.startswith("Educação Digital") for a in s.get("aulas", [])):
                linhas.append('<div class="hs-item">• Trocar o professor de Educação Digital deste turno, ou rever a disponibilidade dele</div>')
            else:
                for a in s.get("aulas", [])[:3]:
                    linhas.append(f'<div class="hs-item">• Atribuir {_e(a)} a outro professor em “Quem dá cada aula”</div>')
            linhas.append(f'<div class="hs-item">• Conferir as observações de {nome} na coleta (pode haver turno parcial combinado)</div>')
            if s.get("analisada") and not s.get("muito") and not s.get("pouco"):
                linhas.append('<div class="hs-obs">O gerador não encontrou outro turno que resolva esta pendência sem criar outras.</div>')
        partes.append('<div class="hs-bloco">' + "".join(linhas) + "</div>")
    return "".join(partes)


_CSS_SUGESTOES = """
<style>
.hs-painel { display:flex; flex-wrap:wrap; gap:28px; align-items:center; justify-content:center; padding:18px; }
.hs-cont { text-align:center; cursor:pointer; background:none; border:none; font-family:inherit; color:inherit; padding:4px 10px; border-radius:8px; }
.hs-cont:hover { background: var(--bg-subtle, #f1f5f9); }
.hs-cont b { display:block; font-size:13px; text-transform:uppercase; letter-spacing:.5px; color:var(--text-muted); }
.hs-cont span { font-family: ui-monospace, Menlo, Consolas, monospace; font-size:34px; font-weight:700; }
.hs-roda { width:56px; height:56px; border-radius:50%; position:relative; overflow:hidden; animation: hs-gira 1.6s linear infinite; }
.hs-roda i { position:absolute; width:50%; height:50%; }
.hs-roda i:nth-child(1) { left:0; top:0; background:#b5cf3c; } .hs-roda i:nth-child(2) { right:0; top:0; background:#4c9a48; }
.hs-roda i:nth-child(3) { left:0; bottom:0; background:#7cc24a; } .hs-roda i:nth-child(4) { right:0; bottom:0; background:#e5e7eb; }
.hs-roda.parada { animation:none; }
@keyframes hs-gira { to { transform: rotate(360deg); } }
dialog.hs-caixa { width:min(920px, 94vw); max-height:86vh; border:none; border-radius:10px; padding:0; box-shadow:0 20px 60px rgba(0,0,0,.35); }
dialog.hs-caixa::backdrop { background: rgba(15,23,42,.45); }
.hs-cab { background:#16a34a; color:white; padding:12px 18px; font-weight:700; font-size:15px; display:flex; justify-content:space-between; align-items:center; position:sticky; top:0; }
.hs-cab button { background:white; color:#166534; border:none; border-radius:6px; padding:4px 12px; font-weight:700; cursor:pointer; }
.hs-corpo { padding:14px 22px 20px; font-family: ui-monospace, Menlo, Consolas, monospace; font-size:13px; line-height:1.55; color:#0f172a; background:#fafafa; }
.hs-bloco { margin-bottom:18px; padding-bottom:12px; border-bottom:1px dashed #cbd5e1; }
.hs-titulo { display:inline-block; border:1px solid #64748b; padding:2px 14px; margin:6px 0 8px; font-weight:700; }
.hs-pend { margin-bottom:6px; } .hs-nivel { margin-top:8px; font-weight:700; } .hs-item { padding-left:12px; } .hs-obs { margin-top:6px; color:#64748b; }
</style>
"""


def _caixa_sugestoes_html(titulo: str, conteudo_html: str) -> str:
    return f"""
    <dialog id="hs-caixa" class="hs-caixa">
        <div class="hs-cab"><span>{_e(titulo)}</span><button type="button" onclick="document.getElementById('hs-caixa').close()">Fechar ✕</button></div>
        <div class="hs-corpo" id="hs-corpo">{conteudo_html}</div>
    </dialog>"""


@router.get("/horarios/geracao/versao/{versao_id}/status")
def geracao_status(request: Request, versao_id: int):
    """Números ao vivo para a tela de progresso (consultado a cada 2 segundos)."""
    from fastapi.responses import JSONResponse
    prof, negado = _exige_gestao(request)
    if negado:
        return JSONResponse({"erro": "acesso"}, status_code=403)
    conn = _db()
    v = conn.execute("SELECT status, ano_letivo, limite_segundos FROM horario_versoes WHERE id = ?", (versao_id,)).fetchone()
    horas = _horas_blocos(conn, v["ano_letivo"]) if v else {}
    conn.close()
    if not v:
        return JSONResponse({"status": "inexistente"})
    info = dict(_GERACOES.get(versao_id, {}))
    lista = info.get("pendencias_lista") or []
    sugest = [dict(x, n=i, muito=[], pouco=[]) for i, x in enumerate(lista, start=1)]
    return JSONResponse({
        "status": v["status"], "etapa": info.get("etapa", ""),
        "decorrido": int(_time.time() - info.get("inicio", _time.time())), "limite": info.get("limite", v["limite_segundos"] or 0),
        "flexivel": bool(info.get("flexivel")), "pendencias": info.get("pendencias"), "sugestoes": info.get("sugestoes"),
        "html": _html_sugestoes(sugest, horas, ao_vivo=True) if sugest else "",
    })


# ---------- tela: ver uma versão ----------

_CSS_GRADE = """
<style>
.hg-grade { border-collapse: collapse; width: 100%; font-size: 11px; table-layout: fixed; }
.hg-grade th { background: var(--bg-subtle, #f1f5f9); padding: 4px; font-size: 11px; }
.hg-grade td { border: 1px solid var(--border, #e2e8f0); padding: 4px 5px; vertical-align: top; height: 34px; }
.hg-grade td.hg-hora { width: 52px; font-weight: 600; color: var(--text-muted); text-align: center; vertical-align: middle; }
.hg-comp { font-weight: 700; font-size: 11.5px; line-height: 1.2; }
.hg-prof { color: var(--text-muted); font-size: 10.5px; }
.hg-ed { color: #0f766e; font-size: 10px; font-weight: 600; }
.hg-fora { background: #fef2f2 !important; outline: 2px solid #ef4444; outline-offset: -2px; }
.hg-cartoes { display: grid; grid-template-columns: repeat(auto-fill, minmax(420px, 1fr)); gap: 14px; }
.hg-cartao { border: 1px solid var(--border, #e2e8f0); border-radius: 8px; padding: 10px; break-inside: avoid; background: var(--bg, #fff); }
.hg-cartao h3 { margin: 0 0 6px; font-size: 14px; }
@media print {
  .sidebar, .mobile-topbar, .hg-nao-imprimir, .page-actions { display: none !important; }
  .main { padding: 0 !important; }
  .hg-cartoes { grid-template-columns: 1fr 1fr; gap: 8px; }
  .hg-cartao { page-break-inside: avoid; }
}
@media (max-width: 560px) { .hg-cartoes { grid-template-columns: 1fr; } }
</style>
"""


def _horas_blocos(conn, ano: int) -> dict:
    return {(r["turno"], r["bloco"]): (r["inicio"] or "") for r in conn.execute("SELECT * FROM horario_blocos WHERE ano_letivo = ?", (ano,)).fetchall()}


def _curto(nome: str) -> str:
    partes = (nome or "").split()
    return f"{partes[0]} {partes[-1]}" if len(partes) > 2 else (nome or "")


@router.get("/horarios/geracao/versao/{versao_id}", response_class=HTMLResponse)
def geracao_versao(request: Request, versao_id: int, ver: str = "turmas"):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    conn = _db()
    v = conn.execute("SELECT * FROM horario_versoes WHERE id = ?", (versao_id,)).fetchone()
    if not v:
        conn.close()
        return RedirectResponse("/horarios/geracao", status_code=303)
    ano_i = v["ano_letivo"]

    if v["status"] == "gerando":
        conn.close()
        conteudo = _cabecalho_config(f"⏳ Gerando o horário (versão #{versao_id})", "Pode sair desta página: a geração continua e a versão aparece na lista quando terminar.",
                                     "geracao", ano_i) + _CSS_SUGESTOES + f"""
            <div class="card" style="padding:18px; text-align:center;">
                <div class="hs-painel">
                    <button type="button" class="hs-cont" onclick="hsAbrir()" title="Clique para ver as pendências">
                        <b>Pendências</b><span id="hs-pend">--</span></button>
                    <div class="hs-roda" id="hs-roda"><i></i><i></i><i></i><i></i></div>
                    <button type="button" class="hs-cont" onclick="hsAbrir()" title="Clique para ver as sugestões">
                        <b>Sugestões</b><span id="hs-sug">--</span></button>
                </div>
                <div id="hs-etapa" style="font-size:14px; margin:4px 0 10px;">Preparando…</div>
                <div style="height:10px; background:var(--bg-subtle,#e2e8f0); border-radius:6px; overflow:hidden; max-width:520px; margin:0 auto;">
                    <div id="hs-barra" style="height:100%; width:0%; background:#16a34a; transition:width 1s;"></div></div>
                <div id="hs-tempo" style="font-size:12px; color:var(--text-muted); margin-top:6px;"></div>
                <button type="button" class="btn" style="margin-top:12px;" onclick="hsAbrir()">🔍 Ver pendências e sugestões</button>
            </div>
            {_caixa_sugestoes_html(f"Sugestões para pendências — Horário {ano_i} — versão #{versao_id}", '<p>Ainda não há pendências: o gerador está procurando um horário que respeite toda a disponibilidade.</p>')}
            <script>
            function hsAbrir() {{ var d = document.getElementById('hs-caixa'); if (d.showModal) d.showModal(); else d.setAttribute('open', ''); }}
            function hsPad(n) {{ return (n === null || n === undefined) ? '--' : String(n).padStart(2, '0'); }}
            function hsAtualiza() {{
              fetch('/horarios/geracao/versao/{versao_id}/status', {{credentials: 'same-origin'}}).then(function (r) {{ return r.json(); }}).then(function (j) {{
                if (j.status !== 'gerando') {{ location.reload(); return; }}
                document.getElementById('hs-etapa').textContent = j.etapa || 'Gerando…';
                document.getElementById('hs-pend').textContent = j.flexivel ? hsPad(j.pendencias) : '00';
                document.getElementById('hs-sug').textContent = j.flexivel ? hsPad(j.sugestoes) : '00';
                var pct = Math.min(99, Math.round(100 * j.decorrido / Math.max(1, j.limite)));
                document.getElementById('hs-barra').style.width = pct + '%';
                document.getElementById('hs-tempo').textContent = j.decorrido + 's de até ' + j.limite + 's nesta etapa';
                if (j.html) document.getElementById('hs-corpo').innerHTML = j.html;
              }}).catch(function () {{}}).finally(function () {{ setTimeout(hsAtualiza, 2000); }});
            }}
            hsAtualiza();
            </script>"""
        return _pagina(request, "Gerando horário", conteudo, "horarios-geracao")

    if v["status"] == "erro":
        conn.close()
        conteudo = _cabecalho_config(f"❌ Versão #{versao_id}", "", "geracao", ano_i) + _msg_box(_e(v["mensagem"] or "Erro."), erro=True) + \
            f'<a href="/horarios/geracao?ano={ano_i}" class="btn">← Voltar</a>'
        return _pagina(request, "Versão do horário", conteudo, "horarios-geracao")

    aulas = conn.execute("SELECT * FROM horario_aulas WHERE versao_id = ? ORDER BY turma_nome, dia, bloco", (versao_id,)).fetchall()
    horas = _horas_blocos(conn, ano_i)
    turmas_db = {t["id"]: t for t in _turmas(conn, ano_i)}
    conn.close()
    ind = _json.loads(v["indicadores"] or "{}")
    ed_info = _json.loads(v["pares_ed"] or "{}")
    fora = _json.loads(v["fora"] or "[]")
    ed_anos = {int(k): set(vv) for k, vv in _json.loads(v["ed_anos"] or "{}").items()}
    fora_set = {(f["professor_id"], f["dia"], f["turno"]) for f in fora}
    ano_da_turma = {}
    for a in aulas:
        t = turmas_db.get(a["turma_id"])
        ano_da_turma[a["turma_id"]] = t["ano_escolaridade"] if t else int(str(a["turma_nome"])[0]) if str(a["turma_nome"])[:1].isdigit() else 0

    def tem_ed(a):
        return ano_da_turma.get(a["turma_id"]) in ed_anos.get(a["componente_id"], set())

    # ---- topo: indicadores ----
    def card(valor, rotulo, cor="inherit"):
        return (f'<div class="card" style="padding:10px 14px; flex:1 1 150px;"><div style="font-size:24px; font-weight:800; color:{cor};">{valor}</div>'
                f'<div style="font-size:12px; color:var(--text-muted);">{rotulo}</div></div>')
    n_fora = ind.get("fora_disponibilidade", 0)
    cards = (card("✅" if v["modo"] == "estrito" else n_fora, "dentro da disponibilidade" if v["modo"] == "estrito" else "turnos fora da disponibilidade",
                  "#16a34a" if v["modo"] == "estrito" else "#b91c1c")
             + card(ind.get("janelas", 0), "janelas de professor")
             + card(ind.get("turnos_usados", 0), "turnos de presença (somando todos os professores)"))
    pares_txt = " · ".join(
        f'{NOME_TURNO[k]}: semana 1 {"+".join(str(a) + "º" for a in info["pares"][0])}, semana 2 {"+".join(str(a) + "º" for a in info["pares"][1])}'
        for k, info in sorted(ed_info.items()))
    bloco_fora = ""
    sugestoes = _json.loads(v["sugestoes"] or "[]") if "sugestoes" in v.keys() else []
    if fora:
        n_pend = ind.get("pendencias", len(fora))
        n_sug = ind.get("sugestoes", len(sugestoes))
        bloco_fora = _CSS_SUGESTOES + f"""
            <div class="card" style="padding:6px 16px 14px; margin-bottom:14px; border:2px solid #fcd34d;">
                <div class="hs-painel" style="padding:10px;">
                    <button type="button" class="hs-cont" onclick="hsAbrir()"><b>Pendências</b><span>{n_pend:02d}</span></button>
                    <div class="hs-roda parada"><i></i><i></i><i></i><i></i></div>
                    <button type="button" class="hs-cont" onclick="hsAbrir()"><b>Sugestões</b><span>{n_sug:02d}</span></button>
                </div>
                <div style="font-size:13px; text-align:center;">
                    <strong>⚠️ Não existe horário dentro da disponibilidade informada.</strong> Este é o horário com o menor número de pendências
                    (aulas em turnos que o professor marcou como indisponível) — elas aparecem com contorno vermelho.<br>
                    <button type="button" class="btn btn-primary" style="margin-top:10px;" onclick="hsAbrir()">🔍 Ver pendências e sugestões</button>
                    <span style="font-size:12px; color:var(--text-muted); display:block; margin-top:6px;">Combine com os professores, ajuste no
                    <a href="/horarios/disponibilidade/painel">painel da coleta</a> e gere de novo.</span>
                </div>
            </div>
            {_caixa_sugestoes_html(f"Sugestões para pendências — Horário {ano_i} — versão #{versao_id}", _html_sugestoes(sugestoes, horas) if sugestoes else
               "".join(f'<div class="hs-item">• {_e(f["professor_nome"])} deixar disponível {_e(_quando(f["dia"], f["turno"], [], horas))}</div>' for f in fora))}
            <script>function hsAbrir() {{ var d = document.getElementById('hs-caixa'); if (d.showModal) d.showModal(); else d.setAttribute('open', ''); }}</script>"""

    abas = "".join(
        f'<a href="/horarios/geracao/versao/{versao_id}?ver={k}" class="btn" style="margin:0 6px 6px 0; {"background:#0f766e; color:white; border-color:#0f766e;" if ver == k else ""}">{r}</a>'
        for k, r in (("turmas", "👥 Por turma"), ("professores", "👩‍🏫 Por professor"), ("laboratorio", "💻 Laboratório")))

    def grade(celulas, turno):
        """celulas[(dia, bloco)] = html. Grade 3 blocos × 5 dias."""
        cab = "".join(f"<th>{NOME_DIA[d][:3]}</th>" for d in DIAS_SEMANA)
        linhas = ""
        for b in range(1, BLOCOS_POR_TURNO + 1):
            tds = "".join(celulas.get((di, b), '<td></td>') for di in range(len(DIAS_SEMANA)))
            linhas += f'<tr><td class="hg-hora">{_e(horas.get((turno, b), "")) or str(b) + "º"}</td>{tds}</tr>'
        return f'<table class="hg-grade"><tr><th style="width:52px;"></th>{cab}</tr>{linhas}</table>'

    corpo = ""
    if ver == "professores":
        por_prof = {}
        for a in aulas:
            por_prof.setdefault((a["professor_nome"], a["professor_id"]), []).append(a)
        ed_por_prof = {}   # professor de ED: em cada bloco, até 2 turmas (uma por semana da quinzena)
        for k, info in ed_info.items():
            if info.get("professor_id"):
                chave_p = (info["professor_nome"], info["professor_id"])
                por_prof.setdefault(chave_p, [])
                for a in aulas:
                    if a["turno"] == k and tem_ed(a):
                        semana = 1 if ano_da_turma.get(a["turma_id"]) in info["pares"][0] else 2
                        ed_por_prof.setdefault(chave_p, {}).setdefault((k, a["dia"], a["bloco"]), []).append((semana, a["turma_nome"]))
        cartoes = ""
        for (nome, pid), lista in sorted(por_prof.items(), key=lambda x: (x[0][0] or "").lower()):
            eds = ed_por_prof.get((nome, pid), {})
            turnos = sorted({a["turno"] for a in lista} | {k for k, _, _ in eds}, key=lambda k: 0 if k == "m" else 1)
            grades = ""
            for k in turnos:
                cel = {}
                for a in [a for a in lista if a["turno"] == k]:
                    classe = ' class="hg-fora"' if (pid, a["dia"], k) in fora_set else ""
                    cel[(a["dia"], a["bloco"])] = f'<td{classe}><div class="hg-comp">{_e(a["turma_nome"])}</div><div class="hg-prof">{_e(a["componente_nome"])}</div></td>'
                for (kk, d, b), turmas_ed in eds.items():
                    if kk != k:
                        continue
                    classe = ' class="hg-fora"' if (pid, d, k) in fora_set else ""
                    txt = "<br>".join(f"S{sem}: {_e(tn)}" for sem, tn in sorted(turmas_ed))
                    cel[(d, b)] = f'<td{classe}><div class="hg-comp">💻 {txt}</div></td>'
                grades += f'<div style="font-size:11px; font-weight:700; margin:6px 0 2px;">{NOME_TURNO[k]}</div>' + grade(cel, k)
            n_blocos = len({(a["turno"], a["dia"], a["bloco"]) for a in lista} | set(eds.keys()))
            cartoes += f'<div class="hg-cartao"><h3>{_e(nome)} <span style="font-weight:400; font-size:12px; color:var(--text-muted);">· {n_blocos} blocos</span></h3>{grades}</div>'
        corpo = f'<div class="hg-cartoes">{cartoes}</div>'
    elif ver == "laboratorio":
        blocos_lab = ""
        for k, info in sorted(ed_info.items(), key=lambda x: 0 if x[0] == "m" else 1):
            for semana, par in ((1, info["pares"][0]), (2, info["pares"][1])):
                cel = {}
                for a in aulas:
                    if a["turno"] == k and tem_ed(a) and ano_da_turma.get(a["turma_id"]) in par:
                        cel[(a["dia"], a["bloco"])] = f'<td><div class="hg-comp">{_e(a["turma_nome"])}</div><div class="hg-prof">{_e(a["componente_nome"])}</div></td>'
                titulo = f'{NOME_TURNO[k]} — semana {semana}: {" e ".join(str(a) + "º" for a in par)} ano'
                blocos_lab += (f'<div class="hg-cartao"><h3>{titulo}</h3><div style="font-size:11px; color:var(--text-muted); margin-bottom:4px;">'
                               f'Professor(a): {_e(info.get("professor_nome") or "—")} · {len(cel)} blocos</div>{grade(cel, k)}</div>')
        corpo = (f'<p style="font-size:12px; color:var(--text-muted);" class="hg-nao-imprimir">O Laboratório recebe uma turma por bloco. A cada semana, um par de anos. '
                 f'Divisão escolhida pelo gerador — {_e(pares_txt) or "sem Educação Digital"}.</p><div class="hg-cartoes">{blocos_lab}</div>')
    else:
        por_turma = {}
        for a in aulas:
            por_turma.setdefault((a["turno"], a["turma_nome"]), []).append(a)
        cartoes = ""
        for (k, nome), lista in sorted(por_turma.items(), key=lambda x: (0 if x[0][0] == "m" else 1, x[0][1])):
            cel = {}
            for a in lista:
                classe = ' class="hg-fora"' if (a["professor_id"], a["dia"], k) in fora_set else ""
                ed_txt = '<div class="hg-ed">💻 + Educação Digital</div>' if tem_ed(a) else ""
                cel[(a["dia"], a["bloco"])] = (f'<td{classe}><div class="hg-comp">{_e(a["componente_nome"])}</div>'
                                               f'<div class="hg-prof">{_e(_curto(a["professor_nome"]))}</div>{ed_txt}</td>')
            cartoes += f'<div class="hg-cartao"><h3>Turma {_e(nome)} <span style="font-weight:400; font-size:12px; color:var(--text-muted);">· {NOME_TURNO[k]}</span></h3>{grade(cel, k)}</div>'
        corpo = f'<div class="hg-cartoes">{cartoes}</div>'

    if ind.get("sem_resposta"):
        bloco_fora += (f'<div class="hd-avisos" style="display:block; margin-bottom:14px;">ℹ️ Sem resposta na coleta (considerados disponíveis em todos os turnos): '
                       f'<strong>{_e(", ".join(ind["sem_resposta"]))}</strong>. Confira o horário deles antes de usar.</div>')
    modo_txt = "respeitando toda a disponibilidade" if v["modo"] == "estrito" else "modo diagnóstico (disponibilidade flexível)"
    conteudo = (_CSS_FORM + _CSS_GRADE + f"""
        <div class="page-header">
            <h1>🗓️ Horário {ano_i} — versão #{versao_id}</h1>
            <p class="subtitle">Gerado em {_ts_br(v["criada_em"])} em {v["tempo"] or "?"}s, {modo_txt}. Educação Digital — {_e(pares_txt) or "—"}.</p>
        </div>
        <div class="hg-nao-imprimir">
            <div style="display:flex; flex-wrap:wrap; gap:10px; margin-bottom:14px;">{cards}</div>
            {bloco_fora}
            <div style="display:flex; flex-wrap:wrap; align-items:center; margin-bottom:10px;">{abas}
                <span style="flex:1;"></span>
                <a href="/horarios/geracao/versao/{versao_id}/exportar" class="btn" style="margin:0 6px 6px 0;">📥 Excel</a>
                <button type="button" class="btn" style="margin:0 6px 6px 0;" onclick="window.print()">🖨️ Imprimir</button>
                <a href="/horarios/geracao?ano={ano_i}" class="btn" style="margin:0 0 6px 0;">← Versões</a>
            </div>
        </div>
        {corpo}""")
    return _pagina(request, f"Horário {ano_i} — versão {versao_id}", conteudo, "horarios-geracao")


@router.get("/horarios/geracao/versao/{versao_id}/exportar")
def geracao_exportar(request: Request, versao_id: int):
    prof, negado = _exige_gestao(request)
    if negado:
        return negado
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    conn = _db()
    v = conn.execute("SELECT * FROM horario_versoes WHERE id = ? AND status = 'pronta'", (versao_id,)).fetchone()
    if not v:
        conn.close()
        return RedirectResponse("/horarios/geracao", status_code=303)
    aulas = conn.execute("SELECT * FROM horario_aulas WHERE versao_id = ?", (versao_id,)).fetchall()
    horas = _horas_blocos(conn, v["ano_letivo"])
    conn.close()
    wb = Workbook()
    negrito = Font(bold=True, color="FFFFFF")
    fundo = PatternFill("solid", fgColor="0F766E")
    fina = Side(style="thin", color="CBD5E1")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)
    quebra = Alignment(wrap_text=True, vertical="top")

    def aba_grade(titulo, grupos, texto_celula):
        ws = wb.create_sheet(titulo)
        linha = 1
        for (k, nome), lista in grupos:
            ws.cell(linha, 1, f"{nome} — {NOME_TURNO[k]}").font = Font(bold=True, size=12)
            linha += 1
            ws.cell(linha, 1, "")
            for i, d in enumerate(DIAS_SEMANA):
                c = ws.cell(linha, i + 2, NOME_DIA[d]); c.font, c.fill = negrito, fundo
            linha += 1
            for b in range(1, BLOCOS_POR_TURNO + 1):
                ws.cell(linha, 1, horas.get((k, b), "") or f"{b}º bloco").font = Font(bold=True)
                for i in range(len(DIAS_SEMANA)):
                    txt = "\n".join(texto_celula(a) for a in lista if a["dia"] == i and a["bloco"] == b)
                    c = ws.cell(linha, i + 2, txt); c.alignment, c.border = quebra, borda
                ws.row_dimensions[linha].height = 32
                linha += 1
            linha += 1
        ws.column_dimensions["A"].width = 12
        for i in range(len(DIAS_SEMANA)):
            ws.column_dimensions[chr(66 + i)].width = 26

    por_turma = {}
    for a in aulas:
        por_turma.setdefault((a["turno"], a["turma_nome"]), []).append(a)
    aba_grade("Por turma", sorted(por_turma.items(), key=lambda x: (0 if x[0][0] == "m" else 1, x[0][1])),
              lambda a: f'{a["componente_nome"]} ({_curto(a["professor_nome"])})')
    por_prof = {}
    for a in aulas:
        por_prof.setdefault(a["professor_nome"], []).append(a)
    grupos_prof = []
    for nome, lista in sorted(por_prof.items(), key=lambda x: (x[0] or "").lower()):
        for k in ("m", "t"):
            sub = [a for a in lista if a["turno"] == k]
            if sub:
                grupos_prof.append(((k, nome), sub))
    aba_grade("Por professor", grupos_prof, lambda a: f'{a["turma_nome"]} — {a["componente_nome"]}')
    ws = wb.create_sheet("Lista")
    ws.append(["Turma", "Turno", "Dia", "Bloco", "Componente", "Professor"])
    for a in sorted(aulas, key=lambda a: (a["turma_nome"], a["dia"], a["bloco"])):
        ws.append([a["turma_nome"], NOME_TURNO[a["turno"]], NOME_DIA[DIAS_SEMANA[a["dia"]]], a["bloco"], a["componente_nome"], a["professor_nome"]])
    for c in ws[1]:
        c.font, c.fill = negrito, fundo
    del wb["Sheet"]
    saida = BytesIO()
    wb.save(saida)
    saida.seek(0)
    return StreamingResponse(saida, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f'attachment; filename="horario_{v["ano_letivo"]}_versao_{versao_id}.xlsx"'})


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
        itens += nav_item("/horarios/configuracao", "horarios-config", "⚙️", "Configuração do horário")
        itens += nav_item("/horarios/geracao", "horarios-geracao", "🚀", "Gerar horário")
    return '<div class="sidebar-section">Horário Escolar</div>' + itens


def grupo_atalhos(prof: Optional[dict]):
    """Grupo de atalhos para a tela inicial (mesmo formato dos outros grupos)."""
    if not _pode_responder(prof):
        return None
    return ("Horário Escolar", "#0f766e", [
        ("🗓️", "Minha disponibilidade", "/horarios/disponibilidade", True),
        ("📊", "Painel de disponibilidade", "/horarios/disponibilidade/painel", _eh_gestao(prof)),
        ("⚙️", "Configuração do horário", "/horarios/configuracao", _eh_gestao(prof)),
        ("🚀", "Gerar horário", "/horarios/geracao", _eh_gestao(prof)),
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