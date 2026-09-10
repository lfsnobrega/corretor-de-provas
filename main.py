from fastapi import FastAPI, Form, UploadFile, File, Request, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional, List
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # Adicionado estilos
from openpyxl.utils import get_column_letter                           # Adicionado utilitário
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
import sqlite3
import os
import re
import uuid
import asyncio
import time
import math
from collections import deque
import qrcode
import base64
import html
import secrets
import hashlib
import json
import urllib.parse
import unicodedata

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


GOOGLE_DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
# Pasta separada pros atestados médicos de ALUNOS (já compartilhada com a conta de
# serviço da VM, confirmado por Felipe em 27/08/2026) — pode ser sobrescrita por
# variável de ambiente se a pasta mudar no futuro.
GOOGLE_DRIVE_FOLDER_ID_ALUNOS = os.environ.get("GOOGLE_DRIVE_FOLDER_ID_ALUNOS", "1uLJrM4BiCiGnBkLHPI1f5SuknK24f9sL")
GOOGLE_DRIVE_CREDENTIALS_JSON = os.environ.get("GOOGLE_DRIVE_CREDENTIALS_JSON", "")

# Carômetro — link do Canva por turma (27/08/2026). Chave é o "nome" da turma como
# cadastrado no sistema (ex: "601"), valor é o link de edição do Canva.
CAROMETRO_LINKS = {
    "601": "https://www.canva.com/design/DAHEN-ksHO8/9Ykhjf1LGyyPnZZl1MdXig/edit",
    "602": "https://www.canva.com/design/DAHEOcTmzFg/zfv-NRhcFJGNJ34LzNXYCw/edit",
    "603": "https://www.canva.com/design/DAHFEfYfRyc/cspiSRDtFhyPB9mNlYRciQ/edit",
    "604": "https://www.canva.com/design/DAHFEjXeXe4/F48EAv74k1MQcTJYei4etA/edit",
    "605": "https://www.canva.com/design/DAHFEtGHDZ8/eO-56qic40iGCwmDJP7MSA/edit",
    "701": "https://www.canva.com/design/DAHEUIxdAs4/z_prhoP1-IlXYoGHfne-eg/edit",
    "702": "https://www.canva.com/design/DAHEaeggrrk/Eqdi8aLJoZMiOv28atmkRw/edit",
    "703": "https://www.canva.com/design/DAHIUidV7KU/vfEeF87Yy31MAZpRsF9vIw/edit",
    "704": "https://www.canva.com/design/DAHFEzqBCFM/3gD1U23C9X76eeA9uaA5hQ/edit",
    "705": "https://www.canva.com/design/DAHFE2T_vNo/jdNRTyK_kcQmc_qBYV0r6Q/edit",
    "706": "https://www.canva.com/design/DAHFKEgR98A/mTY7EfaRYobnaevCHW9R-Q/edit",
    "801": "https://www.canva.com/design/DAHIuXourb0/aFZN75HuC2k4W0I0U9y2Sw/edit",
    "802": "https://www.canva.com/design/DAHEau91N0w/j9-dj3Xyqr9Lyvi3pYlbAg/edit",
    "803": "https://www.canva.com/design/DAHE4GBygM4/zBZjLgRg1Bagvk3yNSQplg/edit",
    "804": "https://www.canva.com/design/DAHFKHbOhG8/dOUjrABAI9syhR51-pLHXA/edit",
    "805": "https://www.canva.com/design/DAHFKIhCLT4/3AMLtREHd0TOb3FACiufJQ/edit",
    "806": "https://www.canva.com/design/DAHFKSTLDtw/ReXpbS2WNRdQkh3-070d_A/edit",
    "901": "https://www.canva.com/design/DAHE4AwrbtA/DdoUEWVqY458entcNtcFaQ/edit",
    "902": "https://www.canva.com/design/DAHE4OalXbE/DFNYtX9dYX6a7UsVzyM1DQ/edit",
    "903": "https://www.canva.com/design/DAHE4DxyszI/0R8spuPXuwRmZBsMu4Vt-A/edit",
    "904": "https://www.canva.com/design/DAHFEVSy02A/Uhsae-hV4zkHRiI3Vg5OPQ/edit",
    "905": "https://www.canva.com/design/DAHFKbtUt_Q/Y9gnyo1QrSiKU40UnQEvSQ/edit",
    "906": "https://www.canva.com/design/DAHFKqjvA7g/nDlp_bUFmU6miVmYWvv2qA/edit",
}


# Tipos de ocorrência disciplinar (Orientação Educacional) — 28/08/2026. Lista fechada
# pra estatística fazer sentido; ajustável depois se a escola usar outra categorização.
TIPOS_OCORRENCIA = {
    "indisciplina_sala": "Indisciplina em sala de aula",
    "uso_celular": "Uso indevido de celular",
    "desrespeito_colega": "Desrespeito a colega",
    "desrespeito_funcionario": "Desrespeito a professor/funcionário",
    "agressao_fisica": "Agressão física",
    "agressao_verbal": "Agressão verbal",
    "dano_patrimonio": "Dano ao patrimônio escolar",
    "atraso_recorrente": "Atraso recorrente",
    "evasao_sala": "Evasão/fuga da sala",
    "outro": "Outro",
}

# Tipos de ocorrência em que existe uma SEGUNDA pessoa (aluno) diretamente envolvida —
# o formulário abre um campo extra pro nome dela (02/09/2026, a pedido).
TIPOS_COM_OUTRO_ALUNO = {"agressao_fisica"}

# Natureza do ato (Art. 74 do Regimento — "Do Regime Disciplinar Aplicado ao Corpo
# Discente", documento anexado por Felipe em 02/09/2026) e as medidas educativas
# disciplinares correspondentes (Art. 80), pra manter o registro alinhado ao regimento.
NATUREZA_ATO = {
    "leve": "Leve (Art. 75)",
    "medio": "Médio (Art. 76)",
    "grave": "Grave (Art. 77)",
    "infracional": "Ato infracional (Art. 78/79)",
}
MEDIDAS_DISCIPLINARES = {
    "leve": [
        "Advertência verbal",
        "Advertência por escrito",
        "Repreensão, com encaminhamento à diretoria/coordenação",
        "Suspensão das atividades recreativas",
    ],
    "medio": [
        "Advertência por escrito",
        "Repreensão, com encaminhamento à diretoria/coordenação",
        "Suspensão temporária de participação em programas extracurriculares",
        "Suspensão das aulas por 1 dia letivo (em domicílio, com atividades pedagógicas)",
        "Mudança de turma",
    ],
    "grave": [
        "Repreensão, com encaminhamento à diretoria/coordenação",
        "Suspensão das aulas por 2 dias letivos (em domicílio, com atividades pedagógicas)",
        "Mudança de turno",
    ],
    "infracional": [
        "Encaminhamento ao Conselho Tutelar (até 12 anos incompletos)",
        "Registro de Boletim de Ocorrência (12 a 17 anos, com comunicação ao Conselho Tutelar)",
        "Registro de Boletim de Ocorrência (a partir de 18 anos)",
        "Suspensão das aulas por 3 dias letivos (em domicílio, com atividades pedagógicas)",
        "Transferência para outra unidade escolar (via Conselho de Classe extraordinário)",
    ],
}

# Motivos de saída da escola (06/09/2026) — usado quando um aluno sai de vez (transfere
# pra outra escola, muda de cidade, evade etc.), sem excluir o cadastro nem o histórico.
MOTIVOS_SAIDA_ALUNO = [
    "Transferência para outra escola",
    "Mudança de cidade/estado",
    "Evasão escolar",
    "Falecimento",
    "Outro",
]


TIPOS_AFASTAMENTO = {
    "atestado_medico": "Atestado médico",
    "permissao_ausencia": "Permissão de ausência",
    "abono_1_3": "Abono 1/3",
    "abono_2_3": "Abono 2/3",
    "folga_tre": "Folga TRE",
    "abono_integral": "Abono Integral",
    "declaracao_comparecimento": "Declaração de comparecimento",
    "hora_extra": "Hora extra",
    "compensacao_horas": "Compensação de horas",
}
# Tipos que precisam de horário específico (não é o dia inteiro) — 26/08/2026, ampliado
# em 02/09/2026 pra Hora extra e Compensação de horas (precisam do intervalo de horas).
TIPOS_COM_HORARIO = {"permissao_ausencia", "declaracao_comparecimento", "hora_extra", "compensacao_horas"}


def _fmt_data_br(data_iso: str) -> str:
    """Converte data ISO (AAAA-MM-DD, como fica salva no banco) pro formato brasileiro
    DD/MM/AAAA, só pra exibição — os campos <input type=date> continuam recebendo ISO
    normalmente, essa função é só pra texto mostrado na tela (28/08/2026)."""
    if not data_iso:
        return "—"
    try:
        return date.fromisoformat(data_iso).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return data_iso


def _extrair_matricula(email: str) -> str:
    """Extrai a matrícula do e-mail de cadastro do profissional. Formato usado na rede:
    nome.MATRICULA@smevr.com.br — pega o último trecho separado por ponto, se for só
    dígitos. Se não achar (email fora do padrão), retorna '—' (25/08/2026)."""
    if not email or "@" not in email:
        return "—"
    local = email.split("@")[0]
    partes = local.split(".")
    ultima = partes[-1] if partes else ""
    return ultima if ultima.isdigit() else "—"


def _matricula_professor(prof_row) -> str:
    """Matrícula de exibição pra relatórios: prioriza o campo explícito 'matricula'
    (obrigatório pra contas locais sem e-mail institucional) e só cai pra extração do
    e-mail se esse campo estiver vazio — 27/08/2026."""
    matricula_explicita = prof_row["matricula"] if "matricula" in prof_row.keys() else None
    if matricula_explicita:
        return matricula_explicita
    return _extrair_matricula(prof_row["email"])


def _hash_senha(senha: str, salt: Optional[str] = None) -> tuple:
    """Gera hash de senha com PBKDF2-SHA256 (biblioteca padrão do Python, sem dependência
    nova) — 100.000 iterações, salt aleatório de 16 bytes. Retorna (hash_hex, salt_hex).
    Se salt for passado, reusa (pra verificar uma senha existente) — 27/08/2026."""
    if salt is None:
        salt = secrets.token_hex(16)
    hash_bytes = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), bytes.fromhex(salt), 100_000)
    return hash_bytes.hex(), salt


def _verificar_senha(senha: str, hash_salvo: str, salt: str) -> bool:
    hash_calculado, _ = _hash_senha(senha, salt)
    return secrets.compare_digest(hash_calculado, hash_salvo or "")


def _gerar_senha_temporaria(tamanho: int = 8) -> str:
    """Gera senha provisória legível (sem caracteres ambíguos tipo 0/O, 1/l) — 27/08/2026."""
    alfabeto = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789"
    return "".join(secrets.choice(alfabeto) for _ in range(tamanho))


# Referencial Curricular de Arte — extraído do documento oficial enviado por Felipe em
# 04/09/2026 (RF_Arte.pdf). Cada bloco: ano + trimestre + unidade temática + objetos de
# conhecimento (texto livre, um item por linha) + códigos de habilidades BNCC (ligados à
# tabela habilidades_bncc que já existe no sistema).
REFERENCIAL_CURRICULAR_ARTE = [
    {"ano": "6º ano", "trimestre": 1, "unidade": "Arte", "objetos":
        "O que é Arte?\nElementos das Artes Visuais: Ponto; Linha; Cores e suas Classificações em primárias e secundárias, quentes, frias e neutras\nFigurativo e Abstrato\nArte da Pré-História\nArte Indígena",
        "habilidades": ["EF69AR01", "EF69AR02", "EF69AR04", "EF69AR05", "EF69AR06", "EF69AR08", "EF69AR33"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Dança e Música", "objetos":
        "Dança e sua origem\nDança de origem Indígena\nMovimento da Dança: Espaço – Fluência; Peso – Tempo\nMúsica Popular e Música Erudita\nElementos Musicais: Intensidade; Melodia; Harmonia; Ritmo\nSurgimento da MPB – Chorinho",
        "habilidades": ["EF69AR10", "EF69AR11", "EF69AR13", "EF69AR14", "EF69AR16", "EF69AR17", "EF69AR18", "EF69AR19", "EF69AR20", "EF69AR31", "EF69AR34"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Teatro", "objetos":
        "Teatro e sua origem\nHistória do Teatro no Brasil\nElementos do Teatro: Plateia; Personagens (Protagonista, Antagonista, Coadjuvante); Figurino - Maquiagem - Cenário",
        "habilidades": ["EF69AR24", "EF69AR25", "EF69AR26", "EF69AR34"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Arte", "objetos":
        "O que é Arte?\nElementos de Linguagem das Artes Visuais\nCultura afro e sua influência\nArte Popular e Arte Erudita\nMovimentos Artísticos: Arte Bizantina, Arte Românica, Arte Gótica",
        "habilidades": ["EF69AR01", "EF69AR02", "EF69AR03", "EF69AR04", "EF69AR06", "EF69AR07", "EF69AR08", "EF69AR31", "EF69AR34"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Dança e Música", "objetos":
        "Dança e sua origem\nDança Clássica e Dança Moderna\nMúsica Popular e Música Erudita\nElementos Musicais: Intensidade – Melodia – Harmonia – Ritmo\nSagrado no Canto e Surgimento das Notas Musicais\nMPB – Músicas Regionais",
        "habilidades": ["EF69AR09", "EF69AR10", "EF69AR13", "EF69AR15", "EF69AR16", "EF69AR17", "EF69AR18", "EF69AR19", "EF69AR20", "EF69AR21", "EF69AR22", "EF69AR31", "EF69AR34"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Artes Visuais, Teatro e Artes Integradas", "objetos":
        "Movimento Artístico – Renascimento\nArtistas Renascentistas\nTeatro e sua origem\nHistória do Teatro no Brasil\nElementos do Teatro: Plateia, Personagem (Protagonista, Antagonista, Coadjuvante), Figurino, Maquiagem, Cenário\nAs Máscaras no teatro Grego e Romano\nTeatro Renascentista: A Commedia dell'Arte e as Máscaras\nPersonagens",
        "habilidades": ["EF69AR01", "EF69AR02", "EF69AR24", "EF69AR25", "EF69AR26", "EF69AR30", "EF69AR32", "EF69AR33"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Arte", "objetos":
        "Movimentos Artísticos: Arte Barroca, Barroco no Brasil, Neoclássico, Missão Artística Francesa\nElementos de Linguagens Visuais: Ponto/Linha; Suporte/Volume; Espaço/Movimento; Simetria/Assimetria; Luz/Sombra\nMonocromia e Policromia",
        "habilidades": ["EF69AR01", "EF69AR02", "EF69AR04", "EF69AR05", "EF69AR06", "EF69AR07", "EF69AR31"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Dança e Música", "objetos":
        "Dança Contemporânea\nGrupos de Dança Nacionais e Internacionais\nDança - Flash Mobs\nInfluência Afro na MPB\nInstrumentos musicais de origem afro\nMúsica – Samba\nPersonalidades Negras de destaque na cultura brasileira",
        "habilidades": ["EF69AR09", "EF69AR10", "EF69AR12", "EF69AR13", "EF69AR14", "EF69AR15", "EF69AR17", "EF69AR18", "EF69AR19", "EF69AR21", "EF69AR31", "EF69AR34", "EF69AR35"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Teatro e Artes Integradas", "objetos":
        "Surgimento do Teatro (Grécia/Roma)\nTeatro Moderno\nTeatro Contemporâneo\nTeatro no Brasil na Atualidade\nA diversidade das técnicas formadas: Formas animadas\nComo e por que se faz teatro\nPara quem e onde se faz teatro\nTexto Teatral",
        "habilidades": ["EF69AR24", "EF69AR25", "EF69AR27", "EF69AR28", "EF69AR30", "EF69AR33", "EF69AR34"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Arte", "objetos":
        "Movimentos Artísticos: Impressionismo, Expressionismo, Cubismo, Surrealismo, Pop Arte\nArte Moderna no Brasil (Semana de Arte Moderna)\nArtistas Modernistas Brasileiros\nArte Pós-Moderna\nArtes digitais",
        "habilidades": ["EF69AR01", "EF69AR02", "EF69AR03", "EF69AR05", "EF69AR06", "EF69AR07", "EF69AR08"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Dança e Música", "objetos":
        "Dança Clássica, Moderna e Contemporânea\nDança de Rua (Street Dance)\nRock e Rock Nacional\nMúsica Popular Brasileira: Bossa Nova\nHip Hop - Cultura de Rua - Rap - Funk",
        "habilidades": ["EF69AR09", "EF69AR10", "EF69AR12", "EF69AR13", "EF69AR14", "EF69AR15", "EF69AR16", "EF69AR17", "EF69AR18", "EF69AR19", "EF69AR23", "EF69AR31", "EF69AR32", "EF69AR34"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Teatro e Artes Integradas", "objetos":
        "Gêneros Teatrais: Comédia, Tragédia, Tragicomédia (Sátira)\nTeatro Contemporâneo\nTeatro Musical",
        "habilidades": ["EF69AR24", "EF69AR25", "EF69AR26", "EF69AR27", "EF69AR28", "EF69AR29", "EF69AR35"]},
]

# Referencial Curricular de Educação Física — extraído de RF_Educação_Física.pdf
# (enviado por Felipe em 04/09/2026). Quando duas unidades "Esporte" apareciam no mesmo
# trimestre com objetos diferentes (ex: Invasão + Técnico Combinatório), foram unificadas
# numa linha só, já que compartilhavam exatamente as mesmas habilidades BNCC.
REFERENCIAL_CURRICULAR_EDUCACAO_FISICA = [
    {"ano": "6º ano", "trimestre": 1, "unidade": "Esporte", "objetos":
        "Esporte de Marca: Atletismo (corridas de velocidade, revezamento, saltos)",
        "habilidades": ["EF67EF03", "EF67EF04"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Brincadeiras e Jogos", "objetos":
        "Jogos Eletrônicos: Games, brincadeiras e jogos eletrônicos adaptados",
        "habilidades": ["EF67EF01"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Ginástica", "objetos":
        "Ginástica de Condicionamento Físico: Valências físicas e Funcional",
        "habilidades": ["EF67EF08", "EF67EF10"]},

    {"ano": "6º ano", "trimestre": 2, "unidade": "Esporte", "objetos":
        "Esporte de Invasão: Futsal e Handebol\nEsporte Técnico Combinatório: Artística, rítmica, acrobática e circense",
        "habilidades": ["EF67EF03", "EF67EF04"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Danças", "objetos":
        "Danças Urbanas: Funk",
        "habilidades": ["EF67EF11"]},

    {"ano": "6º ano", "trimestre": 3, "unidade": "Práticas Corporais de Aventura", "objetos":
        "PCA Urbanas: Skate, Parkour e Slackline",
        "habilidades": ["EF67EF18", "EF67EF19"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Lutas", "objetos":
        "Lutas do Brasil: Capoeira e Huka Huka",
        "habilidades": ["EF67EF14", "EF67EF16"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Esporte de Precisão", "objetos":
        "Esporte de Precisão: Boliche, zarabatana, elástico de dedo",
        "habilidades": ["EF67EF03", "EF67EF04"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Esporte", "objetos":
        "Esporte de Marca (Revisão do ano anterior, corridas de resistência, arremesso e lançamentos): Corrida de rua",
        "habilidades": ["EF67EF03", "EF67EF05", "EF67EF06", "EF67EF07"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Brincadeiras e Jogos", "objetos":
        "Jogos Eletrônicos: Games",
        "habilidades": ["EF67EF01", "EF67EF02"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Ginástica", "objetos":
        "Ginástica de Condicionamento Físico: Funcional",
        "habilidades": ["EF67EF08", "EF67EF09", "EF67EF10"]},

    {"ano": "7º ano", "trimestre": 2, "unidade": "Esporte", "objetos":
        "Esporte de Invasão: Futsal e Handebol\nEsporte Técnico Combinatório: Artística, rítmica, acrobática e circense",
        "habilidades": ["EF67EF05", "EF67EF06", "EF67EF07"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Danças", "objetos":
        "Danças Urbanas: Funk, break, vogue, popping, house dance",
        "habilidades": ["EF67EF11", "EF67EF12", "EF67EF13"]},

    {"ano": "7º ano", "trimestre": 3, "unidade": "Práticas Corporais de Aventura", "objetos":
        "PCA Urbanas: Skate, Parkour e Slackline",
        "habilidades": ["EF67EF20", "EF67EF21"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Lutas", "objetos":
        "Lutas do Brasil: Capoeira, Marajoara, Jiu Jitsu brasileiro",
        "habilidades": ["EF67EF14", "EF67EF15", "EF67EF17"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Esporte de Precisão", "objetos":
        "Esporte de Precisão: Dardo e Golfe",
        "habilidades": ["EF67EF05", "EF67EF06", "EF67EF07"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Esporte", "objetos":
        "Esporte de Rede e Parede: Voleibol e Badminton",
        "habilidades": ["EF89EF02", "EF89EF04", "EF89EF06"]},
    {"ano": "8º ano", "trimestre": 1, "unidade": "Ginástica", "objetos":
        "Ginástica de Condicionamento Físico: Localizada e crossfit\nGinástica de Consciência Corporal: Yoga e alongamento",
        "habilidades": ["EF89EF07", "EF89EF10"]},

    {"ano": "8º ano", "trimestre": 2, "unidade": "Esporte", "objetos":
        "Esporte Taco ou Campo: Softbol e taco",
        "habilidades": ["EF89EF02", "EF89EF04", "EF89EF06"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Práticas Corporais de Aventura", "objetos":
        "PCA de Natureza: Slackline, escalada horizontal e vertical",
        "habilidades": ["EF89EF19", "EF89EF20", "EF89EF21"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Dança", "objetos":
        "Dança de Salão: Forró",
        "habilidades": ["EF89EF12", "EF89EF13"]},

    {"ano": "8º ano", "trimestre": 3, "unidade": "Esporte", "objetos":
        "Esporte de Invasão: Basquete e Futsal\nEsporte de Combate: Esgrima, Boxe",
        "habilidades": ["EF89EF02", "EF89EF04", "EF89EF06"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Lutas", "objetos":
        "Lutas do Mundo: Karatê e taekwondo",
        "habilidades": ["EF89EF16", "EF89EF18"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Esporte", "objetos":
        "Esporte de Rede e Parede: Voleibol, Tênis de quadra, Futvolei",
        "habilidades": ["EF89EF01", "EF89EF02", "EF89EF03", "EF89EF05"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Ginástica", "objetos":
        "Ginástica de Condicionamento Físico: Crossfit e musculação\nGinástica de Consciência Corporal: Yoga, Pilates, Mindfulness",
        "habilidades": ["EF89EF08", "EF89EF09", "EF89EF10"]},

    {"ano": "9º ano", "trimestre": 2, "unidade": "Esporte", "objetos":
        "Esporte Taco ou Campo: Beisebol, taco, críquete",
        "habilidades": ["EF89EF01", "EF89EF03", "EF89EF05", "EF89EF06"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Práticas Corporais de Aventura", "objetos":
        "PCA de Natureza: Slackline, trekking e campismo",
        "habilidades": ["EF89EF19", "EF89EF20", "EF89EF21"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Dança", "objetos":
        "Dança de Salão: Forró, valsa, bolero, tango, samba de gafieira",
        "habilidades": ["EF89EF12", "EF89EF13", "EF89EF14", "EF89EF15"]},

    {"ano": "9º ano", "trimestre": 3, "unidade": "Esporte", "objetos":
        "Esporte de Invasão: Basquete, Futsal e/ou futebol de campo\nEsporte de Combate: Boxe e judô",
        "habilidades": ["EF89EF01", "EF89EF02", "EF89EF03", "EF89EF05"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Lutas", "objetos":
        "Lutas do Mundo: Taekwondo, Muay Thay e Jiu Jitsu",
        "habilidades": ["EF89EF16", "EF89EF17", "EF89EF18"]},
]

# Referencial Curricular de Ciências — extraído de RF_Ciências.pdf (04/09/2026). Alguns
# trimestres do PDF combinavam 2 unidades temáticas no mesmo bloco (ex: 6º 1º tri tinha
# "Matéria e Energia" e "Vida e evolução" juntos) — foram separados em linhas distintas,
# uma por unidade, seguindo o mesmo padrão das outras disciplinas.
REFERENCIAL_CURRICULAR_CIENCIAS = [
    {"ano": "6º ano", "trimestre": 1, "unidade": "Matéria e Energia", "objetos":
        "Misturas homogêneas e heterogêneas\nTransformações químicas\nSeparação de materiais\nMateriais sintéticos",
        "habilidades": ["EF06CI01", "EF06CI02", "EF06CI03", "EF06CI04"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Vida e Evolução", "objetos":
        "Célula como unidade da vida",
        "habilidades": ["EF06CI05"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Vida e Evolução", "objetos":
        "Célula como unidade da vida\nInteração entre os sistemas locomotor e nervoso\nLentes corretivas",
        "habilidades": ["EF06CI06", "EF06CI07", "EF06CI08", "EF06CI09", "EF06CI10"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Terra e Universo", "objetos":
        "Forma, estrutura e movimentos da Terra",
        "habilidades": ["EF06CI11", "EF06CI12", "EF06CI13", "EF06CI14"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Terra e Universo", "objetos":
        "Composição do ar\nEfeito estufa\nCamada de ozônio\nFenômenos naturais (vulcões, terremotos e tsunamis)\nPlacas tectônicas e deriva continental",
        "habilidades": ["EF07CI12", "EF07CI13", "EF07CI14", "EF07CI15", "EF07CI16"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Vida e Evolução", "objetos":
        "Diversidade de ecossistemas\nFenômenos naturais e impactos ambientais\nProgramas e indicadores de saúde pública",
        "habilidades": ["EF07CI07", "EF07CI08", "EF07CI09", "EF07CI10"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Matéria e Energia", "objetos":
        "Máquinas simples\nFormas de propagação do calor\nEquilíbrio termodinâmico e vida na Terra\nHistória dos combustíveis e das máquinas térmicas",
        "habilidades": ["EF07CI01", "EF07CI02", "EF07CI03", "EF07CI04", "EF07CI05", "EF07CI06"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Matéria e Energia", "objetos":
        "Fontes e tipos de energia\nTransformação de energia\nCálculo de consumo de energia elétrica\nCircuitos elétricos\nUso consciente de energia elétrica",
        "habilidades": ["EF08CI01", "EF08CI02", "EF08CI03", "EF08CI04", "EF08CI05", "EF08CI06"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Vida e Evolução", "objetos":
        "Mecanismos reprodutivos\nSexualidade",
        "habilidades": ["EF08CI07", "EF08CI08", "EF08CI09", "EF08CI10", "EF08CI11"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Terra e Universo", "objetos":
        "Sistema Sol, Terra e Lua\nClima",
        "habilidades": ["EF08CI12", "EF08CI13", "EF08CI14", "EF08CI15", "EF08CI16"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Vida e Evolução", "objetos":
        "Hereditariedade\nIdeias evolucionistas\nPreservação da biodiversidade",
        "habilidades": ["EF09CI08", "EF09CI09", "EF09CI10", "EF09CI11", "EF09CI12", "EF09CI13"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Matéria e Energia", "objetos":
        "Aspectos quantitativos das transformações químicas\nEstrutura da matéria\nRadiações e suas aplicações na saúde",
        "habilidades": ["EF09CI01", "EF09CI03", "EF09CI01VR", "EF09CI02", "EF09CI06", "EF09CI07", "EF09CI08VR", "EF09CI04", "EF09CI05"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Matéria e Energia", "objetos":
        "Dinâmica\nCinemática",
        "habilidades": ["EF09CI03VR", "EF09CI04VR", "EF09CI05VR", "EF09CI06VR", "EF09CI07VR"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Terra e Universo", "objetos":
        "Composição, estrutura e localização do Sistema Solar no Universo\nAstronomia e cultura\nVida humana fora da Terra\nOrdem de grandeza astronômica\nEvolução estelar",
        "habilidades": ["EF09CI14", "EF09CI15", "EF09CI16", "EF09CI17"]},
]

# Referencial Curricular de História — extraído de RF_História.pdf (04/09/2026). O
# documento original usa "Eixo Temático" em vez de "Unidade Temática" — mapeado no mesmo
# campo unidade_tematica. Vários trimestres tinham eixos entrelaçados/repetidos na mesma
# página; consolidados em blocos por eixo, preservando todas as habilidades citadas.
REFERENCIAL_CURRICULAR_HISTORIA = [
    {"ano": "6º ano", "trimestre": 1, "unidade": "História: tempo, espaço e formas de registros", "objetos":
        "A questão do tempo, sincronias e diacronias: reflexões sobre o sentido das cronologias\nFormas de registro da história e da produção do conhecimento histórico\nAs origens da humanidade, seus deslocamentos e os processos de sedentarização",
        "habilidades": ["EF06HI01", "EF06HI02", "EF06HI03", "EF06HI04", "EF06HI05", "EF06HI06"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "A invenção do mundo clássico e o contraponto com outras sociedades", "objetos":
        "Povos da Antiguidade na África (egípcios), no Oriente Médio (mesopotâmicos) e nas Américas (pré-colombianos)\nOs povos indígenas originários do atual território brasileiro e seus hábitos culturais e sociais\nO Ocidente Clássico: aspectos da cultura na Grécia e em Roma",
        "habilidades": ["EF06HI07", "EF06HI08", "EF06HI09"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Lógicas de organização política", "objetos":
        "As noções de cidadania e política na Grécia e em Roma",
        "habilidades": ["EF06HI10"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Lógicas de organização política", "objetos":
        "As noções de cidadania e política na Grécia e em Roma\nDomínios e expansão das culturas grega e romana\nSignificados do conceito de 'império' e as lógicas de conquista, conflito e negociação\nAs diferentes formas de organização política na África: reinos, impérios, cidades-estados e sociedades linhageiras ou aldeias\nA passagem do mundo antigo para o mundo medieval\nA fragmentação do poder político na Idade Média\nO Mediterrâneo como espaço de interação entre as sociedades da Europa, da África e do Oriente Médio",
        "habilidades": ["EF06HI11", "EF06HI12", "EF06HI13", "EF06HI14", "EF06HI15"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Trabalho e formas de organização social e cultural", "objetos":
        "Senhores medievais e servos no mundo antigo e no mundo medieval\nEscravidão e trabalho livre em diferentes temporalidades e espaços (Roma Antiga, Europa medieval e África)\nLógicas comerciais na Antiguidade romana e no mundo medieval\nO papel da religião cristã, dos mosteiros e da cultura na Idade Média",
        "habilidades": ["EF06HI16", "EF06HI17", "EF06HI18"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "O mundo moderno e a conexão entre sociedades africanas, americanas e europeias", "objetos":
        "Saberes dos povos africanos e pré-colombianos expressos na cultura material e imaterial\nA construção da ideia de modernidade e seus impactos na concepção de História\nA ideia de 'Novo Mundo' ante o Mundo Antigo: permanências e rupturas de saberes e práticas na emergência do mundo moderno",
        "habilidades": ["EF07HI01", "EF07HI02", "EF07HI03"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "A organização do poder e as dinâmicas do mundo colonial americano", "objetos":
        "A formação e o funcionamento das monarquias europeias: a lógica da centralização política e os conflitos na Europa",
        "habilidades": ["EF07HI07"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Lógicas comerciais e mercantis da modernidade", "objetos":
        "As lógicas mercantis e o domínio europeu sobre os mares e o contraponto oriental",
        "habilidades": ["EF07HI13"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Humanismos, Renascimentos e o Novo Mundo", "objetos":
        "Humanismos: uma nova visão de ser humano e de mundo\nRenascimentos artísticos e culturais",
        "habilidades": ["EF07HI04"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Humanismos, Renascimentos e o Novo Mundo", "objetos":
        "As Reformas religiosas: a cristandade fragmentada\nAs descobertas científicas e a expansão marítima",
        "habilidades": ["EF07HI05", "EF07HI06"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "A organização do poder e as dinâmicas do mundo colonial americano", "objetos":
        "A conquista da América e as formas de organização política dos indígenas e europeus: conflitos, dominação e conciliação\nA estruturação dos vice-reinos nas Américas\nResistências indígenas, invasões e expansão na América portuguesa",
        "habilidades": ["EF07HI08", "EF07HI09", "EF07HI01-VR1", "EF07HI10"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "A organização do poder e as dinâmicas do mundo colonial americano", "objetos":
        "A estruturação dos vice-reinos nas Américas\nResistências indígenas, invasões e expansão na América portuguesa",
        "habilidades": ["EF07HI11", "EF07HI12"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Lógicas comerciais e mercantis da modernidade", "objetos":
        "As lógicas internas das sociedades africanas\nAs formas de organização das sociedades ameríndias\nA escravidão moderna e o tráfico de escravizados\nA emergência do capitalismo",
        "habilidades": ["EF07HI14", "EF07HI15", "EF07HI16", "EF07HI17"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "O mundo contemporâneo: o Antigo Regime em crise", "objetos":
        "As revoluções inglesas e os princípios do liberalismo\nA questão do iluminismo e da ilustração\nRevolução Industrial e seus impactos na produção e circulação de povos, produtos e culturas\nRevolução Francesa e seus desdobramentos\nRebeliões na América portuguesa: as conjurações mineira e baiana",
        "habilidades": ["EF08HI01", "EF08HI02", "EF08HI03", "EF08HI04", "EF08HI05"]},
    {"ano": "8º ano", "trimestre": 1, "unidade": "Os processos de independência nas Américas", "objetos":
        "A revolução dos escravizados em São Domingo e seus múltiplos significados: o caso do Haiti\nIndependência dos Estados Unidos da América\nIndependências na América espanhola\nOs caminhos até a independência do Brasil",
        "habilidades": ["EF08HI06", "EF08HI07", "EF08HI09", "EF08HI10", "EF08HI11"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Os processos de independência nas Américas", "objetos":
        "Independência dos Estados Unidos da América\nIndependências na América espanhola\nOs caminhos até a independência do Brasil\nA tutela da população indígena, a escravidão dos negros e a tutela dos egressos da escravidão",
        "habilidades": ["EF08HI08", "EF08HI12", "EF08HI13", "EF08HI14"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "O Brasil no século XIX", "objetos":
        "Brasil: Primeiro Reinado\nO Período Regencial e as contestações ao poder central\nO Brasil do Segundo Reinado: política e economia\nA Lei de Terras e seus desdobramentos na política do Segundo Reinado\nTerritórios e fronteiras: a Guerra do Paraguai\nO escravismo no Brasil do século XIX: plantations, revoltas de escravizados, abolicionismo e políticas migratórias no Brasil Imperial",
        "habilidades": ["EF08HI15", "EF08HI16", "EF08HI17", "EF08HI18", "EF08HI-VR1", "EF08HI19"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "O Brasil no século XIX", "objetos":
        "O escravismo no Brasil do século XIX: consequências da escravidão nas Américas\nPolíticas de extermínio do indígena durante o Império\nA produção do imaginário nacional brasileiro: cultura popular, representações visuais, letras e o Romantismo no Brasil",
        "habilidades": ["EF08HI20", "EF08HI-VR2", "EF08HI21", "EF08HI22"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Configurações do mundo no século XIX", "objetos":
        "Nacionalismo, revoluções e as novas nações europeias\nUma nova ordem econômica: as demandas do capitalismo industrial e o lugar das economias africanas e asiáticas nas dinâmicas globais\nOs Estados Unidos da América e a América Latina no século XIX\nO imperialismo europeu e a partilha da África e da Ásia\nPensamento e cultura no século XIX: darwinismo e racismo\nO discurso civilizatório nas Américas e a resistência dos povos indígenas",
        "habilidades": ["EF08HI23", "EF08HI24", "EF08HI25", "EF08HI26", "EF08HI27"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "O nascimento da República no Brasil e os processos históricos até a metade do século XX", "objetos":
        "Experiências republicanas e práticas autoritárias\nA Proclamação da República e seus primeiros desdobramentos\nA questão da inserção dos negros no período republicano pós-abolição\nOs movimentos sociais e a imprensa negra\nA cultura afro-brasileira como elemento de resistência e superação das discriminações\nPrimeira República e suas características\nContestações e dinâmicas da vida cultural no Brasil entre 1900 e 1930\nO período varguista e suas contradições\nA emergência da vida urbana e a segregação espacial\nO trabalhismo e seu protagonismo político\nAnarquismo e protagonismo feminino",
        "habilidades": ["EF09HI01", "EF09HI02", "EF09HI03", "EF09HI04", "EF09HI05", "EF09HI06", "EF09HI08", "EF09HI09"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Totalitarismos e conflitos mundiais", "objetos":
        "O colonialismo na África\nAs guerras mundiais, a crise do colonialismo e o advento dos nacionalismos africanos e asiáticos\nO mundo em conflito: a Primeira Guerra Mundial\nA questão da Palestina\nA Revolução Russa\nA crise capitalista de 1929\nA emergência do fascismo e do nazismo\nA Segunda Guerra Mundial\nJudeus e outras vítimas do holocausto\nA Organização das Nações Unidas (ONU) e a questão dos Direitos Humanos",
        "habilidades": ["EF09HI14", "EF09HI10", "EF09HI13", "EF09HI15"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Totalitarismos e conflitos mundiais", "objetos":
        "A Organização das Nações Unidas (ONU) e a questão dos Direitos Humanos\nA Guerra Fria: confrontos de dois modelos políticos\nA Revolução Chinesa e as tensões entre China e Rússia\nA Revolução Cubana e as tensões entre Estados Unidos da América e Cuba",
        "habilidades": ["EF09HI16", "EF09HI28"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Modernização, ditadura civil-militar e redemocratização: o Brasil após 1946", "objetos":
        "O Brasil da era JK e o ideal de uma nação moderna: a urbanização e seus desdobramentos\nOs anos 1960: revolução cultural?\nA ditadura civil-militar e os processos de resistência\nAs questões indígena e negra e a ditadura",
        "habilidades": ["EF09HI17", "EF09HI18", "EF09HI19", "EF09HI20", "EF09HI21"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "O nascimento da República no Brasil e os processos históricos até a metade do século XX", "objetos":
        "A questão indígena durante a República (até 1964)",
        "habilidades": ["EF09HI07"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "A história recente", "objetos":
        "As experiências ditatoriais na América Latina\nO processo de redemocratização\nA Constituição de 1988 e a emancipação das cidadanias (analfabetos, indígenas, negros, jovens etc.)",
        "habilidades": ["EF09HI29", "EF09HI30", "EF09HI22", "EF09HI23"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "A história recente", "objetos":
        "A história recente do Brasil: transformações políticas, econômicas, sociais e culturais de 1989 aos dias atuais\nOs protagonismos da sociedade civil\nA questão da violência contra populações marginalizadas\nO Brasil e suas relações internacionais na era da globalização\nOs processos de descolonização na África e na Ásia\nO fim da Guerra Fria e o processo de globalização\nPolíticas econômicas na América Latina\nOs conflitos do século XXI e a questão do terrorismo\nPluralidades e diversidades identitárias na atualidade\nAs pautas dos povos indígenas no século XXI",
        "habilidades": ["EF09HI24", "EF09HI25", "EF09HI26", "EF09HI27", "EF09HI-VR1", "EF09HI31", "EF09HI32", "EF09HI33", "EF09HI34", "EF09HI35", "EF09HI36"]},
]

# Referencial Curricular de Geografia — extraído de RF_Geografia.pdf (04/09/2026).
# NORMALIZAÇÃO: o PDF original tinha alguns códigos com erro de digitação (ex:
# "EF06GEO2", "EF06GEO8", "EF06GEO9", "EF06GEO11") — corrigidos pro padrão oficial BNCC
# (EF06GE02, EF06GE08, EF06GE09, EF06GE11), inclusive confirmado por descrição idêntica
# repetida no próprio documento (EF06GEO11 e EF06GE11 tinham exatamente o mesmo texto).
REFERENCIAL_CURRICULAR_GEOGRAFIA = [
    {"ano": "6º ano", "trimestre": 1, "unidade": "O sujeito e seu lugar no mundo", "objetos":
        "Conceitos geográficos: Lugar, Paisagem, Espaço geográfico, Território, Região",
        "habilidades": ["EF06GE01", "EF06GE02"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "Localização e orientação no espaço: pontos cardeais, coordenadas geográficas (paralelos/latitude, meridianos/longitude), movimentos da Terra (fusos horários)\nInterações das sociedades com a natureza e a biodiversidade",
        "habilidades": ["EF06GE11"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Conexões e escalas", "objetos":
        "Os movimentos do planeta e a circulação atmosférica\nA Terra: placas tectônicas, formação do relevo, agentes internos e externos, ação humana sobre o relevo",
        "habilidades": ["EF06GE03", "EF06GE05"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Formas de representação e pensamento espacial", "objetos":
        "Alfabetização Cartográfica: importância dos mapas, escalas, legendas, convenções cartográficas, análise de mapas, técnicas modernas",
        "habilidades": ["EF06GE08", "EF06GE09"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Conexões e escalas", "objetos":
        "Hidrosfera: Oceanos e Mares (salinidade, marés, correntes marítimas, tipos de mares); Águas Continentais (rios, bacias hidrográficas, aproveitamento econômico)\nAtmosfera: Clima e Tempo (previsão, fatores climáticos, zonas climáticas)",
        "habilidades": ["EF06GE04", "EF06GE03"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "Consumo de recursos hídricos e uso das bacias hidrográficas no Brasil e no mundo\nA interferência humana na atmosfera: potencialização do efeito estufa e aquecimento global",
        "habilidades": ["EF06GE12", "EF06GE13"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Conexões e escalas", "objetos":
        "Biomas terrestres",
        "habilidades": ["EF06GE05"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "A degradação ambiental resultante das atividades econômicas (queimadas, desmatamentos etc.)",
        "habilidades": ["EF06GE10", "EF06GE11"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Mundo do trabalho", "objetos":
        "Formas de apropriação e organização do espaço: extrativismo, agricultura, pecuária, indústria, comércio, turismo\nO trabalho humano transformando a natureza e o espaço",
        "habilidades": ["EF06GE06", "EF06GE07"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Conexões e escalas", "objetos":
        "O Brasil e o Mundo: Localização, Extensão, Fusos Horários, formação do território brasileiro",
        "habilidades": ["EF07GE02", "EF07GE03"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Formas de representação e pensamento espacial", "objetos":
        "O Brasil e suas paisagens: regiões brasileiras (IBGE), diferenças entre regiões, regiões geoeconômicas, desigualdades socioespaciais e fluxos migratórios regionais e extrarregionais",
        "habilidades": ["EF07GE10", "EF07GE0VR"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Mundo do trabalho", "objetos":
        "A População brasileira: formação, indicadores populacionais, crescimento, diversidade étnico-cultural, distribuição e fluxos migratórios",
        "habilidades": ["EF07GE04", "EF07GE02"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "Aspectos físicos do Brasil: relevo geral, climas do Brasil (áreas de predomínio, características, influência), vegetação (primitiva e atual)",
        "habilidades": ["EF07GE11", "EF07GE12"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Conexões e escalas", "objetos":
        "Hidrografia: potencial hídrico do país, principais bacias, degradação dos recursos hídricos, importância para populações ribeirinhas (Rio Paraíba do Sul e problemas socioambientais)",
        "habilidades": ["EF06GE04", "EF07GE0VR"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Conexões e escalas", "objetos":
        "A urbanização do espaço brasileiro: transformações nas paisagens das cidades, problemas urbanos (moradia, transporte, emprego, escolas)",
        "habilidades": ["EF07GE02"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Mundo do trabalho", "objetos":
        "A Geografia no Campo: atividades agropecuárias, degradação dos ecossistemas (desmatamento, queimadas), remanescentes quilombolas, reforma agrária, reservas indígenas\nA dinâmica da economia brasileira: atividades industriais, comércio, transportes e comunicações",
        "habilidades": ["EF06GE06", "EF06GE10", "EF07GE07", "EF07GE08"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Conexões e escalas", "objetos":
        "Geopolítica e relações internacionais: Estado, Nação, Território, País, organizações internacionais (FMI, ONU, Banco Mundial); Guerra Fria\nDiferentes formas de regionalização: Primeiro/Segundo/Terceiro Mundo, desenvolvidos/subdesenvolvidos/em desenvolvimento, os BRICS",
        "habilidades": ["EF08GE05", "EF08GE06", "EF08GE09"]},
    {"ano": "8º ano", "trimestre": 1, "unidade": "O sujeito e seu lugar no mundo", "objetos":
        "População Mundial\nMigrações",
        "habilidades": ["EF08GE03", "EF08GE01", "EF08GE0VR"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "Continente americano: divisões (América do Norte/Central/Sul, América Latina e Anglo-Saxônica)\nAspectos humanos do continente americano: povos nativos, herança cultural dos nativos e africanos, dominação europeia e formas de resistência",
        "habilidades": ["EF08GE23", "EF08GE20"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Mundo do trabalho", "objetos":
        "Aspectos físicos do continente americano: relevo, clima, hidrografia, vegetação",
        "habilidades": ["EF08GE15"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Conexões e escalas", "objetos":
        "Os processos de colonização\nAspectos econômicos do continente americano (extrativismo, atividades agropecuárias, indústria, comércio, comunicações)\nO domínio dos EUA na América Latina (político, econômico, tecnológico, cultural)\nConflito comercial entre EUA e China\nA integração regional no continente americano e o papel do Brasil",
        "habilidades": ["EF08GE08", "EF08GE07"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "África: aspectos físicos (relevo, clima, hidrografia, vegetação)\nA descolonização e suas consequências: guerras civis, o Apartheid na África do Sul, aspectos econômicos e a questão ambiental",
        "habilidades": ["EF08GEVR", "EF08GE20"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Conexões e escalas", "objetos":
        "A África Pré-Colonial\nA colonização da África no século XIX: a partilha e a desestruturação das sociedades africanas",
        "habilidades": ["EF08GE05"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Conexões e escalas", "objetos":
        "Socialismo e Capitalismo e suas características\nA Nova Ordem Mundial: fim da Guerra Fria, crise da URSS, mundo multipolar, revolução técnico-científica-informacional, globalização e a nova DIT",
        "habilidades": ["EF09GE06", "EF09GE05"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Formas de representação e pensamento espacial", "objetos":
        "A Nova Ordem Mundial: principais fluxos da globalização, diferentes níveis de integração econômica",
        "habilidades": ["EF09GE14"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Mundo do trabalho", "objetos":
        "Impactos do processo de industrialização na produção e circulação de produtos e culturas na Europa, Ásia e Oceania\nMudanças técnicas e científicas decorrentes da industrialização e as transformações no trabalho",
        "habilidades": ["EF09GE10", "EF09GE11"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Natureza, ambientes e qualidade de vida", "objetos":
        "Europa: aspectos físicos (relevo, clima, hidrografia, vegetação) e determinantes histórico-geográficos da divisão Europa/Ásia",
        "habilidades": ["EF09GE07", "EF09GE16", "EF09GE17"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "O sujeito e seu lugar no mundo", "objetos":
        "Conflitos étnicos\nAspectos econômicos: formação e ampliação dos blocos econômicos, características da União Europeia, desenvolvimento econômico e degradação ambiental\nO Leste Europeu\nRússia",
        "habilidades": ["EF09GE01"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Conexões e escalas", "objetos":
        "Transformações territoriais na Europa, Ásia e Oceania: fronteiras, tensões, conflitos e múltiplas regionalidades\nCaracterísticas de países europeus, asiáticos e da Oceania\nÁsia: aspectos físicos e econômicos",
        "habilidades": ["EF09GE08", "EF09GE09", "EF09GE07"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Conexões e escalas", "objetos":
        "Países e regiões de destaque por peculiaridades socioeconômicas e políticas: China, Oriente Médio, Japão, Tigres Asiáticos, Índia",
        "habilidades": ["EF09GE08", "EF09GE09"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "O sujeito e seu lugar no mundo", "objetos":
        "Oceania: diferenças de paisagens e modos de viver de diferentes povos na Europa, Ásia e Oceania",
        "habilidades": ["EF09GE04", "EF09GE10"]},
]



# Referencial Curricular de Matemática — extraído de RF__Matematica.pdf (enviado por
# Felipe em 04/09/2026). O PDF trata as unidades temáticas de Matemática (Números,
# Álgebra, Geometria, Grandezas e medidas, Probabilidade e Estatística) em dois blocos
# separados por ano (Álgebra e Geometria), então quando a mesma unidade temática se
# repetia no mesmo ano/trimestre com objetos de conhecimento diferentes, os objetos
# foram unificados numa única linha (um por parágrafo), somando as habilidades BNCC
# de todos eles — mesmo padrão já usado nas outras disciplinas.
REFERENCIAL_CURRICULAR_MATEMATICA = [
    # ---------- 6º ANO ----------
    {"ano": "6º ano", "trimestre": 1, "unidade": "Números", "objetos":
        "Sistema de numeração decimal; características, leitura, escrita e comparação de números naturais e de números racionais representados na forma decimal.\n"
        "Operações (adição, subtração, multiplicação, divisão e potenciação) com números naturais. Raiz quadrada de números naturais. Divisão euclidiana.\n"
        "Aproximação de números para múltiplos de potências de 10.\n"
        "Fluxograma para determinar a paridade de um número natural. Múltiplos e divisores de um número natural. Números primos e compostos.\n"
        "Frações: significados (parte/todo, quociente), equivalência, comparação, adição e subtração; cálculo da fração de um número natural.",
        "habilidades": ["EF06MA01", "EF06MA02", "EF06MA03", "EF06MA12", "EF06MA04", "EF06MA05", "EF06MA06",
                        "EF06MA07", "EF06MA08", "EF06MA09", "EF06MA10"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Grandezas e medidas", "objetos":
        "Ângulos: noção, usos e medida",
        "habilidades": ["EF06MA25", "EF06MA26", "EF06MA27"]},
    {"ano": "6º ano", "trimestre": 1, "unidade": "Geometria", "objetos":
        "Polígonos: classificações quanto ao número de vértices, às medidas de lados e ângulos e ao paralelismo e perpendicularismo dos lados\n"
        "Construção de retas paralelas e perpendiculares, fazendo uso de réguas, esquadros e softwares",
        "habilidades": ["EF06MA18", "EF06MA19", "EF06MA20", "EF06MA22", "EF06MA23"]},

    {"ano": "6º ano", "trimestre": 2, "unidade": "Números", "objetos":
        "Operações (adição, subtração, multiplicação, divisão e potenciação) com números racionais. Raiz quadrada de números racionais positivos.\n"
        "Cálculo de porcentagens por meio de estratégias diversas, sem fazer uso da \"regra de três\"",
        "habilidades": ["EF06MA11", "EF06MA13"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Probabilidade e Estatística", "objetos":
        "Cálculo de probabilidade como a razão entre o número de resultados favoráveis e o total de resultados possíveis em um espaço amostral equiprovável\n"
        "Cálculo de probabilidade por meio de muitas repetições de um experimento (frequências de ocorrências e probabilidade frequentista)",
        "habilidades": ["EF06MA30"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Álgebra", "objetos":
        "Problemas que tratam da partição de um todo em duas partes desiguais, envolvendo razões entre as partes e entre uma das partes e o todo",
        "habilidades": ["EF06MA15"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Geometria", "objetos":
        "Plano cartesiano: associação dos vértices de um polígono a pares ordenados\n"
        "Construção de figuras semelhantes: ampliação e redução de figuras planas em malhas quadriculadas",
        "habilidades": ["EF06MA16", "EF06MA21"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Grandezas e medidas", "objetos":
        "Perímetro de um quadrado como grandeza proporcional à medida do lado\n"
        "Plantas baixas e vistas aéreas",
        "habilidades": ["EF06MA29", "EF06MA28"]},

    {"ano": "6º ano", "trimestre": 3, "unidade": "Álgebra", "objetos":
        "Propriedades da Igualdade",
        "habilidades": ["EF06MA14"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Probabilidade e Estatística", "objetos":
        "Leitura e interpretação de tabelas e gráficos (de colunas ou barras simples ou múltiplas) referentes a variáveis categóricas e variáveis numéricas\n"
        "Coleta de dados, organização e registro. Construção de diferentes tipos de gráficos para representá-los e interpretação das informações\n"
        "Diferentes tipos de representação de informações: gráficos e fluxogramas",
        "habilidades": ["EF06MA31", "EF06MA32", "EF06MA33", "EF06MA34"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Grandezas e medidas", "objetos":
        "Problemas sobre medidas envolvendo grandezas como comprimento, massa, tempo, temperatura, área, capacidade e volume.",
        "habilidades": ["EF06MA24"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Geometria", "objetos":
        "Prismas e pirâmides: planificações e relações entre seus elementos (vértices, faces e arestas)",
        "habilidades": ["EF06MA17"]},

    # ---------- 7º ANO ----------
    {"ano": "7º ano", "trimestre": 1, "unidade": "Números", "objetos":
        "Múltiplos e divisores de um número natural\n"
        "Números inteiros: usos, história, ordenação, associação com pontos da reta numérica e operações\n"
        "Fração e seus significados: como parte de inteiros, resultado da divisão, razão e operador.\n"
        "Números racionais na representação fracionária e na decimal: usos, ordenação e associação com pontos da reta numérica e operações.\n"
        "Cálculo de porcentagens e de acréscimos e decréscimos simples",
        "habilidades": ["EF07MA01", "EF07MA03", "EF07MA04", "EF07MA05", "EF07MA06", "EF07MA07", "EF07MA08",
                        "EF07MA09", "EF07MA10", "EF07MA11", "EF07MA12", "EF07MA02"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Geometria", "objetos":
        "Transformações geométricas de polígonos no plano cartesiano: multiplicação das coordenadas por um número inteiro e obtenção de simétricos em relação aos eixos e à origem\n"
        "Simetrias de translação, rotação e reflexão\n"
        "Relações entre os ângulos formados por retas paralelas intersectadas por uma transversal",
        "habilidades": ["EF07MA19", "EF07MA20", "EF07MA21", "EF07MA23"]},
    {"ano": "7º ano", "trimestre": 1, "unidade": "Grandezas e medidas", "objetos":
        "A circunferência como lugar geométrico",
        "habilidades": ["EF07MA22"]},

    {"ano": "7º ano", "trimestre": 2, "unidade": "Álgebra", "objetos":
        "Linguagem algébrica: variável e incógnita\n"
        "Equivalência de expressões algébricas: identificação da regularidade de uma sequência numérica\n"
        "Equações polinomiais do 1o grau\n"
        "Problemas envolvendo grandezas diretamente proporcionais e grandezas inversamente proporcionais",
        "habilidades": ["EF07MA13", "EF07MA14", "EF07MA15", "EF07MA16", "EF07MA18", "EF07MA17"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Geometria", "objetos":
        "Triângulos: construção, condição de existência e soma das medidas dos ângulos internos.\n"
        "Polígonos regulares: quadrado e triângulo equilátero",
        "habilidades": ["EF07MA24", "EF07MA25", "EF07MA26", "EF07MA27", "EF07MA28"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Grandezas e medidas", "objetos":
        "Problemas envolvendo medições\n"
        "Cálculo de volume de blocos retangulares, utilizando unidades de medida convencionais mais usuais",
        "habilidades": ["EF07MA29", "EF07MA30"]},

    {"ano": "7º ano", "trimestre": 3, "unidade": "Probabilidade e estatística", "objetos":
        "Experimentos aleatórios: espaço amostral e estimativa de probabilidade por meio de frequência de ocorrências\n"
        "Estatística: média aritmética simples, média aritmética ponderada e amplitude de um conjunto de dados.\n"
        "Pesquisa amostral e pesquisa censitária. Planejamento de pesquisa, coleta e organização dos dados, construção de tabelas e gráficos e interpretação das informações\n"
        "Gráficos de setores: interpretação, pertinência e construção para representar conjunto de dados.",
        "habilidades": ["EF07MA34", "EF07MA35", "EF07MA36", "EF07MA37"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Grandezas e medidas", "objetos":
        "Equivalência de área de figuras planas: cálculo de áreas de figuras que podem ser decompostas por outras áreas, cujas áreas podem ser facilmente determinadas como triângulos e quadriláteros.\n"
        "Medida do comprimento da circunferência",
        "habilidades": ["EF07MA31", "EF07MA32", "EF07MA33"]},

    # ---------- 8º ANO ----------
    {"ano": "8º ano", "trimestre": 1, "unidade": "Números", "objetos":
        "Potenciação e radiciação\n"
        "Notação Científica\n"
        "Dízimas periódicas: fração geratriz\n"
        "Variação de grandezas: diretamente proporcionais, inversamente proporcionais ou não proporcionais\n"
        "Porcentagens",
        "habilidades": ["EF08MA02", "EF08MA01", "EF08MA05", "EF08MA12", "EF08MA13", "EF08MA04"]},
    {"ano": "8º ano", "trimestre": 1, "unidade": "Geometria", "objetos":
        "Congruência de triângulos e demonstrações de propriedades de quadriláteros\n"
        "Construções geométricas: ângulos de 90º, 60º, 45º e 30º e polígonos regulares\n"
        "Mediatriz e bissetriz como lugares geométricos: construção e problemas",
        "habilidades": ["EF08MA14", "EF08MA15", "EF08MA16", "EF08MA17"]},

    {"ano": "8º ano", "trimestre": 2, "unidade": "Álgebra", "objetos":
        "Sequências recursivas e não recursivas\n"
        "Valor numérico de expressões algébricas\n"
        "Sistema de equações polinomiais de 1º grau: resolução algébrica e representação no plano cartesiano\n"
        "Associação de uma equação linear de 1º grau a uma reta no plano cartesiano\n"
        "Equação polinomial de 2º grau do tipo ax2 = b",
        "habilidades": ["EF08MA10", "EF08MA11", "EF08MA06", "EF08MA08", "EF08MA07", "EF08MA09"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Geometria", "objetos":
        "Transformações geométricas: simetrias de translação, reflexão e rotação\n"
        "Área de figuras planas. Área do círculo e comprimento de sua circunferência",
        "habilidades": ["EF08MA18", "EF08MA19"]},

    {"ano": "8º ano", "trimestre": 3, "unidade": "Probabilidade e estatística", "objetos":
        "O princípio multiplicativo da contagem\n"
        "Soma das probabilidades de todos os elementos de um espaço amostral\n"
        "Gráficos de barras, colunas, linhas ou setores e seus elementos constitutivos e adequação para determinado conjunto de dados\n"
        "Organização dos dados de uma variável contínua em classes\n"
        "Medidas de tendência central e de dispersão\n"
        "Pesquisas censitária ou amostral. Planejamento e execução de pesquisa amostral",
        "habilidades": ["EF08MA03", "EF08MA22", "EF08MA23", "EF08MA24", "EF08MA25", "EF08MA26", "EF08MA27"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Grandezas e medidas", "objetos":
        "Volume do cubo, do paralelepípedo e do cilindro reto. Medidas de capacidade",
        "habilidades": ["EF08MA20", "EF08MA21"]},

    # ---------- 9º ANO ----------
    {"ano": "9º ano", "trimestre": 1, "unidade": "Números", "objetos":
        "Necessidade dos números reais para medir qualquer segmento de reta. Números irracionais: reconhecimento e localização de alguns na reta numérica\n"
        "Potências com expoentes negativos e fracionários\n"
        "Números reais: notação científica e problemas\n"
        "Porcentagens: problemas que envolvem cálculo de percentuais sucessivos",
        "habilidades": ["EF09MA01", "EF09MA02", "EF09MA03", "EF09MA04", "EF09MA05"]},
    {"ano": "9º ano", "trimestre": 1, "unidade": "Geometria", "objetos":
        "Retas paralelas cortadas por transversais: teoremas de proporcionalidade e verificações experimentais\n"
        "Semelhança de triângulos\n"
        "Teorema de Pitágoras: verificações experimentais e demonstração. Relações métricas no triângulo retângulo",
        "habilidades": ["EF09MA14", "EF09MA12", "EF09MA13"]},

    {"ano": "9º ano", "trimestre": 2, "unidade": "Álgebra", "objetos":
        "Razão entre grandezas de espécies diferentes. Expressões algébricas: fatoração e produtos notáveis\n"
        "Grandezas diretamente proporcionais e grandezas inversamente proporcionais\n"
        "Resolução de equações polinomiais do 2o grau por meio de fatorações e por meio da fórmula resolutiva de Bhaskara.\n"
        "Funções: representações numérica, algébrica e gráfica",
        "habilidades": ["EF09MA07", "EF09MA08", "EF09MA09", "EF09MA06"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Geometria", "objetos":
        "Distância entre pontos no plano cartesiano\n"
        "Demonstrações de relações entre os ângulos formados por retas paralelas intersectadas por uma transversal\n"
        "Vistas ortogonais de figuras espaciais",
        "habilidades": ["EF09MA16", "EF09MA10", "EF09MA17"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Grandezas e medidas", "objetos":
        "Volume de prismas e cilindros",
        "habilidades": ["EF09MA19"]},

    {"ano": "9º ano", "trimestre": 3, "unidade": "Probabilidade e estatística", "objetos":
        "Análise de probabilidade de eventos aleatórios: eventos dependentes e independentes\n"
        "Análise de gráficos divulgados pela mídia: elementos que podem induzir a erros de leitura ou de interpretação\n"
        "Leitura, interpretação e representação de dados de pesquisa expressos em tabelas de dupla entrada, gráficos de colunas simples e agrupadas, gráficos de barras e de setores e gráficos pictóricos\n"
        "Planejamento e execução de pesquisa amostral e apresentação de relatório.",
        "habilidades": ["EF09MA20", "EF09MA21", "EF09MA22", "EF09MA23"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Geometria", "objetos":
        "Unidades de medida para medir distâncias muito grandes e muito pequenas. Unidades de medida utilizadas na informática\n"
        "Relações entre arcos e ângulos na circunferência de um círculo\n"
        "Polígonos regulares",
        "habilidades": ["EF09MA18", "EF09MA11", "EF09MA15"]},
]


# Referencial Curricular de Língua Inglesa — extraído de RF__Língua_Inglesa.pdf (enviado
# por Felipe em 04/09/2026). Diferente das outras disciplinas, o PDF organiza a maior
# parte do conteúdo por EIXO (Oralidade, Leitura, Escrita, Dimensão Intercultural) e
# marca essas seções como valendo para "1º, 2º e 3º Trimestres" ao mesmo tempo — só o
# eixo "Conhecimentos Específicos" (léxico/gramática) é dividido por trimestre. Para
# manter o mesmo modelo de banco usado nas outras disciplinas (uma linha por
# ano+trimestre+unidade), essas seções dos 4 primeiros eixos foram DUPLICADAS nos 3
# trimestres (o conteúdo é idêntico nos três) — assim elas aparecem como sugestão do
# Referencial Curricular independente do trimestre que o professor estiver planejando.
# Já o eixo "Estudo do léxico/Gramática" tem um bloco diferente por trimestre, como no
# PDF original. Usa o nome de disciplina "Inglês" para bater com o já usado no resto do
# sistema (boletim, e-cidade etc.), embora o PDF chame a disciplina de "Língua Inglesa".

_ING_ORALIDADE_INTERACAO = {
    "6º ano": ("Construção de laços afetivos e convívio social\n"
               "Funções e usos da língua inglesa em sala de aula (Classroom language)",
               ["EF06LI01", "EF06LI02", "EF06LI03"]),
    "7º ano": ("Funções e usos da língua inglesa: convivência e colaboração em sala de aula\n"
               "Práticas investigativas",
               ["EF07LI01", "EF07LI02"]),
    "8º ano": ("Negociação de sentidos (mal-entendidos no uso da língua inglesa e conflito de opiniões)\n"
               "Usos de recursos linguísticos e paralinguísticos no intercâmbio oral",
               ["EF08LI01", "EF08LI02"]),
    "9º ano": ("Funções e usos da língua inglesa: persuasão",
               ["EF09LI01"]),
}
_ING_ORALIDADE_COMPREENSAO = {
    "6º ano": ("Estratégias de compreensão de textos orais: palavras cognatas e pistas do contexto discursivo",
               ["EF06LI04"]),
    "7º ano": ("Estratégias de compreensão de textos orais: conhecimentos prévios\n"
               "Compreensão de textos orais de cunho descritivo ou narrativo",
               ["EF07LI03", "EF07LI04"]),
    "8º ano": ("Estratégias de compreensão de textos orais: conhecimentos prévios",
               ["EF08LI03"]),
    "9º ano": ("Compreensão de textos orais, multimodais, de cunho argumentativo",
               ["EF09LI02", "EF09LI03"]),
}
_ING_ORALIDADE_PRODUCAO = {
    "6º ano": ("Produção de textos orais, com a mediação do professor",
               ["EF06LI01", "EF06LI02"]),
    "7º ano": ("Produção de textos orais, com mediação do professor",
               ["EF07LI05"]),
    "8º ano": ("Produção de textos orais com autonomia",
               ["EF08LI04"]),
    "9º ano": ("Produção de textos orais com autonomia",
               ["EF09LI04"]),
}
_ING_LEITURA_ESTRATEGIAS = {
    "6º ano": ("Hipóteses sobre a finalidade de um texto\n"
               "Compreensão geral e específica: leitura rápida (skimming, scanning)",
               ["EF06LI07", "EF06LI08", "EF06LI09"]),
    "7º ano": ("Compreensão geral e específica: leitura rápida (skimming, scanning)\n"
               "Construção do sentido global do texto",
               ["EF07LI06", "EF07LI07", "EF07LI08"]),
    "8º ano": ("Construção de sentidos por meio de inferências e reconhecimento de implícitos",
               ["EF08LI05"]),
    "9º ano": ("Recursos de persuasão\n"
               "Recursos de argumentação",
               ["EF09LI05", "EF09LI06", "EF09LI07"]),
}
_ING_LEITURA_PRATICAS = {
    "6º ano": ("Práticas de leitura e construção de repertório lexical", "Construção de repertório lexical e autonomia leitora",
               ["EF06LI10", "EF06LI11"]),
    "7º ano": ("Práticas de leitura e construção de repertório lexical", "Objetivos de leitura\nLeitura de textos digitais para estudo",
               ["EF07LI09", "EF07LI10"]),
    "8º ano": ("Práticas de leitura e fruição", "Leitura de textos de cunho artístico/literário",
               ["EF08LI06", "EF08LI07"]),
    "9º ano": ("Práticas de leitura e novas tecnologias", "Informações em ambientes virtuais",
               ["EF09LI08"]),
}
_ING_LEITURA_ATITUDES = {
    "6º ano": ("Atitudes e disposições favoráveis do leitor", "Partilha de leitura, com mediação do professor",
               ["EF06LI12"]),
    "7º ano": ("Atitudes e disposições favoráveis do leitor", "Partilha de leitura",
               ["EF07LI11"]),
    "8º ano": ("Avaliação dos textos lidos", "Reflexão pós-leitura",
               ["EF08LI08"]),
    "9º ano": ("Avaliação dos textos lidos", "Reflexão pós-leitura",
               ["EF09LI09"]),
}
_ING_ESCRITA_PRE = {
    "6º ano": ("Estratégias de escrita: pré-escrita", "Planejamento do texto: brainstorming\nPlanejamento do texto: organização de ideias",
               ["EF06LI13", "EF06LI14"]),
    "7º ano": ("Estratégias de escrita: pré-escrita", "Pré-escrita: planejamento de produção escrita, com mediação do professor\n"
               "Escrita: organização em parágrafos ou tópicos, com mediação do professor",
               ["EF07LI12", "EF07LI13"]),
    "8º ano": ("Estratégias de escrita: escrita e pós-escrita", "Revisão de textos com a mediação do professor",
               ["EF08LI09", "EF08LI10"]),
    "9º ano": ("Estratégias de escrita", "Escrita: construção da argumentação\nEscrita: construção da persuasão",
               ["EF09LI10", "EF09LI11"]),
}
_ING_ESCRITA_PRATICAS = {
    "6º ano": ("Produção de textos escritos, em formatos diversos, com a mediação do professor",
               ["EF06LI15"]),
    "7º ano": ("Produção de textos escritos, em formatos diversos, com mediação do professor",
               ["EF07LI14"]),
    "8º ano": ("Produção de textos escritos com mediação do professor/colegas",
               ["EF08LI11"]),
    "9º ano": ("Produção de textos escritos, com mediação do professor/colegas",
               ["EF09LI12"]),
}
_ING_INTERCULTURAL_1 = {
    "6º ano": ("A língua inglesa no mundo", "Países que têm a língua inglesa como língua materna e/ou oficial",
               ["EF06LI24"]),
    "7º ano": ("A língua inglesa no mundo", "A língua inglesa como língua global na sociedade contemporânea",
               ["EF07LI21"]),
    "8º ano": ("Manifestações culturais", "Construção de repertório artístico-cultural",
               ["EF08LI18"]),
    "9º ano": ("A língua inglesa no mundo", "Expansão da língua inglesa: contexto histórico\n"
               "A língua inglesa e seu papel no intercâmbio científico, econômico e político",
               ["EF09LI17", "EF09LI18"]),
}
_ING_INTERCULTURAL_2 = {
    "6º ano": ("A língua inglesa no cotidiano da sociedade brasileira/comunidade", "Presença da língua inglesa no cotidiano",
               ["EF06LI25", "EF06LI26"]),
    "7º ano": ("Comunicação intercultural", "Variação linguística",
               ["EF07LI22", "EF07LI23"]),
    "8º ano": ("Comunicação intercultural", "Impacto de aspectos culturais na comunicação",
               ["EF08LI19", "EF08LI20"]),
    "9º ano": ("Comunicação intercultural", "Construção de identidades no mundo globalizado",
               ["EF09LI19"]),
}
_ING_LEXICO_GRAMATICA = {
    "6º ano": {
        1: ("Cognates; Family Members; Greetings, Leave Takings, and Introductions; Countries, Nationalities, and "
            "Languages; Cardinal Numbers (1-100); Telling the Time; Interrogatives – WH Questions; Subject Pronouns; "
            "Forms of Address; Verb To Be – Present Tense (Affirmative Form)",
            ["EF06LI16", "EF06LI17", "EF06LI18"]),
        2: ("School Objects and School Subjects; People at School; Days of the Week; Verb To Be – Present Tense "
            "(Interrogative and Negative Forms); Possessive Adjectives; Imperative; Possessive Case",
            ["EF06LI17", "EF06LI18", "EF06LI19", "EF06LI21", "EF06LI22", "EF06LI23"]),
        3: ("Physical Activities; Sports and Free-Time Activities; Animals; Colors; Daily Routine; Simple Present "
            "Tense (Routine); Present Continuous Tense (Affirmative, Interrogative, and Negative Forms); Adverbs of "
            "Frequency; Definite and Indefinite Articles; Plural Forms of Nouns",
            ["EF06LI16", "EF06LI17", "EF06LI19", "EF06LI20"]),
    },
    "7º ano": {
        1: ("Means of Transportation; Giving Directions; Age-Appropriate Activities (Leisure Activities); Cardinal "
            "Numbers (1–100); Telling the Time; Ordinal Numbers (1st–31st); Prepositions (in, on, at, next to, "
            "under, between, behind, in front of); Demonstrative Pronouns (This/That – These/Those); Interrogatives "
            "– WH Questions; Adjectives; Modal Verbs: Can/Could (Abilities and Possibilities)",
            ["EF06LI17", "EF07LI15", "EF07LI20"]),
        2: ("Parts of the Body; Clothes; Movie Genres; Adjectives to Describe Movies and Characters; Same Words, "
            "Different Meanings (polissemia); Verb To Be – Present Tense (Review); Present Continuous Tense "
            "(Affirmative, Negative, and Interrogative Forms); Verb To Be – Past Tense; There To Be – Present and "
            "Past Tenses; Object Pronouns; Linking Words",
            ["EF06LI17", "EF07LI15", "EF07LI17", "EF07LI18", "EF07LI19"]),
        3: ("Personality Adjectives; Dates (Months of the Year); Parts of the House; House Items; Describing Houses "
            "and Rooms (Using Adjectives); Pets; Past Continuous Tense; Simple Past Tense (Regular and Irregular "
            "Verbs); Prepositions of Time and Place: in, on, at; Simple Past and Past Continuous (Review and "
            "Contrast); Subject Pronouns vs. Object Pronouns",
            ["EF06LI17", "EF07LI15", "EF07LI18"]),
    },
    "8º ano": {
        1: ("Healthy food; Cooking Techniques and Measurements; Countable and Uncountable Nouns; Quantifiers; Some, "
            "Any, Much, Many, A Lot Of",
            ["EF06LI17", "EF08LI18", "EF08LI16"]),
        2: ("Physical Appearance; Adjectives Describing Personality and Character; Adjectives Describing Places; "
            "The World of Music; Pronouns; Adjectives and Their Order; Comparatives and Superlatives; Relative "
            "Pronouns; Prefixes and Suffixes",
            ["EF06LI17", "EF08LI18", "EF08LI08", "EF08LI15", "EF08LI17", "EF08LI13"]),
        3: ("The Environment; Natural Hazards; Travel Vocabulary; Expressions of Time and Probability; Verbs in the "
            "Future: Will / Going to; Adverbs of time, Frequency, and Place in the Future and Simple Past",
            ["EF06LI17", "EF08LI18", "EF08LI12", "EF08LI14", "EF08LI04"]),
    },
    "9º ano": {
        1: ("Internet Language; Use of Internet Slang; Health Issues; Interrogatives – WH Questions (Review); "
            "Phrasal Verbs; Modal Verbs (Should, Must, Have To, May, Might); Adjectives (Review)",
            ["EF06LI17", "EF09LI12", "EF09LI13", "EF09LI16"]),
        2: ("Environmental Problems; Environmentally Friendly Attitudes; Traffic and Transportation; Review of "
            "Simple Present, Simple Past, and Future Tenses; First and Second Conditionals",
            ["EF06LI17", "EF09LI02", "EF09LI03", "EF09LI12", "EF09LI15"]),
        3: ("Occupations; Food and Services; Modal and Phrasal Verbs (Review); Question Tags; Linking Words; "
            "Connectors",
            ["EF06LI17", "EF09LI05", "EF09LI16", "EF09LI14"]),
    },
}


def _montar_referencial_ingles():
    """Monta a lista final combinando os eixos (duplicados nos 3 trimestres) com o eixo
    Conhecimentos Específicos (um bloco distinto por trimestre)."""
    blocos = []
    for ano in ["6º ano", "7º ano", "8º ano", "9º ano"]:
        for trimestre in (1, 2, 3):
            obj, hab = _ING_ORALIDADE_INTERACAO[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Interação discursiva",
                            "objetos": obj, "habilidades": hab})
            obj, hab = _ING_ORALIDADE_COMPREENSAO[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Compreensão oral",
                            "objetos": obj, "habilidades": hab})
            obj, hab = _ING_ORALIDADE_PRODUCAO[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Produção oral",
                            "objetos": obj, "habilidades": hab})
            obj, hab = _ING_LEITURA_ESTRATEGIAS[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Estratégias de leitura",
                            "objetos": obj, "habilidades": hab})
            unidade, obj, hab = _ING_LEITURA_PRATICAS[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": unidade,
                            "objetos": obj, "habilidades": hab})
            unidade, obj, hab = _ING_LEITURA_ATITUDES[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": unidade,
                            "objetos": obj, "habilidades": hab})
            unidade, obj, hab = _ING_ESCRITA_PRE[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": unidade,
                            "objetos": obj, "habilidades": hab})
            obj, hab = _ING_ESCRITA_PRATICAS[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Práticas de escrita",
                            "objetos": obj, "habilidades": hab})
            unidade, obj, hab = _ING_INTERCULTURAL_1[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": unidade,
                            "objetos": obj, "habilidades": hab})
            unidade, obj, hab = _ING_INTERCULTURAL_2[ano]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": unidade,
                            "objetos": obj, "habilidades": hab})
            obj, hab = _ING_LEXICO_GRAMATICA[ano][trimestre]
            blocos.append({"ano": ano, "trimestre": trimestre, "unidade": "Estudo do léxico / Gramática",
                            "objetos": obj, "habilidades": hab})
    return blocos


REFERENCIAL_CURRICULAR_INGLES = _montar_referencial_ingles()



# Referencial Curricular de Língua Portuguesa — extraído de RF_Língua_Portuguesa.pdf
# (enviado por Felipe em 04/09/2026). Diferente das outras disciplinas, este PDF organiza
# o conteúdo em 3 grandes eixos (Gramática/Análise Linguística, Leitura, Produção de
# Texto), cada um cruzando os 4 campos de atuação da BNCC de Língua Portuguesa
# (Artístico-Literário, Jornalístico-Midiático, Atuação na Vida Pública, Práticas de
# Estudo e Pesquisa) dentro do mesmo bloco de trimestre — o PDF não separa por campo,
# então os 3 eixos foram usados como "unidade_tematica" (uma linha por ano+trimestre+
# eixo). O campo "objetos_conhecimento" reúne os gêneros/tópicos listados com marcador
# "•" no PDF (ortografia, gêneros textuais, tópicos gramaticais etc.); as habilidades
# BNCC são as citadas em cada bloco (muitas compartilhadas entre os 4 campos, como
# EF69LP55/EF69LP56, que valem para qualquer campo de atuação).
REFERENCIAL_CURRICULAR_PORTUGUES = [
    # =================== GRAMÁTICA (Análise Linguística/Semiótica) ===================
    {"ano": "6º ano", "trimestre": 1, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Comunicação e linguagem.\nLinguagem verbal, não verbal e mista.\nElementos da comunicação.\n"
        "Fonética e fonologia (fonema x letra, encontros consonantais e vocálicos, dígrafos).\nVariação linguística.\n"
        "Sinonímia e antonímia.\nSemântica (Hiperonímia e Hiponímia – como mecanismo de coesão).\n"
        "Ortografia (Mas/ Mais/Más).\nTipos de frase / pontuação.",
        "habilidades": ["EF69LP56", "EF69LP55", "EF06LP03", "EF67LP34", "EF06LP12", "EF67LP32", "EF67LP33", "EF67LP38"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Substantivo (classificação, gênero, número e grau).\nArtigo (Identificação e valor semântico).\n"
        "Adjetivo (gênero, número e grau; alteração de sentido – Ordem dos substantivos e dos adjetivos).\n"
        "Numeral (emprego e valor semântico).\nOrtografia mal/mau, há/a.\nInterjeição.\n"
        "Regras de acentuação gráfica (regras gerais).\nUso do hífen.",
        "habilidades": ["EF06LP04", "EF06LP06", "EF67LP32", "EF69LP56", "EF67LP35"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Verbo: Formação dos tempos verbais.\nModos: Indicativo, Subjuntivo e Imperativo.\nFormas Nominais.\n"
        "Pronomes (pessoais, demonstrativos, possessivos, indefinidos e interrogativos – identificação e valor "
        "semântico – pronominalização: retos e oblíquos – mecanismo de coesão).\n"
        "Ortografia (uso dos porquês e uso do G e J).",
        "habilidades": ["EF06LP04", "EF06LP05", "EF06LP06", "EF06LP08", "EF06LP09", "EF06LP10", "EF06LP11",
                        "EF69LP56", "EF67LP32"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Revisão das classes gramaticais (substantivo, adjetivo, artigo, numeral e pronome).\nAdjunto adnominal.\n"
        "Revisão de verbo (Modos Imperativo, Subjuntivo, Indicativo).\n"
        "Verbo e sua estrutura (conjugação, verbos regulares e irregulares, flexão, grafia de verbos irregulares – "
        "tem/têm etc.).\nAdvérbio.\nPreposição (essenciais e acidentais – valor semântico).\n"
        "Conjunções (valor semântico).\nOrtografia (uso do X e do CH).",
        "habilidades": ["EF07LP06", "EF07LP08", "EF07LP12", "EF07LP13", "EF69LP54", "EF07LP04", "EF07LP05",
                        "EF07LP10", "EF69LP17", "EF69LP28", "EF07LP09", "EF69LP56", "EF07LP11", "EF67LP36",
                        "EF69LP18", "EF67LP32"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Tipos de predicados.\nVerbos de ação e de ligação (foco no predicado nominal e verbo de ligação).\n"
        "Predicativo do sujeito.\nTransitividade verbal (retomar preposições).\nComplementos verbais.\n"
        "Ortografia (uso do C, Ç, S, SS, SC, SÇ, XC).",
        "habilidades": ["EF07LP07", "EF69LP55", "EF07LP08", "EF67LP26", "EF67LP32"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Estrutura e processo de formação das palavras.\nFrase, oração e período / pontuação.\n"
        "Sujeito (tipos de sujeito).\nOrtografia (uso do S, Z, X).",
        "habilidades": ["EF07LP03", "EF67LP34", "EF67LP35", "EF07LP04", "EF07LP10", "EF67LP33", "EF07LP07", "EF67LP32"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Revisão das funções sintáticas (termos essenciais e integrantes da oração).\nRevisão (tipos de predicado).\n"
        "Aposto e vocativo.\nPredicativo do sujeito e do objeto.\nAdjunto adnominal do objeto.\n"
        "Complemento nominal (apresentar a diferença entre complemento nominal x objeto indireto x adjunto "
        "adnominal).\nOrtografia (uso do S e Z nas terminações ES / ESA, EZ / EZA).\n"
        "Pontuação do período simples e composto.",
        "habilidades": ["EF08LP07", "EF08LP06", "EF69LP55", "EF08LP09", "EF69LP56", "EF08LP04"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Adjunto adverbial.\nSemântica (polissemia, ambiguidade, sinonímia, antonímia, hiponímia, hiperonímia, "
        "paronímia e homonímia).\nVozes verbais / Agente da passiva.\nGrafia dos verbos abundantes.",
        "habilidades": ["EF08LP10", "EF08LP16", "EF69LP17", "EF69LP56", "EF08LP08", "EF08LP14", "EF89LP16"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Conotação e Denotação.\nFiguras de linguagem.\n"
        "Conjunções coordenativas (uso e valor semântico)/Período composto por coordenação/Orações coordenadas.\n"
        "Concordância verbal e nominal (noções).\nOrtografia (uso dos porquês).",
        "habilidades": ["EF69LP55", "EF08LP14", "EF08LP15", "EF69LP18", "EF08LP13", "EF08LP11", "EF08LP04", "EF89LP29"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Revisão dos termos da oração.\nRevisão das orações coordenadas.\nOrações subordinadas substantivas.\n"
        "Pronome relativo (função sintática).\nOrações subordinadas adjetivas (valor semântico da vírgula).\n"
        "Ortografia (uso de este, esse, aquele e variações).",
        "habilidades": ["EF09LP04", "EF09LP05", "EF09LP06", "EF09LP08", "EF09LP11", "EF09LP09", "EF69LP55", "EF69LP56"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Concordância nominal (casos relevantes).\nOrações subordinadas adverbiais.\nOrações reduzidas.\n"
        "Colocação pronominal.\nOrtografia (uso de Onde/Aonde; Se não/Senão; Eu/Mim; Trás/Traz).",
        "habilidades": ["EF09LP04", "EF09LP08", "EF09LP11", "EF09LP10", "EF69LP56"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Análise Linguística/Semiótica (Gramática)", "objetos":
        "Concordância verbal (casos relevantes).\n"
        "Ortografia (A fim/Afim; Ao invés de/Em vez de; A par/Ao par; Ao encontro de/De encontro a).\n"
        "Regência verbal e nominal.\nCrase.\nFormação de palavras (revisão).\nFonética e Fonologia (revisão).\n"
        "Figuras de sintaxe (elipse / zeugma / silepse / hipérbato ou inversão / pleonasmo / assíndeto / "
        "polissíndeto / anáfora / anacoluto).",
        "habilidades": ["EF09LP04", "EF09LP07", "EF09LP12", "EF69LP40", "EF69LP56", "EF89LP30", "EF89LP37"]},

    # =================== LEITURA ===================
    {"ano": "6º ano", "trimestre": 1, "unidade": "Leitura", "objetos":
        "Lenda x Mito (trabalhar a figura de linguagem Personificação).\n"
        "Contos populares (contos de fadas e/ou contos maravilhosos observando: tipos de discurso e elementos "
        "da narrativa).\nNarrativa de aventura (paradidático).\nRomance / Novela / Antologia de crônicas ou contos.\n"
        "Provérbios e ditos populares.",
        "habilidades": ["EF67LP27", "EF67LP28", "EF69LP44", "EF69LP46", "EF69LP47", "EF69LP48", "EF69LP53", "EF69LP54"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Leitura", "objetos":
        "Relato pessoal.\nAutobiografia.\nDiário.\nReportagem.\nCampanha publicitária (propaganda - slogan - "
        "anúncio).\nBlog.",
        "habilidades": ["EF06LP01", "EF06LP02", "EF67LP03", "EF67LP04", "EF67LP06", "EF69LP03", "EF67LP02", "EF69LP02", "EF69LP04"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Leitura", "objetos":
        "Carta do leitor.\nCarta pessoal.\nE-mail.\nArtigo de divulgação científica.\n"
        "Verbete (dicionário e enciclopédico).\nResenha.",
        "habilidades": ["EF67LP05", "EF67LP16", "EF67LP18", "EF69LP31", "EF67LP20", "EF67LP26", "EF69LP29",
                        "EF69LP32", "EF69LP34", "EF69LP45"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Leitura", "objetos":
        "Cordel x poesia trovadoresca (noções de versificação).\n"
        "Figuras de linguagem (metáfora, ironia, assonância, aliteração, sinestesia).\n"
        "Conto / Crônica de humor (leitura dramatizada / tipos de discurso e elementos da narrativa).\n"
        "Romance / Novela / Antologia de crônicas ou contos.",
        "habilidades": ["EF67LP27", "EF67LP38", "EF69LP48", "EF69LP54", "EF69LP47", "EF69LP53", "EF67LP28", "EF69LP44", "EF69LP49", "EF69LP46"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Leitura", "objetos":
        "Notícia (fake news).\nCartum x Charge.\nCrônica jornalística.\nResenha crítica x comentário.\n"
        "Carta denúncia.\nArtigo de opinião.",
        "habilidades": ["EF07LP01", "EF07LP02", "EF67LP03", "EF67LP04", "EF69LP05", "EF67LP08", "EF69LP01",
                        "EF69LP03", "EF67LP07", "EF69LP34", "EF67LP17", "EF67LP18", "EF67LP20", "EF67LP16",
                        "EF67LP26", "EF69LP16", "EF69LP30", "EF69LP42"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Leitura", "objetos":
        "Textos instrucionais (manual de instrução / regras de jogo / receita...).\nEstatutos.\nDepoimento.\n"
        "Artigo de divulgação científica.",
        "habilidades": ["EF67LP15", "EF69LP20", "EF69LP21", "EF69LP24", "EF69LP29", "EF69LP40"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Leitura", "objetos":
        "Conto de suspense (tipos de discurso e elementos da narrativa).\nAntologia de crônicas ou contos.\n"
        "Novela ou Romance [de ficção científica / suspense] (paradidático).\nPoema verbal e visual.",
        "habilidades": ["EF69LP44", "EF69LP46", "EF69LP47", "EF69LP49", "EF69LP53", "EF69LP54", "EF89LP32", "EF89LP33", "EF89LP37", "EF69LP48"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Leitura", "objetos":
        "Anúncio / Campanha publicitária.\nCharge x Cartum.\nCarta do leitor / Carta denúncia.",
        "habilidades": ["EF08LP01", "EF08LP02", "EF69LP02", "EF69LP04", "EF89LP01", "EF89LP05", "EF89LP07",
                        "EF69LP05", "EF89LP02", "EF69LP13", "EF89LP03", "EF89LP04", "EF89LP06"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Leitura", "objetos":
        "Documentos legais e normativos (declaração dos direitos humanos / petição / legislações).\nRelatório.\n"
        "Carta argumentativa.\nDissertação acadêmica.",
        "habilidades": ["EF69LP20", "EF69LP27", "EF69LP28", "EF89LP17", "EF69LP32", "EF89LP23", "EF89LP24", "EF69LP21", "EF69LP29"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Leitura", "objetos":
        "Roteiro de TV/Cinema.\nRomance (paradidático).\nRomance/Novela/Antologia de crônicas ou contos.\n"
        "Conto Social e Psicológico (tipos de discurso e elementos da narrativa).",
        "habilidades": ["EF89LP32", "EF89LP34", "EF69LP46", "EF89LP33", "EF69LP49", "EF69LP44", "EF69LP47", "EF69LP53", "EF69LP48", "EF69LP54", "EF89LP37"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Leitura", "objetos":
        "Texto informativo (reportagem / textos didáticos / anúncios / comunicados...).\n"
        "Crônica jornalística (esportiva).\nEditorial.\nCarta Aberta/Manifesto.",
        "habilidades": ["EF69LP03", "EF69LP04", "EF09LP01", "EF09LP02", "EF89LP01", "EF89LP16", "EF89LP03",
                        "EF89LP05", "EF89LP06", "EF89LP04", "EF89LP18", "EF89LP19", "EF89LP20", "EF69LP21"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Leitura", "objetos":
        "Artigo científico.\nResenha crítica.\nArtigo de Opinião / Artigo de Lei.\nBiografia (foco no ambiente acadêmico).",
        "habilidades": ["EF69LP29", "EF89LP17", "EF89LP24", "EF69LP30", "EF89LP28", "EF69LP34", "EF69LP32", "EF69LP20", "EF69LP27", "EF69LP28", "EF69LP42"]},

    # =================== PRODUÇÃO DE TEXTO ===================
    {"ano": "6º ano", "trimestre": 1, "unidade": "Produção de Texto", "objetos":
        "História em quadrinhos (trabalhar a figura de linguagem onomatopeia).\n"
        "Transformar conto em HQ (sequência dos gêneros trabalhados na frente de leitura).\n"
        "Poema / Classificado poético.\nPontuação e ortografia.",
        "habilidades": ["EF67LP30", "EF67LP31", "EF69LP51"]},
    {"ano": "6º ano", "trimestre": 2, "unidade": "Produção de Texto", "objetos":
        "Notícia/Entrevista (Jornal falado).\nCampanha publicitária (valorização do ambiente escolar).\n"
        "Foto denúncia (sobre o ambiente escolar).\nPontuação e Ortografia.",
        "habilidades": ["EF67LP09", "EF67LP10", "EF67LP14", "EF69LP09", "EF67LP13", "EF69LP06", "EF67LP19", "EF69LP07", "EF67LP32", "EF69LP10"]},
    {"ano": "6º ano", "trimestre": 3, "unidade": "Produção de Texto", "objetos":
        "Resumo / Relato.\nNormas de conduta no ambiente escolar (organização, disciplina, cronograma de "
        "estudo, ambiente escolar...) - (Retomada do trabalho do SOE).\n"
        "Carta de reclamação (de acordo com as normas de conduta propostas).\n"
        "Debate (sobre as normas de conduta no ambiente escolar).\nParáfrase.\n"
        "Biografia (hiperlink – hipertexto para conhecimento).",
        "habilidades": ["EF67LP22", "EF69LP20", "EF69LP23", "EF69LP22", "EF67LP24", "EF69LP25", "EF69LP26", "EF67LP25", "EF69LP35"]},

    {"ano": "7º ano", "trimestre": 1, "unidade": "Produção de Texto", "objetos":
        "Texto teatral (oral) / esquete.\nPoema narrativo.\nCordel e poesia (sarau).\nMemórias literárias (oficinas OLP).",
        "habilidades": ["EF67LP29", "EF69LP52", "EF67LP31", "EF69LP50"]},
    {"ano": "7º ano", "trimestre": 2, "unidade": "Produção de Texto", "objetos":
        "Entrevista / Reportagem (Telerreportagem oral – escrita).\nPontuação e Ortografia.\n"
        "Infográfico (criar infográfico para compor uma reportagem).\nPropaganda (cartaz/ folheto).",
        "habilidades": ["EF67LP14", "EF69LP10", "EF69LP08", "EF67LP13", "EF69LP06", "EF69LP07", "EF69LP09"]},
    {"ano": "7º ano", "trimestre": 3, "unidade": "Produção de Texto", "objetos":
        "Carta de reclamação x Carta de solicitação (argumentação).\nRegimentos (assembleia - oralidade).\n"
        "Roteiro de Podcast científico (gênero oral).\nRelatório (visita, aula, etc...).\n"
        "Fichamento (obra literária do bimestre).",
        "habilidades": ["EF67LP17", "EF67LP18", "EF67LP19", "EF69LP22", "EF67LP23", "EF69LP24", "EF69LP26", "EF67LP21", "EF69LP35", "EF69LP37", "EF69LP36"]},

    {"ano": "8º ano", "trimestre": 1, "unidade": "Produção de Texto", "objetos":
        "Canção / paródia / paráfrase (intertextualidade).\nPontuação e Ortografia.\n"
        "Criação de texto dramático a partir de crônicas literárias.",
        "habilidades": ["EF69LP51", "EF89LP36", "EF69LP50", "EF69LP52"]},
    {"ano": "8º ano", "trimestre": 2, "unidade": "Produção de Texto", "objetos":
        "Memes.\nNotícia (vídeo minuto).\nEditorial.",
        "habilidades": ["EF69LP05", "EF69LP06", "EF69LP07", "EF69LP08", "EF69LP17", "EF69LP18"]},
    {"ano": "8º ano", "trimestre": 3, "unidade": "Produção de Texto", "objetos":
        "Artigo de opinião.\nResenha crítica.\nDissertação escolar.\nRegras de debate.\nSeminário.",
        "habilidades": ["EF08LP03", "EF89LP10", "EF69LP08", "EF69LP43", "EF69LP13", "EF69LP15", "EF89LP12", "EF69LP07", "EF89LP26"]},

    {"ano": "9º ano", "trimestre": 1, "unidade": "Produção de Texto", "objetos":
        "Conto (terror, humor, social...).\nCrônica de ficção científica.\nPontuação e ortografia.",
        "habilidades": ["EF89LP35", "EF69LP51", "EF69LP53"]},
    {"ano": "9º ano", "trimestre": 2, "unidade": "Produção de Texto", "objetos":
        "Infográfico.\nProdução de reportagem impressa e digital.\nEntrevista oral e escrita.",
        "habilidades": ["EF69LP06", "EF69LP07", "EF69LP08", "EF89LP08", "EF89LP09", "EF69LP12", "EF89LP13", "EF69LP10"]},
    {"ano": "9º ano", "trimestre": 3, "unidade": "Produção de Texto", "objetos":
        "Debate / mesa redonda (oral).\nSeminário (oral – foco no mundo do trabalho) ou júri simulado.\n"
        "Enquete/Pesquisa de opinião (foco no ambiente acadêmico).\n"
        "Pesquisa de opinião (foco na inserção do jovem no mercado de trabalho).\nCurrículo.\n"
        "Entrevista de emprego (oral).",
        "habilidades": ["EF89LP12", "EF89LP14", "EF89LP15", "EF69LP25", "EF69LP26", "EF69LP38", "EF69LP41",
                        "EF89LP21", "EF69LP22", "EF89LP22", "EF89LP25", "EF69LP36", "EF69LP37", "EF69LP39", "EF69LP43"]},
]


def _migrar_atividades_por_docente(conn):
    """Copia o texto que já existia em documento_norteador_semanas.atividade pra nova
    tabela por docente (documento_norteador_semana_atividades), uma cópia pra cada
    docente vinculado ao ano daquela semana — assim ninguém perde o que já tinha
    escrito quando essa separação (atividade por professor) entrou. Idempotente via
    INSERT OR IGNORE: não sobrescreve nada que algum docente já tenha editado na tabela
    nova, e roda de novo a cada start só pra cobrir docentes vinculados depois da
    primeira migração (06/09/2026)."""
    semanas_com_atividade = conn.execute(
        "SELECT id, documento_id, ano_escolaridade, atividade FROM documento_norteador_semanas WHERE atividade IS NOT NULL AND atividade != ''"
    ).fetchall()
    for s in semanas_com_atividade:
        docentes = conn.execute(
            "SELECT id FROM documento_norteador_docentes WHERE documento_id=? AND ano_escolaridade=?",
            (s["documento_id"], s["ano_escolaridade"])
        ).fetchall()
        for d in docentes:
            conn.execute(
                "INSERT OR IGNORE INTO documento_norteador_semana_atividades (semana_id, docente_vinculo_id, atividade, atualizado_em) VALUES (?, ?, ?, CURRENT_TIMESTAMP)",
                (s["id"], d["id"], s["atividade"])
            )


# Consolidação de disciplinas duplicadas (06/09/2026, a pedido de Felipe): o Documento
# Norteador acabou criando linhas de disciplina com nomes diferentes do mesmo componente
# curricular ("Português" x "Língua Portuguesa", "Inglês" x "Língua Inglesa", "Ed. Física"
# x "Educação Física"/variações). O nome "sobrevivente" de cada grupo é sempre o que já é
# usado pelo sistema de Boletim/Conselho de Classe (que casa com a exportação oficial da
# rede/e-cidade e tem lógica própria espalhada pelo código) — trocar esse nome quebraria
# a importação e recriaria a duplicata a cada import. Isso é "menor impacto": o Boletim
# nem percebe a mudança, e o Documento Norteador passa a ter uma única disciplina por
# componente, com o Referencial Curricular corretamente ligado a ela.
DISCIPLINAS_UNIFICAR = {
    "Português": ["Língua Portuguesa", "Lingua Portuguesa", "Portugues"],
    "Inglês": ["Língua Inglesa", "Lingua Inglesa", "Ingles"],
    "Ed. Física": ["Educação Física", "Educação Fisica", "Educacao Fisica", "Educacao Física", "Ed. Fisica", "Ed Física", "Ed Fisica"],
}

# Tabelas que têm uma coluna disciplina_id — usadas tanto pela consolidação acima quanto
# como referência de quais tabelas precisam ser repontadas ao unificar duas disciplinas.
_TABELAS_COM_DISCIPLINA_ID = [
    "questoes", "referencial_curricular", "documentos_norteadores",
    "simulado_blocos", "boletim_medias", "boletim_faltas",
    "boletim_analise", "boletim_professor_turma",
]


def _consolidar_disciplinas_duplicadas(conn):
    """Une disciplinas duplicadas (mesmo componente curricular, nomes diferentes) numa
    única linha, migrando tudo que apontava pras duplicadas pro nome canônico (o que já
    é usado pelo Boletim). Idempotente — depois da primeira vez que roda, não encontra
    mais duplicata nenhuma e não faz nada (06/09/2026)."""
    for nome_canonico, variantes in DISCIPLINAS_UNIFICAR.items():
        ids_do_grupo = []
        vistos = set()
        for nome in [nome_canonico] + variantes:
            row = conn.execute("SELECT id FROM disciplinas WHERE nome = ?", (nome,)).fetchone()
            if row and row["id"] not in vistos:
                ids_do_grupo.append(row["id"])
                vistos.add(row["id"])
        if len(ids_do_grupo) <= 1:
            if ids_do_grupo:
                conn.execute("UPDATE disciplinas SET nome = ? WHERE id = ?", (nome_canonico, ids_do_grupo[0]))
            continue

        sobrevivente_id = ids_do_grupo[0]
        for dup_id in ids_do_grupo[1:]:
            for tabela in _TABELAS_COM_DISCIPLINA_ID:
                if tabela == "boletim_professor_turma":
                    # UNIQUE(professor_id, turma_id, disciplina_id) — se já existir a
                    # combinação no sobrevivente, descarta a duplicada em vez de repontar.
                    for r in conn.execute("SELECT id, professor_id, turma_id FROM boletim_professor_turma WHERE disciplina_id = ?", (dup_id,)).fetchall():
                        existe = conn.execute(
                            "SELECT 1 FROM boletim_professor_turma WHERE professor_id=? AND turma_id=? AND disciplina_id=?",
                            (r["professor_id"], r["turma_id"], sobrevivente_id)
                        ).fetchone()
                        if existe:
                            conn.execute("DELETE FROM boletim_professor_turma WHERE id = ?", (r["id"],))
                        else:
                            conn.execute("UPDATE boletim_professor_turma SET disciplina_id = ? WHERE id = ?", (sobrevivente_id, r["id"]))
                elif tabela == "referencial_curricular":
                    # Evita duplicar blocos idênticos (mesmo ano+trimestre+unidade
                    # temática) — mantém o bloco do sobrevivente quando já existir igual.
                    for r in conn.execute(
                        "SELECT id, ano_escolaridade, trimestre, unidade_tematica FROM referencial_curricular WHERE disciplina_id = ?", (dup_id,)
                    ).fetchall():
                        existe = conn.execute(
                            "SELECT id FROM referencial_curricular WHERE disciplina_id=? AND ano_escolaridade=? AND trimestre=? AND unidade_tematica=?",
                            (sobrevivente_id, r["ano_escolaridade"], r["trimestre"], r["unidade_tematica"])
                        ).fetchone()
                        if existe:
                            conn.execute("DELETE FROM referencial_curricular_habilidades WHERE referencial_id = ?", (r["id"],))
                            conn.execute("DELETE FROM referencial_curricular WHERE id = ?", (r["id"],))
                        else:
                            conn.execute("UPDATE referencial_curricular SET disciplina_id = ? WHERE id = ?", (sobrevivente_id, r["id"]))
                elif tabela == "documentos_norteadores":
                    # Se já existir um Documento Norteador igual (mesmo trimestre+ano) no
                    # sobrevivente, não mexe nesse — fica pro admin resolver na mão (só
                    # acontece se os DOIS nomes já tinham documento cadastrado pro mesmo
                    # período, caso raro).
                    for r in conn.execute("SELECT id, trimestre, ano_letivo FROM documentos_norteadores WHERE disciplina_id = ?", (dup_id,)).fetchall():
                        existe = conn.execute(
                            "SELECT id FROM documentos_norteadores WHERE disciplina_id=? AND trimestre=? AND ano_letivo=?",
                            (sobrevivente_id, r["trimestre"], r["ano_letivo"])
                        ).fetchone()
                        if not existe:
                            conn.execute("UPDATE documentos_norteadores SET disciplina_id = ? WHERE id = ?", (sobrevivente_id, r["id"]))
                else:
                    conn.execute(f"UPDATE {tabela} SET disciplina_id = ? WHERE disciplina_id = ?", (sobrevivente_id, dup_id))

            # Só apaga a duplicata se realmente não sobrou nenhuma referência a ela —
            # evita deixar dado orfão no caso raro de conflito acima.
            ainda_referenciado = any(
                conn.execute(f"SELECT 1 FROM {t} WHERE disciplina_id = ? LIMIT 1", (dup_id,)).fetchone()
                for t in _TABELAS_COM_DISCIPLINA_ID
            )
            if not ainda_referenciado:
                conn.execute("DELETE FROM disciplinas WHERE id = ?", (dup_id,))

        conn.execute("UPDATE disciplinas SET nome = ? WHERE id = ?", (nome_canonico, sobrevivente_id))


def _seed_referencial_curricular(conn):
    """Popula o Referencial Curricular a partir das listas REFERENCIAL_CURRICULAR_*
    (uma por disciplina). Idempotente — só insere o que ainda não existe, então pode
    rodar em todo startup sem duplicar nem sobrescrever edições feitas pela tela depois.
    Habilidades BNCC que já existem em habilidades_bncc não são alteradas; só cria a
    linha se o código realmente não existir ainda (04/09/2026)."""
    for disciplina_nome, blocos in [
        ("Arte", REFERENCIAL_CURRICULAR_ARTE),
        ("Ed. Física", REFERENCIAL_CURRICULAR_EDUCACAO_FISICA),
        ("Ciências", REFERENCIAL_CURRICULAR_CIENCIAS),
        ("História", REFERENCIAL_CURRICULAR_HISTORIA),
        ("Geografia", REFERENCIAL_CURRICULAR_GEOGRAFIA),
        ("Matemática", REFERENCIAL_CURRICULAR_MATEMATICA),
        ("Inglês", REFERENCIAL_CURRICULAR_INGLES),
        ("Português", REFERENCIAL_CURRICULAR_PORTUGUES),
    ]:
        disc = conn.execute("SELECT id FROM disciplinas WHERE nome = ?", (disciplina_nome,)).fetchone()
        if not disc:
            cur = conn.execute("INSERT INTO disciplinas (nome) VALUES (?)", (disciplina_nome,))
            disciplina_id = cur.lastrowid
        else:
            disciplina_id = disc["id"]

        for bloco in blocos:
            existe = conn.execute(
                "SELECT id FROM referencial_curricular WHERE disciplina_id=? AND ano_escolaridade=? AND trimestre=? AND unidade_tematica=?",
                (disciplina_id, bloco["ano"], bloco["trimestre"], bloco["unidade"])
            ).fetchone()
            if existe:
                continue
            cur = conn.execute(
                "INSERT INTO referencial_curricular (disciplina_id, ano_escolaridade, trimestre, unidade_tematica, objetos_conhecimento, ordem) VALUES (?, ?, ?, ?, ?, ?)",
                (disciplina_id, bloco["ano"], bloco["trimestre"], bloco["unidade"], bloco["objetos"], bloco["trimestre"])
            )
            referencial_id = cur.lastrowid
            for codigo in bloco["habilidades"]:
                hab = conn.execute("SELECT id FROM habilidades_bncc WHERE codigo = ?", (codigo,)).fetchone()
                if not hab:
                    cur2 = conn.execute("INSERT INTO habilidades_bncc (codigo, descricao) VALUES (?, ?)", (codigo, None))
                    habilidade_id = cur2.lastrowid
                else:
                    habilidade_id = hab["id"]
                conn.execute(
                    "INSERT INTO referencial_curricular_habilidades (referencial_id, habilidade_id) VALUES (?, ?)",
                    (referencial_id, habilidade_id)
                )


def _gerar_username(nome_completo: str, conn) -> str:
    """Gera um username único a partir do nome (primeiro.ultimo, minúsculo, sem acento),
    adicionando um número se já existir — 27/08/2026."""
    base = _boletim_normalizar(nome_completo).replace(" ", ".")
    base = "".join(c for c in base if c.isalnum() or c == ".")
    candidato = base
    n = 1
    while conn.execute("SELECT 1 FROM professores WHERE username = ?", (candidato,)).fetchone():
        n += 1
        candidato = f"{base}{n}"
    return candidato


def _drive_upload_arquivo(nome_arquivo: str, conteudo_bytes: bytes, mime_type: str, folder_id: Optional[str] = None):
    """Sobe um arquivo pra uma pasta do Google Drive. Se folder_id não for passado, usa
    GOOGLE_DRIVE_FOLDER_ID (pasta dos afastamentos de profissionais) — mantido assim pra
    não quebrar quem já chama essa função sem passar pasta. Pra outras pastas (ex: atestados
    de alunos), passe folder_id explicitamente — 27/08/2026.

    Usa a IDENTIDADE DA PRÓPRIA VM (Application Default Credentials) em vez de uma chave
    JSON baixada — não precisa de service account key (a organização pode bloquear a
    criação dessas chaves por política de segurança, como aconteceu aqui em 25/08/2026).
    A VM precisa ter o escopo do Drive habilitado (drive, não só cloud-platform — o
    Drive fica de fora do cloud-platform) e a pasta compartilhada com o e-mail da conta
    de serviço da própria VM. Se GOOGLE_DRIVE_CREDENTIALS_JSON estiver definida, ainda é
    aceita como alternativa (ambientes fora do GCP, ou se a política mudar).

    IMPORTANTE sobre o escopo: usamos 'drive' (completo) e não 'drive.file'. O escopo
    'drive.file' só enxerga arquivos/pastas que O PRÓPRIO APP criou (ou que foram abertos
    via seletor do Google) — uma pasta compartilhada manualmente pela interface do Drive
    fica invisível pra esse escopo, mesmo com a permissão de Editor certinha. Descobrimos
    isso em produção em 25/08/2026 (erro 404 'File not found' na pasta, mesmo compartilhada
    corretamente) — trocar pra 'drive' resolve.

    Retorna (file_id, link, erro). Se não der pra autenticar, retorna erro claro em vez de
    quebrar — o registro do afastamento é salvo de qualquer forma."""
    folder_id = folder_id or GOOGLE_DRIVE_FOLDER_ID
    if not folder_id:
        return None, None, "Google Drive não configurado ainda (falta o ID da pasta no servidor)."
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        import io as _io

        DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
        creds = None
        if GOOGLE_DRIVE_CREDENTIALS_JSON:
            from google.oauth2 import service_account
            info = json.loads(GOOGLE_DRIVE_CREDENTIALS_JSON)
            creds = service_account.Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)
        else:
            import google.auth
            creds, _ = google.auth.default(scopes=DRIVE_SCOPES)

        service = build("drive", "v3", credentials=creds)
        metadata = {"name": nome_arquivo, "parents": [folder_id]}
        media = MediaIoBaseUpload(_io.BytesIO(conteudo_bytes), mimetype=mime_type, resumable=False)
        # supportsAllDrives=True é obrigatório se a pasta estiver dentro de um Drive
        # Compartilhado (não "Meu Drive") — sem isso a API nem enxerga a pasta e retorna
        # 404 "File not found", mesmo com permissão certa (achado em produção, 25/08/2026).
        arquivo = service.files().create(
            body=metadata, media_body=media, fields="id, webViewLink", supportsAllDrives=True
        ).execute()
        return arquivo.get("id"), arquivo.get("webViewLink"), None
    except ImportError:
        return None, None, "Biblioteca do Google Drive não instalada no servidor (google-api-python-client / google-auth)."
    except Exception as e:
        import traceback
        # Algumas exceções do google-auth/googleapiclient têm str(e) vazio — cai pro
        # repr(e) (mostra ao menos o tipo da exceção) e, se for erro HTTP da API,
        # extrai o motivo (25/08/2026, depois de ver esse caso em produção).
        detalhe = str(e).strip() or repr(e)
        try:
            from googleapiclient.errors import HttpError
            if isinstance(e, HttpError):
                motivo = e._get_reason().strip() if hasattr(e, "_get_reason") else ""
                detalhe = f"HTTP {e.resp.status} — {motivo or e.content}"
        except Exception:
            pass
        # Log completo (com traceback) vai pro journalctl, mesmo que a mensagem mostrada
        # ao usuário seja curta — é o que a gente olha pra diagnosticar de verdade.
        print(f"[Drive upload] Falha ao enviar '{nome_arquivo}': {detalhe}")
        print(traceback.format_exc())
        return None, None, f"Erro ao enviar pro Google Drive: {detalhe}"


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
    permitidas = {"strong", "b", "em", "i", "u", "br", "p", "div", "span", "ul", "ol", "li", "blockquote",
                   "table", "thead", "tbody", "tr", "th", "td", "img", "figure", "figcaption"}
    def _filtrar_tag(m):
        tag_full = m.group(0)
        tag_name = m.group(1).lower()
        if tag_name not in permitidas:
            return ""
        # img: mantém src e alt, remove outros atributos perigosos
        if tag_name == "img":
            src = _re.search(r'src\s*=\s*["\']([^"\'>]+)["\']', tag_full)
            if not src: return ""
            alt = _re.search(r'alt\s*=\s*["\']([^"\'>]*)["\' ]', tag_full)
            alt_val = alt.group(1) if alt else ""
            return f'<img src="{src.group(1)}" alt="{alt_val}" style="max-width:100%; height:auto;">'
        # table/th/td: mantém style de bordas
        if tag_name in ("table", "th", "td"):
            style = _re.search(r'style\s*=\s*["\']([^"\'>]+)["\' ]', tag_full)
            style_attr = f' style="{style.group(1)}"' if style else ""
            if tag_full.startswith("</"): return f"</{tag_name}>"
            return f'<{tag_name}{style_attr}>'
        # thead/tbody/tr/figure/figcaption: sem atributos
        if tag_name in ("thead", "tbody", "tr", "figure", "figcaption"):
            if tag_full.startswith("</"): return f"</{tag_name}>"
            return f"<{tag_name}>"
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




_JS_MATH_BUTTONS = r"""
            function _inserirTexto(editor, sync, txt) {
                editor.focus();
                var sel = window.getSelection();
                if (sel && sel.rangeCount) {
                    var rng = sel.getRangeAt(0);
                    if (!editor.contains(rng.commonAncestorContainer)) {
                        rng = document.createRange();
                        rng.selectNodeContents(editor);
                        rng.collapse(false);
                    }
                    rng.deleteContents();
                    var node = document.createTextNode(txt);
                    rng.insertNode(node);
                    rng.setStartAfter(node);
                    rng.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(rng);
                } else {
                    document.execCommand("insertText", false, txt);
                }
                sync();
            }
            var btnFrac = toolbar.querySelector(".btn-insert-frac");
            if (btnFrac) {
                btnFrac.addEventListener("click", function(e) {
                    e.preventDefault();
                    var num = prompt("Numerador da fração:");
                    if (num === null) return;
                    var den = prompt("Denominador da fração:");
                    if (den === null) return;
                    _inserirTexto(editor, sync, "$\\frac{" + num + "}{" + den + "}$");
                });
            }
            var btnPot = toolbar.querySelector(".btn-insert-pot");
            if (btnPot) {
                btnPot.addEventListener("click", function(e) {
                    e.preventDefault();
                    var base = prompt("Base (ex: 2, x, 2x):");
                    if (base === null) return;
                    var expoente = prompt("Expoente (ex: 2, 3, n):");
                    if (expoente === null) return;
                    _inserirTexto(editor, sync, "$" + base + "^{" + expoente + "}$");
                });
            }
            var btnTab = toolbar.querySelector(".btn-insert-tab");
            if (btnTab) {
                btnTab.addEventListener("click", function(e) {
                    e.preventDefault();
                    var nlin = parseInt(prompt("Número de linhas:", "3"));
                    if (!nlin || nlin < 1) return;
                    var ncol = parseInt(prompt("Número de colunas:", "3"));
                    if (!ncol || ncol < 1) return;
                    var tbl = '<table style="border-collapse:collapse;width:100%;margin:8px 0;">';
                    tbl += "<thead><tr>";
                    for (var c = 0; c < ncol; c++) {
                        tbl += '<th style="border:1px solid #999;padding:6px 10px;background:#f0f0f0;font-weight:600;">Col ' + (c+1) + "</th>";
                    }
                    tbl += "</tr></thead><tbody>";
                    for (var r = 0; r < nlin - 1; r++) {
                        tbl += "<tr>";
                        for (var c2 = 0; c2 < ncol; c2++) {
                            tbl += '<td style="border:1px solid #999;padding:6px 10px;">&nbsp;</td>';
                        }
                        tbl += "</tr>";
                    }
                    tbl += "</tbody></table><p></p>";
                    document.execCommand("insertHTML", false, tbl);
                    sync();
                });
            }
"""

_JS_DETECTAR_ALTS = r"""
            function detectarAlternativas(texto) {
                texto = texto.replace(/\r\n/g, '\n').replace(/\u00A0/g, ' ').trim();
                var padrao = /(?:^|\n)[ \t]*[(]?([A-Da-d])[)]?[ \t]*[-).,:][ \t]*/g;
                var matches = Array.from(texto.matchAll(padrao));
                var idxA=-1, idxB=-1, idxC=-1, idxD=-1;
                for (var mi=0; mi<matches.length; mi++) {
                    var letra = matches[mi][1].toUpperCase();
                    var pos = matches[mi].index;
                    if (letra==='A' && idxA===-1) idxA=pos;
                    else if (letra==='B' && idxB===-1 && idxA!==-1 && pos>idxA) idxB=pos;
                    else if (letra==='C' && idxC===-1 && idxB!==-1 && pos>idxB) idxC=pos;
                    else if (letra==='D' && idxD===-1 && idxC!==-1 && pos>idxC) idxD=pos;
                }
                if (idxA===-1 || idxB===-1 || idxC===-1 || idxD===-1) return null;
                var enunciado = texto.slice(0, idxA).trim();
                function ext(s,e) { return texto.slice(s,e).replace(/^\n?[ \t]*[(]?[A-Da-d][)]?[ \t]*[-).,:][ \t]*/, "").trim(); }
                return { enunciado:enunciado, alternativas:[ext(idxA,idxB),ext(idxB,idxC),ext(idxC,idxD),ext(idxD,texto.length)] };
            }
            function aplicarAlternativas(texto) {
                var r = detectarAlternativas(texto);
                if (!r) { document.execCommand("insertText", false, texto); return; }
                var trunc = function(s) { return s.length > 60 ? s.slice(0,60)+"..." : s; };
                var nl = "\n";
                var msg = "Detectei 4 alternativas. Aplicar automaticamente?" + nl + nl
                        + (r.enunciado ? "Enunciado: " + trunc(r.enunciado) + nl : "")
                        + "A) " + trunc(r.alternativas[0]) + nl
                        + "B) " + trunc(r.alternativas[1]) + nl
                        + "C) " + trunc(r.alternativas[2]) + nl
                        + "D) " + trunc(r.alternativas[3]);
                if (!confirm(msg)) { document.execCommand("insertText", false, texto); return; }
                editor.innerHTML = r.enunciado ? r.enunciado.replace(/\n/g, "<br>") : "";
                hidden.value = editor.innerHTML;
                refreshPlaceholder();
                ["a","b","c","d"].forEach(function(letra, idx) {
                    var altEd = document.querySelector(".editor-content[data-target=\"alt_"+letra+"\"]");
                    var altHid = document.getElementById("alt_"+letra+"_hidden");
                    if (altEd && altHid) {
                        altEd.innerHTML = r.alternativas[idx].replace(/\n/g, "<br>");
                        altHid.value = altEd.innerHTML;
                        altEd.removeAttribute("data-ph-shown");
                    }
                });
            }
            editor.addEventListener("paste", function(e) {
                var cb = e.clipboardData || window.clipboardData;
                if (!cb) return;
                var items = cb.items ? Array.from(cb.items) : [];
                var imgItem = items.find(function(it) { return it.type.startsWith("image/"); });
                if (imgItem) {
                    e.preventDefault();
                    var blob = imgItem.getAsFile();
                    if (!blob) return;
                    var fd = new FormData();
                    fd.append("arquivo", blob, "imagem_colada.png");
                    fetch("/upload-imagem-inline", { method: "POST", body: fd })
                        .then(function(r) { return r.json(); })
                        .then(function(data) {
                            if (data.url) {
                                document.execCommand("insertHTML", false,
                                    "<img src=\"" + data.url + "\" style=\"max-width:100%; height:auto; display:block; margin:4px 0;\" alt=\"\">");
                                sync();
                            }
                        })
                        .catch(function() { alert("Erro ao fazer upload da imagem."); });
                    return;
                }
                var texto = cb.getData("text/plain") || "";
                if (!texto) return;
                e.preventDefault();
                aplicarAlternativas(texto);
            });
"""

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
    bot_fracao = f'<button type="button" class="btn-insert-frac" title="Inserir fração como $\\frac{{num}}{{den}}$" style="{btn_style}">½ fração</button>'
    bot_potencia = f'<button type="button" class="btn-insert-pot" title="Inserir potência como $base^{{exp}}$" style="{btn_style}">x² potência</button>'
    bot_tabela = f'<button type="button" class="btn-insert-tab" title="Inserir tabela" style="{btn_style}">⊞ tabela</button>'

    toolbar_buttons = bot_basicos + sep + bot_fracao + bot_potencia + bot_tabela + sep + bot_limpar if compact else bot_basicos + sep + bot_extra + bot_fracao + bot_potencia + bot_tabela + sep + bot_limpar

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
            {_JS_MATH_BUTTONS}

            {_JS_DETECTAR_ALTS if detectar_alternativas else ""}

            // Paste de imagem (todos os campos, incluindo alternativas)
            editor.addEventListener('paste', function(e) {{
                var cb = e.clipboardData || window.clipboardData;
                if (!cb) return;
                var items = cb.items ? Array.from(cb.items) : [];
                var imgItem = items.find(function(it) {{ return it.type.startsWith('image/'); }});
                if (!imgItem) return;  // texto é tratado pelo handler acima (se existir) ou pelo browser
                e.preventDefault();
                var blob = imgItem.getAsFile();
                if (!blob) return;
                var fd = new FormData();
                fd.append('arquivo', blob, 'imagem_colada.png');
                fetch('/upload-imagem-inline', {{ method: 'POST', body: fd }})
                    .then(function(r) {{ return r.json(); }})
                    .then(function(data) {{
                        if (data.url) {{
                            document.execCommand('insertHTML', false,
                                '<img src="' + data.url + '" style="max-width:100%; height:auto; display:block; margin:4px 0;" alt="">');
                            sync();
                        }}
                    }})
                    .catch(function() {{ alert('Erro ao fazer upload da imagem.'); }});
            }}, 