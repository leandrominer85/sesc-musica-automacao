"""
sesc_musica_capital.py
======================
Extrai programação de MÚSICA das unidades da CAPITAL do Sesc SP
e salva como planilha Excel (.xlsx).

Dependências:
    pip install requests pdfplumber openpyxl
"""

import re
import sys
import io
import os
from datetime import datetime
from pathlib import Path

def _req(pkg, install=None):
    try:
        __import__(pkg)
    except ImportError:
        print(f"❌ '{pkg}' não encontrado.\n   pip install {install or pkg}")
        sys.exit(1)

_req("requests")
_req("pdfplumber")
_req("openpyxl")

import requests
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

MESES_PT = {
    1:"janeiro", 2:"fevereiro", 3:"março",     4:"abril",
    5:"maio",    6:"junho",     7:"julho",      8:"agosto",
    9:"setembro",10:"outubro",  11:"novembro",  12:"dezembro",
}

UNIDADES_CAPITAL = [
    "14 BIS","24 DE MAIO","AVENIDA PAULISTA","BELENZINHO","BOM RETIRO",
    "CAMPO LIMPO","CASA VERDE","CONSOLAÇÃO","INTERLAGOS","IPIRANGA",
    "ITAQUERA","PINHEIROS","POMPEIA","SANTANA","SANTO AMARO","SÃO BENTO",
    "SÃO JOÃO CLÍMACO","VILA MARIANA","CARMO","PARQUE DOM PEDRO",
    "BRÁS","BRAS",
]

# Seções não-musicais — ao aparecerem como cabeçalho, param a coleta
CORTE_SECAO = [
    "TEATRO","DANÇA","CIRCO","EXPOSIÇÕES","CINEMA","CRIANÇAS","JOVENS",
    "PESSOAS IDOSAS","ESPORTE","TURISMO SOCIAL","TRABALHO E","MEIO AMBIENTE",
    "DIREITOS HUMANOS","TECNOLOGIAS","SAÚDE","ALIMENTAÇÃO","LITERATURA",
    "CENTRO DE PESQUISA","LOJAS SESC","SHOWS E ESPETÁCULOS","INGRESSOS",
]

# Títulos de subsessão que devem ser ignorados (não são artistas)
IGNORAR_LINHAS = {
    "SHOWS E ESPETÁCULOS","INGRESSOS","EM CARTAZ","JUNHO 2026","JULHO 2026",
    "AGOSTO 2026","PROGRAMAÇÃO","EVENTOS","MÚSICA","MUSICA",
}

PREPOSICOES = {
    "a","ao","à","às","de","do","da","dos","das","em","no","na","nos","nas",
    "por","para","pelo","pela","pelos","pelas","com","entre","e","ou",
    "que","se","sua","seu","seus","suas","um","uma","os","as","o",
}

PREFIXOS_DESC = (
    "part.","com ","curadoria","mediação","apresentação de","realização",
    "intervenção","produção","direção","texto ","arranjo","criação ",
    "foto:","#","consulte","local:","informações",
)

DIAS_SEMANA = r'(?:Segunda|Terça|Quarta|Quinta|Sexta|Sábado|Domingo|' \
              r'Seg\.|Ter\.|Qua\.|Qui\.|Sex\.|Sáb\.|Dom\.)'

COR_VERMELHO = "C00000"
COR_VERM_ESC = "8B0000"
COR_BRANCO   = "FFFFFF"
COR_CINZA    = "F5F5F5"

COLUNAS  = ["Artista / Título", "Unidade", "Data", "Horário", "Preço", "Classificação"]
LARGURAS = [42, 24, 26, 12, 22, 16]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.sescsp.org.br/",
}

# ─────────────────────────────────────────────────────────────────────────────
# PASSO 1 – LOCALIZAR E BAIXAR O PDF
# ─────────────────────────────────────────────────────────────────────────────

def _pdf_existe(url):
    try:
        r = requests.head(url, headers=HEADERS, timeout=10, allow_redirects=True)
        ct = r.headers.get("Content-Type", "").lower()
        return r.status_code == 200 and "pdf" in ct
    except Exception:
        return False


def descobrir_url_pdf(ano, mes):
    m       = MESES_PT[mes]
    mes_pub = mes - 1 if mes > 1 else 12
    ano_pub = ano if mes > 1 else ano - 1
    candidatos = [
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V12_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V10_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz_{ano}_{m}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz-{m}-{ano}.pdf",
    ]
    for url in candidatos:
        print(f"  🔎 {url.split('uploads/')[-1]}")
        if _pdf_existe(url):
            return url
    return None


def buscar_pdf_no_html(ano, mes):
    m = MESES_PT[mes]
    for base in ["https://www.sescsp.org.br/programacao/", "https://www.sescsp.org.br/"]:
        try:
            r = requests.get(base, headers=HEADERS, timeout=15)
            pdfs = re.findall(
                r'https://www\.sescsp\.org\.br/wp-content/uploads/[^\s"\'<>]*'
                r'[Ee]m[- _][Cc]artaz[^\s"\'<>]*\.pdf', r.text
            )
            for u in pdfs:
                if m in u.lower() or str(ano) in u:
                    return u
        except Exception:
            pass
    return None


def baixar_pdf(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=45)
        return r.content if r.status_code == 200 else None
    except Exception as e:
        print(f"  ❌ {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 2 – EXTRAIR TEXTO DO PDF
# ─────────────────────────────────────────────────────────────────────────────

def _corrigir_espelhado(texto: str) -> str:
    """Corrige palavras extraídas invertidas por fontes não-padrão."""
    linhas = []
    for linha in texto.split("\n"):
        palavras = []
        for p in linha.split():
            rev = p[::-1]
            if (len(p) >= 3 and p[0].islower() and rev[0].isupper()
                    and rev[-1].islower() and p.isalpha()):
                palavras.append(rev)
            else:
                palavras.append(p)
        linhas.append(" ".join(palavras))
    return "\n".join(linhas)


def _reconstruir_por_coords(words: list) -> str:
    """Reconstrói linhas a partir de palavras com coordenadas XY."""
    if not words:
        return ""
    linhas: dict[int, list] = {}
    for w in words:
        texto_w = w.get("text", "").strip()
        if not texto_w:
            continue
        y = round(float(w.get("top", 0)) / 3) * 3
        linhas.setdefault(y, []).append((float(w.get("x0", 0)), texto_w))
    resultado = []
    for y in sorted(linhas):
        resultado.append(" ".join(p[1] for p in sorted(linhas[y])))
    return "\n".join(resultado)


def extrair_texto(conteudo: bytes) -> str:
    """
    Extrai texto separando as 2 colunas do PDF por coordenada X,
    garantindo que conteúdo de seções diferentes não se misture.
    """
    partes = []
    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for num_pag, page in enumerate(pdf.pages):
            largura = float(page.width)
            words = page.extract_words(x_tolerance=3, y_tolerance=3,
                                       keep_blank_chars=False, use_text_flow=False)
            if not words:
                continue

            # Detecta ponto de corte entre colunas (maior gap horizontal)
            xs = sorted(set(round(float(w["x0"])) for w in words))
            corte = largura / 2
            zona = [x for x in xs if largura * 0.25 <= x <= largura * 0.75]
            if len(zona) > 1:
                gaps = [(zona[i+1]-zona[i], zona[i]) for i in range(len(zona)-1)]
                maior = max(gaps, key=lambda g: g[0])
                corte = maior[1] + maior[0] / 2

            esq = _reconstruir_por_coords([w for w in words if float(w["x0"]) < corte])
            dir_ = _reconstruir_por_coords([w for w in words if float(w["x0"]) >= corte])

            if esq.strip():
                partes.append(f"[P{num_pag+1}E]\n{esq}")
            if dir_.strip():
                partes.append(f"[P{num_pag+1}D]\n{dir_}")

    return _corrigir_espelhado("\n".join(partes))


def ler_pdf_local(caminho: str) -> str:
    with open(caminho, "rb") as f:
        return extrair_texto(f.read())


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3 – PARSER DA SEÇÃO MÚSICA
# ─────────────────────────────────────────────────────────────────────────────

def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()

def _sem_marcador(linha: str) -> str:
    """Remove marcadores de página/coluna [P1E], [P2D] etc."""
    return re.sub(r'^\[P\d+[ED]\]\s*', '', linha).strip()

def _eh_unidade_capital(ls: str) -> bool:
    lu = re.sub(r'\s*\(.*?\)', '', ls.upper()).strip()
    # Garante que é EXATAMENTE uma unidade, não "CAMPO LIMPO torcer" etc.
    return any(u == lu or lu.startswith(u) and (len(lu) == len(u) or lu[len(u)] in ' (')
               for u in UNIDADES_CAPITAL)

def _eh_linha_unidade(ls: str) -> bool:
    if not (3 <= len(ls) <= 50):
        return False
    base = re.sub(r'\s*\(.*?\)', '', ls).strip()
    if not base or base != base.upper() or re.search(r'\d', base):
        return False
    return _eh_unidade_capital(ls)

def _eh_secao_corte(ls: str) -> bool:
    lu = ls.upper().strip()
    return any(lu == s or lu.startswith(s + " ") for s in CORTE_SECAO)

def _eh_continuacao_metadado(ls: str) -> bool:
    """
    Detecta linha que é continuação de metadado do evento anterior
    (preço, horário, classificação soltos), não um novo título.
    """
    if not ls:
        return False
    # Começa com R$
    if re.match(r'^R\$\s*\d', ls):
        return True
    # Começa com horário (HHh)
    if re.match(r'^\d{1,2}h', ls):
        return True
    # Começa com dia da semana
    if re.match(DIAS_SEMANA, ls, re.I):
        return True
    # É só classificação e/ou GRÁTIS
    if re.match(r'^(AL|A\d{1,2}|Livre|Todas as|GRÁTIS)', ls, re.I):
        return True
    return False

def _extrair_preco(texto: str) -> str:
    vals = re.findall(r'R\$\s*\d+', texto)
    if vals:
        return " / ".join(vals)
    if re.search(r'GRÁTIS', texto, re.I):
        return "GRÁTIS"
    return ""

def _extrair_hora(texto: str) -> str:
    m = re.search(r'(\d{1,2}h\d{0,2}(?:\s*às\s*\d{1,2}h\d{0,2})?)', texto, re.I)
    return m.group(1) if m else ""

def _extrair_classif(texto: str) -> str:
    m = re.search(r'\b(A(?:L|\d{1,2})|Livre|Todas as idades)\b', texto, re.I)
    return m.group(1) if m else ""

def _parse_linha_data(ls: str) -> dict | None:
    """Detecta linha que contém data de evento."""
    m = re.match(
        r'(?:Dia[s]?\s+|De\s+)?'
        r'(\d{1,2}(?:\s*(?:e|a|,)\s*\d{1,2})?'
        r'(?:\s*/\s*\d{1,2})?/\d{1,2}'
        r'(?:\s*(?:e|a|,)\s*\d{1,2}/\d{1,2})?)'
        r'([^.]*)\.\s*(.*)',
        ls, re.IGNORECASE
    )
    if not m:
        return None
    data_raw = _norm(m.group(1) + m.group(2))
    resto    = _norm(m.group(3))
    return {
        "Data":          data_raw,
        "Horário":       _extrair_hora(resto),
        "Preço":         _extrair_preco(resto),
        "Classificação": _extrair_classif(resto),
        "_resto":        resto,
    }

def _titulo_valido(ls: str) -> bool:
    if not ls or len(ls) < 3:
        return False
    if ls.upper() in IGNORAR_LINHAS:
        return False
    if ls[0].islower():
        return False
    primeira = ls.split()[0].lower().rstrip('.,;:()')
    if primeira in PREPOSICOES:
        return False
    ls_low = ls.lower()
    if any(ls_low.startswith(p) for p in PREFIXOS_DESC):
        return False
    if re.match(r'Dia[s]?\s+\d|De\s+\d|\d{1,2}/\d{1,2}', ls, re.I):
        return False
    if re.match(r'^\[P\d+', ls):
        return False
    if _eh_continuacao_metadado(ls):
        return False
    return True

def _limpa_titulo(t: str) -> str:
    t = re.sub(r'\bFoto:.*', '', t, flags=re.I)
    t = re.sub(r'\s*\|.*$', '', t)
    t = re.sub(r'^\[P\d+[ED]\]\s*', '', t)
    # Remove texto duplicado: "Brasa: Uma Dança Brasa: Uma Dança" → "Brasa: Uma Dança"
    words = t.split()
    meio = len(words) // 2
    if meio > 2 and words[:meio] == words[meio:]:
        t = " ".join(words[:meio])
    return _norm(t)[:90]


def parse_musica(texto: str) -> list[dict]:
    # Localiza seção MÚSICA
    idx = -1
    for pat in [r'\bMÚSICA\b', r'\bMUSICA\b']:
        for m in re.finditer(pat, texto, re.IGNORECASE):
            if idx == -1 or m.start() < idx:
                idx = m.start()
    if idx == -1:
        print("  ⚠️  Seção 'MÚSICA' não encontrada.")
        print(texto[:500])
        return []

    # Estado do parser
    eventos       = []
    unidade_atual = None
    ev_aberto     = None   # evento sendo construído (pode ter data sem preço ainda)
    titulo_atual  = ""
    tem_desc      = False
    em_musica     = True

    def _fechar_evento():
        nonlocal ev_aberto, titulo_atual, tem_desc
        if ev_aberto and titulo_atual:
            ev_aberto["Artista / Título"] = _limpa_titulo(titulo_atual)
            ev_aberto["Unidade"]          = f"Sesc {unidade_atual.title()}"
            eventos.append(ev_aberto)
        ev_aberto    = None
        titulo_atual = ""
        tem_desc     = False

    for linha_raw in texto[idx:].split("\n"):
        linha  = _norm(linha_raw)
        ls     = _sem_marcador(linha)

        if not ls:
            tem_desc = False
            continue

        # Marcador de coluna → reseta contexto de seção mas não de unidade
        if re.match(r'^\[P\d+[ED]\]$', linha):
            continue

        # Seção não-musical → para de coletar desta coluna
        if _eh_secao_corte(ls):
            em_musica = False
            continue

        # Volta a coletar se encontrar cabeçalho MÚSICA novamente (outra coluna)
        if re.match(r'^(MÚSICA|MUSICA)$', ls, re.I):
            em_musica = True
            continue

        if not em_musica:
            continue

        # ── Linha de UNIDADE ────────────────────────────────────────────────
        if _eh_linha_unidade(ls):
            _fechar_evento()
            unidade_atual = ls
            continue

        if not unidade_atual:
            continue

        # ── Linha ignorada (cabeçalhos internos) ────────────────────────────
        if ls.upper() in IGNORAR_LINHAS:
            continue

        # ── Continuação de metadado (preço/hora soltos após data) ───────────
        if _eh_continuacao_metadado(ls):
            if ev_aberto is not None:
                # Enriquece o evento aberto com dados faltantes
                if not ev_aberto.get("Horário"):
                    h = _extrair_hora(ls)
                    if h:
                        ev_aberto["Horário"] = h
                if not ev_aberto.get("Preço"):
                    p = _extrair_preco(ls)
                    if p:
                        ev_aberto["Preço"] = p
                if not ev_aberto.get("Classificação"):
                    c = _extrair_classif(ls)
                    if c:
                        ev_aberto["Classificação"] = c
            # Se não há evento aberto, ignora essa linha
            continue

        # ── Linha de DATA ───────────────────────────────────────────────────
        ev_data = _parse_linha_data(ls)
        if ev_data:
            # Fecha evento anterior se havia um aberto
            if ev_aberto is not None:
                _fechar_evento()
            # Abre novo evento com os dados da linha de data
            ev_aberto = {
                "Data":          ev_data["Data"],
                "Horário":       ev_data["Horário"],
                "Preço":         ev_data["Preço"],
                "Classificação": ev_data["Classificação"],
            }
            # Título já deve ter sido capturado antes
            if not titulo_atual:
                titulo_atual = ""
            tem_desc = False
            continue

        # ── Título ou descrição ─────────────────────────────────────────────
        if not tem_desc:
            if _titulo_valido(ls):
                # Novo título: fecha evento anterior se aberto
                if ev_aberto is not None:
                    _fechar_evento()
                titulo_atual = _limpa_titulo(ls)
                tem_desc     = True
        # Linhas subsequentes são descrição — não sobrescreve titulo_atual

    # Fecha último evento pendente
    _fechar_evento()

    # Remove duplicatas exatas (mesmo artista + unidade + data)
    vistos = set()
    unicos = []
    for ev in eventos:
        chave = (ev.get("Artista / Título",""), ev.get("Unidade",""), ev.get("Data",""))
        if chave not in vistos:
            vistos.add(chave)
            unicos.append(ev)

    return unicos


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4 – GERAR O EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel(eventos: list[dict], caminho: Path, mes_ref: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Programação Musical"
    nc = len(COLUNAS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(row=1, column=1)
    c.value     = f"Programação Musical – Sesc SP Capital | {mes_ref}"
    c.font      = Font(name="Arial", bold=True, size=13, color=COR_BRANCO)
    c.fill      = PatternFill("solid", fgColor=COR_VERMELHO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    c = ws.cell(row=2, column=1)
    c.value     = (f"Extraído em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                   "   |   Fonte: PDF Em Cartaz – sescsp.org.br")
    c.font      = Font(name="Arial", size=9, italic=True, color="555555")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    for ci, nome in enumerate(COLUNAS, 1):
        c = ws.cell(row=3, column=ci, value=nome)
        c.font      = Font(name="Arial", bold=True, size=10, color=COR_BRANCO)
        c.fill      = PatternFill("solid", fgColor=COR_VERM_ESC)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = Border(bottom=Side(style="thin", color=COR_BRANCO))
    ws.row_dimensions[3].height = 22

    borda = Border(bottom=Side(style="thin", color="DDDDDD"),
                   right =Side(style="thin", color="DDDDDD"))
    for i, ev in enumerate(eventos):
        row  = i + 4
        fill = PatternFill("solid", fgColor=COR_CINZA if i % 2 == 0 else COR_BRANCO)
        for ci, col in enumerate(COLUNAS, 1):
            c = ws.cell(row=row, column=ci, value=ev.get(col, ""))
            c.font      = Font(name="Arial", size=10)
            c.fill      = fill
            c.alignment = Alignment(vertical="center", wrap_text=(ci == 1))
            c.border    = borda
        ws.row_dimensions[row].height = 18

    lr = len(eventos) + 4
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=nc)
    c = ws.cell(row=lr, column=1)
    c.value     = f"Total de eventos musicais: {len(eventos)}"
    c.font      = Font(name="Arial", bold=True, size=10, color=COR_BRANCO)
    c.fill      = PatternFill("solid", fgColor=COR_VERMELHO)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[lr].height = 20

    for ci, larg in enumerate(LARGURAS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = larg
    ws.auto_filter.ref = f"A3:{get_column_letter(nc)}3"
    ws.freeze_panes   = "A4"
    wb.save(caminho)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 58)
    print("  Sesc SP – Programação Musical da Capital")
    print("  Extrator via PDF 'Em Cartaz'")
    print("=" * 58)

    agora   = datetime.now()
    ano, mes = agora.year, agora.month
    mes_ref  = f"{MESES_PT[mes].capitalize()} {ano}"

    print(f"\n📅 Mês de referência: {mes_ref}")
    print("\n🔍 Procurando PDF 'Em Cartaz'...")
    url_pdf = descobrir_url_pdf(ano, mes)
    texto   = ""

    if not url_pdf:
        print("  ⚠️  Buscando no HTML do site...")
        url_pdf = buscar_pdf_no_html(ano, mes)

    if url_pdf:
        print(f"\n📥 Baixando: {url_pdf.split('uploads/')[-1]}")
        conteudo = baixar_pdf(url_pdf)
        if conteudo:
            print("📄 Extraindo texto...")
            texto = extrair_texto(conteudo)
            print(f"   {len(texto):,} chars extraídos")

    if not texto.strip():
        caminho_env = os.environ.get("PDF_LOCAL_PATH", "").strip().strip('"')
        if caminho_env and Path(caminho_env).exists():
            print(f"📄 PDF local: {caminho_env}")
            texto = ler_pdf_local(caminho_env)

    if not texto.strip():
        if sys.stdin.isatty():
            print("\n⚠️  PDF não obtido automaticamente.")
            caminho = input("   Caminho do PDF: ").strip().strip('"')
            if caminho and Path(caminho).exists():
                texto = ler_pdf_local(caminho)
            else:
                print("❌ Caminho inválido."); sys.exit(1)
        else:
            print(f"\n❌ PDF não encontrado (CI). Mês: {mes_ref}")
            print("   O PDF pode ainda não ter sido publicado pelo Sesc.")
            sys.exit(1)

    if not texto.strip():
        print("❌ Não foi possível extrair texto."); sys.exit(1)

    print("🎵 Parseando programação musical...")
    eventos = parse_musica(texto)

    if not eventos:
        print("\n⚠️  Nenhum evento encontrado.")
        idx = max(texto.find("MÚSICA"), texto.find("MUSICA"), 0)
        print(texto[idx:idx+800])
        sys.exit(1)

    eventos.sort(key=lambda x: (x.get("Unidade",""), x.get("Data","")))
    nome    = f"sesc_musica_capital_{ano}_{mes:02d}.xlsx"
    destino = Path(__file__).parent / nome
    print(f"\n📊 Gerando: {nome}")
    gerar_excel(eventos, destino, mes_ref)
    print(f"\n✅ Concluído! {len(eventos)} eventos → {destino}")


if __name__ == "__main__":
    main()
