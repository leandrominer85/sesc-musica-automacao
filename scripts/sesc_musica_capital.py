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

CORTE_SECAO = [
    "TEATRO","DANÇA","CIRCO","EXPOSIÇÕES","CINEMA","CRIANÇAS","JOVENS",
    "PESSOAS IDOSAS","ESPORTE","TURISMO SOCIAL","TRABALHO E","MEIO AMBIENTE",
    "DIREITOS HUMANOS","TECNOLOGIAS","SAÚDE","ALIMENTAÇÃO","LITERATURA",
    "CENTRO DE PESQUISA","LOJAS SESC","INGRESSOS",
]

IGNORAR_LINHAS = {
    "SHOWS E ESPETÁCULOS","INGRESSOS","EM CARTAZ","PROGRAMAÇÃO",
    "EVENTOS","MÚSICA","MUSICA",
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


# Padrão para detectar qualquer PDF "Em Cartaz" no HTML
_RE_PDF = re.compile(
    r'https://www\.sescsp\.org\.br/wp-content/uploads/[^\s"\'<>]*'
    r'[Ee][Mm][- _][Cc][Aa][Rr][Tt][Aa][Zz][^\s"\'<>]*\.pdf',
    re.IGNORECASE
)

# Abreviações dos meses usadas nos nomes de arquivo
MESES_ABREV = {
    1:"JAN", 2:"FEV", 3:"MAR", 4:"ABR", 5:"MAI", 6:"JUN",
    7:"JUL", 8:"AGO", 9:"SET", 10:"OUT", 11:"NOV", 12:"DEZ",
}


def buscar_pdf_no_html(ano: int, mes: int) -> str | None:
    """
    Estratégia principal: raspa o HTML do site para encontrar o link real do PDF.
    O nome do arquivo muda a cada edição — essa é a única forma confiável.
    """
    m_lower = MESES_PT[mes]          # "julho"
    m_abrev = MESES_ABREV[mes]       # "JUL"

    paginas = [
        "https://www.sescsp.org.br/programacao/",
        "https://www.sescsp.org.br/",
        f"https://www.sescsp.org.br/wp-sitemap.xml",
    ]

    for url_busca in paginas:
        try:
            r = requests.get(url_busca, headers=HEADERS, timeout=20)
            if r.status_code != 200:
                continue

            pdfs = _RE_PDF.findall(r.text)
            print(f"  📄 {len(pdfs)} PDF(s) encontrado(s) em {url_busca.split('sescsp.org.br/')[-1]}")

            for u in pdfs:
                u_low = u.lower()
                # Confere se o PDF é do mês/ano correto
                if (str(ano) in u and
                        (m_lower in u_low or m_abrev.lower() in u_low)):
                    print(f"  ✅ Match: {u.split('uploads/')[-1]}")
                    return u

        except Exception as e:
            print(f"  ⚠️  Erro ao acessar {url_busca}: {e}")

    return None


def descobrir_url_pdf(ano: int, mes: int) -> str | None:
    """
    Primeiro tenta raspar o HTML (mais confiável).
    Fallback: testa URLs candidatas com padrões históricos.
    """
    # 1. Raspagem do HTML — método primário
    print("  🌐 Buscando link no HTML do site...")
    url = buscar_pdf_no_html(ano, mes)
    if url:
        return url

    # 2. Fallback: candidatos baseados em padrões históricos
    print("  🔎 Testando URLs candidatas...")
    m      = MESES_PT[mes]
    m_abr  = MESES_ABREV[mes]
    # Mês de publicação = mês anterior
    mes_pub = mes - 1 if mes > 1 else 12
    ano_pub = ano if mes > 1 else ano - 1

    candidatos = [
        # Padrões recentes (2026)
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/Em-Cartaz_{ano}_{m}_compressed.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V12_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V10_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V11_EM-CARTAZ_{m_abr}_{ano}_bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V9_EM-CARTAZ_{m_abr}_{ano}.pdf",
        # Com ano na pasta
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz_{ano}_{m}_compressed.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz_{ano}_{m}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz-{m}-{ano}.pdf",
        # Pasta 2025 (padrão visto no PDF de julho 2026)
        f"https://www.sescsp.org.br/wp-content/uploads/2025/01/Em-Cartaz_{ano}_{m}_compressed.pdf",
    ]

    for url in candidatos:
        nome = url.split('uploads/')[-1]
        print(f"    {nome}")
        if _pdf_existe(url):
            print(f"  ✅ Encontrado: {nome}")
            return url

    return None


def baixar_pdf(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=45)
        return r.content if r.status_code == 200 else None
    except Exception as e:
        print(f"  ❌ {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 2 – EXTRAIR TEXTO DO PDF (sem separação de colunas)
# ─────────────────────────────────────────────────────────────────────────────
# Estratégia: extrair palavras com coordenadas XY e reconstruir por linha,
# mas SEM separar colunas — deixamos o parser lidar com a mistura.
# A separação de colunas criava blocos desconexos que quebravam o contexto.

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


def extrair_texto(conteudo: bytes) -> str:
    """
    Extrai texto página por página com pdfplumber.
    Reconstrói cada linha agrupando palavras pelo mesmo Y,
    ordenadas por X — sem dividir em colunas.
    O parser cuida de ignorar conteúdo de outras seções.
    """
    partes = []
    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=3, y_tolerance=3,
                keep_blank_chars=False, use_text_flow=False,
            )
            if not words:
                continue

            # Agrupa por Y (tolerância de 3 pt) e ordena por X
            linhas: dict[int, list] = {}
            for w in words:
                y = round(float(w["top"]) / 3) * 3
                linhas.setdefault(y, []).append(
                    (float(w["x0"]), w["text"])
                )

            for y in sorted(linhas):
                linha = " ".join(t for _, t in sorted(linhas[y]))
                if linha.strip():
                    partes.append(linha)

            partes.append("")  # linha em branco entre páginas

    return _corrigir_espelhado("\n".join(partes))


def ler_pdf_local(caminho: str) -> str:
    with open(caminho, "rb") as f:
        return extrair_texto(f.read())


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3 – PARSER DA SEÇÃO MÚSICA
# ─────────────────────────────────────────────────────────────────────────────

def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()


def _eh_unidade_capital(ls: str) -> bool:
    lu = re.sub(r'\s*\(.*?\)', '', ls.upper()).strip()
    # Match exato ou com qualificador entre parênteses
    for u in UNIDADES_CAPITAL:
        if lu == u:
            return True
        if lu.startswith(u) and len(lu) > len(u) and lu[len(u)] in ' (':
            return True
    return False


def _eh_linha_unidade(ls: str) -> bool:
    if not (3 <= len(ls) <= 55):
        return False
    base = re.sub(r'\s*\(.*?\)', '', ls).strip()
    if not base or base != base.upper() or re.search(r'\d', base):
        return False
    return _eh_unidade_capital(ls)


def _eh_secao_corte(ls: str) -> bool:
    lu = ls.upper().strip()
    return any(lu == s or lu.startswith(s + " ") for s in CORTE_SECAO)


def _eh_continuacao_metadado(ls: str) -> bool:
    """Linha de preço/hora/classif solta — não é novo título."""
    if not ls:
        return False
    if re.match(r'^R\$\s*\d', ls):
        return True
    if re.match(r'^\d{1,2}h', ls):
        return True
    if re.match(DIAS_SEMANA, ls, re.I):
        return True
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
    if _eh_continuacao_metadado(ls):
        return False
    return True


def _limpa_titulo(t: str) -> str:
    t = re.sub(r'\bFoto:.*', '', t, flags=re.I)
    t = re.sub(r'\s*\|.*$', '', t)
    # Remove duplicação "X X" → "X"
    words = t.split()
    meio = len(words) // 2
    if meio > 2 and words[:meio] == words[meio:]:
        t = " ".join(words[:meio])
    return _norm(t)[:90]


def parse_musica(texto: str) -> list[dict]:
    # Localiza início da seção MÚSICA
    idx = -1
    for pat in [r'\bMÚSICA\b', r'\bMUSICA\b']:
        for m in re.finditer(pat, texto, re.IGNORECASE):
            if idx == -1 or m.start() < idx:
                idx = m.start()

    if idx == -1:
        print("  ⚠️  'MÚSICA' não encontrada no texto.")
        print("  Amostra do texto extraído:")
        print(texto[:600])
        return []

    print(f"  ✅ Seção MÚSICA encontrada na posição {idx}")

    # Estado
    eventos       = []
    unidade_atual = None
    ev_aberto     = None
    titulo_atual  = ""
    tem_desc      = False
    em_musica     = True
    unidades_vistas = []

    def _fechar():
        nonlocal ev_aberto, titulo_atual, tem_desc
        if ev_aberto is not None and titulo_atual:
            ev_aberto["Artista / Título"] = _limpa_titulo(titulo_atual)
            ev_aberto["Unidade"]          = f"Sesc {unidade_atual.title()}"
            eventos.append(ev_aberto)
        ev_aberto    = None
        titulo_atual = ""
        tem_desc     = False

    for linha_raw in texto[idx:].split("\n"):
        ls = _norm(linha_raw)
        if not ls:
            tem_desc = False
            continue

        # Seção não-musical → desativa coleta
        if _eh_secao_corte(ls):
            em_musica = False
            continue

        # Reabre coleta se aparecer MÚSICA novamente
        if re.match(r'^(MÚSICA|MUSICA)$', ls, re.I):
            em_musica = True
            continue

        if not em_musica:
            continue

        # Linha ignorada (cabeçalhos internos)
        if ls.upper() in IGNORAR_LINHAS:
            continue

        # Linha de UNIDADE
        if _eh_linha_unidade(ls):
            _fechar()
            unidade_atual = ls
            unidades_vistas.append(ls)
            continue

        if not unidade_atual:
            continue

        # Continuação de metadado (preço/hora soltos)
        if _eh_continuacao_metadado(ls):
            if ev_aberto is not None:
                if not ev_aberto.get("Horário"):
                    h = _extrair_hora(ls)
                    if h: ev_aberto["Horário"] = h
                if not ev_aberto.get("Preço"):
                    p = _extrair_preco(ls)
                    if p: ev_aberto["Preço"] = p
                if not ev_aberto.get("Classificação"):
                    c = _extrair_classif(ls)
                    if c: ev_aberto["Classificação"] = c
            continue

        # Linha de DATA
        ev_data = _parse_linha_data(ls)
        if ev_data:
            if ev_aberto is not None:
                _fechar()
            ev_aberto = dict(ev_data)
            tem_desc  = False
            continue

        # Título ou descrição
        if not tem_desc:
            if _titulo_valido(ls):
                if ev_aberto is not None:
                    _fechar()
                titulo_atual = _limpa_titulo(ls)
                tem_desc     = True

    _fechar()

    print(f"  Unidades detectadas: {unidades_vistas}")

    # Remove duplicatas
    vistos, unicos = set(), []
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

    # Diagnóstico: mostra contexto ao redor de MÚSICA no texto
    idx_musica = texto.find("MÚSICA")
    if idx_musica == -1:
        idx_musica = texto.find("MUSICA")
    if idx_musica >= 0:
        print(f"\n📋 Contexto da seção MÚSICA (chars {idx_musica}-{idx_musica+300}):")
        print(texto[idx_musica:idx_musica+300])
        print("...")

    print("\n🎵 Parseando programação musical...")
    eventos = parse_musica(texto)

    if not eventos:
        print("\n⚠️  Nenhum evento encontrado.")
        print("Amostra completa do texto extraído (primeiros 2000 chars):")
        print(texto[:2000])
        sys.exit(1)

    eventos.sort(key=lambda x: (x.get("Unidade",""), x.get("Data","")))
    nome    = f"sesc_musica_capital_{ano}_{mes:02d}.xlsx"
    destino = Path(__file__).parent / nome
    print(f"\n📊 Gerando: {nome}")
    gerar_excel(eventos, destino, mes_ref)
    print(f"\n✅ Concluído! {len(eventos)} eventos → {destino}")


if __name__ == "__main__":
    main()
