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

# Seções não-musicais — interrompem a coleta ao aparecerem como cabeçalho
CORTE_SECAO = [
    "TEATRO","DANÇA","CIRCO","EXPOSIÇÕES","CINEMA","CRIANÇAS","JOVENS",
    "PESSOAS IDOSAS","ESPORTE","TURISMO SOCIAL","TRABALHO E","MEIO AMBIENTE",
    "DIREITOS HUMANOS","TECNOLOGIAS","SAÚDE","ALIMENTAÇÃO","LITERATURA",
    "CENTRO DE PESQUISA","LOJAS SESC",
]

PREPOSICOES = {
    "a","ao","à","às","de","do","da","dos","das","em","no","na","nos","nas",
    "por","para","pelo","pela","pelos","pelas","com","entre","e","ou",
    "que","se","sua","seu","seus","suas","um","uma","os","as","o",
}

PREFIXOS_DESC = (
    "part.","com ","curadoria","mediação","apresentação de","realização",
    "intervenção","produção","direção","texto ","arranjo","criação ",
    "foto:","#",
)

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
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/em-cartaz-{m}-{ano}.pdf",
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

def extrair_texto(conteudo: bytes) -> str:
    """
    Extrai texto página por página com pdfplumber (sem crop de colunas).
    Usa extract_text() simples que preserva a ordem de leitura natural do PDF.
    """
    partes = []
    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt and txt.strip():
                partes.append(txt)
    return "\n".join(partes)


def ler_pdf_local(caminho: str) -> str:
    with open(caminho, "rb") as f:
        return extrair_texto(f.read())


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3 – PARSER DA SEÇÃO MÚSICA
# ─────────────────────────────────────────────────────────────────────────────

def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()


def _eh_unidade_capital(linha: str) -> bool:
    lu = re.sub(r'\s*\(.*?\)', '', linha.upper()).strip()
    return any(u in lu for u in UNIDADES_CAPITAL)


def _eh_linha_unidade(linha: str) -> bool:
    """
    Linha de unidade Sesc: toda em maiúsculas, sem dígitos, máx 70 chars.
    Ex: 'BELENZINHO', 'AVENIDA PAULISTA', 'CONSOLAÇÃO'
    """
    ls = linha.strip()
    if not (3 <= len(ls) <= 70):
        return False
    base = re.sub(r'\s*\(.*?\)', '', ls).strip()
    if not base or base != base.upper():
        return False
    if re.search(r'\d', base):
        return False
    # Deve ser uma unidade conhecida da capital
    return _eh_unidade_capital(ls)


def _parse_linha_data(linha: str) -> dict | None:
    """
    Detecta linha de metadados do evento.
    Padrão: 'Dia DD/MM. Dia-da-semana, HHh. Classif R$XX R$YY'
    """
    m = re.match(
        r'(?:Dia[s]?\s+|De\s+)?'
        r'(\d{1,2}(?:\s*(?:e|a|,)\s*\d{1,2})?'
        r'(?:\s*/\s*\d{1,2})?/\d{1,2}'
        r'(?:\s*(?:e|a|,)\s*\d{1,2}/\d{1,2})?)'
        r'([^.]*)\.\s*(.*)',
        linha, re.IGNORECASE
    )
    if not m:
        return None

    data_raw = _norm(m.group(1) + m.group(2))
    resto    = _norm(m.group(3))

    hora_m = re.search(r'(\d{1,2}h\d{0,2}(?:\s*às\s*\d{1,2}h\d{0,2})?)', resto, re.I)

    # Captura múltiplos valores de preço: R$50 R$25 R$15
    preco_vals = re.findall(r'R\$\s*\d+', resto)
    if preco_vals:
        preco = " / ".join(preco_vals)
    elif re.search(r'GRÁTIS', resto, re.I):
        preco = "GRÁTIS"
    else:
        preco = ""

    clas_m = re.search(r'\b(A(?:L|\d{1,2})|Livre|Todas as idades)\b', resto, re.I)

    return {
        "Data":          data_raw,
        "Horário":       hora_m.group(1) if hora_m else "",
        "Preço":         preco,
        "Classificação": clas_m.group(1) if clas_m else "",
    }


def _titulo_valido(linha: str) -> bool:
    """Retorna True se a linha pode ser um nome de artista/show."""
    ls = linha.strip()
    if not ls or len(ls) < 3:
        return False
    # Começa com minúscula → continuação de frase
    if ls[0].islower():
        return False
    # Primeira palavra é preposição/artigo
    primeira = ls.split()[0].lower().rstrip('.,;:')
    if primeira in PREPOSICOES:
        return False
    # Prefixo de crédito
    ls_low = ls.lower()
    if any(ls_low.startswith(p) for p in PREFIXOS_DESC):
        return False
    # É linha de data
    if re.match(r'Dia[s]?\s+\d|De\s+\d|\d{1,2}/\d{1,2}', ls, re.I):
        return False
    return True


def _limpa_titulo(t: str) -> str:
    t = re.sub(r'\bFoto:.*', '', t, flags=re.I)
    t = re.sub(r'\s*\|\s*.*$', '', t)
    return _norm(t)[:90]


def parse_musica(texto: str) -> list[dict]:
    # ── 1. Localiza início da seção MÚSICA ───────────────────────────────────
    idx = -1
    for pat in [r'\bMÚSICA\b', r'\bMUSICA\b']:
        m = re.search(pat, texto, re.IGNORECASE)
        if m and (idx == -1 or m.start() < idx):
            idx = m.start()

    if idx == -1:
        print("  ⚠️  Seção 'MÚSICA' não encontrada no PDF.")
        return []

    secao_musica = texto[idx:]

    # ── 2. Corta ao encontrar próxima seção não-musical ──────────────────────
    corte_idx = len(secao_musica)
    for secao in CORTE_SECAO:
        # Busca a seção como linha isolada (cabeçalho)
        c = re.search(
            rf'(?:^|\n)[ \t]*{re.escape(secao)}[ \t]*(?:\n|$)',
            secao_musica, re.IGNORECASE
        )
        if c and c.start() > 100 and c.start() < corte_idx:
            corte_idx = c.start()

    secao_musica = secao_musica[:corte_idx]

    # ── 3. Parse linha a linha ───────────────────────────────────────────────
    eventos       = []
    unidade_atual = None
    titulo_atual  = ""
    tem_descricao = False   # True depois de registrar o título, até encontrar a data

    for linha_raw in secao_musica.split("\n"):
        linha = _norm(linha_raw)
        if not linha:
            tem_descricao = False
            continue

        # Detecta cabeçalho de unidade
        if _eh_linha_unidade(linha):
            unidade_atual = linha
            titulo_atual  = ""
            tem_descricao = False
            continue

        if not unidade_atual:
            continue

        # Detecta linha de data/metadados → fecha o evento
        ev = _parse_linha_data(linha)
        if ev:
            if titulo_atual:
                ev["Artista / Título"] = _limpa_titulo(titulo_atual)
                ev["Unidade"]          = f"Sesc {unidade_atual.title()}"
                eventos.append(ev)
            titulo_atual  = ""
            tem_descricao = False
            continue

        # Linha de título ou descrição
        if not tem_descricao:
            # Primeira linha após unidade ou após evento fechado: deve ser título
            if _titulo_valido(linha):
                titulo_atual  = _limpa_titulo(linha)
                tem_descricao = True
        # Se já tem título, ignora linhas de descrição (não sobrescreve)

    return eventos


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4 – GERAR O EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel(eventos: list[dict], caminho: Path, mes_ref: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Programação Musical"
    nc = len(COLUNAS)

    # Linha 1: título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(row=1, column=1)
    c.value     = f"Programação Musical – Sesc SP Capital | {mes_ref}"
    c.font      = Font(name="Arial", bold=True, size=13, color=COR_BRANCO)
    c.fill      = PatternFill("solid", fgColor=COR_VERMELHO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Linha 2: subtítulo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    c = ws.cell(row=2, column=1)
    c.value     = (f"Extraído em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                   "   |   Fonte: PDF Em Cartaz – sescsp.org.br")
    c.font      = Font(name="Arial", size=9, italic=True, color="555555")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Linha 3: cabeçalho
    for ci, nome in enumerate(COLUNAS, 1):
        c = ws.cell(row=3, column=ci, value=nome)
        c.font      = Font(name="Arial", bold=True, size=10, color=COR_BRANCO)
        c.fill      = PatternFill("solid", fgColor=COR_VERM_ESC)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = Border(bottom=Side(style="thin", color=COR_BRANCO))
    ws.row_dimensions[3].height = 22

    borda = Border(
        bottom=Side(style="thin", color="DDDDDD"),
        right =Side(style="thin", color="DDDDDD"),
    )
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
    print("  Extrator via PDF 'Em Cartaz' (pdfplumber)")
    print("=" * 58)

    agora   = datetime.now()
    ano     = agora.year
    mes     = agora.month
    mes_ref = f"{MESES_PT[mes].capitalize()} {ano}"

    print(f"\n📅 Mês de referência: {mes_ref}")
    print("\n🔍 Procurando PDF 'Em Cartaz'...")
    url_pdf = descobrir_url_pdf(ano, mes)

    texto = ""

    if not url_pdf:
        print("  ⚠️  Buscando link no HTML do site...")
        url_pdf = buscar_pdf_no_html(ano, mes)

    if url_pdf:
        print(f"\n📥 Baixando: {url_pdf.split('uploads/')[-1]}")
        conteudo = baixar_pdf(url_pdf)
        if conteudo:
            print("📄 Extraindo texto com pdfplumber...")
            texto = extrair_texto(conteudo)
            if texto.strip():
                print(f"   Texto extraído: {len(texto):,} chars")
        else:
            print("❌ Falha ao baixar o PDF.")

    if not texto.strip():
        print("\n⚠️  PDF não obtido automaticamente.")
        print("   Baixe o PDF 'Em Cartaz' em sescsp.org.br/programacao")
        caminho_manual = input("   Informe o caminho do PDF: ").strip().strip('"')
        if not caminho_manual:
            sys.exit(1)
        if not Path(caminho_manual).exists():
            print(f"❌ Arquivo não encontrado: {caminho_manual}")
            sys.exit(1)
        print("📄 Lendo PDF local...")
        texto = ler_pdf_local(caminho_manual)

    if not texto.strip():
        print("❌ Não foi possível extrair texto do PDF.")
        sys.exit(1)

    print("🎵 Parseando programação musical da capital...")
    eventos = parse_musica(texto)

    if not eventos:
        print("\n⚠️  Nenhum evento encontrado.")
        print("   Trecho do texto extraído para diagnóstico:")
        idx = texto.find("MÚSICA")
        if idx == -1:
            idx = 0
        print(texto[idx:idx+800])
        sys.exit(1)

    eventos.sort(key=lambda x: (x.get("Unidade",""), x.get("Data","")))

    nome    = f"sesc_musica_capital_{ano}_{mes:02d}.xlsx"
    destino = Path(__file__).parent / nome
    print(f"\n📊 Gerando planilha: {nome}")
    gerar_excel(eventos, destino, mes_ref)

    print(f"\n✅ Concluído!")
    print(f"   Arquivo: {destino}")
    print(f"   Eventos encontrados: {len(eventos)}")


if __name__ == "__main__":
    main()
