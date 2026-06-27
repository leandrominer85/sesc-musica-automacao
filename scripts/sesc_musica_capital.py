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
    "foto:","#","consulte",
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
# PASSO 2 – EXTRAIR TEXTO DO PDF COM SEPARAÇÃO DE COLUNAS
# ─────────────────────────────────────────────────────────────────────────────

def _corrigir_texto_espelhado(texto: str) -> str:
    """
    O PDF do Sesc usa fontes com glifos mapeados de forma não-padrão,
    resultando em palavras extraídas de trás para frente.
    Esta função detecta e corrige palavras espelhadas.
    """
    linhas_corrigidas = []
    for linha in texto.split("\n"):
        palavras = linha.split()
        palavras_corrigidas = []
        for palavra in palavras:
            # Detecta palavra espelhada: contém só letras mas está invertida
            # Heurística: se reverter a palavra produz algo mais legível
            # (começa com maiúscula válida após reversão)
            rev = palavra[::-1]
            # Verifica se a palavra invertida começa com maiúscula e termina
            # com minúscula (padrão de nome próprio normal)
            if (len(palavra) >= 3
                    and palavra[0].islower()
                    and rev[0].isupper()
                    and rev[-1].islower()
                    and palavra.isalpha()):
                palavras_corrigidas.append(rev)
            else:
                palavras_corrigidas.append(palavra)
        linhas_corrigidas.append(" ".join(palavras_corrigidas))
    return "\n".join(linhas_corrigidas)


def _reconstruir_linhas_por_coordenadas(words: list) -> str:
    """
    Reconstrói o texto a partir de palavras com coordenadas XY.
    Agrupa palavras na mesma linha (mesmo Y) e ordena por X.
    Isso garante a ordem correta mesmo em PDFs com layout complexo.
    """
    if not words:
        return ""

    # Agrupa por linha (tolerância vertical de 3 pontos)
    linhas = {}
    for w in words:
        texto_w = w.get("text", "").strip()
        if not texto_w:
            continue
        y = round(float(w.get("top", 0)) / 3) * 3  # arredonda para agrupar
        if y not in linhas:
            linhas[y] = []
        linhas[y].append((float(w.get("x0", 0)), texto_w))

    # Ordena linhas por Y crescente, palavras por X crescente
    resultado = []
    for y in sorted(linhas.keys()):
        palavras_linha = sorted(linhas[y], key=lambda p: p[0])
        resultado.append(" ".join(p[1] for p in palavras_linha))

    return "\n".join(resultado)


def extrair_texto_por_coluna(conteudo: bytes) -> str:
    """
    Extrai texto separando as duas colunas do PDF pela posição X central.
    Processa coluna esquerda e direita separadamente para cada página,
    garantindo que conteúdo de seções diferentes não se misture.
    """
    todas_partes = []

    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for num_pag, page in enumerate(pdf.pages):
            largura  = float(page.width)
            altura   = float(page.height)

            # Extrai todas as palavras com coordenadas
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            )

            if not words:
                continue

            # Determina o ponto de divisão das colunas
            # Analisa distribuição de X para encontrar o "vão" central
            xs = sorted(set(round(float(w["x0"])) for w in words))
            meio_pagina = largura / 2

            # Encontra o melhor ponto de corte perto do meio
            # (onde há menos palavras = área de margem/gutter)
            melhor_corte = meio_pagina
            if len(xs) > 10:
                # Janela de busca: 30% a 70% da largura
                zona_corte = [x for x in xs if largura * 0.30 <= x <= largura * 0.70]
                if zona_corte:
                    # Encontra maior gap entre posições X consecutivas
                    gaps = [(zona_corte[i+1] - zona_corte[i], zona_corte[i])
                            for i in range(len(zona_corte)-1)]
                    if gaps:
                        melhor_corte = max(gaps, key=lambda g: g[0])[1] + \
                                       max(gaps, key=lambda g: g[0])[0] / 2

            # Separa palavras em coluna esquerda e direita
            words_esq = [w for w in words if float(w["x0"]) < melhor_corte]
            words_dir = [w for w in words if float(w["x0"]) >= melhor_corte]

            txt_esq = _reconstruir_linhas_por_coordenadas(words_esq)
            txt_dir = _reconstruir_linhas_por_coordenadas(words_dir)

            if txt_esq.strip():
                todas_partes.append(f"[PAG{num_pag+1}_ESQ]\n{txt_esq}")
            if txt_dir.strip():
                todas_partes.append(f"[PAG{num_pag+1}_DIR]\n{txt_dir}")

    texto_completo = "\n".join(todas_partes)
    return _corrigir_texto_espelhado(texto_completo)


def extrair_texto(conteudo: bytes) -> str:
    return extrair_texto_por_coluna(conteudo)


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
    ls = linha.strip()
    # Remove marcador de página/coluna
    ls = re.sub(r'^\[PAG\d+_(ESQ|DIR)\]\s*', '', ls)
    if not (3 <= len(ls) <= 70):
        return False
    base = re.sub(r'\s*\(.*?\)', '', ls).strip()
    if not base or base != base.upper():
        return False
    if re.search(r'\d', base):
        return False
    return _eh_unidade_capital(ls)


def _parse_linha_data(linha: str) -> dict | None:
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
    ls = linha.strip()
    if not ls or len(ls) < 3:
        return False
    if ls[0].islower():
        return False
    primeira = ls.split()[0].lower().rstrip('.,;:')
    if primeira in PREPOSICOES:
        return False
    ls_low = ls.lower()
    if any(ls_low.startswith(p) for p in PREFIXOS_DESC):
        return False
    if re.match(r'Dia[s]?\s+\d|De\s+\d|\d{1,2}/\d{1,2}', ls, re.I):
        return False
    # Descarta marcadores de página/coluna
    if re.match(r'\[PAG\d+', ls):
        return False
    return True


def _limpa_titulo(t: str) -> str:
    t = re.sub(r'\bFoto:.*', '', t, flags=re.I)
    t = re.sub(r'\s*\|\s*.*$', '', t)
    t = re.sub(r'^\[PAG\d+_(ESQ|DIR)\]\s*', '', t)
    return _norm(t)[:90]


def _detecta_secao_corte(linha: str) -> bool:
    """Retorna True se a linha é cabeçalho de seção não-musical."""
    ls = _norm(linha).upper()
    # Remove marcador de coluna para comparar
    ls = re.sub(r'^\[PAG\d+_(ESQ|DIR)\]\s*', '', ls)
    return any(ls == secao or ls.startswith(secao + " ") for secao in CORTE_SECAO)


def parse_musica(texto: str) -> list[dict]:
    # ── 1. Localiza início da seção MÚSICA ───────────────────────────────────
    # Busca tanto no texto normal quanto dentro dos marcadores de coluna
    idx = -1
    for pat in [r'\bMÚSICA\b', r'\bMUSICA\b']:
        for m in re.finditer(pat, texto, re.IGNORECASE):
            if idx == -1 or m.start() < idx:
                idx = m.start()

    if idx == -1:
        print("  ⚠️  Seção 'MÚSICA' não encontrada no PDF.")
        print("  Primeiros 500 chars do texto extraído:")
        print(texto[:500])
        return []

    secao = texto[idx:]

    # ── 2. Parse linha a linha dentro da seção MÚSICA ───────────────────────
    eventos       = []
    unidade_atual = None
    titulo_atual  = ""
    tem_desc      = False
    em_musica     = True   # flag: ainda estamos na seção MÚSICA

    for linha_raw in secao.split("\n"):
        linha = _norm(linha_raw)

        # Remove marcadores de coluna do texto visível
        linha_limpa = re.sub(r'^\[PAG\d+_(ESQ|DIR)\]\s*', '', linha).strip()

        if not linha_limpa:
            tem_desc = False
            continue

        # Detecta entrada em nova seção não-musical → para de coletar
        if _detecta_secao_corte(linha_limpa):
            em_musica = False
            unidade_atual = None  # reseta para não capturar da nova seção
            continue

        # Reativa coleta se voltarmos para um bloco de MÚSICA
        # (em PDFs com múltiplas colunas, MÚSICA pode aparecer em ambas)
        if re.match(r'\b(MÚSICA|MUSICA)\b', linha_limpa, re.I):
            em_musica = True
            continue

        if not em_musica:
            continue

        # Detecta linha de unidade Sesc
        if _eh_linha_unidade(linha_limpa):
            unidade_atual = linha_limpa
            titulo_atual  = ""
            tem_desc      = False
            continue

        if not unidade_atual:
            continue

        # Detecta linha de data/metadados → fecha o evento
        ev = _parse_linha_data(linha_limpa)
        if ev:
            if titulo_atual:
                ev["Artista / Título"] = _limpa_titulo(titulo_atual)
                ev["Unidade"]          = f"Sesc {unidade_atual.title()}"
                eventos.append(ev)
            titulo_atual = ""
            tem_desc     = False
            continue

        # Título ou descrição
        if not tem_desc:
            if _titulo_valido(linha_limpa):
                titulo_atual = _limpa_titulo(linha_limpa)
                tem_desc     = True
        # Linhas subsequentes são descrição — não sobrescreve o título

    return eventos


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
    print("  Extrator via PDF 'Em Cartaz' (pdfplumber + coords XY)")
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
            print("📄 Extraindo texto por coordenadas XY...")
            texto = extrair_texto(conteudo)
            if texto.strip():
                print(f"   Texto extraído: {len(texto):,} chars")
        else:
            print("❌ Falha ao baixar o PDF.")

    # Fallback: variável de ambiente com caminho local
    if not texto.strip():
        caminho_env = os.environ.get("PDF_LOCAL_PATH", "").strip().strip('"')
        if caminho_env and Path(caminho_env).exists():
            print(f"📄 Lendo PDF local (PDF_LOCAL_PATH): {caminho_env}")
            texto = ler_pdf_local(caminho_env)

    if not texto.strip():
        if sys.stdin.isatty():
            print("\n⚠️  PDF não obtido automaticamente.")
            print("   Baixe o PDF 'Em Cartaz' em sescsp.org.br/programacao")
            caminho_manual = input("   Informe o caminho do PDF: ").strip().strip('"')
            if caminho_manual and Path(caminho_manual).exists():
                texto = ler_pdf_local(caminho_manual)
            else:
                print("❌ Caminho inválido.")
                sys.exit(1)
        else:
            print("\n❌ PDF não encontrado automaticamente (ambiente CI).")
            print(f"   Mês buscado: {mes_ref}")
            print("   O PDF pode ainda não ter sido publicado pelo Sesc.")
            print("   Aguarde e re-execute o workflow quando estiver disponível.")
            sys.exit(1)

    if not texto.strip():
        print("❌ Não foi possível extrair texto do PDF.")
        sys.exit(1)

    print("🎵 Parseando programação musical da capital...")
    eventos = parse_musica(texto)

    if not eventos:
        print("\n⚠️  Nenhum evento encontrado.")
        idx = max(texto.find("MÚSICA"), texto.find("MUSICA"), 0)
        print("Trecho do texto extraído (seção MÚSICA):")
        print(texto[idx:idx+1000])
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
