"""
sesc_musica_capital.py
======================
Extrai programação de MÚSICA das unidades da CAPITAL do Sesc SP
e salva como planilha Excel (.xlsx).

ESTRATÉGIA:
  1. Descobre automaticamente a URL do PDF "Em Cartaz" do mês atual
  2. Baixa e extrai o texto do PDF via pypdf
  3. Parseia a seção MÚSICA filtrando apenas unidades da capital
  4. Gera o .xlsx formatado e pronto para uso

Dependências (instalar apenas uma vez):
    pip install requests pypdf openpyxl

Uso:
    python sesc_musica_capital.py

Arquivo gerado: sesc_musica_capital_AAAA_MM.xlsx
"""

import re
import sys
import io
from datetime import datetime
from pathlib import Path

# ── Verifica dependências ─────────────────────────────────────────────────────
def _check(pkg, install):
    try:
        __import__(pkg)
    except ImportError:
        print(f"❌ '{pkg}' não encontrado. Instale com:\n   pip install {install}")
        sys.exit(1)

_check("requests", "requests")
_check("pypdf", "pypdf")
_check("openpyxl", "openpyxl")

import requests
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

MESES_PT = {
    1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
    7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro",
}

# Unidades Sesc na capital de SP
UNIDADES_CAPITAL = [
    "14 BIS","24 DE MAIO","AVENIDA PAULISTA","BELENZINHO","BOM RETIRO",
    "CAMPO LIMPO","CASA VERDE","CONSOLAÇÃO","INTERLAGOS","IPIRANGA",
    "ITAQUERA","PINHEIROS","POMPEIA","SANTANA","SANTO AMARO","SÃO BENTO",
    "SÃO JOÃO CLÍMACO","VILA MARIANA","CARMO","PARQUE DOM PEDRO",
    "BRAS", "BRÁS",
]

# Palavras-chave que marcam início de seção não-musical → interrompe a coleta
CORTE_SECAO = ["TEATRO","DANÇA","CIRCO","EXPOSIÇÕES","CINEMA","CRIANÇAS",
               "JOVENS","PESSOAS IDOSAS","ESPORTE","TURISMO SOCIAL",
               "TRABALHO E","MEIO AMBIENTE","DIREITOS HUMANOS","TECNOLOGIAS",
               "SAÚDE","ALIMENTAÇÃO","LITERATURA","CENTRO DE PESQUISA","LOJAS SESC"]

# Estilo Excel
COR_VERMELHO  = "C00000"
COR_VERM_ESC  = "8B0000"
COR_BRANCO    = "FFFFFF"
COR_CINZA     = "F5F5F5"

COLUNAS  = ["Artista / Título", "Unidade", "Data", "Horário", "Preço", "Classificação"]
LARGURAS = [42, 24, 26, 12, 22, 16]

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/pdf,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9",
    "Referer": "https://www.sescsp.org.br/",
}

# ─────────────────────────────────────────────────────────────────────────────
# PASSO 1 – LOCALIZAR O PDF "Em Cartaz"
# ─────────────────────────────────────────────────────────────────────────────

def _pdf_existe(url: str) -> bool:
    try:
        r = requests.head(url, headers=HEADERS_HTTP, timeout=10, allow_redirects=True)
        ct = r.headers.get("Content-Type", "")
        return r.status_code == 200 and "pdf" in ct
    except Exception:
        return False


def descobrir_url_pdf(ano: int, mes: int) -> str | None:
    """Testa variações comuns da URL do PDF 'Em Cartaz'."""
    m = MESES_PT[mes]
    mc = m.capitalize()
    # Mês anterior (publicado na última semana do mês anterior)
    mes_pub = mes - 1 if mes > 1 else 12
    ano_pub = ano if mes > 1 else ano - 1

    candidatos = [
        # Padrão observado em 2026 (V11 = versão 11)
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V12_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/V10_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        # Padrões mais simples
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz_{ano}_{m}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz-{m}-{ano}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/em-cartaz-{m}-{ano}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mes:02d}/Em-Cartaz_{mc}_{ano}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano_pub}/{mes_pub:02d}/Em-Cartaz_{ano}_{m}.pdf",
    ]

    for url in candidatos:
        print(f"  🔎 Testando: {url.split('uploads/')[-1]}")
        if _pdf_existe(url):
            return url
    return None


def buscar_pdf_no_html(ano: int, mes: int) -> str | None:
    """Fallback: raspa o HTML do site buscando links de PDF 'Em Cartaz'."""
    m = MESES_PT[mes]
    for url_busca in ["https://www.sescsp.org.br/programacao/", "https://www.sescsp.org.br/"]:
        try:
            r = requests.get(url_busca, headers=HEADERS_HTTP, timeout=15)
            pdfs = re.findall(
                r'https://www\.sescsp\.org\.br/wp-content/uploads/[^\s"\'<>]*'
                r'[Ee]m[- _][Cc]artaz[^\s"\'<>]*\.pdf',
                r.text
            )
            for pdf_url in pdfs:
                if m in pdf_url.lower() or str(ano) in pdf_url:
                    return pdf_url
        except Exception:
            pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 2 – BAIXAR E EXTRAIR TEXTO DO PDF
# ─────────────────────────────────────────────────────────────────────────────

def baixar_pdf(url: str) -> bytes | None:
    try:
        r = requests.get(url, headers=HEADERS_HTTP, timeout=45)
        return r.content if r.status_code == 200 else None
    except Exception as e:
        print(f"  ❌ Erro ao baixar PDF: {e}")
        return None


def extrair_texto_pdf(conteudo: bytes) -> str:
    reader = PdfReader(io.BytesIO(conteudo))
    partes = []
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            partes.append(txt)
    return "\n".join(partes)


def ler_pdf_local(caminho: str) -> str:
    with open(caminho, "rb") as f:
        return extrair_texto_pdf(f.read())


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3 – PARSER DA SEÇÃO MÚSICA
# ─────────────────────────────────────────────────────────────────────────────

def _normalizar(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()


def _eh_unidade_capital(linha: str) -> bool:
    lu = linha.upper().strip()
    return any(u in lu for u in UNIDADES_CAPITAL)


def _eh_linha_unidade(linha: str) -> bool:
    """
    Linha de unidade: toda em maiúsculas, sem dígitos, entre 3 e 40 chars.
    Permite parênteses como em 'BELENZINHO (PERIFERIAS: DO TAMBOR AO SISTEMA DE SOM)'.
    """
    ls = linha.strip()
    if not ls or len(ls) < 3 or len(ls) > 60:
        return False
    # Remove conteúdo entre parênteses para checar caixa alta
    base = re.sub(r'\(.*?\)', '', ls).strip()
    if not base:
        return False
    if base != base.upper():
        return False
    if re.search(r'\d', ls):
        return False
    return _eh_unidade_capital(ls)


def _parse_linha_evento(linha: str) -> dict | None:
    """
    Tenta extrair data, horário, preço e classificação de uma linha.
    Retorna dict com os campos ou None se não for linha de evento.
    """
    # Padrão: "Dia(s) DD[/e/a DD]/MM. Dia-da-semana, HHh[mm]. [A12] [R$XX] [GRÁTIS]"
    m = re.match(
        r'(?:Dia[s]?\s+|De\s+)?'
        r'(\d{1,2}(?:\s*(?:e|a)\s*\d{1,2})?(?:\s*,\s*\d{1,2})?/\d{1,2}(?:\s*(?:e|a|,)\s*\d{1,2}/\d{1,2})?)'
        r'([^.]*)\.\s*(.*)',
        linha, re.IGNORECASE
    )
    if not m:
        return None

    data_raw = _normalizar(m.group(1) + m.group(2))
    resto    = _normalizar(m.group(3))

    hora_m  = re.search(r'(\d{1,2}h\d{0,2}(?:\s*às\s*\d{1,2}h\d{0,2})?)', resto, re.IGNORECASE)
    preco_m = re.search(r'(GRÁTIS|R\$\s*\d[\d,\.]*(?:\s+R\$\s*\d[\d,\.]*)*)', resto, re.IGNORECASE)
    clas_m  = re.search(r'\b(A(?:L|\d{1,2})|Livre|Todas as idades)\b', resto, re.IGNORECASE)

    return {
        "Data":          data_raw,
        "Horário":       hora_m.group(1)  if hora_m  else "",
        "Preço":         preco_m.group(1) if preco_m else "",
        "Classificação": clas_m.group(1)  if clas_m  else "",
    }


def _limpa_titulo(t: str) -> str:
    """Remove prefixos de tipo de evento e texto de foto/crédito."""
    t = re.sub(r'^(SHOW|BAILE|RODA|APRESENTAÇÃO|CONCERTO|RECITAL|FESTIVAL|BAILE)\s+', '', t, flags=re.I)
    t = re.sub(r'\bFoto:\s*\S+', '', t, flags=re.I)
    return _normalizar(t)[:90]


def parse_musica(texto: str) -> list[dict]:
    # ── 1. Isola seção MÚSICA ─────────────────────────────────────────────────
    idx = max(
        (m.start() for pat in [r'\bMÚSICA\b', r'\bMUSICA\b']
         for m in [re.search(pat, texto, re.IGNORECASE)] if m),
        default=-1
    )
    if idx == -1:
        print("  ⚠️  Palavra 'MÚSICA' não encontrada no PDF.")
        return []
    texto = texto[idx:]

    # ── 2. Corta na próxima seção não-musical ────────────────────────────────
    for secao in CORTE_SECAO:
        c = re.search(rf'\n\s*{re.escape(secao)}\s*\n', texto, re.IGNORECASE)
        if c and c.start() > 100:
            texto = texto[:c.start()]
            break

    # ── 3. Percorre linha a linha ─────────────────────────────────────────────
    eventos      = []
    unidade_atual = None
    titulo_atual  = ""

    for linha_raw in texto.split("\n"):
        linha = _normalizar(linha_raw)
        if not linha:
            continue

        # Detecta unidade
        if _eh_linha_unidade(linha):
            if unidade_atual and titulo_atual:
                # Título sem data → provavelmente é série/projeto; ignora
                titulo_atual = ""
            unidade_atual = linha
            titulo_atual  = ""
            continue

        if not unidade_atual:
            continue

        # Ignora créditos de foto e marcadores de seção interna
        if re.match(r'Foto:', linha, re.I) or linha.startswith("#"):
            continue

        # Tenta parsear como linha de data/evento
        ev = _parse_linha_evento(linha)
        if ev:
            if titulo_atual:
                ev["Artista / Título"] = _limpa_titulo(titulo_atual)
                ev["Unidade"]          = f"Sesc {unidade_atual.title()}"
                eventos.append(ev)
            titulo_atual = ""
        else:
            # Acumula como título (pode ser nome do artista + subtítulo)
            if titulo_atual:
                # Se a linha nova parece ser a descrição, substitui apenas se o
                # título atual for só o nome do artista (sem descrição)
                titulo_atual = _limpa_titulo(linha) or titulo_atual
            else:
                titulo_atual = _limpa_titulo(linha)

    return eventos


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4 – GERAR O EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel(eventos: list[dict], caminho: Path, mes_ref: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Programação Musical"
    nc = len(COLUNAS)

    # Linha 1: Título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(row=1, column=1)
    c.value     = f"Programação Musical – Sesc SP Capital | {mes_ref}"
    c.font      = Font(name="Arial", bold=True, size=13, color=COR_BRANCO)
    c.fill      = PatternFill("solid", fgColor=COR_VERMELHO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Linha 2: Subtítulo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    c = ws.cell(row=2, column=1)
    c.value     = (f"Extraído em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                   "   |   Fonte: PDF Em Cartaz – sescsp.org.br")
    c.font      = Font(name="Arial", size=9, italic=True, color="555555")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Linha 3: Cabeçalho
    for ci, nome in enumerate(COLUNAS, 1):
        c = ws.cell(row=3, column=ci, value=nome)
        c.font      = Font(name="Arial", bold=True, size=10, color=COR_BRANCO)
        c.fill      = PatternFill("solid", fgColor=COR_VERM_ESC)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = Border(bottom=Side(style="thin", color=COR_BRANCO))
    ws.row_dimensions[3].height = 22

    # Dados
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

    # Rodapé com total
    lr = len(eventos) + 4
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=nc)
    c = ws.cell(row=lr, column=1)
    c.value     = f"Total de eventos musicais: {len(eventos)}"
    c.font      = Font(name="Arial", bold=True, size=10, color=COR_BRANCO)
    c.fill      = PatternFill("solid", fgColor=COR_VERMELHO)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[lr].height = 20

    # Larguras, filtro e congelamento
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
    ano     = agora.year
    mes     = agora.month
    mes_ref = f"{MESES_PT[mes].capitalize()} {ano}"

    print(f"\n📅 Mês de referência: {mes_ref}")

    # ── Localiza o PDF ────────────────────────────────────────────────────────
    print("\n🔍 Procurando PDF 'Em Cartaz'...")
    url_pdf = descobrir_url_pdf(ano, mes)

    texto = ""

    if not url_pdf:
        print("  ⚠️  Buscando link no HTML do site...")
        url_pdf = buscar_pdf_no_html(ano, mes)

    if url_pdf:
        print(f"\n📥 Baixando PDF: {url_pdf.split('uploads/')[-1]}")
        conteudo = baixar_pdf(url_pdf)
        if conteudo:
            print("📄 Extraindo texto...")
            texto = extrair_texto_pdf(conteudo)
        else:
            print("❌ Falha ao baixar o PDF.")
    
    if not texto.strip():
        print("\n⚠️  Não foi possível obter o PDF automaticamente.")
        print("   Acesse https://www.sescsp.org.br/programacao/ e baixe o PDF 'Em Cartaz'.")
        caminho_manual = input("   Informe o caminho do PDF baixado (ou Enter para sair): ").strip().strip('"')
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

    # ── Parseia eventos musicais ──────────────────────────────────────────────
    print("🎵 Extraindo programação musical da capital...")
    eventos = parse_musica(texto)

    if not eventos:
        print("\n⚠️  Nenhum evento musical encontrado.")
        print("   O formato do PDF pode ter mudado.")
        print("   Primeiros 600 caracteres extraídos do PDF:")
        print(texto[:600])
        sys.exit(1)

    # Ordena por unidade → data
    eventos.sort(key=lambda x: (x.get("Unidade", ""), x.get("Data", "")))

    # ── Gera o Excel ──────────────────────────────────────────────────────────
    nome    = f"sesc_musica_capital_{ano}_{mes:02d}.xlsx"
    destino = Path(__file__).parent / nome

    print(f"\n📊 Gerando planilha: {nome}")
    gerar_excel(eventos, destino, mes_ref)

    print(f"\n✅ Concluído!")
    print(f"   Arquivo: {destino}")
    print(f"   Total de eventos musicais na capital: {len(eventos)}")


if __name__ == "__main__":
    main()
