import os
import re
import sys
import json
import urllib.parse
import pymupdf as fitz  # Evita o FutureWarning do PyMuPDF no log
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
from typing import List, Literal, Optional
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================================
# CONFIGURAÇÕES E DIRETÓRIOS
# =====================================================================

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "marco", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

MESES_PT_CAPITAL = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

SCRIPT_DIR = Path(__file__).parent
PROJETO_DIR = SCRIPT_DIR.parent

# Opcional: se quiser forçar um PDF específico, cole a URL abaixo
URL_PDF_FIXA = os.environ.get("URL_PDF_SESC", "").strip()

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}


# =====================================================================
# 1. ESQUEMA DE VALIDAÇÃO ESTRITA (PYDANTIC)
# =====================================================================

UnidadesCapital = Literal[
    "Sesc 24 de Maio",
    "Sesc Avenida Paulista",
    "Sesc Belenzinho",
    "Sesc Bom Retiro",
    "Sesc Campo Limpo",
    "Sesc Carmo",
    "Sesc Casa Verde",
    "Sesc CineSesc",
    "Sesc Consolação",
    "Sesc Florêncio de Abreu",
    "Sesc Interlagos",
    "Sesc Ipiranga",
    "Sesc Itaquera",
    "Sesc Pinheiros",
    "Sesc Pompeia",
    "Sesc Santana",
    "Sesc Santo Amaro",
    "Sesc Vila Mariana",
]


class EventoMusical(BaseModel):
    artista_titulo: str = Field(
        description=(
            "APENAS o nome do artista, banda ou título oficial do show (em destaque no topo da nota). "
            "NUNCA inclua frases da sinopse (ex: 'Grupo apresenta...', 'Show inspirado...'), "
            "nem duração ('60 min.'), nem códigos como 'AL', 'VIVÊNCIA' ou 'BATE-PAPO'."
        )
    )
    unidade: UnidadesCapital = Field(
        description="Unidade oficial do Sesc na Capital. Ignore estações de metrô como 'Metrô São Bento'."
    )
    data: str = Field(
        description="Datas do show no formato DD/MM (ex: '24/10', '9 e 10/10', '1 a 12/10')."
    )
    horario: Optional[str] = Field(
        default=None,
        description="Horário do show (ex: '20h', '21h', '19h30'). Deixe null se não informado."
    )
    preco: Optional[str] = Field(
        default=None,
        description="Preços no formato 'R$XX / R$XX / R$XX' ou 'GRÁTIS'."
    )
    classificacao: Optional[Literal["AL", "A10", "A12", "A14", "A16", "A18"]] = Field(
        default=None,
        description="Classificação indicativa oficial: AL, A10, A12, A14, A16 ou A18."
    )


class ListaProgramacao(BaseModel):
    eventos: List[EventoMusical]


# =====================================================================
# 2. LOCALIZAÇÃO AUTOMÁTICA DO PDF EM 5 CAMADAS
# =====================================================================

PADRAO_URL_PDF_SESC = re.compile(
    r"https?://(?:www\.)?sescsp\.org\.br/wp-content/uploads/\d{4}/\d{2}/[^\s\"'<>\\]+?\.pdf",
    re.IGNORECASE,
)


def validar_url_pdf(url: str) -> bool:
    """Verifica se o link realmente existe e aponta para um PDF válido."""
    try:
        r = requests.get(url, headers=HEADERS_HTTP, timeout=20, stream=True)
        content_type = r.headers.get("Content-Type", "").lower()
        valido = r.status_code == 200 and ("pdf" in content_type or url.lower().endswith(".pdf"))
        r.close()
        return valido
    except Exception:
        return False


def extrair_pdfs_cartaz_de_texto(texto_bruto: str) -> List[str]:
    """Desescapa barras JSON (\\/) e extrai links de PDF que contenham 'cartaz'."""
    texto_limpo = texto_bruto.replace(r"\/", "/")
    texto_limpo = urllib.parse.unquote(texto_limpo)
    todos_pdfs = PADRAO_URL_PDF_SESC.findall(texto_limpo)
    pdfs_cartaz = [u for u in todos_pdfs if "cartaz" in u.lower()]
    # Remove duplicados preservando a ordem
    return list(dict.fromkeys(pdfs_cartaz))


def obter_url_pdf(agora: datetime) -> str:
    # CAMADA 0: URL fixa se configurada
    if URL_PDF_FIXA:
        print(f"🔗 [Camada 0] Usando URL fixa: {URL_PDF_FIXA}")
        return URL_PDF_FIXA

    # CAMADA 1: API REST do WordPress do Sesc SP (Mídias, Posts e Páginas)
    print("🔍 [Camada 1] Consultando API WordPress do Sesc SP...")
    endpoints_wp = [
        "https://www.sescsp.org.br/wp-json/wp/v2/media?mime_type=application/pdf&per_page=100",
        "https://www.sescsp.org.br/wp-json/wp/v2/media?search=Cartaz&per_page=50",
        "https://www.sescsp.org.br/wp-json/wp/v2/media?search=Em-Cartaz&per_page=50",
        "https://www.sescsp.org.br/wp-json/wp/v2/posts?search=Em+Cartaz&per_page=20",
        "https://www.sescsp.org.br/wp-json/wp/v2/pages?search=Em+Cartaz&per_page=20",
    ]
    for ep in endpoints_wp:
        try:
            resp = requests.get(ep, headers=HEADERS_HTTP, timeout=25)
            if resp.status_code == 200:
                candidatos = extrair_pdfs_cartaz_de_texto(resp.text)
                for url_cand in candidatos:
                    if validar_url_pdf(url_cand):
                        print(f"✅ PDF localizado via API WordPress: {url_cand}")
                        return url_cand
        except Exception as e:
            print(f"   Aviso WP ({ep[:55]}...): {e}")

    # CAMADA 2: Varredura das páginas do Sesc SP (desescapando JSON embutido)
    print("🔍 [Camada 2] Varrendo páginas públicas do Sesc SP...")
    paginas_sesc = [
        "https://www.sescsp.org.br/programacao/",
        "https://www.sescsp.org.br/revista-em-cartaz/",
        "https://www.sescsp.org.br/editorial/revista-em-cartaz/",
        "https://www.sescsp.org.br/publicacoes/",
        "https://www.sescsp.org.br/",
    ]
    for pag in paginas_sesc:
        try:
            resp = requests.get(pag, headers=HEADERS_HTTP, timeout=25)
            if resp.status_code == 200:
                candidatos = extrair_pdfs_cartaz_de_texto(resp.text)
                for url_cand in candidatos:
                    if validar_url_pdf(url_cand):
                        print(f"✅ PDF localizado na página {pag}: {url_cand}")
                        return url_cand
        except Exception as e:
            print(f"   Aviso página ({pag}): {e}")

    # CAMADA 3: Busca via DuckDuckGo HTML (captura nomes com prefixos como V11_Em-Cartaz_...)
    print("🔍 [Camada 3] Buscando PDF indexado via buscador web...")
    ano = agora.year
    mes_nome = MESES_PT[agora.month]
    queries_busca = [
        f'site:sescsp.org.br/wp-content/uploads/{ano} "Cartaz" filetype:pdf',
        f'site:sescsp.org.br "Em-Cartaz" {ano} {mes_nome} filetype:pdf',
        f'site:sescsp.org.br "Em Cartaz" {ano} filetype:pdf',
    ]
    for q in queries_busca:
        try:
            ddg_url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(q)}"
            resp = requests.get(ddg_url, headers=HEADERS_HTTP, timeout=25)
            if resp.status_code == 200:
                candidatos = extrair_pdfs_cartaz_de_texto(resp.text)
                for url_cand in candidatos:
                    if validar_url_pdf(url_cand):
                        print(f"✅ PDF localizado via busca web: {url_cand}")
                        return url_cand
        except Exception as e:
            print(f"   Aviso busca web: {e}")

    # CAMADA 4: Gemini 2.5 Flash com Google Search Grounding oficial
    api_key = os.environ.get("GEMINI_API_KEY")
    if api_key:
        print("🔍 [Camada 4] Usando Gemini com Google Search Grounding para localizar o PDF...")
        try:
            client = genai.Client(api_key=api_key)
            prompt_busca = (
                f"Encontre a URL direta do arquivo PDF da revista 'Em Cartaz' do Sesc SP de {ano} "
                f"(hospedado em https://www.sescsp.org.br/wp-content/uploads/...). "
                f"Retorne a URL completa terminada em .pdf."
            )
            resp_ia = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt_busca,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.0,
                ),
            )
            texto_busca = resp_ia.text or ""
            # Também inspeciona metadados de grounding se disponíveis
            if resp_ia.candidates and resp_ia.candidates[0].grounding_metadata:
                gm = resp_ia.candidates[0].grounding_metadata
                if gm.grounding_chunks:
                    for chunk in gm.grounding_chunks:
                        if chunk.web and chunk.web.uri:
                            texto_busca += " " + chunk.web.uri

            candidatos = extrair_pdfs_cartaz_de_texto(texto_busca)
            for url_cand in candidatos:
                if validar_url_pdf(url_cand):
                    print(f"✅ PDF localizado via Gemini Google Search: {url_cand}")
                    return url_cand
        except Exception as e:
            print(f"   Aviso Gemini Search: {e}")

    raise RuntimeError(
        "❌ Nenhuma das 4 camadas automáticas encontrou o link do PDF. "
        "Cole o link direto do PDF na variável URL_PDF_FIXA no topo de scripts/sesc_musica_capital.py."
    )


def baixar_pdf(url_pdf: str, destino: Path) -> Path:
    print(f"⬇️ Baixando PDF: {url_pdf}")
    resp = requests.get(url_pdf, headers=HEADERS_HTTP, timeout=60, stream=True)
    resp.raise_for_status()
    with open(destino, "wb") as f:
        for chunk in resp.iter_content(chunk_size=65536):
            if chunk:
                f.write(chunk)
    tamanho_mb = destino.stat().st_size / (1024 * 1024)
    print(f"✅ Download concluído ({tamanho_mb:.2f} MB)")
    return destino


# =====================================================================
# 3. FILTRO DE PÁGINAS COM PYMUPDF
# =====================================================================

def filtrar_paginas_musica(caminho_pdf_entrada: str, caminho_pdf_saida: str) -> int:
    doc = fitz.open(caminho_pdf_entrada)
    novo_doc = fitz.open()

    editorias_excluir = [
        "TEATRO", "DANÇA", "CIRCO", "CINEMA", "ARTES VISUAIS",
        "LITERATURA", "CRIANÇAS", "TURISMO", "IDOSOS", "ESPORTE", "JOVENS"
    ]

    em_secao_musica = False
    paginas_musica = []

    for i in range(len(doc)):
        pagina = doc[i]
        blocos = pagina.get_text("blocks")
        texto_cabecalho = " ".join(
            [b[4].upper() for b in blocos if b[1] < 130 or b[0] < 70 or b[0] > pagina.rect.width - 70]
        )

        if re.search(r"\bM[ÚU]SICA\b", texto_cabecalho):
            em_secao_musica = True
        elif em_secao_musica and any(re.search(rf"\b{ed}\b", texto_cabecalho) for ed in editorias_excluir):
            em_secao_musica = False

        if em_secao_musica:
            paginas_musica.append(i)

    if not paginas_musica:
        print("ℹ️ Tarja textual 'MÚSICA' não isolada; enviando PDF completo para a IA filtrar.")
        doc.save(caminho_pdf_saida)
        total = len(doc)
        doc.close()
        return total

    for p in paginas_musica:
        novo_doc.insert_pdf(doc, from_page=p, to_page=p)

    novo_doc.save(caminho_pdf_saida)
    total = len(novo_doc)
    novo_doc.close()
    doc.close()
    return total


# =====================================================================
# 4. EXTRAÇÃO COM GEMINI 2.5 FLASH
# =====================================================================

def extrair_eventos_pdf(caminho_pdf: str) -> List[EventoMusical]:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("❌ A variável GEMINI_API_KEY não foi encontrada nos Secrets do GitHub!")

    client = genai.Client(api_key=api_key)
    pdf_filtrado = str(PROJETO_DIR / "temp_secao_musica.pdf")

    qtd_paginas = filtrar_paginas_musica(caminho_pdf, pdf_filtrado)
    print(f"🤖 Enviando {qtd_paginas} página(s) para extração estruturada com Gemini 2.5 Flash...")

    arquivo_remoto = client.files.upload(file=pdf_filtrado)

    prompt = """
    Você está lendo a revista 'Em Cartaz' do Sesc São Paulo.
    Extraia TODOS os shows da seção de MÚSICA que ocorrem APENAS nas 18 unidades da CAPITAL de São Paulo.

    REGRAS OBRIGATÓRIAS:
    1. Cada página possui 3 colunas verticais. Leia cada coluna verticalmente de cima para baixo, sem jamais juntar textos de colunas vizinhas.
    2. No campo 'artista_titulo', extraia SOMENTE o título do show ou nome do artista/banda (que fica em negrito/destaque no início do bloco).
       PROIBIDO incluir frases da descrição/sinopse (ex: 'Grupo apresenta...', 'Show inspirado no...', 'Reúne músicas...').
       PROIBIDO incluir horários, preços, durações ('60 min.') ou classificações ('AL', 'A12') dentro de 'artista_titulo'.
    3. EXCLUA qualquer evento de Teatro, Dança, Circo, Infantil/Crianças, Contação de Histórias, Bate-papo, Cursos, Oficinas, Vivências ou Turismo Social.
    4. EXCLUA unidades do interior, litoral ou Grande SP (como Santo André, São Caetano, Guarulhos, Osasco, Mogi das Cruzes, Jundiaí, etc.).
    5. Atenção: 'Metrô São Bento', 'Metrô República' ou 'Casa Verde' no meio de um endereço são referências de localização; classifique apenas pela tarja oficial da Unidade do Sesc.
    """

    try:
        resposta = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[arquivo_remoto, prompt],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_schema=ListaProgramacao,
            ),
        )
        resultado = ListaProgramacao.model_validate_json(resposta.text)
    finally:
        client.files.delete(name=arquivo_remoto.name)
        if os.path.exists(pdf_filtrado):
            os.remove(pdf_filtrado)

    vistos = set()
    eventos_unicos = []
    for ev in resultado.eventos:
        chave = (ev.unidade, ev.artista_titulo.lower().strip(), ev.data.strip())
        if chave not in vistos:
            vistos.add(chave)
            eventos_unicos.append(ev)

    return sorted(eventos_unicos, key=lambda x: (x.unidade, x.data))


# =====================================================================
# 5. GERAÇÃO DO EXCEL
# =====================================================================

def salvar_planilha_excel(eventos: List[EventoMusical], mes_ano_label: str, caminho_xlsx: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Programação Musical"

    ws.merge_cells("A1:F1")
    ws["A1"] = f"Programação Musical – Sesc SP Capital | {mes_ano_label}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:F2")
    data_extracao = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws["A2"] = f"Extraído em: {data_extracao}   |   Fonte: PDF Em Cartaz – sescsp.org.br"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    colunas = ["Artista / Título", "Unidade", "Data", "Horário", "Preço", "Classificação"]
    for idx_col, nome_col in enumerate(colunas, start=1):
        celula = ws.cell(row=3, column=idx_col, value=nome_col)
        celula.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="8B0000")
        celula.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 24

    borda = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for idx_linha, ev in enumerate(eventos, start=4):
        cor_zebra = "F5F5F5" if idx_linha % 2 == 0 else "FFFFFF"
        preenchimento = PatternFill("solid", fgColor=cor_zebra)
        valores = [ev.artista_titulo, ev.unidade, ev.data, ev.horario, ev.preco, ev.classificacao]

        for idx_col, valor in enumerate(valores, start=1):
            c = ws.cell(row=idx_linha, column=idx_col, value=valor)
            c.number_format = "@"
            c.fill = preenchimento
            c.border = borda
            c.alignment = Alignment(
                horizontal="left" if idx_col in (1, 2) else "center",
                vertical="center",
                wrap_text=True,
            )

    linha_final = len(eventos) + 4
    ws.merge_cells(f"A{linha_final}:F{linha_final}")
    ws[f"A{linha_final}"] = f"Total de eventos musicais: {len(eventos)}"
    ws[f"A{linha_final}"].font = Font(name="Calibri", size=10, bold=True)
    ws[f"A{linha_final}"].alignment = Alignment(horizontal="right", vertical="center")

    larguras = {"A": 42.0, "B": 24.0, "C": 22.0, "D": 14.0, "E": 22.0, "F": 16.0}
    for col_letra, largura in larguras.items():
        ws.column_dimensions[col_letra].width = largura

    wb.save(caminho_xlsx)
    print(f"✅ Planilha gerada com sucesso: {caminho_xlsx} ({len(eventos)} shows válidos)")


# =====================================================================
# 6. EXECUÇÃO PRINCIPAL (MAIN)
# =====================================================================

def main():
    agora = datetime.now()
    mes_ano_label = f"{MESES_PT_CAPITAL[agora.month]} {agora.year}"
    nome_xlsx = f"sesc_musica_capital_{agora.year}_{agora.month:02d}.xlsx"
    caminho_xlsx = PROJETO_DIR / nome_xlsx
    caminho_pdf_temp = PROJETO_DIR / "em_cartaz_atual.pdf"

    try:
        url_pdf = obter_url_pdf(agora)
        baixar_pdf(url_pdf, caminho_pdf_temp)
        eventos = extrair_eventos_pdf(str(caminho_pdf_temp))
        salvar_planilha_excel(eventos, mes_ano_label, str(caminho_xlsx))
    finally:
        if caminho_pdf_temp.exists():
            caminho_pdf_temp.unlink()


if __name__ == "__main__":
    main()
