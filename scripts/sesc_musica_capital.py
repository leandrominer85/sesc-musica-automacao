import os
import re
import sys
import time
import json
import urllib.parse
import types as py_types
import pymupdf as fitz
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
# 2. SELEÇÃO AUTOMÁTICA DO MODELO GEMINI (gemini-3.0-flash)
# =====================================================================

def obter_modelo_gemini(client: genai.Client) -> str:
    """Seleciona gemini-3.0-flash ou descobre dinamicamente o modelo ativo na chave."""
    preferidos = [
        "gemini-3.0-flash",
        "gemini-3-flash",
        "gemini-3.0-pro",
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash",
    ]
    try:
        disponiveis = []
        for m in client.models.list():
            nome = (m.name or "").replace("models/", "")
            acoes = (
                getattr(m, "supported_actions", None)
                or getattr(m, "supported_generation_methods", None)
                or []
            )
            if not acoes or "generateContent" in acoes:
                disponiveis.append(nome)
        print(f"ℹ️ Modelos Gemini disponíveis na chave: {disponiveis[:10]}")
        for pref in preferidos:
            if pref in disponiveis:
                print(f"✅ Modelo selecionado: {pref}")
                return pref
        for disp in disponiveis:
            if "flash" in disp.lower() and not any(x in disp.lower() for x in ["tts", "image", "audio", "embedding"]):
                print(f"✅ Modelo Flash selecionado: {disp}")
                return disp
    except Exception as e:
        print(f"⚠️ Aviso ao listar modelos: {e}")
    return "gemini-3.0-flash"


# =====================================================================
# 3. LOCALIZAÇÃO E DOWNLOAD AUTOMÁTICO DO PDF EM CARTAZ
# =====================================================================

PADRAO_URL_PDF_SESC = re.compile(
    r"https?://(?:www\.)?sescsp\.org\.br/wp-content/uploads/\d{4}/\d{2}/[^\s\"'<>\\]+?\.pdf",
    re.IGNORECASE,
)

PADRAO_RELATIVO_PDF_SESC = re.compile(
    r"/wp-content/uploads/\d{4}/\d{2}/[^\s\"'<>\\]+?\.pdf",
    re.IGNORECASE,
)


class _PdfCapturadoSucesso(Exception):
    pass


def tentar_baixar_com_codigo_original_do_repo(destino_pdf: Path) -> bool:
    """
    Recupera a lógica de download original do repositório (que gerou o PDF antes da refatoração)
    e intercepta o arquivo PDF assim que ele é entregue ao pdfplumber.open().
    """
    print("🔍 [Camada 1] Executando método original de download do repositório...")
    repo = os.environ.get("GITHUB_REPOSITORY", "leandrominer85/sesc-musica-automacao")
    token = os.environ.get("GITHUB_TOKEN", "")
    headers_gh = {"Accept": "application/vnd.github.v3+json", "User-Agent": "sesc-automacao"}
    if token:
        headers_gh["Authorization"] = f"Bearer {token}"

    try:
        url_commits = f"https://api.github.com/repos/{repo}/commits?path=scripts/sesc_musica_capital.py&per_page=15"
        r_commits = requests.get(url_commits, headers=headers_gh, timeout=20)
        if r_commits.status_code != 200:
            return False

        commits = r_commits.json()
        codigo_original = None
        for c in commits:
            sha = c["sha"]
            url_blob = f"https://raw.githubusercontent.com/{repo}/{sha}/scripts/sesc_musica_capital.py"
            r_code = requests.get(url_blob, headers=headers_gh, timeout=20)
            if r_code.status_code == 200 and "pydantic" not in r_code.text and "pdfplumber" in r_code.text:
                codigo_original = r_code.text
                print(f"   Encontrado downloader original (commit {sha[:7]}). Baixando PDF...")
                break

        if not codigo_original:
            return False

        modulo_fake = py_types.ModuleType("pdfplumber")

        def interceptar_open(file_or_path, *args, **kwargs):
            if isinstance(file_or_path, (str, Path)):
                destino_pdf.write_bytes(Path(file_or_path).read_bytes())
            elif isinstance(file_or_path, (bytes, bytearray)):
                destino_pdf.write_bytes(file_or_path)
            elif hasattr(file_or_path, "getvalue"):
                destino_pdf.write_bytes(file_or_path.getvalue())
            elif hasattr(file_or_path, "read"):
                if hasattr(file_or_path, "seek"):
                    file_or_path.seek(0)
                destino_pdf.write_bytes(file_or_path.read())
            raise _PdfCapturadoSucesso("PDF capturado com sucesso!")

        modulo_fake.open = interceptar_open
        modulo_anterior = sys.modules.get("pdfplumber")
        sys.modules["pdfplumber"] = modulo_fake

        try:
            escopo = {"__name__": "__main__", "__file__": str(Path(__file__).resolve())}
            exec(compile(codigo_original, "<original_sesc_download>", "exec"), escopo)
        except _PdfCapturadoSucesso:
            if destino_pdf.exists() and destino_pdf.stat().st_size > 10000:
                tamanho_mb = destino_pdf.stat().st_size / (1024 * 1024)
                print(f"✅ PDF baixado e capturado com sucesso ({tamanho_mb:.2f} MB)!")
                return True
        finally:
            if modulo_anterior is not None:
                sys.modules["pdfplumber"] = modulo_anterior
            else:
                sys.modules.pop("pdfplumber", None)

    except Exception as e:
        print(f"   Aviso na Camada 1: {e}")

    return False


def validar_e_baixar_url(url: str, destino_pdf: Path) -> bool:
    """Baixa a URL (seguindo redirecionamentos) e verifica se a assinatura binária é %PDF."""
    try:
        r = requests.get(url, headers=HEADERS_HTTP, timeout=45, stream=True, allow_redirects=True)
        if r.status_code != 200:
            r.close()
            return False

        chunks = []
        primeiro_chunk = None
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                if primeiro_chunk is None:
                    primeiro_chunk = chunk
                    if not primeiro_chunk.startswith(b"%PDF"):
                        r.close()
                        return False
                chunks.append(chunk)
        r.close()

        if primeiro_chunk and sum(len(c) for c in chunks) > 10000:
            with open(destino_pdf, "wb") as f:
                for c in chunks:
                    f.write(c)
            tamanho_mb = destino_pdf.stat().st_size / (1024 * 1024)
            print(f"✅ PDF baixado com sucesso ({tamanho_mb:.2f} MB): {r.url}")
            return True
    except Exception:
        pass
    return False


def extrair_pdfs_cartaz_de_texto(texto_bruto: str) -> List[str]:
    texto_limpo = texto_bruto.replace(r"\/", "/")
    texto_limpo = urllib.parse.unquote(texto_limpo)
    absolutos = PADRAO_URL_PDF_SESC.findall(texto_limpo)
    relativos = [f"https://www.sescsp.org.br{m}" for m in PADRAO_RELATIVO_PDF_SESC.findall(texto_limpo)]
    todos = absolutos + relativos
    pdfs_cartaz = [u for u in todos if "cartaz" in u.lower()]
    outros_pdfs = [u for u in todos if "cartaz" not in u.lower()]
    return list(dict.fromkeys(pdfs_cartaz + outros_pdfs))


def obter_e_baixar_pdf(agora: datetime, destino_pdf: Path) -> Path:
    if URL_PDF_FIXA:
        print(f"🔗 [Camada 0] Baixando da URL fixa: {URL_PDF_FIXA}")
        if validar_e_baixar_url(URL_PDF_FIXA, destino_pdf):
            return destino_pdf

    # CAMADA 1: Método original do próprio repositório
    if tentar_baixar_com_codigo_original_do_repo(destino_pdf):
        return destino_pdf

    # CAMADA 2: API REST do WordPress e páginas do Sesc SP
    print("🔍 [Camada 2] Consultando API WordPress e páginas do Sesc SP...")
    fontes_sesc = [
        "https://www.sescsp.org.br/wp-json/wp/v2/media?mime_type=application/pdf&per_page=100",
        "https://www.sescsp.org.br/wp-json/wp/v2/media?search=Cartaz&per_page=50",
        "https://www.sescsp.org.br/wp-json/wp/v2/posts?search=Em+Cartaz&per_page=20",
        "https://www.sescsp.org.br/wp-json/wp/v2/pages?search=Em+Cartaz&per_page=20",
        "https://www.sescsp.org.br/programacao/",
        "https://www.sescsp.org.br/revista-em-cartaz/",
        "https://www.sescsp.org.br/publicacoes/",
    ]
    for url_fonte in fontes_sesc:
        try:
            resp = requests.get(url_fonte, headers=HEADERS_HTTP, timeout=25)
            if resp.status_code == 200:
                for cand in extrair_pdfs_cartaz_de_texto(resp.text):
                    if validar_e_baixar_url(cand, destino_pdf):
                        return destino_pdf
        except Exception:
            continue

    # CAMADA 3: Gemini 3.0 Flash com Google Search Grounding
    api_key = os.environ.get("GEMINI_API_KEY")
    if api_key:
        print("🔍 [Camada 3] Localizando PDF via Gemini 3.0 Flash com Google Search...")
        try:
            client = genai.Client(api_key=api_key)
            modelo_ativo = obter_modelo_gemini(client)
            ano = agora.year
            mes_nome = MESES_PT[agora.month]
            prompt_busca = (
                f"Qual é a URL direta do arquivo PDF da revista 'Em Cartaz' do Sesc SP de {mes_nome} de {ano} "
                f"(ou a edição mais recente de {ano} hospedada em https://www.sescsp.org.br/wp-content/uploads/...)? "
                f"Retorne todas as URLs completas terminadas em .pdf encontradas."
            )
            resp_ia = client.models.generate_content(
                model=modelo_ativo,
                contents=prompt_busca,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.0,
                ),
            )
            texto_busca = resp_ia.text or ""
            urls_grounding = []
            if resp_ia.candidates and resp_ia.candidates[0].grounding_metadata:
                gm = resp_ia.candidates[0].grounding_metadata
                if gm.grounding_chunks:
                    for chunk in gm.grounding_chunks:
                        if chunk.web and chunk.web.uri:
                            urls_grounding.append(chunk.web.uri)
                        if chunk.web and chunk.web.title:
                            texto_busca += " " + chunk.web.title

            for cand in extrair_pdfs_cartaz_de_texto(texto_busca) + urls_grounding:
                if validar_e_baixar_url(cand, destino_pdf):
                    return destino_pdf
        except Exception as e:
            print(f"   Aviso Gemini Search: {e}")

    raise RuntimeError(
        "❌ Não foi possível baixar o PDF automaticamente. "
        "Cole o link direto do PDF na variável URL_PDF_FIXA no topo de scripts/sesc_musica_capital.py."
    )


# =====================================================================
# 4. FILTRO DE PÁGINAS COM PYMUPDF
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
# 5. EXTRAÇÃO ESTRUTURADA COM GEMINI 3.0 FLASH
# =====================================================================

def extrair_eventos_pdf(caminho_pdf: str) -> List[EventoMusical]:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("❌ A variável GEMINI_API_KEY não foi encontrada nos Secrets do GitHub!")

    client = genai.Client(api_key=api_key)
    modelo_ativo = obter_modelo_gemini(client)
    pdf_filtrado = str(PROJETO_DIR / "temp_secao_musica.pdf")

    qtd_paginas = filtrar_paginas_musica(caminho_pdf, pdf_filtrado)
    print(f"🤖 Enviando {qtd_paginas} página(s) para extração estruturada com {modelo_ativo}...")

    arquivo_remoto = client.files.upload(file=pdf_filtrado)

    # Aguarda caso o PDF grande entre em estado PROCESSING na API
    while getattr(arquivo_remoto, "state", None) and getattr(arquivo_remoto.state, "name", "") == "PROCESSING":
        print("⏳ Aguardando processamento do arquivo PDF nos servidores Gemini...")
        time.sleep(3)
        arquivo_remoto = client.files.get(name=arquivo_remoto.name)

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
            model=modelo_ativo,
            contents=[arquivo_remoto, prompt],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_schema=ListaProgramacao,
            ),
        )
        resultado = ListaProgramacao.model_validate_json(resposta.text)
    finally:
        try:
            client.files.delete(name=arquivo_remoto.name)
        except Exception:
            pass
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
# 6. GERAÇÃO DO EXCEL
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
# 7. EXECUÇÃO PRINCIPAL (MAIN)
# =====================================================================

def main():
    agora = datetime.now()
    mes_ano_label = f"{MESES_PT_CAPITAL[agora.month]} {agora.year}"
    nome_xlsx = f"sesc_musica_capital_{agora.year}_{agora.month:02d}.xlsx"
    caminho_xlsx = PROJETO_DIR / nome_xlsx
    caminho_pdf_temp = PROJETO_DIR / "em_cartaz_atual.pdf"

    try:
        obter_e_baixar_pdf(agora, caminho_pdf_temp)
        eventos = extrair_eventos_pdf(str(caminho_pdf_temp))
        salvar_planilha_excel(eventos, mes_ano_label, str(caminho_xlsx))
    finally:
        if caminho_pdf_temp.exists():
            caminho_pdf_temp.unlink()


if __name__ == "__main__":
    main()
