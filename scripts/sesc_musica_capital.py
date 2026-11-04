import os
import re
import sys
import time
import json
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
import pymupdf as fitz
import requests
from datetime import datetime
from pathlib import Path
from typing import List, Literal, Optional, Tuple
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

MESES_ABREV = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
}

MESES_PT_CAPITAL = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

SCRIPT_DIR = Path(__file__).parent
PROJETO_DIR = SCRIPT_DIR.parent

URL_PDF_FIXA = os.environ.get("URL_PDF_SESC", "").strip()
MES_ALVO_ENV = os.environ.get("MES_ALVO_SESC", "").strip()

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
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
# 2. FILA DE MODELOS GEMINI (COM FALLBACK AUTOMÁTICO)
# =====================================================================

def obter_fila_modelos_gemini(client: genai.Client) -> List[str]:
    modelos_bloqueados = {"gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"}
    disponiveis = []

    try:
        for m in client.models.list():
            nome = (m.name or "").replace("models/", "")
            if nome in modelos_bloqueados:
                continue
            if any(x in nome.lower() for x in ["tts", "image", "audio", "embedding", "vision", "aqa", "bison", "gecko"]):
                continue
            acoes = (
                getattr(m, "supported_actions", None)
                or getattr(m, "supported_generation_methods", None)
                or []
            )
            if not acoes or "generateContent" in acoes:
                disponiveis.append(nome)
    except Exception as e:
        print(f"⚠️ Aviso ao listar modelos: {e}")

    modelos_v3 = [m for m in disponiveis if "gemini-3" in m.lower()]

    preferidos_fixos = [
        "gemini-3.8-flash",
        "gemini-3.0-flash",
        "gemini-3-flash-preview",
        "gemini-flash-latest",
        "gemini-2.5-flash-lite",
        "gemini-flash-lite-latest",
        "gemini-2.5-pro",
        "gemini-pro-latest",
    ]

    outros_gemini = [m for m in disponiveis if m.startswith("gemini-")]

    fila = []
    for cand in modelos_v3 + preferidos_fixos + outros_gemini:
        if cand not in modelos_bloqueados and cand not in fila:
            fila.append(cand)

    if not fila:
        fila = ["gemini-3.8-flash", "gemini-3.0-flash", "gemini-flash-latest"]

    print(f"✅ Fila de modelos Gemini configurada: {fila[:6]}")
    return fila


# =====================================================================
# 3. LOCALIZAÇÃO E DOWNLOAD DO PDF (MÊS SEGUINTE / MÊS CORRENTE)
# =====================================================================

PADRAO_URL_PDF_SESC = re.compile(
    r"https?://(?:www\.)?sescsp\.org\.br/wp-content/uploads/\d{4}/\d{2}/[^\s\"'<>\\]+?\.pdf",
    re.IGNORECASE,
)

PADRAO_RELATIVO_PDF_SESC = re.compile(
    r"/wp-content/uploads/\d{4}/\d{2}/[^\s\"'<>\\]+?\.pdf",
    re.IGNORECASE,
)


def validar_e_baixar_url(url: str, destino_pdf: Path) -> bool:
    try:
        r = requests.get(url, headers=HEADERS_HTTP, timeout=30, stream=True, allow_redirects=True)
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


def gerar_urls_candidatas_para_mes(mes_idx: int, ano: int) -> List[str]:
    """
    Gera variações de URLs que o Sesc SP utiliza nos uploads do WordPress,
    testando tanto a pasta do próprio mês quanto a pasta do mês anterior e retrasado.
    """
    mes_nome = MESES_PT[mes_idx]
    mes_cap = MESES_PT_CAPITAL[mes_idx].replace("ç", "c")
    mes_abrev = MESES_ABREV[mes_idx]

    mes_ant1 = mes_idx - 1 if mes_idx > 1 else 12
    ano_ant1 = ano if mes_idx > 1 else ano - 1

    mes_ant2 = mes_ant1 - 1 if mes_ant1 > 1 else 12
    ano_ant2 = ano_ant1 if mes_ant1 > 1 else ano_ant1 - 1

    pastas = [
        f"{ano}/{mes_idx:02d}",
        f"{ano_ant1}/{mes_ant1:02d}",
        f"{ano_ant2}/{mes_ant2:02d}",
    ]

    nomes_arquivos = [
        f"Em-Cartaz_{ano}_{mes_nome}.pdf",
        f"Em-Cartaz_{ano}_{mes_nome}_compressed.pdf",
        f"Em-Cartaz-{mes_cap}-{ano}.pdf",
        f"Em-Cartaz-{ano}-{mes_nome}.pdf",
        f"em_cartaz_{mes_nome}{ano}_final_web.pdf",
        f"em_cartaz_{mes_nome}_{ano}_final_web.pdf",
        f"em_cartaz_{mes_nome}_{ano}_site.pdf",
        f"em-cartaz-{mes_nome}-{ano}-site.pdf",
        f"AF_EM-CARTAZ_{mes_abrev}{ano}_BaixaDuplas-1.pdf",
        f"AF_EM-CARTAZ_{mes_abrev}{ano}_BaixaDuplas.pdf",
        f"AF_EM-CARTAZ_{mes_abrev}_{ano}_BaixaDuplas-1.pdf",
    ]

    for v in range(7, 17):
        nomes_arquivos.extend([
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}_web.pdf",
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}_bx.pdf",
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}.pdf",
            f"V{v}_Em-Cartaz_{ano}_{mes_nome}_web.pdf",
            f"V{v}_Em-Cartaz_{ano}_{mes_nome}_bx.pdf",
        ])
        for dia in range(10, 26):
            nomes_arquivos.append(f"V{v}_Em-Cartaz_{ano}_{mes_nome}_{dia}-bx.pdf")

    urls = []
    for pasta in pastas:
        for nome in nomes_arquivos:
            urls.append(f"https://www.sescsp.org.br/wp-content/uploads/{pasta}/{nome}")
    return urls


def testar_urls_em_paralelo(urls: List[str], destino_pdf: Path) -> bool:
    """Testa centenas de URLs candidatas em paralelo em poucos segundos."""
    def checar_url(u: str) -> Optional[str]:
        try:
            r = requests.head(u, headers=HEADERS_HTTP, timeout=8, allow_redirects=True)
            if r.status_code == 200 and "pdf" in r.headers.get("Content-Type", "").lower():
                return u
        except Exception:
            pass
        return None

    with ThreadPoolExecutor(max_workers=20) as executor:
        futuros = {executor.submit(checar_url, u): u for u in urls}
        for fut in as_completed(futuros):
            url_valida = fut.result()
            if url_valida:
                print(f"🎯 URL encontrada no servidor do Sesc: {url_valida}")
                if validar_e_baixar_url(url_valida, destino_pdf):
                    return True
    return False


def definir_meses_candidatos(agora: datetime) -> List[Tuple[int, int]]:
    """
    Define a ordem de meses (mes, ano) para procurar o PDF:
    - Se MES_ALVO_SESC foi informado (ex: '12/2026'), busca exatamente esse mês.
    - Caso contrário, prioriza buscar os meses futuros disponíveis (ex: Dezembro / Novembro).
    """
    if MES_ALVO_ENV and MES_ALVO_ENV.lower() != "auto":
        m_match = re.match(r"^(\d{1,2})[/\-](\d{4})$", MES_ALVO_ENV.strip())
        if m_match:
            return [(int(m_match.group(1)), int(m_match.group(2)))]

    m_atual = agora.month
    a_atual = agora.year

    m_prox = m_atual + 1 if m_atual < 12 else 1
    a_prox = a_atual if m_atual < 12 else a_atual + 1

    m_seg = m_prox + 1 if m_prox < 12 else 1
    a_seg = a_prox if m_prox < 12 else a_prox + 1

    planilha_atual_existe = (PROJETO_DIR / f"sesc_musica_capital_{a_atual}_{m_atual:02d}.xlsx").exists()

    # Se estivermos no fim do mês (dia >= 20) OU se o mês atual já foi extraído,
    # tenta primeiro o mês seguinte (e m+2, como Dezembro) antes do mês atual
    if agora.day >= 20 or planilha_atual_existe:
        return [(m_prox, a_prox), (m_seg, a_seg), (m_atual, a_atual)]
    else:
        return [(m_prox, a_prox), (m_atual, a_atual)]


def obter_e_baixar_pdf(agora: datetime, destino_pdf: Path) -> Optional[Tuple[int, int]]:
    meses_candidatos = definir_meses_candidatos(agora)

    if URL_PDF_FIXA:
        print(f"🔗 [Camada 0] Baixando da URL informada: {URL_PDF_FIXA}")
        if validar_e_baixar_url(URL_PDF_FIXA, destino_pdf):
            # Tenta identificar o mês pelo próprio nome do arquivo na URL
            for m_idx, nome_m in MESES_PT.items():
                if nome_m in URL_PDF_FIXA.lower() or f"_{MESES_ABREV[m_idx].lower()}" in URL_PDF_FIXA.lower():
                    return (m_idx, agora.year)
            return meses_candidatos[0]

    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key) if api_key else None
    fila_modelos = obter_fila_modelos_gemini(client) if client else []

    for mes_idx, ano_val in meses_candidatos:
        mes_nome = MESES_PT[mes_idx]
        mes_abrev = MESES_ABREV[mes_idx].lower()
        print(f"\n🔎 === Verificando disponibilidade da revista de {MESES_PT_CAPITAL[mes_idx].upper()} {ano_val} ===")

        # 1. Tenta via Gemini Search Grounding (usando a fila de modelos válida)
        if client and fila_modelos:
            prompt_busca = (
                f"Localize a URL direta do arquivo PDF da revista 'Em Cartaz' do Sesc SP referente ao mês de "
                f"{mes_nome} de {ano_val} (hospedada em https://www.sescsp.org.br/wp-content/uploads/...). "
                f"Retorne todas as URLs diretas terminadas em .pdf."
            )
            for mod in fila_modelos[:3]:
                try:
                    print(f"🤖 [Camada 1] Buscando PDF de {mes_nome}/{ano_val} no Google via {mod}...")
                    resp_ia = client.models.generate_content(
                        model=mod,
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

                    for cand in extrair_pdfs_cartaz_de_texto(texto_busca) + urls_grounding:
                        if mes_nome in cand.lower() or f"_{mes_abrev}" in cand.lower() or f"-{mes_abrev}" in cand.lower():
                            if validar_e_baixar_url(cand, destino_pdf):
                                return (mes_idx, ano_val)
                    break
                except Exception as e:
                    print(f"   Aviso no modelo {mod}: {str(e)[:90]}")
                    continue

        # 2. Tenta via Buscadores Web (DuckDuckGo HTML)
        print(f"🌐 [Camada 2] Buscando indexação web para {mes_nome}/{ano_val}...")
        queries = [
            f'site:sescsp.org.br/wp-content/uploads/ "Cartaz" {mes_nome} {ano_val} filetype:pdf',
            f'site:sescsp.org.br "Em Cartaz" {mes_nome} {ano_val} filetype:pdf',
        ]
        for q in queries:
            try:
                ddg_url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(q)}"
                resp = requests.get(ddg_url, headers=HEADERS_HTTP, timeout=20)
                if resp.status_code == 200:
                    for cand in extrair_pdfs_cartaz_de_texto(resp.text):
                        if mes_nome in cand.lower() or mes_abrev in cand.lower():
                            if validar_e_baixar_url(cand, destino_pdf):
                                return (mes_idx, ano_val)
            except Exception:
                pass

        # 3. Tenta via API WordPress e páginas do Sesc SP filtrando pelo mês alvo
        print(f"📡 [Camada 3] Consultando servidor do Sesc SP para {mes_nome}/{ano_val}...")
        fontes_sesc = [
            f"https://www.sescsp.org.br/wp-json/wp/v2/media?search={mes_nome}&per_page=50",
            "https://www.sescsp.org.br/wp-json/wp/v2/media?search=Cartaz&per_page=100",
            "https://www.sescsp.org.br/revista-em-cartaz/",
            "https://www.sescsp.org.br/publicacoes/",
        ]
        for url_fonte in fontes_sesc:
            try:
                resp = requests.get(url_fonte, headers=HEADERS_HTTP, timeout=20)
                if resp.status_code == 200:
                    for cand in extrair_pdfs_cartaz_de_texto(resp.text):
                        if mes_nome in cand.lower() or mes_abrev in cand.lower():
                            if validar_e_baixar_url(cand, destino_pdf):
                                return (mes_idx, ano_val)
            except Exception:
                continue

        # 4. Tenta URLs candidatas em paralelo (nas pastas do próprio mês, mês anterior e retrasado)
        urls_cand = gerar_urls_candidatas_para_mes(mes_idx, ano_val)
        print(f"⚡ [Camada 4] Testando {len(urls_cand)} padrões de URL do Sesc em paralelo para {mes_nome}/{ano_val}...")
        if testar_urls_em_paralelo(urls_cand, destino_pdf):
            return (mes_idx, ano_val)

    return None


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
# 5. EXTRAÇÃO ESTRUTURADA COM GEMINI
# =====================================================================

def extrair_eventos_pdf(caminho_pdf: str) -> List[EventoMusical]:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("❌ A variável GEMINI_API_KEY não foi encontrada!")

    client = genai.Client(api_key=api_key)
    fila_modelos = obter_fila_modelos_gemini(client)
    pdf_filtrado = str(PROJETO_DIR / "temp_secao_musica.pdf")

    qtd_paginas = filtrar_paginas_musica(caminho_pdf, pdf_filtrado)
    print(f"📄 {qtd_paginas} página(s) da seção Música isoladas.")

    arquivo_remoto = client.files.upload(file=pdf_filtrado)

    while getattr(arquivo_remoto, "state", None) and getattr(arquivo_remoto.state, "name", "") == "PROCESSING":
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

    resultado = None
    ultimo_erro = None

    try:
        for modelo_atual in fila_modelos:
            try:
                print(f"🤖 Tentando extração com o modelo: {modelo_atual}...")
                resposta = client.models.generate_content(
                    model=modelo_atual,
                    contents=[arquivo_remoto, prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        response_mime_type="application/json",
                        response_schema=ListaProgramacao,
                    ),
                )
                resultado = ListaProgramacao.model_validate_json(resposta.text)
                print(f"✅ Extração concluída com sucesso usando: {modelo_atual} ({len(resultado.eventos)} shows)!")
                break
            except Exception as e:
                ultimo_erro = e
                msg_erro = str(e)
                print(f"⚠️ Modelo '{modelo_atual}' indisponível. Tentando próximo...")
                match_sugerido = re.search(r"use models/([a-zA-Z0-9._-]+)", msg_erro)
                if match_sugerido:
                    modelo_sugerido = match_sugerido.group(1)
                    if modelo_sugerido not in fila_modelos:
                        fila_modelos.append(modelo_sugerido)

        if resultado is None:
            raise RuntimeError(f"❌ Nenhum modelo conseguiu processar. Último erro: {ultimo_erro}")
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
    caminho_pdf_temp = PROJETO_DIR / "em_cartaz_atual.pdf"

    try:
        resultado_mes = obter_e_baixar_pdf(agora, caminho_pdf_temp)
        if not resultado_mes:
            print(
                "\nℹ️ Nenhum novo PDF de programação encontrado nos servidores hoje. "
                "Encerrando sem erros (status 0) — o workflow verificará novamente na próxima execução diária."
            )
            sys.exit(0)

        mes_encontrado, ano_encontrado = resultado_mes
        mes_ano_label = f"{MESES_PT_CAPITAL[mes_encontrado]} {ano_encontrado}"
        nome_xlsx = f"sesc_musica_capital_{ano_encontrado}_{mes_encontrado:02d}.xlsx"
        caminho_xlsx = PROJETO_DIR / nome_xlsx

        eventos = extrair_eventos_pdf(str(caminho_pdf_temp))
        salvar_planilha_excel(eventos, mes_ano_label, str(caminho_xlsx))
    finally:
        if caminho_pdf_temp.exists():
            caminho_pdf_temp.unlink()


if __name__ == "__main__":
    main()
