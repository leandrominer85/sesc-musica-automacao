import os
import re
import sys
import time
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
}

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
        description="APENAS o nome do artista, banda ou título oficial do show. Sem frases da sinopse."
    )
    unidade: UnidadesCapital = Field(
        description="Unidade oficial do Sesc na Capital."
    )
    data: str = Field(
        description="Datas do show no formato DD/MM (ex: '14 e 15/11', '06/11')."
    )
    horario: Optional[str] = Field(
        default=None,
        description="Horário do show (ex: '20h', '21h')."
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
            disponiveis.append(nome)
    except Exception:
        pass

    modelos_v3 = [m for m in disponiveis if "gemini-3" in m.lower()]
    preferidos = [
        "gemini-3.8-flash",
        "gemini-3.0-flash",
        "gemini-flash-latest",
        "gemini-2.5-flash-lite",
        "gemini-2.5-pro",
    ]
    outros = [m for m in disponiveis if m.startswith("gemini-")]

    fila = []
    for c in modelos_v3 + preferidos + outros:
        if c not in modelos_bloqueados and c not in fila:
            fila.append(c)
    return fila or ["gemini-3.8-flash", "gemini-flash-latest"]


def validar_e_baixar_url(url: str, destino_pdf: Path) -> bool:
    try:
        r = requests.get(url, headers=HEADERS_HTTP, timeout=30, stream=True, allow_redirects=True)
        if r.status_code != 200:
            r.close()
            return False
        chunks = []
        primeiro = None
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                if primeiro is None:
                    primeiro = chunk
                    if not primeiro.startswith(b"%PDF"):
                        r.close()
                        return False
                chunks.append(chunk)
        r.close()
        if primeiro and sum(len(c) for c in chunks) > 10000:
            with open(destino_pdf, "wb") as f:
                for c in chunks:
                    f.write(c)
            print(f"✅ PDF baixado ({destino_pdf.stat().st_size / (1024*1024):.2f} MB): {r.url}")
            return True
    except Exception:
        pass
    return False


def gerar_urls_candidatas_para_mes(mes_idx: int, ano: int) -> List[str]:
    mes_nome = MESES_PT[mes_idx]
    mes_cap = MESES_PT_CAPITAL[mes_idx].replace("ç", "c")
    mes_abrev = MESES_ABREV[mes_idx]

    mes_ant = mes_idx - 1 if mes_idx > 1 else 12
    ano_ant = ano if mes_idx > 1 else ano - 1

    pastas = [f"{ano_ant}/{mes_ant:02d}", f"{ano}/{mes_idx:02d}"]

    nomes = [
        f"V13_Em-Cartaz_{ano}_{mes_nome}_VF.pdf",
        f"Em-Cartaz_{ano}_{mes_nome}.pdf",
        f"Em-Cartaz_{ano}_{mes_nome}_VF.pdf",
        f"Em-Cartaz_{ano}_{mes_nome}_compressed.pdf",
        f"Em-Cartaz-{mes_cap}-{ano}.pdf",
        f"em_cartaz_{mes_nome}{ano}_final_web.pdf",
        f"em_cartaz_{mes_nome}_{ano}_site.pdf",
    ]
    for v in range(8, 18):
        nomes.extend([
            f"V{v}_Em-Cartaz_{ano}_{mes_nome}_VF.pdf",
            f"V{v}_Em-Cartaz_{ano}_{mes_nome}_web.pdf",
            f"V{v}_Em-Cartaz_{ano}_{mes_nome}_bx.pdf",
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}_VF.pdf",
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}_bx.pdf",
            f"V{v}_EM-CARTAZ_{mes_abrev}_{ano}_web.pdf",
        ])
        for dia in range(8, 25):
            nomes.append(f"V{v}_Em-Cartaz_{ano}_{mes_nome}_{dia:02d}-bx.pdf")
            nomes.append(f"V{v}_Em-Cartaz_{ano}_{mes_nome}_{dia:02d}-sem-marcas.pdf")

    return [f"https://www.sescsp.org.br/wp-content/uploads/{p}/{n}" for p in pastas for n in nomes]


def testar_urls_em_paralelo(urls: List[str], destino_pdf: Path) -> bool:
    def checar(u: str) -> Optional[str]:
        try:
            r = requests.head(u, headers=HEADERS_HTTP, timeout=8, allow_redirects=True)
            if r.status_code == 200 and "pdf" in r.headers.get("Content-Type", "").lower():
                return u
        except Exception:
            pass
        return None

    with ThreadPoolExecutor(max_workers=24) as ex:
        futs = {ex.submit(checar, u): u for u in urls}
        for f in as_completed(futs):
            u_ok = f.result()
            if u_ok and validar_e_baixar_url(u_ok, destino_pdf):
                return True
    return False


def definir_meses_candidatos(agora: datetime) -> List[Tuple[int, int]]:
    if MES_ALVO_ENV and MES_ALVO_ENV.lower() != "auto":
        m = re.match(r"^(\d{1,2})[/\-](\d{4})$", MES_ALVO_ENV.strip())
        if m:
            return [(int(m.group(1)), int(m.group(2)))]

    m_at, a_at = agora.month, agora.year
    m_pr = m_at + 1 if m_at < 12 else 1
    a_pr = a_at if m_at < 12 else a_at + 1

    if agora.day >= 22:
        return [(m_pr, a_pr), (m_at, a_at)]
    return [(m_at, a_at), (m_pr, a_pr)]


def obter_e_baixar_pdf(agora: datetime, destino_pdf: Path) -> Optional[Tuple[int, int]]:
    candidatos = definir_meses_candidatos(agora)
    if URL_PDF_FIXA and validar_e_baixar_url(URL_PDF_FIXA, destino_pdf):
        return candidatos[0]

    for m_idx, a_val in candidatos:
        print(f"\n🔎 Buscando PDF de {MESES_PT_CAPITAL[m_idx].upper()} {a_val}...")
        urls = gerar_urls_candidatas_para_mes(m_idx, a_val)
        if testar_urls_em_paralelo(urls, destino_pdf):
            return (m_idx, a_val)
    return None


def selecionar_paginas_musica(caminho_pdf: str) -> List[int]:
    """
    Identifica todas as páginas da revista que contêm programação de MÚSICA
    ou cabeçalhos das unidades na seção de Shows, sem desligar quando encontrar 'Local: Teatro'.
    """
    doc = fitz.open(caminho_pdf)
    paginas_alvo = []

    for i in range(len(doc)):
        pag = doc[i]
        # Verifica apenas os 70px do topo e rodapé onde fica a editoria da página
        blocos = pag.get_text("blocks")
        faixa_editorial = " ".join([b[4].upper() for b in blocos if b[1] < 75 or b[3] > pag.rect.height - 65])
        texto_pag = pag.get_text("text")

        # Captura páginas da editoria MÚSICA ou páginas de Destaques/Shows que contenham 'SHOW' e preços/datas
        if re.search(r"\bM[ÚU]SICA\b", faixa_editorial) or re.search(r"\bSHOW\b", texto_pag):
            paginas_alvo.append(i)

    doc.close()
    return paginas_alvo if paginas_alvo else list(range(len(fitz.open(caminho_pdf))))


def extrair_eventos_em_lotes(caminho_pdf: str, mes_idx: int) -> List[EventoMusical]:
    """
    Divide as páginas selecionadas em lotes de 3 páginas por chamada ao Gemini,
    impedindo que a IA trunque a lista após ~38 shows e garantindo todas as 18 unidades.
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)
    fila_modelos = obter_fila_modelos_gemini(client)
    modelo_ativo = None

    paginas_indices = selecionar_paginas_musica(caminho_pdf)
    print(f"📄 Total de páginas com shows/música identificadas no PDF: {len(paginas_indices)}")

    doc_orig = fitz.open(caminho_pdf)
    todos_eventos: List[EventoMusical] = []
    tamanho_lote = 3

    prompt = f"""
    Analise estas páginas da revista 'Em Cartaz' do Sesc São Paulo (mês {mes_idx:02d}).
    Extraia ABSOLUTAMENTE TODOS os shows musicais das 18 unidades da CAPITAL de São Paulo:
    (24 de Maio, Avenida Paulista, Belenzinho, Bom Retiro, Campo Limpo, Carmo, Casa Verde, CineSesc,
    Consolação, Florêncio de Abreu, Interlagos, Ipiranga, Itaquera, Pinheiros, Pompeia, Santana, Santo Amaro, Vila Mariana).

    REGRAS:
    1. Leia as 3 colunas de cada página verticalmente sem omitir nenhuma unidade ou artista.
    2. Em 'artista_titulo', coloque APENAS o nome do artista/banda/show em destaque.
    3. Ignore peças de teatro, dança, circo, bate-papos, oficinas, cursos ou unidades do interior/ABC/Guarulhos/Osasco.
    """

    for inicio in range(0, len(paginas_indices), tamanho_lote):
        lote_pags = paginas_indices[inicio:inicio + tamanho_lote]
        doc_lote = fitz.open()
        for p in lote_pags:
            doc_lote.insert_pdf(doc_orig, from_page=p, to_page=p)

        caminho_lote = str(PROJETO_DIR / f"temp_lote_{inicio}.pdf")
        doc_lote.save(caminho_lote)
        doc_lote.close()

        arq_remoto = client.files.upload(file=caminho_lote)
        try:
            modelos_tentar = [modelo_ativo] if modelo_ativo else fila_modelos
            for mod in modelos_tentar:
                try:
                    resp = client.models.generate_content(
                        model=mod,
                        contents=[arq_remoto, prompt],
                        config=types.GenerateContentConfig(
                            temperature=0.0,
                            response_mime_type="application/json",
                            response_schema=ListaProgramacao,
                        ),
                    )
                    res_lote = ListaProgramacao.model_validate_json(resp.text)
                    modelo_ativo = mod
                    todos_eventos.extend(res_lote.eventos)
                    print(f"   Lote págs {lote_pags}: +{len(res_lote.eventos)} shows extraídos")
                    break
                except Exception as e:
                    if modelo_ativo:
                        modelo_ativo = None
                    continue
        finally:
            try:
                client.files.delete(name=arq_remoto.name)
            except Exception:
                pass
            if os.path.exists(caminho_lote):
                os.remove(caminho_lote)

    doc_orig.close()
    return todos_eventos


def extrair_eventos_portal_web(mes_idx: int) -> List[EventoMusical]:
    """
    Usa Playwright (Chromium Headless) para abrir a página dinâmica do Sesc SP
    (/programacao/?tipos=shows-espetaculos-e-performances) e capturar shows extras não listados no PDF.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return []

    print("🌐 Consultando programação online em sescsp.org.br/programacao/ via Playwright...")
    texto_portal = ""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            url_portal = "https://www.sescsp.org.br/programacao/?tipos=shows-espetaculos-e-performances&linguagens=musica"
            page.goto(url_portal, timeout=45000, wait_until="networkidle")
            for _ in range(6):
                page.mouse.wheel(0, 4000)
                page.wait_for_timeout(1500)
            texto_portal = page.inner_text("body")
            browser.close()
    except Exception as e:
        print(f"⚠️ Aviso ao consultar portal web: {e}")
        return []

    if len(texto_portal.strip()) < 200:
        return []

    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)
    for mod in obter_fila_modelos_gemini(client):
        try:
            resp = client.models.generate_content(
                model=mod,
                contents=[
                    f"Extraia todos os shows de MÚSICA nas 18 unidades da CAPITAL do mês {mes_idx:02d} presentes no texto abaixo:\n\n{texto_portal[:60000]}"
                ],
                config=types.GenerateContentConfig(
                    temperature=0.0,
                    response_mime_type="application/json",
                    response_schema=ListaProgramacao,
                ),
            )
            res = ListaProgramacao.model_validate_json(resp.text)
            print(f"✅ Portal Web adicionou {len(res.eventos)} registros para mesclagem!")
            return res.eventos
        except Exception:
            continue
    return []


# (Mantenha a sua função salvar_planilha_excel exatamente igual e atualize o main abaixo)
