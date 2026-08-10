"""
sesc_musica_capital.py
======================
Extrai programação de MÚSICA das unidades da CAPITAL do Sesc SP
e salva como planilha Excel (.xlsx).

Dependências:
    pip install requests pdfplumber openpyxl
"""

import re, sys, io, os
from datetime import datetime
from pathlib import Path

def _req(pkg, install=None):
    try: __import__(pkg)
    except ImportError:
        print(f"❌ '{pkg}' não encontrado.\n   pip install {install or pkg}"); sys.exit(1)

_req("requests"); _req("pdfplumber"); _req("openpyxl")

import requests, pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── CONFIGURAÇÕES ───────────────────────────────────────────────────────────

MESES_PT = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
            7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}
MESES_ABR = {1:"JAN",2:"FEV",3:"MAR",4:"ABR",5:"MAI",6:"JUN",
             7:"JUL",8:"AGO",9:"SET",10:"OUT",11:"NOV",12:"DEZ"}

UNIDADES_CAP = {
    "14 BIS","24 DE MAIO","AVENIDA PAULISTA","BELENZINHO","BOM RETIRO",
    "CAMPO LIMPO","CASA VERDE","CONSOLAÇÃO","INTERLAGOS","IPIRANGA",
    "ITAQUERA","PINHEIROS","POMPEIA","SANTANA","SANTO AMARO","SÃO BENTO",
    "SÃO JOÃO CLÍMACO","VILA MARIANA","CARMO","PARQUE DOM PEDRO","BRÁS","BRAS",
}

CORTE = {"TEATRO","DANÇA","CIRCO","EXPOSIÇÕES","CINEMA","CRIANÇAS","JOVENS",
         "PESSOAS IDOSAS","ESPORTE","TURISMO SOCIAL","TRABALHO E","MEIO AMBIENTE",
         "DIREITOS HUMANOS","TECNOLOGIAS","SAÚDE","ALIMENTAÇÃO","LITERATURA",
         "CENTRO DE PESQUISA","LOJAS SESC","INGRESSOS"}

PREP = {"a","ao","à","às","de","do","da","dos","das","em","no","na","nos","nas",
        "por","para","pelo","pela","pelos","pelas","com","entre","e","ou","que",
        "se","sua","seu","seus","suas","um","uma","os","as","o"}

PREF_DESC = ("part.","com ","curadoria","mediação","apresentação de",
             "realização","intervenção","produção","direção","texto ",
             "arranjo","criação ","foto:","#","consulte","local:","informações",
             "integra o projeto","integra a ação","transmissão","ingressos")

DIAS_SEM = r'(?:Segunda|Terça|Quarta|Quinta|Sexta|Sábado|Domingo|Seg\.|Ter\.|Qua\.|Qui\.|Sex\.|Sáb\.|Dom\.)'

HEADERS = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0",
           "Referer":"https://www.sescsp.org.br/"}

COR_V = "C00000"; COR_VE = "8B0000"; BRANCO = "FFFFFF"; CINZA = "F5F5F5"
COLS = ["Artista / Título","Unidade","Data","Horário","Preço","Classificação"]
LARG = [42,24,26,12,22,16]

# ─── LOCALIZAR E BAIXAR PDF ──────────────────────────────────────────────────

# Regex para qualquer PDF no site do Sesc
_RE_PDF = re.compile(
    r'https://www\.sescsp\.org\.br/wp-content/uploads/'
    r'[^\s"\'<>]*[Ee][Mm][- _][Cc][Aa][Rr][Tt][Aa][Zz][^\s"\'<>]*\.pdf', re.I)

def _ok(url):
    try:
        r = requests.head(url, headers=HEADERS, timeout=10, allow_redirects=True)
        return r.status_code==200 and "pdf" in r.headers.get("Content-Type","").lower()
    except: return False

def _match_mes(url, ano, mes):
    ul=url.lower(); m=MESES_PT[mes]; ma=MESES_ABR[mes].lower()
    return str(ano) in url and (m in ul or ma in ul or m[:3] in ul)

def buscar_html(ano, mes):
    paginas=["https://www.sescsp.org.br/editorial/emcartaz/","https://www.sescsp.org.br/programacao/","https://www.sescsp.org.br/","https://www.sescsp.org.br/wp-sitemap.xml"]
    for base in paginas:
        try:
            r=requests.get(base,headers=HEADERS,timeout=20)
            if r.status_code!=200: continue
            todos=_RE_PDF.findall(r.text)
            em_c=[u for u in todos if re.search(r'[Ee]m.?[Cc]artaz|EM.?CARTAZ',u)]
            print(f"  📄 {len(em_c)} PDF(s) Em Cartaz em {base.split('sescsp.org.br/')[-1]}")
            for u in em_c:
                if _match_mes(u,ano,mes):
                    print(f"  ✅ Match HTML: {u.split('uploads/')[-1]}"); return u
        except Exception as e: print(f"  ⚠️  {base}: {e}")
    return None

def candidatos(ano, mes):
    m=MESES_PT[mes]; mc=m.capitalize(); ma=MESES_ABR[mes]
    mp=mes-1 if mes>1 else 12; ap=ano if mes>1 else ano-1
    return [
        # Padrão ago/2026: ano/mes-anterior/Em-Cartaz_ano_mes.pdf
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mp:02d}/Em-Cartaz_{ano}_{m}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mp:02d}/Em-Cartaz_{ano}_{m}_compressed.pdf",
        # Padrão com hífen e maiúscula: ago/2024, ago/2023
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/Em-Cartaz-{mc}-{ano}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mp:02d}/Em-Cartaz-{mc}-{ano}.pdf",
        # Padrão V11 jun/2026
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/V11_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/V12_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/V10_Em-Cartaz_{ano}_{m}_15-bx.pdf",
        # EM-CARTAZ maiúsculo
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/V11_EM-CARTAZ_{ma}_{ano}_bx.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/V9_EM-CARTAZ_{ma}_{ano}.pdf",
        # Pasta 2025/01 (jul/2026, ago/2025)
        f"https://www.sescsp.org.br/wp-content/uploads/2025/01/Em-Cartaz_{ano}_{m}_compressed.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/2025/01/Em-Cartaz-{mc}-{ano}.pdf",
        # AF_ ago/2025
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/AF_EM-CARTAZ_{ma}{ano}_BaixaDuplas-1.pdf",
        # Ano corrente pasta mes-anterior (variações extras)
        f"https://www.sescsp.org.br/wp-content/uploads/{ano}/{mp:02d}/Em-Cartaz-{ano}-{m}.pdf",
        f"https://www.sescsp.org.br/wp-content/uploads/{ap}/{mp:02d}/Em-Cartaz_{ano}_{m}.pdf",
    ]

def descobrir_pdf(ano, mes):
    print("  🌐 Raspando HTML do site...")
    u=buscar_html(ano,mes)
    if u: return u
    print("  🔎 Testando URLs candidatas...")
    for u in candidatos(ano,mes):
        print(f"    {u.split('uploads/')[-1]}")
        if _ok(u): print(f"  ✅ {u.split('uploads/')[-1]}"); return u
    return None


def baixar(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=45)
        return r.content if r.status_code==200 else None
    except Exception as e: print(f"  ❌ {e}"); return None

# ─── EXTRAIR TEXTO DO PDF ────────────────────────────────────────────────────

def _fix_espelho(t):
    linhas=[]
    for l in t.split("\n"):
        ws=[]
        for p in l.split():
            r=p[::-1]
            if len(p)>=3 and p[0].islower() and r[0].isupper() and r[-1].islower() and p.isalpha(): ws.append(r)
            else: ws.append(p)
        linhas.append(" ".join(ws))
    return "\n".join(linhas)

def extrair(conteudo):
    """
    Extrai texto preservando a ordem de leitura natural.
    Agrupa palavras por linha (coordenada Y) e ordena por X dentro de cada linha.
    NÃO divide em colunas — deixa o parser lidar com o conteúdo de seções.
    """
    partes=[]
    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for page in pdf.pages:
            ws = page.extract_words(x_tolerance=3, y_tolerance=3,
                                    keep_blank_chars=False, use_text_flow=False)
            if not ws: continue
            linhas={}
            for w in ws:
                y=round(float(w["top"])/3)*3
                linhas.setdefault(y,[]).append((float(w["x0"]),w["text"]))
            for y in sorted(linhas):
                l=" ".join(t for _,t in sorted(linhas[y]))
                if l.strip(): partes.append(l)
            partes.append("")
    return _fix_espelho("\n".join(partes))

def ler_local(p):
    with open(p,"rb") as f: return extrair(f.read())

# ─── PARSER ──────────────────────────────────────────────────────────────────

def _n(t): return re.sub(r"\s+"," ",t).strip()

def _unidade(ls):
    lu=re.sub(r'\s*\(.*?\)','',ls.upper()).strip()
    if not (3<=len(lu)<=55): return False
    if lu!=lu.upper() or re.search(r'\d',lu): return False
    return any(u==lu or (lu.startswith(u) and len(lu)>len(u) and lu[len(u)] in' (')
               for u in UNIDADES_CAP)

def _corte(ls):
    lu=ls.upper().strip()
    return any(lu==s or lu.startswith(s+" ") for s in CORTE)

def _meta_solta(ls):
    """Linha que é continuação de metadado: R$, hora, dia-da-semana, classif."""
    if not ls: return False
    if re.match(r'^R\$\s*\d',ls): return True
    if re.match(r'^\d{1,2}h',ls): return True
    if re.match(DIAS_SEM,ls,re.I): return True
    if re.match(r'^(AL|A\d{1,2}|Livre|Todas as|GRÁTIS)$',ls,re.I): return True
    return False

def _data(ls):
    m=re.match(
        r'^(?:Dia[s]?\s+|De\s+)?'
        r'(\d{1,2}(?:\s*(?:e|a|,)\s*\d{1,2})?'
        r'(?:\s*/\s*\d{1,2})?/\d{1,2}'
        r'(?:\s*(?:e|a|,)\s*\d{1,2}/\d{1,2})?)'
        r'([^.]*)\.\s*(.*)',ls,re.I)
    if not m: return None
    data=_n(m.group(1)+m.group(2)); resto=_n(m.group(3))
    hm=re.search(r'(\d{1,2}h\d{0,2}(?:\s*às\s*\d{1,2}h\d{0,2})?)',resto,re.I)
    pm=re.findall(r'R\$\s*\d+',resto)
    gm=re.search(r'GRÁTIS',resto,re.I)
    cm=re.search(r'\b(A(?:L|\d{1,2})|Livre|Todas as idades)\b',resto,re.I)
    return {"Data":data,"Horário":hm.group(1) if hm else "",
            "Preço":" / ".join(pm) if pm else ("GRÁTIS" if gm else ""),
            "Classificação":cm.group(1) if cm else ""}

def _titulo_ok(ls):
    if not ls or len(ls)<3: return False
    if ls.upper() in {"SHOWS E ESPETÁCULOS","INGRESSOS","EM CARTAZ","PROGRAMAÇÃO",
                      "EVENTOS","MÚSICA","MUSICA","SHOWS","ESPETÁCULOS","DESTAQUES DA PROGRAMAÇÃO"}: return False
    if ls[0].islower(): return False
    prim=ls.split()[0].lower().rstrip('.,;:()')
    if prim in PREP: return False
    ll=ls.lower()
    if any(ll.startswith(p) for p in PREF_DESC): return False
    if re.match(r'^Dia[s]?\s+\d|^De\s+\d|\d{1,2}/\d{1,2}',ls,re.I): return False
    if _meta_solta(ls): return False
    if re.match(r'^\d+\s+EM CARTAZ',ls,re.I): return False
    return True

def _limpa(t):
    t=re.sub(r'\bFoto:.*','',t,flags=re.I)
    t=re.sub(r'\s*\|.*$','',t)
    ws=t.split(); meio=len(ws)//2
    if meio>2 and ws[:meio]==ws[meio:]: t=" ".join(ws[:meio])
    return _n(t)[:90]

def _serie(ls):
    """Detecta nome de série/projeto em maiúsculas (ex: 'RUA DO ROCK', 'SALVE SAMBA')."""
    if not ls: return False
    lu=ls.upper().strip()
    if lu==ls.strip() and not re.search(r'\d',lu) and 5<len(lu)<60: return True
    return False

def parse(texto):
    idx=-1
    for pat in [r'\bMÚSICA\b',r'\bMUSICA\b']:
        for m in re.finditer(pat,texto,re.I):
            if idx==-1 or m.start()<idx: idx=m.start()
    if idx==-1:
        print("  ⚠️  'MÚSICA' não encontrada.")
        print(texto[:600]); return []

    print(f"  ✅ Seção MÚSICA na posição {idx}")

    eventos=[]; unid=None; ev=None; tit=""; desc=False
    serie_preco=""; serie_classif=""  # herança de série/projeto

    def fecha():
        nonlocal ev,tit,desc
        if ev is not None and tit and unid:
            ev["Artista / Título"]=_limpa(tit)
            ev["Unidade"]=f"Sesc {unid.title()}"
            eventos.append(ev)
        ev=None; tit=""; desc=False

    for lr in texto[idx:].split("\n"):
        ls=_n(lr)
        if not ls: desc=False; continue

        if _corte(ls): unid=None; serie_preco=""; serie_classif=""; continue
        if re.match(r'^(MÚSICA|MUSICA)$',ls,re.I): continue

        if _unidade(ls):
            fecha(); unid=ls; serie_preco=""; serie_classif=""; continue

        if not unid: continue

        # Linha de metadado solto (preço/hora/classif após data)
        if _meta_solta(ls):
            if ev is not None:
                if not ev.get("Horário"):
                    hm=re.search(r'(\d{1,2}h\d{0,2})',ls,re.I)
                    if hm: ev["Horário"]=hm.group(1)
                if not ev.get("Preço"):
                    pm=re.findall(r'R\$\s*\d+',ls)
                    if pm: ev["Preço"]=" / ".join(pm)
                    elif re.search(r'GRÁTIS',ls,re.I): ev["Preço"]="GRÁTIS"
                if not ev.get("Classificação"):
                    cm=re.search(r'\b(A(?:L|\d{1,2})|Livre|Todas as idades)\b',ls,re.I)
                    if cm: ev["Classificação"]=cm.group(1)
            elif ls.startswith("R$") or re.match(r'^(AL|A\d{1,2})$',ls,re.I):
                # Metadado de série (sem evento aberto ainda)
                pm=re.findall(r'R\$\s*\d+',ls)
                if pm: serie_preco=" / ".join(pm)
                cm=re.search(r'\b(A(?:L|\d{1,2})|Livre)\b',ls,re.I)
                if cm: serie_classif=cm.group(1)
            continue

        # Linha de data
        ev_d=_data(ls)
        if ev_d:
            if ev is not None: fecha()
            ev=dict(ev_d)
            # Herda da série se campos vazios
            if not ev["Preço"] and serie_preco: ev["Preço"]=serie_preco
            if not ev["Classificação"] and serie_classif: ev["Classificação"]=serie_classif
            desc=False; continue

        # Título ou descrição
        if not desc:
            if _titulo_ok(ls):
                if ev is not None: fecha()
                tit=_limpa(ls); desc=True
            elif _serie(ls):
                # Nome de série em maiúsculas: reseta herança
                serie_preco=""; serie_classif=""

    fecha()

    print(f"  Unidades vistas: {sorted(set(e['Unidade'] for e in eventos))}")

    # Deduplicação
    visto=set(); unicos=[]
    for e in eventos:
        ch=(e.get("Artista / Título",""),e.get("Unidade",""),e.get("Data",""))
        if ch not in visto: visto.add(ch); unicos.append(e)
    return unicos

# ─── GERAR EXCEL ─────────────────────────────────────────────────────────────

def excel(evs, path, mes_ref):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Programação Musical"; nc=len(COLS)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=nc)
    c=ws.cell(row=1,column=1); c.value=f"Programação Musical – Sesc SP Capital | {mes_ref}"
    c.font=Font(name="Arial",bold=True,size=13,color=BRANCO)
    c.fill=PatternFill("solid",fgColor=COR_V); c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=28
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=nc)
    c=ws.cell(row=2,column=1)
    c.value=(f"Extraído em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
             "   |   Fonte: PDF Em Cartaz – sescsp.org.br")
    c.font=Font(name="Arial",size=9,italic=True,color="555555")
    c.alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[2].height=16
    for ci,n in enumerate(COLS,1):
        c=ws.cell(row=3,column=ci,value=n)
        c.font=Font(name="Arial",bold=True,size=10,color=BRANCO)
        c.fill=PatternFill("solid",fgColor=COR_VE)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=Border(bottom=Side(style="thin",color=BRANCO))
    ws.row_dimensions[3].height=22
    brd=Border(bottom=Side(style="thin",color="DDDDDD"),right=Side(style="thin",color="DDDDDD"))
    for i,ev in enumerate(evs):
        row=i+4; fill=PatternFill("solid",fgColor=CINZA if i%2==0 else BRANCO)
        for ci,col in enumerate(COLS,1):
            c=ws.cell(row=row,column=ci,value=ev.get(col,""))
            c.font=Font(name="Arial",size=10); c.fill=fill
            c.alignment=Alignment(vertical="center",wrap_text=(ci==1)); c.border=brd
        ws.row_dimensions[row].height=18
    lr=len(evs)+4; ws.merge_cells(start_row=lr,start_column=1,end_row=lr,end_column=nc)
    c=ws.cell(row=lr,column=1); c.value=f"Total de eventos musicais: {len(evs)}"
    c.font=Font(name="Arial",bold=True,size=10,color=BRANCO)
    c.fill=PatternFill("solid",fgColor=COR_V); c.alignment=Alignment(horizontal="right",vertical="center")
    ws.row_dimensions[lr].height=20
    for ci,larg in enumerate(LARG,1): ws.column_dimensions[get_column_letter(ci)].width=larg
    ws.auto_filter.ref=f"A3:{get_column_letter(nc)}3"; ws.freeze_panes="A4"; wb.save(path)

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("="*58)
    print("  Sesc SP – Programação Musical da Capital")
    print("  Extrator via PDF 'Em Cartaz'")
    print("="*58)
    agora=datetime.now(); ano,mes=agora.year,agora.month
    mes_ref=f"{MESES_PT[mes].capitalize()} {ano}"
    print(f"\n📅 {mes_ref}")
    print("\n🔍 Localizando PDF...")
    url=descobrir_pdf(ano,mes); texto=""
    if url:
        print(f"\n📥 Baixando: {url.split('uploads/')[-1]}")
        c=baixar(url)
        if c:
            print("📄 Extraindo texto...")
            texto=extrair(c); print(f"   {len(texto):,} chars")
    if not texto.strip():
        env=os.environ.get("PDF_LOCAL_PATH","").strip().strip('"')
        if env and Path(env).exists(): print(f"📄 PDF_LOCAL_PATH: {env}"); texto=ler_local(env)
    if not texto.strip():
        if sys.stdin.isatty():
            print("\n⚠️  PDF não obtido automaticamente.")
            cp=input("   Caminho do PDF: ").strip().strip('"')
            if cp and Path(cp).exists(): texto=ler_local(cp)
            else: print("❌ Inválido."); sys.exit(1)
        else:
            print(f"\n❌ PDF não encontrado (CI). Mês: {mes_ref}")
            print("   PDF pode ainda não ter sido publicado pelo Sesc."); sys.exit(1)
    if not texto.strip(): print("❌ Texto vazio."); sys.exit(1)
    idx=max(texto.find("MÚSICA"),texto.find("MUSICA"),0)
    print(f"\n📋 Contexto MÚSICA (chars {idx}–{idx+300}):")
    print(texto[idx:idx+300]); print("...")
    print("\n🎵 Parseando...")
    evs=parse(texto)
    if not evs:
        print("\n⚠️  Nenhum evento. Texto extraído (2000 chars):"); print(texto[:2000]); sys.exit(1)
    evs.sort(key=lambda x:(x.get("Unidade",""),x.get("Data","")))
    nome=f"sesc_musica_capital_{ano}_{mes:02d}.xlsx"
    dest=Path(__file__).parent/nome
    print(f"\n📊 Gerando {nome}..."); excel(evs,dest,mes_ref)
    print(f"\n✅ Concluído! {len(evs)} eventos → {dest}")

if __name__=="__main__": main()
