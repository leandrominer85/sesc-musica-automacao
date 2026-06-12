"""
comparar_e_notificar.py
=======================
Compara a planilha gerada com a versão anterior (salva no repositório)
e envia e-mail de notificação se houver mudanças.

Variáveis de ambiente necessárias (configuradas nos Secrets do GitHub):
  GMAIL_USER     → seu endereço Gmail (ex: voce@gmail.com)
  GMAIL_APP_PASS → senha de app do Gmail (não é a senha normal)
  EMAIL_DESTINO  → e-mail(s) que receberão a notificação (separe por vírgula)
"""

import os
import sys
import json
import hashlib
import smtplib
import email.mime.multipart as mime_multi
import email.mime.text as mime_text
import email.mime.base as mime_base
import email.encoders as encoders
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("❌ openpyxl não encontrado.")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

MESES_PT = {
    1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
    7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro",
}

SCRIPT_DIR   = Path(__file__).parent
PROJETO_DIR  = SCRIPT_DIR.parent
HISTORICO_DIR = PROJETO_DIR / "historico"
HASH_FILE    = PROJETO_DIR / "ultimo_hash.json"


# ─────────────────────────────────────────────────────────────────────────────
# COMPARAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def extrair_dados_excel(caminho: Path) -> list[dict]:
    """Extrai todas as linhas de dados do Excel (ignora cabeçalho/rodapé)."""
    wb = openpyxl.load_workbook(caminho)
    ws = wb.active
    dados = []
    cabecalho = None
    for row in ws.iter_rows(values_only=True):
        if cabecalho is None:
            # Linha 3 é o cabeçalho real
            if row[0] and str(row[0]).startswith("Artista"):
                cabecalho = list(row)
            continue
        if row[0] and not str(row[0]).startswith("Total"):
            dados.append(dict(zip(cabecalho, row)))
    return dados


def calcular_hash(dados: list[dict]) -> str:
    """Hash MD5 do conteúdo de dados para detectar mudanças."""
    conteudo = json.dumps(dados, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.md5(conteudo.encode()).hexdigest()


def carregar_hash_anterior() -> str | None:
    if HASH_FILE.exists():
        try:
            return json.loads(HASH_FILE.read_text())["hash"]
        except Exception:
            pass
    return None


def salvar_hash(hash_val: str) -> None:
    HASH_FILE.write_text(json.dumps({"hash": hash_val, "atualizado": datetime.now().isoformat()}))


def diff_eventos(antigos: list[dict], novos: list[dict]) -> dict:
    """Compara duas listas de eventos e retorna adicionados/removidos."""
    def chave(ev):
        return (
            str(ev.get("Artista / Título", "")),
            str(ev.get("Unidade", "")),
            str(ev.get("Data", "")),
        )

    set_antigos = {chave(e): e for e in antigos}
    set_novos   = {chave(e): e for e in novos}

    adicionados = [e for k, e in set_novos.items()   if k not in set_antigos]
    removidos   = [e for k, e in set_antigos.items() if k not in set_novos]

    return {"adicionados": adicionados, "removidos": removidos}


# ─────────────────────────────────────────────────────────────────────────────
# E-MAIL
# ─────────────────────────────────────────────────────────────────────────────

def formatar_tabela_html(eventos: list[dict], cor_header: str) -> str:
    if not eventos:
        return "<p><em>Nenhum evento nesta categoria.</em></p>"

    colunas = ["Artista / Título", "Unidade", "Data", "Horário", "Preço"]
    linhas  = []
    linhas.append(
        f'<table border="0" cellspacing="0" cellpadding="6" '
        f'style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px;">'
    )
    linhas.append("<thead><tr>")
    for col in colunas:
        linhas.append(
            f'<th style="background:{cor_header};color:#fff;text-align:left;'
            f'padding:8px 10px;border-bottom:2px solid #fff;">{col}</th>'
        )
    linhas.append("</tr></thead><tbody>")

    for i, ev in enumerate(eventos):
        bg = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        linhas.append(f'<tr style="background:{bg};">')
        for col in colunas:
            valor = ev.get(col, "") or ""
            linhas.append(f'<td style="padding:7px 10px;border-bottom:1px solid #eee;">{valor}</td>')
        linhas.append("</tr>")

    linhas.append("</tbody></table>")
    return "\n".join(linhas)


def montar_email_html(
    mes_ref: str,
    total_eventos: int,
    diff: dict,
    tem_mudancas: bool,
) -> str:
    agora = datetime.now().strftime("%d/%m/%Y às %H:%M")

    if tem_mudancas:
        status_badge = (
            '<span style="background:#e74c3c;color:#fff;padding:4px 12px;'
            'border-radius:20px;font-size:12px;font-weight:bold;">🔔 COM ALTERAÇÕES</span>'
        )
        resumo = (
            f'<p style="font-size:15px;">A programação musical de <strong>{mes_ref}</strong> '
            f'foi atualizada. Foram identificadas <strong>{len(diff["adicionados"])} inclusões</strong> '
            f'e <strong>{len(diff["removidos"])} remoções</strong> em relação ao mês anterior.</p>'
        )
    else:
        status_badge = (
            '<span style="background:#27ae60;color:#fff;padding:4px 12px;'
            'border-radius:20px;font-size:12px;font-weight:bold;">✅ SEM ALTERAÇÕES</span>'
        )
        resumo = (
            f'<p style="font-size:15px;">A programação musical de <strong>{mes_ref}</strong> '
            f'foi verificada. Nenhuma alteração em relação à versão anterior.</p>'
        )

    adds_html = formatar_tabela_html(diff["adicionados"], "#27ae60") if diff["adicionados"] else ""
    rems_html = formatar_tabela_html(diff["removidos"],   "#e74c3c") if diff["removidos"]   else ""

    secao_adds = ""
    if diff["adicionados"]:
        secao_adds = f"""
        <h3 style="color:#27ae60;margin-top:28px;">
            ➕ Eventos adicionados ({len(diff["adicionados"])})
        </h3>
        {adds_html}
        """

    secao_rems = ""
    if diff["removidos"]:
        secao_rems = f"""
        <h3 style="color:#e74c3c;margin-top:28px;">
            ➖ Eventos removidos ({len(diff["removidos"])})
        </h3>
        {rems_html}
        """

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f0f0;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f0f0;padding:30px 0;">
<tr><td align="center">
<table width="620" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.12);">

  <!-- CABEÇALHO -->
  <tr>
    <td style="background:#C00000;padding:24px 32px;">
      <p style="margin:0;color:#fff;font-size:11px;letter-spacing:2px;text-transform:uppercase;">Sesc SP</p>
      <h1 style="margin:6px 0 0;color:#fff;font-size:22px;font-weight:700;">
        Programação Musical – Capital
      </h1>
      <p style="margin:6px 0 0;color:rgba(255,255,255,.75);font-size:13px;">{mes_ref}</p>
    </td>
  </tr>

  <!-- CORPO -->
  <tr>
    <td style="padding:28px 32px;">
      <div style="margin-bottom:16px;">{status_badge}</div>
      {resumo}

      <table style="width:100%;background:#f8f8f8;border-radius:6px;padding:16px;margin:20px 0;" cellpadding="0" cellspacing="0">
        <tr>
          <td style="text-align:center;padding:0 20px;">
            <p style="margin:0;font-size:28px;font-weight:700;color:#C00000;">{total_eventos}</p>
            <p style="margin:4px 0 0;font-size:12px;color:#666;">eventos no total</p>
          </td>
          <td style="text-align:center;padding:0 20px;border-left:1px solid #ddd;">
            <p style="margin:0;font-size:28px;font-weight:700;color:#27ae60;">{len(diff["adicionados"])}</p>
            <p style="margin:4px 0 0;font-size:12px;color:#666;">adicionados</p>
          </td>
          <td style="text-align:center;padding:0 20px;border-left:1px solid #ddd;">
            <p style="margin:0;font-size:28px;font-weight:700;color:#e74c3c;">{len(diff["removidos"])}</p>
            <p style="margin:4px 0 0;font-size:12px;color:#666;">removidos</p>
          </td>
        </tr>
      </table>

      <p style="font-size:13px;color:#666;">
        A planilha completa está em anexo neste e-mail.
      </p>

      {secao_adds}
      {secao_rems}
    </td>
  </tr>

  <!-- RODAPÉ -->
  <tr>
    <td style="background:#f5f5f5;padding:16px 32px;border-top:1px solid #eee;">
      <p style="margin:0;font-size:11px;color:#999;">
        Verificação automática executada em {agora} via GitHub Actions.<br>
        Fonte: <a href="https://www.sescsp.org.br/programacao/" style="color:#C00000;">sescsp.org.br/programacao</a>
      </p>
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


def enviar_email(
    assunto: str,
    corpo_html: str,
    anexo_path: Path | None = None,
) -> bool:
    gmail_user  = os.environ.get("GMAIL_USER", "")
    gmail_pass  = os.environ.get("GMAIL_APP_PASS", "")
    destinarios_raw = os.environ.get("EMAIL_DESTINO", gmail_user)

    if not gmail_user or not gmail_pass:
        print("⚠️  GMAIL_USER ou GMAIL_APP_PASS não configurados. E-mail não enviado.")
        return False

    destinatarios = [d.strip() for d in destinarios_raw.split(",") if d.strip()]

    msg = mime_multi.MIMEMultipart("mixed")
    msg["Subject"] = assunto
    msg["From"]    = f"Sesc Automação <{gmail_user}>"
    msg["To"]      = ", ".join(destinatarios)

    # Corpo HTML
    parte_html = mime_text.MIMEText(corpo_html, "html", "utf-8")
    msg.attach(parte_html)

    # Anexo Excel
    if anexo_path and anexo_path.exists():
        with open(anexo_path, "rb") as f:
            parte_anexo = mime_base.MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            parte_anexo.set_payload(f.read())
        encoders.encode_base64(parte_anexo)
        parte_anexo.add_header(
            "Content-Disposition",
            f'attachment; filename="{anexo_path.name}"',
        )
        msg.attach(parte_anexo)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_pass)
            smtp.sendmail(gmail_user, destinatarios, msg.as_bytes())
        print(f"✅ E-mail enviado para: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    agora = datetime.now()
    mes_ref = f"{MESES_PT[agora.month]} {agora.year}"

    # Localiza o Excel recém-gerado
    padrao = f"sesc_musica_capital_{agora.year}_{agora.month:02d}.xlsx"
    xlsx_path = PROJETO_DIR / padrao
    if not xlsx_path.exists():
        # Procura na pasta scripts também
        xlsx_path = SCRIPT_DIR.parent / padrao
    if not xlsx_path.exists():
        print(f"❌ Arquivo não encontrado: {padrao}")
        sys.exit(1)

    print(f"📊 Lendo planilha: {xlsx_path.name}")
    dados_novos = extrair_dados_excel(xlsx_path)
    hash_novo   = calcular_hash(dados_novos)
    hash_antigo = carregar_hash_anterior()

    print(f"   Hash atual:    {hash_novo[:12]}...")
    print(f"   Hash anterior: {hash_antigo[:12] if hash_antigo else 'nenhum'}...")

    tem_mudancas = (hash_antigo is None) or (hash_novo != hash_antigo)

    # Carrega dados anteriores para diff (se existirem)
    dados_antigos = []
    if HISTORICO_DIR.exists():
        arquivos_anteriores = sorted(HISTORICO_DIR.glob("*.xlsx"))
        if arquivos_anteriores:
            try:
                dados_antigos = extrair_dados_excel(arquivos_anteriores[-1])
            except Exception:
                pass

    diff = diff_eventos(dados_antigos, dados_novos)

    # Se não há mudanças e não é primeira execução, ainda envia e-mail mensal
    # (comentar as linhas abaixo para enviar só quando houver mudanças)
    forcar_envio = os.environ.get("FORCAR_ENVIO", "false").lower() == "true"

    if tem_mudancas or forcar_envio:
        status = "🔔 COM ALTERAÇÕES" if tem_mudancas else "✅ Programação do mês"
        assunto = f"[Sesc SP] {status} – Música Capital | {mes_ref}"
        corpo = montar_email_html(mes_ref, len(dados_novos), diff, tem_mudancas)
        enviado = enviar_email(assunto, corpo, xlsx_path)
        if not enviado:
            print("⚠️  E-mail não pôde ser enviado (verifique os Secrets do GitHub).")
    else:
        print("ℹ️  Sem mudanças detectadas. E-mail não enviado.")

    # Salva hash e copia para histórico
    salvar_hash(hash_novo)
    HISTORICO_DIR.mkdir(exist_ok=True)
    import shutil
    shutil.copy2(xlsx_path, HISTORICO_DIR / xlsx_path.name)
    print(f"💾 Histórico salvo: historico/{xlsx_path.name}")


if __name__ == "__main__":
    main()
