# 🎵 Sesc SP – Automação da Programação Musical

Extrai mensalmente a programação de **Música** das unidades do **Sesc SP na capital**,
gera uma planilha Excel e envia notificação por **e-mail** quando há mudanças.

Roda automaticamente todo dia **25 de cada mês** via **GitHub Actions** (gratuito).

---

## 📁 Estrutura do projeto

```
sesc-automation/
│
├── .github/
│   └── workflows/
│       └── sesc_musica.yml       ← Automação do GitHub Actions
│
├── scripts/
│   ├── sesc_musica_capital.py    ← Extrator principal (baixa PDF, gera Excel)
│   └── comparar_e_notificar.py  ← Detecta mudanças e envia e-mail
│
├── historico/                    ← Versões anteriores (criado automaticamente)
│   └── sesc_musica_capital_2026_06.xlsx
│
├── ultimo_hash.json              ← Controle de versão interno (criado automaticamente)
├── sesc_musica_capital_2026_07.xlsx  ← Planilha mais recente
└── README.md
```

---

## 🚀 Configuração inicial (faça uma vez)

### Passo 1 – Criar o repositório no GitHub

1. Acesse [github.com](https://github.com) e faça login (crie uma conta se não tiver)
2. Clique em **"New repository"** (botão verde no canto superior direito)
3. Configure:
   - **Nome:** `sesc-musica-automacao` (ou qualquer nome)
   - **Visibilidade:** ✅ **Private** (seus e-mails ficarão protegidos)
   - **Add a README file:** deixe desmarcado
4. Clique em **"Create repository"**

### Passo 2 – Fazer upload dos arquivos

No repositório criado, clique em **"uploading an existing file"** e envie:

```
.github/workflows/sesc_musica.yml
scripts/sesc_musica_capital.py
scripts/comparar_e_notificar.py
```

> **Atenção:** mantenha a estrutura de pastas exatamente como está.

Alternativamente, use Git:
```bash
git clone https://github.com/SEU_USUARIO/sesc-musica-automacao.git
# Copie os arquivos e então:
git add .
git commit -m "Setup inicial"
git push
```

### Passo 3 – Criar senha de app no Gmail

O script usa uma **senha de app** (não sua senha normal) para enviar e-mails.

1. Acesse [myaccount.google.com/security](https://myaccount.google.com/security)
2. Ative a **Verificação em duas etapas** (obrigatório)
3. Volte à segurança e procure **"Senhas de app"**
4. Clique em **"Selecionar app"** → Outro → Digite: `Sesc Automação`
5. Clique em **"Gerar"**
6. Copie a senha de 16 caracteres exibida (ex: `abcd efgh ijkl mnop`)
   > Guarde essa senha agora — ela não é exibida novamente!

### Passo 4 – Configurar os Secrets no GitHub

Os dados sensíveis ficam protegidos como "Secrets":

1. No repositório, clique em **Settings** (aba superior)
2. No menu lateral: **Secrets and variables → Actions**
3. Clique em **"New repository secret"** e adicione os 3 secrets:

| Nome           | Valor                                      |
|----------------|--------------------------------------------|
| `GMAIL_USER`   | Seu endereço Gmail (ex: `voce@gmail.com`)  |
| `GMAIL_APP_PASS` | A senha de app gerada no Passo 3 (sem espaços) |
| `EMAIL_DESTINO` | E-mail(s) que receberão a notificação (separe múltiplos por vírgula) |

### Passo 5 – Testar manualmente

1. No repositório, clique na aba **Actions**
2. No menu lateral, clique em **"🎵 Sesc SP – Programação Musical Mensal"**
3. Clique em **"Run workflow"** → **"Run workflow"** (botão verde)
4. Aguarde ~2 minutos e verifique o resultado
5. Cheque seu e-mail!

---

## 📅 Agendamento automático

O workflow roda automaticamente todo **dia 25 de cada mês às 9h** (horário de Brasília).

Isso garante que o PDF "Em Cartaz" do mês seguinte já esteja disponível
(o Sesc publica normalmente na última semana do mês anterior).

Para mudar o dia, edite `.github/workflows/sesc_musica.yml`:
```yaml
# Linha 9 — formato: "minuto hora dia mês dia-da-semana"
- cron: "0 12 25 * *"   # dia 25 às 12h UTC (= 9h Brasília)
- cron: "0 12 28 * *"   # dia 28 às 9h Brasília
- cron: "0 12 20 * *"   # dia 20 às 9h Brasília
```

---

## 📬 E-mails gerados

### Quando há mudanças na programação:
- Assunto: `[Sesc SP] 🔔 COM ALTERAÇÕES – Música Capital | Julho 2026`
- Mostra eventos adicionados ✅ e removidos ❌
- Planilha Excel completa em anexo

### Quando não há mudanças:
- Sem e-mail (comportamento padrão)
- Para receber e-mail mesmo sem mudanças, rode manualmente marcando **"Forçar envio"**

---

## 🔧 Executar localmente (opcional)

Se quiser rodar no seu computador sem esperar o agendamento:

```bash
# Instale as dependências (uma vez)
pip install requests pypdf openpyxl

# Rode o extrator
python scripts/sesc_musica_capital.py

# Para comparar e enviar e-mail, configure as variáveis:
export GMAIL_USER="voce@gmail.com"
export GMAIL_APP_PASS="abcdefghijklmnop"
export EMAIL_DESTINO="destino@email.com"
python scripts/comparar_e_notificar.py
```

---

## 🛠️ Resolução de problemas

### "Nenhum evento encontrado"
O formato do PDF "Em Cartaz" mudou. Abra `scripts/sesc_musica_capital.py`
e ajuste a função `parse_musica()`.

### E-mail não chegou
1. Verifique se os 3 Secrets estão configurados corretamente
2. Confirme que a verificação em duas etapas está ativa no Gmail
3. Veja os logs em: **Actions → último workflow → "Comparar mudanças e notificar"**

### Workflow não executa automaticamente
Repositórios sem atividade recente podem ter o agendamento pausado pelo GitHub.
Execute manualmente uma vez para reativar.

---

## 📊 Histórico

Cada execução salva a planilha em `historico/` e faz commit automaticamente.
Isso cria um histórico completo de todas as programações mensais diretamente no GitHub.

---

*Projeto gerado com assistência do Claude (Anthropic)*
