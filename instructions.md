# AutoMail — Instruções de Uso

## 1. Pré-requisitos

### Python 3.8 ou superior

Verificar se já está instalado:
```
python --version
```

Se não tiver, baixe em: https://www.python.org/downloads/

Durante a instalação, marque a opção **"Add Python to PATH"**.

---

### Dependências do projeto

O script usa apenas bibliotecas nativas do Python, portanto **não precisa instalar nada** para o modo IMAP básico.

As únicas dependências opcionais são:

| Biblioteca | Para que serve | Como instalar |
|---|---|---|
| `py7zr` | Extrair arquivos `.7z` automaticamente | `pip install py7zr` |
| `pypff` | Ler arquivos `.pst` (modo PST) | `pip install pypff` |

Para instalar as opcionais de uma vez:
```
pip install py7zr pypff
```

Se não instalar nenhuma, o script funciona normalmente — apenas avisa no terminal quando encontrar um `.7z` ou `.pst` e pula a extração.

---

### Verificar instalação

```
python -c "import imaplib, email, zipfile, json; print('OK')"
```

Se imprimir `OK`, está tudo certo para rodar.

---

## 2. Configurar o config.json

Abra o arquivo `config.json` e preencha os campos:

| Campo | O que preencher | Exemplo |
|---|---|---|
| `imap_server` | Endereço do servidor IMAP do seu e-mail | `mail.suaempresa.com.br` |
| `imap_port` | Porta IMAP — **993** para SSL (padrão), **143** para STARTTLS | `993` |
| `email_user` | Seu endereço de e-mail completo | `voce@suaempresa.com.br` |
| `email_pass` | Senha da conta de e-mail | `minha_senha` |
| `only_unseen` | `true` para buscar só não lidos, `false` para todos | `false` |
| `date_range.start` | Data inicial do período (formato DD/MM/AAAA) | `01/02/2026` |
| `date_range.end` | Data final do período (formato DD/MM/AAAA) | `28/02/2026` |
| `filter_emails` | Remetentes para baixar. Lista vazia `[]` aceita todos | `["fornecedor@email.com"]` |
| `output_base_path` | Pasta onde os arquivos serão salvos | `./auditoria_arquivos` |
| `auto_extract` | `true` para extrair `.zip` e `.7z` automaticamente | `true` |
| `allowed_extensions` | Extensões permitidas. Lista vazia `[]` baixa **tudo** | `[]` |
| `pst_file` | Caminho do arquivo `.pst` — deixe `""` para usar IMAP | `""` |

### Exemplo completo preenchido:
```json
{
  "imap_server": "mail.suaempresa.com.br",
  "imap_port": 993,
  "email_user": "voce@suaempresa.com.br",
  "email_pass": "sua_senha",
  "only_unseen": false,
  "date_range": { "start": "01/02/2026", "end": "28/02/2026" },
  "filter_emails": ["fornecedor@email.com", "contabilidade@parceiro.com.br"],
  "output_base_path": "./auditoria_arquivos",
  "auto_extract": true,
  "allowed_extensions": [],
  "pst_file": ""
}
```

### Filtrar só alguns tipos de arquivo:
```json
"allowed_extensions": [".xml", ".pdf", ".zip", ".xlsx", ".ofx"]
```
Deixar `[]` baixa qualquer extensão.

### Modo PST (arquivos históricos offline):
```json
"pst_file": "C:/Users/Sistemas/Desktop/backup.pst"
```
Quando `pst_file` estiver preenchido, o IMAP é ignorado completamente.

---

## 3. Como habilitar IMAP na conta de e-mail

### Servidor corporativo (cPanel, Zimbra, Postfix, etc.)
- Use as mesmas credenciais do webmail
- O servidor IMAP normalmente é `mail.suaempresa.com.br`
- Confirme com o administrador que a porta **993 (SSL)** está liberada

### Gmail
1. Acesse **Configurações → Ver todas as configurações → Encaminhamento e POP/IMAP**
2. Ative **Acesso IMAP**
3. Se usar autenticação em dois fatores, gere uma **Senha de App** em:
   `Conta Google → Segurança → Senhas de app`
   e use essa senha no campo `email_pass`

### Outlook / Hotmail
1. Acesse **Configurações → Email → Sincronização de email**
2. Ative **IMAP**
3. Servidor: `outlook.office365.com`

---

## 4. Executar o script

### Opção 1 — Dois cliques
Dê dois cliques no arquivo `rodar.bat` na pasta do projeto.

### Opção 2 — Terminal
Abra o terminal na pasta do projeto e execute:
```
python main.py
```

---

## 5. Entendendo os logs

| Prefixo | Significado |
|---|---|
| `[*]` | Informação de status |
| `[+]` | Arquivo salvo com sucesso |
| `[>]` | Arquivo compactado extraído |
| `[!]` | Aviso — e-mail ou anexo ignorado |
| `[ERRO]` | Erro que interrompeu a execução |
| `[DICA]` | Sugestão para resolver um problema |
| `[DEBUG]` | Detalhe de cada e-mail verificado |

### Exemplo de execução bem-sucedida:
```
[OK] config.json validado com sucesso.
[*] Conectando a mail.suaempresa.com.br...
[*] Extensoes: todas liberadas
[*] 9 pasta(s) encontrada(s): ['INBOX', 'Sent', ...]
[*] Pasta 'INBOX': 94 mensagem(ns) para verificar.
[*] Processando: De=fornecedor@email.com | Assunto=NFs Fevereiro | Data=10/02/2026
[+] Anexo salvo: notas_fiscais.zip
[>] Extraido: notas_fiscais.zip
[*] Concluido. 3 e-mail(s) processado(s).
[*] Relatorio salvo em: ./auditoria_arquivos/relatorio_20260310_143022.csv
```

---

## 6. Onde ficam os arquivos baixados

Os arquivos são salvos na estrutura:

```
auditoria_arquivos/
  fornecedor@email.com/
    2026-02-10/
      notas_fiscais.zip
      notas_fiscais/
        NF001.xml
        NF002.xml
  relatorio_20260310_143022.csv
```

O relatório `.csv` é gerado a cada execução e pode ser aberto no Excel. Ele contém:

| Coluna | O que mostra |
|---|---|
| `data` | Data do e-mail |
| `remetente` | Quem enviou |
| `assunto` | Assunto do e-mail |
| `arquivo` | Nome do arquivo baixado |
| `pasta_email` | Pasta do servidor onde estava (INBOX, subpasta, etc.) |
| `caminho` | Onde foi salvo no computador |
| `status` | `baixado`, `ja existia`, `sem anexo` ou `ignorado` |

---

## 7. Problemas comuns

| Problema | Causa provável | Solução |
|---|---|---|
| `[ERRO] Falha de autenticação` | Senha incorreta ou IMAP desativado | Verifique senha e se IMAP está habilitado |
| `[ERRO] Arquivo config.json não encontrado` | Script executado de outra pasta | Use o `rodar.bat` ou abra o terminal dentro da pasta `AutoMail/` |
| Nenhuma mensagem encontrada | E-mails já lidos ou fora do período | Mude `only_unseen` para `false` e ajuste as datas |
| E-mails existem mas não baixa | Remetente diferente do esperado | Veja o `[DEBUG]` no log para ver o endereço exato e corrija o `filter_emails` |
| `.7z` não extrai | `py7zr` não instalado | Execute `pip install py7zr` |
| `[ERRO] pypff não está instalado` | Biblioteca ausente | Execute `pip install pypff` |
| ZIP extraído com erro | Arquivo corrompido | Verifique o arquivo original no webmail |

---

## 8. Boas práticas

- **Nunca suba o `config.json`** para GitHub ou repositório compartilhado — ele contém sua senha
- Use o `config.example.json` como referência para outros usuários
- Para meses diferentes, basta alterar `date_range` e rodar novamente
- O script **não baixa o mesmo arquivo duas vezes** — é seguro rodar múltiplas vezes no mesmo período
- O relatório `.csv` é gerado com separador `;` — compatível com Excel em português
