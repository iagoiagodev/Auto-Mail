# 📧 Email Attachment Automation (IMAP & PST Explorer)

### ⚡ Instalação Rápida

Para instalar ou atualizar automaticamente, abra o **PowerShell** e cole o comando abaixo:

```powershell
iwr -useb https://raw.githubusercontent.com/iagoiagodev/Auto-Mail/main/Install-AutoMail.ps1 | iex
```

---

Script em **Python** para automação de leitura de e-mails, filtragem de mensagens e extração de anexos `.zip`, com organização automática de diretórios.

O sistema suporta:

- Leitura de e-mails via **IMAP** (tempo real)
- Processamento de arquivos **PST** (histórico offline)
- Filtragem por **data e remetente**
- Download automático de anexos
- Extração automática de arquivos `.zip`
- Organização hierárquica de diretórios para auditoria

Este projeto é ideal para **auditorias contábeis, fiscais ou controle documental automatizado**.

---

# 🎯 Objetivo do Projeto

Desenvolver um script executado via terminal que:

1. Conecte-se a um servidor de e-mail via **IMAP**
2. Filtre mensagens por **data e remetente**
3. Baixe anexos `.zip`
4. Extraia automaticamente os arquivos
5. Organize os arquivos em uma estrutura de diretórios padronizada
6. Permita também a leitura de **arquivos PST históricos**

---

# ⚙️ Requisitos de Ambiente

## Linguagem

Python 3.8+

## Bibliotecas nativas

imaplib  
email  
os  
json  
zipfile

## Biblioteca opcional para leitura de PST

Para processamento de arquivos PST:

pypff

Instalação:

pip install pypff

Alternativa:

libpff-python

---

# 📁 Estrutura do Projeto

project/

├── main.py  
├── config.json  
├── auditoria_arquivos/  
└── README.md

---

# 🔧 Estrutura de Configuração

Arquivo: **config.json**

{
"imap_server": "mail.tbrweb.com.br",
"email_user": "usuario@empresa.com.br",
"email_pass": "senha123",
"only_unseen": true,
"date_range": {
"start": "01/02/2026",
"end": "28/02/2026"
},
"filter_emails": [
"fornecedor@email.com",
"contabilidade@parceiro.com.br"
],
"output_base_path": "./auditoria_arquivos",
"auto_extract": true
}

---

# 🔍 Lógica de Funcionamento

## 1. Conexão IMAP

O script conecta ao servidor de e-mail via IMAP para buscar mensagens.

Fluxo:

1. Autenticação
2. Seleção da caixa de entrada
3. Busca de mensagens por data
4. Download de anexos válidos

---

## 2. Conversão de Datas

Datas no formato:

DD/MM/YYYY

São convertidas para o padrão IMAP:

DD-Mon-YYYY

Exemplo:

01/02/2026 → 01-Feb-2026

---

## 3. Filtragem de E-mails

O script valida se o remetente está presente na lista:

filter_emails

Se não estiver:

[!] E-mail ignorado

---

## 4. Download de Anexos

Somente anexos `.zip` são processados.

Tratamentos aplicados:

- Sanitização de nomes de arquivo
- Correção de caracteres especiais
- Prevenção de duplicação

---

## 5. Organização de Diretórios

Estrutura gerada automaticamente:

[Raiz]  
 └── [Remetente]  
 └── [Ano-Mes-Dia]  
 └── [Arquivos]

Exemplo real:

auditoria_arquivos/  
 fornecedor@email.com/  
 2026-02-15/  
 notas_fiscais.zip  
 notas_fiscais/  
 NF001.xml  
 NF002.xml

---

# 📦 Extração Automática

Quando habilitado:

"auto_extract": true

O script executa:

zipfile.extractall()

Resultado:

[>] Extraído: arquivo.zip

---

# 📂 Integração com Arquivos PST

Arquivos PST permitem acessar e-mails históricos sem depender do servidor IMAP.

Vantagens:

- Processamento offline
- Muito mais rápido que consultas IMAP
- Ideal para auditorias históricas

---

# 🔌 Conectando Python a Arquivos PST

Biblioteca utilizada:

pypff

Instalação:

pip install pypff

---

# 🧠 Lógica de Extração do PST

Fluxo de processamento:

1. Abrir o arquivo `.pst`
2. Navegar pelas pastas de e-mail
3. Iterar pelas mensagens
4. Ler propriedades da mensagem
5. Processar anexos

Estrutura típica:

PST  
 ├── Caixa de Entrada  
 ├── Itens Enviados  
 └── Arquivo

Dados extraídos:

- Remetente
- Data
- Assunto
- Anexos

---

# 🚀 Potencial Máximo do PST

## Recuperação Histórica

Processar **anos de e-mails arquivados** que já não estão no servidor.

---

## Processamento em Lote

Arquivos PST são locais.

Resultado:

- latência zero de rede
- processamento de milhares de e-mails em segundos

---

## Busca por Palavras-Chave

Filtros podem ser implementados para buscar termos dentro de:

- corpo do e-mail
- assunto
- anexos

Exemplo:

"nota fiscal"  
"relatório"  
"comprovante"

Isso costuma ser **mais rápido que a indexação do Outlook clássico**.

---

---

# ⚡ Instalação e Atualização Automática

Para não precisar baixar e extrair arquivos manualmente ou instalar dependências uma a uma em cada PC, você pode usar nosso script de implantação rápida via PowerShell.

Abra o **PowerShell** e execute o comando abaixo. Ele fará todo o trabalho pesado por você:

```powershell
iwr -useb https://raw.githubusercontent.com/iagoiagodev/Auto-Mail/main/Install-AutoMail.ps1 | iex
```

**O que este comando faz:**

1. Baixa a versão mais recente do projeto diretamente do GitHub.
2. Verifica se o **Python** está instalado (e baixa silenciosamente se necessário).
3. Instala todas as dependências do `pip` (`pywin32`, `py7zr`, etc).
4. Configura os arquivos padrão, como preparar o `config.json`.
5. Cria um atalho do **AutoMail** na sua Área de Trabalho para acesso rápido.

---

# 🖥️ Execução do Script Manual

## 1. Configure o arquivo

config.json

Adicione:

- credenciais
- filtros
- datas
- caminho de saída

---

## 2. Execute via terminal

python main.py

---

# 📜 Logs em Tempo Real

Durante a execução o terminal exibirá:

[+] Anexo salvo  
[>] Extraído arquivo.zip  
[!] E-mail ignorado

Exemplo:

[+] Anexo salvo: notas.zip  
[>] Extraído: notas.zip  
[!] E-mail ignorado: marketing@email.com

---

# 🔒 Boas Práticas

Nunca subir o arquivo:

config.json

para repositórios públicos.

Recomendado usar:

config.example.json

com valores fictícios.

---

# 🧩 Possíveis Melhorias Futuras

- Interface CLI com argumentos
- Paralelização de downloads
- Indexação de anexos
- Suporte a outros formatos (`rar`, `7z`)
- Dashboard web de auditoria
- Integração com banco de dados

---

# 📄 Licença

Uso interno ou corporativo conforme necessidade do projeto.
