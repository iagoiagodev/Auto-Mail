import imaplib
import email
import email.utils
import os
import json
import zipfile
import re
import sys
import csv
from datetime import datetime, timedelta

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

try:
    import py7zr
    PY7ZR_AVAILABLE = True
except ImportError:
    PY7ZR_AVAILABLE = False



def validate_config(config):
    errors = []
    warnings = []

    # Campos obrigatórios (output_base_path agora é opcional, default ./auditoria_arquivos)
    required = ["imap_server", "email_user", "email_pass", "date_range"]
    for field in required:
        if field not in config:
            errors.append(f"Campo obrigatorio ausente: '{field}'")
        elif isinstance(config[field], str) and not config[field].strip():
            errors.append(f"Campo '{field}' esta vazio")

    # Validar date_range
    if "date_range" in config and isinstance(config["date_range"], dict):
        for key in ("start", "end"):
            val = config["date_range"].get(key, "")
            if not val:
                errors.append(f"date_range.{key} esta vazio")
            else:
                try:
                    datetime.strptime(val, "%d/%m/%Y")
                except ValueError:
                    errors.append(f"date_range.{key} = '{val}' formato invalido — use DD/MM/AAAA")
        if not errors:
            start = datetime.strptime(config["date_range"]["start"], "%d/%m/%Y")
            end = datetime.strptime(config["date_range"]["end"], "%d/%m/%Y")
            if start > end:
                errors.append(f"date_range.start ({config['date_range']['start']}) e posterior a date_range.end ({config['date_range']['end']})")
    else:
        if "date_range" in config:
            errors.append("Campo 'date_range' deve ser um objeto com 'start' e 'end'")

    # Validar porta
    port = config.get("imap_port", 993)
    if not isinstance(port, int) or port <= 0:
        errors.append(f"imap_port = '{port}' invalido — use 993 (SSL) ou 143 (STARTTLS)")

    # Validar filter_emails (aceita string ou {"email": "...", "output_path": "..."})
    if "filter_emails" in config:
        if not isinstance(config["filter_emails"], list):
            errors.append("'filter_emails' deve ser uma lista")
        elif len(config["filter_emails"]) == 0:
            warnings.append("filter_emails esta vazio — todos os remetentes serao aceitos")
        else:
            for i, entry in enumerate(config["filter_emails"]):
                if isinstance(entry, str):
                    pass  # formato simples, ok
                elif isinstance(entry, dict):
                    if "email" not in entry or not entry["email"].strip():
                        errors.append(f"filter_emails[{i}]: campo 'email' ausente ou vazio")
                else:
                    errors.append(f"filter_emails[{i}]: valor invalido — use string ou {{\"email\": \"...\", \"output_path\": \"...\"}}")

    # Validar pst_file
    pst = config.get("pst_file", "").strip()
    if pst and not os.path.exists(pst):
        errors.append(f"pst_file = '{pst}' — arquivo nao encontrado")

    # Exibir resultado
    if warnings:
        for w in warnings:
            print(f"[AVISO] {w}")

    if errors:
        print("[ERRO] Problemas encontrados no config.json:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)

    print("[OK] config.json validado com sucesso.")


def _app_dir():
    """Retorna o diretório do executável (frozen) ou do script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    config_path = os.path.join(_app_dir(), "config.json")
    if not os.path.exists(config_path):
        print("[ERRO] Arquivo config.json nao encontrado. Copie config.example.json e preencha com suas credenciais.")
        sys.exit(1)
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        print(f"[ERRO] config.json com formato invalido (JSON quebrado): {e}")
        sys.exit(1)
    validate_config(config)
    return config


def build_email_paths(config):
    """Retorna dict {email_lower: output_path} e a lista de emails para filtro."""
    default_path = config.get("output_base_path", "./auditoria_arquivos").strip() or "./auditoria_arquivos"
    email_paths = {}
    filter_emails = []
    for entry in config.get("filter_emails", []):
        if isinstance(entry, str):
            addr = entry.lower().strip()
            email_paths[addr] = default_path
            filter_emails.append(addr)
        elif isinstance(entry, dict):
            addr = entry["email"].lower().strip()
            path = entry.get("output_path", "").strip() or default_path
            email_paths[addr] = path
            filter_emails.append(addr)
    return filter_emails, email_paths, default_path


def convert_date_imap(date_str):
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    return dt.strftime("%d-%b-%Y")


def sanitize_filename(name):
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    return name.strip() or "attachment"


def create_output_dir(base_path, sender, date_str):
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    date_folder = dt.strftime("%d-%m-%Y")
    output_dir = os.path.join(base_path, sender, date_folder)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def save_and_extract(data, filename, output_dir, auto_extract):
    filepath = os.path.join(output_dir, filename)
    if os.path.exists(filepath):
        return
    with open(filepath, "wb") as f:
        f.write(data)
    print(f"[+] Anexo salvo: {filename}")

    if not auto_extract:
        return

    ext = os.path.splitext(filename)[1].lower()
    extract_dir = os.path.join(output_dir, os.path.splitext(filename)[0])

    if ext == ".zip":
        if not zipfile.is_zipfile(filepath):
            print(f"[!] Arquivo invalido (nao e um zip real): {filename}")
            return
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
        except zipfile.BadZipFile:
            print(f"[!] Falha ao extrair (BadZipFile): {filename}")

    elif ext == ".7z":
        if not PY7ZR_AVAILABLE:
            print(f"[!] py7zr nao instalado — nao foi possivel extrair {filename}. Execute: pip install py7zr")
            return
        try:
            with py7zr.SevenZipFile(filepath, mode="r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
        except Exception as e:
            print(f"[!] Falha ao extrair {filename}: {e}")


def process_imap(config):
    host = config["imap_server"]
    port = config.get("imap_port", 993)
    user = config["email_user"]
    password = config["email_pass"]
    only_unseen = config.get("only_unseen", True)
    filter_emails, email_paths, base_path = build_email_paths(config)
    auto_extract = config.get("auto_extract", True)
    allowed_extensions = {e.lower() for e in config.get("allowed_extensions", [])}
    # lista vazia = aceita tudo
    if allowed_extensions:
        print(f"[*] Filtrando extensoes: {sorted(allowed_extensions)}")
    else:
        print("[*] Extensoes: todas liberadas")

    print(f"[*] Conectando a {host}...")
    try:
        mail = imaplib.IMAP4_SSL(host, port)
        mail.login(user, password)
    except imaplib.IMAP4.error as e:
        print(f"[ERRO] Falha de autenticação: {e}")
        return

    # Listar todas as pastas disponíveis no servidor
    folders = []
    for list_args in [(), ('""', '%'), ('""', '*'), ('"INBOX"', '*'), ('"INBOX."', '*')]:
        try:
            _, raw_folder_list = mail.list(*list_args)
            if not raw_folder_list or raw_folder_list == [None]:
                continue
            for f in raw_folder_list:
                if not f:
                    continue
                raw_str = f.decode()
                match = re.search(r'"[./]"\s+(.*)', raw_str)
                if match:
                    folder_name = match.group(1).strip().strip('"')
                    if folder_name not in folders:
                        folders.append(folder_name)
        except Exception:
            continue
    print(f"[*] {len(folders)} pasta(s) encontrada(s): {folders}")

    start_dt = datetime.strptime(config["date_range"]["start"], "%d/%m/%Y")
    end_dt = datetime.strptime(config["date_range"]["end"], "%d/%m/%Y")

    seen_part = "UNSEEN" if only_unseen else "ALL"
    start_imap = start_dt.strftime("%d-%b-%Y")
    end_imap = (end_dt + timedelta(days=1)).strftime("%d-%b-%Y")
    date_part = f"SINCE {start_imap} BEFORE {end_imap}"

    total_processados = 0
    relatorio = []  # lista de dicts para o CSV

    for folder in folders:
        try:
            rv, _ = mail.select(f'"{folder}"')
            if rv != "OK":
                continue
        except Exception:
            continue

        all_ids = set()
        if filter_emails:
            for addr in filter_emails:
                q = f'{seen_part} {date_part} FROM "{addr}"'
                status, data = mail.search(None, q)
                if status == "OK" and data[0]:
                    all_ids.update(data[0].split())
        else:
            status, data = mail.search(None, f"{seen_part} {date_part}")
            if status == "OK" and data[0]:
                all_ids.update(data[0].split())

        if not all_ids:
            continue

        message_ids = sorted(all_ids)
        print(f"[*] Pasta '{folder}': {len(message_ids)} mensagem(ns) para verificar.")

        for msg_id in message_ids:
            # Busca só o header para checar data antes de baixar o e-mail inteiro
            _, hdata = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM DATE SUBJECT)])")
            if not hdata or not hdata[0]:
                continue

            raw_bytes = hdata[0][1] if isinstance(hdata[0], tuple) else hdata[0]
            hdrs = email.message_from_bytes(raw_bytes)

            raw_date = hdrs.get("Date", "")
            msg_dt = None
            # Tenta parsedate_to_datetime primeiro
            try:
                parsed_date = email.utils.parsedate_to_datetime(raw_date)
                msg_dt = parsed_date.replace(tzinfo=None)
            except Exception:
                pass
            # Fallback para parsedate (mais tolerante a formatos não padrão)
            if msg_dt is None:
                try:
                    t = email.utils.parsedate(raw_date)
                    if t:
                        msg_dt = datetime(*t[:6])
                except Exception:
                    pass
            if msg_dt is None:
                print(f"[!] Nao foi possivel ler a data do e-mail (id={msg_id.decode()}): '{raw_date}'")
                continue

            _, sender_addr = email.utils.parseaddr(hdrs.get("From", ""))
            sender_addr = sender_addr.lower()
            if not (start_dt <= msg_dt <= end_dt.replace(hour=23, minute=59, second=59)):
                continue

            if filter_emails and sender_addr not in filter_emails:
                print(f"[!] E-mail ignorado (remetente fora do filtro): {sender_addr}")
                continue

            # Só agora baixa o e-mail completo
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = msg.get("Subject", "(sem assunto)")
            date_str = msg_dt.strftime("%d/%m/%Y")

            print(f"[*] Processando: De={sender_addr} | Assunto={subject} | Data={date_str}")

            sender_path = email_paths.get(sender_addr, base_path)
            output_dir = create_output_dir(sender_path, sender_addr, date_str)

            attachment_found = False
            for part in msg.walk():
                if part.get_content_disposition() != "attachment":
                    continue
                filename = part.get_filename()
                if not filename:
                    continue
                ext = os.path.splitext(filename)[1].lower()
                if allowed_extensions and ext not in allowed_extensions:
                    print(f"[!] Anexo ignorado (extensao nao permitida): {filename}")
                    relatorio.append({
                        "data": date_str,
                        "remetente": sender_addr,
                        "assunto": subject,
                        "arquivo": filename,
                        "pasta_email": folder,
                        "caminho": "",
                        "status": f"ignorado (extensao {ext} nao permitida)",
                    })
                    continue
                attachment_found = True
                decoded = email.header.decode_header(filename)[0]
                filename = sanitize_filename(decoded[0] if isinstance(decoded[0], str)
                                             else decoded[0].decode(decoded[1] or "utf-8"))
                data_bytes = part.get_payload(decode=True)
                if data_bytes:
                    filepath = os.path.join(output_dir, filename)
                    ja_existia = os.path.exists(filepath)
                    save_and_extract(data_bytes, filename, output_dir, auto_extract)
                    relatorio.append({
                        "data": date_str,
                        "remetente": sender_addr,
                        "assunto": subject,
                        "arquivo": filename,
                        "pasta_email": folder,
                        "caminho": output_dir,
                        "status": "ja existia" if ja_existia else "baixado",
                    })

            if not attachment_found:
                relatorio.append({
                    "data": date_str,
                    "remetente": sender_addr,
                    "assunto": subject,
                    "arquivo": "(sem anexo valido)",
                    "pasta_email": folder,
                    "caminho": "",
                    "status": "sem anexo",
                })
                print(f"[!] Nenhum anexo valido encontrado neste e-mail.")

            total_processados += 1

    mail.logout()

    if total_processados == 0:
        print(f"[*] Nenhum e-mail encontrado no periodo {config['date_range']['start']} a {config['date_range']['end']}.")
        print("[DICA] Se os e-mails ja foram lidos, mude 'only_unseen' para false no config.json.")
    else:
        print(f"[*] Concluido. {total_processados} e-mail(s) processado(s).")

    if relatorio:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = os.path.join(config["output_base_path"], f"relatorio_{ts}.csv")
        os.makedirs(config["output_base_path"], exist_ok=True)
        campos = ["data", "remetente", "assunto", "arquivo", "pasta_email", "caminho", "status"]
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=campos, delimiter=";")
            writer.writeheader()
            writer.writerows(relatorio)
        print(f"[*] Relatorio salvo em: {csv_path}")


def process_pst(config):
    if not WIN32COM_AVAILABLE:
        print("[ERRO] pywin32 nao esta instalado. Execute: pip install pywin32")
        return

    pst_file = config.get("pst_file", "").strip()
    pst_file = os.path.abspath(pst_file)
    if not os.path.exists(pst_file):
        print(f"[ERRO] Arquivo PST nao encontrado: {pst_file}")
        return

    filter_emails, email_paths, base_path = build_email_paths(config)
    base_path = os.path.abspath(base_path)
    email_paths = {k: os.path.abspath(v) for k, v in email_paths.items()}
    auto_extract = config.get("auto_extract", True)
    allowed_extensions = {e.lower() for e in config.get("allowed_extensions", [])}
    start_dt = datetime.strptime(config["date_range"]["start"], "%d/%m/%Y")
    end_dt = datetime.strptime(config["date_range"]["end"], "%d/%m/%Y").replace(hour=23, minute=59, second=59)
    relatorio = []
    total_processados = 0

    print(f"[*] Abrindo PST: {pst_file}")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"[ERRO] Nao foi possivel iniciar o Outlook: {e}")
        return

    def normalizar(p):
        return p.replace("\\", "/").lower().strip("/")

    # Registra caminhos de stores já existentes antes de adicionar
    stores_antes = set()
    try:
        for i in range(1, outlook.Stores.Count + 1):
            try:
                fp = getattr(outlook.Stores.Item(i), "FilePath", "") or ""
                if fp:
                    stores_antes.add(normalizar(fp))
            except Exception:
                pass
        count_antes = outlook.Stores.Count
    except Exception:
        count_antes = 0

    try:
        outlook.AddStoreEx(pst_file, 3)  # 3 = olStoreUnicode
    except Exception as e:
        print(f"[ERRO] Nao foi possivel abrir o PST: {e}")
        return

    # Aguarda o Outlook registrar a store (necessário em alguns ambientes)
    import time
    time.sleep(1)

    pst_norm = normalizar(pst_file)
    pst_basename = normalizar(os.path.basename(pst_file))

    # 1ª tentativa: caminho exato
    store = None
    for i in range(1, outlook.Stores.Count + 1):
        try:
            s = outlook.Stores.Item(i)
            fp = getattr(s, "FilePath", "") or ""
            if normalizar(fp) == pst_norm:
                store = s
                break
        except Exception:
            continue

    # 2ª tentativa: nome do arquivo (útil quando o Outlook resolve o caminho diferente)
    if not store:
        for i in range(1, outlook.Stores.Count + 1):
            try:
                s = outlook.Stores.Item(i)
                fp = getattr(s, "FilePath", "") or ""
                if pst_basename and normalizar(os.path.basename(fp)) == pst_basename:
                    store = s
                    print(f"[*] Store identificada pelo nome do arquivo: {getattr(store, 'DisplayName', '?')}")
                    break
            except Exception:
                continue

    # 3ª tentativa: store nova pelo índice (última da lista)
    if not store:
        try:
            novo_count = outlook.Stores.Count
            if novo_count > count_antes:
                store = outlook.Stores.Item(novo_count)
                print(f"[*] Store identificada pelo indice ({novo_count}): {getattr(store, 'DisplayName', '?')}")
        except Exception:
            pass

    # 4ª tentativa: qualquer store cujo caminho não existia antes
    if not store:
        for i in range(1, outlook.Stores.Count + 1):
            try:
                s = outlook.Stores.Item(i)
                fp = getattr(s, "FilePath", "") or ""
                if fp and normalizar(fp) not in stores_antes:
                    store = s
                    print(f"[*] Store identificada como nova: {getattr(store, 'DisplayName', '?')}")
                    break
            except Exception:
                continue

    if not store:
        print("[ERRO] PST aberto mas nao encontrado nas stores do Outlook.")
        return

    def process_folder(folder):
        nonlocal total_processados
        try:
            folder_name = folder.Name
        except Exception:
            folder_name = "?"

        # Processa subpastas recursivamente
        try:
            for i in range(folder.Folders.Count):
                process_folder(folder.Folders.Item(i + 1))
        except Exception:
            pass

        # Processa mensagens da pasta
        try:
            items = folder.Items
            count = items.Count
        except Exception:
            return

        if count == 0:
            return

        print(f"[*] PST pasta '{folder_name}': {count} item(ns)")

        pasta_processados = 0
        fora_data = []
        fora_remetente = set()

        for i in range(1, count + 1):
            try:
                msg = items.Item(i)
                if msg.Class != 43:  # 43 = olMail
                    continue

                sender_addr = (msg.SenderEmailAddress or "").lower()
                subject = msg.Subject or "(sem assunto)"

                try:
                    msg_dt = msg.ReceivedTime
                    msg_dt = datetime(msg_dt.year, msg_dt.month, msg_dt.day,
                                      msg_dt.hour, msg_dt.minute, msg_dt.second)
                except Exception:
                    continue

                if not (start_dt <= msg_dt <= end_dt):
                    fora_data.append(msg_dt.strftime("%d/%m/%Y"))
                    continue

                if filter_emails and sender_addr not in filter_emails:
                    fora_remetente.add(sender_addr)
                    continue

                pasta_processados += 1

                date_str = msg_dt.strftime("%d/%m/%Y")
                sender_path = email_paths.get(sender_addr, base_path)
                output_dir = create_output_dir(sender_path, sender_addr or "desconhecido", date_str)

                print(f"[*] Processando: De={sender_addr} | Assunto={subject} | Data={date_str}")

                attachment_found = False
                for j in range(1, msg.Attachments.Count + 1):
                    try:
                        att = msg.Attachments.Item(j)
                        filename = att.FileName or ""
                        if not filename:
                            continue
                        ext = os.path.splitext(filename)[1].lower()
                        if allowed_extensions and ext not in allowed_extensions:
                            relatorio.append({
                                "data": date_str, "remetente": sender_addr,
                                "assunto": subject, "arquivo": filename,
                                "pasta_email": folder_name, "caminho": "",
                                "status": f"ignorado (extensao {ext} nao permitida)",
                            })
                            continue
                        filename = sanitize_filename(filename)
                        filepath = os.path.join(output_dir, filename)
                        ja_existia = os.path.exists(filepath)
                        if not ja_existia:
                            att.SaveAsFile(filepath)
                            print(f"[+] Anexo salvo: {filename}")
                            if auto_extract:
                                _extrair_se_compactado(filepath, filename, output_dir)
                        attachment_found = True
                        relatorio.append({
                            "data": date_str, "remetente": sender_addr,
                            "assunto": subject, "arquivo": filename,
                            "pasta_email": folder_name, "caminho": output_dir,
                            "status": "ja existia" if ja_existia else "baixado",
                        })
                    except Exception as e:
                        print(f"[!] Erro ao salvar anexo: {e}")

                if not attachment_found:
                    relatorio.append({
                        "data": date_str, "remetente": sender_addr,
                        "assunto": subject, "arquivo": "(sem anexo valido)",
                        "pasta_email": folder_name, "caminho": "",
                        "status": "sem anexo",
                    })

                total_processados += 1

            except Exception:
                continue

        # Resumo da pasta se nada foi processado
        if pasta_processados == 0 and count > 0:
            if fora_data:
                datas_unicas = sorted(set(fora_data))
                print(f"[!] '{folder_name}': {len(fora_data)} e-mail(s) fora do periodo. Datas encontradas: {datas_unicas[:5]}{'...' if len(datas_unicas) > 5 else ''}")
            if fora_remetente:
                print(f"[!] '{folder_name}': remetentes encontrados no periodo mas fora do filtro: {sorted(fora_remetente)}")

    root = store.GetRootFolder()
    process_folder(root)

    # Remove a store do Outlook após processar
    try:
        outlook.RemoveStore(store.GetRootFolder())
    except Exception:
        pass

    if total_processados == 0:
        print(f"[*] Nenhum e-mail encontrado no PST para o periodo {config['date_range']['start']} a {config['date_range']['end']}.")
    else:
        print(f"[*] PST concluido. {total_processados} e-mail(s) processado(s).")

    if relatorio:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = os.path.join(base_path, f"relatorio_pst_{ts}.csv")
        os.makedirs(base_path, exist_ok=True)
        campos = ["data", "remetente", "assunto", "arquivo", "pasta_email", "caminho", "status"]
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=campos, delimiter=";")
            writer.writeheader()
            writer.writerows(relatorio)
        print(f"[*] Relatorio salvo em: {csv_path}")


def _extrair_se_compactado(filepath, filename, output_dir):
    ext = os.path.splitext(filename)[1].lower()
    extract_dir = os.path.join(output_dir, os.path.splitext(filename)[0])
    if ext == ".zip":
        if not zipfile.is_zipfile(filepath):
            return
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
        except zipfile.BadZipFile:
            print(f"[!] Falha ao extrair (BadZipFile): {filename}")
    elif ext == ".7z" and PY7ZR_AVAILABLE:
        try:
            with py7zr.SevenZipFile(filepath, mode="r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
        except Exception as e:
            print(f"[!] Falha ao extrair {filename}: {e}")


_LAST_RUN = os.path.join(_app_dir(), "last_run.json")


def _load_last_run():
    try:
        with open(_LAST_RUN, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_last_run(data):
    try:
        with open(_LAST_RUN, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _make_date_entry(parent, value=""):
    """Campo único DD/MM/AAAA — barras inseridas automaticamente ao digitar."""
    import tkinter as tk

    frm = tk.Frame(parent)
    ent = tk.Entry(frm, width=11, font=("Segoe UI", 9), justify="center")
    ent.pack()

    _lock = [False]

    def _reformat(event=None):
        if _lock[0]:
            return
        _lock[0] = True
        try:
            raw    = ent.get()
            cursor = ent.index("insert")

            digits = "".join(c for c in raw if c.isdigit())[:8]

            fmt = ""
            for i, d in enumerate(digits):
                if i in (2, 4):
                    fmt += "/"
                fmt += d

            if fmt == raw:
                return  # nada mudou, não mexe no cursor

            ent.delete(0, "end")
            ent.insert(0, fmt)

            # Reposiciona cursor: conta dígitos antes da posição antiga
            dc = min(sum(1 for c in raw[:cursor] if c.isdigit()), len(digits))
            new_pos = len(fmt)
            counted = 0
            for i, c in enumerate(fmt):
                if counted == dc:
                    new_pos = i
                    break
                if c.isdigit():
                    counted += 1
            ent.icursor(new_pos)
        finally:
            _lock[0] = False

    ent.bind("<KeyRelease>", _reformat)
    ent.bind("<FocusIn>", lambda e: ent.after(1, lambda: ent.select_range(0, "end")))

    if value:
        ent.insert(0, value)

    frm.get = ent.get
    return frm


def _sel_all(widget):
    """Bind select-all ao focar em um Entry."""
    widget.bind("<FocusIn>",
                lambda e, w=widget: w.after(1, lambda: w.select_range(0, "end")))


def show_gui(config):
    """Interface completa: PST e IMAP em abas + todas as configurações."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    last   = _load_last_run()
    result = {}
    PAD    = 10

    root = tk.Tk()
    root.title("AutoMail")
    root.resizable(True, True)
    root.minsize(640, 480)
    root.attributes("-topmost", True)

    # Coluna única expande horizontalmente; linha 2 (remetentes) expande verticalmente
    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)

    # ── ABAS: PST / IMAP ─────────────────────────────────────────────────
    notebook = ttk.Notebook(root)
    notebook.grid(row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 4))

    # ── ABA PST ───────────────────────────────────────────────────────────
    tab_pst = tk.Frame(notebook, padx=PAD, pady=PAD)
    notebook.add(tab_pst, text="  PST  ")

    pst_var = tk.StringVar(value=last.get("pst_file", config.get("pst_file", "")))
    tk.Label(tab_pst, text="Arquivo PST:", font=("Segoe UI", 9)).pack(anchor="w")
    frm_pst_row = tk.Frame(tab_pst)
    frm_pst_row.pack(fill="x", pady=(2, 0))
    ent_pst = tk.Entry(frm_pst_row, textvariable=pst_var, width=60, font=("Segoe UI", 9))
    ent_pst.pack(side="left", padx=(0, 6))
    _sel_all(ent_pst)

    def browse_pst():
        p = filedialog.askopenfilename(
            title="Selecione o arquivo PST",
            filetypes=[("Arquivo PST do Outlook", "*.pst"), ("Todos os arquivos", "*.*")]
        )
        if p:
            pst_var.set(p)

    tk.Button(frm_pst_row, text="…", width=3, command=browse_pst,
              font=("Segoe UI", 9)).pack(side="left")

    # ── ABA IMAP ──────────────────────────────────────────────────────────
    tab_imap = tk.Frame(notebook, padx=PAD, pady=PAD)
    notebook.add(tab_imap, text="  IMAP  ")

    last_imap = last.get("imap", {})

    def imap_row(parent, label, key, cfg_key, width=32, show=""):
        val = last_imap.get(key, config.get(cfg_key, ""))
        row = tk.Frame(parent)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, font=("Segoe UI", 9),
                 width=12, anchor="w").pack(side="left")
        var = tk.StringVar(value=str(val) if val != "" else "")
        ent = tk.Entry(row, textvariable=var, width=width,
                       font=("Segoe UI", 9), show=show)
        ent.pack(side="left")
        _sel_all(ent)
        return var

    imap_server_var = imap_row(tab_imap, "Servidor:",  "imap_server", "imap_server", 28)
    imap_port_var   = imap_row(tab_imap, "Porta:",     "imap_port",   "imap_port",   6)
    imap_user_var   = imap_row(tab_imap, "Usuário:",   "email_user",  "email_user",  34)
    imap_pass_var   = imap_row(tab_imap, "Senha:",     "email_pass",  "email_pass",  34, show="•")

    unseen_val = last_imap.get("only_unseen", config.get("only_unseen", True))
    only_unseen_var = tk.BooleanVar(value=bool(unseen_val))
    tk.Checkbutton(tab_imap, text="Apenas e-mails não lidos",
                   variable=only_unseen_var,
                   font=("Segoe UI", 9)).pack(anchor="w", pady=(6, 0))

    # ── PERÍODO (compartilhado) ───────────────────────────────────────────
    frm_date = tk.LabelFrame(root, text="Período", font=("Segoe UI", 9),
                              padx=PAD, pady=6)
    frm_date.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(0, 4))

    last_dr = last.get("date_range", config.get("date_range", {}))
    tk.Label(frm_date, text="De:", font=("Segoe UI", 9)).pack(side="left")
    date_start = _make_date_entry(frm_date, last_dr.get("start", ""))
    date_start.pack(side="left", padx=(4, 20))
    tk.Label(frm_date, text="Até:", font=("Segoe UI", 9)).pack(side="left")
    date_end = _make_date_entry(frm_date, last_dr.get("end", ""))
    date_end.pack(side="left", padx=(4, 0))

    # ── REMETENTES (compartilhado) ────────────────────────────────────────
    frm_table = tk.LabelFrame(root, text="Remetentes e pastas de destino",
                               font=("Segoe UI", 9), padx=PAD, pady=6)
    frm_table.grid(row=2, column=0, sticky="nsew", padx=PAD, pady=(0, 4))
    # Coluna 0 expande; linha 1 (canvas) expande verticalmente
    frm_table.columnconfigure(0, weight=1)
    frm_table.rowconfigure(1, weight=1)

    # Cabeçalho (fora do canvas para ficar fixo)
    frm_header = tk.Frame(frm_table)
    frm_header.grid(row=0, column=0, sticky="ew", padx=(0, 17))
    tk.Label(frm_header, text="E-mail do remetente",
             font=("Segoe UI", 9, "bold"), width=32, anchor="w").pack(side="left", padx=(0, 6))
    tk.Label(frm_header, text="Pasta de destino",
             font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left")

    # Canvas com barra de rolagem — altura inicial para 5 linhas, expande com a janela
    _ROW_H = 34
    frm_canvas = tk.Frame(frm_table)
    frm_canvas.grid(row=1, column=0, sticky="nsew")
    frm_canvas.rowconfigure(0, weight=1)
    frm_canvas.columnconfigure(0, weight=1)

    canvas = tk.Canvas(frm_canvas, height=_ROW_H * 5, highlightthickness=0, bd=0)
    scrollbar = tk.Scrollbar(frm_canvas, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    inner = tk.Frame(canvas)
    _win  = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _sync_canvas(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _fit_inner(event):
        canvas.itemconfig(_win, width=event.width)

    inner.bind("<Configure>", _sync_canvas)
    canvas.bind("<Configure>", _fit_inner)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind("<MouseWheel>", _on_mousewheel)
    inner.bind("<MouseWheel>", _on_mousewheel)

    rows = []

    def add_row(email="", path=""):
        frm = tk.Frame(inner)
        frm.pack(fill="x", pady=2)

        email_var = tk.StringVar(value=email)
        path_var  = tk.StringVar(value=path)

        ent_email = tk.Entry(frm, textvariable=email_var, width=32,
                             font=("Segoe UI", 9))
        ent_email.pack(side="left", padx=(0, 6))
        _sel_all(ent_email)

        ent_path = tk.Entry(frm, textvariable=path_var, width=34,
                            font=("Segoe UI", 9))
        ent_path.pack(side="left", padx=(0, 4))
        _sel_all(ent_path)

        def browse_dir(pv=path_var):
            p = filedialog.askdirectory(title="Selecione a pasta de destino")
            if p:
                pv.set(p.replace("/", "\\"))

        tk.Button(frm, text="…", width=2, command=browse_dir,
                  font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        row_ref = [email_var, path_var, frm]

        def _remove(f=frm, r=row_ref):
            f.destroy()
            rows.remove(r)
            _sync_canvas()

        tk.Button(frm, text="✕", width=2, fg="red", font=("Segoe UI", 9),
                  command=_remove).pack(side="left")

        # Propaga mousewheel para o canvas em todos os filhos da linha
        for w in (frm, ent_email, ent_path):
            w.bind("<MouseWheel>", _on_mousewheel)

        rows.append(row_ref)
        canvas.after(50, lambda: canvas.yview_moveto(1.0))

    default_path = config.get("output_base_path", "")
    saved_emails = last.get("filter_emails", config.get("filter_emails", []))
    if saved_emails:
        for entry in saved_emails:
            if isinstance(entry, str):
                add_row(email=entry, path=default_path)
            elif isinstance(entry, dict):
                add_row(email=entry.get("email", ""),
                        path=entry.get("output_path", default_path))
    else:
        add_row()

    tk.Button(frm_table, text="+ Adicionar remetente",
              font=("Segoe UI", 9), command=add_row).grid(
        row=2, column=0, pady=(8, 2), sticky="w")

    # ── OPÇÕES ────────────────────────────────────────────────────────────
    frm_opts = tk.LabelFrame(root, text="Opções", font=("Segoe UI", 9),
                              padx=PAD, pady=6)
    frm_opts.grid(row=3, column=0, sticky="ew", padx=PAD, pady=(0, 4))
    frm_opts.columnconfigure(0, weight=1)

    # Pasta padrão de saída
    out_val = last.get("output_base_path", config.get("output_base_path", "./auditoria_arquivos"))
    out_var = tk.StringVar(value=out_val)
    row_out = tk.Frame(frm_opts)
    row_out.pack(fill="x", pady=2)
    tk.Label(row_out, text="Pasta padrão:", font=("Segoe UI", 9),
             width=16, anchor="w").pack(side="left")
    ent_out = tk.Entry(row_out, textvariable=out_var, width=46, font=("Segoe UI", 9))
    ent_out.pack(side="left", padx=(0, 4))
    _sel_all(ent_out)
    def browse_out():
        p = filedialog.askdirectory(title="Pasta padrão de saída")
        if p:
            out_var.set(p.replace("/", "\\"))
    tk.Button(row_out, text="…", width=2, command=browse_out,
              font=("Segoe UI", 9)).pack(side="left")

    # auto_extract
    auto_val = last.get("auto_extract", config.get("auto_extract", True))
    auto_var = tk.BooleanVar(value=bool(auto_val))
    tk.Checkbutton(frm_opts, text="Extrair arquivos compactados automaticamente",
                   variable=auto_var, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 2))

    # allowed_extensions
    row_ext = tk.Frame(frm_opts)
    row_ext.pack(fill="x", pady=2)
    tk.Label(row_ext, text="Extensões permitidas:",
             font=("Segoe UI", 9)).pack(side="left")
    tk.Label(row_ext, text="(vazio = todas)",
             font=("Segoe UI", 8), fg="#888").pack(side="left", padx=(4, 8))
    exts_raw = last.get("allowed_extensions", config.get("allowed_extensions", []))
    exts_str = ", ".join(exts_raw) if isinstance(exts_raw, list) else str(exts_raw)
    exts_var = tk.StringVar(value=exts_str)
    ent_exts = tk.Entry(row_ext, textvariable=exts_var, width=36, font=("Segoe UI", 9))
    ent_exts.pack(side="left")
    _sel_all(ent_exts)

    # ── BOTÕES ────────────────────────────────────────────────────────────
    frm_btn = tk.Frame(root)
    frm_btn.grid(row=4, column=0, pady=(0, PAD))

    def on_ok():
        start = date_start.get()
        end   = date_end.get()
        try:
            datetime.strptime(start, "%d/%m/%Y")
            datetime.strptime(end,   "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Datas inválidas. Use DD/MM/AAAA.")
            return

        default_out = out_var.get().strip() or "./auditoria_arquivos"
        entries = []
        for email_var, path_var, _ in rows:
            e = email_var.get().strip().lower()
            p = path_var.get().strip() or default_out
            if e:
                entries.append({"email": e, "output_path": p})

        exts_input = exts_var.get().strip()
        allowed_exts = [x.strip().lower() for x in exts_input.split(",")
                        if x.strip()] if exts_input else []

        tab_idx = notebook.index(notebook.select())
        mode    = "pst" if tab_idx == 0 else "imap"

        if mode == "pst":
            pst = pst_var.get().strip()
            if not pst:
                messagebox.showerror("Erro", "Selecione um arquivo PST.")
                return
            result["pst_file"] = pst
        else:
            srv  = imap_server_var.get().strip()
            port = imap_port_var.get().strip()
            usr  = imap_user_var.get().strip()
            pwd  = imap_pass_var.get().strip()
            if not all([srv, port, usr, pwd]):
                messagebox.showerror("Erro", "Preencha todos os campos IMAP.")
                return
            try:
                port_int = int(port)
            except ValueError:
                messagebox.showerror("Erro", "Porta IMAP inválida.")
                return
            result.update({
                "imap_server": srv,
                "imap_port":   port_int,
                "email_user":  usr,
                "email_pass":  pwd,
                "only_unseen": only_unseen_var.get(),
            })

        result.update({
            "mode":               mode,
            "date_range":         {"start": start, "end": end},
            "filter_emails":      entries,
            "output_base_path":   default_out,
            "auto_extract":       auto_var.get(),
            "allowed_extensions": allowed_exts,
        })

        # Persiste tudo exceto senha
        to_save = {k: v for k, v in result.items() if k != "email_pass"}
        if mode == "imap":
            to_save["imap"] = {k: result[k] for k in
                               ("imap_server", "imap_port", "email_user", "only_unseen")
                               if k in result}
        _save_last_run(to_save)
        root.destroy()

    def on_cancel():
        root.destroy()

    tk.Button(frm_btn, text="Iniciar", width=12, font=("Segoe UI", 9, "bold"),
              bg="#0078D4", fg="white", command=on_ok).pack(side="left", padx=6)
    tk.Button(frm_btn, text="Cancelar", width=10, font=("Segoe UI", 9),
              command=on_cancel).pack(side="left", padx=6)

    # Seleciona a aba do último modo usado
    last_mode = last.get("mode", "pst" if config.get("pst_file", "") else "imap")
    notebook.select(0 if last_mode == "pst" else 1)

    root.mainloop()
    return result


def main():
    config = load_config()

    cfg = show_gui(config)
    if not cfg:
        print("[*] Cancelado pelo usuário.")
        return

    config.update(cfg)

    if cfg.get("mode") == "imap":
        process_imap(config)
    else:
        process_pst(config)


if __name__ == "__main__":
    main()
