import imaplib
import email
import email.header
import email.utils
import os
import json
import zipfile
import re
import sys
import csv
import socket
import threading
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
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

try:
    import rarfile
    RAR_AVAILABLE = True
except ImportError:
    RAR_AVAILABLE = False


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def validate_config(config):
    errors = []
    warnings = []

    required = ["imap_server", "email_user", "email_pass", "date_range"]
    for field in required:
        if field not in config:
            errors.append(f"Campo obrigatorio ausente: '{field}'")
        elif isinstance(config[field], str) and not config[field].strip():
            errors.append(f"Campo '{field}' esta vazio")

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
                errors.append(
                    f"date_range.start ({config['date_range']['start']}) "
                    f"e posterior a date_range.end ({config['date_range']['end']})"
                )
    else:
        if "date_range" in config:
            errors.append("Campo 'date_range' deve ser um objeto com 'start' e 'end'")

    port = config.get("imap_port", 993)
    if not isinstance(port, int) or port <= 0:
        errors.append(f"imap_port = '{port}' invalido — use 993 (SSL) ou 143 (STARTTLS)")

    if "filter_emails" in config:
        if not isinstance(config["filter_emails"], list):
            errors.append("'filter_emails' deve ser uma lista")
        elif len(config["filter_emails"]) == 0:
            warnings.append("filter_emails esta vazio — todos os remetentes serao aceitos")
        else:
            for i, entry in enumerate(config["filter_emails"]):
                if isinstance(entry, str):
                    pass
                elif isinstance(entry, dict):
                    if "email" not in entry or not entry["email"].strip():
                        errors.append(f"filter_emails[{i}]: campo 'email' ausente ou vazio")
                else:
                    errors.append(
                        f"filter_emails[{i}]: valor invalido — "
                        f"use string ou {{\"email\": \"...\", \"output_path\": \"...\"}}"
                    )

    if config.get("allowed_extensions") and config.get("blocked_extensions"):
        errors.append(
            "Use apenas 'allowed_extensions' OU 'blocked_extensions', nao os dois ao mesmo tempo"
        )

    pst = config.get("pst_file", "").strip()
    if pst and not os.path.exists(pst):
        errors.append(f"pst_file = '{pst}' — arquivo nao encontrado")

    if warnings:
        for w in warnings:
            print(f"[AVISO] {w}")

    if errors:
        print("[AVISO] Faltam informacoes no config.json. Use a interface para preencher.")
        for e in errors:
            print(f"  - {e}")
        return False

    print("[OK] config.json validado com sucesso.")
    return True


def _app_dir():
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# PR 1 — Extração unificada: .zip / .7z / .rar
# ---------------------------------------------------------------------------

def _extract_archive(filepath, filename, output_dir):
    """Extrai .zip, .7z ou .rar para um subdiretório. Retorna True em sucesso."""
    ext = os.path.splitext(filename)[1].lower()
    extract_dir = os.path.join(output_dir, os.path.splitext(filename)[0])

    if ext == ".zip":
        if not zipfile.is_zipfile(filepath):
            print(f"[!] Arquivo invalido (nao e um ZIP real): {filename}")
            return False
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
            return True
        except zipfile.BadZipFile:
            print(f"[!] Falha ao extrair (BadZipFile): {filename}")
        except Exception as e:
            print(f"[!] Falha ao extrair ZIP {filename}: {e}")
        return False

    elif ext == ".7z":
        if not PY7ZR_AVAILABLE:
            print(f"[!] py7zr nao instalado — nao foi possivel extrair {filename}. Execute: pip install py7zr")
            return False
        try:
            with py7zr.SevenZipFile(filepath, mode="r") as zf:
                zf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
            return True
        except Exception as e:
            print(f"[!] Falha ao extrair 7z {filename}: {e}")
        return False

    elif ext == ".rar":
        if not RAR_AVAILABLE:
            print(f"[!] rarfile nao instalado — nao foi possivel extrair {filename}. Execute: pip install rarfile")
            return False
        try:
            with rarfile.RarFile(filepath) as rf:
                rf.extractall(extract_dir)
            print(f"[>] Extraido: {filename}")
            return True
        except rarfile.BadRarFile:
            print(f"[!] Falha ao extrair (BadRarFile): {filename}")
        except Exception as e:
            print(f"[!] Falha ao extrair RAR {filename}: {e}")
        return False

    return False  # formato nao suportado para extracao


def save_and_extract(data, filename, output_dir, auto_extract):
    """Salva o arquivo em disco e extrai se compactado. Retorna status string."""
    filepath = os.path.join(output_dir, filename)
    if os.path.exists(filepath):
        return "ja existia"
    try:
        with open(filepath, "wb") as f:
            f.write(data)
        print(f"[+] Anexo salvo: {filename}")
    except OSError as e:
        print(f"[!] Erro ao salvar {filename}: {e}")
        return "erro ao salvar"

    if auto_extract:
        _extract_archive(filepath, filename, output_dir)
    return "baixado"


# ---------------------------------------------------------------------------
# PR 3 — Log de auditoria CSV centralizado
# ---------------------------------------------------------------------------

def write_audit_log(relatorio, output_path, prefix="relatorio"):
    """Persiste o relatório de auditoria em CSV (UTF-8-BOM para Excel)."""
    if not relatorio:
        return
    try:
        os.makedirs(output_path, exist_ok=True)
    except OSError as e:
        print(f"[!] Nao foi possivel criar pasta para o relatorio: {e}")
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(output_path, f"{prefix}_{ts}.csv")
    campos = ["data", "remetente", "assunto", "arquivo", "pasta_email", "caminho", "status"]
    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=campos, delimiter=";")
            writer.writeheader()
            writer.writerows(relatorio)
        print(f"[*] Relatorio salvo em: {csv_path}")
    except OSError as e:
        print(f"[!] Erro ao salvar relatorio CSV: {e}")


# ---------------------------------------------------------------------------
# PR 2 — Worker paralelo para processamento de anexos
# ---------------------------------------------------------------------------

def _process_email_attachments(raw_msg_bytes, sender_addr, msg_dt, folder,
                                email_paths, base_path, auto_extract,
                                allowed_extensions, blocked_extensions):
    """
    Processa os anexos de um único e-mail.
    Executado em thread separada — não usa IMAP, apenas I/O de disco.
    Retorna lista de dicts para o relatório de auditoria.
    """
    entries = []
    try:
        msg = email.message_from_bytes(raw_msg_bytes)
    except Exception as e:
        print(f"[!] Erro ao parsear mensagem de {sender_addr}: {e}")
        return entries

    subject = msg.get("Subject", "(sem assunto)")
    date_str = msg_dt.strftime("%d/%m/%Y")
    sender_path = email_paths.get(sender_addr, base_path)

    try:
        output_dir = create_output_dir(sender_path, sender_addr, date_str)
    except OSError as e:
        print(f"[!] Erro ao criar diretorio para {sender_addr}: {e}")
        return entries

    attachment_found = False
    for part in msg.walk():
        if part.get_content_disposition() != "attachment":
            continue
        raw_name = part.get_filename()
        if not raw_name:
            continue

        ext = os.path.splitext(raw_name)[1].lower()
        if allowed_extensions and ext not in allowed_extensions:
            print(f"[!] Anexo ignorado (extensao nao permitida): {raw_name}")
            entries.append({
                "data": date_str, "remetente": sender_addr, "assunto": subject,
                "arquivo": raw_name, "pasta_email": folder, "caminho": "",
                "status": f"ignorado (extensao {ext} nao permitida)",
            })
            continue
        if blocked_extensions and ext in blocked_extensions:
            print(f"[!] Anexo ignorado (extensao bloqueada): {raw_name}")
            entries.append({
                "data": date_str, "remetente": sender_addr, "assunto": subject,
                "arquivo": raw_name, "pasta_email": folder, "caminho": "",
                "status": f"ignorado (extensao {ext} bloqueada)",
            })
            continue

        attachment_found = True

        # Decodifica nome do arquivo com fallback robusto
        try:
            decoded_parts = email.header.decode_header(raw_name)
            decoded_val, charset = decoded_parts[0]
            if isinstance(decoded_val, bytes):
                filename = decoded_val.decode(charset or "utf-8", errors="replace")
            else:
                filename = decoded_val
        except Exception:
            filename = raw_name
        filename = sanitize_filename(filename)

        data_bytes = part.get_payload(decode=True)
        if not data_bytes:
            continue

        status = save_and_extract(data_bytes, filename, output_dir, auto_extract)
        entries.append({
            "data": date_str, "remetente": sender_addr, "assunto": subject,
            "arquivo": filename, "pasta_email": folder, "caminho": output_dir,
            "status": status,
        })

    if not attachment_found:
        entries.append({
            "data": date_str, "remetente": sender_addr, "assunto": subject,
            "arquivo": "(sem anexo valido)", "pasta_email": folder, "caminho": "",
            "status": "sem anexo",
        })
        print(f"[!] Nenhum anexo valido encontrado neste e-mail (De={sender_addr}).")

    return entries


# ---------------------------------------------------------------------------
# IMAP
# ---------------------------------------------------------------------------

IMAP_TIMEOUT = 30  # segundos


def process_imap(config):
    host = config["imap_server"]
    port = config.get("imap_port", 993)
    user = config["email_user"]
    password = config["email_pass"]
    only_unseen = config.get("only_unseen", True)
    filter_emails, email_paths, base_path = build_email_paths(config)
    auto_extract = config.get("auto_extract", True)
    max_workers = config.get("max_workers", 4)
    allowed_extensions = {e.lower() for e in config.get("allowed_extensions", [])}
    blocked_extensions = {e.lower() for e in config.get("blocked_extensions", [])}

    if allowed_extensions:
        print(f"[*] Extensoes permitidas (whitelist): {sorted(allowed_extensions)}")
    elif blocked_extensions:
        print(f"[*] Extensoes bloqueadas (blacklist): {sorted(blocked_extensions)}")
    else:
        print("[*] Extensoes: todas liberadas")

    print(f"[*] Conectando a {host}...")

    # Bug fix: aplica timeout para evitar travamento silencioso
    old_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(IMAP_TIMEOUT)
    mail = None
    try:
        try:
            mail = imaplib.IMAP4_SSL(host, port)
            mail.login(user, password)
        except imaplib.IMAP4.error as e:
            print(f"[ERRO] Falha de autenticacao: {e}")
            return
        except (socket.timeout, OSError) as e:
            print(f"[ERRO] Timeout/conexao ao conectar em {host}:{port} — {e}")
            return

        # Descoberta de pastas
        folders = _list_imap_folders(mail)
        print(f"[*] {len(folders)} pasta(s) encontrada(s): {folders}")

        start_dt = datetime.strptime(config["date_range"]["start"], "%d/%m/%Y")
        end_dt = datetime.strptime(config["date_range"]["end"], "%d/%m/%Y")

        seen_part = "UNSEEN" if only_unseen else "ALL"
        start_imap = start_dt.strftime("%d-%b-%Y")
        end_imap = (end_dt + timedelta(days=1)).strftime("%d-%b-%Y")
        date_part = f"SINCE {start_imap} BEFORE {end_imap}"

        # ── Fase 1: coleta sequencial (IMAP não é thread-safe) ───────────
        # Acumula (raw_bytes, sender_addr, msg_dt, folder_name)
        pending = []

        for folder in folders:
            try:
                rv, _ = mail.select(f'"{folder}"')
                if rv != "OK":
                    continue
            except Exception as e:
                print(f"[!] Nao foi possivel selecionar pasta '{folder}': {e}")
                continue

            all_ids = _search_folder(mail, filter_emails, seen_part, date_part)
            if not all_ids:
                continue

            message_ids = sorted(all_ids)
            print(f"[*] Pasta '{folder}': {len(message_ids)} mensagem(ns) para verificar.")

            for msg_id in message_ids:
                try:
                    _, hdata = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM DATE SUBJECT)])")
                except Exception as e:
                    print(f"[!] Erro ao buscar cabecalho (id={msg_id}): {e}")
                    continue

                if not hdata or not hdata[0]:
                    continue

                raw_bytes = hdata[0][1] if isinstance(hdata[0], tuple) else hdata[0]
                hdrs = email.message_from_bytes(raw_bytes)

                msg_dt = _parse_email_date(hdrs.get("Date", ""), msg_id)
                if msg_dt is None:
                    continue

                if not (start_dt <= msg_dt <= end_dt.replace(hour=23, minute=59, second=59)):
                    continue

                _, sender_addr = email.utils.parseaddr(hdrs.get("From", ""))
                sender_addr = sender_addr.lower()

                if filter_emails and sender_addr not in filter_emails:
                    print(f"[!] E-mail ignorado (remetente fora do filtro): {sender_addr}")
                    continue

                # Só baixa o corpo completo após passar todos os filtros
                try:
                    status, msg_data = mail.fetch(msg_id, "(RFC822)")
                    if status != "OK" or not msg_data or not msg_data[0]:
                        continue
                except Exception as e:
                    print(f"[!] Erro ao baixar mensagem (id={msg_id}): {e}")
                    continue

                raw_full = msg_data[0][1]
                print(f"[*] Enfileirado: De={sender_addr} | Data={msg_dt.strftime('%d/%m/%Y')}")
                pending.append((raw_full, sender_addr, msg_dt, folder))

    finally:
        socket.setdefaulttimeout(old_timeout)
        if mail:
            try:
                mail.logout()
            except Exception:
                pass

    if not pending:
        print(
            f"[*] Nenhum e-mail encontrado no periodo "
            f"{config['date_range']['start']} a {config['date_range']['end']}."
        )
        print("[DICA] Se os e-mails ja foram lidos, mude 'only_unseen' para false no config.json.")
        return

    # ── Fase 2: processamento paralelo (I/O de disco) ────────────────────
    print(f"[*] Processando {len(pending)} e-mail(s) com {max_workers} worker(s)...")
    relatorio = []
    relatorio_lock = threading.Lock()

    def _worker(args):
        raw, sender, dt, folder_name = args
        return _process_email_attachments(
            raw, sender, dt, folder_name,
            email_paths, base_path, auto_extract, allowed_extensions, blocked_extensions,
        )

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_worker, item): item for item in pending}
        for future in as_completed(futures):
            raw, sender, dt, _ = futures[future]
            try:
                entries = future.result()
                with relatorio_lock:
                    relatorio.extend(entries)
            except Exception as e:
                print(f"[!] Erro inesperado ao processar e-mail de {sender} ({dt.strftime('%d/%m/%Y')}): {e}")

    print(f"[*] Concluido. {len(pending)} e-mail(s) processado(s).")
    write_audit_log(relatorio, base_path, prefix="relatorio_imap")


def _list_imap_folders(mail):
    """Retorna lista de nomes de pastas disponíveis no servidor."""
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
    return folders


def _search_folder(mail, filter_emails, seen_part, date_part):
    """Executa a busca IMAP em uma pasta e retorna set de IDs."""
    all_ids = set()
    if filter_emails:
        for addr in filter_emails:
            q = f'{seen_part} {date_part} FROM "{addr}"'
            try:
                status, data = mail.search(None, q)
                if status == "OK" and data[0]:
                    all_ids.update(data[0].split())
            except imaplib.IMAP4.error as e:
                print(f"[!] Erro na busca IMAP (FROM={addr}): {e}")
    else:
        try:
            status, data = mail.search(None, f"{seen_part} {date_part}")
            if status == "OK" and data[0]:
                all_ids.update(data[0].split())
        except imaplib.IMAP4.error as e:
            print(f"[!] Erro na busca IMAP: {e}")
    return all_ids


def _parse_email_date(raw_date, msg_id=b"?"):
    """Parseia a data de um e-mail com dois fallbacks. Retorna datetime ou None."""
    try:
        parsed = email.utils.parsedate_to_datetime(raw_date)
        return parsed.replace(tzinfo=None)
    except Exception:
        pass
    try:
        t = email.utils.parsedate(raw_date)
        if t:
            return datetime(*t[:6])
    except Exception:
        pass
    mid = msg_id.decode() if isinstance(msg_id, bytes) else str(msg_id)
    print(f"[!] Nao foi possivel ler a data do e-mail (id={mid}): '{raw_date}'")
    return None


# ---------------------------------------------------------------------------
# PST
# ---------------------------------------------------------------------------

# Formato de data que o Outlook espera na string de Restrict (locale-independente via @SQL)
_RESTRICT_FMT = "%Y-%m-%d %H:%M:%S"


def _find_pst_store(outlook, pst_file, count_antes, stores_antes):
    """
    Tenta localizar a store do PST recém-adicionado usando 4 estratégias.
    Retorna o objeto store ou None.
    """
    def normalizar(p):
        return p.replace("\\", "/").lower().strip("/")

    pst_norm = normalizar(pst_file)
    pst_basename = normalizar(os.path.basename(pst_file))

    # Retry loop: aguarda o Outlook registrar a store (máximo 10 segundos)
    for _ in range(20):
        if outlook.Stores.Count > count_antes:
            break
        time.sleep(0.5)

    # 1ª tentativa: caminho exato
    for i in range(1, outlook.Stores.Count + 1):
        try:
            s = outlook.Stores.Item(i)
            fp = getattr(s, "FilePath", "") or ""
            if normalizar(fp) == pst_norm:
                return s
        except Exception:
            continue

    # 2ª tentativa: nome do arquivo (Outlook às vezes resolve o caminho diferente)
    for i in range(1, outlook.Stores.Count + 1):
        try:
            s = outlook.Stores.Item(i)
            fp = getattr(s, "FilePath", "") or ""
            if pst_basename and normalizar(os.path.basename(fp)) == pst_basename:
                print(f"[*] Store identificada pelo nome: {getattr(s, 'DisplayName', '?')}")
                return s
        except Exception:
            continue

    # 3ª tentativa: última store da lista (nova pelo índice)
    try:
        novo_count = outlook.Stores.Count
        if novo_count > count_antes:
            s = outlook.Stores.Item(novo_count)
            print(f"[*] Store identificada pelo indice ({novo_count}): {getattr(s, 'DisplayName', '?')}")
            return s
    except Exception:
        pass

    # 4ª tentativa: qualquer store com caminho que não existia antes
    for i in range(1, outlook.Stores.Count + 1):
        try:
            s = outlook.Stores.Item(i)
            fp = getattr(s, "FilePath", "") or ""
            if fp and normalizar(fp) not in stores_antes:
                print(f"[*] Store identificada como nova: {getattr(s, 'DisplayName', '?')}")
                return s
        except Exception:
            continue

    return None


def _pst_restrict_by_date(items, start_dt, end_dt):
    """
    Filtra a coleção de itens do Outlook pela data de recebimento usando Restrict.
    Muito mais rápido que iterar tudo: o Outlook aplica o filtro internamente.
    Retorna a coleção filtrada, ou a original em caso de falha.
    """
    try:
        # Sintaxe DASL — funciona independente de locale regional
        restr = (
            "@SQL=\"urn:schemas:httpmail:datereceived\" >= "
            f"'{start_dt.strftime(_RESTRICT_FMT)}' AND "
            "\"urn:schemas:httpmail:datereceived\" <= "
            f"'{end_dt.strftime(_RESTRICT_FMT)}'"
        )
        restricted = items.Restrict(restr)
        count_original = items.Count
        count_filtered = restricted.Count
        if count_original > 0:
            reducao = round((1 - count_filtered / count_original) * 100)
            print(f"    Restrict: {count_original} → {count_filtered} itens ({reducao}% filtrados pelo Outlook)")
        return restricted
    except Exception as e:
        print(f"    [!] Restrict nao suportado nesta pasta, iterando tudo: {e}")
        return items


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
    max_workers = config.get("max_workers", 4)
    allowed_extensions = {e.lower() for e in config.get("allowed_extensions", [])}
    blocked_extensions = {e.lower() for e in config.get("blocked_extensions", [])}

    if allowed_extensions:
        print(f"[*] Extensoes permitidas (whitelist): {sorted(allowed_extensions)}")
    elif blocked_extensions:
        print(f"[*] Extensoes bloqueadas (blacklist): {sorted(blocked_extensions)}")
    else:
        print("[*] Extensoes: todas liberadas")

    start_dt = datetime.strptime(config["date_range"]["start"], "%d/%m/%Y")
    end_dt = datetime.strptime(config["date_range"]["end"], "%d/%m/%Y").replace(
        hour=23, minute=59, second=59
    )

    relatorio = []
    total_processados = 0
    # Acumula arquivos salvos que precisam de extração — processados em paralelo depois
    pending_extractions = []  # lista de (filepath, filename, output_dir)

    print(f"[*] Abrindo PST: {pst_file}")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"[ERRO] Nao foi possivel iniciar o Outlook: {e}")
        return

    def normalizar(p):
        return p.replace("\\", "/").lower().strip("/")

    stores_antes = set()
    count_antes = 0
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
        pass

    try:
        outlook.AddStoreEx(pst_file, 3)  # 3 = olStoreUnicode
    except Exception as e:
        print(f"[ERRO] Nao foi possivel abrir o PST: {e}")
        return

    store = _find_pst_store(outlook, pst_file, count_antes, stores_antes)
    if not store:
        print("[ERRO] PST aberto mas nao encontrado nas stores do Outlook.")
        return

    print(f"[*] Store: {getattr(store, 'DisplayName', pst_file)}")

    # ── Fase 1: travessia iterativa (evita estouro de pilha em PSTs profundos) ──
    # Os objetos COM são STA — todo o acesso acontece nesta thread.
    queue = deque([store.GetRootFolder()])
    total_folders = 0

    while queue:
        folder = queue.popleft()

        try:
            folder_name = folder.Name
        except Exception:
            folder_name = "?"

        # Enfileira subpastas
        try:
            for i in range(1, folder.Folders.Count + 1):
                queue.append(folder.Folders.Item(i))
        except Exception:
            pass

        # Lê itens da pasta
        try:
            raw_items = folder.Items
            total_raw = raw_items.Count
        except Exception:
            continue

        if total_raw == 0:
            continue

        total_folders += 1
        print(f"[*] Pasta '{folder_name}': {total_raw} item(ns) total")

        # Melhoria principal: deixa o Outlook filtrar por data antes de iterar
        items = _pst_restrict_by_date(raw_items, start_dt, end_dt)

        try:
            count = items.Count
        except Exception:
            continue

        if count == 0:
            print(f"    Nenhum item no periodo em '{folder_name}'.")
            continue

        fora_remetente = set()

        for i in range(1, count + 1):
            try:
                msg = items.Item(i)
            except Exception as e:
                print(f"    [!] Nao foi possivel ler item {i} de '{folder_name}': {e}")
                continue

            try:
                if msg.Class != 43:  # 43 = olMail
                    continue

                sender_addr = (msg.SenderEmailAddress or "").lower()
                subject = msg.Subject or "(sem assunto)"

                try:
                    raw_dt = msg.ReceivedTime
                    msg_dt = datetime(raw_dt.year, raw_dt.month, raw_dt.day,
                                      raw_dt.hour, raw_dt.minute, raw_dt.second)
                except Exception:
                    print(f"    [!] Nao foi possivel ler data do item {i} de '{folder_name}'")
                    continue

                # Restrict já filtrou por data, mas valida por segurança
                if not (start_dt <= msg_dt <= end_dt):
                    continue

                if filter_emails and sender_addr not in filter_emails:
                    fora_remetente.add(sender_addr)
                    continue

                date_str = msg_dt.strftime("%d/%m/%Y")
                sender_path = email_paths.get(sender_addr, base_path)

                try:
                    output_dir = create_output_dir(sender_path, sender_addr or "desconhecido", date_str)
                except OSError as e:
                    print(f"    [!] Erro ao criar diretorio para '{sender_addr}': {e}")
                    continue

                print(f"    Processando: De={sender_addr} | Assunto={subject} | Data={date_str}")

                attachment_found = False
                att_count = 0
                try:
                    att_count = msg.Attachments.Count
                except Exception:
                    pass

                for j in range(1, att_count + 1):
                    att_name = "?"
                    try:
                        att = msg.Attachments.Item(j)
                        att_name = att.FileName or ""
                        if not att_name:
                            continue

                        ext = os.path.splitext(att_name)[1].lower()
                        if allowed_extensions and ext not in allowed_extensions:
                            relatorio.append({
                                "data": date_str, "remetente": sender_addr,
                                "assunto": subject, "arquivo": att_name,
                                "pasta_email": folder_name, "caminho": "",
                                "status": f"ignorado (extensao {ext} nao permitida)",
                            })
                            continue
                        if blocked_extensions and ext in blocked_extensions:
                            relatorio.append({
                                "data": date_str, "remetente": sender_addr,
                                "assunto": subject, "arquivo": att_name,
                                "pasta_email": folder_name, "caminho": "",
                                "status": f"ignorado (extensao {ext} bloqueada)",
                            })
                            continue

                        safe_name = sanitize_filename(att_name)
                        filepath = os.path.join(output_dir, safe_name)
                        ja_existia = os.path.exists(filepath)

                        if not ja_existia:
                            att.SaveAsFile(filepath)
                            print(f"    [+] Salvo: {safe_name}")
                            if auto_extract:
                                # Enfileira para extração paralela (sem COM)
                                pending_extractions.append((filepath, safe_name, output_dir))

                        attachment_found = True
                        relatorio.append({
                            "data": date_str, "remetente": sender_addr,
                            "assunto": subject, "arquivo": safe_name,
                            "pasta_email": folder_name, "caminho": output_dir,
                            "status": "ja existia" if ja_existia else "baixado",
                        })

                    except Exception as e:
                        print(f"    [!] Erro ao salvar anexo '{att_name}' (item {i}, att {j}): {e}")

                if not attachment_found:
                    relatorio.append({
                        "data": date_str, "remetente": sender_addr,
                        "assunto": subject, "arquivo": "(sem anexo valido)",
                        "pasta_email": folder_name, "caminho": "",
                        "status": "sem anexo",
                    })

                total_processados += 1

            except Exception as e:
                print(f"    [!] Erro inesperado no item {i} de '{folder_name}': {e}")
                continue

        if fora_remetente:
            print(f"    Remetentes no periodo mas fora do filtro: {sorted(fora_remetente)}")

    # Libera a store do Outlook
    try:
        outlook.RemoveStore(store.GetRootFolder())
    except Exception:
        pass

    # ── Fase 2: extração paralela (puro I/O de disco, sem COM) ──────────────
    if pending_extractions:
        print(f"[*] Extraindo {len(pending_extractions)} arquivo(s) com {max_workers} worker(s)...")
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(_extract_archive, fp, fn, od): fn
                for fp, fn, od in pending_extractions
            }
            for future in as_completed(futures):
                fn = futures[future]
                try:
                    future.result()
                except Exception as e:
                    print(f"[!] Erro na extracao de '{fn}': {e}")

    if total_processados == 0:
        print(
            f"[*] Nenhum e-mail encontrado no PST para o periodo "
            f"{config['date_range']['start']} a {config['date_range']['end']}."
        )
    else:
        print(f"[*] PST concluido. {total_processados} e-mail(s) em {total_folders} pasta(s).")

    write_audit_log(relatorio, base_path, prefix="relatorio_pst")


# ---------------------------------------------------------------------------
# Persistência da última execução
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def _make_date_entry(parent, value=""):
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
            raw = ent.get()
            cursor = ent.index("insert")
            digits = "".join(c for c in raw if c.isdigit())[:8]
            fmt = ""
            for i, d in enumerate(digits):
                if i in (2, 4):
                    fmt += "/"
                fmt += d
            if fmt == raw:
                return
            ent.delete(0, "end")
            ent.insert(0, fmt)
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
    widget.bind("<FocusIn>",
                lambda e, w=widget: w.after(1, lambda: w.select_range(0, "end")))


def show_gui(config):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    last = _load_last_run()
    result = {}
    PAD = 10

    root = tk.Tk()
    root.title("AutoMail")
    root.resizable(True, True)
    root.minsize(640, 480)
    root.attributes("-topmost", True)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)

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
            filetypes=[("Arquivo PST do Outlook", "*.pst"), ("Todos os arquivos", "*.*")],
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

    imap_server_var = imap_row(tab_imap, "Servidor:", "imap_server", "imap_server", 28)
    imap_port_var = imap_row(tab_imap, "Porta:", "imap_port", "imap_port", 6)
    imap_user_var = imap_row(tab_imap, "Usuario:", "email_user", "email_user", 34)
    imap_pass_var = imap_row(tab_imap, "Senha:", "email_pass", "email_pass", 34, show="•")

    unseen_val = last_imap.get("only_unseen", config.get("only_unseen", True))
    only_unseen_var = tk.BooleanVar(value=bool(unseen_val))
    tk.Checkbutton(tab_imap, text="Apenas e-mails nao lidos",
                   variable=only_unseen_var,
                   font=("Segoe UI", 9)).pack(anchor="w", pady=(6, 0))

    # ── PERÍODO ───────────────────────────────────────────────────────────
    frm_date = tk.LabelFrame(root, text="Periodo", font=("Segoe UI", 9),
                              padx=PAD, pady=6)
    frm_date.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(0, 4))

    last_dr = last.get("date_range", config.get("date_range", {}))
    tk.Label(frm_date, text="De:", font=("Segoe UI", 9)).pack(side="left")
    date_start = _make_date_entry(frm_date, last_dr.get("start", ""))
    date_start.pack(side="left", padx=(4, 20))
    tk.Label(frm_date, text="Ate:", font=("Segoe UI", 9)).pack(side="left")
    date_end = _make_date_entry(frm_date, last_dr.get("end", ""))
    date_end.pack(side="left", padx=(4, 0))

    # ── REMETENTES ────────────────────────────────────────────────────────
    frm_table = tk.LabelFrame(root, text="Remetentes e pastas de destino",
                               font=("Segoe UI", 9), padx=PAD, pady=6)
    frm_table.grid(row=2, column=0, sticky="nsew", padx=PAD, pady=(0, 4))
    frm_table.columnconfigure(0, weight=1)
    frm_table.rowconfigure(1, weight=1)

    frm_header = tk.Frame(frm_table)
    frm_header.grid(row=0, column=0, sticky="ew", padx=(0, 17))
    tk.Label(frm_header, text="E-mail do remetente",
             font=("Segoe UI", 9, "bold"), width=32, anchor="w").pack(side="left", padx=(0, 6))
    tk.Label(frm_header, text="Pasta de destino",
             font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left")

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
    _win = canvas.create_window((0, 0), window=inner, anchor="nw")

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
        path_var = tk.StringVar(value=path)

        ent_email = tk.Entry(frm, textvariable=email_var, width=32, font=("Segoe UI", 9))
        ent_email.pack(side="left", padx=(0, 6))
        _sel_all(ent_email)

        ent_path = tk.Entry(frm, textvariable=path_var, width=34, font=("Segoe UI", 9))
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

        tk.Button(frm, text="x", width=2, fg="red", font=("Segoe UI", 9),
                  command=_remove).pack(side="left")

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
    frm_opts = tk.LabelFrame(root, text="Opcoes", font=("Segoe UI", 9),
                              padx=PAD, pady=6)
    frm_opts.grid(row=3, column=0, sticky="ew", padx=PAD, pady=(0, 4))
    frm_opts.columnconfigure(0, weight=1)

    out_val = last.get("output_base_path", config.get("output_base_path", "./auditoria_arquivos"))
    out_var = tk.StringVar(value=out_val)
    row_out = tk.Frame(frm_opts)
    row_out.pack(fill="x", pady=2)
    tk.Label(row_out, text="Pasta padrao:", font=("Segoe UI", 9),
             width=16, anchor="w").pack(side="left")
    ent_out = tk.Entry(row_out, textvariable=out_var, width=46, font=("Segoe UI", 9))
    ent_out.pack(side="left", padx=(0, 4))
    _sel_all(ent_out)

    def browse_out():
        p = filedialog.askdirectory(title="Pasta padrao de saida")
        if p:
            out_var.set(p.replace("/", "\\"))

    tk.Button(row_out, text="…", width=2, command=browse_out,
              font=("Segoe UI", 9)).pack(side="left")

    auto_val = last.get("auto_extract", config.get("auto_extract", True))
    auto_var = tk.BooleanVar(value=bool(auto_val))
    tk.Checkbutton(frm_opts, text="Extrair arquivos compactados automaticamente",
                   variable=auto_var, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 2))

    row_ext = tk.Frame(frm_opts)
    row_ext.pack(fill="x", pady=2)
    tk.Label(row_ext, text="Extensoes permitidas:",
             font=("Segoe UI", 9)).pack(side="left")
    tk.Label(row_ext, text="(vazio = todas)",
             font=("Segoe UI", 8), fg="#888").pack(side="left", padx=(4, 8))
    exts_raw = last.get("allowed_extensions", config.get("allowed_extensions", []))
    exts_str = ", ".join(exts_raw) if isinstance(exts_raw, list) else str(exts_raw)
    exts_var = tk.StringVar(value=exts_str)
    ent_exts = tk.Entry(row_ext, textvariable=exts_var, width=36, font=("Segoe UI", 9))
    ent_exts.pack(side="left")
    _sel_all(ent_exts)

    row_blk = tk.Frame(frm_opts)
    row_blk.pack(fill="x", pady=2)
    tk.Label(row_blk, text="Extensoes bloqueadas:",
             font=("Segoe UI", 9)).pack(side="left")
    tk.Label(row_blk, text="(vazio = nenhuma)",
             font=("Segoe UI", 8), fg="#888").pack(side="left", padx=(4, 8))
    blk_raw = last.get("blocked_extensions", config.get("blocked_extensions", []))
    blk_str = ", ".join(blk_raw) if isinstance(blk_raw, list) else str(blk_raw)
    blk_var = tk.StringVar(value=blk_str)
    ent_blk = tk.Entry(row_blk, textvariable=blk_var, width=36, font=("Segoe UI", 9))
    ent_blk.pack(side="left")
    _sel_all(ent_blk)

    # ── BOTÕES ────────────────────────────────────────────────────────────
    frm_btn = tk.Frame(root)
    frm_btn.grid(row=4, column=0, pady=(0, PAD))

    def on_ok():
        start = date_start.get()
        end = date_end.get()
        try:
            datetime.strptime(start, "%d/%m/%Y")
            datetime.strptime(end, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Datas invalidas. Use DD/MM/AAAA.")
            return

        default_out = out_var.get().strip() or "./auditoria_arquivos"
        entries = []
        for email_var, path_var, _ in rows:
            e = email_var.get().strip().lower()
            p = path_var.get().strip() or default_out
            if e:
                entries.append({"email": e, "output_path": p})

        exts_input = exts_var.get().strip()
        allowed_exts = (
            [x.strip().lower() for x in exts_input.split(",") if x.strip()]
            if exts_input else []
        )

        blk_input = blk_var.get().strip()
        blocked_exts = (
            [x.strip().lower() for x in blk_input.split(",") if x.strip()]
            if blk_input else []
        )

        if allowed_exts and blocked_exts:
            messagebox.showerror(
                "Erro",
                "Preencha apenas 'Extensoes permitidas' OU 'Extensoes bloqueadas', nao os dois."
            )
            return

        tab_idx = notebook.index(notebook.select())
        mode = "pst" if tab_idx == 0 else "imap"

        if mode == "pst":
            pst = pst_var.get().strip()
            if not pst:
                messagebox.showerror("Erro", "Selecione um arquivo PST.")
                return
            result["pst_file"] = pst
        else:
            srv = imap_server_var.get().strip()
            port = imap_port_var.get().strip()
            usr = imap_user_var.get().strip()
            pwd = imap_pass_var.get().strip()
            if not all([srv, port, usr, pwd]):
                messagebox.showerror("Erro", "Preencha todos os campos IMAP.")
                return
            try:
                port_int = int(port)
            except ValueError:
                messagebox.showerror("Erro", "Porta IMAP invalida.")
                return
            result.update({
                "imap_server": srv,
                "imap_port": port_int,
                "email_user": usr,
                "email_pass": pwd,
                "only_unseen": only_unseen_var.get(),
            })

        result.update({
            "mode": mode,
            "date_range": {"start": start, "end": end},
            "filter_emails": entries,
            "output_base_path": default_out,
            "auto_extract": auto_var.get(),
            "allowed_extensions": allowed_exts,
            "blocked_extensions": blocked_exts,
        })

        to_save = {k: v for k, v in result.items() if k != "email_pass"}
        if mode == "imap":
            to_save["imap"] = {
                k: result[k]
                for k in ("imap_server", "imap_port", "email_user", "only_unseen")
                if k in result
            }
        _save_last_run(to_save)
        root.destroy()

    def on_cancel():
        root.destroy()

    tk.Button(frm_btn, text="Iniciar", width=12, font=("Segoe UI", 9, "bold"),
              bg="#0078D4", fg="white", command=on_ok).pack(side="left", padx=6)
    tk.Button(frm_btn, text="Cancelar", width=10, font=("Segoe UI", 9),
              command=on_cancel).pack(side="left", padx=6)

    last_mode = last.get("mode", "pst" if config.get("pst_file", "") else "imap")
    notebook.select(0 if last_mode == "pst" else 1)

    root.mainloop()
    return result


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main():
    config = load_config()

    cfg = show_gui(config)
    if not cfg:
        print("[*] Cancelado pelo usuario.")
        return

    config.update(cfg)

    if cfg.get("mode") == "imap":
        process_imap(config)
    else:
        process_pst(config)


if __name__ == "__main__":
    main()
