"""
Posta2Excel - Backend: estrazione contatti da PST/OST e export Excel.
"""

from __future__ import annotations

import html
import os
import queue
import re
import traceback
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

import pandas as pd
import pypff
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

GENERIC_DOMAIN_BLACKLIST = {
    "gmail",
    "outlook",
    "hotmail",
    "yahoo",
    "live",
    "icloud",
    "libero",
    "virgilio",
    "alice",
    "tiscali",
    "fastwebnet",
    "poste",
    "me",
    "msn",
    "ymail",
    "aol",
}

ROBOT_PREFIX_BLACKLIST = {
    "noreply",
    "no-reply",
    "newsletter",
    "notifiche",
    "notifications",
    "bounce",
    "delivery",
    "mailer-daemon",
}

GENERIC_NAME_WORDS = {
    "info",
    "ufficio",
    "commerciale",
    "vendite",
    "amministrazione",
    "supporto",
}

PHONE_KEYWORDS = (
    "tel",
    "cell",
    "mob",
    "phone",
    "t:",
    "m:",
    "ufficio",
    "whatsapp",
    "telefono",
)

PHONE_REGEX = re.compile(
    r"(?:(?:\+|00)39)?\s?(?:[0-9]{2,4})[\s.-]?[0-9]{6,8}",
    re.IGNORECASE,
)

COMPANY_REGEX = re.compile(
    r"\b[A-Za-z0-9\s.&-]+\s(?:S\.r\.l\.|Srl|S\.p\.A\.|Spa|S\.A\.S\.|Sas|S\.n\.c\.|GmbH)\b",
    re.IGNORECASE,
)

EMAIL_IN_TEXT_REGEX = re.compile(
    r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
    re.IGNORECASE,
)

EMAIL_HEADER_REGEX = re.compile(
    r"(?:^|[\s,;])(?:\"?([^\"<>;,]+)\"?\s*)?<([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})>",
    re.IGNORECASE,
)

EMAIL_STANDALONE_REGEX = re.compile(
    r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
    re.IGNORECASE,
)

THREAD_MARKERS = (
    re.compile(r"^\s*-----Original Message-----\s*$", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^\s*---\s*Inizio messaggio inoltrato\s*---\s*$", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^\s*---\s*Forwarded message\s*---\s*$", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^\s*Da:\s", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^\s*From:\s", re.IGNORECASE | re.MULTILINE),
    re.compile(r"^\s*>", re.MULTILINE),
)

HTML_TAG_REGEX = re.compile(r"<[^>]+>")

ERROR_LOG = "posta2excel_errors.log"

HEADER_FILL = PatternFill(start_color="2F3542", end_color="2F3542", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
DATA_ALIGNMENT = Alignment(vertical="top", wrap_text=True)

ContactRecord = Dict[str, Any]
ContactsDict = Dict[str, ContactRecord]


def _safe_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value).strip()


def _normalize_email(email: str) -> str:
    return _safe_str(email).lower()


def _split_name(display_name: str, email: str = "") -> Tuple[str, str]:
    name = _clean_outlook_name(display_name, email)
    if not name:
        return "", ""
    if " " not in name:
        return name, ""
    parts = name.split(None, 1)
    return parts[0], parts[1] if len(parts) > 1 else ""


def _clean_outlook_name(name: str, email: str = "") -> str:
    cleaned = _safe_str(name)
    if not cleaned:
        return ""
    cleaned = re.sub(r"^SMTP:\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip("\"'")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    email_norm = _normalize_email(email)
    if email_norm and cleaned.lower() == email_norm:
        return ""
    if EMAIL_STANDALONE_REGEX.fullmatch(cleaned):
        return ""
    return cleaned


def _is_generic_name(nome: str, cognome: str) -> bool:
    combined = f"{nome} {cognome}".strip().lower()
    if not combined:
        return True
    tokens = re.split(r"[\s|/\\-]+", combined)
    return any(token in GENERIC_NAME_WORDS for token in tokens)


def _is_better_name(
    current_nome: str,
    current_cognome: str,
    new_nome: str,
    new_cognome: str,
) -> bool:
    if not new_nome and not new_cognome:
        return False
    if _is_generic_name(new_nome, new_cognome):
        return False
    if not current_nome and not current_cognome:
        return True
    if _is_generic_name(current_nome, current_cognome):
        return True
    current_len = len(current_nome) + len(current_cognome)
    new_len = len(new_nome) + len(new_cognome)
    if new_nome and new_cognome and not (current_nome and current_cognome):
        return True
    return new_len > current_len


def _domain_to_company(email: str) -> str:
    email_norm = _normalize_email(email)
    if "@" not in email_norm:
        return ""
    domain = email_norm.split("@", 1)[1]
    domain_root = domain.split(".", 1)[0].lower()
    if domain_root in GENERIC_DOMAIN_BLACKLIST:
        return ""
    label = domain_root.replace("-", " ").replace("_", " ")
    if not label:
        return ""
    return label[0].upper() + label[1:]


def _apply_generic_name_fallback(
    nome: str,
    cognome: str,
    email: str,
) -> Tuple[str, str, str]:
    societa = ""
    if not _is_generic_name(nome, cognome) and (nome or cognome):
        return nome, cognome, societa
    company = _domain_to_company(email)
    if not company:
        return nome, cognome, societa
    if not nome and not cognome:
        return company, "", company
    if _is_generic_name(nome, cognome):
        return company, cognome if cognome and not _is_generic_name("", cognome) else "", company
    return nome, cognome, societa


def _is_robot_email(email: str) -> bool:
    email_norm = _normalize_email(email)
    if "@" not in email_norm:
        return True
    local_part = email_norm.split("@", 1)[0]
    return local_part in ROBOT_PREFIX_BLACKLIST


def _extract_emails_from_header(value: str) -> List[Tuple[str, str]]:
    results: List[Tuple[str, str]] = []
    seen: Set[str] = set()
    if not value:
        return results
    for match in EMAIL_HEADER_REGEX.finditer(value):
        display_name = _clean_outlook_name(match.group(1) or "", match.group(2))
        email = _normalize_email(match.group(2))
        if email and email not in seen:
            seen.add(email)
            results.append((display_name, email))
    if not results:
        for match in EMAIL_STANDALONE_REGEX.finditer(value):
            email = _normalize_email(match.group(0))
            if email and email not in seen:
                seen.add(email)
                results.append(("", email))
    return results


def _parse_transport_headers(headers: str) -> Dict[str, str]:
    parsed: Dict[str, str] = {}
    if not headers:
        return parsed
    for line in headers.replace("\r\n", "\n").split("\n"):
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key_norm = key.strip().lower()
        val = val.strip()
        if not val:
            continue
        if key_norm in parsed:
            parsed[key_norm] = f"{parsed[key_norm]}, {val}"
        else:
            parsed[key_norm] = val
    return parsed


def _get_message_body(msg: Any) -> str:
    plain = getattr(msg, "plain_text_body", None)
    if plain is None and hasattr(msg, "get_plain_text_body"):
        plain = msg.get_plain_text_body()
    plain_text = _safe_str(plain)
    if plain_text:
        return plain_text
    html_body = getattr(msg, "html_body", None)
    if html_body is None and hasattr(msg, "get_html_body"):
        html_body = msg.get_html_body()
    html_text = _safe_str(html_body)
    if not html_text:
        return ""
    html_text = HTML_TAG_REGEX.sub(" ", html_text)
    html_text = html.unescape(html_text)
    html_text = re.sub(r"[ \t]+", " ", html_text)
    html_text = re.sub(r"\n\s*\n+", "\n\n", html_text)
    return html_text.strip()


def _isolate_latest_message_block(body: str) -> str:
    if not body:
        return ""
    earliest_index: Optional[int] = None
    for pattern in THREAD_MARKERS:
        match = pattern.search(body)
        if match:
            index = match.start()
            if earliest_index is None or index < earliest_index:
                earliest_index = index
    if earliest_index is None:
        return body
    return body[:earliest_index].strip()


def _last_non_empty_lines(text: str, count: int = 15) -> List[str]:
    lines = [line.strip() for line in text.replace("\r\n", "\n").split("\n")]
    non_empty = [line for line in lines if line]
    return non_empty[-count:]


def _line_is_email_free(line: str) -> bool:
    return "@" not in line and EMAIL_IN_TEXT_REGEX.search(line) is None


def _phone_near_keyword(lines: List[str], line_index: int) -> bool:
    indices = {line_index}
    if line_index > 0:
        indices.add(line_index - 1)
    if line_index + 1 < len(lines):
        indices.add(line_index + 1)
    for idx in indices:
        lower_line = lines[idx].lower()
        if any(keyword in lower_line for keyword in PHONE_KEYWORDS):
            return True
    return False


def _extract_phones_from_lines(lines: List[str]) -> Set[str]:
    phones: Set[str] = set()
    for index, line in enumerate(lines):
        for match in PHONE_REGEX.finditer(line):
            if _phone_near_keyword(lines, index):
                phones.add(match.group(0).strip())
    return phones


def _extract_company_from_lines(lines: List[str]) -> str:
    for line in lines:
        match = COMPANY_REGEX.search(line)
        if match:
            return match.group(0).strip()
    return ""


def _parse_delivery_time(msg: Any) -> Optional[datetime]:
    delivery = getattr(msg, "delivery_time", None)
    if delivery is None and hasattr(msg, "get_delivery_time"):
        delivery = msg.get_delivery_time()
    if delivery is None:
        delivery = getattr(msg, "client_submit_time", None)
    if delivery is None and hasattr(msg, "get_client_submit_time"):
        delivery = msg.get_client_submit_time()
    if delivery is None:
        return None
    if isinstance(delivery, datetime):
        if delivery.tzinfo is None:
            return delivery.replace(tzinfo=timezone.utc)
        return delivery
    return None


def _log_message_error(
    exc: Exception,
    folder_name: str = "",
    index: int = -1,
    msg: Any = None,
) -> None:
    subject = ""
    if msg is not None:
        try:
            subject = _safe_str(getattr(msg, "subject", ""))
        except Exception:
            subject = ""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = (
        f"{timestamp} | Folder={folder_name} | Index={index} | "
        f"Subject={subject} | Error={exc}\n{traceback.format_exc()}\n"
    )
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as log_file:
            log_file.write(entry)
    except Exception:
        pass


def _join_unique(series: pd.Series) -> str:
    values: Set[str] = set()
    for value in series:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not text:
            continue
        for part in text.split("\n"):
            part = part.strip()
            if part:
                values.add(part)
    return "\n".join(sorted(values))


class ContactExtractor:
    def __init__(self) -> None:
        self.contatti: ContactsDict = {}
        self._processed_messages = 0
        self._status_callback: Optional[Callable[[str], None]] = None
        self._progress_callback: Optional[Callable[[float], None]] = None

    def set_callbacks(
        self,
        status_callback: Optional[Callable[[str], None]] = None,
        progress_callback: Optional[Callable[[float], None]] = None,
    ) -> None:
        self._status_callback = status_callback
        self._progress_callback = progress_callback

    def _emit_status(self, text: str) -> None:
        if self._status_callback:
            self._status_callback(text)

    def _emit_progress(self, value: float) -> None:
        if self._progress_callback:
            self._progress_callback(max(0.0, min(1.0, value)))

    def _upsert_contact(
        self,
        email: str,
        nome: str = "",
        cognome: str = "",
        telefoni: Optional[Set[str]] = None,
        societa: str = "",
        delivery_time: Optional[datetime] = None,
        from_metadata: bool = True,
    ) -> None:
        email_norm = _normalize_email(email)
        if not email_norm or _is_robot_email(email_norm):
            return
        if not from_metadata:
            if email_norm not in self.contatti:
                return

        nome, cognome, fallback_societa = _apply_generic_name_fallback(nome, cognome, email_norm)

        if email_norm not in self.contatti:
            if not from_metadata:
                return
            self.contatti[email_norm] = {
                "email": email_norm,
                "nome": nome,
                "cognome": cognome,
                "telefoni": set(),
                "societa": "",
                "societa_updated_at": None,
            }

        record = self.contatti[email_norm]
        if _is_better_name(record["nome"], record["cognome"], nome, cognome):
            record["nome"] = nome
            record["cognome"] = cognome

        if telefoni:
            record["telefoni"].update(telefoni)

        societa_candidate = societa or fallback_societa or _domain_to_company(email_norm)
        if societa_candidate:
            current_ts = record.get("societa_updated_at")
            if delivery_time is None:
                if not record.get("societa"):
                    record["societa"] = societa_candidate
            elif current_ts is None or delivery_time >= current_ts:
                record["societa"] = societa_candidate
                record["societa_updated_at"] = delivery_time

    def _extract_sender(self, msg: Any, headers: Dict[str, str]) -> Tuple[str, str]:
        email = ""
        display_name = _safe_str(getattr(msg, "sender_name", ""))

        sender_address = getattr(msg, "sender_delivery_address", None)
        if sender_address:
            email = _normalize_email(sender_address)

        if not email:
            for key in ("from", "sender"):
                if key in headers:
                    extracted = _extract_emails_from_header(headers[key])
                    if extracted:
                        if not display_name:
                            display_name = extracted[0][0]
                        email = extracted[0][1]
                        break

        if not email:
            match = EMAIL_STANDALONE_REGEX.search(display_name)
            if match:
                email = _normalize_email(match.group(0))

        display_name = _clean_outlook_name(display_name, email)
        return display_name, email

    def _extract_recipients(self, msg: Any, headers: Dict[str, str]) -> List[Tuple[str, str]]:
        recipients: List[Tuple[str, str]] = []
        seen: Set[str] = set()

        if hasattr(msg, "get_number_of_recipients") and hasattr(msg, "get_recipient"):
            try:
                count = msg.get_number_of_recipients()
                for index in range(count):
                    recipient = msg.get_recipient(index)
                    display_name = _safe_str(
                        getattr(recipient, "name", "")
                        or getattr(recipient, "display_name", "")
                    )
                    email = _normalize_email(
                        getattr(recipient, "email_address", "")
                        or getattr(recipient, "delivery_address", "")
                        or getattr(recipient, "address", "")
                    )
                    if email and email not in seen:
                        seen.add(email)
                        recipients.append((_clean_outlook_name(display_name, email), email))
            except Exception:
                pass

        for key in ("to", "cc", "bcc"):
            if key in headers:
                for display_name, email in _extract_emails_from_header(headers[key]):
                    if email not in seen:
                        seen.add(email)
                        recipients.append((_clean_outlook_name(display_name, email), email))

        return recipients

    def _parse_signature_for_enrichment(
        self,
        body: str,
        sender_email: str,
        delivery_time: Optional[datetime],
    ) -> None:
        sender_email = _normalize_email(sender_email)
        if not sender_email or sender_email not in self.contatti:
            return
        block = _isolate_latest_message_block(body)
        lines = [
            line
            for line in _last_non_empty_lines(block, 15)
            if _line_is_email_free(line)
        ]
        phones = _extract_phones_from_lines(lines)
        company = _extract_company_from_lines(lines)
        self._upsert_contact(
            email=sender_email,
            telefoni=phones,
            societa=company,
            delivery_time=delivery_time,
            from_metadata=False,
        )

    def _process_single_message(
        self,
        msg: Any,
        folder_name: str,
        message_index: int,
    ) -> None:
        self._processed_messages += 1
        if self._processed_messages % 25 == 0:
            pseudo_progress = min(0.95, self._processed_messages / max(self._processed_messages + 500, 1))
            self._emit_progress(pseudo_progress)
        self._emit_status(f"Stato: Scansione {folder_name} (Mail {message_index})...")

        headers_text = _safe_str(getattr(msg, "transport_headers", ""))
        headers = _parse_transport_headers(headers_text)
        delivery_time = _parse_delivery_time(msg)

        sender_name, sender_email = self._extract_sender(msg, headers)
        if sender_email:
            nome, cognome = _split_name(sender_name, sender_email)
            self._upsert_contact(
                email=sender_email,
                nome=nome,
                cognome=cognome,
                delivery_time=delivery_time,
                from_metadata=True,
            )

        for display_name, email in self._extract_recipients(msg, headers):
            nome, cognome = _split_name(display_name, email)
            self._upsert_contact(
                email=email,
                nome=nome,
                cognome=cognome,
                delivery_time=delivery_time,
                from_metadata=True,
            )

        body = _get_message_body(msg)
        if body and sender_email:
            self._parse_signature_for_enrichment(body, sender_email, delivery_time)

    def _traverse_folder(self, folder: Any, folder_name: str = "") -> None:
        current_name = folder_name or _safe_str(
            folder.get_name() if hasattr(folder, "get_name") else getattr(folder, "name", "")
        ) or "Cartella"

        try:
            message_count = folder.get_number_of_sub_messages()
        except Exception as exc:
            _log_message_error(exc, folder_name=current_name)
            message_count = 0

        for index in range(message_count):
            msg = None
            try:
                msg = folder.get_sub_message(index)
                self._process_single_message(msg, current_name, index + 1)
            except Exception as exc:
                _log_message_error(exc, folder_name=current_name, index=index, msg=msg)
            finally:
                del msg

        try:
            subfolder_count = folder.get_number_of_sub_folders()
        except Exception as exc:
            _log_message_error(exc, folder_name=current_name)
            subfolder_count = 0

        for sub_index in range(subfolder_count):
            sub = None
            try:
                sub = folder.get_sub_folder(sub_index)
                sub_name = _safe_str(
                    sub.get_name() if hasattr(sub, "get_name") else getattr(sub, "name", "")
                )
                next_name = f"{current_name}/{sub_name}" if sub_name else current_name
                self._traverse_folder(sub, next_name)
            except Exception as exc:
                _log_message_error(exc, folder_name=current_name, index=sub_index)
            finally:
                if sub is not None:
                    del sub

    def extract_from_pst(self, pst_path: str) -> None:
        pst_file = pypff.file()
        pst_file.open(pst_path)
        try:
            root = pst_file.get_root_folder()
            self._emit_status("Stato: Analisi archivio in corso...")
            self._traverse_folder(root)
        finally:
            pst_file.close()


def _records_to_rows(contatti: ContactsDict) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for record in contatti.values():
        rows.append(
            {
                "Nome": record.get("nome", ""),
                "Cognome": record.get("cognome", ""),
                "Email": record.get("email", ""),
                "Telefono": "\n".join(sorted(record.get("telefoni", set()))),
                "Società": record.get("societa", ""),
            }
        )
    return rows


def _build_export_dataframe(contatti: ContactsDict) -> pd.DataFrame:
    if not contatti:
        return pd.DataFrame(columns=["Nome", "Cognome", "Email", "Telefono", "Società"])

    df = pd.DataFrame(_records_to_rows(contatti))
    named_mask = df["Nome"].astype(str).str.strip() != ""
    named_mask &= df["Cognome"].astype(str).str.strip() != ""

    df_named = df[named_mask].copy()
    df_anon = df[~named_mask].copy()

    if not df_named.empty:
        df_grouped = df_named.groupby(["Nome", "Cognome"], as_index=False).agg(
            {
                "Email": _join_unique,
                "Telefono": _join_unique,
                "Società": _join_unique,
            }
        )
    else:
        df_grouped = pd.DataFrame(columns=["Nome", "Cognome", "Email", "Telefono", "Società"])

    result = pd.concat([df_grouped, df_anon], ignore_index=True)
    sort_cols = [col for col in ("Cognome", "Nome") if col in result.columns]
    if sort_cols:
        result = result.sort_values(by=sort_cols, kind="stable")
    return result.reset_index(drop=True)


def _autofit_columns(worksheet: Any, dataframe: pd.DataFrame) -> None:
    for column_index, column_name in enumerate(dataframe.columns, start=1):
        max_length = len(str(column_name))
        column_series = dataframe[column_name].astype(str)
        for value in column_series:
            if value == "nan":
                continue
            for segment in value.split("\n"):
                max_length = max(max_length, len(segment))
        adjusted_width = max_length + 3
        worksheet.column_dimensions[get_column_letter(column_index)].width = adjusted_width


def _style_worksheet(worksheet: Any, dataframe: pd.DataFrame) -> None:
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGNMENT

    worksheet.views.sheetView[0].showGridLines = True

    for row in worksheet.iter_rows(
        min_row=2,
        max_row=worksheet.max_row,
        min_col=1,
        max_col=worksheet.max_column,
    ):
        for cell in row:
            cell.alignment = DATA_ALIGNMENT

    _autofit_columns(worksheet, dataframe)


def export_to_excel(contatti: ContactsDict, output_path: str) -> None:
    dataframe = _build_export_dataframe(contatti)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Rubrica")
        worksheet = writer.sheets["Rubrica"]
        _style_worksheet(worksheet, dataframe)


def run_extraction(pst_path: str, output_dir: str, event_queue: queue.Queue) -> None:
    def push(event_type: str, **payload: Any) -> None:
        event_queue.put({"type": event_type, **payload})

    def status_callback(text: str) -> None:
        push("status", text=text)

    def progress_callback(value: float) -> None:
        push("progress", value=value)

    try:
        if not os.path.isfile(pst_path):
            raise FileNotFoundError("Il file PST/OST selezionato non esiste.")
        if not os.path.isdir(output_dir):
            raise FileNotFoundError("La cartella di destinazione non esiste.")

        extractor = ContactExtractor()
        extractor.set_callbacks(status_callback, progress_callback)

        push("status", text="Stato: Apertura archivio PST/OST...")
        push("progress", value=0.0)
        extractor.extract_from_pst(pst_path)

        push("status", text="Stato: Generazione Report Excel...")
        push("progress", value=0.97)
        output_path = os.path.join(output_dir, "rubrica_estratta.xlsx")
        export_to_excel(extractor.contatti, output_path)

        push("progress", value=1.0)
        push("status", text="Stato: Processo Completato con successo!")
        push("done")
    except Exception as exc:
        _log_message_error(exc)
        push("error", text=f"Errore: {exc}")
        push("progress", value=0.0)
