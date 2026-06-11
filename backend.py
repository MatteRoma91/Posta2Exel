"""
Posta2Excel - Backend: estrazione contatti da PST/OST/MBOX e export Excel.
"""

from __future__ import annotations

import html
import mailbox
import os
import queue
import re
import traceback
import unicodedata
from datetime import datetime, timezone
from email.header import decode_header
from email.utils import parsedate_to_datetime
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

import pandas as pd
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
    "donotreply",
    "do-not-reply",
    "postmaster",
    "announce",
    "reminder",
    "feedback",
    "preventivo",
    "comunicazioni",
    "news",
    "hello",
    "team",
    "author",
    "push",
}

ROLE_LOCAL_PARTS = {
    "info",
    "ufficio",
    "commerciale",
    "vendite",
    "amministrazione",
    "supporto",
}

GENERIC_NAME_WORDS = ROLE_LOCAL_PARTS

NAME_PARTICLES = {
    "di",
    "de",
    "del",
    "della",
    "degli",
    "delle",
    "da",
    "van",
    "von",
    "le",
    "la",
    "du",
    "den",
}

ORG_NAME_REJECT_WORDS = {
    "team",
    "servizio",
    "servizi",
    "notifica",
    "store",
    "communications",
    "marketplace",
    "fresh",
    "music",
    "calendar",
    "cloud",
    "studio",
    "certified",
    "beyond",
    "communications",
    "recensioni",
    "prenota",
    "antigravity",
    "assicurazioni",
    "banca",
    "scadenzabollo",
    "clienti",
}

AUTOMATED_DISPLAY_KEYWORDS = (
    "amazon",
    "google calendar",
    "google ai",
    "google antigravity",
    "paypal",
    "call of duty",
    "privalia",
    "myheritage",
    "mondo convenienza",
    "ita airways",
    "lego certified",
    "protezione civile",
    "facile.it",
    "segugio",
    "cooltra",
    "d&d beyond",
    "ovhcloud",
    "aruba.it",
    "github",
    "mailjet",
    "brevosend",
    "meetingdash",
    "percassi",
    "teatro sistina",
    "as roma store",
    "ollama",
)

AUTOMATED_LOCAL_PART_PATTERNS = (
    re.compile(r"^conferma-", re.IGNORECASE),
    re.compile(r"promotion", re.IGNORECASE),
    re.compile(r"shipment", re.IGNORECASE),
    re.compile(r"order-update", re.IGNORECASE),
    re.compile(r"account-update", re.IGNORECASE),
    re.compile(r"payments-update", re.IGNORECASE),
    re.compile(r"marketplace-", re.IGNORECASE),
    re.compile(r"customer-reviews", re.IGNORECASE),
    re.compile(r"calendar-notification", re.IGNORECASE),
    re.compile(r"^unsub-", re.IGNORECASE),
    re.compile(r"-noreply$", re.IGNORECASE),
    re.compile(r"^no_reply", re.IGNORECASE),
    re.compile(r"^google", re.IGNORECASE),
    re.compile(r"^reserve-noreply", re.IGNORECASE),
    re.compile(r"^announce", re.IGNORECASE),
    re.compile(r"^reminder", re.IGNORECASE),
    re.compile(r"^preventivo", re.IGNORECASE),
    re.compile(r"^ricerca", re.IGNORECASE),
    re.compile(r"^comunicazioni", re.IGNORECASE),
    re.compile(r"^pickup-point", re.IGNORECASE),
    re.compile(r"^shipment-tracking", re.IGNORECASE),
    re.compile(r"^order-update", re.IGNORECASE),
    re.compile(r"^payments-update", re.IGNORECASE),
    re.compile(r"^amazon-offers", re.IGNORECASE),
    re.compile(r"^notice@", re.IGNORECASE),
    re.compile(r"^news@", re.IGNORECASE),
    re.compile(r"^hello@", re.IGNORECASE),
    re.compile(r"^vendite@", re.IGNORECASE),
)

AUTOMATED_DOMAIN_PATTERNS = (
    re.compile(r"^noreply\.", re.IGNORECASE),
    re.compile(r"\.mailjet\.com$", re.IGNORECASE),
    re.compile(r"^bnc\d+\.mailjet\.com$", re.IGNORECASE),
    re.compile(r"\.brevosend\.com$", re.IGNORECASE),
    re.compile(r"\.meetingdash\.com$", re.IGNORECASE),
    re.compile(r"^emails\.", re.IGNORECASE),
    re.compile(r"^remind\.", re.IGNORECASE),
    re.compile(r"^enews\.", re.IGNORECASE),
    re.compile(r"^from\.", re.IGNORECASE),
    re.compile(r"^timetoride\.", re.IGNORECASE),
    re.compile(r"^update\.", re.IGNORECASE),
    re.compile(r"^staff\.", re.IGNORECASE),
    re.compile(r"^communications\.", re.IGNORECASE),
    re.compile(r"^nl-it\.", re.IGNORECASE),
    re.compile(r"^comms\.", re.IGNORECASE),
    re.compile(r"^news\.", re.IGNORECASE),
)

SPAM_LOCAL_PART_REGEX = re.compile(
    r"^[a-z]+[0-9]{2,}[a-z0-9]*$|^[a-z]{2,}[0-9]{4,}[a-z0-9]*$",
    re.IGNORECASE,
)

SUBJECT_LINE_MARKERS = re.compile(r"[|!]|SHE[I1]N|misterios|archiviazione|scato", re.IGNORECASE)

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

PYPFF_ERROR_MESSAGE = (
    "Per analizzare i file di Outlook è necessario installare la libreria aggiuntiva (pypff). "
    "Esegui: pip install libpff-python  (Windows: pip install libpff-python-windows)"
)

HEADER_FILL = PatternFill(start_color="2F3542", end_color="2F3542", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
DATA_ALIGNMENT = Alignment(vertical="top", wrap_text=True)

ContactRecord = Dict[str, Any]
ContactsDict = Dict[str, ContactRecord]


def _import_pypff() -> Any:
    try:
        import pypff
        return pypff
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(PYPFF_ERROR_MESSAGE) from exc


def _safe_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value).strip()


def _normalize_email(email: str) -> str:
    return _safe_str(email).lower()


def _normalize_text_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", _safe_str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return text.casefold().strip()


def _email_parts(email: str) -> Tuple[str, str]:
    email_norm = _normalize_email(email)
    if "@" not in email_norm:
        return "", ""
    local_part, domain = email_norm.split("@", 1)
    return local_part, domain


def _is_consumer_domain(domain: str) -> bool:
    domain = domain.lower()
    root = domain.split(".", 1)[0]
    return root in GENERIC_DOMAIN_BLACKLIST


def _business_domain_key(email: str) -> str:
    _, domain = _email_parts(email)
    return domain


def _is_automated_local_part(local_part: str) -> bool:
    if not local_part:
        return True
    base = local_part.split("+", 1)[0]
    if base in ROBOT_PREFIX_BLACKLIST:
        return True
    if "noreply" in base or "no-reply" in base or "donotreply" in base:
        return True
    for prefix in ROBOT_PREFIX_BLACKLIST:
        if base.startswith(f"{prefix}-") or base.startswith(f"{prefix}."):
            return True
    return any(pattern.search(base) for pattern in AUTOMATED_LOCAL_PART_PATTERNS)


def _is_automated_domain(domain: str) -> bool:
    if not domain:
        return True
    return any(pattern.search(domain) for pattern in AUTOMATED_DOMAIN_PATTERNS)


def _is_automated_display_name(display_name: str) -> bool:
    if not display_name:
        return False
    lower_name = display_name.lower()
    if any(keyword in lower_name for keyword in AUTOMATED_DISPLAY_KEYWORDS):
        return True
    if re.search(r"\b(?:su|on)\s+amazon\b", lower_name):
        return True
    if re.search(r"\.(?:it|com|eu|org|net)\b", lower_name):
        return True
    if " - " in display_name and not _is_plausible_person_name(display_name):
        return True
    return False


def _looks_like_spam_local_part(local_part: str) -> bool:
    base = local_part.split("+", 1)[0]
    if len(base) < 8:
        return False
    if SPAM_LOCAL_PART_REGEX.match(base):
        return True
    digit_count = sum(char.isdigit() for char in base)
    return digit_count >= 4 and digit_count / max(len(base), 1) > 0.25


def _looks_like_subject_line(display_name: str) -> bool:
    if not display_name:
        return False
    if SUBJECT_LINE_MARKERS.search(display_name):
        return True
    words = display_name.split()
    if len(words) > 5:
        return True
    return False


def _is_plausible_person_name(display_name: str) -> bool:
    name = _safe_str(display_name)
    if not name:
        return False
    if _looks_like_subject_line(name):
        return False
    if _is_automated_display_name(name):
        return False
    lower_name = name.lower()
    if any(word in lower_name for word in ORG_NAME_REJECT_WORDS):
        return False
    if re.search(r"[0-9/@#]", name):
        return False
    words = name.split()
    if not words or len(words) > 5:
        return False
    alpha_words = 0
    for word in words:
        letters = re.sub(r"[^A-Za-zÀ-ÿ'-]", "", word)
        if len(letters) >= 2:
            alpha_words += 1
    return alpha_words >= 1 and (alpha_words >= 2 or len(words[0]) >= 3)


def _classify_contact(email: str, display_name: str) -> str:
    email_norm = _normalize_email(email)
    if not email_norm or "@" not in email_norm:
        return "spam"
    local_part, domain = _email_parts(email_norm)
    cleaned_name = _clean_outlook_name(display_name, email_norm)
    local_base = local_part.split("+", 1)[0]

    if local_base in ROLE_LOCAL_PARTS and not _is_consumer_domain(domain):
        if not _is_automated_domain(domain):
            return "role"

    if (
        _is_automated_local_part(local_part)
        or _is_automated_domain(domain)
        or _is_automated_display_name(cleaned_name)
    ):
        return "automated"

    if _is_consumer_domain(domain):
        if _looks_like_spam_local_part(local_part):
            return "spam"
        if not cleaned_name or not _is_plausible_person_name(cleaned_name):
            return "spam"

    if cleaned_name and _is_plausible_person_name(cleaned_name):
        return "person"

    if local_base in ROLE_LOCAL_PARTS and not _is_consumer_domain(domain):
        return "role"

    return "automated"


def _decode_mime_header(value: str) -> str:
    if not value or "=?" not in value:
        return value
    try:
        parts: List[str] = []
        for fragment, charset in decode_header(value):
            if isinstance(fragment, bytes):
                parts.append(fragment.decode(charset or "utf-8", errors="replace"))
            else:
                parts.append(fragment)
        return "".join(parts).strip()
    except Exception:
        return value


def _title_case_word(word: str) -> str:
    if not word:
        return ""
    lower = word.lower()
    if lower in NAME_PARTICLES:
        return lower.capitalize()
    if word.isupper() and len(word) > 1:
        return word.capitalize()
    return word[:1].upper() + word[1:].lower()


def _title_case_name(text: str) -> str:
    return " ".join(_title_case_word(word) for word in text.split())


def _names_from_email_local_part(email: str) -> Optional[Tuple[str, str]]:
    local_part, _ = _email_parts(email)
    if not local_part:
        return None
    base = local_part.split("+", 1)[0]
    base = re.sub(r"\d+$", "", base)
    tokens = [token for token in re.split(r"[._-]+", base) if token and not token.isdigit()]
    tokens = [token for token in tokens if len(re.sub(r"[^A-Za-zÀ-ÿ]", "", token)) >= 2]
    if len(tokens) < 2:
        return None

    for index, token in enumerate(tokens):
        if index > 0 and token.lower() in NAME_PARTICLES:
            nome = _title_case_name(" ".join(tokens[:index]))
            cognome = _title_case_name(" ".join(tokens[index:]))
            return nome, cognome

    return _title_case_name(tokens[0]), _title_case_name(" ".join(tokens[1:]))


def _display_name_tokens(nome: str, cognome: str) -> Set[str]:
    tokens: Set[str] = set()
    for part in (nome, cognome):
        for word in part.split():
            token = _normalize_text_key(word)
            if token:
                tokens.add(token)
    return tokens


def _parse_person_name(display_name: str, email: str = "") -> Optional[Tuple[str, str]]:
    name = _clean_outlook_name(display_name, email)
    if not _is_plausible_person_name(name):
        return None

    words = name.split()
    hint = _names_from_email_local_part(email)
    display_tokens = {_normalize_text_key(word) for word in words}

    if len(words) == 1:
        if hint and _normalize_text_key(words[0]) in {
            _normalize_text_key(hint[0]),
            _normalize_text_key(hint[1]),
        }:
            return hint
        return _title_case_name(words[0]), ""

    if words[0].lower() in NAME_PARTICLES:
        return _title_case_name(words[-1]), _title_case_name(" ".join(words[:-1]))

    if len(words) >= 2 and words[1].lower() in NAME_PARTICLES:
        return _title_case_name(words[0]), _title_case_name(" ".join(words[1:]))

    if hint:
        hint_tokens = _display_name_tokens(hint[0], hint[1])
        if hint_tokens & display_tokens:
            return hint

    if len(words) == 2:
        return _title_case_name(words[1]), _title_case_name(words[0])

    return _title_case_name(words[0]), _title_case_name(" ".join(words[1:]))


def _normalize_name_key(nome: str, cognome: str) -> Tuple[str, str]:
    return _normalize_text_key(nome), _normalize_text_key(cognome)


def _canonical_person_key(record: ContactRecord) -> Tuple[str, ...]:
    nome = record.get("nome", "")
    cognome = record.get("cognome", "")
    tokens = _display_name_tokens(nome, cognome)
    hint = _names_from_email_local_part(record.get("email", ""))
    if hint:
        tokens |= _display_name_tokens(hint[0], hint[1])
    if len(tokens) >= 2:
        return ("tokens",) + tuple(sorted(tokens))
    return ("name",) + _normalize_name_key(nome, cognome)


def _split_name(display_name: str, email: str = "") -> Tuple[str, str]:
    name = _clean_outlook_name(display_name, email)
    if not name:
        return "", ""
    if " " not in name:
        return name, ""
    parts = name.split(None, 1)
    return parts[0], parts[1] if len(parts) > 1 else ""


def _clean_outlook_name(name: str, email: str = "") -> str:
    cleaned = _decode_mime_header(_safe_str(name))
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


def _name_quality_score(
    nome: str,
    cognome: str,
    email: str = "",
) -> int:
    score = 0
    if nome and cognome:
        score += 10
    hint = _names_from_email_local_part(email)
    if hint and (nome, cognome) == hint:
        score += 20
    score += len(nome) + len(cognome)
    return score


def _is_better_name(
    current_nome: str,
    current_cognome: str,
    new_nome: str,
    new_cognome: str,
    email: str = "",
) -> bool:
    if not new_nome and not new_cognome:
        return False
    if _is_generic_name(new_nome, new_cognome):
        return False
    if not current_nome and not current_cognome:
        return True
    if _is_generic_name(current_nome, current_cognome):
        return True
    current_score = _name_quality_score(current_nome, current_cognome, email)
    new_score = _name_quality_score(new_nome, new_cognome, email)
    if new_score != current_score:
        return new_score > current_score
    current_has_both = bool(current_nome and current_cognome)
    new_has_both = bool(new_nome and new_cognome)
    if new_has_both and not current_has_both:
        return True
    if current_has_both and not new_has_both:
        return False
    return (len(new_nome) + len(new_cognome)) > (len(current_nome) + len(current_cognome))


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
    local_part, domain = _email_parts(email_norm)
    return (
        _is_automated_local_part(local_part)
        or _is_automated_domain(domain)
    )


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


def _html_to_plain(html_text: str) -> str:
    if not html_text:
        return ""
    text = HTML_TAG_REGEX.sub(" ", html_text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def _decode_email_payload(part: Any) -> str:
    payload = part.get_payload(decode=True)
    if not payload:
        return ""
    charset = part.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def _get_mbox_header(msg: Any, name: str) -> str:
    value = msg.get(name, "")
    if value is None:
        return ""
    return _safe_str(value)


def _parse_mbox_date(msg: Any) -> Optional[datetime]:
    date_header = _get_mbox_header(msg, "Date")
    if not date_header:
        return None
    try:
        parsed = parsedate_to_datetime(date_header)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed
    except (TypeError, ValueError, OverflowError):
        return None


def _get_mbox_message_body(msg: Any) -> str:
    if msg.is_multipart():
        plain_parts: List[str] = []
        html_parts: List[str] = []
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            content_type = part.get_content_type()
            if content_type == "text/plain":
                plain_parts.append(_decode_email_payload(part))
            elif content_type == "text/html":
                html_parts.append(_decode_email_payload(part))
        if plain_parts:
            return "\n\n".join(part for part in plain_parts if part.strip())
        if html_parts:
            return _html_to_plain("\n".join(html_parts))
        return ""
    content_type = msg.get_content_type()
    if content_type == "text/plain":
        return _decode_email_payload(msg)
    if content_type == "text/html":
        return _html_to_plain(_decode_email_payload(msg))
    return ""


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
    return _html_to_plain(_safe_str(html_body))


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


def _line_has_phone_keyword(line: str) -> bool:
    lower_line = line.lower()
    return any(keyword in lower_line for keyword in PHONE_KEYWORDS)


def _normalize_phone_match(match_text: str) -> str:
    return re.sub(r"\s+", " ", match_text.strip())


def _is_plausible_phone(match_text: str, line: str) -> bool:
    if not _line_has_phone_keyword(line):
        return False
    compact = re.sub(r"[\s.-]", "", match_text)
    digits = re.sub(r"\D", "", compact)
    if not digits:
        return False
    if re.match(r"^\d{3}-\d{7}$", match_text.replace(" ", "")):
        return False
    if len(digits) >= 9 and not compact.startswith(("+39", "0039", "0", "+")):
        return False
    if len(digits) > 13:
        return False
    return True


def _extract_phones_from_lines(lines: List[str]) -> Set[str]:
    phones: Set[str] = set()
    for line in lines:
        for match in PHONE_REGEX.finditer(line):
            candidate = _normalize_phone_match(match.group(0))
            if _is_plausible_phone(candidate, line):
                phones.add(candidate)
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


def _message_subject(msg: Any) -> str:
    if msg is None:
        return ""
    if hasattr(msg, "get"):
        return _safe_str(msg.get("Subject", ""))
    return _safe_str(getattr(msg, "subject", ""))


def _log_message_error(
    exc: Exception,
    folder_name: str = "",
    index: int = -1,
    msg: Any = None,
) -> None:
    subject = _message_subject(msg)
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

    def _tick_progress(self, status_text: str) -> None:
        self._processed_messages += 1
        if self._processed_messages % 25 == 0:
            pseudo_progress = min(
                0.95,
                self._processed_messages / max(self._processed_messages + 500, 1),
            )
            self._emit_progress(pseudo_progress)
        self._emit_status(status_text)

    def _upsert_contact(
        self,
        email: str,
        nome: str = "",
        cognome: str = "",
        telefoni: Optional[Set[str]] = None,
        societa: str = "",
        delivery_time: Optional[datetime] = None,
        from_metadata: bool = True,
        contact_type: str = "person",
    ) -> None:
        email_norm = _normalize_email(email)
        if not email_norm:
            return
        if not from_metadata:
            if email_norm not in self.contatti:
                return
            if self.contatti[email_norm].get("contact_type") != "person":
                return

        fallback_societa = ""
        if contact_type == "role":
            company = societa or _domain_to_company(email_norm)
            nome = ""
            cognome = ""
            fallback_societa = company
        elif contact_type == "person":
            if from_metadata and not nome and not cognome:
                return
        else:
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
                "contact_type": contact_type,
                "domain_key": _business_domain_key(email_norm) if contact_type == "role" else "",
            }

        record = self.contatti[email_norm]
        if from_metadata:
            record["contact_type"] = contact_type
            if contact_type == "role":
                record["domain_key"] = _business_domain_key(email_norm)

        if contact_type == "person" and _is_better_name(
            record["nome"], record["cognome"], nome, cognome, email_norm
        ):
            record["nome"] = nome
            record["cognome"] = cognome
        elif contact_type == "role":
            record["nome"] = ""
            record["cognome"] = ""

        if telefoni:
            record["telefoni"].update(telefoni)

        societa_candidate = societa or fallback_societa
        if societa_candidate:
            current_ts = record.get("societa_updated_at")
            if delivery_time is None:
                if not record.get("societa"):
                    record["societa"] = societa_candidate
            elif current_ts is None or delivery_time >= current_ts:
                record["societa"] = societa_candidate
                record["societa_updated_at"] = delivery_time

    def _upsert_from_header(
        self,
        display_name: str,
        email: str,
        delivery_time: Optional[datetime] = None,
    ) -> None:
        contact_type = _classify_contact(email, display_name)
        if contact_type in ("automated", "spam"):
            return

        email_norm = _normalize_email(email)
        if contact_type == "role":
            company = _domain_to_company(email_norm)
            self._upsert_contact(
                email=email_norm,
                societa=company,
                delivery_time=delivery_time,
                from_metadata=True,
                contact_type="role",
            )
            return

        parsed = _parse_person_name(display_name, email_norm)
        if not parsed:
            return
        nome, cognome = parsed
        self._upsert_contact(
            email=email_norm,
            nome=nome,
            cognome=cognome,
            delivery_time=delivery_time,
            from_metadata=True,
            contact_type="person",
        )

    def _parse_signature_for_enrichment(
        self,
        body: str,
        sender_email: str,
        delivery_time: Optional[datetime],
    ) -> None:
        sender_email = _normalize_email(sender_email)
        if not sender_email or sender_email not in self.contatti:
            return
        if self.contatti[sender_email].get("contact_type") != "person":
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

    def _process_contacts_from_headers(
        self,
        sender_name: str,
        sender_email: str,
        recipients: List[Tuple[str, str]],
        body: str,
        delivery_time: Optional[datetime],
    ) -> None:
        if sender_email:
            self._upsert_from_header(sender_name, sender_email, delivery_time)

        for display_name, email in recipients:
            self._upsert_from_header(display_name, email, delivery_time)

        if body and sender_email:
            self._parse_signature_for_enrichment(body, sender_email, delivery_time)

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

    def _process_single_message(
        self,
        msg: Any,
        folder_name: str,
        message_index: int,
    ) -> None:
        self._tick_progress(f"Stato: Scansione {folder_name} (Mail {message_index})...")

        headers_text = _safe_str(getattr(msg, "transport_headers", ""))
        headers = _parse_transport_headers(headers_text)
        delivery_time = _parse_delivery_time(msg)

        sender_name, sender_email = self._extract_sender(msg, headers)
        recipients = self._extract_recipients(msg, headers)
        body = _get_message_body(msg)

        self._process_contacts_from_headers(
            sender_name, sender_email, recipients, body, delivery_time
        )

    def _process_mbox_message(self, msg: Any, message_index: int) -> None:
        self._tick_progress(f"Stato: Scansione MBOX (Mail {message_index})...")

        delivery_time = _parse_mbox_date(msg)

        from_header = _get_mbox_header(msg, "From")
        sender_entries = _extract_emails_from_header(from_header)
        sender_name = ""
        sender_email = ""
        if sender_entries:
            sender_name, sender_email = sender_entries[0]

        recipients: List[Tuple[str, str]] = []
        seen: Set[str] = set()
        for header_name in ("To", "Cc"):
            for display_name, email in _extract_emails_from_header(_get_mbox_header(msg, header_name)):
                if email not in seen:
                    seen.add(email)
                    recipients.append((display_name, email))

        body = _get_mbox_message_body(msg)
        self._process_contacts_from_headers(
            sender_name, sender_email, recipients, body, delivery_time
        )

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
        pypff = _import_pypff()
        pst_file = pypff.file()
        pst_file.open(pst_path)
        try:
            root = pst_file.get_root_folder()
            self._emit_status("Stato: Analisi archivio in corso...")
            self._traverse_folder(root)
        finally:
            pst_file.close()

    def extract_from_mbox(self, mbox_path: str) -> None:
        mbox_label = os.path.basename(mbox_path) or "MBOX"
        mbox = mailbox.mbox(mbox_path, create=False)
        try:
            self._emit_status("Stato: Analisi archivio MBOX in corso...")
            for index, message in enumerate(mbox, start=1):
                try:
                    self._process_mbox_message(message, index)
                except Exception as exc:
                    _log_message_error(exc, folder_name=mbox_label, index=index, msg=message)
                finally:
                    del message
        finally:
            mbox.close()


def _merge_contact_records(
    records: List[ContactRecord],
    contact_type: str = "person",
) -> Dict[str, str]:
    emails: Set[str] = set()
    phones: Set[str] = set()
    companies: Set[str] = set()
    nome = ""
    cognome = ""

    for record in records:
        emails.add(record.get("email", ""))
        phones.update(record.get("telefoni", set()))
        if record.get("societa"):
            companies.add(record["societa"])
        if contact_type == "person":
            candidate_nome = record.get("nome", "")
            candidate_cognome = record.get("cognome", "")
            record_email = record.get("email", "")
            if _is_better_name(
                nome, cognome, candidate_nome, candidate_cognome, record_email
            ):
                nome = candidate_nome
                cognome = candidate_cognome

    if contact_type == "role":
        return {
            "Nome": "",
            "Cognome": "",
            "Email": "\n".join(sorted(email for email in emails if email)),
            "Telefono": "\n".join(sorted(phones)),
            "Società": "\n".join(sorted(companies)),
        }

    return {
        "Nome": nome,
        "Cognome": cognome,
        "Email": "\n".join(sorted(email for email in emails if email)),
        "Telefono": "\n".join(sorted(phones)),
        "Società": "\n".join(sorted(companies)),
    }


def _build_export_dataframe(contatti: ContactsDict) -> pd.DataFrame:
    columns = ["Nome", "Cognome", "Email", "Telefono", "Società"]
    if not contatti:
        return pd.DataFrame(columns=columns)

    person_groups: Dict[Tuple[str, ...], List[ContactRecord]] = {}
    role_groups: Dict[str, List[ContactRecord]] = {}

    for record in contatti.values():
        contact_type = record.get("contact_type", "person")
        if contact_type == "role":
            domain_key = record.get("domain_key") or _business_domain_key(record.get("email", ""))
            role_groups.setdefault(domain_key, []).append(record)
        elif contact_type == "person":
            if not record.get("nome") and not record.get("cognome"):
                continue
            person_groups.setdefault(_canonical_person_key(record), []).append(record)

    rows: List[Dict[str, str]] = []
    for group_records in person_groups.values():
        rows.append(_merge_contact_records(group_records, "person"))
    for group_records in role_groups.values():
        rows.append(_merge_contact_records(group_records, "role"))

    if not rows:
        return pd.DataFrame(columns=columns)

    result = pd.DataFrame(rows, columns=columns)
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


def run_extraction(source_path: str, output_dir: str, event_queue: queue.Queue) -> None:
    def push(event_type: str, **payload: Any) -> None:
        event_queue.put({"type": event_type, **payload})

    def status_callback(text: str) -> None:
        push("status", text=text)

    def progress_callback(value: float) -> None:
        push("progress", value=value)

    try:
        if not os.path.isfile(source_path):
            raise FileNotFoundError("Il file sorgente selezionato non esiste.")
        if not os.path.isdir(output_dir):
            raise FileNotFoundError("La cartella di destinazione non esiste.")

        ext = os.path.splitext(source_path)[1].lower()
        if ext not in (".pst", ".ost", ".mbox"):
            raise ValueError("Formato non supportato. Usa .pst, .ost o .mbox")

        extractor = ContactExtractor()
        extractor.set_callbacks(status_callback, progress_callback)
        push("progress", value=0.0)

        if ext in (".pst", ".ost"):
            push("status", text="Stato: Apertura archivio Outlook...")
            extractor.extract_from_pst(source_path)
        else:
            push("status", text="Stato: Apertura archivio MBOX...")
            extractor.extract_from_mbox(source_path)

        push("status", text="Stato: Generazione Report Excel...")
        push("progress", value=0.97)
        output_path = os.path.join(output_dir, "rubrica_estratta.xlsx")
        export_to_excel(extractor.contatti, output_path)

        push("progress", value=1.0)
        push("status", text="Stato: Processo Completato con successo!")
        push("done")
    except ModuleNotFoundError as exc:
        _log_message_error(exc)
        push("error", text=f"Errore: {exc}")
        push("progress", value=0.0)
    except Exception as exc:
        _log_message_error(exc)
        push("error", text=f"Errore: {exc}")
        push("progress", value=0.0)
