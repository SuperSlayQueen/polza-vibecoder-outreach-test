#!/usr/bin/env python3
"""
Polza Agency — скрипт персонализации для холодного аутрича.

На входе: CSV/XLSX со столбцами company / website (и опционально name, email).
На выходе: тот же файл + столбец «Персонализация» (1–2 предложения по сайту)
и служебные столбцы для QA (доступность сайта, домен email, флаги ловушек).

Запуск:
  python personalize.py --input data/companies_seed.csv --output data/polza_outreach_base.xlsx
  python personalize.py --input data/task4_input.csv --output data/task4_personalized.xlsx --strict
"""

from __future__ import annotations

import argparse
import csv
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import tldextract
from bs4 import BeautifulSoup

USER_AGENT = (
    "Mozilla/5.0 (compatible; PolzaOutreachResearch/1.0; +https://polzaagency.ru)"
)
TIMEOUT = 12
MAX_TEXT_CHARS = 6000

# Страницы, с которых чаще всего берём смысл компании
PATH_CANDIDATES = [
    "",
    "/about",
    "/about/",
    "/o-nas",
    "/o-kompanii",
    "/company",
    "/products",
    "/solutions",
    "/uslugi",
    "/services",
]


@dataclass
class SiteSnapshot:
    url: str
    reachable: bool
    title: str = ""
    meta_description: str = ""
    h1: str = ""
    body_excerpt: str = ""
    emails_found: list[str] = field(default_factory=list)
    error: str = ""


@dataclass
class QaFlags:
    site_unreachable: bool = False
    email_domain_mismatch: bool = False
    name_site_mismatch: bool = False
    likely_b2c: bool = False
    empty_or_placeholder: bool = False
    notes: list[str] = field(default_factory=list)

    def as_text(self) -> str:
        flags = []
        if self.site_unreachable:
            flags.append("сайт недоступен")
        if self.email_domain_mismatch:
            flags.append("домен email ≠ сайт")
        if self.name_site_mismatch:
            flags.append("название ≠ сайт")
        if self.likely_b2c:
            flags.append("похоже на B2C")
        if self.empty_or_placeholder:
            flags.append("пустые/плейсхолдер данные")
        if self.notes:
            flags.extend(self.notes)
        return "; ".join(flags)


def normalize_url(raw: str) -> str:
    raw = (raw or "").strip()
    if not raw:
        return ""
    if not raw.startswith(("http://", "https://")):
        raw = "https://" + raw
    return raw.rstrip("/")


def registrable_domain(host_or_url: str) -> str:
    if not host_or_url:
        return ""
    if "://" not in host_or_url and "@" in host_or_url:
        host_or_url = host_or_url.split("@", 1)[1]
    if "://" not in host_or_url:
        host_or_url = "https://" + host_or_url
    host = urlparse(host_or_url).hostname or ""
    ext = tldextract.extract(host)
    if not ext.domain:
        return host.lower()
    return f"{ext.domain}.{ext.suffix}".lower()


def clean_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def fetch_html(url: str, session: requests.Session) -> tuple[Optional[str], str]:
    try:
        resp = session.get(url, timeout=TIMEOUT, allow_redirects=True)
        if resp.status_code >= 400:
            return None, f"HTTP {resp.status_code}"
        ctype = resp.headers.get("Content-Type", "")
        if "text/html" not in ctype and "application/xhtml" not in ctype and ctype:
            # некоторые сайты не отдают content-type — всё равно пробуем
            if "text" not in ctype and "html" not in ctype:
                return None, f"не HTML ({ctype})"
        resp.encoding = resp.apparent_encoding or resp.encoding or "utf-8"
        return resp.text, ""
    except requests.RequestException as exc:
        return None, str(exc.__class__.__name__)


def extract_emails(text: str) -> list[str]:
    found = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text or "")
    # отсекаем картинки и мусор
    skip = (
        "example.com",
        "sentry.io",
        "wixpress",
        "schema.org",
        ".png",
        ".jpg",
        ".jpeg",
        ".webp",
        ".gif",
        ".svg",
        "2x.webp",
        "sentry",
    )
    out = []
    for e in found:
        low = e.lower()
        if any(s in low for s in skip):
            continue
        if low not in out:
            out.append(low)
    return out[:8]


def parse_page(html: str, base_url: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()

    title = clean_text(soup.title.get_text()) if soup.title else ""
    meta = ""
    md = soup.find("meta", attrs={"name": re.compile("^description$", re.I)})
    if md and md.get("content"):
        meta = clean_text(md["content"])
    if not meta:
        og = soup.find("meta", attrs={"property": "og:description"})
        if og and og.get("content"):
            meta = clean_text(og["content"])

    h1 = ""
    h1_tag = soup.find("h1")
    if h1_tag:
        h1 = clean_text(h1_tag.get_text())

    body = clean_text(soup.get_text(" ", strip=True))[:MAX_TEXT_CHARS]
    emails = extract_emails(html)
    return {
        "title": title,
        "meta": meta,
        "h1": h1,
        "body": body,
        "emails": emails,
        "base_url": base_url,
    }


def snapshot_site(website: str, session: requests.Session) -> SiteSnapshot:
    url = normalize_url(website)
    if not url:
        return SiteSnapshot(url="", reachable=False, error="нет сайта")

    # главная
    html, err = fetch_html(url, session)
    if not html:
        # иногда http редиректит иначе — пробуем www
        alt = url.replace("://", "://www.") if "://www." not in url else url
        html, err2 = fetch_html(alt, session)
        if html:
            url = alt
            err = ""
        else:
            return SiteSnapshot(url=url, reachable=False, error=err or err2)

    main = parse_page(html, url)
    title, meta, h1, body = main["title"], main["meta"], main["h1"], main["body"]
    emails = list(main["emails"])

    # если meta пустая — добираем /about
    if len(meta) < 40 or len(body) < 200:
        for path in PATH_CANDIDATES[1:4]:
            sub = urljoin(url + "/", path.lstrip("/"))
            sub_html, _ = fetch_html(sub, session)
            if not sub_html:
                continue
            page = parse_page(sub_html, sub)
            if len(page["meta"]) > len(meta):
                meta = page["meta"]
            if len(page["body"]) > len(body):
                body = page["body"]
            if not h1 and page["h1"]:
                h1 = page["h1"]
            for e in page["emails"]:
                if e not in emails:
                    emails.append(e)
            time.sleep(0.3)
            break

    return SiteSnapshot(
        url=url,
        reachable=True,
        title=title,
        meta_description=meta,
        h1=h1,
        body_excerpt=body,
        emails_found=emails,
    )


def pick_signal_sentence(snapshot: SiteSnapshot, company: str) -> str:
    """Достаём факт с сайта без LLM — только то, что реально есть в тексте."""
    chunks: list[str] = []
    for raw in (snapshot.meta_description, snapshot.h1, snapshot.title):
        raw = clean_text(raw)
        if raw and raw.lower() not in {c.lower() for c in chunks}:
            chunks.append(raw)

    # из body — предложения с продуктовыми маркерами
    markers = (
        "помог",
        "разраб",
        "платформ",
        "сервис",
        "решен",
        "клиент",
        "автоматиз",
        "B2B",
        "бизнес",
        "продаж",
        "интеграц",
        "облач",
        "CRM",
        "логист",
        "производ",
        "маркетинг",
        "рекрут",
        "подбор",
        "доставк",
        "склад",
        "аналитик",
        "управлен",
    )
    sentences = re.split(r"(?<=[.!?…])\s+", snapshot.body_excerpt)
    for s in sentences:
        s = clean_text(s)
        if 40 <= len(s) <= 220 and any(m.lower() in s.lower() for m in markers):
            if s not in chunks:
                chunks.append(s)
            if len(chunks) >= 4:
                break

    if not chunks:
        return ""

    # собираем 1–2 коротких предложения
    primary = chunks[0]
    secondary = ""
    for c in chunks[1:]:
        if c.lower() not in primary.lower() and len(c) > 30:
            secondary = c
            break

    name = clean_text(company) or "Компания"
    # убираем дублирование названия в начале
    fact = primary
    if secondary and len(primary) < 100:
        fact = f"{primary} {secondary}"
    fact = re.sub(r"\s+", " ", fact).strip()
    if len(fact) > 280:
        fact = fact[:277].rsplit(" ", 1)[0] + "…"

    # человеческий формат под вставку в письмо
    if name.lower() not in fact.lower()[:50]:
        # короче и менее шаблонно, чем «акцент на следующее»
        return f"{name}: {fact}"
    return fact


FREEMAIL = {
    "gmail.com",
    "yandex.ru",
    "ya.ru",
    "mail.ru",
    "bk.ru",
    "inbox.ru",
    "list.ru",
    "outlook.com",
    "hotmail.com",
    "icloud.com",
    "126.com",
    "163.com",
    "qq.com",
    "yeah.net",
    "sina.com",
    "sohu.com",
}

# частые латинские написания для сверки названия с title
NAME_ALIASES = {
    "кайтен": ["kaiten"],
    "хантфлоу": ["huntflow"],
    "поток": ["potok"],
    "майндбокс": ["mindbox"],
    "битрикс": ["bitrix"],
    "клеверенс": ["cleverence"],
    "контур": ["kontur"],
    "тензор": ["tensor", "sbis"],
    "селектел": ["selectel"],
    "мойсклад": ["moysklad"],
    "эльба": ["elba", "e-kontur"],
    "простые": ["prostiezvonki"],
    "хабр": ["habr"],
    "хедхантер": ["headhunter", "hh.ru"],
    "авито": ["avito"],
    "яндекс": ["yandex"],
    "диасофт": ["diasoft"],
    "наумен": ["naumen"],
}


def detect_b2c(text: str, company: str = "") -> bool:
    t = f"{text} {company}".lower()
    # явные B2C-бренды / ниши вне ICP Polza
    hard = (
        "яндекс лавка",
        "lavka",
        "доставка еды",
        "продукт за час",
        "супермаркет",
        "заказ продуктов",
        "для всей семьи",
    )
    if any(h in t for h in hard):
        return True
    b2c_hits = sum(
        1
        for w in (
            "интернет-магазин",
            "купить сейчас",
            "корзина",
            "доставка до двери",
            "скидка дня",
            "личный кабинет покупателя",
        )
        if w in t
    )
    b2b_hits = sum(
        1
        for w in ("b2b", "для бизнеса", "корпоратив", "оптовым", "интеграц", "api", "лпр")
        if w in t
    )
    return b2c_hits >= 2 and b2b_hits == 0


def name_matches_site(company: str, snapshot: SiteSnapshot) -> bool:
    blob = f"{snapshot.title} {snapshot.h1} {snapshot.url} {snapshot.meta_description}".lower()
    tokens = [
        t
        for t in re.split(r"[^a-zA-Zа-яА-ЯёЁ0-9]+", company.lower())
        if len(t) >= 3
        and t
        not in {
            "ооо",
            "ао",
            "пао",
            "llc",
            "ltd",
            "group",
            "для",
            "бизнеса",
            "агентство",
            "digital",
            "работа",
            "pro",
            "center",
            "technology",
            "speech",
        }
    ]
    if not tokens:
        return True
    if any(t in blob for t in tokens):
        return True
    # домен часто содержит бренд латиницей
    dom = registrable_domain(snapshot.url)
    if any(t in dom for t in tokens if t.isascii()):
        return True
    for t in tokens:
        for alias in NAME_ALIASES.get(t, []):
            if alias in blob or alias in dom:
                return True
    return False


def qa_check(
    company: str,
    website: str,
    email: str,
    snapshot: SiteSnapshot,
) -> QaFlags:
    flags = QaFlags()
    company = clean_text(company)
    website = clean_text(website)
    email = clean_text(email).lower()

    if not company or company.lower() in {"test", "тест", "n/a", "-", "demo company"}:
        flags.empty_or_placeholder = True
        flags.notes.append("подозрительное название")
    if (
        not website
        or website.lower() in {"http://", "https://", "n/a", "-"}
        or "example.com" in website.lower()
    ):
        flags.empty_or_placeholder = True
        flags.notes.append("нет сайта / example")
    if not email or "@" not in email or email.endswith("@example.com"):
        flags.empty_or_placeholder = True
        flags.notes.append("email пустой/плейсхолдер")

    # email vs заявленный сайт — проверяем даже если сайт сейчас лежит
    site_dom = registrable_domain(snapshot.url or website)
    mail_dom = registrable_domain(email) if email and "@" in email else ""
    if site_dom and mail_dom and site_dom != mail_dom:
        if mail_dom in FREEMAIL:
            flags.email_domain_mismatch = True
            flags.notes.append("freemail вместо корпоративного домена")
        else:
            flags.email_domain_mismatch = True

    if not snapshot.reachable:
        flags.site_unreachable = True
        return flags

    if company and (snapshot.title or snapshot.h1 or snapshot.url):
        if not name_matches_site(company, snapshot):
            flags.name_site_mismatch = True

    page_text = f"{snapshot.meta_description} {snapshot.body_excerpt}"
    if detect_b2c(page_text, company):
        flags.likely_b2c = True

    return flags


def personalize_row(
    company: str,
    website: str,
    email: str,
    session: requests.Session,
) -> dict:
    snap = snapshot_site(website, session)
    flags = qa_check(company, website, email, snap)

    if not snap.reachable:
        personalization = (
            "Сайт компании на момент сбора недоступен — персонализацию "
            "нужно добрать вручную (LinkedIn / новости / 2ГИС)."
        )
        source = "unavailable"
    else:
        personalization = pick_signal_sentence(snap, company)
        source = "website"
        if not personalization:
            personalization = (
                f"По сайту {snap.url} не удалось вытащить устойчивый продуктовый "
                f"сигнал (мало текста / лендинг на JS) — нужен ручной ресёрч."
            )
            source = "weak_signal"
            flags.notes.append("слабый сигнал с сайта")

    return {
        "Персонализация": personalization,
        "Источник_персонализации": source,
        "Сайт_финальный": snap.url,
        "Title": snap.title,
        "Meta": snap.meta_description,
        "Emails_на_сайте": ", ".join(snap.emails_found),
        "QA_флаги": flags.as_text(),
        "Ошибка_сбора": snap.error,
    }


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in {"company", "компания", "название", "name_company", "org"}:
            mapping[col] = "company"
        elif key in {"website", "site", "сайт", "url", "domain"}:
            mapping[col] = "website"
        elif key in {"name", "имя", "contact", "лпр", "contact_name", "фио"}:
            mapping[col] = "name"
        elif key in {"email", "e-mail", "почта", "mail"}:
            mapping[col] = "email"
        elif key in {"title", "должность", "position", "role"}:
            mapping[col] = "title"
        elif key in {"niche", "ниша", "segment", "сегмент"}:
            mapping[col] = "niche"
    out = df.rename(columns=mapping)
    for need in ("company", "website", "name", "email"):
        if need not in out.columns:
            out[need] = ""
    return out


def run(input_path: Path, output_path: Path, delay: float, limit: Optional[int]) -> None:
    df = normalize_columns(read_table(input_path))
    if limit:
        df = df.head(limit).copy()

    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru-RU,ru;q=0.9"})

    rows = []
    total = len(df)
    print(f"Personalizing {total} rows from {input_path.name}...")

    for i, row in df.iterrows():
        company = str(row.get("company") or "").strip()
        website = str(row.get("website") or "").strip()
        email = str(row.get("email") or "").strip()
        print(f"[{len(rows)+1}/{total}] {company or website or '-'}")
        enriched = personalize_row(company, website, email, session)
        merged = {
            "Компания": company,
            "Сайт": website,
            "Имя": str(row.get("name") or "").strip(),
            "Email": email,
            "Должность": str(row.get("title") or "").strip(),
            "Ниша": str(row.get("niche") or "").strip(),
            **enriched,
        }
        # сохраняем исходные доп. столбцы
        for col in df.columns:
            if col not in {"company", "website", "name", "email", "title", "niche"}:
                merged[f"src_{col}"] = row.get(col)
        rows.append(merged)
        time.sleep(delay)

    out = pd.DataFrame(rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.suffix.lower() in {".xlsx", ".xls"}:
        out.to_excel(output_path, index=False)
    else:
        out.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"Done: {output_path}")
    flagged = out[out["QA_флаги"].astype(str).str.len() > 0]
    if len(flagged):
        print(f"QA flagged rows: {len(flagged)}")
        for _, r in flagged.iterrows():
            try:
                print(f"  - {r['Компания']}: {r['QA_флаги']}")
            except UnicodeEncodeError:
                print("  - (flagged row)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Персонализация базы для Polza outreach")
    parser.add_argument("--input", "-i", required=True, help="CSV/XLSX со списком компаний")
    parser.add_argument("--output", "-o", required=True, help="Куда сохранить результат")
    parser.add_argument("--delay", type=float, default=0.6, help="Пауза между сайтами, сек")
    parser.add_argument("--limit", type=int, default=None, help="Ограничить число строк")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Только подсказка: в Task 4 смотрите столбец QA_флаги на ловушки",
    )
    args = parser.parse_args()
    if args.strict:
        print(
            "Режим --strict: скрипт сам не удаляет строки, но пишет QA_флаги "
            "(несовпадение домена, недоступный сайт, B2C, плейсхолдеры)."
        )
    run(Path(args.input), Path(args.output), delay=args.delay, limit=args.limit)


if __name__ == "__main__":
    main()
