#!/usr/bin/env python3
"""Проверка домена email (MX) и лёгкий поиск mailto на /contacts."""

from __future__ import annotations

import csv
import re
import smtplib
import socket
import time
from pathlib import Path
from urllib.parse import urljoin

import dns.resolver
import pandas as pd
import requests
from bs4 import BeautifulSoup

UA = "Mozilla/5.0 (compatible; PolzaLeadQA/1.0)"


def has_mx(domain: str) -> tuple[bool, str]:
    try:
        answers = dns.resolver.resolve(domain, "MX")
        hosts = sorted(str(r.exchange).rstrip(".") for r in answers)
        return True, ", ".join(hosts[:3])
    except Exception as exc:  # noqa: BLE001
        return False, exc.__class__.__name__


def scrape_emails(website: str) -> list[str]:
    session = requests.Session()
    session.headers["User-Agent"] = UA
    urls = [website]
    for p in ("/contacts", "/contact", "/kontakty", "/about", "/company"):
        urls.append(urljoin(website.rstrip("/") + "/", p.lstrip("/")))
    found: list[str] = []
    for url in urls[:5]:
        try:
            r = session.get(url, timeout=10, allow_redirects=True)
            if r.status_code >= 400:
                continue
            for m in re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", r.text):
                low = m.lower()
                if any(x in low for x in ("example.", "sentry.", ".png", "wixpress")):
                    continue
                if low not in found:
                    found.append(low)
            soup = BeautifulSoup(r.text, "lxml")
            for a in soup.select('a[href^="mailto:"]'):
                href = a.get("href", "")
                mail = href.replace("mailto:", "").split("?")[0].strip().lower()
                if mail and mail not in found:
                    found.append(mail)
        except requests.RequestException:
            continue
        time.sleep(0.25)
    return found[:10]


def main() -> None:
    src = Path("data/companies_seed.csv")
    df = pd.read_csv(src)
    df = df[df["email"].astype(str).str.contains("@")]
    df = df[~df["email"].astype(str).str.endswith("@skip.local")]
    df = df.drop_duplicates(subset=["website"], keep="first")

    rows = []
    for _, row in df.iterrows():
        email = str(row["email"]).strip().lower()
        domain = email.split("@", 1)[1]
        ok, mx = has_mx(domain)
        site_emails = []
        try:
            site_emails = scrape_emails(str(row["website"]))
        except Exception as exc:  # noqa: BLE001
            site_emails = [f"ERR:{exc.__class__.__name__}"]
        preferred = email
        # если на сайте нашли sales/info — и исходный не найден, подсветим
        note = ""
        if site_emails and email not in site_emails:
            role = [e for e in site_emails if e.startswith(("sales@", "info@", "hello@", "office@", "commerce@"))]
            if role:
                note = f"на сайте также: {', '.join(role[:3])}"
        rows.append(
            {
                **row.to_dict(),
                "email": preferred,
                "mx_ok": ok,
                "mx": mx,
                "emails_on_site": ", ".join(site_emails),
                "qa_note": note,
            }
        )
        print(f"{row['company']}: MX={'OK' if ok else 'FAIL'} | site={site_emails[:3]}")

    out = pd.DataFrame(rows)
    out_path = Path("data/companies_verified.csv")
    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Saved {len(out)} → {out_path}")
    print(f"MX fail: {(~out['mx_ok']).sum()}")


if __name__ == "__main__":
    main()
