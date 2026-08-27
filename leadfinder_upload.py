"""
leadfinder_upload.py
--------------------
Bron-neutrale upload- en verrijkingsmodule voor VoiceStamp leadfinder.

Flow:
  1. Upload een CSV of Excel met leads (naam + optioneel website/plaats/etc.).
  2. Kies welke kolom de bedrijfsnaam en welke de website/domein is.
  3. De tool crawlt elke website (contact/over-ons/colofon/privacy + footer),
     haalt e-mailadressen eruit, en valt terug op info@domein als het niks vindt.
  4. Elk adres wordt gratis geverifieerd via een MX-check (+ optionele SMTP-probe).
  5. Je standaardmail (per segment) wordt per lead ingevuld via {placeholders}.
  6. Export naar Excel en naar een Smartlead/Instantly-klare CSV.

Draai de crawl/verificatie LOKAAL. Streamlit Community Cloud blokkeert vrijwel
zeker uitgaande SMTP (poort 25) en kan uitgaande HTTP knijpen. De upload-,
template- en exportstappen werken prima in de cloud; het scrapen niet.

Zo start je lokaal:
    pip install -r requirements.txt
    streamlit run leadfinder_upload.py

Integratie in je bestaande app.py: importeer render_page() en roep 'm aan
binnen een eigen tab, of hergebruik de losse functies (find_emails_on_site,
verify_email, render_template). De verrijkingslaag is identiek voor
OSM-leads en geüploade leads, dus je pijplijn wordt één controlepunt.
"""

from __future__ import annotations

import io
import re
import time
import socket
import smtplib
from dataclasses import dataclass, field
from urllib.parse import urljoin, urlparse
from urllib import robotparser
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup

try:
    import dns.resolver  # dnspython
    _HAS_DNS = True
except Exception:  # pragma: no cover
    _HAS_DNS = False


# --------------------------------------------------------------------------- #
# E-mail extractie
# --------------------------------------------------------------------------- #

EMAIL_RE = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"
)

# Pagina's waar een contactadres vrijwel altijd staat (AVG dwingt privacy af).
CONTACT_HINTS = (
    "contact", "over-ons", "over_ons", "overons", "colofon",
    "privacy", "privacybeleid", "about", "team", "kontakt",
)

# Adressen/domeinen die we NOOIT als lead willen (assets, trackers, voorbeelden).
NOISE_SUBSTRINGS = (
    "example.", "sentry.", "wixpress.", "wix.com", "godaddy",
    "your-email", "yourdomain", "domain.com", "email.com", "test@",
    "@2x", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg",
    "u003e", "core-js", "react", "schema.org", "sentry-next",
)

# Veelvoorkomende obfuscaties -> normaliseren naar een echt adres.
_DEOBFUSCATE = [
    (re.compile(r"\s*\[\s*at\s*\]\s*", re.I), "@"),
    (re.compile(r"\s*\(\s*at\s*\)\s*", re.I), "@"),
    (re.compile(r"\s+at\s+", re.I), "@"),
    (re.compile(r"\s*\[\s*dot\s*\]\s*", re.I), "."),
    (re.compile(r"\s*\(\s*dot\s*\)\s*", re.I), "."),
    (re.compile(r"\s+dot\s+", re.I), "."),
]

DEFAULT_HEADERS = {
    "User-Agent": (
        "VoiceStampLeadFinder/1.0 (+contact via website; polite crawler)"
    )
}


def _clean_email(raw: str) -> str | None:
    e = raw.strip().strip(".,;:<>()[]\"'").lower()
    if not e or "@" not in e:
        return None
    if any(n in e for n in NOISE_SUBSTRINGS):
        return None
    m = EMAIL_RE.search(e)
    return m.group(0) if m else None


def _deobfuscate(text: str) -> str:
    for pat, repl in _DEOBFUSCATE:
        text = pat.sub(repl, text)
    return text


def _normalize_url(url: str) -> str | None:
    if not url or not isinstance(url, str):
        return None
    url = url.strip()
    if not url or url.lower() in ("nan", "none", "-"):
        return None
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    p = urlparse(url)
    if not p.netloc:
        return None
    return url


def domain_of(url_or_email: str) -> str | None:
    if not url_or_email:
        return None
    if "@" in url_or_email:
        return url_or_email.rsplit("@", 1)[-1].lower().strip()
    p = urlparse(_normalize_url(url_or_email) or "")
    host = p.netloc.lower()
    return host[4:] if host.startswith("www.") else host or None


def _robots_ok(url: str, cache: dict) -> bool:
    """Best-effort robots.txt-respect. Faalt open (True) bij twijfel."""
    try:
        p = urlparse(url)
        base = f"{p.scheme}://{p.netloc}"
        rp = cache.get(base)
        if rp is None:
            rp = robotparser.RobotFileParser()
            rp.set_url(urljoin(base, "/robots.txt"))
            try:
                rp.read()
            except Exception:
                rp = False  # kon robots niet lezen -> niet blokkeren
            cache[base] = rp
        if rp is False:
            return True
        return rp.can_fetch(DEFAULT_HEADERS["User-Agent"], url)
    except Exception:
        return True


def _emails_from_html(html: str) -> set[str]:
    found: set[str] = set()
    if not html:
        return found
    soup = BeautifulSoup(html, "lxml")
    # 1) mailto: links zijn het schoonst
    for a in soup.select("a[href^=mailto]"):
        href = a.get("href", "")
        addr = href.split("mailto:", 1)[-1].split("?", 1)[0]
        c = _clean_email(addr)
        if c:
            found.add(c)
    # 2) zichtbare tekst (met de-obfuscatie)
    text = _deobfuscate(soup.get_text(" ", strip=True))
    for m in EMAIL_RE.findall(text):
        c = _clean_email(m)
        if c:
            found.add(c)
    return found


def _candidate_contact_links(base_url: str, html: str, limit: int = 4) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    out, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith(("mailto:", "tel:", "#", "javascript:")):
            continue
        label = (a.get_text(" ", strip=True) + " " + href).lower()
        if any(h in label for h in CONTACT_HINTS):
            full = urljoin(base_url, href)
            if domain_of(full) == domain_of(base_url) and full not in seen:
                seen.add(full)
                out.append(full)
        if len(out) >= limit:
            break
    return out


def find_emails_on_site(
    website: str,
    session: requests.Session | None = None,
    max_pages: int = 5,
    timeout: int = 8,
    respect_robots: bool = True,
    robots_cache: dict | None = None,
) -> tuple[set[str], str]:
    """
    Crawlt homepage + een paar contactpagina's. Geeft (emails, bron) terug.
    bron in {"mailto/tekst", "geen"} — 'guess' wordt elders toegevoegd.
    """
    url = _normalize_url(website)
    if not url:
        return set(), "geen"
    sess = session or requests.Session()
    robots_cache = robots_cache if robots_cache is not None else {}
    emails: set[str] = set()
    to_visit = [url]
    visited: set[str] = set()

    while to_visit and len(visited) < max_pages:
        current = to_visit.pop(0)
        if current in visited:
            continue
        visited.add(current)
        if respect_robots and not _robots_ok(current, robots_cache):
            continue
        try:
            r = sess.get(current, headers=DEFAULT_HEADERS, timeout=timeout,
                         allow_redirects=True)
            if r.status_code >= 400 or "text/html" not in \
                    r.headers.get("Content-Type", "text/html"):
                continue
            html = r.text
        except Exception:
            continue
        emails |= _emails_from_html(html)
        if current == url:  # alleen vanaf de homepage extra links volgen
            for link in _candidate_contact_links(url, html):
                if link not in visited:
                    to_visit.append(link)
        time.sleep(0.4)  # beleefd

    return emails, ("mailto/tekst" if emails else "geen")


# --------------------------------------------------------------------------- #
# Verificatie (gratis): MX + optionele SMTP-probe
# --------------------------------------------------------------------------- #

@dataclass
class Verdict:
    email: str
    mx_ok: bool = False
    mx_host: str | None = None
    smtp_status: str = "niet gecheckt"   # ok / afgewezen / onbekend / geblokkeerd
    catch_all: bool = False
    confidence: str = "laag"             # hoog / midden / laag / dood


_mx_cache: dict[str, tuple[bool, str | None]] = {}


def mx_lookup(domain: str) -> tuple[bool, str | None]:
    if not _HAS_DNS or not domain:
        return (False, None)
    if domain in _mx_cache:
        return _mx_cache[domain]
    ok, host = False, None
    try:
        answers = dns.resolver.resolve(domain, "MX", lifetime=6)
        hosts = sorted((r.preference, str(r.exchange).rstrip(".")) for r in answers)
        if hosts:
            ok, host = True, hosts[0][1]
    except Exception:
        # geen MX? soms accepteert het A-record alsnog mail
        try:
            dns.resolver.resolve(domain, "A", lifetime=6)
            ok, host = True, domain
        except Exception:
            ok, host = False, None
    _mx_cache[domain] = (ok, host)
    return ok, host


def smtp_probe(email: str, mx_host: str, mail_from: str,
               timeout: int = 8) -> tuple[str, bool]:
    """
    Best-effort RCPT-check ZONDER iets te versturen. Geeft (status, catch_all).
    Werkt vaak niet vanaf cloud-IP's (poort 25 dicht) -> dan 'geblokkeerd'.
    """
    domain = email.rsplit("@", 1)[-1]
    try:
        server = smtplib.SMTP(timeout=timeout)
        server.connect(mx_host, 25)
        server.helo("voicestamp.example")
        server.mail(mail_from)
        code, _ = server.rcpt(email)
        # catch-all test: bestaat-vast-niet-adres ook geaccepteerd?
        catch = False
        probe = f"zzz-nietbestaand-{int(time.time())}@{domain}"
        try:
            c2, _ = server.rcpt(probe)
            catch = c2 in (250, 251)
        except Exception:
            pass
        server.quit()
        if code in (250, 251):
            return ("ok", catch)
        if code in (550, 551, 553, 501):
            return ("afgewezen", catch)
        return ("onbekend", catch)
    except (socket.timeout, ConnectionRefusedError, OSError):
        return ("geblokkeerd", False)
    except Exception:
        return ("onbekend", False)


def _score(v: Verdict, source: str) -> str:
    if not v.mx_ok:
        return "dood"
    if v.smtp_status == "afgewezen":
        return "dood"
    base = {"mailto/tekst": "hoog", "guess": "laag"}.get(source, "midden")
    if v.smtp_status == "ok" and not v.catch_all:
        base = "hoog"
    if v.catch_all and base == "hoog":
        base = "midden"
    return base


def verify_email(email: str, source: str, mail_from: str,
                 do_smtp: bool = False) -> Verdict:
    v = Verdict(email=email)
    domain = email.rsplit("@", 1)[-1]
    v.mx_ok, v.mx_host = mx_lookup(domain)
    if v.mx_ok and do_smtp and v.mx_host:
        v.smtp_status, v.catch_all = smtp_probe(email, v.mx_host, mail_from)
    v.confidence = _score(v, source)
    return v


# --------------------------------------------------------------------------- #
# Per-lead verrijking
# --------------------------------------------------------------------------- #

@dataclass
class EnrichResult:
    website: str
    email: str = ""
    email_source: str = "geen"     # mailto/tekst | guess | bestaand | geen
    mx_ok: bool = False
    smtp_status: str = "niet gecheckt"
    catch_all: bool = False
    confidence: str = "dood"
    all_emails: str = ""           # ; -gescheiden, voor de nieuwsgierigen


def enrich_row(website: str, existing_email: str, mail_from: str,
               do_smtp: bool, respect_robots: bool,
               session: requests.Session, robots_cache: dict) -> EnrichResult:
    res = EnrichResult(website=website or "")

    # 0) al een adres in de upload? gebruik dat, verifieer alleen.
    if existing_email and _clean_email(existing_email):
        e = _clean_email(existing_email)
        v = verify_email(e, "bestaand", mail_from, do_smtp)
        res.email, res.email_source = e, "bestaand"
        res.mx_ok, res.smtp_status = v.mx_ok, v.smtp_status
        res.catch_all, res.confidence = v.catch_all, v.confidence
        res.all_emails = e
        return res

    dom = domain_of(website or "")
    found: set[str] = set()
    source = "geen"
    if website:
        found, source = find_emails_on_site(
            website, session=session, respect_robots=respect_robots,
            robots_cache=robots_cache,
        )
    res.all_emails = "; ".join(sorted(found))

    # 1) gevonden adres op eigen domein wint
    chosen, chosen_src = None, "geen"
    if found:
        same = [e for e in found if domain_of(e) == dom] if dom else []
        chosen = sorted(same)[0] if same else sorted(found)[0]
        chosen_src = "mailto/tekst"
    # 2) fallback: info@domein construeren
    elif dom:
        chosen, chosen_src = f"info@{dom}", "guess"

    if not chosen:
        return res  # niks te doen

    v = verify_email(chosen, chosen_src, mail_from, do_smtp)
    res.email, res.email_source = chosen, chosen_src
    res.mx_ok, res.smtp_status = v.mx_ok, v.smtp_status
    res.catch_all, res.confidence = v.catch_all, v.confidence
    return res


# --------------------------------------------------------------------------- #
# Mail-merge
# --------------------------------------------------------------------------- #

_FIELD_RE = re.compile(r"\{([a-zA-Z0-9_]+)\}")


def render_template(template: str, row: dict) -> str:
    """Vervangt {kolomnaam} door de waarde. Onbekende velden blijven staan."""
    def repl(m):
        key = m.group(1)
        val = row.get(key, m.group(0))
        return "" if val is None else str(val)
    return _FIELD_RE.sub(repl, template or "")


def template_fields(template: str) -> list[str]:
    return sorted(set(_FIELD_RE.findall(template or "")))


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #

def render_page():  # pragma: no cover  (UI, niet los te testen)
    import pandas as pd
    import streamlit as st

    st.header("Leads uploaden, verrijken & mailen")
    st.caption(
        "Bron-neutraal: CSV/Excel uit OSM, een directory of een Gemini-lijst "
        "gaan door dezelfde verrijking en verificatie voordat er iets uitgaat."
    )

    up = st.file_uploader("CSV of Excel", type=["csv", "xlsx", "xls"])
    if not up:
        st.info("Upload een bestand om te beginnen.")
        return

    # inlezen (csv-delimiter automatisch raden)
    try:
        if up.name.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(up)
        else:
            df = pd.read_csv(up, sep=None, engine="python")
    except Exception as e:
        st.error(f"Kon het bestand niet lezen: {e}")
        return

    df.columns = [str(c).strip() for c in df.columns]
    st.write(f"**{len(df)} rijen** ingelezen.")
    st.dataframe(df.head(8), use_container_width=True)

    cols = list(df.columns)
    none = "— geen —"
    c1, c2, c3 = st.columns(3)
    name_col = c1.selectbox("Bedrijfsnaam", cols)
    web_col = c2.selectbox("Website / domein", cols,
                           index=_guess(cols, ("website", "url", "site", "web")))
    mail_col = c3.selectbox("Bestaand e-mailadres (optioneel)",
                            [none] + cols)

    with st.expander("Instellingen"):
        mail_from = st.text_input(
            "MAIL FROM voor verificatie (jouw verzendadres)",
            value="outreach@voicestamp.nl",
        )
        do_smtp = st.checkbox(
            "SMTP-probe (nauwkeuriger, maar traag en vaak geblokkeerd in de cloud)",
            value=False,
        )
        respect_robots = st.checkbox("robots.txt respecteren", value=True)
        workers = st.slider("Parallelle workers", 1, 12, 6)
        dedupe = st.checkbox("Ontdubbelen op domein", value=True)

    if st.button("Verrijk & verifieer", type="primary"):
        rows = df.to_dict("records")
        results: list[dict] = [None] * len(rows)
        robots_cache: dict = {}
        prog = st.progress(0.0, text="Bezig…")
        done = 0

        def work(i, row):
            sess = requests.Session()
            website = str(row.get(web_col, "") or "")
            existing = "" if mail_col == none else str(row.get(mail_col, "") or "")
            r = enrich_row(website, existing, mail_from, do_smtp,
                           respect_robots, sess, robots_cache)
            return i, r

        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = [ex.submit(work, i, row) for i, row in enumerate(rows)]
            for fut in as_completed(futs):
                i, r = fut.result()
                merged = dict(rows[i])
                merged.update({
                    "email": r.email,
                    "email_bron": r.email_source,
                    "mx_ok": r.mx_ok,
                    "smtp": r.smtp_status,
                    "catch_all": r.catch_all,
                    "confidence": r.confidence,
                    "alle_emails": r.all_emails,
                })
                results[i] = merged
                done += 1
                prog.progress(done / len(rows),
                              text=f"{done}/{len(rows)} verwerkt")

        out = pd.DataFrame(results)
        if dedupe:
            out["_dom"] = out["email"].apply(
                lambda e: e.rsplit("@", 1)[-1] if isinstance(e, str) and "@" in e else "")
            out = out.sort_values(
                "confidence",
                key=lambda s: s.map({"hoog": 0, "midden": 1, "laag": 2, "dood": 3}),
            ).drop_duplicates(subset="_dom", keep="first").drop(columns="_dom")

        st.session_state["enriched"] = out
        _summary(st, out)

    out = st.session_state.get("enriched")
    if out is None:
        return

    st.subheader("Resultaat")
    st.dataframe(out, use_container_width=True)

    # ---- Mailmerge ---- #
    st.subheader("Standaardmail")
    st.caption("Gebruik {kolomnaam} als variabele, bv. {" + name_col + "}.")
    subject = st.text_input("Onderwerp", value="Geen tweede app — juist minder drukte")
    body = st.text_area(
        "Bericht",
        height=220,
        value=(f"Hoi,\n\nIk zag {{{name_col}}} en vroeg me iets af...\n\n"
               "Groet,\nYusuf"),
    )
    used = [f for f in template_fields(subject + " " + body) if f in out.columns]
    st.caption("Herkende velden: " + (", ".join(used) or "geen"))

    if len(out):
        preview = out.iloc[0].to_dict()
        with st.expander("Voorbeeld (eerste lead)"):
            st.write("**Onderwerp:** " + render_template(subject, preview))
            st.text(render_template(body, preview))

    # alleen bruikbare adressen exporteren
    good = out[out["confidence"].isin(["hoog", "midden", "laag"])].copy()
    good["subject"] = good.apply(lambda r: render_template(subject, r.to_dict()), axis=1)
    good["body"] = good.apply(lambda r: render_template(body, r.to_dict()), axis=1)

    st.write(f"**{len(good)}** leads met bruikbaar adres (dood/geen weggelaten).")

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇︎ Volledig Excel",
        data=_to_excel(out),
        file_name="leads_verrijkt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    smart = good[["email", "subject", "body"] +
                 [c for c in (name_col,) if c in good.columns]].rename(
        columns={"email": "email", name_col: "company"})
    d2.download_button(
        "⬇︎ Smartlead/Instantly CSV",
        data=smart.to_csv(index=False).encode("utf-8"),
        file_name="leads_smartlead.csv",
        mime="text/csv",
    )

    with st.expander("⚠︎ Direct versturen via SMTP (alleen kleine test)"):
        st.caption(
            "Voor echte campagnes: gebruik de CSV in Smartlead/Instantly, waar je "
            "warmup en throttling geregeld hebt. Rauw versturen vanuit deze tool "
            "omzeilt dat en schaadt je deliverability. Let ook op art. 11.7 Tw."
        )
        host = st.text_input("SMTP-host", value="smtp.zoho.eu")
        port = st.number_input("Poort", value=587)
        user = st.text_input("Gebruiker")
        pw = st.text_input("Wachtwoord", type="password")
        n = st.number_input("Aantal (max 5 voor test)", 1, 5, 1)
        if st.button("Verstuur test"):
            _send_test(st, good.head(int(n)), host, int(port), user, pw)


def _guess(cols, keys):
    for i, c in enumerate(cols):
        if any(k in c.lower() for k in keys):
            return i
    return 0


def _summary(st, out):
    counts = out["confidence"].value_counts().to_dict()
    st.success(
        "Klaar. "
        f"hoog: {counts.get('hoog', 0)} · "
        f"midden: {counts.get('midden', 0)} · "
        f"laag: {counts.get('laag', 0)} · "
        f"dood/geen: {counts.get('dood', 0)}"
    )


def _to_excel(df) -> bytes:
    buf = io.BytesIO()
    with __import__("pandas").ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="leads")
    return buf.getvalue()


def _send_test(st, df, host, port, user, pw):  # pragma: no cover
    sent = 0
    try:
        s = smtplib.SMTP(host, port, timeout=15)
        s.starttls()
        s.login(user, pw)
        for _, r in df.iterrows():
            msg = (f"From: {user}\r\nTo: {r['email']}\r\n"
                   f"Subject: {r['subject']}\r\n\r\n{r['body']}")
            s.sendmail(user, [r["email"]], msg.encode("utf-8"))
            sent += 1
        s.quit()
        st.success(f"{sent} testmail(s) verstuurd.")
    except Exception as e:
        st.error(f"Versturen mislukt: {e}")


if __name__ == "__main__":
    render_page()
