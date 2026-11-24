"""
VoiceStamp Lead Finder - pipeline-module (geinspireerd op Sellbird)
===================================================================
Losse logica zonder Streamlit, zodat het testbaar is:

  1. Beslisser vinden      - eigenaar/manager van de eigen teampagina's van een bedrijf
  2. E-mailpatroon         - patroon leren van gepubliceerde adressen, kandidaat-adres bouwen
  3. Adres controleren     - MX + (optioneel) Reoon of ZeroBounce; elk adres krijgt een status
  4. Reeks plannen         - mail 1, 2 en 3 op werkdagen
  5. Controle voor verzenden ('gate') - zonder AI, blokkeert fouten
  6. Antwoorden            - IMAP lezen, bounce / afmelding / reactie herkennen

Alleen website-pagina's van het bedrijf zelf worden gelezen (geen LinkedIn e.d.).
De e-mailadres-statussen zijn dezelfde als bij Sellbird:
published (stond op hun eigen site), verified, risky, guess, invalid.
"""

import re, json, imaplib, ssl, html as _html, unicodedata
from datetime import date, timedelta
from email import message_from_bytes, policy
from email.header import decode_header, make_header
from email.utils import parseaddr

import requests

try:
    import dns.resolver
    HAVE_DNS = True
except Exception:
    HAVE_DNS = False

# ------------------------------------------------------------------ constanten
S_PUBLISHED, S_VERIFIED, S_RISKY, S_GUESS, S_INVALID = "published", "verified", "risky", "guess", "invalid"
STATUS_LABEL = {S_PUBLISHED: "Gepubliceerd", S_VERIFIED: "Geverifieerd", S_RISKY: "Risicovol",
                S_GUESS: "Gok", S_INVALID: "Ongeldig", None: "Onbekend", "": "Onbekend"}

FREE_MAIL = {"gmail.com", "hotmail.com", "hotmail.nl", "outlook.com", "outlook.nl", "yahoo.com",
             "yahoo.nl", "icloud.com", "live.nl", "live.com", "ziggo.nl", "kpnmail.nl", "planet.nl",
             "xs4all.nl", "me.com", "msn.com", "casema.nl", "home.nl", "online.nl", "telfort.nl"}

TEAM_PATHS = ["", "over-ons", "about", "team", "ons-team", "het-team", "wie-zijn-wij",
              "over-ons/team", "contact"]

ROLE_WEIGHT = {
    "eigenaar": 10, "eigenaresse": 10, "oprichter": 10, "oprichtster": 10, "founder": 10,
    "co-founder": 10, "owner": 10, "directeur": 9, "directrice": 9, "ceo": 9,
    "algemeen directeur": 9, "general manager": 8, "bedrijfsleider": 8, "manager": 7,
    "beheerder": 7, "campingbeheerder": 7, "hr manager": 7, "hr-manager": 7, "p&o": 6,
    "marketing": 5, "gastheer": 4, "gastvrouw": 4, "host": 4,
}
_ROLE_ALT = "|".join(sorted((re.escape(r) for r in ROLE_WEIGHT), key=len, reverse=True))

_NAMEPART = r"[A-ZÀ-Ý][a-zà-ÿ'’\-]{1,20}"
_PARTICLE = r"(?:van|de|der|den|ten|ter|te|von|op|'t)"
NAME_RE = rf"{_NAMEPART}(?:\s+(?:{_PARTICLE}\s+){{0,2}}{_NAMEPART}){{1,2}}"
_PAT_ROLE_FIRST = re.compile(rf"(?i:\b(?P<role>{_ROLE_ALT}))\s*[:\-–—,|]?\s*(?P<name>{NAME_RE})")
_PAT_NAME_FIRST = re.compile(rf"(?P<name>{NAME_RE})\s*[,\-–—|(]\s*(?i:(?P<role>{_ROLE_ALT}))\b")

NAME_STOP = {
    "camping", "het", "de", "een", "onze", "ons", "team", "bed", "breakfast", "hotel", "restaurant",
    "contact", "over", "welkom", "kamer", "boerderij", "vakantiepark", "mail", "telefoon", "tel",
    "email", "route", "open", "info", "home", "menu", "lees", "meer", "bekijk", "boek", "reserveren",
    "facebook", "instagram", "linkedin", "bezoek", "ons", "wij", "jullie", "ontdek", "ontvang",
    "neem", "bel", "stuur", "klik", "alle", "nieuw", "nieuws", "privacy", "cookie", "cookies",
    "algemene", "voorwaarden", "locatie", "adres", "kosten", "prijzen", "tarieven", "natuur",
    "park", "tuin", "huis", "winkel", "museum", "kasteel", "gastenverblijf",
}

FIRST_PERSON_JSONLD = ("Person",)

# ------------------------------------------------------------------ 1. beslisser vinden
def html_to_text(html):
    """Zichtbare tekst, met ' | ' tussen blokken zodat naam en functie los blijven."""
    html = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", html)
    t = re.sub(r"(?s)<[^>]+>", " | ", html)
    t = _html.unescape(t)
    t = re.sub(r"[ \t\r\n\u00a0]+", " ", t)
    t = re.sub(r"(\|\s*)+", "| ", t)
    return t

def _norm_ascii(s):
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z]", "", s.lower())

def split_name(full):
    """'Jan van der Berg' -> ('Jan', 'van der Berg')"""
    parts = full.split()
    return parts[0], " ".join(parts[1:])

def _naam_ok(name, company):
    toks = name.split()
    if len(toks) < 2:
        return False
    low = [t.lower().strip("'’-") for t in toks]
    if low[0] in NAME_STOP or low[-1] in NAME_STOP:
        return False
    comp = {c for c in re.findall(r"[a-zà-ÿ]{3,}", (company or "").lower())}
    if low[0] in comp:
        return False
    if sum(1 for t in low if t in NAME_STOP) >= 1 and len(toks) == 2:
        return False
    return True

def _jsonld_people(html):
    out = []
    for m in re.finditer(r'(?is)<script[^>]*application/ld\+json[^>]*>(.*?)</script>', html):
        try:
            data = json.loads(m.group(1).strip())
        except Exception:
            continue
        stack = [data]
        while stack:
            x = stack.pop()
            if isinstance(x, list):
                stack.extend(x)
            elif isinstance(x, dict):
                typ = x.get("@type")
                typs = typ if isinstance(typ, list) else [typ]
                if "Person" in typs and x.get("name"):
                    out.append((str(x["name"]).strip(), str(x.get("jobTitle", "")).strip()))
                for key in ("founder", "founders", "employee", "employees", "member", "author"):
                    if key in x:
                        stack.append(x[key])
                for v in x.values():
                    if isinstance(v, (list, dict)):
                        stack.append(v)
    return out

def extract_people(html, company=""):
    """Haal (naam, functie, gewicht) uit een pagina. Alleen functies uit ROLE_WEIGHT."""
    found = {}
    text = html_to_text(html)
    for pat in (_PAT_ROLE_FIRST, _PAT_NAME_FIRST):
        for m in pat.finditer(text):
            name, role = m.group("name").strip(), m.group("role").strip().lower()
            if not _naam_ok(name, company):
                continue
            w = ROLE_WEIGHT.get(role, 3)
            if name not in found or found[name][1] < w:
                found[name] = (role, w)
    for name, title in _jsonld_people(html):
        low = title.lower()
        role = next((r for r in sorted(ROLE_WEIGHT, key=len, reverse=True) if r in low), "")
        if role and _naam_ok(name, company):
            w = ROLE_WEIGHT[role] + 1  # gestructureerde data is betrouwbaarder
            if name not in found or found[name][1] < w:
                found[name] = (role, w)
    return sorted(((n, r, w) for n, (r, w) in found.items()), key=lambda x: -x[2])

def published_emails(html, domain=None):
    """Alle e-mailadressen die op de pagina staan (ook mailto en [at])."""
    bad = ("example.", "sentry.", "wixpress.", ".png", ".jpg", ".gif", "@2x", "your-email",
           "email@", "name@", "domain.com", "@sentry", "u003e")
    t = _html.unescape(html)
    t = re.sub(r"\s*[\[(]\s*(?:at|apenstaartje)\s*[\])]\s*", "@", t, flags=re.I)
    cands = re.findall(r"mailto:([^\"'?\s>]+)", t) + re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", t)
    out = []
    for c in cands:
        c = c.strip().strip(".").lower()
        if not re.fullmatch(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", c) or any(b in c for b in bad):
            continue
        if domain and c.split("@")[1] != domain:
            continue
        if c not in out:
            out.append(c)
    return out

def find_decision_makers(base, company, fetch, max_people=2):
    """
    base  : 'https://www.voorbeeld.nl' (zonder slash)
    fetch : callable(base, path) -> html-tekst of None  (netjes met robots.txt en pauzes)
    Geeft {'people': [(naam, functie)], 'emails': [adressen], 'domain': str, 'pages': int}
    """
    domain = re.sub(r"^www\.", "", re.sub(r"^https?://", "", base).split("/")[0])
    people, emails, pages = {}, [], 0
    for p in TEAM_PATHS:
        html = None
        try:
            html = fetch(base, p)
        except Exception:
            html = None
        if not html:
            continue
        pages += 1
        for e in published_emails(html, domain):
            if e not in emails:
                emails.append(e)
        for n, r, w in extract_people(html, company):
            if n not in people or people[n][1] < w:
                people[n] = (r, w)
        if pages >= 6:
            break
    best = sorted(people.items(), key=lambda kv: -kv[1][1])[:max_people]
    return {"people": [(n, r) for n, (r, w) in best], "emails": emails, "domain": domain, "pages": pages}

# ------------------------------------------------------------------ 2. e-mailpatroon
PATTERNS = ["first", "first.last", "flast", "f.last", "firstlast", "last"]

def build_local(pattern, first, last):
    f, l = _norm_ascii(first), _norm_ascii(last.split()[-1] if last else "")
    lj = _norm_ascii(last)
    if not f:
        return ""
    return {
        "first": f,
        "first.last": f"{f}.{lj}" if lj else "",
        "flast": f"{f[0]}{lj}" if lj else "",
        "f.last": f"{f[0]}.{lj}" if lj else "",
        "firstlast": f"{f}{lj}" if lj else "",
        "last": l,
    }.get(pattern, "")

def learn_pattern(emails):
    """Raad het patroon uit gepubliceerde persoonlijke adressen. None = niet te zien."""
    role_words = {"info", "contact", "welkom", "boeking", "boekingen", "reserveringen", "reservering",
                  "hallo", "mail", "receptie", "office", "sales", "administratie", "verkoop",
                  "support", "service", "klantenservice", "bureau", "post", "secretariaat"}
    votes = {}
    for e in emails:
        lok = e.split("@")[0]
        if lok in role_words:
            continue
        if re.fullmatch(r"[a-z]{3,14}", lok):
            k = "first"
        elif re.fullmatch(r"[a-z]{2,}\.[a-z]{2,}", lok):
            k = "first.last"
        elif re.fullmatch(r"[a-z]\.[a-z]{2,}", lok):
            k = "f.last"
        elif re.fullmatch(r"[a-z][a-z]{3,}", lok) and len(lok) > 7:
            k = "firstlast"
        else:
            continue
        votes[k] = votes.get(k, 0) + 1
    return max(votes, key=votes.get) if votes else None

def email_candidates(person, domain, known_emails=(), pattern=None):
    """
    Geeft [(adres, basis)] op volgorde van waarschijnlijkheid.
    basis: 'gepubliceerd' (stond zo op de site), 'patroon' (volgt hun patroon), 'gok'.
    """
    first, last = split_name(person)
    pattern = pattern or learn_pattern(known_emails)
    order = ([pattern] if pattern else []) + [p for p in ("first", "first.last", "flast", "f.last") if p != pattern]
    out, seen = [], set()
    for pat in order:
        lok = build_local(pat, first, last)
        if not lok:
            continue
        addr = f"{lok}@{domain}"
        if addr in seen:
            continue
        seen.add(addr)
        basis = "gepubliceerd" if addr in known_emails else ("patroon" if pat == pattern else "gok")
        out.append((addr, basis))
    return out

# ------------------------------------------------------------------ 3. adres controleren
def has_mx(email_or_domain):
    if not HAVE_DNS:
        return None
    dom = email_or_domain.split("@")[-1]
    try:
        return len(dns.resolver.resolve(dom, "MX")) > 0
    except Exception:
        return False

def verify_email(email, reoon_key=None, zerobounce_key=None, timeout=20, http=requests):
    """
    Geeft (status, toelichting). Zonder dienst kun je alleen MX controleren:
    dan is het hoogstens een 'guess'. Een adres dat op de eigen site stond blijft 'published'
    (dat bepaalt de aanroeper).
    LET OP: de antwoordformaten van Reoon en ZeroBounce zijn hier niet live getest.
    """
    mx = has_mx(email)
    if mx is False:
        return S_INVALID, "domein ontvangt geen mail (geen MX)"
    try:
        if reoon_key:
            r = http.get("https://emailverifier.reoon.com/api/v1/verify",
                         params={"email": email, "key": reoon_key, "mode": "power"}, timeout=timeout)
            st = str(r.json().get("status", "")).lower()
            if st in ("safe", "valid"):
                return S_VERIFIED, "Reoon: veilig"
            if st in ("catch_all", "role_account", "unknown", "inbox_full"):
                return S_RISKY, f"Reoon: {st}"
            return S_INVALID, f"Reoon: {st or 'onbekend antwoord'}"
        if zerobounce_key:
            r = http.get("https://api.zerobounce.net/v2/validate",
                         params={"api_key": zerobounce_key, "email": email}, timeout=timeout)
            st = str(r.json().get("status", "")).lower()
            if st == "valid":
                return S_VERIFIED, "ZeroBounce: geldig"
            if st in ("catch-all", "unknown"):
                return S_RISKY, f"ZeroBounce: {st}"
            return S_INVALID, f"ZeroBounce: {st or 'onbekend antwoord'}"
    except Exception as e:
        return S_GUESS, f"verificatiedienst niet bereikbaar ({e.__class__.__name__})"
    return S_GUESS, "geen verificatiedienst ingesteld"

def pick_contact_email(candidates, published, reoon_key=None, zerobounce_key=None, max_checks=2):
    """
    Kies het beste adres voor de persoon. Een adres dat op hun site stond telt als published.
    Anders: verifieer maximaal `max_checks` kandidaten (zoals Sellbird: een of twee, niet allemaal).
    Geeft (adres, status, toelichting) of ('', None, 'geen kandidaat').
    """
    if not candidates:
        return "", None, "geen kandidaat"
    for addr, basis in candidates:
        if addr in published:
            return addr, S_PUBLISHED, "stond op hun eigen website"
    if not (reoon_key or zerobounce_key):
        addr, basis = candidates[0]
        st, why = verify_email(addr)
        return addr, (S_INVALID if st == S_INVALID else S_GUESS), f"{basis}; {why}"
    last = ("", None, "geen geldig adres gevonden")
    for addr, basis in candidates[:max_checks]:
        st, why = verify_email(addr, reoon_key, zerobounce_key)
        if st in (S_VERIFIED, S_RISKY):
            return addr, st, f"{basis}; {why}"
        last = (addr, st, f"{basis}; {why}")
    return last

# ------------------------------------------------------------------ 4. reeks plannen
def volgende_werkdag(d):
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def plan_dates(start, offsets=(0, 4, 8)):
    """Mail 1, 2 en 3: standaard op dag 0, 4 en 8, altijd naar een werkdag geschoven."""
    return [volgende_werkdag(start + timedelta(days=o)) for o in offsets]

# ------------------------------------------------------------------ 5. gate voor verzenden
STOP_STATUS_ALTIJD = {"Do not contact", "Afgewezen"}
STOP_STATUS_VERVOLG = {"Reactie", "Afspraak", "Klant"}
ADRES_OK = {S_PUBLISHED, S_VERIFIED}

OWN_LINK_PREFIXES = ("https://www.voicestamp.nl", "https://voicestamp.nl", "http://www.voicestamp.nl",
                     "https://www.instagram.com/voicestamp.nl", "https://www.linkedin.com/company/voicestamp")
BANNED = ("ik hoop dat deze mail je goed vindt", "in de huidige snel veranderende", "revolutionair",
          "baanbrekend", "game-changer", "game changer", "synergie", "naadloos", "unlock",
          "ik zal niet om de hete brij heen draaien", "hopelijk gaat het goed met je")
OPTOUT_RE = re.compile(r"afmeld|uitschrijv|geen mail|niet meer mailen|liever geen|laat.*weten.*niet", re.I)
TAG_RE = re.compile(r"\{[a-zA-Z_]+\}|\[[A-Za-z _]{2,30}\]|<<[^>]+>>")
EMAIL_SYNTAX = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def gate(item, lead_status="Nieuw", email_status=None, allow_risky=False, mx_check=True,
         own_prefixes=OWN_LINK_PREFIXES, max_words=220, full_text=None):
    """
    item: dict met to_email, subject, body, step (1..3).
    full_text: de volledige mail incl. handtekening (voor de afmeldregel-check), optioneel.
    Geeft (fouten, waarschuwingen). Bij fouten wordt NIET verstuurd.
    """
    err, warn = [], []
    to = (item.get("to_email") or "").strip()
    step = int(item.get("step") or 1)
    body, subj = item.get("body") or "", item.get("subject") or ""
    if not EMAIL_SYNTAX.match(to):
        err.append("Ongeldig e-mailadres.")
    if not subj.strip() or not body.strip():
        err.append("Onderwerp of tekst is leeg.")
    if lead_status in STOP_STATUS_ALTIJD:
        err.append(f"Lead staat op '{lead_status}'.")
    if step > 1 and lead_status in STOP_STATUS_VERVOLG:
        err.append(f"Lead heeft al gereageerd of is klant ('{lead_status}'): reeks gestopt.")
    st = email_status
    if st == S_INVALID:
        err.append("Adres is als ongeldig gemarkeerd.")
    elif st == S_GUESS:
        err.append("Adres is een gok en niet geverifieerd: wordt niet verstuurd.")
    elif st == S_RISKY and not allow_risky:
        err.append("Adres is risicovol (catch-all/onzeker). Zet 'risicovolle adressen toestaan' aan als je dat wilt.")
    if TAG_RE.search(subj + "\n" + body):
        err.append("Niet ingevulde plaatshouder in de tekst: " + TAG_RE.search(subj + "\n" + body).group(0))
    if mx_check and EMAIL_SYNTAX.match(to) and has_mx(to) is False:
        err.append("Het domein van dit adres ontvangt geen mail.")
    low = (subj + " " + body).lower()
    for b in BANNED:
        if b in low:
            warn.append(f"Standaard AI-zin gevonden: '{b}'.")
    words = len(re.findall(r"\w+", body))
    if words > max_words:
        warn.append(f"Lange mail ({words} woorden). Korter wordt vaker gelezen.")
    for u in re.findall(r"https?://[^\s<>()]+", body):
        if not u.startswith(own_prefixes):
            warn.append(f"Link naar een ander domein: {u}")
    if full_text is not None and not OPTOUT_RE.search(full_text):
        warn.append("Geen afmeldregel in de mail. Voor zakelijke mail is een eenvoudige afmeldmogelijkheid verplicht.")
    return err, warn

# ------------------------------------------------------------------ 6. antwoorden
BOUNCE_FROM = re.compile(r"mailer-daemon|postmaster|mail delivery|noreply.*google", re.I)
BOUNCE_SUBJ = re.compile(r"undeliver|delivery status|delivery failure|failure notice|returned mail|"
                         r"niet bezorgd|kon niet worden bezorgd|bezorging mislukt|mail delivery failed|"
                         r"address not found|adres niet gevonden", re.I)
AUTO_SUBJ = re.compile(r"automatisch antwoord|automatic reply|auto.?reply|out of office|afwezigheid|"
                       r"afwezig|buiten kantoor|vakantie|autoreply", re.I)
OPTOUT = re.compile(r"afmelden|uitschrijven|schrijf me uit|schrijf ons uit|verwijder (?:mij|ons|me)|"
                    r"unsubscribe|stop met mailen|niet meer mailen|geen mails? meer|geen e-?mails? meer|"
                    r"haal (?:mij|ons|me) van (?:de|jullie) lijst|laat mij met rust", re.I)
NEGATIVE = re.compile(r"geen interesse|niet ge[iï]nteresseerd|past niet bij ons|geen behoefte|"
                      r"hebben we niet nodig|niet van toepassing|bedankt maar nee|nee,? bedankt|"
                      r"op dit moment niet|geen budget", re.I)
POSITIVE = re.compile(r"interessant|ge[iï]nteresseerd|graag|laat maar zien|vertel (?:me )?meer|"
                      r"meer informatie|stuur (?:me |ons )?(?:maar )?(?:wat )?(?:info|informatie)|"
                      r"afspraak|inplannen|bel (?:me|ons|mij)|demo|wanneer kun|tijd voor|kennismaken|"
                      r"klinkt goed|klinkt leuk|leuk idee", re.I)
SCHEDULE = re.compile(r"afspraak|inplannen|kennismaken|bel (?:me|ons|mij)|wanneer kun|tijd voor", re.I)

def strip_quoted(text):
    out = []
    for line in (text or "").splitlines():
        if re.match(r"^\s*(Op .{3,80} schreef|On .{3,80} wrote|-{3,}\s*(Original|Oorspronkelijk)|"
                    r"Van:|From:|Verzonden:|Sent:)", line, re.I):
            break
        if line.lstrip().startswith(">"):
            continue
        out.append(line)
    return "\n".join(out).strip()

def classify_reply(subject="", body="", from_addr="", auto_submitted="", our_address=""):
    """
    Geeft (soort, status_voor_crm, uitleg).
    soort: bounce | auto | optout | negative | positive | neutral
    status_voor_crm: None (niets wijzigen), 'Reactie', 'Afgewezen' of 'Do not contact'
    """
    subj, fresh = subject or "", strip_quoted(body or "")
    if BOUNCE_FROM.search(from_addr or "") or BOUNCE_SUBJ.search(subj):
        return "bounce", None, "bezorgmelding: het adres bestaat niet of is vol"
    if (auto_submitted or "").lower() not in ("", "no") or AUTO_SUBJ.search(subj):
        return "auto", None, "automatisch antwoord (niet beantwoorden)"
    if OPTOUT.search(fresh):
        return "optout", "Do not contact", "vraagt om geen mail meer te sturen"
    if NEGATIVE.search(fresh) and not SCHEDULE.search(fresh):
        return "negative", "Afgewezen", "geen interesse"
    if POSITIVE.search(fresh):
        return "positive", "Reactie", "positieve of vragende reactie"
    return "neutral", "Reactie", "reactie, lees zelf even mee"

def failed_recipient(body, our_address=""):
    """Zoek bij een bounce het adres dat niet bezorgd kon worden."""
    for m in re.finditer(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", body or ""):
        a = m.group(0).lower().strip(".")
        if our_address and a == our_address.lower():
            continue
        if re.match(r"(mailer-daemon|postmaster)@", a):
            continue
        return a
    return ""

def imap_host_for(smtp_host):
    h = (smtp_host or "").strip().lower()
    return {"smtp.zoho.eu": "imap.zoho.eu", "smtp.zoho.com": "imap.zoho.com",
            "smtppro.zoho.eu": "imappro.zoho.eu", "smtppro.zoho.com": "imappro.zoho.com"}.get(
        h, h.replace("smtp", "imap", 1))

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def _imap_date(d):
    return f"{d.day:02d}-{_MONTHS[d.month - 1]}-{d.year}"

def _decode(h):
    try:
        return str(make_header(decode_header(h or "")))
    except Exception:
        return h or ""

def _body_text(msg):
    plain, html = None, None
    parts = msg.walk() if msg.is_multipart() else [msg]
    for p in parts:
        ct = p.get_content_type()
        if p.get_content_disposition() == "attachment":
            continue
        try:
            payload = p.get_content()
        except Exception:
            continue
        if ct == "text/plain" and plain is None and isinstance(payload, str):
            plain = payload
        elif ct == "text/html" and html is None and isinstance(payload, str):
            html = payload
    if plain:
        return plain
    return re.sub(r"\s+", " ", re.sub(r"(?s)<[^>]+>", " ", _html.unescape(html or ""))).strip()

def fetch_replies(host, user, password, since_days=14, max_messages=150, port=993, timeout=25):
    """
    Leest recente mails uit je inbox (alleen lezen, markeert niets als gelezen).
    Geeft een lijst dicts: from, subject, date, message_id, in_reply_to, auto, body.
    Vereist dat IMAP aan staat voor je account (bij Zoho: niet op het gratis pakket).
    """
    ctx = ssl.create_default_context()
    M = imaplib.IMAP4_SSL(host, port, ssl_context=ctx, timeout=timeout)
    try:
        M.login(user.strip(), password.replace(" ", "").strip())
        M.select("INBOX", readonly=True)
        typ, data = M.search(None, "SINCE", _imap_date(date.today() - timedelta(days=since_days)))
        ids = (data[0] or b"").split()[-max_messages:]
        out = []
        for i in ids:
            typ, md = M.fetch(i, "(BODY.PEEK[])")
            if typ != "OK" or not md or not isinstance(md[0], tuple):
                continue
            msg = message_from_bytes(md[0][1], policy=policy.default)
            out.append({
                "from": parseaddr(str(msg.get("From", "")))[1].lower(),
                "subject": _decode(str(msg.get("Subject", ""))),
                "date": str(msg.get("Date", "")),
                "message_id": str(msg.get("Message-ID", "")).strip(),
                "in_reply_to": str(msg.get("In-Reply-To", "")).strip(),
                "auto": str(msg.get("Auto-Submitted", "")).strip(),
                "body": _body_text(msg)[:6000],
            })
        return out
    finally:
        try:
            M.logout()
        except Exception:
            pass

def match_lead(addr, leads):
    """
    Koppel afzender aan een lead: eerst exact adres, dan hetzelfde bedrijfsdomein
    (een antwoord van iemand anders bij hetzelfde bedrijf telt ook, zoals bij Sellbird).
    leads: lijst dicts met id, email, contact_email, website.
    """
    addr = (addr or "").lower()
    if not addr or "@" not in addr:
        return None
    for l in leads:
        if addr in {(l.get("email") or "").lower(), (l.get("contact_email") or "").lower()}:
            return l
    dom = addr.split("@")[1]
    if dom in FREE_MAIL:
        return None
    for l in leads:
        doms = set()
        for e in (l.get("email"), l.get("contact_email")):
            if e and "@" in e:
                doms.add(e.split("@")[1].lower())
        w = re.sub(r"^www\.", "", re.sub(r"^https?://", "", (l.get("website") or "").lower()).split("/")[0])
        if w:
            doms.add(w)
        doms -= FREE_MAIL
        if dom in doms:
            return l
    return None
