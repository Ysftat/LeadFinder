"""
VoiceStamp Lead Finder — Streamlit web-app (snelle versie)
==========================================================
Nieuw t.o.v. de vorige versie:
  - VEEL SNELLER: websites worden parallel verwerkt (meerdere tegelijk).
  - MEER LEADS: bredere OpenStreetMap-categorieen per campagne (gratis, legaal).
  - Social links: Instagram / LinkedIn / Facebook per lead (handig voor je LinkedIn-aanpak).
  - OSM-cache: dezelfde provincie opnieuw ophalen gaat direct.
  - Live voortgang: je ziet meteen wat er gevonden wordt.

Lokaal: pip install streamlit requests dnspython openpyxl pandas ; streamlit run app.py
Online: zie LEES_MIJ_streamlit.md
"""

import re, io, time, csv, hashlib, random
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import dns.resolver
    HAVE_DNS = True
except Exception:
    HAVE_DNS = False

# Sellbird-onderdelen (beslisser vinden, adres controleren, reeks, gate, antwoorden)
try:
    import sellbird_pipeline as SP
    SP_OK = True
except Exception:
    SP = None
    SP_OK = False

# Campagnes met hoofdgroepen en subcategorieën (campagnes.py moet naast app.py staan)
try:
    import campagnes as CP
    CP_OK = True
except Exception:
    CP = None
    CP_OK = False

# ------------------------------------------------------------------ CRM (Supabase, optioneel)
try:
    from supabase import create_client as _create_client
    SUPABASE_LIB = True
except Exception:
    SUPABASE_LIB = False

STATUS_OPTIES = ["Nieuw", "Gemaild", "Reactie", "Afspraak", "Klant", "Afgewezen", "Do not contact"]

def crm_beschikbaar():
    if not SUPABASE_LIB:
        return False
    try:
        return bool(st.secrets.get("SUPABASE_URL")) and bool(st.secrets.get("SUPABASE_ANON_KEY"))
    except Exception:
        return False

def crm_auth(email, password, registreren=False):
    client = _create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_ANON_KEY"])
    if registreren:
        res = client.auth.sign_up({"email": email, "password": password})
    else:
        res = client.auth.sign_in_with_password({"email": email, "password": password})
    sess = getattr(res, "session", None)
    if sess and getattr(sess, "access_token", None):
        try:
            client.postgrest.auth(sess.access_token)
        except Exception:
            pass
    return client, getattr(res, "user", None)

def crm_save(client, user_id, leads):
    bestaand = set()
    try:
        r = client.table("leads").select("email").eq("user_id", user_id).execute()
        bestaand = {row["email"] for row in (r.data or []) if row.get("email")}
    except Exception:
        pass
    rows = []
    for l in leads:
        e = l.get("email")
        if not e or e in bestaand:
            continue
        bestaand.add(e)
        rows.append({"user_id": user_id, "company_name": l["naam"], "segment": l["type"],
                     "website": l.get("website", ""), "email": e, "phone": l.get("telefoon", ""),
                     "score": l.get("_score"), "hook": l.get("haak", ""),
                     "channel_advice": l.get("_kanaal", ""), "status": "Nieuw"})
    if rows:
        client.table("leads").insert(rows).execute()
    return len(rows)

def crm_load(client, user_id):
    r = client.table("leads").select("*").eq("user_id", user_id).order("score", desc=True).execute()
    return r.data or []

def crm_update(client, row_id, user_id, fields):
    # Defensief: filter op id EN user_id (bovenop row-level security in Supabase).
    client.table("leads").update(fields).eq("id", row_id).eq("user_id", user_id).execute()

def crm_delete(client, row_id, user_id):
    client.table("leads").delete().eq("id", row_id).eq("user_id", user_id).execute()

def crm_delete_all(client, user_id):
    r = client.table("leads").delete().eq("user_id", user_id).execute()
    return len(r.data or [])

def crm_add_activity(client, user_id, lead_id, type_, description=""):
    try:
        client.table("activities").insert({
            "user_id": user_id, "lead_id": lead_id, "type": type_, "description": description
        }).execute()
    except Exception:
        pass

def crm_load_activities(client, user_id, lead_id):
    try:
        r = (client.table("activities").select("*")
             .eq("user_id", user_id).eq("lead_id", lead_id)
             .order("created_at", desc=True).execute())
        return r.data or []
    except Exception:
        return []

def crm_lead_id_by_email(client, user_id, email):
    try:
        r = client.table("leads").select("id").eq("user_id", user_id).eq("email", email).limit(1).execute()
        if r.data:
            return r.data[0]["id"]
    except Exception:
        pass
    return None

def crm_sent_today(client, user_id):
    """Aantal 'Mail verstuurd'-activiteiten van vandaag (voor de dagelijkse limiet)."""
    from datetime import date
    try:
        r = (client.table("activities").select("id")
             .eq("user_id", user_id).eq("type", "Mail verstuurd")
             .gte("created_at", date.today().isoformat()).execute())
        return len(r.data or [])
    except Exception:
        return 0

# ---- Outbox (mailreeks per lead): helpers
def ob_create(client, user_id, rows):
    client.table("outbox").insert([dict(r, user_id=user_id) for r in rows]).execute()

def ob_load(client, user_id):
    r = (client.table("outbox").select("*").eq("user_id", user_id)
         .order("scheduled_for").order("step").execute())
    return r.data or []

def ob_update(client, user_id, row_id, fields):
    client.table("outbox").update(fields).eq("id", row_id).eq("user_id", user_id).execute()

def ob_cancel_for_lead(client, user_id, lead_id, reden):
    """Stop alle nog niet verstuurde mails van deze lead (bij antwoord, afmelding of bounce)."""
    try:
        (client.table("outbox").update({"status": "cancelled", "error": reden})
         .eq("user_id", user_id).eq("lead_id", lead_id)
         .in_("status", ["pending_review", "approved"]).execute())
    except Exception:
        pass

# ---------------------------------------------------------------- Zoho-mail (SMTP)
import smtplib, ssl, html as _html
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid

MAIL_DAG_LIMIET = 15

_URL_RE = re.compile(r"(https?://[^\s<>()]+|www\.[^\s<>()]+)")

def _naar_html(tekst):
    """Zet platte tekst om naar simpele, nette HTML met klikbare links (sober = beter voor spam)."""
    veilig = _html.escape(tekst)
    def _link(m):
        u = m.group(0)
        href = u if u.startswith("http") else "https://" + u
        return f'<a href="{href}">{u}</a>'
    veilig = _URL_RE.sub(_link, veilig)
    veilig = veilig.replace("\n", "<br>")
    return (f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
            f'line-height:1.5;color:#222;">{veilig}</div>')

def zoho_send(host, port, user, app_pw, from_name, to_addr, subject, body, in_reply_to=None):
    """Verstuurt de mail en geeft het Message-ID terug (voor het koppelen van opvolgmails)."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name, user))
    msg["To"] = to_addr
    msg_id = make_msgid(domain=user.split("@")[-1])
    msg["Message-ID"] = msg_id
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = in_reply_to
    msg.attach(MIMEText(body, "plain", "utf-8"))
    msg.attach(MIMEText(_naar_html(body), "html", "utf-8"))
    ctx = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(host, int(port), context=ctx, timeout=20) as s:
            s.login(user.strip(), app_pw.replace(" ", "").strip())
            s.sendmail(user.strip(), [to_addr], msg.as_string())
        return msg_id
    except smtplib.SMTPAuthenticationError:
        raise RuntimeError(
            f"Zoho weigert de inlog op {host}. Meestal is het: (1) je gebruikt je gewone wachtwoord "
            "in plaats van een app-wachtwoord, (2) de server klopt niet met je account (probeer "
            "smtp.zoho.com, of smtppro.zoho.eu / smtppro.zoho.com voor een bedrijfsaccount), "
            "(3) het app-wachtwoord is ingetrokken of verkeerd gekopieerd, of (4) SMTP staat niet "
            "aan voor je account.")

# Koppel Nederlandse (of Engelse) kolomnamen aan de CRM-velden.
IMPORT_MAP = {
    "company_name": ["naam", "bedrijf", "company_name", "company", "bedrijfsnaam"],
    "email": ["e-mail", "email", "e-mail (beste ingang)", "mail", "e-mailadres"],
    "segment": ["type", "segment", "categorie"],
    "website": ["website", "url", "site"],
    "phone": ["telefoon", "phone", "tel", "telefoonnummer"],
    "status": ["status"],
    "note": ["notities", "notitie", "note", "opmerkingen"],
}

def _kies_kolom(kolommen_laag, opties):
    for o in opties:
        if o in kolommen_laag:
            return kolommen_laag[o]
    return None

def crm_import_uit_bytes(client, user_id, data_bytes, filename):
    """Lees een geuploade Excel/CSV en zet de leads in het CRM. Geeft (toegevoegd, overgeslagen)."""
    if filename.lower().endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(io.BytesIO(data_bytes))
    else:
        df = pd.read_csv(io.BytesIO(data_bytes))
    # kolomnamen -> {kleine letter: originele naam}
    kol = {str(c).strip().lower(): c for c in df.columns}
    mapping = {veld: _kies_kolom(kol, opties) for veld, opties in IMPORT_MAP.items()}
    if not mapping.get("email"):
        raise ValueError("Geen e-mailkolom gevonden. Zorg dat er een kolom 'E-mail' in staat.")
    # bestaande e-mails ophalen om dubbel te voorkomen
    bestaand = set()
    try:
        r = client.table("leads").select("email").eq("user_id", user_id).execute()
        bestaand = {row["email"] for row in (r.data or []) if row.get("email")}
    except Exception:
        pass
    rows, overgeslagen = [], 0
    for _, rij in df.iterrows():
        email = str(rij.get(mapping["email"], "") or "").strip().lower()
        if not email or "@" not in email or email in bestaand:
            overgeslagen += 1
            continue
        bestaand.add(email)
        def val(veld):
            c = mapping.get(veld)
            if not c: return ""
            v = rij.get(c, "")
            return "" if pd.isna(v) else str(v).strip()
        rows.append({"user_id": user_id, "company_name": val("company_name"), "email": email,
                     "segment": val("segment"), "website": val("website"), "phone": val("phone"),
                     "status": val("status") or "Nieuw", "note": val("note")})
    # in blokjes wegschrijven
    for i in range(0, len(rows), 100):
        client.table("leads").insert(rows[i:i+100]).execute()
    return len(rows), overgeslagen


# ------------------------------------------------------------------ constants
OVERPASS = ["https://overpass-api.de/api/interpreter",
            "https://overpass.kumi.systems/api/interpreter",
            "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
            "https://overpass.openstreetmap.ru/api/interpreter",
            "https://overpass.private.coffee/api/interpreter"]
UA = "VoiceStamp-LeadFinder/streamlit (zakelijk gebruik; contact via voicestamp.nl)"
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
BAD = ("example.", "sentry.", "wixpress.", ".png", ".jpg", ".gif", "@2x",
       "your-email", "email@", "name@", "domain.com", "@sentry")
# Kortere padlijst = sneller. Homepage + de paar meest waarschijnlijke contactpaginas.
PATHS = ["", "contact", "over-ons", "info", "reserveren", "about"]
ROLE = ("info", "contact", "welkom", "boeking", "boekingen", "reserveringen",
        "reservering", "hallo", "mail", "receptie", "office", "sales")
KETEN = ("hotels", "resorts", "group", "vakantieparken", "landal", "roompot",
         "rcn", "huttopia", "fletcher", "ardoer", "europarcs", "oostappen")
SOCIAL_RE = {
    "instagram": re.compile(r"https?://(?:www\.)?instagram\.com/[A-Za-z0-9_.]+"),
    "linkedin": re.compile(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/(?:company|in)/[A-Za-z0-9_%\-]+"),
    "facebook": re.compile(r"https?://(?:www\.)?facebook\.com/[A-Za-z0-9_.\-]+"),
}

# --- Netjes blijven tegenover de bronnen: rate limiting + robots + retries ---
import threading
from urllib import robotparser

_domain_lock = threading.Lock()
_last_hit = {}          # domein -> laatste tijdstip
_robots_cache = {}      # domein -> RobotFileParser (of None)
MIN_INTERVAL = 1.0      # minimaal 1 seconde tussen verzoeken naar hetzelfde domein

def _respect_delay(domain):
    """Wacht indien nodig zodat we niet te snel achter elkaar hetzelfde domein raken."""
    with _domain_lock:
        now = time.time()
        wait = MIN_INTERVAL - (now - _last_hit.get(domain, 0))
        if wait > 0:
            time.sleep(wait)
        _last_hit[domain] = time.time()

def _robots_ok(base, path):
    """Controleer robots.txt; bij twijfel toestaan (maar netjes blijven qua tempo)."""
    domain = urlparse(base).netloc
    rp = _robots_cache.get(domain, "missing")
    if rp == "missing":
        rp = robotparser.RobotFileParser()
        try:
            rp.set_url(f"{urlparse(base).scheme}://{domain}/robots.txt")
            rp.read()
        except Exception:
            rp = None
        _robots_cache[domain] = rp
    if rp is None:
        return True
    try:
        return rp.can_fetch(UA, f"{base}/{path}".rstrip("/"))
    except Exception:
        return True

def polite_get(url, domain, timeout=8, retries=2):
    """GET met rate limiting en een enkele nette retry bij een tijdelijke fout."""
    for attempt in range(retries):
        _respect_delay(domain)
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=timeout)
            if r.status_code in (429, 503) and attempt < retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
            return r
        except Exception:
            if attempt < retries - 1:
                time.sleep(1)
                continue
            raise
    return None
PROVINCIES = ["Groningen", "Fryslân", "Drenthe", "Overijssel", "Flevoland",
              "Gelderland", "Utrecht", "Noord-Holland", "Zuid-Holland",
              "Zeeland", "Noord-Brabant", "Limburg"]

# Bredere presets = meer leads. Elke osm-value krijgt (label, template-segment).
PRESETS = {
    "Verblijf (campings, hotels, B&B's)": {
        "typemap": {
            "camp_site": ("Camping", "natuurcamping"), "caravan_site": ("Camperplaats", "natuurcamping"),
            "guest_house": ("B&B / guest house", "bnb"), "hotel": ("Hotel", "hotel"),
            "hostel": ("Hostel", "hotel"), "chalet": ("Chalet/huisjes", "natuurcamping"),
            "motel": ("Motel", "hotel"), "apartment": ("Vakantieappartement", "bnb"),
            "alpine_hut": ("Berghut/verblijf", "natuurcamping"),
        },
        "keys": {"tourism": ["camp_site", "caravan_site", "guest_house", "hotel",
                             "hostel", "chalet", "motel", "apartment", "alpine_hut"]},
    },
    "Streek (wijn, kaas, brouwers, boerderijwinkels)": {
        "typemap": {
            "winery": ("Wijngaard", "streek"), "brewery": ("Brouwerij", "streek"),
            "distillery": ("Distilleerderij", "streek"), "cheese_making": ("Kaasmakerij", "streek"),
            "wine": ("Wijnhandel", "streek"), "cheese": ("Kaaswinkel", "streek"),
            "farm": ("Boerderijwinkel", "streek"), "dairy": ("Zuivel/boerderij", "streek"),
            "greengrocer": ("Groente/streekwinkel", "streek"), "deli": ("Delicatessen", "streek"),
            "bakery": ("Bakkerij", "streek"), "butcher": ("Slagerij", "streek"),
            "confectionery": ("Chocolaterie/snoep", "streek"), "chocolate": ("Chocolatier", "streek"),
            "honey": ("Imker/honing", "streek"),
        },
        "keys": {"craft": ["winery", "brewery", "distillery", "cheese_making"],
                 "shop": ["wine", "cheese", "farm", "dairy", "greengrocer", "deli",
                          "bakery", "butcher", "confectionery", "chocolate", "honey"]},
    },
    "Attracties (dierentuinen, kinderboerderijen, parken)": {
        "typemap": {
            "zoo": ("Dierentuin/kinderboerderij", "attractie"),
            "theme_park": ("Attractiepark", "attractie"), "aquarium": ("Aquarium", "attractie"),
            "water_park": ("Waterpark", "attractie"),
        },
        "keys": {"tourism": ["zoo", "theme_park", "aquarium"], "leisure": ["water_park"]},
    },
    "Erfgoed (musea, kastelen, landgoederen)": {
        "typemap": {
            "museum": ("Museum", "erfgoed"), "gallery": ("Galerie", "erfgoed"),
            "castle": ("Kasteel", "erfgoed"), "manor": ("Landgoed", "erfgoed"),
            "fort": ("Fort", "erfgoed"), "monastery": ("Klooster", "erfgoed"),
        },
        "keys": {"tourism": ["museum", "gallery"],
                 "historic": ["castle", "manor", "fort", "monastery"]},
    },
    "Horeca (restaurants, cafés)": {
        "typemap": {
            "restaurant": ("Restaurant", "horeca"), "cafe": ("Café/lunchroom", "horeca"),
            "pub": ("Eetcafé/pub", "horeca"), "bistro": ("Bistro", "horeca"),
            "ice_cream": ("IJssalon", "horeca"),
        },
        "keys": {"amenity": ["restaurant", "cafe", "pub", "bistro", "ice_cream"]},
    },
    "Winkels met een verhaal (cadeau, kunst, sieraden)": {
        "typemap": {
            "gift": ("Cadeauwinkel", "winkel"), "art": ("Kunstwinkel/galerie", "winkel"),
            "books": ("Boekwinkel", "winkel"), "jewelry": ("Sieradenwinkel", "winkel"),
            "antiques": ("Antiek", "winkel"), "interior_decoration": ("Interieurwinkel", "winkel"),
            "pottery": ("Keramiek/aardewerk", "winkel"), "craft": ("Hobby/ambacht", "winkel"),
            "musical_instrument": ("Muziekwinkel", "winkel"), "watches": ("Horlogewinkel", "winkel"),
        },
        "keys": {"shop": ["gift", "art", "books", "jewelry", "antiques",
                          "interior_decoration", "pottery", "craft",
                          "musical_instrument", "watches"]},
    },
    "Bloemisten": {
        "typemap": {"florist": ("Bloemist", "bloemist"), "garden_centre": ("Tuincentrum", "bloemist")},
        "keys": {"shop": ["florist", "garden_centre"]},
    },
}

# Omgekeerde koppeling: opgeslagen type-label (bijv. "Camping") terug naar segment ("natuurcamping").
LABEL_TO_SEG = {}
for _p in PRESETS.values():
    for _osmval, (_label, _seg) in _p["typemap"].items():
        LABEL_TO_SEG[_label] = _seg
if CP_OK:
    LABEL_TO_SEG.update(CP.label_naar_segment())

LINKS = ("\u2022 Website: https://www.voicestamp.nl\n"
         "\u2022 Hoe het werkt: https://www.voicestamp.nl/how-it-works\n"
         "\u2022 Instagram: https://www.instagram.com/voicestamp.nl\n"
         "\u2022 LinkedIn: https://www.linkedin.com/company/voicestamp")
SIG = ("Met vriendelijke groet,\nYusuf Tatlicioglu\n0646756497\n\n"
       "Founder VoiceStamp\nhttps://www.voicestamp.nl/")
APP_LINE = ("Ik zag dat jullie al met een app werken, dus voor de duidelijkheid: VoiceStamp is "
            "g\u00e9\u00e9n tweede app. Een gast hoeft niets te downloaden en geen account te maken. "
            "Ik zie het naast jullie app staan, niet in de weg.")

def _mid_placeholder():
    pass

MID = {
 "natuurcamping": ("Wat als een gast dat verhaal niet alleen leest, maar het ook rechtstreeks van "
   "jullie hoort?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op een plek, en "
   "opent direct een persoonlijke audioboodschap met een landingspagina. Geen app, geen account: "
   "juist minder drukte, niet meer. Een warm welkom, een verhaal over de omgeving, of een wandeltip."),
 "bnb": ("Wat als een gast dat verhaal ook van jullie hoort, op het moment dat hij binnenkomt?\n\n"
   "Met VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op de kamer, en opent direct een "
   "persoonlijke audioboodschap met een landingspagina. Geen app, geen account. Een warm welkom in "
   "je eigen stem, het verhaal van het pand, of een tip voor de omgeving."),
 "hotel": ("Wat als een gast dat verhaal hoort op het moment dat hij binnenkomt?\n\nMet VoiceStamp "
   "scan je een eenvoudige stempel (VoiceStamp) op de kamer of in de lobby, en opent direct een "
   "audioboodschap met een landingspagina. Geen app, geen account. Een warm welkom, het verhaal van "
   "de plek, of info die nu aan de balie wordt gevraagd."),
 "streek": ("Want dat verhaal vertel je in de winkel, maar het gaat niet mee met het product dat "
   "iemand koopt.\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op de verpakking, en "
   "opent direct een audioboodschap met een landingspagina. Geen app, geen account. Waarom een "
   "ingredi\u00ebnt is gekozen, hoe iets gemaakt wordt, of een welkom bij een proeverij."),
 "attractie": ("Wat als een bezoeker bij een dier of attractie een verzorger het hoort vertellen, in "
   "plaats van een bordje te lezen?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) bij "
   "een verblijf, en opent direct een audioboodschap met een landingspagina. Geen app \u2014 en het "
   "werkt ook voor kinderen die nog niet lezen. Het verhaal van een dier, een welkom, of het "
   "dagprogramma."),
 "erfgoed": ("Wat als een bezoeker het verhaal hoort in de stem van een gids of conservator?\n\nMet "
   "VoiceStamp scan je een eenvoudige stempel (VoiceStamp) bij een object of monument, en opent "
   "direct een audioboodschap met een landingspagina. Geen app, geen account. Zo maak je het verhaal "
   "toegankelijk, ook buiten de rondleidingen om."),
 "horeca": ("Wat als een gast niet alleen leest wat er op de kaart staat, maar de chef zelf hoort "
   "vertellen waarom een gerecht er staat?\n\nMet VoiceStamp scan je een eenvoudige stempel "
   "(VoiceStamp) op de menukaart of op tafel, en opent direct een audioboodschap met een "
   "landingspagina. Geen app, geen account. Een welkom bij binnenkomst, het verhaal achter een "
   "streekgerecht of leverancier, of een tip voor een volgende keer."),
 "winkel": ("Want dat verhaal vertel je in de winkel, maar het gaat niet mee met het product dat "
   "iemand koopt.\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op het product of de "
   "verpakking, en opent direct een audioboodschap met een landingspagina. Geen app, geen account. "
   "Waarom een stuk gemaakt is, wie de maker is, of hoe je het het beste gebruikt of onderhoudt."),
 "bloemist": ("Wat als een boeket zelf een boodschap kon meegeven, in de stem van de afzender?\n\n"
   "Met VoiceStamp scan je een eenvoudige stempel (VoiceStamp) bij het boeket, in plaats van het "
   "kleine kaartje, en opent direct een persoonlijke audioboodschap met een landingspagina. Geen "
   "app, geen account. Juist rond momenten als Moederdag of Valentijn maakt een stem het verschil."),
}
SUBJECTS = {
 "natuurcamping": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Een stem bij jullie plek, zonder app"],
 "bnb": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Een persoonlijk welkom op de kamer"],
 "hotel": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Minder vragen aan de balie"],
 "streek": ["Wat als jullie product zelf zijn verhaal kon vertellen?", "Het verhaal achter het product"],
 "attractie": ["Wat als jullie dieren hun eigen verhaal konden vertellen?", "Een stem bij elk verblijf"],
 "erfgoed": ["Wat als jullie collectie zelf haar verhaal kon vertellen?", "Een stem bij het monument"],
 "horeca": ["Wat als jullie gerechten hun eigen verhaal konden vertellen?", "Een woord van de chef, op tafel"],
 "winkel": ["Wat als jullie producten zelf hun verhaal konden vertellen?", "Het verhaal achter het product, bij de klant thuis"],
 "bloemist": ["Wat als een boeket zelf iets kon zeggen?", "Een stem bij de bloemen, geen kaartje"],
}
FOLLOWUP_2 = ("Beste {aanhef},\n\nKorte opvolging op mijn vorige bericht. Ik snap dat het druk is, "
              "dus kort: met {merk} horen jullie mensen een stem op de plek zelf, in plaats van een "
              "bordje te lezen. Geen app, geen account.\n\nZal ik het een keer laten zien?\n\n"
              "Met vriendelijke groet,")
FOLLOWUP_3 = ("Beste {aanhef},\n\nLaatste bericht van mijn kant, daarna laat ik het rusten. Mocht je "
              "ooit willen dat jullie verhaal ook echt te h\u00f3ren is op de plek zelf, dan weet je me "
              "te vinden.\n\nMet vriendelijke groet,")

# ---------------------------------------------------------------- HR / onboarding
HR_SUBJECT = "Nieuwe medewerkers sneller op weg, zonder extra app"
HR_KERN = (
    "Wat als een nieuwe medewerker de belangrijkste info niet hoeft te lezen, maar gewoon kan "
    "horen?\n\n"
    "Een welkom van het team op de eerste dag. Uitleg bij een werkplek, precies waar iemand het "
    "nodig heeft. Of een collega die op het juiste moment zijn kennis deelt.\n\n"
    "Met VoiceStamp scant iemand een eenvoudige stempel en hoort direct de boodschap. Geen app, "
    "geen account, geen inlog.\n\n"
    "Geen app. Geen account. Gewoon luisteren.\n\n"
    "En dat is nog maar het begin. De toepassingen zijn eindeloos: overal waar kennis of een "
    "boodschap op het juiste moment moet landen, kan een VoiceStamp staan. In jullie eigen stem, "
    "of die van een collega.\n\n"
    "Juist bij {invalshoek_zin} scheelt dat tijd en herhaling: de uitleg staat klaar, telkens "
    "hetzelfde, ook als het druk is.")

def maak_hr_mail(naam, functie, invalshoek, brand, links):
    functie = (functie or "").strip()
    aanhef = functie if functie else "team"
    inval = (invalshoek or "").strip().rstrip(".")
    inval_zin = inval[0].lower() + inval[1:] if inval else "het inwerken van nieuwe mensen"
    kern = HR_KERN.format(invalshoek_zin=inval_zin).replace("VoiceStamp", brand["merk"])
    linksblok = "Benieuwd hoe dat eruitziet? Neem gerust een kijkje:\n" + links
    body = (f"Beste {aanhef},\n\n{kern}\n\n{linksblok}\n\n"
            f"Ik laat het graag zien als je denkt dat {brand['merk']} bij {naam} past.\n\n"
            "Met vriendelijke groet,")
    return HR_SUBJECT, body


# Prikkelende onderwerpregels per segment (in plaats van een 'wat als'-vraag).
SUBJECTS_STD = {
    "natuurcamping": "Wat als gasten jullie camping ook konden horen?",
    "bnb": "Wat als gasten jullie plek ook konden horen?",
    "hotel": "Wat als gasten jullie hotel ook konden horen?",
    "streek": "Wat als jullie producten iets konden vertellen?",
    "winkel": "Wat als jullie producten iets konden vertellen?",
    "horeca": "Wat als jullie gerechten iets konden vertellen?",
    "bloemist": "Wat als een boeket iets kon zeggen?",
    "attractie": "Wat als jullie dieren iets konden vertellen?",
    "erfgoed": "Wat als jullie plek ook te horen was?",
}

# Kerntekst per segment (zonder aanhef, zonder links, zonder groet).
STD = {
 "natuurcamping": {"segwoord": "camping", "kern":
   "Wat als gasten jullie camping niet alleen zien, maar ook kunnen horen?\n\n"
   "Een welkom bij aankomst. Het verhaal achter de plek. Een persoonlijke tip voor de mooiste "
   "wandeling.\n\n"
   "Met VoiceStamp scant een gast een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan staan op de kampeerplek, bij de receptie of aan het begin van een route. In "
   "jullie eigen stem.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal op de camping een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je gasten iets mee dat op een bordje vaak verloren gaat, zonder extra drukte voor je "
   "team."},
 "bnb": {"segwoord": "B&B", "kern":
   "Wat als gasten jullie plek niet alleen zien, maar ook kunnen horen?\n\n"
   "Een welkom bij aankomst. Het verhaal van het pand. Een persoonlijke tip voor de omgeving.\n\n"
   "Met VoiceStamp scant een gast een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan op de kamer, in de gezamenlijke ruimte of bij een bijzonder detail van het "
   "pand. In jullie eigen stem.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal in jullie B&B een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je gasten iets mee dat op een briefje vaak verloren gaat, zonder extra drukte voor je "
   "team."},
 "hotel": {"segwoord": "hotel", "kern":
   "Wat als gasten jullie hotel niet alleen zien, maar ook kunnen horen?\n\n"
   "Een welkom bij aankomst. Het verhaal van het pand. Een tip voor de omgeving.\n\n"
   "Met VoiceStamp scant een gast een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan op de kamer, in de lobby of in het restaurant. In jullie eigen stem.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal in jullie hotel een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je gasten iets mee dat op een gelamineerd vel vaak verloren gaat, en vang je meteen "
   "vragen af die anders bij de receptie komen."},
 "streek": {"segwoord": "zaak", "kern":
   "Wat als jullie producten niet alleen gezien worden, maar ook iets kunnen vertellen?\n\n"
   "Waarom een druif, een kaas of een ingredi\u00ebnt gekozen is. Hoe iets gemaakt wordt. Een woord "
   "van de maker zelf.\n\n"
   "Met VoiceStamp scant een klant een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan op de verpakking, bij een schap of aan de proeftafel. In jullie eigen stem.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan bij elk product een verhaal, boodschap of moment "
   "een stem geven.\n\n"
   "Zo geef je klanten iets mee dat op een etiket vaak verloren gaat, ook nog thuis aan tafel."},
 "winkel": {"segwoord": "winkel", "kern":
   "Wat als jullie producten niet alleen gezien worden, maar ook iets kunnen vertellen?\n\n"
   "Wie het maakte. Waarom het bijzonder is. Hoe je het het beste gebruikt.\n\n"
   "Met VoiceStamp scant een klant een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan op het product, op de verpakking of bij een schap. In de stem van de maker.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan bij elk product een verhaal, boodschap of moment "
   "een stem geven.\n\n"
   "Zo geef je klanten iets mee dat in de winkel verteld wordt, maar anders niet meegaat naar huis."},
 "horeca": {"segwoord": "zaak", "kern":
   "Wat als gasten jullie gerechten niet alleen proeven, maar er ook iets over kunnen horen?\n\n"
   "Waarom een gerecht op de kaart staat. Waar een product vandaan komt. Een woord van de chef.\n\n"
   "Met VoiceStamp scant een gast een eenvoudige stempel en hoort direct jullie boodschap, op het "
   "moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan op de menukaart, op tafel of bij binnenkomst. In jullie eigen stem.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal in jullie zaak een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je gasten iets mee dat op een drukke avond vaak verloren gaat, zonder extra werk voor de "
   "bediening."},
 "bloemist": {"segwoord": "winkel", "kern":
   "Wat als een boeket niet alleen gezien wordt, maar ook iets kan zeggen?\n\n"
   "Een persoonlijke felicitatie. Een woord van troost. Gewoon: ik denk aan je.\n\n"
   "Met VoiceStamp scant de ontvanger een eenvoudige stempel en hoort direct de boodschap, op het "
   "moment dat het telt.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan bij het boeket, in plaats van het kleine kaartje. In de eigen stem van de "
   "afzender.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan bij elk boeket een boodschap of moment een stem "
   "geven.\n\n"
   "Juist rond Moederdag en Valentijn maakt een stem het verschil, zonder extra gedoe voor jullie."},
 "attractie": {"segwoord": "park", "kern":
   "Wat als bezoekers jullie dieren en attracties niet alleen zien, maar er ook iets over kunnen "
   "horen?\n\n"
   "Het verhaal van een dier. Een welkom bij binnenkomst. Het programma van die dag.\n\n"
   "Met VoiceStamp scant een bezoeker een eenvoudige stempel en hoort direct jullie boodschap, op "
   "het moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan bij een verblijf, langs de route of bij de ingang. Verteld door een "
   "verzorger.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal in jullie park een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je bezoekers iets mee dat op een bordje vaak verloren gaat, en dat ook werkt voor "
   "kinderen die nog niet lezen."},
 "erfgoed": {"segwoord": "organisatie", "kern":
   "Wat als bezoekers jullie plek niet alleen zien, maar ook kunnen horen?\n\n"
   "Het verhaal achter een stuk. Een route langs de hoogtepunten. Een welkom bij binnenkomst.\n\n"
   "Met VoiceStamp scant een bezoeker een eenvoudige stempel en hoort direct jullie boodschap, op "
   "het moment en de plek waar die relevant is.\n\n"
   "Geen app. Geen account. Gewoon luisteren.\n\n"
   "Een VoiceStamp kan bij een object, een monument of langs een route. In de stem van een gids of "
   "conservator.\n\n"
   "En dat is maar \u00e9\u00e9n voorbeeld. VoiceStamp kan overal in jullie collectie een verhaal, boodschap "
   "of moment een stem geven.\n\n"
   "Zo maak je jullie verhaal toegankelijk, ook buiten de rondleidingen om."},
}

DEFAULT_OPENER = {
 "natuurcamping": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan rust en natuur.",
 "bnb": "Tijdens het bekijken van jullie website viel me op hoe persoonlijk jullie gasten ontvangen.",
 "hotel": "Tijdens het bekijken van jullie website viel me op hoeveel karakter jullie plek heeft.",
 "streek": "Tijdens het bekijken van jullie website viel me op hoeveel verhaal er in jullie producten zit.",
 "attractie": "Tijdens het bekijken van jullie website viel me op hoeveel er bij jullie te beleven is.",
 "erfgoed": "Tijdens het bekijken van jullie website viel me op hoeveel verhaal jullie plek draagt.",
 "horeca": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan de gerechten en de sfeer.",
 "winkel": "Tijdens het bekijken van jullie website viel me op hoeveel verhaal er in jullie producten zit.",
 "bloemist": "Tijdens het bekijken van jullie website viel me op hoeveel gevoel jullie in jullie boeketten leggen.",
}
if CP_OK:
    _OPENERS_NIEUW = {
        "zorg": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan een goede uitleg voor patiënten.",
        "wellness": "Tijdens het bekijken van jullie website viel me op hoeveel rust en aandacht jullie plek uitstraalt.",
        "vastgoed": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan wonen en de mensen die er wonen.",
        "werkgever": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan jullie mensen.",
        "evenement": "Tijdens het bekijken van jullie website viel me op hoeveel er bij jullie te beleven valt.",
        "entertainment": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan het programma.",
        "onderwijs": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan studenten en leerlingen.",
        "dealer": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan advies en service.",
        "vervoer": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan de reiziger.",
        "recreatie": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan de omgeving en de bezoeker.",
    }
    for _seg, _d in CP.SEGMENTEN.items():
        STD[_seg] = {"segwoord": _d["segwoord"], "kern": _d["kern"]}
        SUBJECTS_STD[_seg] = _d["onderwerp"]
        DEFAULT_OPENER[_seg] = _OPENERS_NIEUW[_seg]

KEYWORDS = ("rust", "natuur", "gastvrij", "persoonlijk", "familie", "verhaal", "duurzaam",
            "welkom", "beleef", "genieten", "monument", "ambacht", "streek", "traditie",
            "avontuur", "kinderen", "dieren")

# --- Haak-vinder: signalen die een reden geven om contact op te nemen ---
# label -> regex (op kleine letters). De eerste match per lead wordt de "haak".
SIGNALS = {
    "Prijs/onderscheiding": re.compile(
        r"\b(award|bekroond|genomineerd|winnaar|onderscheiding|pincamp|green key|"
        r"michelin|beste [a-z]+ van|verkozen tot)\b"),
    "Lange historie": re.compile(
        r"\b(sinds (1[6-9]\d\d|20[01]\d)|opgericht in|al \d{2,3} jaar|"
        r"\d+e generatie|generatie op generatie|eeuwenoud)\b"),
    "Familiebedrijf": re.compile(r"\b(familiebedrijf|van vader op zoon|onze familie|familie [a-z]+ runt)\b"),
    "Duurzaam / B Corp": re.compile(r"\b(duurzaam|biologisch|b corp|co2-neutraal|klimaatneutraal|permacultuur)\b"),
    "Net geopend / nieuw": re.compile(r"\b(net geopend|sinds kort|nieuw seizoen|onlangs geopend|kersvers)\b"),
    "Eigen verhaal/oprichter": re.compile(r"\b(ons verhaal|het verhaal achter|de oprichter|opgericht door)\b"),
    "Ambachtelijk / eigen productie": re.compile(r"\b(ambachtelijk|zelfgemaakt|eigen productie|met de hand gemaakt|hoeve-eigen)\b"),
}
# Als een van deze matcht heeft het bedrijf al een app -> pas de pitch aan.
APP_HINTS = re.compile(r"(download (?:onze|de) app|in de app store|google play|via onze app|onze eigen app)")
SIGNAL_WORDS = ("award", "bekroond", "genomineerd", "winnaar", "michelin", "green key",
                "familiebedrijf", "generatie", "sinds 1", "sinds 20", "ambachtelijk",
                "duurzaam", "oprichter", "ons verhaal", "b corp")

def detect_signals(text):
    """Geef (haak-label of '', heeft_app-bool) terug op basis van de sitetekst."""
    tl = text.lower()
    haak = ""
    for label, rx in SIGNALS.items():
        if rx.search(tl):
            haak = label
            break
    return haak, bool(APP_HINTS.search(tl))

# ------------------------------------------------------------------ OSM
def q_build(area, keys, whole):
    parts = []
    for k, vals in keys.items():
        for v in vals:
            parts.append(f'node["{k}"="{v}"](area.a);way["{k}"="{v}"](area.a);'
                         f'relation["{k}"="{v}"](area.a);')
    ab = ('area["ISO3166-1"="NL"][admin_level=2]->.a;' if whole
          else f'area["name"="{area}"]["admin_level"="4"]->.a;')
    return f'[out:json][timeout:180];{ab}({"".join(parts)});out center tags;'

def _overpass_post(q):
    laatste = None
    # 2 rondes over alle servers, met een korte pauze, zodat tijdelijke 500/502-fouten worden opgevangen
    for ronde in range(2):
        for ep in OVERPASS:
            try:
                r = requests.post(ep, data={"data": q},
                                  headers={"User-Agent": UA}, timeout=180)
                if r.status_code in (429, 500, 502, 503, 504):
                    laatste = f"{r.status_code} van {urlparse(ep).netloc}"
                    time.sleep(1.5)
                    continue
                r.raise_for_status()
                return r.json().get("elements", [])
            except Exception as e:
                laatste = f"{type(e).__name__} bij {urlparse(ep).netloc}"
                time.sleep(1.0)
                continue
        time.sleep(2)
    raise RuntimeError(
        "De OpenStreetMap-servers reageren nu niet (dit ligt aan hun kant, niet aan de app). "
        f"Laatste melding: {laatste}. Probeer het over een paar minuten opnieuw, of kies "
        "eventueel een kleinere provincie.")

@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def osm_cached(area, keys_items, whole):
    keys = {k: list(v) for k, v in keys_items}
    return _overpass_post(q_build(area, keys, whole))

@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def osm_cached_cl(area, clauses, whole):
    """Zoekopdracht voor de campagnes met subcategorieën (campagnes.py)."""
    ab = ('area["ISO3166-1"="NL"][admin_level=2]->.a;' if whole
          else f'area["name"="{area}"]["admin_level"="4"]->.a;')
    q = f'[out:json][timeout:180];{ab}({CP.overpass_stukken(clauses)});out center tags;'
    return _overpass_post(q)

def _lead_uit_tags(t, label, seg):
    email = (t.get("contact:email") or t.get("email") or "").strip().lower()
    return {"naam": t.get("name"), "type": label, "seg": seg,
            "plaats": (t.get("addr:city") or t.get("addr:place") or "").strip(),
            "email": email if EMAIL_RE.fullmatch(email or "") else "",
            "website": (t.get("contact:website") or t.get("website") or "").strip(),
            "telefoon": (t.get("contact:phone") or t.get("phone") or "").strip(),
            "instagram": (t.get("contact:instagram") or ""),
            "linkedin": "", "facebook": (t.get("contact:facebook") or ""),
            "haak": "", "heeft_app": False,
            "bron": "OpenStreetMap", "opener": ""}

def parse(els, typemap):
    out = []
    for el in els:
        t = el.get("tags", {}); name = t.get("name")
        if not name: continue
        osmval = next((t.get(k) for k in ("tourism", "craft", "shop", "historic", "leisure", "amenity")
                       if t.get(k) in typemap), None)
        if not osmval: continue
        label, seg = typemap[osmval]
        out.append(_lead_uit_tags(t, label, seg))
    return out

def parse_subs(els, subs):
    """Deelt elke gevonden plek in bij de eerste subcategorie waar hij bij past."""
    out = []
    for el in els:
        t = el.get("tags", {})
        if not t.get("name"): continue
        sub = CP.deel_in(t, subs)
        if sub: out.append(_lead_uit_tags(t, sub["label"], sub["seg"]))
    return out

# ------------------------------------------------------------------ website
def norm(u):
    if not u: return ""
    if not u.startswith(("http://", "https://")): u = "https://" + u
    return u.rstrip("/")

def detag(h):
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
    return re.sub(r"<[^>]+>", " ", h)

# Vaste, altijd-correcte openingszinnen per gevonden haak (geen zinsontleding nodig).
HOOK_OPENERS = {
    "Prijs/onderscheiding": "Tijdens het bekijken van jullie website viel me op dat jullie in de prijzen zijn gevallen.",
    "Lange historie": "Tijdens het bekijken van jullie website viel me op hoeveel geschiedenis er in jullie plek zit.",
    "Familiebedrijf": "Tijdens het bekijken van jullie website viel me op dat het een echt familiebedrijf is.",
    "Duurzaam / B Corp": "Tijdens het bekijken van jullie website viel me op hoe bewust jullie met duurzaamheid omgaan.",
    "Net geopend / nieuw": "Tijdens het bekijken van jullie website viel me op dat jullie nog niet zo lang geleden zijn begonnen.",
    "Eigen verhaal/oprichter": "Tijdens het bekijken van jullie website viel me op hoeveel persoonlijk verhaal er achter jullie plek zit.",
    "Ambachtelijk / eigen productie": "Tijdens het bekijken van jullie website viel me op hoe ambachtelijk jullie te werk gaan.",
}

def _schone_citaatzin(text):
    """Zoek een korte, nette zin op de site die we LETTERLIJK kunnen citeren (altijd grammaticaal)."""
    text = re.sub(r"\s+", " ", text).strip()
    best, score = None, 0
    for z in re.split(r"(?<=[.!?]) ", text):
        z = z.strip()
        zl = z.lower()
        # streng: begint met hoofdletter, nette lengte, geen rare tekens of webrommel
        if not (35 <= len(z) <= 140): continue
        if not z[0].isupper(): continue
        if any(x in zl for x in ("cookie", "javascript", "©", "http", "@", "lees meer", "klik")): continue
        if any(ch in z for ch in "{}<>|[]"): continue
        s = sum(k in zl for k in KEYWORDS)
        if s > score: best, score = z.rstrip(".!?"), s
    return best if score >= 1 else None

def maak_opener(sigtext, seg, plaats, haak):
    """Altijd grammaticaal correcte openingszin: haak > letterlijk citaat > standaard."""
    if haak and haak in HOOK_OPENERS:
        return HOOK_OPENERS[haak]
    zin = _schone_citaatzin(sigtext or "")
    if zin:
        return f'Tijdens het bekijken van jullie website bleef deze zin me bij: "{zin}."'
    if plaats:
        return DEFAULT_OPENER[seg].rstrip(".") + f", hier in {plaats}."
    return DEFAULT_OPENER[seg]

def enrich_one(l):
    """Verrijk een enkele lead: e-mail, openingszin en social links. Muteert en geeft terug."""
    base = norm(l["website"])
    if not base:
        l["opener"] = l["opener"] or DEFAULT_OPENER[l["seg"]]
        return l
    dom = urlparse(base).netloc.replace("www.", "")
    email = l["email"]
    sigtext = ""  # verzamelde tekst van home + over-ons voor de haak-vinder en de openingszin
    for p in PATHS:
        url = base if p == "" else f"{base}/{p}"
        if not _robots_ok(base, p):
            continue
        try:
            r = polite_get(url, dom, timeout=8)
            if not r or r.status_code != 200 or "html" not in r.headers.get("content-type", ""): continue
            html = r.text
            if not email:
                c = re.findall(r"mailto:([^\"'?\s>]+)", html) + EMAIL_RE.findall(html)
                c = [x.strip().strip(".").lower() for x in c]
                c = [x for x in c if EMAIL_RE.fullmatch(x) and not any(b in x for b in BAD)]
                same = [x for x in c if x.split("@")[-1] == dom]; pool = same or c
                for pref in ("info@", "contact@", "welkom@", "boeking@"):
                    hit = next((x for x in pool if x.startswith(pref)), "")
                    if hit: email = hit; break
                if not email and pool: email = pool[0]
            for netwerk, rx in SOCIAL_RE.items():
                if not l.get(netwerk):
                    m = rx.search(html)
                    if m: l[netwerk] = m.group(0)
            if p in ("", "over-ons", "about"):
                sigtext += " " + detag(html)
            if email and sigtext: break
        except Exception:
            continue
    haak, heeft_app = detect_signals(sigtext)
    l["haak"] = haak
    l["heeft_app"] = heeft_app
    l["email"] = email
    if email and l["bron"] == "OpenStreetMap": l["bron"] = "OSM + website"
    l["opener"] = maak_opener(sigtext, l["seg"], l["plaats"], haak)
    return l

def has_mx(e):
    if not HAVE_DNS or "@" not in e: return None
    try: return len(dns.resolver.resolve(e.split("@")[-1], "MX")) > 0
    except Exception: return False

def voornaam(email):
    if not email: return ""
    lok = email.split("@")[0]
    if lok in ROLE or not re.fullmatch(r"[a-zA-Z]{3,14}", lok): return ""
    return lok.capitalize()

def ingang(l):
    e = l["email"]
    if not e: return "Adres ontbreekt"
    lok = e.split("@")[0]
    pers = lok not in ROLE and not any(lok.startswith(r) for r in ROLE)
    keten = any(k in (l["website"] + l["naam"]).lower() for k in KETEN)
    return f"{'Persoonlijk' if pers else 'Algemene inbox'} \u2022 {'keten (traag)' if keten else 'klein (snel)'}"

def maak_mail(l, brand):
    seg = l["seg"] if l["seg"] in STD else "natuurcamping"
    subject = SUBJECTS_STD.get(seg, SUBJECTS_STD["natuurcamping"])
    if l.get("heeft_app"):
        subject = "Geen tweede app \u2014 juist minder gedoe"
    # Aanhef: 'algemeen' = "Beste team," ; anders "Beste team van <naam>," (of voornaam indien bekend)
    if brand.get("aanhef_modus") == "algemeen":
        aanhef = "team"
    else:
        aanhef = l.get("_vn") or f"team van {l['naam']}"
    # Volledige-mail-override: vervangt de hele mail (alle segmenten).
    if brand.get("eigen_body", "").strip():
        ctx = {"aanhef": aanhef, "naam": l["naam"], "plaats": l["plaats"] or "jullie omgeving",
               "haak": l.get("haak", ""), "merk": brand["merk"], "links": brand["links"]}
        try:
            body = brand["eigen_body"].format(**ctx)
        except Exception:
            body = brand["eigen_body"]
        fu2 = FOLLOWUP_2.format(aanhef=aanhef, merk=brand["merk"])
        fu3 = FOLLOWUP_3.format(aanhef=aanhef)
        return subject, body, fu2, fu3
    if brand["eigen_mid"].strip():
        try:
            kern = brand["eigen_mid"].format(naam=l["naam"], plaats=l["plaats"] or "jullie omgeving",
                                             haak=l.get("haak", ""))
        except Exception:
            kern = brand["eigen_mid"]
    else:
        kern = STD[seg]["kern"]
    kern = kern.replace("VoiceStamp", brand["merk"])
    if l.get("heeft_app"):
        kern += "\n\n" + APP_LINE.replace("VoiceStamp", brand["merk"])
    linksblok = "Benieuwd hoe dat eruitziet? Neem gerust een kijkje:\n" + brand["links"]
    body = (f"Beste {aanhef},\n\n{kern}\n\n{linksblok}\n\n"
            f"Ik laat het graag zien als je denkt dat {brand['merk']} bij jullie past.\n\n"
            "Met vriendelijke groet,")
    fu2 = FOLLOWUP_2.format(aanhef=aanhef, merk=brand["merk"])
    fu3 = FOLLOWUP_3.format(aanhef=aanhef)
    return subject, body, fu2, fu3

# ------------------------------------------------ Sellbird-onderdelen: beslisser, reeks, antwoorden
def _fetch_voor_sp(base, path):
    """Pagina ophalen met robots.txt en pauze per domein (zelfde regels als de rest van de app)."""
    dom = urlparse(base).netloc.replace("www.", "")
    if not _robots_ok(base, path):
        return None
    url = base if path == "" else f"{base}/{path}"
    r = polite_get(url, dom, timeout=8)
    if r is not None and r.status_code == 200 and "html" in r.headers.get("content-type", ""):
        return r.text
    return None

def _verifier_keys():
    try:
        return st.secrets.get("REOON_API_KEY"), st.secrets.get("ZEROBOUNCE_API_KEY")
    except Exception:
        return None, None

def zoek_beslisser(row, rk=None, zk=None):
    """
    Zoek eigenaar of manager op de eigen site van het bedrijf en bouw het bijpassende adres.
    Geeft (velden_voor_crm of None, uitleg).
    """
    base = norm(row.get("website") or "")
    if not base:
        return None, "geen website bij deze lead"
    res = SP.find_decision_makers(base, row.get("company_name") or "", _fetch_voor_sp)
    upd = {}
    hoofd = (row.get("email") or "").lower()
    if hoofd and hoofd in res["emails"]:
        upd["email_status"] = SP.S_PUBLISHED
    if not res["people"]:
        return (upd or None), f"geen eigenaar of manager gevonden ({res['pages']} pagina's gelezen)"
    naam, rol = res["people"][0]
    cands = SP.email_candidates(naam, res["domain"], res["emails"])
    addr, status, why = SP.pick_contact_email(cands, set(res["emails"]), rk, zk)
    upd.update({"contact_name": naam, "contact_role": rol,
                "contact_email": addr, "contact_email_status": status or ""})
    return upd, f"{naam} ({rol}): {addr or 'geen adres'} [{SP.STATUS_LABEL.get(status, status)}] - {why}"

def kies_ontvanger(row, allow_risky=False):
    """Persoonlijk adres als dat gepubliceerd of geverifieerd is, anders de gewone inbox."""
    ok = {SP.S_PUBLISHED, SP.S_VERIFIED} | ({SP.S_RISKY} if allow_risky else set())
    ce = (row.get("contact_email") or "").strip()
    cs = row.get("contact_email_status") or ""
    if ce and cs in ok and row.get("contact_name"):
        return ce, cs, row["contact_name"].split()[0]
    return (row.get("email") or "").strip(), (row.get("email_status") or None), ""

def status_voor(lead, to):
    to = (to or "").lower()
    if to and to == (lead.get("contact_email") or "").lower():
        return lead.get("contact_email_status") or None
    return lead.get("email_status") or None

def auto_dagbatch(cl, uid, rows, n, brand, rk, zk, start, bestaand_ids, voortgang=None):
    """
    Kiest zelf de beste leads: status Nieuw, nog geen reeks, hoogste score eerst. Zoekt voor de beste
    kandidaten de eigenaar of manager, bouwt de reeksen en zet ze klaar voor goedkeuring.
    Geeft (aantal_reeksen, nieuwe_outbox_rijen, logregels).
    """
    log = []
    kandidaten = [r for r in rows
                  if (r.get("status") or "Nieuw") == "Nieuw" and r["id"] not in bestaand_ids
                  and ((r.get("email") or "").strip() or (r.get("website") or "").strip())]
    kandidaten.sort(key=lambda r: -(r.get("score") or 0))
    pool = kandidaten[: max(n * 2, n + 5)]            # iets meer dan nodig: niet elke lead levert een reeks op
    te_zoeken = [r for r in pool if not r.get("contact_name") and (r.get("website") or "").strip() and SP_OK]
    if te_zoeken:
        with ThreadPoolExecutor(max_workers=4) as ex:
            futs = {ex.submit(zoek_beslisser, r, rk, zk): r for r in te_zoeken}
            for k, f in enumerate(as_completed(futs), 1):
                r = futs[f]
                try:
                    upd, uitleg = f.result()
                except Exception:
                    upd, uitleg = None, "zoeken mislukt"
                if upd:
                    r.update(upd)
                    try:
                        crm_update(cl, r["id"], uid, upd)
                        if upd.get("contact_name"):
                            crm_add_activity(cl, uid, r["id"], "Contactpersoon gevonden", uitleg)
                    except Exception:
                        pass
                if voortgang: voortgang(k / len(futs))
    nieuw_rows, gemaakt = [], 0
    for r in pool:
        if gemaakt >= n:
            break
        items, why = maak_reeks(r, brand, start, False, None)
        if items is None:
            log.append(f"{r.get('company_name')}: overgeslagen ({why})")
            continue
        nieuw_rows.extend(items); gemaakt += 1
    return gemaakt, nieuw_rows, log

def volledige_tekst(body, sig, optout=""):
    t = body.rstrip() + "\n" + (sig or "")
    if (optout or "").strip():
        t = t.rstrip() + "\n\n" + optout.strip()
    return t

def maak_reeks(row, brand, start, allow_risky=False, onbekend_seg=None):
    """Mail 1, 2 en 3 voor een lead, met geplande datums. Geeft (items, reden_als_het_niet_lukt)."""
    to, st_, vn = kies_ontvanger(row, allow_risky)
    if not to:
        return None, "geen e-mailadres"
    seg = LABEL_TO_SEG.get(str(row.get("segment") or ""), onbekend_seg)
    if not seg:
        return None, f"type '{row.get('segment') or '-'}' niet herkend"
    lead_obj = {"seg": seg, "naam": row.get("company_name") or "", "plaats": "", "_vn": vn,
                "haak": row.get("hook") or "", "heeft_app": False, "email": to}
    subj, m1, m2, m3 = maak_mail(lead_obj, brand)
    data_ = SP.plan_dates(start)
    return [{"lead_id": row["id"], "step": 1, "to_email": to, "subject": subj, "body": m1,
             "status": "pending_review", "scheduled_for": data_[0].isoformat()},
            {"lead_id": row["id"], "step": 2, "to_email": to, "subject": "Re: " + subj, "body": m2,
             "status": "pending_review", "scheduled_for": data_[1].isoformat()},
            {"lead_id": row["id"], "step": 3, "to_email": to, "subject": "Re: " + subj, "body": m3,
             "status": "pending_review", "scheduled_for": data_[2].isoformat()}], ""

def verwerk_antwoord(cl, uid, lead, soort, nieuwe_status, onderwerp, tekst, mid=""):
    """Zet het antwoord door naar het CRM en stopt de rest van de reeks. Geeft de nieuwe status."""
    from datetime import date as _d
    lid, huidig = lead["id"], (lead.get("status") or "Nieuw")
    velden = {}
    if nieuwe_status == "Do not contact":
        velden = {"status": "Do not contact", "dnc_reason": "Afgemeld per mail",
                  "next_action": "", "next_action_date": None}
    elif nieuwe_status == "Afgewezen" and huidig != "Klant":
        velden = {"status": "Afgewezen", "next_action": "", "next_action_date": None}
    elif nieuwe_status == "Reactie" and huidig not in ("Klant", "Afspraak"):
        velden = {"status": "Reactie", "next_action": "Antwoord geven",
                  "next_action_date": _d.today().isoformat()}
    if velden:
        crm_update(cl, lid, uid, velden)
    ob_cancel_for_lead(cl, uid, lid, f"gestopt: {soort}")
    snip = re.sub(r"\s+", " ", SP.strip_quoted(tekst))[:200]
    pre = f"[{mid}] " if mid else ""
    crm_add_activity(cl, uid, lid, "Antwoord ontvangen", f"{pre}{soort}: {onderwerp} — {snip}")
    return velden.get("status", huidig)

def verwerk_bounce(cl, uid, lead, adres, mid=""):
    if adres and adres.lower() == (lead.get("contact_email") or "").lower():
        crm_update(cl, lead["id"], uid, {"contact_email_status": SP.S_INVALID})
    else:
        crm_update(cl, lead["id"], uid, {"email_status": SP.S_INVALID})
    ob_cancel_for_lead(cl, uid, lead["id"], "gestopt: bounce")
    pre = f"[{mid}] " if mid else ""
    crm_add_activity(cl, uid, lead["id"], "Bounce", f"{pre}adres onbereikbaar: {adres}")

def verwerk_inbox(cl, uid, mails, leads, eigen):
    """Loopt de inbox door: bounce, afmelding, afwijzing of reactie, en past het CRM aan."""
    try:
        r = (cl.table("activities").select("description").eq("user_id", uid)
             .in_("type", ["Antwoord ontvangen", "Bounce"])
             .order("created_at", desc=True).limit(800).execute())
        bekend = " ".join(x.get("description") or "" for x in (r.data or []))
    except Exception:
        bekend = ""
    uit = []
    for m in mails:
        if m["from"] == (eigen or "").lower():
            continue
        mid = m["message_id"]
        if mid and f"[{mid}]" in bekend:
            continue
        soort, st_, uitleg = SP.classify_reply(m["subject"], m["body"], m["from"], m["auto"], eigen)
        if soort == "auto":
            continue
        if soort == "bounce":
            adres = SP.failed_recipient(m["body"], eigen)
            lead = SP.match_lead(adres, leads)
            if not lead:
                continue
            verwerk_bounce(cl, uid, lead, adres, mid)
            uit.append({"Bedrijf": lead.get("company_name"), "Soort": "Bounce", "Gevolg": f"{adres} ongeldig, reeks gestopt",
                        "Onderwerp": m["subject"]})
            continue
        lead = SP.match_lead(m["from"], leads)
        # alleen leads aan wie we gemaild hebben; anders is het geen antwoord maar gewone post
        if not lead or (not lead.get("last_mailed") and (lead.get("status") or "Nieuw") == "Nieuw"):
            continue
        nieuw = verwerk_antwoord(cl, uid, lead, soort, st_, m["subject"], m["body"], mid)
        uit.append({"Bedrijf": lead.get("company_name"), "Soort": uitleg, "Gevolg": f"status: {nieuw}, reeks gestopt",
                    "Onderwerp": m["subject"]})
    return uit

def score_lead(l):
    s = 0
    if l["email"]: s += 40
    if l.get("mx") is True: s += 10
    ing = l.get("_ingang", "")
    if "Persoonlijk" in ing: s += 20
    if "klein" in ing: s += 15
    if l.get("haak"): s += 15
    if l.get("instagram") or l.get("linkedin"): s += 5
    return min(s, 100)

def kanaal_advies(l):
    if not l["email"]:
        if l.get("linkedin") or l.get("instagram"):
            return "Geen e-mail \u2014 benader via social"
        return "Geen e-mail \u2014 bel of zoek contact"
    ing = l.get("_ingang", "")
    if "keten" in ing:
        return "Keten \u2014 zoek de marketing/beslisser via LinkedIn"
    if "Persoonlijk" in ing:
        return "Mail de eigenaar direct (persoonlijk adres)"
    return "Mail (algemene inbox) \u2014 vraag om doorzetten"

def laad_verstuurd(bytes_data, filename):
    s = set()
    if not bytes_data: return s
    try:
        if filename.lower().endswith((".xlsx", ".xlsm")):
            wb = load_workbook(io.BytesIO(bytes_data), read_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for c in row:
                        if isinstance(c, str) and EMAIL_RE.fullmatch(c.strip().lower()):
                            s.add(c.strip().lower())
        else:
            for m in EMAIL_RE.findall(bytes_data.decode("utf-8", "ignore")): s.add(m.strip().lower())
    except Exception: pass
    return s

# ------------------------------------------------------------------ export
def build_xlsx(leads):
    wb = Workbook(); ws = wb.active; ws.title = "Leads"
    H = ["Score", "Naam", "Type", "Plaats", "E-mail", "Voornaam", "Haak (reden voor contact)",
         "Heeft app?", "Beste ingang", "Kanaal-advies", "E-mail geldig?", "Verzendbatch",
         "Website", "Telefoon", "Instagram", "LinkedIn", "Facebook", "Concept onderwerp",
         "Concept mail (1)", "Opvolgmail (2)", "Opvolgmail (3)", "Status", "Datum verstuurd", "Bron"]
    hf = PatternFill("solid", fgColor="2E5D4B"); hfont = Font("Arial", bold=True, color="FFFFFF")
    th = Side(style="thin", color="D0D0D0"); bd = Border(th, th, th, th)
    for c, h in enumerate(H, 1):
        x = ws.cell(1, c, h); x.fill = hf; x.font = hfont
        x.alignment = Alignment(vertical="center", wrap_text=True); x.border = bd
    for r, l in enumerate(leads, 2):
        mx = l.get("mx"); mxt = "ja" if mx is True else ("nee" if mx is False else "onbekend")
        vals = [l.get("_score", ""), l["naam"], l["type"], l["plaats"], l["email"], l.get("_vn", ""),
                l.get("haak", ""), "ja" if l.get("heeft_app") else "", l["_ingang"], l.get("_kanaal", ""),
                mxt, l.get("_batch", ""), l["website"], l["telefoon"], l.get("instagram", ""),
                l.get("linkedin", ""), l.get("facebook", ""), l["_subject"], l["_body"],
                l.get("_fu2", ""), l.get("_fu3", ""), "", "", l["bron"]]
        for c, v in enumerate(vals, 1):
            x = ws.cell(r, c, v); x.font = Font("Arial", size=10)
            x.alignment = Alignment(vertical="top", wrap_text=True); x.border = bd
    for i, w in enumerate([7, 24, 16, 13, 25, 10, 21, 9, 22, 30, 11, 10, 22, 13, 19, 19, 19, 28, 54, 40, 40, 10, 12, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 28
    ws.auto_filter.ref = f"A1:{get_column_letter(len(H))}{len(leads)+1}"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_csv(leads):
    cols = ["score", "email", "first_name", "company", "city", "hook", "has_app", "channel_advice",
            "custom_opener", "custom_subject", "custom_body", "followup_2", "followup_3",
            "send_batch", "website", "phone", "instagram", "linkedin", "facebook", "beste_ingang"]
    buf = io.StringIO(); w = csv.DictWriter(buf, fieldnames=cols); w.writeheader()
    for l in leads:
        if not l["email"]: continue
        w.writerow({"score": l.get("_score", ""), "email": l["email"], "first_name": l.get("_vn", ""),
                    "company": l["naam"], "city": l["plaats"], "hook": l.get("haak", ""),
                    "has_app": "yes" if l.get("heeft_app") else "", "channel_advice": l.get("_kanaal", ""),
                    "custom_opener": l.get("opener", ""), "custom_subject": l["_subject"],
                    "custom_body": l["_body"], "followup_2": l.get("_fu2", ""),
                    "followup_3": l.get("_fu3", ""), "send_batch": l.get("_batch", ""),
                    "website": l["website"], "phone": l["telefoon"], "instagram": l.get("instagram", ""),
                    "linkedin": l.get("linkedin", ""), "facebook": l.get("facebook", ""),
                    "beste_ingang": l["_ingang"]})
    return buf.getvalue().encode("utf-8")

# ------------------------------------------------------------------ pipeline
def draai(preset, area, whole, scrape, mx_check, dedupe_dom, n_per_dag, max_sites,
          workers, verstuurd, brand, progress, status):
    status("OpenStreetMap ophalen ...")
    if preset.get("subs"):
        els = osm_cached_cl(area, CP.clauses(preset["subs"]), whole)
        leads = parse_subs(els, preset["subs"])
    else:
        keys_items = tuple((k, tuple(v)) for k, v in preset["keys"].items())
        els = osm_cached(area, keys_items, whole)
        leads = parse(els, preset["typemap"])
    seen, uniek = set(), []
    for l in leads:
        k = (l["naam"].lower(), l["plaats"].lower())
        if k not in seen: seen.add(k); uniek.append(l)
    leads = uniek
    if verstuurd:
        leads = [l for l in leads if l["email"].lower() not in verstuurd]
    status(f"{len(leads)} plekken gevonden. Websites verrijken ...")
    progress(0.15)

    if scrape:
        todo = [l for l in leads if l["website"]]
        if max_sites: todo = todo[:max_sites]
        done = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(enrich_one, l): l for l in todo}
            for _ in as_completed(futs):
                done += 1
                if done % 5 == 0 or done == len(todo):
                    status(f"Websites verrijken … {done}/{len(todo)}")
                    progress(0.15 + 0.7 * done / max(len(todo), 1))
    progress(0.86)

    if dedupe_dom:
        gz, dd = set(), []
        for l in leads:
            d = (l["email"].split("@")[-1] if l["email"] else l["website"]).lower()
            if d and d in gz: continue
            gz.add(d); dd.append(l)
        leads = dd

    for l in leads:
        l["_vn"] = voornaam(l["email"]); l["_ingang"] = ingang(l)

    if mx_check and HAVE_DNS:
        status("E-mailadressen controleren (MX) ...")
        with ThreadPoolExecutor(max_workers=workers) as ex:
            res = {ex.submit(has_mx, l["email"]): l for l in leads if l["email"]}
            for fut in as_completed(res):
                res[fut]["mx"] = fut.result()
        for l in leads:
            if "mx" not in l: l["mx"] = None
    else:
        for l in leads: l["mx"] = None
    progress(0.93)

    for l in leads:
        l["_score"] = score_lead(l)
        l["_kanaal"] = kanaal_advies(l)
        l["_subject"], l["_body"], l["_fu2"], l["_fu3"] = maak_mail(l, brand)

    leads.sort(key=lambda l: (-l["_score"], l["type"], l["naam"]))
    t = 0
    for l in leads:
        if l["email"] and l.get("mx") is not False:
            l["_batch"] = f"Dag {t // max(n_per_dag, 1) + 1}"; t += 1
        else: l["_batch"] = ""
    progress(1.0)
    return leads

# ------------------------------------------------------------------ UI
st.set_page_config(page_title="VoiceStamp Lead Finder", page_icon="🎙️", layout="wide")

# ---- Huisstijl: de donkere VoiceStamp-look (kleuren en vormen uit de VoiceStamp-app) ----
VS = {"bg": "#07060b", "card": "rgba(20,16,32,0.9)", "line": "rgba(255,255,255,0.08)",
      "amber": "#FFB547", "on_amber": "#1b1206", "lilac": "#c4b5fd", "purple": "#7c3aed",
      "soft": "#8b84a3", "body": "#cfc9de"}
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');

  /* lettertype: alleen op tekstelementen, zodat de pijltjes en icoontjes heel blijven */
  .stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
  .stApp button, .stApp input, .stApp textarea, .stApp [data-baseweb="select"] {
      font-family: 'Manrope', system-ui, -apple-system, 'Segoe UI', sans-serif; }
  .stApp { background: radial-gradient(900px 480px at 0% -8%, rgba(124,58,237,0.20), transparent 60%), #07060b; }
  .block-container { max-width: 1240px; padding-top: 4.2rem; }
  h1, h2, h3, h4 { font-weight: 800 !important; letter-spacing: -0.01em; }
  h3 { font-size: 18px !important; margin-top: 4px; }
  h4 { font-size: 15px !important; }
  hr { border-color: rgba(255,255,255,0.08) !important; margin: 26px 0 10px 0 !important; }
  [data-testid="stCaptionContainer"], .stApp small { color: #8b84a3; }

  /* bewegingen, zoals in de VoiceStamp-app: zacht inkomen (450 ms, 12 px omhoog), knoppen drukken 4% in */
  @keyframes vsIn { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: none; } }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > [data-testid="stElementContainer"],
  [role="tabpanel"] [data-testid="stVerticalBlock"] > .element-container { animation: vsIn .45s cubic-bezier(.215,.61,.355,1) both; }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > :nth-child(2) { animation-delay: .05s; }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > :nth-child(3) { animation-delay: .10s; }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > :nth-child(4) { animation-delay: .15s; }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > :nth-child(5) { animation-delay: .20s; }
  [role="tabpanel"] [data-testid="stVerticalBlock"] > :nth-child(n+6) { animation-delay: .25s; }
  .vs-kpi, .vs-card, .vs-insight { animation: vsIn .45s cubic-bezier(.215,.61,.355,1) both; }
  .vs-kpi:nth-child(2) { animation-delay: .06s; } .vs-kpi:nth-child(3) { animation-delay: .12s; }
  .vs-kpi:nth-child(4) { animation-delay: .18s; } .vs-kpi:nth-child(5) { animation-delay: .24s; }
  .vs-kpi:nth-child(6) { animation-delay: .30s; }
  .vs-fill { animation: vsGrow .8s cubic-bezier(.215,.61,.355,1) both; transform-origin: left; }
  @keyframes vsGrow { from { transform: scaleX(0); } to { transform: scaleX(1); } }
  @media (prefers-reduced-motion: reduce) { * { animation: none !important; transition: none !important; } }

  /* kop */
  .vs-top { display:flex; align-items:center; gap:12px; margin-bottom: 4px; }
  .vs-logo { width:44px; height:44px; border-radius:14px; background:#FFB547; flex:0 0 auto;
             display:flex; align-items:center; justify-content:center; }
  .vs-title { font-size:18px; font-weight:800; line-height:1.15; color:#fff; }
  .vs-sub { color:#8b84a3; font-size:12.5px; margin-top:1px; }

  /* tabbladen: een vlakke balk, het actieve tabblad is een lichtgrijs vlak */
  [role="tablist"] { gap:6px; background:rgba(255,255,255,0.05); border-radius:12px; padding:4px;
        border-bottom:none !important; box-shadow:none !important; margin-bottom: 6px; }
  [role="tab"] { flex:1 1 0; justify-content:center; border-radius:9px !important; padding:9px 8px !important;
        height:auto !important; border:none !important; background:transparent !important;
        transition: background .2s ease, color .2s ease; }
  [role="tab"] p, [role="tab"] div, [role="tab"] span { margin:0; color:#8b84a3 !important; font-size:12.5px !important; font-weight:700 !important; }
  [role="tab"]:hover { background:rgba(255,255,255,0.05) !important; }
  [role="tab"][aria-selected="true"] { background:rgba(255,255,255,0.10) !important; }
  [role="tab"][aria-selected="true"] p, [role="tab"][aria-selected="true"] div,
  [role="tab"][aria-selected="true"] span { color:#fff !important; }
  [data-baseweb="tab-highlight"], [data-baseweb="tab-border"], .react-aria-SelectionIndicator, [role="tablist"]::after { display:none !important; }
  [data-testid="stTabs"] > div:first-child { border-bottom:none !important; box-shadow:none !important; }
  /* binnenste tabbladen (Mail 1, 2, 3): kleinere pillen */
  [role="tabpanel"] [role="tablist"] { background:transparent; padding:0; gap:6px; }
  [role="tabpanel"] [role="tab"] { flex:0 0 auto; border:1px solid rgba(255,255,255,0.12) !important; border-radius:999px !important; padding:7px 12px !important; }
  [role="tabpanel"] [role="tab"][aria-selected="true"] { background:#FFB547 !important; border-color:#FFB547 !important; }
  [role="tabpanel"] [role="tab"][aria-selected="true"] p, [role="tabpanel"] [role="tab"][aria-selected="true"] div,
  [role="tabpanel"] [role="tab"][aria-selected="true"] span { color:#1b1206 !important; }

  /* knoppen: hoofdknop = volle amber pil, de rest = pil met amber rand */
  .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
      border-radius:999px; font-weight:700; padding:0.45rem 1.2rem; background:transparent;
      border:1.5px solid rgba(255,181,71,0.5); color:#FFB547;
      transition: transform .15s ease, filter .15s ease, background .2s ease, border-color .2s ease; }
  .stButton > button p, .stDownloadButton > button p { color:#FFB547; font-weight:700; }
  .stButton > button:hover, .stDownloadButton > button:hover { background:rgba(255,181,71,0.10); border-color:#FFB547; color:#FFB547; }
  .stButton > button:active, .stDownloadButton > button:active, .stFormSubmitButton > button:active { transform:scale(0.96); }
  .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"],
  button[data-testid="stBaseButton-primary"], button[data-testid="stBaseButton-primaryFormSubmit"] {
      background:#FFB547; border-color:#FFB547; color:#1b1206; }
  .stButton > button[kind="primary"] p, button[data-testid="stBaseButton-primary"] p,
  .stDownloadButton > button[kind="primary"] p, button[data-testid="stBaseButton-primaryFormSubmit"] p { color:#1b1206 !important; font-weight:800; }
  .stButton > button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
      filter:brightness(1.07); background:#FFB547; border-color:#FFB547; color:#1b1206; }
  .stButton > button:disabled { opacity:0.4; }

  /* zijbalk: zelfde kaartstijl als het dashboard */
  [data-testid="stSidebar"] { background:#0c0a14; }
  [data-testid="stSidebar"] [data-testid="stSidebarHeader"] { padding-bottom:0; }
  [data-testid="stSidebar"] h2 { color:#c4b5fd; font-size:10.5px !important; font-weight:800 !important; letter-spacing:1.6px; text-transform:uppercase; margin:6px 0 4px 0; }
  [data-testid="stSidebar"] [data-testid="stExpander"] { border-radius:16px; }
  [data-testid="stSidebar"] label p { font-size:13px; color:#cfc9de; }
  [data-testid="stSidebar"] .stButton > button { width:100%; }
  [data-testid="stSidebar"] hr { margin:16px 0 !important; }
  [data-baseweb="tag"] { background:rgba(196,181,253,0.16) !important; border-radius:999px !important; }
  [data-baseweb="tag"] span { color:#fff !important; font-weight:600; }
  /* kaarten en velden */
  [data-testid="stMetric"] { background:rgba(20,16,32,0.9); border:1px solid rgba(255,255,255,0.08); border-radius:16px; padding:12px 14px; }
  [data-testid="stMetricLabel"] p { color:#8b84a3; font-size:11.5px; font-weight:700; }
  [data-testid="stMetricValue"] { font-weight:900; }
  [data-testid="stExpander"] { background:rgba(20,16,32,0.9); border:1px solid rgba(255,255,255,0.08); border-radius:20px; }
  [data-testid="stExpander"] summary { font-weight:700; }
  [data-testid="stAlert"] { border-radius:16px; }
  [data-testid="stSidebar"] { border-right:1px solid rgba(255,255,255,0.08); }
  [data-testid="stDataFrame"], [data-testid="stDataEditor"] { border:1px solid rgba(255,255,255,0.08); border-radius:16px; overflow:hidden; }
  [data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="select"] > div { border-radius:12px; }
  [data-testid="stVerticalBlockBorderWrapper"] { border-radius:20px; background:rgba(20,16,32,0.9); border-color:rgba(255,255,255,0.08); }

  /* bouwstenen uit het VoiceStamp-dashboard */
  .vs-insight { margin:10px 0 0 0; padding:18px; border-radius:22px; border:1px solid rgba(196,181,253,0.25);
      background: radial-gradient(120% 140% at 0% 0%, rgba(124,58,237,0.55), rgba(124,58,237,0.12) 55%, rgba(20,16,32,0.9)); }
  .vs-insight small { color:#c4b5fd; font-size:10.5px; font-weight:800; letter-spacing:1.6px; }
  .vs-insight h2 { color:#fff; font-size:24px; line-height:29px; font-weight:800; margin:6px 0 0 0; padding:0; letter-spacing:-0.4px; }
  .vs-insight p { color:#ddd6f3; font-size:14px; line-height:20px; margin:8px 0 0 0; }
  .vs-kpis { display:flex; gap:8px; margin:10px 0 4px 0; }
  .vs-kpi { flex:1; min-width:0; padding:12px; border-radius:16px; background:rgba(20,16,32,0.9); border:1px solid rgba(255,255,255,0.08); }
  .vs-kpi .l { color:#8b84a3; font-size:11.5px; font-weight:700; }
  .vs-kpi .v { color:#fff; font-size:22px; font-weight:900; margin-top:4px; }
  .vs-kpi .s { color:#6f6884; font-size:10.5px; margin-top:4px; line-height:14px; }
  .vs-kpi.amber { border-color:rgba(255,181,71,0.35); background:rgba(255,181,71,0.08); }
  .vs-card { padding:16px; border-radius:20px; background:rgba(20,16,32,0.9); border:1px solid rgba(255,255,255,0.08); margin:8px 0; }
  .vs-card h4 { color:#fff; font-size:15px; font-weight:800; margin:0 0 10px 0; padding:0; }
  .vs-f { margin:7px 0; }
  .vs-f .r { display:flex; justify-content:space-between; font-size:13.5px; color:#e2dfea; }
  .vs-f .r b { color:#fff; font-weight:900; }
  .vs-tr { height:9px; border-radius:5px; background:rgba(255,255,255,0.07); margin-top:5px; overflow:hidden; }
  .vs-tr i { display:block; height:9px; border-radius:5px; }
  .vs-f .n { color:#8b84a3; font-size:11.5px; margin-top:3px; }
  .vs-step { min-height:120px; padding:16px 18px; background:rgba(20,16,32,0.9); border:1px solid rgba(255,255,255,0.08); border-radius:20px; height:100%; color:#cfc9de; }
  .vs-step b { color:#FFB547; font-size:15px; }
  .vs-px { display:flex; gap:3px; } .vs-px > div { flex:1; display:flex; flex-direction:column; gap:3px; }
  .vs-px span { height:11px; border-radius:2px; background:rgba(255,255,255,0.04); }
  .vs-px span.on { background:rgba(167,139,250,0.8); } .vs-px > div:last-child span.on { background:#FFB547; }
  .vs-ax { display:flex; justify-content:space-between; color:#6f6884; font-size:11px; margin-top:6px; }
  .vs-heat { display:grid; gap:3px; align-items:center; } .vs-heat i { height:22px; border-radius:3px; display:block; }
  .vs-heat em { color:#6f6884; font-size:10.5px; font-style:normal; }
  .vs-stack { display:flex; height:12px; border-radius:6px; overflow:hidden; gap:2px; }
  .vs-legend { display:flex; flex-wrap:wrap; gap:6px 14px; margin-top:10px; color:#cfc9de; font-size:12.5px; }
  .vs-legend i { display:inline-block; width:9px; height:9px; border-radius:5px; margin-right:6px; }
  .vs-ev { display:flex; gap:12px; align-items:center; margin:8px 0; color:#ece9f3; font-size:13.5px; }
  .vs-ev .ic { width:30px; height:30px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex:none; font-size:14px; }
  .vs-ev small { display:block; color:#6f6884; font-size:11.5px; margin-top:1px; }
  .vs-muted { color:#8b84a3; font-size:13.5px; line-height:19px; }
  /* periodekeuze als pillen (segmented control), rechts uitgelijnd */
  [data-testid="stButtonGroup"] { justify-content:flex-end; }
  [data-testid="stButtonGroup"] button { border-radius:999px !important; border:1px solid rgba(255,255,255,0.12) !important;
        background:transparent !important; padding:7px 12px !important; min-height:0 !important; margin-left:6px; transition:background .2s, transform .15s; }
  [data-testid="stButtonGroup"] button p { font-size:12.5px !important; font-weight:700 !important; color:#cfc9de !important; }
  [data-testid="stButtonGroup"] button[aria-checked="true"], [data-testid="stButtonGroup"] button[aria-pressed="true"],
  [data-testid="stButtonGroup"] button[kind="segmented_controlActive"] { background:#FFB547 !important; border-color:#FFB547 !important; }
  [data-testid="stButtonGroup"] button[aria-checked="true"] p, [data-testid="stButtonGroup"] button[aria-pressed="true"] p,
  [data-testid="stButtonGroup"] button[kind="segmented_controlActive"] p { color:#1b1206 !important; }
  [data-testid="stButtonGroup"] button:active { transform:scale(0.96); }
  /* periodekeuze als pillen (terugval voor oudere Streamlit) */
  [data-testid="stRadio"] [role="radiogroup"] { gap:6px; flex-wrap:wrap; }
  [data-testid="stRadio"] [role="radiogroup"] label { padding:7px 12px; border-radius:999px; border:1px solid rgba(255,255,255,0.12); transition:background .2s, border-color .2s; cursor:pointer; margin:0; }
  [data-testid="stRadio"] [role="radiogroup"] label > div:first-child { display:none; }
  [data-testid="stRadio"] [role="radiogroup"] label p { font-size:12.5px; font-weight:700; color:#cfc9de; }
  [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) { background:#FFB547; border-color:#FFB547; }
  [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) p { color:#1b1206; }
  @media (max-width: 760px) { .vs-kpis { flex-wrap:wrap; } .vs-kpi { flex:1 1 40%; } }
</style>
""", unsafe_allow_html=True)

import html as _html

def vs_insight(label, title, sub=""):
    """Grote paarse kaart bovenaan een tabblad (zoals het hoofdinzicht in het VoiceStamp-dashboard)."""
    st.markdown(
        f'<div class="vs-insight"><small>{_html.escape(label.upper())}</small>'
        f'<h2>{_html.escape(title)}</h2>' + (f'<p>{_html.escape(sub)}</p>' if sub else '') + '</div>',
        unsafe_allow_html=True)

def vs_kpis(items):
    """Rij cijfertegels. items = [(label, waarde, onderschrift, 'amber' of '')]."""
    tegels = "".join(
        f'<div class="vs-kpi {(it[3] if len(it) > 3 else "")}"><div class="l">{_html.escape(str(it[0]))}</div>'
        f'<div class="v">{_html.escape(str(it[1]))}</div>'
        + (f'<div class="s">{_html.escape(str(it[2]))}</div>' if len(it) > 2 and it[2] else '') + '</div>'
        for it in items)
    st.markdown(f'<div class="vs-kpis">{tegels}</div>', unsafe_allow_html=True)

def vs_funnel(titel, rijen):
    """Kaart met voortgangsbalkjes. rijen = [(label, waarde, kleur, notitie)]."""
    top = max([r[1] for r in rijen] + [1])
    delen = []
    for r in rijen:
        breed = max(2, round(r[1] / top * 100)) if r[1] else 0
        note = f'<div class="n">{_html.escape(r[3])}</div>' if len(r) > 3 and r[3] else ''
        delen.append(f'<div class="vs-f"><div class="r"><span>{_html.escape(r[0])}</span><b>{r[1]}</b></div>'
                     f'<div class="vs-tr"><i class="vs-fill" style="width:{breed}%;background:{r[2]}"></i></div>{note}</div>')
    st.markdown(f'<div class="vs-card"><h4>{_html.escape(titel)}</h4>{"".join(delen)}</div>', unsafe_allow_html=True)

def vs_card(titel, inhoud_html, sub=""):
    st.markdown(f'<div class="vs-card"><h4>{_html.escape(titel)}</h4>'
                + (f'<div class="vs-muted" style="margin:-6px 0 10px 0">{_html.escape(sub)}</div>' if sub else '')
                + inhoud_html + '</div>', unsafe_allow_html=True)

def vs_pixels_html(data):
    """Blokjesgrafiek: data = [(label, aantal)]; de laatste kolom (vandaag) is amber."""
    rows = 8; mx = max([d[1] for d in data] + [1])
    kol = []
    for _, n in data:
        lit = round(n / mx * rows)
        kol.append("<div>" + "".join(f'<span class="{"on" if (rows - r <= lit and n > 0) else ""}"></span>' for r in range(rows)) + "</div>")
    return (f'<div class="vs-px">{"".join(kol)}</div>'
            f'<div class="vs-ax"><span>{data[0][0]}</span><span>{data[-1][0]}</span></div>')

def vs_heat_html(grid, uren):
    """Warmtekaart weekdag x uur. grid[dag][uur] = aantal."""
    mx = max([v for r in grid for v in r] + [1])
    dagen = ["ma", "di", "wo", "do", "vr", "za", "zo"]
    kop = "<span></span>" + "".join(f'<em style="text-align:center">{h if h % 3 == 0 else ""}</em>' for h in uren)
    rij = ""
    for d in range(7):
        rij += f"<em>{dagen[d]}</em>" + "".join(
            f'<i style="background:rgba(196,181,253,{0.06 + grid[d][h] / mx * 0.94:.2f})"></i>' for h in uren)
    return f'<div class="vs-heat" style="grid-template-columns:22px repeat({len(uren)},1fr)">{kop}{rij}</div>'

def vs_stack_html(delen):
    """Gestapelde balk met legenda. delen = [(label, aantal, kleur)]."""
    tot = sum(d[1] for d in delen) or 1
    balk = "".join(f'<div style="flex:{d[1]};background:{d[2]}"></div>' for d in delen if d[1])
    leg = "".join(f'<span><i style="background:{d[2]}"></i>{_html.escape(d[0])} {round(d[1] / tot * 100)}%</span>' for d in delen if d[1])
    return f'<div class="vs-stack">{balk}</div><div class="vs-legend">{leg}</div>'

def vs_events_html(events):
    """Net gebeurd. events = [(glyph, kleur, tekst, tijd)]."""
    return "".join(f'<div class="vs-ev"><div class="ic" style="background:{k}22;color:{k}">{g}</div>'
                   f'<div>{_html.escape(t)}<small>{_html.escape(w)}</small></div></div>' for g, k, t, w in events)

def tijd_geleden(dt):
    import pandas as _pd
    sec = (_pd.Timestamp.now(tz="UTC") - dt).total_seconds()
    if sec < 3600: return f"{max(1, int(sec // 60))} min geleden"
    if sec < 86400: return f"{int(sec // 3600)} uur geleden"
    d = int(sec // 86400)
    return "gisteren" if d == 1 else f"{d} dagen geleden"

def crm_recent_activities(client, user_id, days=90):
    try:
        from datetime import datetime as _dtm, timedelta as _td
        sinds = (_dtm.now(__import__("datetime").timezone.utc).replace(tzinfo=None) - _td(days=days)).isoformat()
        r = (client.table("activities").select("*").eq("user_id", user_id)
             .gte("created_at", sinds).order("created_at", desc=True).execute())
        return r.data or []
    except Exception:
        return []

_LOGO = ('<svg width="26" height="26" viewBox="0 0 30 30" fill="none" xmlns="http://www.w3.org/2000/svg">'
         '<rect x="3"  y="12" width="3.6" height="6"  rx="1.8" fill="#1b1206"/>'
         '<rect x="8.6" y="7"  width="3.6" height="16" rx="1.8" fill="#1b1206"/>'
         '<rect x="14.2" y="3" width="3.6" height="24" rx="1.8" fill="#1b1206"/>'
         '<rect x="19.8" y="8" width="3.6" height="14" rx="1.8" fill="#1b1206"/>'
         '<rect x="25.4" y="12" width="3.6" height="6" rx="1.8" fill="#1b1206"/></svg>')
st.markdown(
    f'<div class="vs-top"><div class="vs-logo">{_LOGO}</div>'
    '<div><div class="vs-title">VoiceStamp Lead Finder</div>'
    '<div class="vs-sub">Vind leads, schrijf de mails, beheer je opvolging.</div></div></div>',
    unsafe_allow_html=True)

with st.sidebar:
    st.header("Instellingen")
    if CP_OK:
        groep = st.selectbox("Hoofdgroep", list(CP.CAMPAGNES.keys()))
        _namen = [x["naam"] for x in CP.CAMPAGNES[groep]]
        gekozen = st.multiselect("Categorieën", _namen, default=_namen,
                                 help="Haal er weg wat je nu niet wilt. Alles tegelijk kan, het is één zoekopdracht.")
        for _x in CP.subs_van(groep, gekozen):
            if _x.get("let_op"):
                st.caption(f"**{_x['naam']}:** {_x['let_op']}")
        campagne = groep
    else:
        st.warning("campagnes.py ontbreekt naast app.py. Je ziet nu alleen de oude campagnes.")
        campagne = st.selectbox("Campagne", list(PRESETS.keys()))
    heel = st.checkbox("Heel Nederland (traag, tabblad open houden)", value=False)
    provincie = st.selectbox("Provincie", PROVINCIES, index=3, disabled=heel)
    st.divider()
    scrape = st.checkbox("Websites verrijken (e-mail, socials, openingszin)", value=True)
    mx_check = st.checkbox("E-mail controleren (MX)", value=True, disabled=not HAVE_DNS)
    dedupe_dom = st.checkbox("Max 1 per domein/keten", value=False)
    snelheid = st.select_slider("Snelheid (parallelle verwerking)",
                                options=[4, 6, 8, 10, 12], value=8,
                                help="Hoger = sneller. De app blijft altijd netjes tegenover elk "
                                     "afzonderlijk domein (minimaal 1 sec ertussen).")
    n_per_dag = st.number_input("Aantal per verzenddag", 5, 200, 25, step=5)
    max_sites = st.number_input("Max sites verrijken (0 = alle)", 0, 5000, 0, step=10)
    st.divider()
    up = st.file_uploader("Al-verstuurde lijst (csv/xlsx, optioneel)", type=["csv", "xlsx", "xlsm"])
    with st.expander("Jouw merk & pitch (optioneel)"):
        merk = st.text_input("Productnaam", "VoiceStamp")
        aanhef_keuze = st.radio(
            "Aanhef", ["Beste team van [bedrijfsnaam]", "Beste team"], index=0,
            help="Persoonlijk (met bedrijfsnaam) scoort meestal beter. 'Beste team' is voor "
                 "iedereen identiek.")
        links_blok = st.text_area("Links-blok in de mail", LINKS, height=110)
        eigen_mid = st.text_area(
            "Eigen pitch (optioneel)", "",
            help="Vervangt de standaard segmenttekst (het middenstuk) voor de hele campagne. "
                 "Gebruik {naam}, {plaats}, {haak}. Laat leeg voor de ingebouwde teksten.")
        eigen_body = st.text_area(
            "Volledige standaardmail (optioneel)", "", height=200,
            help="Vervangt de HELE mail (opening + midden + slot) voor alle segmenten. Gebruik "
                 "{aanhef}, {naam}, {plaats}, {haak}, {merk}, {links}. Laat leeg om de ingebouwde "
                 "mails per segment te gebruiken. Eindig zelf met 'Met vriendelijke groet,'.")
    start = st.button("🚀 Start", type="primary", use_container_width=True)
    st.caption("Houd dit tabblad open tijdens het zoeken. Werk per provincie voor een snelle run. "
               "Dezelfde provincie opnieuw ophalen gaat direct (opgeslagen).")

    st.divider()
    with st.expander("Account / CRM", expanded=False):
        if not crm_beschikbaar():
            st.caption("CRM nog niet ingesteld. Zie LEES_MIJ_crm.md om je gratis database te koppelen.")
        elif st.session_state.get("sb_user"):
            st.success(f"Ingelogd als {st.session_state.get('sb_email','')}")
            if st.button("Uitloggen"):
                for k in ("sb_client", "sb_user", "sb_email",
                          "zoho_user", "zoho_pw", "zoho_host", "zoho_from", "zoho_sig"):
                    st.session_state.pop(k, None)
                st.rerun()
        else:
            st.text_input("E-mail", key="login_email")
            st.text_input("Wachtwoord", type="password", key="login_pw")
            cA, cB = st.columns(2)
            _em = (st.session_state.get("login_email") or "").strip()
            _pw = st.session_state.get("login_pw") or ""
            if cA.button("Inloggen"):
                if not _em or not _pw:
                    st.error("Vul een e-mail en wachtwoord in.")
                else:
                    try:
                        c, u = crm_auth(_em, _pw, registreren=False)
                        if u:
                            st.session_state["sb_client"] = c
                            st.session_state["sb_user"] = u.id
                            st.session_state["sb_email"] = _em
                            st.rerun()
                        else:
                            st.error("Inloggen mislukt. Controleer je gegevens.")
                    except Exception as e:
                        st.error(f"Inloggen mislukt: {e}")
            if cB.button("Registreren"):
                if not _em or not _pw:
                    st.error("Vul een e-mail en wachtwoord in.")
                elif len(_pw) < 6:
                    st.error("Wachtwoord moet minstens 6 tekens zijn.")
                else:
                    try:
                        crm_auth(_em, _pw, registreren=True)
                        st.info("Account aangemaakt. Bevestig eventueel je e-mail en log daarna in.")
                    except Exception as e:
                        st.error(f"Registreren mislukt: {e}")

    with st.expander("Mailen (Zoho)", expanded=False):
        if not st.session_state.get("sb_user"):
            st.caption("Log eerst in bij **Account / CRM** hierboven. De mailfunctie is alleen "
                       "beschikbaar als je bent ingelogd.")
        elif st.session_state.get("zoho_user"):
            st.success(f"Mail actief: {st.session_state['zoho_user']}")
            rest = MAIL_DAG_LIMIET
            if st.session_state.get("sb_user"):
                rest = MAIL_DAG_LIMIET - crm_sent_today(st.session_state["sb_client"], st.session_state["sb_user"])
            else:
                rest = MAIL_DAG_LIMIET - st.session_state.get("zoho_sent_today", 0)
            st.caption(f"Nog {max(rest,0)} van {MAIL_DAG_LIMIET} mails vandaag.")
            if st.button("Mail uitloggen"):
                for k in ("zoho_user", "zoho_pw", "zoho_host", "zoho_from", "zoho_sig"):
                    st.session_state.pop(k, None)
                st.rerun()
        else:
            st.caption("Log in om rechtstreeks vanuit de app te mailen (max 15 per dag).")
            zu = st.text_input("Zoho e-mail", key="zoho_email_in")
            zp = st.text_input("App-wachtwoord", type="password", key="zoho_pw_in",
                               help="Maak in Zoho een app-specifiek wachtwoord aan (niet je gewone wachtwoord).")
            zh = st.selectbox("Server", ["smtp.zoho.eu", "smtp.zoho.com"], key="zoho_host_in")
            zf = st.text_input("Afzendernaam", "Yusuf Tatlicioglu", key="zoho_from_in")
            zsig = st.text_area("Handtekening (onder de mail)",
                                "Yusuf Tatlicioglu\n0646756497\n\nFounder VoiceStamp\nhttps://www.voicestamp.nl/",
                                key="zoho_sig_in", height=90,
                                help="Komt onder 'Met vriendelijke groet,'. Volledige webadressen "
                                     "(https://... of www...) worden automatisch klikbaar in de mail.")
            if st.button("Mail inloggen"):
                if not (zu.strip() and zp):
                    st.error("Vul je Zoho-e-mail en app-wachtwoord in.")
                else:
                    st.session_state["zoho_user"] = zu.strip()
                    st.session_state["zoho_pw"] = zp
                    st.session_state["zoho_host"] = zh
                    st.session_state["zoho_from"] = zf.strip() or zu.strip()
                    st.session_state["zoho_sig"] = zsig
                    st.rerun()

brand = {"merk": merk.strip() or "VoiceStamp", "links": links_blok, "eigen_mid": eigen_mid,
         "eigen_body": eigen_body,
         "aanhef_modus": "algemeen" if aanhef_keuze == "Beste team" else "persoonlijk"}

if start and CP_OK and not gekozen:
    st.warning("Kies minstens één categorie in de zijbalk.")
    start = False

if start:
    if CP_OK:
        preset = {"subs": CP.subs_van(campagne, gekozen)}
    else:
        preset = PRESETS[campagne]
    verstuurd = laad_verstuurd(up.getvalue(), up.name) if up else set()
    bar = st.progress(0.0); box = st.empty(); t0 = time.time()
    try:
        leads = draai(preset, None if heel else provincie, heel, scrape, mx_check,
                      dedupe_dom, n_per_dag, max_sites, int(snelheid), verstuurd, brand,
                      lambda p: bar.progress(min(p, 1.0)), lambda m: box.info(m))
        st.session_state["leads"] = leads
        box.success(f"Klaar in {int(time.time()-t0)} seconden.")
    except Exception as e:
        box.error(f"Er ging iets mis: {e}")

tab_res, tab_crm, tab_out, tab_upload = st.tabs(
    ["📋 Resultaten & mails", "📇 Mijn CRM", "📤 Outbox & antwoorden", "📄 Bestand uploaden (HR)"])

with tab_res:
    if st.session_state.get("leads"):
        leads = st.session_state["leads"]
        met = sum(1 for l in leads if l["email"])
        haken = sum(1 for l in leads if l.get("haak"))
        top3 = sorted(leads, key=lambda l: l.get("_score", 0), reverse=True)[:1]
        vs_insight("Deze zoekactie", f"{len(leads)} leads gevonden, {met} met e-mailadres.",
                   (f"Hoogste score: {top3[0]['naam']} ({top3[0].get('_score', 0)}). " if top3 else "")
                   + "Sorteer op score, kopieer de mail of sla ze op in je CRM.")
        vs_kpis([("Leads totaal", len(leads), "gevonden"),
                 ("Met e-mail", met, f"{round(met / max(len(leads), 1) * 100)}% van alle leads"),
                 ("Met haak", haken, "persoonlijke openingszin"),
                 ("Heeft app", sum(1 for l in leads if l.get("heeft_app")), "heeft al een eigen app"),
                 ("Verzenddagen", (met // max(int(n_per_dag), 1)) + 1 if met else 0, f"bij {int(n_per_dag)} per dag")])

        f1, f2, f3 = st.columns([2, 1, 1])
        min_score = f1.slider("Toon alleen leads met score \u2265", 0, 100, 0, step=5)
        alleen_mail = f2.checkbox("Alleen met e-mail", value=False)
        alleen_haak = f3.checkbox("Alleen met haak", value=False)
        zicht = [l for l in leads if l.get("_score", 0) >= min_score
                 and (l["email"] or not alleen_mail) and (l.get("haak") or not alleen_haak)]

        df = pd.DataFrame([{"Score": l.get("_score", 0), "Naam": l["naam"], "Type": l["type"],
                            "Plaats": l["plaats"], "E-mail": l["email"], "Haak": l.get("haak", ""),
                            "App?": "ja" if l.get("heeft_app") else "", "Kanaal": l.get("_kanaal", ""),
                            "Batch": l.get("_batch", "")} for l in zicht])
        st.caption(f"{len(zicht)} van {len(leads)} leads getoond, gesorteerd op score.")
        st.dataframe(df, use_container_width=True, height=380, hide_index=True,
                     column_config={"Score": st.column_config.ProgressColumn(
                         "Score", min_value=0, max_value=100, format="%d")})

        # --- Kopieer-en-plak: volledige mail per lead ---
        st.subheader("Mail kopiëren per lead")
        mailbaar = [l for l in zicht if l["email"]]
        if mailbaar:
            keuze = st.selectbox(
                "Kies een lead",
                range(len(mailbaar)),
                format_func=lambda i: f"{mailbaar[i]['naam']} — {mailbaar[i]['email']}"
                                      f" (score {mailbaar[i].get('_score', 0)})")
            l = mailbaar[keuze]
            info = f"**Aan:** {l['email']}"
            if l.get("_kanaal"): info += f"  •  **Advies:** {l['_kanaal']}"
            if l.get("haak"): info += f"  •  **Haak:** {l['haak']}"
            st.markdown(info)
            tab1, tab2, tab3 = st.tabs(["Mail 1", "Opvolgmail 2", "Opvolgmail 3"])
            with tab1:
                onderwerp = st.text_input("Onderwerp", l["_subject"], key=f"subj_{keuze}")
                mailtekst = st.text_area("Mailtekst (pas gerust aan voor deze lead)",
                                         l["_body"], height=320, key=f"body_{keuze}")
                if st.session_state.get("sb_user") and st.session_state.get("zoho_user"):
                    # dagteller
                    gedaan = crm_sent_today(st.session_state["sb_client"], st.session_state["sb_user"])
                    rest = MAIL_DAG_LIMIET - gedaan
                    if st.button(f"✉️ Verstuur via Zoho ({max(rest,0)} over vandaag)", key=f"send_{keuze}"):
                        if rest <= 0:
                            st.warning("Je dagelijkse 15 mails zijn verstuurd. Morgen weer.")
                        else:
                            body_send = mailtekst + "\n" + st.session_state.get("zoho_sig", "")
                            try:
                                zoho_send(st.session_state["zoho_host"], 465,
                                          st.session_state["zoho_user"], st.session_state["zoho_pw"],
                                          st.session_state.get("zoho_from", st.session_state["zoho_user"]),
                                          l["email"], onderwerp, body_send)
                                st.success(f"Mail verstuurd naar {l['email']}.")
                                st.session_state["zoho_sent_today"] = gedaan + 1
                                if st.session_state.get("sb_user"):
                                    from datetime import date as _d
                                    cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                                    lid = crm_lead_id_by_email(cl, uid, l["email"])
                                    if not lid:
                                        crm_save(cl, uid, [l])
                                        lid = crm_lead_id_by_email(cl, uid, l["email"])
                                    if lid:
                                        crm_update(cl, lid, uid,
                                                   {"status": "Gemaild", "last_mailed": _d.today().isoformat()})
                                        crm_add_activity(cl, uid, lid, "Mail verstuurd", onderwerp)
                            except Exception as e:
                                st.error(f"Versturen mislukt: {e}")
                elif not st.session_state.get("sb_user"):
                    st.caption("Wil je direct vanuit de app mailen? Log links in bij **Account / CRM** "
                               "en daarna bij **Mailen (Zoho)**.")
                else:
                    st.caption("Wil je direct vanuit de app mailen? Log links in bij **Mailen (Zoho)**.")
            with tab2:
                st.code(l.get("_fu2", ""), language=None)
            with tab3:
                st.code(l.get("_fu3", ""), language=None)
            st.caption("Tip: klik rechtsboven in het grijze vak op het kopieer-icoon om de mail te kopiëren.")

            # --- Alle mails van één verzenddag in één keer ---
            st.markdown("**Of pak een hele verzenddag in één keer:**")
            dagen = sorted({l["_batch"] for l in mailbaar if l.get("_batch")},
                           key=lambda d: int(d.split()[-1]) if d.split()[-1].isdigit() else 999)
            if dagen:
                dag = st.selectbox("Verzenddag", dagen)
                dagleads = [l for l in mailbaar if l.get("_batch") == dag]
                blok = "\n\n════════════════════\n\n".join(
                    f"AAN: {l['email']}\nONDERWERP: {l['_subject']}\n\n{l['_body']}" for l in dagleads)
                st.caption(f"{len(dagleads)} mails in {dag}.")
                st.code(blok, language=None)
                st.download_button(f"⬇️ Excel van {dag}", build_xlsx(dagleads),
                                   file_name=f"voicestamp_leads_{dag.replace(' ', '_').lower()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Geen leads met e-mail in de huidige selectie.")

        st.markdown("**Download de volledige selectie:**")
        d1, d2 = st.columns(2)
        d1.download_button("⬇️ Excel (alle kolommen: mails, haak, score, kanaal)", build_xlsx(zicht),
                           file_name="voicestamp_leads.xlsx", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("⬇️ CSV voor Smartlead/Instantly", build_csv(zicht),
                           file_name="voicestamp_leads.csv", mime="text/csv", use_container_width=True)

        if st.session_state.get("sb_user"):
            if st.button("💾 Deze selectie opslaan in mijn CRM", use_container_width=True):
                try:
                    n = crm_save(st.session_state["sb_client"], st.session_state["sb_user"], zicht)
                    # log 'Lead opgeslagen' voor de nieuw toegevoegde leads
                    try:
                        cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                        for l in zicht:
                            if l.get("email"):
                                lid = crm_lead_id_by_email(cl, uid, l["email"])
                                if lid:
                                    acts = crm_load_activities(cl, uid, lid)
                                    if not acts:
                                        crm_add_activity(cl, uid, lid, "Lead opgeslagen", l.get("naam", ""))
                    except Exception:
                        pass
                    st.success(f"{n} nieuwe leads opgeslagen in je CRM (dubbele e-mails overgeslagen).")
                except Exception as e:
                    st.error(f"Opslaan mislukt: {e}")
        elif crm_beschikbaar():
            st.caption("Log links in bij Account / CRM om deze leads op te slaan.")
    else:
        vs_insight("Klaar om te beginnen", "Kies links een campagne en provincie en klik op Start.",
                   "Je krijgt leads met e-mail, een score, een openingszin en een kant-en-klare mail.")
        st.write("")
        s1, s2, s3 = st.columns(3)
        s1.markdown('<div class="vs-step"><b>1. Kies</b><br>Selecteer links een campagne '
                    'en provincie en klik op Start.</div>', unsafe_allow_html=True)
        s2.markdown('<div class="vs-step"><b>2. Bekijk</b><br>Je krijgt leads met e-mail, een haak, '
                    'een score en een kant-en-klare mail per lead.</div>', unsafe_allow_html=True)
        s3.markdown('<div class="vs-step"><b>3. Werk bij</b><br>Kopieer de mails, download alles, '
                    'of sla op in je CRM en volg later op.</div>', unsafe_allow_html=True)

with tab_crm:
    if not st.session_state.get("sb_user"):
        if crm_beschikbaar():
            st.info("Log links in bij **Account / CRM** om je opgeslagen leads te beheren.")
        else:
            st.info("Het CRM is nog niet ingesteld. Zie **LEES_MIJ_crm.md** om je gratis database te koppelen.")
    else:

            with st.expander("📥 Bestaande Excel importeren"):
                st.caption("Zet je bestaande leadlijst (bijv. VoiceStamp_leads.xlsx) in \u00e9\u00e9n keer "
                           "in je CRM. Dubbele e-mailadressen worden overgeslagen. Bestaande statussen "
                           "blijven behouden.")
                imp = st.file_uploader("Kies je Excel of CSV", type=["xlsx", "xlsm", "csv"],
                                       key="crm_import_file")
                if imp is not None and st.button("Importeren", key="crm_import_btn"):
                    try:
                        toe, over = crm_import_uit_bytes(
                            st.session_state["sb_client"], st.session_state["sb_user"],
                            imp.getvalue(), imp.name)
                        st.success(f"{toe} leads ge\u00efmporteerd, {over} overgeslagen (dubbel of geen e-mail).")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Import mislukt: {e}")

            try:
                data = crm_load(st.session_state["sb_client"], st.session_state["sb_user"])
            except Exception as e:
                data = []
                st.error(f"Kon CRM niet laden: {e}")
            if not data:
                st.caption("Nog geen leads opgeslagen. Zoek hierboven en klik op 'Opslaan in mijn CRM'.")
            else:
                from datetime import date as _date
                cdf = pd.DataFrame(data)
                for kol, leeg in [("note", ""), ("status", "Nieuw"), ("segment", ""),
                                  ("company_name", ""), ("email", ""), ("score", 0),
                                  ("next_action", ""), ("next_action_date", None), ("dnc_reason", ""),
                                  ("last_mailed", None), ("email_status", ""), ("contact_name", ""),
                                  ("contact_role", ""), ("contact_email", ""),
                                  ("contact_email_status", "")]:
                    if kol not in cdf.columns:
                        cdf[kol] = leeg
                cdf["note"] = cdf["note"].fillna("")
                cdf["next_action"] = cdf["next_action"].fillna("")
                cdf["dnc_reason"] = cdf["dnc_reason"].fillna("")
                for _k in ("email_status", "contact_name", "contact_role", "contact_email",
                           "contact_email_status"):
                    cdf[_k] = cdf[_k].fillna("")
                # Supabase geeft datums als tekst; de tabel wil echte datums (anders een foutmelding)
                for _k in ("next_action_date", "last_mailed"):
                    cdf[_k] = pd.to_datetime(cdf[_k], errors="coerce").dt.date
                    cdf[_k] = cdf[_k].where(cdf[_k].notna(), None)

                # ---- Dashboard ----
                vandaag = _date.today().isoformat()
                def _telt(s): return int((cdf["status"] == s).sum())
                actief = ~cdf["status"].isin(["Klant", "Afgewezen", "Do not contact"])
                opvolg = cdf[(cdf["next_action_date"].notna()) &
                             (cdf["next_action_date"].astype(str) <= vandaag) & actief]
                _n = len(cdf)
                _gemaild = int(cdf["status"].isin(["Gemaild", "Follow-up nodig", "Reactie", "Afspraak", "Klant"]).sum())
                _reactie = int(cdf["status"].isin(["Reactie", "Afspraak", "Klant"]).sum())
                _afspraak = int(cdf["status"].isin(["Afspraak", "Klant"]).sum())
                vs_insight("Mijn CRM",
                           (f"{len(opvolg)} lead{'s' if len(opvolg) != 1 else ''} wacht{'en' if len(opvolg) != 1 else ''} op opvolging."
                            if len(opvolg) else f"{_n} leads in je CRM, niets dat vandaag opvolging nodig heeft."),
                           f"{_telt('Nieuw')} nieuw, {_telt('Gemaild')} gemaild, {_telt('Reactie')} reacties.")
                vs_kpis([("Nieuw", _telt("Nieuw"), "nog niet gemaild"),
                         ("Gemaild", _telt("Gemaild"), "wacht op antwoord"),
                         ("Follow-up nodig", len(opvolg), "vandaag of eerder", "amber" if len(opvolg) else ""),
                         ("Reacties", _telt("Reactie"), "heeft gereageerd"),
                         ("Afspraken", _telt("Afspraak"), "staat gepland"),
                         ("Klanten", _telt("Klant"), "gewonnen")])
                # ---- Analytics: zoals de analytics-pagina van VoiceStamp ----
                import pandas as _pd
                cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                kop_l, kop_r = st.columns([1, 1])
                _opts = ["7 dagen", "30 dagen", "90 dagen"]
                if hasattr(st, "segmented_control"):
                    periode = st.session_state.get("an_dagen") or "30 dagen"
                    with kop_r:
                        periode = st.segmented_control("Periode", _opts, default="30 dagen", key="an_dagen",
                                                       label_visibility="collapsed") or "30 dagen"
                else:
                    periode = kop_r.radio("Periode", _opts, index=1, horizontal=True, key="an_dagen",
                                          label_visibility="collapsed")
                n_dagen = int(periode.split()[0])
                acts_all = crm_recent_activities(cl, uid, 90)
                if acts_all:
                    adf = _pd.DataFrame(acts_all)
                    adf["t"] = _pd.to_datetime(adf["created_at"], utc=True, errors="coerce")
                    adf = adf.dropna(subset=["t"])
                    adf["lokaal"] = adf["t"].dt.tz_convert("Europe/Amsterdam")
                else:
                    adf = _pd.DataFrame(columns=["type", "lead_id", "t", "lokaal", "description"])
                grens = _pd.Timestamp.now(tz="Europe/Amsterdam").normalize() - _pd.Timedelta(days=n_dagen - 1)
                adf_p = adf[adf["lokaal"] >= grens] if len(adf) else adf
                mails_p = adf_p[adf_p["type"] == "Mail verstuurd"] if len(adf_p) else adf_p

                links, rechts = st.columns([2, 1])
                with links:
                    k1, k2 = st.columns(2)
                    with k1:
                        vs_funnel("Van lead naar klant", [
                            ("In je CRM", _n, "#c4b5fd"),
                            ("Gemaild", _gemaild, "#a78bfa", f"{_n - _gemaild} nog niet gemaild" if _n else ""),
                            ("Reactie gekregen", _reactie, "#38bdf8", f"{100 - round(_reactie / _gemaild * 100)}% reageert niet" if _gemaild else ""),
                            ("Afspraak", _afspraak, "#4ade80"),
                            ("Klant", _telt("Klant"), "#FFB547")])
                    with k2:
                        if "segment" in cdf.columns and _n:
                            tel = cdf["segment"].fillna("").replace("", "Onbekend").value_counts()
                            kl_ = ["#c4b5fd", "#FFB547", "#4ade80", "#38bdf8", "#f0abfc"]
                            delen = [(str(n), int(v), kl_[i]) for i, (n, v) in enumerate(tel.head(4).items())]
                            if len(tel) > 4:
                                delen.append(("Overig", int(tel.iloc[4:].sum()), kl_[4]))
                            vs_card("Waar komen je leads vandaan?", vs_stack_html(delen))
                        else:
                            vs_card("Waar komen je leads vandaan?", '<div class="vs-muted">Nog geen gegevens.</div>')
                    dagen = [(_pd.Timestamp.now(tz="Europe/Amsterdam").normalize() - _pd.Timedelta(days=i)) for i in range(n_dagen - 1, -1, -1)]
                    if len(mails_p):
                        per_dag = mails_p["lokaal"].dt.normalize().value_counts()
                        px_data = [(d.strftime("%d-%m"), int(per_dag.get(d, 0))) for d in dagen]
                    else:
                        px_data = [(d.strftime("%d-%m"), 0) for d in dagen]
                    vs_card("Elke dag gemaild", vs_pixels_html(px_data) if any(x[1] for x in px_data)
                            else '<div class="vs-muted">Nog geen mails verstuurd in deze periode.</div>')
                    if len(mails_p) >= 3:
                        grid = [[0] * 24 for _ in range(7)]
                        for tt in mails_p["lokaal"]:
                            grid[tt.weekday()][tt.hour] += 1
                        gebruikt = [tt.hour for tt in mails_p["lokaal"]]
                        uren = list(range(min(7, min(gebruikt)), max(20, max(gebruikt)) + 1))
                        vs_card("Wanneer mail je?", vs_heat_html(grid, uren), "Hoe lichter, hoe drukker.")
                with rechts:
                    if not len(opvolg):
                        vs_card("🔥 Vandaag opvolgen", '<div class="vs-muted">Niets om vandaag op te volgen.</div>')
                    else:
                        st.markdown('<div class="vs-card" style="margin-bottom:0"><h4 style="margin:0">🔥 Vandaag opvolgen</h4></div>', unsafe_allow_html=True)
                        st.dataframe(
                            opvolg[["company_name", "status", "next_action_date"]]
                            .rename(columns={"company_name": "Bedrijf", "status": "Status", "next_action_date": "Datum"}),
                            use_container_width=True, hide_index=True)
                    naam_van = dict(zip(cdf["id"], cdf["company_name"]))
                    ikoon = {"Mail verstuurd": ("✉", "#38bdf8"), "Antwoord ontvangen": ("↩", "#4ade80"), "Bounce": ("!", "#f87171"),
                             "Lead opgeslagen": ("+", "#c4b5fd"), "Status gewijzigd": ("●", "#FFB547"),
                             "Contactpersoon gevonden": ("@", "#f0abfc"), "Reeks goedgekeurd": ("✓", "#4ade80")}
                    ev = []
                    if len(adf):
                        for _, r in adf.head(6).iterrows():
                            g, k = ikoon.get(r["type"], ("•", "#c4b5fd"))
                            wie = naam_van.get(r.get("lead_id"), "")
                            ev.append((g, k, f"{r['type']}" + (f": {wie}" if wie else ""), tijd_geleden(r["t"])))
                    vs_card("Net gebeurd", vs_events_html(ev) if ev else
                            '<div class="vs-muted">Nog niets bijzonders. Zodra je mailt of een antwoord krijgt, zie je het hier.</div>')

                st.divider()
                st.markdown("### Alle leads")
                stat = st.multiselect("Filter op status", STATUS_OPTIES, default=[])
                toon = cdf[cdf["status"].isin(stat)] if stat else cdf
                st.caption(f"{len(toon)} van {len(cdf)} leads.")
                edit = st.data_editor(
                    toon[["company_name", "segment", "email", "contact_name", "contact_email",
                          "contact_email_status", "score", "status",
                          "next_action", "next_action_date", "last_mailed", "note", "dnc_reason"]],
                    use_container_width=True, height=420, key="crm_editor",
                    column_config={
                        "company_name": st.column_config.TextColumn("Bedrijf", disabled=True),
                        "segment": st.column_config.TextColumn("Type", disabled=True),
                        "email": st.column_config.TextColumn("E-mail", disabled=True),
                        "contact_name": st.column_config.TextColumn("Contactpersoon", disabled=True),
                        "contact_email": st.column_config.TextColumn("Adres contactpersoon", disabled=True),
                        "contact_email_status": st.column_config.TextColumn("Adresstatus", disabled=True,
                            help="published = stond op hun eigen site, verified = gecontroleerd, "
                                 "risky = onzeker, guess = gok (wordt niet gebruikt), invalid = ongeldig"),
                        "score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%d"),
                        "status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIES),
                        "next_action": st.column_config.TextColumn("Volgende actie"),
                        "next_action_date": st.column_config.DateColumn("Datum volg. actie"),
                        "last_mailed": st.column_config.DateColumn("Gemaild op", disabled=True),
                        "note": st.column_config.TextColumn("Notitie"),
                        "dnc_reason": st.column_config.TextColumn("Reden (Do not contact)"),
                    })
                if st.button("Wijzigingen opslaan", type="primary"):
                    fouten = 0
                    cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                    for idx in edit.index:
                        rid = cdf.loc[idx, "id"]
                        nd = edit.loc[idx, "next_action_date"]
                        nd = None if (nd is None or pd.isna(nd)) else str(nd)[:10]
                        nieuwe_status = edit.loc[idx, "status"]
                        velden = {"status": nieuwe_status,
                                  "next_action": edit.loc[idx, "next_action"] or "",
                                  "next_action_date": nd,
                                  "note": edit.loc[idx, "note"] or "",
                                  "dnc_reason": edit.loc[idx, "dnc_reason"] or ""}
                        try:
                            crm_update(cl, rid, uid, velden)
                            if nieuwe_status != cdf.loc[idx, "status"]:
                                crm_add_activity(cl, uid, rid, "Status gewijzigd", f"→ {nieuwe_status}")
                        except Exception:
                            fouten += 1
                    if fouten:
                        st.warning(f"Opgeslagen, maar {fouten} rijen gaven een fout.")
                    else:
                        st.success("Wijzigingen opgeslagen.")

                # ---- Beslisser vinden + mailreeks klaarzetten ----
                st.divider()
                st.markdown("### 🎯 Beslisser vinden & mailreeks klaarzetten")
                if not SP_OK:
                    st.warning("Zet **sellbird_pipeline.py** in dezelfde map als app.py om dit te gebruiken.")
                else:
                    st.caption(
                        "**Stap 1:** de app leest de eigen teampagina's van het bedrijf (geen LinkedIn) en zoekt "
                        "een eigenaar of manager. **Stap 2:** zij bouwt het adres van die persoon en markeert "
                        "het: published, verified, risky of guess. Een gok wordt nooit gebruikt. "
                        "**Stap 3:** zet een reeks van 3 mails klaar in je Outbox, nog niets verstuurd.")
                    _rk, _zk = _verifier_keys()
                    if not (_rk or _zk):
                        st.info("Geen verificatiedienst gekoppeld (REOON_API_KEY of ZEROBOUNCE_API_KEY in je secrets). "
                                "Dan blijft een zelf gebouwd adres een gok en wordt het niet gebruikt; "
                                "alleen adressen die op hun eigen site staan tellen. "
                                "Reoon heeft een gratis pakket.")
                    recs = {r["id"]: r for r in data}
                    st.markdown("#### ✨ Laat de app de beste leads kiezen")
                    st.caption("Geen selectie nodig. De app pakt de leads met de hoogste score die nog niet gemaild "
                               "zijn, zoekt de eigenaar of manager, schrijft de mails en zet alles klaar. "
                               "Jij leest en keurt goed in het tabblad Outbox.")
                    a1, a2 = st.columns([1, 2])
                    n_auto = a1.number_input("Aantal leads", 1, 25, 10, key="auto_n")
                    if a2.button("✨ Kies de beste leads en zet de mails klaar", type="primary", key="auto_go"):
                        from datetime import date as _dd0
                        _cl = st.session_state["sb_client"]; _uid = st.session_state["sb_user"]
                        try:
                            _bestaand = {o["lead_id"] for o in ob_load(_cl, _uid)
                                         if o["status"] in ("pending_review", "approved", "sent")}
                        except Exception as e:
                            _bestaand = set()
                            st.error(f"Outbox niet bereikbaar. Heb je de nieuwe supabase_setup.sql uitgevoerd? ({e})")
                        _bar = st.progress(0.0, text="Beslissers zoeken ...")
                        gemaakt_n, rijen_n, log_n = auto_dagbatch(
                            _cl, _uid, [dict(r) for r in data], int(n_auto), brand, _rk, _zk,
                            _dd0.today(), _bestaand, lambda p: _bar.progress(min(p, 1.0), text="Beslissers zoeken ..."))
                        _bar.empty()
                        if rijen_n:
                            try:
                                ob_create(_cl, _uid, rijen_n)
                                st.success(f"{gemaakt_n} reeksen staan klaar in je Outbox. Open het tabblad "
                                           "**Outbox & antwoorden** en keur ze goed.")
                            except Exception as e:
                                st.error(f"Klaarzetten mislukt: {e}")
                        else:
                            st.warning("Er waren geen leads om klaar te zetten. Zoek eerst nieuwe leads en sla ze op in je CRM.")
                        if log_n:
                            with st.expander(f"{len(log_n)} leads overgeslagen"):
                                st.write("\n".join(f"- {x}" for x in log_n))
                    st.markdown("#### Of kies zelf (optioneel)")
                    modus = st.radio("Welke leads?", ["Gekozen leads", "Eerste 25 zonder contactpersoon"],
                                     horizontal=True, key="sp_modus")
                    if modus == "Gekozen leads":
                        sel_idx = st.multiselect(
                            "Kies leads (max 25 per keer)", options=list(cdf.index), key="sp_sel",
                            format_func=lambda i: f"{cdf.loc[i,'company_name']} ({cdf.loc[i,'email']})")
                    else:
                        sel_idx = [i for i in cdf.index
                                   if not cdf.loc[i, "contact_name"] and cdf.loc[i, "status"] == "Nieuw"
                                   and str(cdf.loc[i].get("website") or "").strip()][:25]
                        st.caption(f"{len(sel_idx)} leads geselecteerd.")
                    sel_rows = [recs[cdf.loc[i, "id"]] for i in sel_idx[:25] if cdf.loc[i, "id"] in recs]

                    if st.button("🔎 Zoek eigenaar of manager + adres", disabled=not sel_rows, key="sp_zoek"):
                        cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                        bar = st.progress(0.0); res_rows = []
                        with ThreadPoolExecutor(max_workers=4) as ex:
                            futs = {ex.submit(zoek_beslisser, r, _rk, _zk): r for r in sel_rows}
                            for n, f in enumerate(as_completed(futs), 1):
                                r = futs[f]
                                try:
                                    upd, uitleg = f.result()
                                except Exception as e:
                                    upd, uitleg = None, f"fout: {e.__class__.__name__}"
                                if upd:
                                    try:
                                        crm_update(cl, r["id"], uid, upd)
                                        if upd.get("contact_name"):
                                            crm_add_activity(cl, uid, r["id"], "Contactpersoon gevonden", uitleg)
                                    except Exception as e:
                                        uitleg += f" (opslaan mislukt: {e})"
                                res_rows.append({"Bedrijf": r.get("company_name"), "Resultaat": uitleg})
                                bar.progress(n / len(futs))
                        st.session_state["sp_res"] = res_rows
                        st.rerun()
                    if st.session_state.get("sp_res"):
                        st.dataframe(pd.DataFrame(st.session_state["sp_res"]),
                                     use_container_width=True, hide_index=True)

                    st.markdown("#### Reeks klaarzetten")
                    from datetime import date as _dd
                    r1, r2 = st.columns(2)
                    start_dat = r1.date_input("Eerste mail op", _dd.today(), key="sp_start")
                    onbekend = r2.selectbox(
                        "Type niet herkend? Gebruik dan de tekst van",
                        ["(overslaan)"] + list(STD.keys()), key="sp_onbekend",
                        help="Voorkomt dat een winkel per ongeluk de campingtekst krijgt.")
                    st.checkbox("Risicovolle adressen toestaan (catch-all of onzeker)", key="sp_allow_risky")
                    st.caption("Mail 2 volgt standaard 4 dagen later en mail 3 na 8 dagen, altijd op een werkdag. "
                               "Stopt automatisch zodra iemand antwoordt of zich afmeldt.")
                    if st.button("📤 Zet reeks klaar in Outbox", disabled=not sel_rows, key="sp_reeks"):
                        cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                        try:
                            bestaand = {o["lead_id"] for o in ob_load(cl, uid)
                                        if o["status"] in ("pending_review", "approved", "sent")}
                        except Exception as e:
                            bestaand = set()
                            st.error(f"Outbox niet bereikbaar. Heb je de nieuwe supabase_setup.sql uitgevoerd? ({e})")
                        nieuw_rows, overgeslagen, gemaakt = [], [], 0
                        for r in sel_rows:
                            if r["id"] in bestaand:
                                overgeslagen.append(f"{r.get('company_name')}: heeft al een reeks")
                                continue
                            if (r.get("status") or "Nieuw") in ("Do not contact", "Afgewezen", "Klant"):
                                overgeslagen.append(f"{r.get('company_name')}: status {r.get('status')}")
                                continue
                            items, why = maak_reeks(r, brand, start_dat,
                                                    bool(st.session_state.get("sp_allow_risky")),
                                                    None if onbekend == "(overslaan)" else onbekend)
                            if items is None:
                                overgeslagen.append(f"{r.get('company_name')}: {why}")
                            else:
                                nieuw_rows.extend(items); gemaakt += 1
                        if nieuw_rows:
                            try:
                                ob_create(cl, uid, nieuw_rows)
                                st.success(f"{gemaakt} reeksen klaargezet. Ga naar het tabblad **Outbox & antwoorden** "
                                           "om ze te lezen en goed te keuren.")
                            except Exception as e:
                                st.error(f"Klaarzetten mislukt: {e}")
                        if overgeslagen:
                            st.warning("Overgeslagen:\n\n" + "\n".join(f"- {x}" for x in overgeslagen))

                # ---- Verwijderen ----
                st.divider()
                st.markdown("### Leads verwijderen")
                cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                vd1, vd2 = st.columns(2)
                with vd1:
                    weg = st.multiselect(
                        "Kies leads om te verwijderen",
                        options=list(cdf.index),
                        format_func=lambda i: f"{cdf.loc[i,'company_name']} ({cdf.loc[i,'email']})")
                    if st.button("Verwijder gekozen leads") and weg:
                        fout = 0
                        for i in weg:
                            try:
                                crm_delete(cl, cdf.loc[i, "id"], uid)
                            except Exception:
                                fout += 1
                        if fout:
                            st.warning(f"{len(weg)-fout} verwijderd, {fout} met een fout.")
                        else:
                            st.success(f"{len(weg)} leads verwijderd.")
                        st.rerun()
                with vd2:
                    st.caption("Alles wissen kan niet ongedaan worden gemaakt.")
                    bevestig = st.checkbox("Ja, verwijder al mijn leads")
                    if st.button("🗑️ Alle leads verwijderen", type="secondary") and bevestig:
                        try:
                            n = crm_delete_all(cl, uid)
                            st.success(f"Alle leads verwijderd ({n}).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Verwijderen mislukt: {e}")

                # ---- Mail versturen vanuit CRM (Zoho) ----
                st.divider()
                st.markdown("### Mail versturen")
                if not st.session_state.get("zoho_user"):
                    st.caption("Log links in bij **Mailen (Zoho)** om vanuit je CRM te mailen.")
                else:
                    mlijst = cdf[cdf["email"].fillna("") != ""]
                    if not len(mlijst):
                        st.caption("Geen leads met e-mail in je CRM.")
                    else:
                        idxs = list(mlijst.index)
                        mk = st.selectbox(
                            "Kies een lead om te mailen", idxs,
                            format_func=lambda i: f"{cdf.loc[i,'company_name']} ({cdf.loc[i,'email']})",
                            key="crm_mail_sel")
                        seg = LABEL_TO_SEG.get(str(cdf.loc[mk, "segment"]), "natuurcamping")
                        lead_obj = {"seg": seg, "naam": cdf.loc[mk, "company_name"] or "",
                                    "plaats": "", "_vn": "", "haak": "", "heeft_app": False,
                                    "email": cdf.loc[mk, "email"]}
                        subj, body, _, _ = maak_mail(lead_obj, brand)
                        ondw = st.text_input("Onderwerp", subj, key=f"crm_mail_subj_{mk}")
                        mtekst = st.text_area("Mailtekst (pas gerust aan voor deze lead)",
                                              body, height=320, key=f"crm_mail_body_{mk}")
                        gedaan = crm_sent_today(st.session_state["sb_client"], st.session_state["sb_user"])
                        rest = MAIL_DAG_LIMIET - gedaan
                        if st.button(f"✉️ Verstuur via Zoho ({max(rest,0)} over vandaag)", key="crm_send"):
                            if rest <= 0:
                                st.warning("Je dagelijkse 15 mails zijn verstuurd. Morgen weer.")
                            else:
                                body_send = mtekst + "\n" + st.session_state.get("zoho_sig", "")
                                try:
                                    from datetime import date as _d
                                    zoho_send(st.session_state["zoho_host"], 465,
                                              st.session_state["zoho_user"], st.session_state["zoho_pw"],
                                              st.session_state.get("zoho_from", st.session_state["zoho_user"]),
                                              lead_obj["email"], ondw, body_send)
                                    cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
                                    crm_update(cl, cdf.loc[mk, "id"], uid,
                                               {"status": "Gemaild", "last_mailed": _d.today().isoformat()})
                                    crm_add_activity(cl, uid, cdf.loc[mk, "id"], "Mail verstuurd", ondw)
                                    st.success(f"Mail verstuurd naar {lead_obj['email']}.")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Versturen mislukt: {e}")

                # ---- Contacthistorie per lead ----
                st.divider()
                st.markdown("### Contacthistorie")
                namen = cdf["company_name"].tolist()
                if namen:
                    kz = st.selectbox("Kies een lead", range(len(namen)),
                                      format_func=lambda i: namen[i], key="hist_sel")
                    lid = cdf.loc[cdf.index[kz], "id"]
                    nieuwe_notitie = st.text_input("Notitie toevoegen", key="hist_note")
                    if st.button("Notitie loggen") and nieuwe_notitie.strip():
                        crm_add_activity(st.session_state["sb_client"], st.session_state["sb_user"],
                                         lid, "Notitie", nieuwe_notitie.strip())
                        st.success("Genoteerd.")
                        st.rerun()
                    acts = crm_load_activities(st.session_state["sb_client"],
                                               st.session_state["sb_user"], lid)
                    if acts:
                        for a in acts:
                            datum = str(a.get("created_at", ""))[:10]
                            oms = f" — {a['description']}" if a.get("description") else ""
                            st.markdown(f"- **{datum}**  {a.get('type','')}{oms}")
                    else:
                        st.caption("Nog geen activiteiten voor deze lead.")

with tab_out:
    if not SP_OK:
        st.warning("Zet **sellbird_pipeline.py** in dezelfde map als app.py om de Outbox te gebruiken.")
    elif not st.session_state.get("sb_user"):
        st.info("Log links in bij **Account / CRM** om je Outbox te gebruiken.")
    else:
        from datetime import date as _dt, datetime as _datetime
        cl = st.session_state["sb_client"]; uid = st.session_state["sb_user"]
        st.caption("Verzenden gebeurt alleen terwijl deze app open staat: klik elke dag op **Verstuur nu** "
                   "voor wat vandaag aan de beurt is.")
        try:
            ob = ob_load(cl, uid); alle_leads = crm_load(cl, uid)
        except Exception as e:
            ob, alle_leads = [], []
            st.error(f"Kon de Outbox niet laden. Heb je de nieuwe **supabase_setup.sql** uitgevoerd? ({e})")
        leadmap = {l["id"]: l for l in alle_leads}
        allow_risky = bool(st.session_state.get("sp_allow_risky"))
        sig = st.session_state.get("zoho_sig", "")

        oa, ob_ = st.columns([1, 3])
        opt_aan = oa.checkbox("Afmeldregel onderaan", value=True, key="ob_optout_aan",
                              help="Bij zakelijke mail hoort een eenvoudige afmeldmogelijkheid.")
        opt_tekst = ob_.text_input("Tekst afmeldregel",
                                   "Liever geen mail meer van mij? Laat het even weten, dan stop ik direct.",
                                   key="ob_optout_tekst", disabled=not opt_aan)
        optout = opt_tekst if opt_aan else ""

        vandaag = _dt.today().isoformat()
        wacht = [o for o in ob if o["status"] == "pending_review"]
        goed = [o for o in ob if o["status"] == "approved"]
        due = [o for o in goed if (o.get("scheduled_for") or "") <= vandaag]
        _verstuurd = sum(1 for o in ob if o["status"] == "sent")
        vs_insight("Outbox",
                   (f"{len(wacht)} mail{'s' if len(wacht) != 1 else ''} wacht{'en' if len(wacht) != 1 else ''} op jouw akkoord."
                    if wacht else (f"{len(due)} mail{'s' if len(due) != 1 else ''} vandaag klaar om te versturen." if due
                                   else "Alles is afgehandeld.")),
                   "Niets gaat de deur uit zonder jouw goedkeuring.")
        vs_kpis([("Wacht op jouw akkoord", len(wacht), "lezen en goedkeuren", "amber" if wacht else ""),
                 ("Goedgekeurd", len(goed), "staat klaar"),
                 ("Vandaag aan de beurt", len(due), "kan nu verstuurd worden"),
                 ("Verstuurd", _verstuurd, "totaal")])

        # ---------------- 1. Goedkeuren
        st.divider()
        st.markdown("### 1. Lezen en goedkeuren")
        if not wacht:
            st.caption("Niets dat wacht. Zet een reeks klaar bij **Mijn CRM → Beslisser vinden & mailreeks klaarzetten**.")
        groepen = {}
        for o in wacht:
            groepen.setdefault(o["lead_id"], []).append(o)
        if groepen:
            if st.button(f"✅ Keur alle {len(groepen)} reeksen in één keer goed", type="primary", key="ob_ok_alle",
                         help="Keurt alle reeksen goed waar geen fout in zit. Reeksen met een fout blijven staan."):
                goed_n, blijft = 0, []
                for lid_, items_ in groepen.items():
                    items_.sort(key=lambda x: x["step"])
                    lead_ = leadmap.get(lid_, {}); to_ = items_[0]["to_email"]; stt_ = status_voor(lead_, to_)
                    fouten_ = []; nieuw_ = []
                    for it_ in items_:
                        subj_ = st.session_state.get(f"obs_{it_['id']}", it_["subject"])
                        body_ = st.session_state.get(f"obb_{it_['id']}", it_["body"])
                        f_, _ = SP.gate({"to_email": to_, "subject": subj_, "body": body_, "step": it_["step"]},
                                        lead_.get("status") or "Nieuw", stt_, allow_risky, mx_check=True,
                                        full_text=volledige_tekst(body_, sig, optout))
                        fouten_ += f_; nieuw_.append((it_, subj_, body_))
                    if fouten_:
                        blijft.append(f"{lead_.get('company_name', '?')}: {' '.join(sorted(set(fouten_)))}")
                        continue
                    for it_, subj_, body_ in nieuw_:
                        ob_update(cl, uid, it_["id"], {"status": "approved", "subject": subj_, "body": body_})
                    crm_add_activity(cl, uid, lid_, "Reeks goedgekeurd", f"{len(nieuw_)} mails gepland")
                    goed_n += 1
                st.session_state["ob_alle_msg"] = (goed_n, blijft)
                st.rerun()
            if st.session_state.get("ob_alle_msg"):
                _g, _b = st.session_state.pop("ob_alle_msg")
                st.success(f"{_g} reeksen goedgekeurd.")
                if _b:
                    st.warning("Niet goedgekeurd:\n\n" + "\n".join(f"- {x}" for x in _b))
            st.caption("Wil je eerst lezen? Open hieronder een lead. Je ziet alle 3 de mails en kunt ze aanpassen.")
        for lid, items in list(groepen.items())[:25]:
            items.sort(key=lambda x: x["step"])
            lead = leadmap.get(lid, {})
            to = items[0]["to_email"]
            stt = status_voor(lead, to)
            with st.expander(f"{lead.get('company_name','?')} → {to}  ({SP.STATUS_LABEL.get(stt, stt or 'niet gecontroleerd')})"):
                waarsch_all, fouten_all, nieuwe = [], [], []
                for it in items:
                    st.markdown(f"**Mail {it['step']}** — gepland op {it['scheduled_for']}")
                    subj = st.text_input("Onderwerp", it["subject"], key=f"obs_{it['id']}")
                    body = st.text_area("Tekst", it["body"], height=200, key=f"obb_{it['id']}")
                    f_, w_ = SP.gate({"to_email": to, "subject": subj, "body": body, "step": it["step"]},
                                     lead.get("status") or "Nieuw", stt, allow_risky, mx_check=False,
                                     full_text=volledige_tekst(body, sig, optout))
                    if stt is None:
                        w_.append("Dit adres is niet gecontroleerd (geimporteerd). De MX-controle gebeurt bij verzenden.")
                    fouten_all += [f"Mail {it['step']}: {x}" for x in f_]
                    waarsch_all += [f"Mail {it['step']}: {x}" for x in w_]
                    nieuwe.append((it, subj, body))
                for x in fouten_all:
                    st.error(x)
                for x in waarsch_all:
                    st.warning(x)
                if st.checkbox("Toon mail 1 met handtekening", key=f"ob_prev_{lid}"):
                    st.text(volledige_tekst(nieuwe[0][2], sig, optout))
                b1, b2 = st.columns(2)
                if b1.button("✅ Opslaan en hele reeks goedkeuren", key=f"ob_ok_{lid}",
                             type="primary", disabled=bool(fouten_all)):
                    slecht = []
                    for it, subj, body in nieuwe:
                        f2, _ = SP.gate({"to_email": to, "subject": subj, "body": body, "step": it["step"]},
                                        lead.get("status") or "Nieuw", stt, allow_risky, mx_check=True)
                        slecht += f2
                    if slecht:
                        st.error("Niet goedgekeurd: " + " ".join(sorted(set(slecht))))
                    else:
                        for it, subj, body in nieuwe:
                            ob_update(cl, uid, it["id"], {"status": "approved", "subject": subj, "body": body})
                        crm_add_activity(cl, uid, lid, "Reeks goedgekeurd", f"{len(nieuwe)} mails gepland")
                        st.rerun()
                if b2.button("❌ Afwijzen", key=f"ob_no_{lid}"):
                    for it, subj, body in nieuwe:
                        ob_update(cl, uid, it["id"], {"status": "rejected"})
                    st.rerun()
        if len(groepen) > 25:
            st.caption(f"Nog {len(groepen)-25} reeksen. Keur eerst deze goed of af.")

        # ---------------- 2. Versturen
        st.divider()
        st.markdown("### 2. Versturen")
        if not st.session_state.get("zoho_user"):
            st.caption("Log links in bij **Mailen (Zoho)** om te kunnen versturen.")
        else:
            rest = max(MAIL_DAG_LIMIET - crm_sent_today(cl, uid), 0)
            st.caption(f"Vandaag aan de beurt: **{len(due)}** mails. Je hebt nog **{rest}** van {MAIL_DAG_LIMIET} over vandaag.")
            pauze = st.slider("Pauze tussen mails (seconden)", 5, 60, 25, key="ob_pauze",
                              help="Een korte, wisselende pauze oogt natuurlijker dan alles tegelijk.")
            if st.button(f"✉️ Verstuur nu (max {rest})", disabled=not (due and rest), type="primary", key="ob_send"):
                verstuurd = 0; log = []
                due_sorted = sorted(due, key=lambda x: (x["scheduled_for"] or "", x["step"]))
                with st.status("Bezig met versturen...", expanded=True) as stat:
                    for it in due_sorted:
                        if verstuurd >= rest:
                            break
                        lead = leadmap.get(it["lead_id"], {})
                        sib = [x for x in ob if x["lead_id"] == it["lead_id"]]
                        naam = lead.get("company_name", "?")
                        # reeks in volgorde en niet twee mails kort na elkaar
                        if it["step"] > 1 and not any(x["step"] == it["step"] - 1 and x["status"] == "sent" for x in sib):
                            log.append(f"{naam}: mail {it['step']} wacht tot mail {it['step']-1} verstuurd is"); continue
                        recent = [x for x in sib if x["status"] == "sent" and (x.get("sent_at") or "")[:10] >=
                                  (_dt.today().fromordinal(_dt.today().toordinal() - 3)).isoformat()]
                        if recent:
                            log.append(f"{naam}: minder dan 3 dagen geleden gemaild, wacht"); continue
                        stt = status_voor(lead, it["to_email"])
                        f_, _w = SP.gate(it, lead.get("status") or "Nieuw", stt, allow_risky, mx_check=True)
                        if f_:
                            stop = (lead.get("status") or "Nieuw") in ("Do not contact", "Afgewezen", "Reactie", "Afspraak", "Klant")
                            ob_update(cl, uid, it["id"], {"status": "cancelled" if stop else "pending_review",
                                                          "error": " ".join(f_)})
                            log.append(f"{naam}: niet verstuurd ({' '.join(f_)})"); continue
                        mid1 = next((x.get("message_id") for x in sib if x["step"] == 1 and x.get("message_id")), None)
                        try:
                            if verstuurd:
                                time.sleep(random.randint(max(pauze - 10, 3), pauze + 10))
                            mid = zoho_send(st.session_state["zoho_host"], 465, st.session_state["zoho_user"],
                                            st.session_state["zoho_pw"],
                                            st.session_state.get("zoho_from", st.session_state["zoho_user"]),
                                            it["to_email"], it["subject"],
                                            volledige_tekst(it["body"], sig, optout),
                                            in_reply_to=mid1 if it["step"] > 1 else None)
                        except Exception as e:
                            ob_update(cl, uid, it["id"], {"error": str(e)})
                            log.append(f"{naam}: versturen mislukt ({e})"); stat.update(label="Gestopt door een fout"); break
                        nu = _dt.today().isoformat()
                        ob_update(cl, uid, it["id"], {"status": "sent", "message_id": mid, "error": None,
                                                      "sent_at": _datetime.now().isoformat()})
                        veld = {"last_mailed": nu}
                        if (lead.get("status") or "Nieuw") == "Nieuw":
                            veld["status"] = "Gemaild"
                        volg = next((x for x in sib if x["step"] == it["step"] + 1 and x["status"] == "approved"), None)
                        if volg:
                            veld["next_action"] = f"Mail {it['step']+1} gepland"
                            veld["next_action_date"] = volg["scheduled_for"]
                            # nooit eerder dan 4 dagen na deze mail
                            vroegst = SP.volgende_werkdag(_dt.fromordinal(_dt.today().toordinal() + 4)).isoformat()
                            if (volg["scheduled_for"] or "") < vroegst:
                                ob_update(cl, uid, volg["id"], {"scheduled_for": vroegst})
                                veld["next_action_date"] = vroegst
                        crm_update(cl, lead["id"], uid, veld)
                        crm_add_activity(cl, uid, lead["id"], "Mail verstuurd", f"Mail {it['step']}: {it['subject']}")
                        verstuurd += 1
                        st.write(f"✅ {naam}: mail {it['step']} verstuurd")
                    stat.update(label=f"{verstuurd} verstuurd", state="complete")
                st.session_state["ob_log"] = log
                st.session_state["ob_verstuurd"] = verstuurd
                time.sleep(1); st.rerun()
            if "ob_verstuurd" in st.session_state:
                st.success(f"{st.session_state.pop('ob_verstuurd')} mails verstuurd.")
            for x in st.session_state.pop("ob_log", []):
                st.warning(x)

        # ---------------- 3. Planning
        if ob:
            with st.expander("📅 Planning en geschiedenis"):
                pl = pd.DataFrame([{"Bedrijf": leadmap.get(o["lead_id"], {}).get("company_name", "?"),
                                    "Mail": o["step"], "Naar": o["to_email"], "Status": o["status"],
                                    "Gepland": o["scheduled_for"], "Opmerking": o.get("error") or ""}
                                   for o in ob])
                st.dataframe(pl, use_container_width=True, hide_index=True)

        # ---------------- 4. Antwoorden
        st.divider()
        st.markdown("### 3. Antwoorden verwerken")
        st.caption("Een antwoord, afmelding of bounce zet de lead op de juiste status en stopt de rest van de reeks. "
                   "Een antwoord van iemand anders bij hetzelfde bedrijf telt ook.")
        if st.session_state.get("zoho_user"):
            dagen = st.selectbox("Kijk terug", [3, 7, 14, 30], index=1, key="ob_dagen")
            if st.button("📥 Check mijn inbox", key="ob_inbox"):
                try:
                    host_imap = SP.imap_host_for(st.session_state["zoho_host"])
                    mails = SP.fetch_replies(host_imap, st.session_state["zoho_user"],
                                             st.session_state["zoho_pw"], since_days=int(dagen))
                    uit = verwerk_inbox(cl, uid, mails, alle_leads, st.session_state["zoho_user"])
                    st.success(f"{len(mails)} mails gelezen, {len(uit)} daarvan gingen over je leads.")
                    if uit:
                        st.dataframe(pd.DataFrame(uit), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"Inbox lezen mislukt: {e}")
                    st.info("IMAP moet aan staan voor je account. Volgens meldingen van gebruikers heeft het gratis "
                            "Zoho-pakket geen POP/IMAP/SMTP meer (dat zou ook inloggen via SMTP blokkeren). "
                            "Gebruik anders hieronder de handmatige invoer.")
        else:
            st.caption("Log links in bij **Mailen (Zoho)** om je inbox te kunnen lezen.")

        with st.expander("✍️ Antwoord handmatig verwerken (als je inbox niet te lezen is)"):
            if alle_leads:
                li = st.selectbox("Van welke lead?", range(len(alle_leads)),
                                  format_func=lambda i: f"{alle_leads[i].get('company_name')} ({alle_leads[i].get('email')})",
                                  key="man_lead")
                m_ond = st.text_input("Onderwerp", "Re:", key="man_ond")
                m_txt = st.text_area("Plak hier hun antwoord", height=140, key="man_txt")
                if st.button("Verwerk dit antwoord", disabled=not m_txt.strip(), key="man_go"):
                    soort, nw, uitleg = SP.classify_reply(m_ond, m_txt, alle_leads[li].get("email") or "")
                    if soort == "bounce":
                        verwerk_bounce(cl, uid, alle_leads[li], alle_leads[li].get("email") or "")
                    elif soort != "auto":
                        verwerk_antwoord(cl, uid, alle_leads[li], soort, nw, m_ond, m_txt)
                    st.success(f"Herkend als: {uitleg}. " + ("Reeks gestopt." if soort != "auto" else "Niets gewijzigd."))
                    st.rerun()

with tab_upload:
    st.header("Bestand uploaden (HR / onboarding)")
    st.caption("Voor je eigen lijst bedrijven (bijv. HR-leads). De mail gebruikt een "
               "onboarding-invalshoek, los van de toeristische mails.")
    try:
        from leadfinder_upload import enrich_row
        _UPLOAD_OK = True
    except Exception as _e:
        _UPLOAD_OK = False
        st.warning("De verrijkingsmodule (leadfinder_upload.py) is niet geladen. Upload werkt, "
                   "maar e-mails van websites halen kan uitstaan. Detail: " + str(_e))

    up_hr = st.file_uploader("CSV of Excel", type=["csv", "xlsx", "xls"], key="hr_upload")
    if up_hr is None:
        st.info("Upload een bestand met minstens een kolom bedrijfsnaam. Website, functie en "
                "invalshoek zijn optioneel maar maken de mail beter.")
    else:
        try:
            if up_hr.name.lower().endswith((".xlsx", ".xls")):
                dfu = pd.read_excel(up_hr)
            else:
                dfu = pd.read_csv(up_hr, sep=None, engine="python")
            dfu.columns = [str(c).strip() for c in dfu.columns]
            # Reparatie: alles in 1 kolom geplakt met ; of tab? Splits het alsnog netjes uit.
            if len(dfu.columns) == 1:
                enige = dfu.columns[0]
                for sep in (";", "\t", ","):
                    if sep in str(enige):
                        nieuwe_kol = [c.strip() for c in str(enige).split(sep)]
                        gesplitst = dfu[enige].astype(str).str.split(sep, expand=True)
                        if gesplitst.shape[1] == len(nieuwe_kol):
                            gesplitst.columns = nieuwe_kol
                            dfu = gesplitst
                            st.info(f"Bestand stond in \u00e9\u00e9n kolom; automatisch uitgesplitst naar "
                                    f"{len(nieuwe_kol)} kolommen.")
                        break
        except Exception as e:
            dfu = None
            st.error(f"Kon het bestand niet lezen: {e}")

        if dfu is not None and len(dfu):
            st.write(f"**{len(dfu)} rijen** ingelezen.")
            st.dataframe(dfu.head(8), use_container_width=True)
            cols = list(dfu.columns)
            geen = "— geen —"
            def _gok(keys, standaard=0):
                for i, c in enumerate(cols):
                    if any(k in c.lower() for k in keys):
                        return i
                return standaard
            c1, c2, c3 = st.columns(3)
            k_naam = c1.selectbox("Bedrijfsnaam", cols, index=_gok(("bedrijf", "naam", "company")))
            k_web = c2.selectbox("Website (optioneel)", [geen] + cols,
                                 index=1 + _gok(("website", "url", "site", "web")) if any(
                                     k in " ".join(cols).lower() for k in ("website", "url", "site")) else 0)
            k_func = c3.selectbox("Functie / aanhef (optioneel)", [geen] + cols,
                                  index=1 + _gok(("functie", "doelfunctie", "rol", "title")) if any(
                                      k in " ".join(cols).lower() for k in ("functie", "rol", "title")) else 0)
            k_mail = c1.selectbox("Bestaand e-mailadres (optioneel)", [geen] + cols,
                                  index=1 + _gok(("mail", "email")) if any(
                                      k in " ".join(cols).lower() for k in ("mail", "email")) else 0)
            k_inval = c2.selectbox("Invalshoek (optioneel)", [geen] + cols,
                                   index=1 + _gok(("invalshoek", "hoek", "reden")) if any(
                                       k in " ".join(cols).lower() for k in ("invalshoek", "hoek")) else 0)

            verrijk = st.checkbox("E-mails van de websites proberen op te halen (traag, werkt "
                                  "vaak niet in de cloud)", value=False,
                                  disabled=not _UPLOAD_OK)
            if st.button("Mails opbouwen", type="primary"):
                rijen = dfu.to_dict("records")
                uit = []
                prog = st.progress(0.0)
                sess = requests.Session()
                robots = {}
                for idx, rij in enumerate(rijen):
                    naam = str(rij.get(k_naam, "") or "").strip()
                    website = "" if k_web == geen else str(rij.get(k_web, "") or "").strip()
                    functie = "" if k_func == geen else str(rij.get(k_func, "") or "").strip()
                    inval = "" if k_inval == geen else str(rij.get(k_inval, "") or "").strip()
                    email = "" if k_mail == geen else str(rij.get(k_mail, "") or "").strip()
                    if not email and verrijk and _UPLOAD_OK and website:
                        try:
                            r = enrich_row(website, "", "outreach@voicestamp.nl", False, True, sess, robots)
                            email = r.email or ""
                        except Exception:
                            email = ""
                    if not email and website:
                        dom = website.replace("https://", "").replace("http://", "").strip("/").split("/")[0]
                        dom = dom[4:] if dom.startswith("www.") else dom
                        email = f"info@{dom}" if dom else ""
                    subj, body = maak_hr_mail(naam, functie, inval, brand, links_blok)
                    uit.append({"Bedrijf": naam, "Functie": functie, "E-mail": email,
                                "Website": website, "Onderwerp": subj, "Mail": body})
                    prog.progress((idx + 1) / len(rijen))
                st.session_state["hr_uit"] = uit
                st.success(f"{len(uit)} mails opgebouwd.")

            uit = st.session_state.get("hr_uit")
            if uit:
                st.subheader("Resultaat")
                st.dataframe(pd.DataFrame([{k: v for k, v in u.items() if k != "Mail"} for u in uit]),
                             use_container_width=True, height=320, hide_index=True)
                # download
                import io as _io
                wb = Workbook(); ws = wb.active; ws.title = "HR-leads"
                kolommen = ["Bedrijf", "Functie", "E-mail", "Website", "Onderwerp", "Mail"]
                ws.append(kolommen)
                for u in uit:
                    ws.append([u.get(k, "") for k in kolommen])
                buf = _io.BytesIO(); wb.save(buf)
                st.download_button("⬇︎ Download Excel", buf.getvalue(),
                                   file_name="hr_leads.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                # per-lead bekijken + mailen
                st.subheader("Mail per lead")
                mail_met = [u for u in uit if u["E-mail"]]
                if mail_met:
                    ki = st.selectbox("Kies een lead", range(len(mail_met)),
                                      format_func=lambda i: f"{mail_met[i]['Bedrijf']} ({mail_met[i]['E-mail']})",
                                      key="hr_pick")
                    u = mail_met[ki]
                    ond = st.text_input("Onderwerp", u["Onderwerp"], key=f"hr_subj_{ki}")
                    txt = st.text_area("Mailtekst", u["Mail"], height=320, key=f"hr_body_{ki}")
                    if st.session_state.get("sb_user") and st.session_state.get("zoho_user"):
                        gedaan = crm_sent_today(st.session_state["sb_client"], st.session_state["sb_user"])
                        rest = MAIL_DAG_LIMIET - gedaan
                        if st.button(f"✉️ Verstuur via Zoho ({max(rest,0)} over vandaag)", key="hr_send"):
                            if rest <= 0:
                                st.warning("Je dagelijkse 15 mails zijn verstuurd. Morgen weer.")
                            else:
                                body_send = txt + "\n" + st.session_state.get("zoho_sig", "")
                                try:
                                    zoho_send(st.session_state["zoho_host"], 465,
                                              st.session_state["zoho_user"], st.session_state["zoho_pw"],
                                              st.session_state.get("zoho_from", st.session_state["zoho_user"]),
                                              u["E-mail"], ond, body_send)
                                    st.success(f"Mail verstuurd naar {u['E-mail']}.")
                                except Exception as e:
                                    st.error(f"Versturen mislukt: {e}")
                    else:
                        st.caption("Log in bij **Account / CRM** en **Mailen (Zoho)** om direct te "
                                   "versturen. Anders kun je de mail kopiëren of downloaden.")

st.divider()
st.caption(
    "Zakelijk gebruik. Benader alleen relevante bedrijven met een duidelijke afmeldmogelijkheid, "
    "en houd je aan de AVG en de Telecommunicatiewet (B2B-uitzondering). De data komt van "
    "OpenStreetMap (ODbL) en van openbare bedrijfswebsites. Verstuur gedoseerd; de verzendbatches "
    "helpen daarbij. Deze tool zoekt geen persoonsgegevens op en is bedoeld als hulpmiddel, niet "
    "als vervanging van je eigen beoordeling per lead.")
