"""
VoiceStamp Lead Finder — Streamlit web-app
==========================================
Een klikbare webversie van de leadtool: kies campagne + provincie, klik Start,
en download de Excel en de Smartlead-CSV. Werkt in de browser, ook op je telefoon.

Lokaal draaien:
    pip install streamlit requests dnspython openpyxl pandas
    streamlit run app.py

Gratis online zetten: zie LEES_MIJ_streamlit.md
"""

import re, io, time, csv, json, hashlib
from urllib.parse import urlparse

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

# ------------------------------------------------------------------ constants
OVERPASS = ["https://overpass-api.de/api/interpreter",
            "https://overpass.kumi.systems/api/interpreter"]
UA = "VoiceStamp-LeadFinder/streamlit (zakelijk gebruik; contact via voicestamp.nl)"
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
BAD = ("example.", "sentry.", "wixpress.", ".png", ".jpg", ".gif", "@2x",
       "your-email", "email@", "name@", "domain.com")
PATHS = ["", "contact", "contact/", "over-ons", "over-ons/", "info", "info/",
         "reserveren", "reserveren/", "about", "about/"]
ROLE = ("info", "contact", "welkom", "boeking", "boekingen", "reserveringen",
        "reservering", "hallo", "mail", "receptie", "office", "sales")
KETEN = ("hotels", "resorts", "group", "vakantieparken", "landal", "roompot",
         "rcn", "huttopia", "fletcher", "ardoer", "europarcs", "oostappen")
PROVINCIES = ["Groningen", "Fryslân", "Drenthe", "Overijssel", "Flevoland",
              "Gelderland", "Utrecht", "Noord-Holland", "Zuid-Holland",
              "Zeeland", "Noord-Brabant", "Limburg"]

PRESETS = {
    "Verblijf (campings, hotels, B&B's)": {
        "key": "verblijf",
        "filters": [("tourism", "camp_site"), ("tourism", "caravan_site"),
                    ("tourism", "guest_house"), ("tourism", "hotel"),
                    ("tourism", "hostel"), ("tourism", "chalet"), ("tourism", "motel")],
        "typemap": {"camp_site": ("Camping", "natuurcamping"),
                    "caravan_site": ("Camperplaats", "natuurcamping"),
                    "guest_house": ("B&B / guest house", "bnb"),
                    "hotel": ("Hotel", "hotel"), "hostel": ("Hostel", "hotel"),
                    "chalet": ("Chalet/huisjes", "natuurcamping"), "motel": ("Motel", "hotel")}},
    "Streek (wijn, kaas, boerderijwinkels)": {
        "key": "streek",
        "filters": [("craft", "winery"), ("shop", "wine"), ("shop", "cheese"),
                    ("shop", "farm"), ("shop", "dairy"), ("shop", "greengrocer")],
        "typemap": {"winery": ("Wijngaard", "streek"), "wine": ("Wijnhandel", "streek"),
                    "cheese": ("Kaaswinkel", "streek"), "farm": ("Boerderijwinkel", "streek"),
                    "dairy": ("Zuivel/boerderij", "streek"), "greengrocer": ("Streekwinkel", "streek")}},
    "Attracties (dierentuinen, kinderboerderijen)": {
        "key": "attracties",
        "filters": [("tourism", "zoo"), ("tourism", "theme_park"), ("tourism", "aquarium")],
        "typemap": {"zoo": ("Dierentuin/kinderboerderij", "attractie"),
                    "theme_park": ("Attractiepark", "attractie"),
                    "aquarium": ("Aquarium", "attractie")}},
    "Erfgoed (musea, kastelen, landgoederen)": {
        "key": "erfgoed",
        "filters": [("tourism", "museum"), ("historic", "castle"), ("historic", "manor")],
        "typemap": {"museum": ("Museum", "erfgoed"), "castle": ("Kasteel", "erfgoed"),
                    "manor": ("Landgoed", "erfgoed")}},
}

LINKS = ("\u2022 Instagram: https://www.instagram.com/voicestamp.nl / @voicestamp.nl\n"
         "\u2022 LinkedIn: https://www.linkedin.com/company/voicestamp / VoiceStamp\n"
         "\u2022 Website: www.voicestamp.nl\n\u2022 Hoe het werkt: www.voicestamp.nl/how-it-works")
SIG = ("Met vriendelijke groet,\nYusuf Tatlicioglu\n0646756497\n\n"
       "Founder VoiceStamp\nhttps://www.voicestamp.nl/")

def _body(mid):
    return ("Beste {aanhef},\n\n{opener} Dat bracht me op een idee.\n\n" + mid +
            "\n\nBenieuwd hoe dat eruitziet? Neem gerust een kijkje:\n" + LINKS +
            "\n\nLijkt het interessant? Dan lees ik graag een reactie!\n\n" + SIG)

MID = {
 "natuurcamping": ("Wat als een gast dat verhaal niet alleen leest, maar het ook rechtstreeks van "
   "jullie hoort?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) \u2014 op een "
   "kampeerplek, bij de receptie of aan de start van een route \u2014 en opent direct een "
   "persoonlijke audioboodschap met een landingspagina. Geen app, geen account: juist minder drukte, "
   "niet meer. Bijvoorbeeld een warm welkom, een verhaal over de omgeving, of een wandeltip."),
 "bnb": ("Wat als een gast dat verhaal niet alleen leest, maar het ook van jullie hoort \u2014 op het "
   "moment dat hij binnenkomt?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op de "
   "kamer, en opent direct een persoonlijke audioboodschap met een landingspagina. Geen app, geen "
   "account. Een warm welkom in je eigen stem, het verhaal van het pand, of een tip voor de omgeving."),
 "hotel": ("Wat als een gast dat verhaal hoort op het moment dat hij binnenkomt?\n\nMet VoiceStamp "
   "scan je een eenvoudige stempel (VoiceStamp) op de kamer of in de lobby, en opent direct een "
   "audioboodschap met een landingspagina. Geen app, geen account. Een warm welkom, het verhaal van "
   "de plek, of info die nu aan de balie wordt gevraagd."),
 "streek": ("Want dat verhaal vertel je in de winkel, maar het gaat niet mee met het product dat "
   "iemand koopt.\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) op de verpakking, en "
   "opent direct een audioboodschap met een landingspagina. Geen app, geen account. Waarom een "
   "ingredi\u00ebnt is gekozen, hoe iets gemaakt wordt, of een welkom bij een proeverij."),
 "attractie": ("Wat als een bezoeker bij een dier niet alleen een bordje leest, maar een verzorger "
   "het hoort vertellen?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) bij een "
   "verblijf, en opent direct een audioboodschap met een landingspagina. Geen app \u2014 en het werkt "
   "ook voor kinderen die nog niet lezen. Het verhaal van een dier, een welkom, of het dagprogramma."),
 "erfgoed": ("Wat als een bezoeker het verhaal niet alleen leest, maar hoort \u2014 in de stem van een "
   "gids of conservator?\n\nMet VoiceStamp scan je een eenvoudige stempel (VoiceStamp) bij een object "
   "of monument, en opent direct een audioboodschap met een landingspagina. Geen app, geen account. "
   "Zo maak je het verhaal toegankelijk, ook buiten de rondleidingen om."),
}
SUBJECTS = {
 "natuurcamping": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Een stem bij jullie plek, zonder app"],
 "bnb": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Een persoonlijk welkom op de kamer"],
 "hotel": ["Wat als jullie plek zelf haar verhaal kon vertellen?", "Minder vragen aan de balie"],
 "streek": ["Wat als jullie product zelf zijn verhaal kon vertellen?", "Het verhaal achter het product"],
 "attractie": ["Wat als jullie dieren hun eigen verhaal konden vertellen?", "Een stem bij elk verblijf"],
 "erfgoed": ["Wat als jullie collectie zelf haar verhaal kon vertellen?", "Een stem bij het monument"],
}
DEFAULT_OPENER = {
 "natuurcamping": "Tijdens het bekijken van jullie website viel me op hoeveel aandacht jullie besteden aan rust en natuur.",
 "bnb": "Tijdens het bekijken van jullie website viel me op hoe persoonlijk jullie gasten ontvangen.",
 "hotel": "Tijdens het bekijken van jullie website viel me op hoeveel karakter jullie plek heeft.",
 "streek": "Tijdens het bekijken van jullie website viel me op hoeveel verhaal er in jullie producten zit.",
 "attractie": "Tijdens het bekijken van jullie website viel me op hoeveel er bij jullie te beleven is.",
 "erfgoed": "Tijdens het bekijken van jullie website viel me op hoeveel verhaal jullie plek draagt.",
}
KEYWORDS = ("rust", "natuur", "gastvrij", "persoonlijk", "familie", "verhaal", "duurzaam",
            "welkom", "beleef", "genieten", "monument", "ambacht", "streek", "traditie",
            "avontuur", "kinderen", "dieren")

# ------------------------------------------------------------------ core logic
def q_build(area, filters, whole):
    tf = "".join(f'node["{k}"="{v}"](area.a);way["{k}"="{v}"](area.a);'
                 f'relation["{k}"="{v}"](area.a);' for k, v in filters)
    ab = ('area["ISO3166-1"="NL"][admin_level=2]->.a;' if whole
          else f'area["name"="{area}"]["admin_level"="4"]->.a;')
    return f"[out:json][timeout:180];{ab}({tf});out center tags;"

def osm(area, filters, whole):
    q = q_build(area, filters, whole); err = None
    for ep in OVERPASS:
        try:
            r = requests.post(ep, data={"data": q}, headers={"User-Agent": UA}, timeout=200)
            r.raise_for_status(); return r.json().get("elements", [])
        except Exception as e:
            err = e; time.sleep(2)
    raise RuntimeError(f"Overpass niet bereikbaar: {err}")

def parse(els, typemap):
    out = []
    for el in els:
        t = el.get("tags", {}); name = t.get("name")
        if not name: continue
        osmval = next((t.get(k) for k in ("tourism", "craft", "shop", "historic")
                       if t.get(k) in typemap), None)
        if not osmval: continue
        label, seg = typemap[osmval]
        email = (t.get("contact:email") or t.get("email") or "").strip().lower()
        out.append({"naam": name, "type": label, "seg": seg,
                    "plaats": (t.get("addr:city") or t.get("addr:place") or "").strip(),
                    "email": email if EMAIL_RE.fullmatch(email or "") else "",
                    "website": (t.get("contact:website") or t.get("website") or "").strip(),
                    "telefoon": (t.get("contact:phone") or t.get("phone") or "").strip(),
                    "bron": "OpenStreetMap", "opener": ""})
    return out

def norm(u):
    if not u: return ""
    if not u.startswith(("http://", "https://")): u = "https://" + u
    return u.rstrip("/")

def detag(h):
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
    return re.sub(r"<[^>]+>", " ", h)

def opener_uit(text, seg, plaats):
    text = re.sub(r"\s+", " ", text).strip(); best, score = None, 0
    for z in re.split(r"(?<=[.!?]) ", text):
        zl = z.lower()
        if not (40 <= len(z) <= 165) or "cookie" in zl: continue
        s = sum(k in zl for k in KEYWORDS)
        if s > score: best, score = z.strip().rstrip("."), s
    if best and score >= 1:
        return f"Tijdens het bekijken van jullie website viel me op dat {best[0].lower()}{best[1:]}."
    if plaats: return DEFAULT_OPENER[seg].rstrip(".") + f", hier in {plaats}."
    return DEFAULT_OPENER[seg]

def enrich(website, seg, plaats):
    base = norm(website)
    if not base: return "", ""
    dom = urlparse(base).netloc.replace("www.", ""); email, opener = "", ""
    for p in PATHS:
        url = base if p == "" else f"{base}/{p}"
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=10)
            if r.status_code != 200 or "html" not in r.headers.get("content-type", ""): continue
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
            if not opener and p in ("", "over-ons", "over-ons/", "about", "about/"):
                opener = opener_uit(detag(html), seg, plaats)
            if email and opener: break
        except Exception: continue
        finally: time.sleep(0.6)
    return email, (opener or (DEFAULT_OPENER[seg] if not plaats else
                   DEFAULT_OPENER[seg].rstrip(".") + f", hier in {plaats}."))

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

def maak_mail(l):
    seg = l["seg"]; subs = SUBJECTS[seg]
    subject = subs[int(hashlib.md5(l['naam'].encode()).hexdigest(), 16) % len(subs)]
    aanhef = l.get("_vn") or f"team van {l['naam']}"
    body = _body(MID[seg]).format(aanhef=aanhef, opener=l.get("opener") or DEFAULT_OPENER[seg])
    return subject, body

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
            for m in EMAIL_RE.findall(bytes_data.decode("utf-8", "ignore")):
                s.add(m.strip().lower())
    except Exception:
        pass
    return s

def build_xlsx(leads):
    wb = Workbook(); ws = wb.active; ws.title = "Leads"
    H = ["Naam", "Type", "Plaats", "E-mail", "Voornaam", "Beste ingang", "E-mail geldig?",
         "Verzendbatch", "Website", "Telefoon", "Concept onderwerp", "Concept mail",
         "Status", "Datum verstuurd", "Bron"]
    hf = PatternFill("solid", fgColor="2E5D4B"); hfont = Font("Arial", bold=True, color="FFFFFF")
    th = Side(style="thin", color="D0D0D0"); bd = Border(th, th, th, th)
    for c, h in enumerate(H, 1):
        x = ws.cell(1, c, h); x.fill = hf; x.font = hfont
        x.alignment = Alignment(vertical="center", wrap_text=True); x.border = bd
    for r, l in enumerate(leads, 2):
        mx = l.get("mx"); mxt = "ja" if mx is True else ("nee" if mx is False else "onbekend")
        vals = [l["naam"], l["type"], l["plaats"], l["email"], l.get("_vn", ""), l["_ingang"], mxt,
                l.get("_batch", ""), l["website"], l["telefoon"], l["_subject"], l["_body"], "", "", l["bron"]]
        for c, v in enumerate(vals, 1):
            x = ws.cell(r, c, v); x.font = Font("Arial", size=10)
            x.alignment = Alignment(vertical="top", wrap_text=True); x.border = bd
    for i, w in enumerate([26, 18, 15, 28, 12, 26, 11, 11, 26, 15, 32, 66, 12, 13, 15], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 28
    ws.auto_filter.ref = f"A1:{get_column_letter(len(H))}{len(leads)+1}"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_csv(leads):
    cols = ["email", "first_name", "company", "city", "custom_opener", "custom_subject",
            "custom_body", "send_batch", "website", "phone", "beste_ingang"]
    buf = io.StringIO(); w = csv.DictWriter(buf, fieldnames=cols); w.writeheader()
    for l in leads:
        if not l["email"]: continue
        w.writerow({"email": l["email"], "first_name": l.get("_vn", ""), "company": l["naam"],
                    "city": l["plaats"], "custom_opener": l.get("opener", ""),
                    "custom_subject": l["_subject"], "custom_body": l["_body"],
                    "send_batch": l.get("_batch", ""), "website": l["website"],
                    "phone": l["telefoon"], "beste_ingang": l["_ingang"]})
    return buf.getvalue().encode("utf-8")

# ------------------------------------------------------------------ pipeline
def draai(preset, area, whole, scrape, mx_check, dedupe_dom, n_per_dag, max_sites,
          verstuurd, progress, status):
    status("OpenStreetMap ophalen ...")
    leads = parse(osm(area, preset["filters"], whole), preset["typemap"])
    seen, uniek = set(), []
    for l in leads:
        k = (l["naam"].lower(), l["plaats"].lower())
        if k not in seen: seen.add(k); uniek.append(l)
    leads = uniek
    if verstuurd:
        leads = [l for l in leads if l["email"].lower() not in verstuurd]
    progress(0.15)

    if scrape:
        todo = [l for l in leads if l["website"] and (not l["email"] or not l["opener"])]
        if max_sites: todo = todo[:max_sites]
        for i, l in enumerate(todo, 1):
            e, o = enrich(l["website"], l["seg"], l["plaats"])
            if e and not l["email"]: l["email"] = e; l["bron"] = "OSM + website"
            l["opener"] = o
            if i % 3 == 0 or i == len(todo):
                status(f"Websites verrijken … {i}/{len(todo)}")
                progress(0.15 + 0.6 * i / max(len(todo), 1))
    progress(0.8)

    if dedupe_dom:
        gz, dd = set(), []
        for l in leads:
            dom = (l["email"].split("@")[-1] if l["email"] else l["website"]).lower()
            if dom and dom in gz: continue
            gz.add(dom); dd.append(l)
        leads = dd

    for l in leads:
        l["_vn"] = voornaam(l["email"]); l["_ingang"] = ingang(l)
        l["_subject"], l["_body"] = maak_mail(l)

    if mx_check and HAVE_DNS:
        status("E-mailadressen controleren (MX) ...")
        for l in leads: l["mx"] = has_mx(l["email"]) if l["email"] else None
    else:
        for l in leads: l["mx"] = None
    progress(0.92)

    leads.sort(key=lambda l: (l["email"] == "", l["type"], l["naam"]))
    t = 0
    for l in leads:
        if l["email"] and l.get("mx") is not False:
            l["_batch"] = f"Dag {t // max(n_per_dag, 1) + 1}"; t += 1
        else: l["_batch"] = ""
    progress(1.0)
    return leads

# ------------------------------------------------------------------ UI
st.set_page_config(page_title="VoiceStamp Lead Finder", page_icon="🎙️", layout="wide")
st.title("🎙️ VoiceStamp Lead Finder")
st.caption("Gratis leads uit OpenStreetMap, verrijkt met e-mail, concept-mail en verzendbatches.")

with st.sidebar:
    st.header("Instellingen")
    campagne = st.selectbox("Campagne", list(PRESETS.keys()))
    heel = st.checkbox("Heel Nederland", value=False)
    provincie = st.selectbox("Provincie", PROVINCIES, index=3, disabled=heel)
    st.divider()
    scrape = st.checkbox("Websites scrapen (e-mail + openingszin)", value=True)
    mx_check = st.checkbox("E-mail controleren (MX)", value=True, disabled=not HAVE_DNS)
    dedupe_dom = st.checkbox("Max 1 per domein/keten", value=False)
    n_per_dag = st.number_input("Aantal per verzenddag", 5, 200, 25, step=5)
    max_sites = st.number_input("Max sites scrapen (0 = alle)", 0, 5000, 0, step=10)
    st.divider()
    up = st.file_uploader("Al-verstuurde lijst (csv/xlsx, optioneel)", type=["csv", "xlsx", "xlsm"])
    start = st.button("🚀 Start", type="primary", use_container_width=True)

if start:
    preset = PRESETS[campagne]
    verstuurd = laad_verstuurd(up.getvalue(), up.name) if up else set()
    bar = st.progress(0.0); status_box = st.empty()
    try:
        leads = draai(preset, None if heel else provincie, heel, scrape, mx_check,
                      dedupe_dom, n_per_dag, max_sites, verstuurd,
                      lambda p: bar.progress(min(p, 1.0)),
                      lambda m: status_box.info(m))
        st.session_state["leads"] = leads
        status_box.success("Klaar!")
    except Exception as e:
        status_box.error(f"Er ging iets mis: {e}")

if st.session_state.get("leads"):
    leads = st.session_state["leads"]
    met = sum(1 for l in leads if l["email"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Leads totaal", len(leads))
    c2.metric("Met e-mail", met)
    c3.metric("Verzenddagen", (met // max(int(n_per_dag), 1)) + 1 if met else 0)

    df = pd.DataFrame([{"Naam": l["naam"], "Type": l["type"], "Plaats": l["plaats"],
                        "E-mail": l["email"], "Ingang": l["_ingang"], "Batch": l.get("_batch", "")}
                       for l in leads])
    st.dataframe(df, use_container_width=True, height=380)

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Excel (met concept-mails)", build_xlsx(leads),
                       file_name="voicestamp_leads.xlsx", use_container_width=True,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇️ CSV voor Smartlead/Instantly", build_csv(leads),
                       file_name="voicestamp_leads.csv", mime="text/csv", use_container_width=True)
else:
    st.info("Kies links een campagne en provincie en klik op **Start**.")
