"""Scoring van leads, met uitleg. Geen Streamlit nodig, dus los te testen.

De score is een som van punten. Elke regel in de uitleg is een reden met zijn punten, zodat je kunt zien
waarom een lead hoog of laag staat. Alle gewichten staan hieronder en zijn gemakkelijk aan te passen.
"""
import re
from datetime import date

# Hoe goed past een branche bij VoiceStamp? Mijn inschatting, geen meting. Pas gerust aan.
BRANCHE_GEWICHT = {
    "winkel": 8, "streek": 8, "natuurcamping": 6, "bnb": 6, "bloemist": 5, "erfgoed": 5,
    "wellness": 5, "dealer": 5, "horeca": 4, "entertainment": 3, "recreatie": 3, "hotel": 3,
    "attractie": 2, "zorg": 1, "vastgoed": 1, "werkgever": 0, "onderwijs": 0, "evenement": 0,
    "vervoer": -5,
}

TECH = {
    "Shopify": r"cdn\.shopify\.com|myshopify\.com",
    "WooCommerce": r"woocommerce",
    "Lightspeed": r"lightspeedhq|webshopapp\.com|seoshop\.nl",
    "Magento": r"magento|/static/version\d+/frontend",
    "Wix": r"wixstatic\.com|static\.parastorage\.com",
    "Squarespace": r"squarespace\.com|static1\.squarespace",
    "WordPress": r"wp-content|wp-includes",
}
BOEKING = {
    "Recras": r"recras",
    "Bookingkit": r"bookingkit",
    "Checkfront": r"checkfront",
    "Beds24": r"beds24",
    "Eviivo": r"eviivo",
    "Cloudbeds": r"cloudbeds",
    "SimpleBooking": r"simplebooking",
    "Tablebooker": r"tablebooker",
    "Formitable": r"formitable",
    "TheFork": r"thefork\.",
    "Resengo": r"resengo",
}
WEBSHOP_TEKST = re.compile(r"winkelwagen|winkelmand|in mijn mandje|add to cart|afrekenen|bestel online|webshop", re.I)
VERHAAL = re.compile(
    r"familiebedrijf|sinds (?:19|20)\d\d|ambacht|handgemaakt|zelfgemaakt|met de hand|eigen (?:productie|bakkerij|"
    r"brouwerij|atelier|werkplaats|recept)|vakmanschap|streekproduct|lokaal|duurzaam", re.I)
TEAM = re.compile(r"(\d{1,3})\s*(?:medewerkers|collega'?s|teamleden|man personeel|vaste medewerkers)", re.I)
TEAM2 = re.compile(r"team van (\d{1,3})", re.I)
VESTIGINGEN = re.compile(r"(\d{1,3})\s*(?:vestigingen|filialen|winkels|locaties|kantoren)", re.I)
JAAR_NABIJ = re.compile(r"(?:©|&copy;|copyright|alle rechten voorbehouden)[^<]{0,60}?((?:20)\d\d)", re.I)
JAAR_META = re.compile(r"(?:dateModified|datePublished|article:modified_time|article:published_time)[^0-9]{0,12}((?:20)\d\d)-\d\d-\d\d")
WHATSAPP = re.compile(r"(?:https?://)?(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\d{8,15})")


def site_signalen(paginas):
    """paginas = [(url, html)]. Geeft een dict met wat we gratis uit de site kunnen halen."""
    alle = " ".join(h for _, h in paginas)[:600_000]
    tekst = re.sub(r"<[^>]+>", " ", alle)
    sig = {"tech": [], "boeking": [], "webshop": False, "laatste_jaar": None, "team": None,
           "vestigingen": None, "verhaalwoorden": 0, "formulier_url": "", "whatsapp": ""}
    low = alle.lower()
    sig["tech"] = [n for n, rx in TECH.items() if re.search(rx, low)]
    sig["boeking"] = [n for n, rx in BOEKING.items() if re.search(rx, low)]
    sig["webshop"] = bool({"Shopify", "WooCommerce", "Lightspeed", "Magento"} & set(sig["tech"])) \
        or len(WEBSHOP_TEKST.findall(tekst)) >= 2
    jaren = [int(j) for j in JAAR_NABIJ.findall(alle)] + [int(j) for j in JAAR_META.findall(alle)]
    jaren = [j for j in jaren if 2000 <= j <= date.today().year + 1]
    sig["laatste_jaar"] = max(jaren) if jaren else None
    m = TEAM.search(tekst) or TEAM2.search(tekst)
    sig["team"] = int(m.group(1)) if m else None
    v = [int(x) for x in VESTIGINGEN.findall(tekst)]
    sig["vestigingen"] = max(v) if v else None
    sig["verhaalwoorden"] = len(set(w.lower() for w in VERHAAL.findall(tekst)))
    for url, html in paginas:
        if re.search(r"<form[^>]*>.*?<textarea", html, re.I | re.S):
            sig["formulier_url"] = url
            break
    m = WHATSAPP.search(alle)
    if m:
        sig["whatsapp"] = m.group(1)
    return sig


def score(l, huidig_jaar=None):
    """Geeft (score 0-100, [(reden, punten), ...]). `l` is een lead-dict uit de app."""
    huidig_jaar = huidig_jaar or date.today().year
    sig = l.get("_sig") or {}
    r = []
    if l.get("email"):
        r.append(("E-mailadres gevonden", 30))
        if l.get("mx") is True:
            r.append(("Domein kan mail ontvangen", 5))
    ing = l.get("_ingang", "")
    if "Persoonlijk" in ing:
        r.append(("Persoonlijk adres", 15))
    if "klein" in ing:
        r.append(("Klein bedrijf, snelle beslissing", 10))
    if "keten" in ing:
        r.append(("Keten, trage beslissing", -15))
    if l.get("haak"):
        r.append((f"Haak op de site: {l['haak']}", 10))
    if l.get("instagram") or l.get("linkedin"):
        r.append(("Actief op social", 5))

    g = BRANCHE_GEWICHT.get(l.get("seg", ""), 0)
    if g:
        r.append((f"Past {'goed' if g > 0 else 'minder'} bij VoiceStamp (branche)", g))
    if sig.get("webshop"):
        r.append(("Heeft een webshop", 8))
    if sig.get("boeking"):
        r.append((f"Werkt met online boeken ({sig['boeking'][0]})", 3))
    jr = sig.get("laatste_jaar")
    if jr:
        if jr >= huidig_jaar - 1:
            r.append((f"Site bijgewerkt in {jr}", 7))
        elif jr <= huidig_jaar - 3:
            r.append((f"Site oogt verlaten (laatste jaar {jr})", -5))
    if sig.get("verhaalwoorden", 0) >= 2:
        r.append(("Vertelt een eigen verhaal (ambacht, familie, streek)", 5))
    t = sig.get("team")
    if t is not None and t <= 15:
        r.append((f"Klein team ({t})", 4))
    v = sig.get("vestigingen")
    if v and v >= 5:
        r.append((f"{v} vestigingen, waarschijnlijk een keten", -15))
    if l.get("heeft_app"):
        r.append(("Heeft al een eigen app", -5))
    totaal = max(0, min(100, sum(p for _, p in r)))
    return totaal, r


def uitleg_tekst(redenen):
    return "; ".join(f"{naam} {p:+d}" for naam, p in redenen)
