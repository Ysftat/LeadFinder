"""Kanaalkaart: waar bereik je een lead, en wat zeg je daar? Geen Streamlit nodig.

De app stuurt hier niets zelf. Ze maakt de links en de voorbeeldberichten, jij klikt en verstuurt.
Dat is bewust: automatisch berichten sturen op LinkedIn of Instagram kan je account kosten.
"""
import re
from urllib.parse import quote_plus

DING = {
    "natuurcamping": "jullie camping", "bnb": "jullie B&B", "hotel": "jullie hotel", "streek": "jullie producten",
    "horeca": "jullie zaak", "winkel": "jullie winkel", "bloemist": "jullie winkel", "attractie": "jullie park",
    "erfgoed": "jullie locatie", "zorg": "jullie praktijk", "wellness": "jullie plek", "vastgoed": "jullie woningen",
    "werkgever": "jullie onboarding", "evenement": "jullie locatie", "entertainment": "jullie zaal",
    "onderwijs": "jullie school", "dealer": "jullie voertuigen", "vervoer": "jullie reizigers",
    "recreatie": "jullie plek",
}


def _tekst_naam(naam):
    return (naam or "").strip()


def linkedin_personen_link(bedrijf, persoon=""):
    q = " ".join(x for x in (persoon, bedrijf) if x)
    return "https://www.linkedin.com/search/results/people/?keywords=" + quote_plus(q)


def linkedin_bedrijf_link(bedrijf):
    return "https://www.linkedin.com/search/results/companies/?keywords=" + quote_plus(bedrijf or "")


def instagram_link(waarde):
    w = (waarde or "").strip()
    if not w:
        return ""
    if w.startswith("http"):
        return w
    return "https://www.instagram.com/" + w.lstrip("@").strip("/") + "/"


def nl_nummer(tel):
    """Maakt van 06-12345678 of +31 6 12345678 het nummer zonder plus (31612345678), of ''."""
    d = re.sub(r"\D", "", tel or "")
    if d.startswith("0031"):
        d = d[2:]
    if d.startswith("31") and len(d) >= 11:
        return d
    if d.startswith("0") and len(d) >= 10:
        return "31" + d[1:]
    return ""


def whatsapp_link(nummer, tekst=""):
    n = re.sub(r"\D", "", nummer or "")
    if not n:
        return ""
    return f"https://wa.me/{n}" + (f"?text={quote_plus(tekst)}" if tekst else "")


def beste_kanalen(lead):
    """Volgorde van meest naar minst kansrijk: [(kanaal, uitleg, link)]."""
    seg = lead.get("segment") or lead.get("seg") or ""
    naam = _tekst_naam(lead.get("company_name") or lead.get("naam"))
    uit = []
    if (lead.get("email") or "").strip() or (lead.get("contact_email") or "").strip():
        uit.append(("E-mail", "Heeft een adres: zet de mailreeks klaar.", ""))
    if (lead.get("instagram") or "").strip():
        uit.append(("Instagram", "Heeft een Instagram-account: stuur een kort bericht.",
                    instagram_link(lead["instagram"])))
    wa = nl_nummer(lead.get("whatsapp") or "") or nl_nummer(lead.get("phone") or lead.get("telefoon") or "")
    if (lead.get("whatsapp") or "").strip():
        uit.append(("WhatsApp", "Er staat een WhatsApp-link op de site.", whatsapp_link(wa)))
    if (lead.get("form_url") or "").strip():
        uit.append(("Contactformulier", "De site heeft een contactformulier.", lead["form_url"]))
    persoon = lead.get("contact_name") or ""
    uit.append(("LinkedIn", "Zoek de eigenaar of manager en stuur een korte uitnodiging.",
                linkedin_personen_link(naam, persoon)))
    if (lead.get("phone") or lead.get("telefoon") or "").strip():
        uit.append(("Telefoon", "Bellen kan altijd, vooral bij kleine zaken.", ""))
    return uit


def maak_berichten(lead, afzender="Yusuf", merk="VoiceStamp", site="https://www.voicestamp.nl"):
    """Korte voorbeeldberichten per kanaal. Jij leest ze en verstuurt ze zelf."""
    seg = lead.get("segment_seg") or lead.get("seg") or ""
    ding = DING.get(seg, "jullie producten")
    voornaam = (lead.get("contact_name") or "").split()[0] if lead.get("contact_name") else ""
    hoi = f"Hoi {voornaam}," if voornaam else "Hoi,"
    pitch = (f"Wat als {ding} niet alleen gezien wordt, maar ook iets kan vertellen? Met {merk} scant iemand een "
             "eenvoudige stempel en hoort direct jullie boodschap. Geen app, geen account.")
    haak = (lead.get("hook") or lead.get("haak") or "").strip()
    haakzin = f"Ik zag dat jullie {haak.lower()}. " if haak and len(haak) < 60 else ""
    linkedin_noot = f"{hoi} {haakzin}Ik maak {merk}: scan een stempel, hoor direct de stem erachter. Geen app. Zin om kennis te maken? {afzender}"
    linkedin_bericht = (f"{hoi}\n\n{haakzin}{pitch}\n\nIk laat het graag zien als je denkt dat het bij jullie past. "
                        f"Meer info: {site}\n\nGroet,\n{afzender}")
    instagram_dm = (f"{hoi} {haakzin}{pitch} Benieuwd? Kijk op {site.replace('https://', '')} "
                    f"of stuur me een bericht. {afzender}")
    formulier = (f"{hoi}\n\n{pitch}\n\nIk laat het graag zien. Kunnen jullie me doorverwijzen naar de juiste persoon?\n\n"
                 f"Meer info: {site}\n\nMet vriendelijke groet,\n{afzender}")
    whatsapp = (f"{hoi} {pitch} Mag ik je er een kort voorbeeld van sturen? {afzender} ({merk})")
    # LinkedIn-notitie kort houden
    if len(linkedin_noot) > 200:
        linkedin_noot = f"{hoi} Ik maak {merk}: scan een stempel, hoor direct de stem erachter. Geen app. Zin om kennis te maken? {afzender}"
    return {"LinkedIn uitnodiging (kort)": linkedin_noot, "LinkedIn bericht": linkedin_bericht,
            "Instagram bericht": instagram_dm, "Contactformulier": formulier, "WhatsApp": whatsapp}
