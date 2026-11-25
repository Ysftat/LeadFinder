"""Contentwachtrij voor LinkedIn en Instagram. Geen Streamlit nodig.

De app plaatst niets zelf. Ze maakt een planning met bijschriften, jij plakt ze in LinkedIn of Instagram
(of in een gratis planner). Teksten komen uit sjablonen, niet uit een taalmodel: lees ze vooraf.
"""
from datetime import date, datetime, timedelta

SITE = "https://www.voicestamp.nl"
LINKS = (f"• Website: {SITE}\n• Hoe het werkt: {SITE}/how-it-works\n"
         "• Instagram: https://www.instagram.com/voicestamp.nl\n"
         "• LinkedIn: https://www.linkedin.com/company/voicestamp")

# Per sector: wie, waar kan een stempel staan, een inzicht, en hashtags.
SECTOREN = {
    "Webshops": {
        "wie": "een webshop", "klant": "klanten", "ding": "pakketje",
        "plekken": ["in het pakketje bij elke bestelling", "op de verpakking", "op de bedankkaart"],
        "inzicht": "Een pakketje komt aan. Het product is mooi. Maar het verhaal erachter zit nergens in de doos.",
        "voorbeeld": "De maker spreekt een persoonlijke bedankboodschap in. De klant scant de stempel in het pakketje "
                     "en hoort wie het maakte en hoe je het het beste gebruikt.",
        "tags": "#webshop #ondernemen #klantbeleving #unboxing #VoiceStamp",
    },
    "Campings & vakantieparken": {
        "wie": "een camping", "klant": "gasten", "ding": "plek",
        "plekken": ["bij de ingang", "op de kaart met wandelroutes", "bij de sanitairgebouwen"],
        "inzicht": "Op een bord bij de ingang lezen gasten de huisregels. Zelden de tip die het verblijf echt bijzonder maakt.",
        "voorbeeld": "De eigenaar vertelt bij de ingang welkom en wijst de mooiste wandeling aan, in eigen stem.",
        "tags": "#camping #vakantiepark #gastvrijheid #recreatie #VoiceStamp",
    },
    "Hotels & B&B's": {
        "wie": "een hotel of B&B", "klant": "gasten", "ding": "kamer",
        "plekken": ["op het nachtkastje", "bij de receptie", "bij het ontbijt"],
        "inzicht": "Een welkomstmapje wordt zelden helemaal gelezen. Een stem wel.",
        "voorbeeld": "De gastheer vertelt op de kamer over de buurt, het ontbijt en zijn favoriete restaurant.",
        "tags": "#hotel #bnb #gastvrijheid #hospitality #VoiceStamp",
    },
    "Restaurants & cafés": {
        "wie": "een restaurant of café", "klant": "gasten", "ding": "menukaart",
        "plekken": ["op het menu", "op de tafel", "bij de kassa"],
        "inzicht": "Het verhaal achter een gerecht past niet op een menukaart. Het past wel in een stem.",
        "voorbeeld": "De kok vertelt bij het menu waar de ingrediënten vandaan komen en wat hij zelf bestelt.",
        "tags": "#horeca #restaurant #koffiebar #lokaal #VoiceStamp",
    },
    "Streekproducten & ambacht": {
        "wie": "een maker van streekproducten", "klant": "klanten", "ding": "product",
        "plekken": ["op het etiket", "op de verpakking", "bij het schap"],
        "inzicht": "Een ambachtelijk product heeft een verhaal. In de winkel krijgt het zelden de tijd om verteld te worden.",
        "voorbeeld": "De maker vertelt op de verpakking hoe het gemaakt is en hoe je het het lekkerst eet.",
        "tags": "#streekproducten #ambacht #lokaal #foodie #VoiceStamp",
    },
    "Winkels & retail": {
        "wie": "een winkel", "klant": "klanten", "ding": "product",
        "plekken": ["bij een schap", "op een product", "bij de kassa"],
        "inzicht": "In de winkel kun je een product zien en voelen. Wie het maakte en waarom, weet je niet.",
        "voorbeeld": "Bij het schap vertelt de eigenaar of maker waarom dit product in de winkel ligt.",
        "tags": "#retail #winkel #lokaalwinkelen #klantbeleving #VoiceStamp",
    },
    "Musea & erfgoed": {
        "wie": "een museum of erfgoedlocatie", "klant": "bezoekers", "ding": "tentoonstelling",
        "plekken": ["bij een object", "bij de ingang", "langs de route"],
        "inzicht": "Bij een object past een kaartje van drie regels. Er is vaak veel meer te vertellen.",
        "voorbeeld": "De conservator of de maker vertelt bij een object het verhaal, zonder audiotour en zonder app.",
        "tags": "#museum #erfgoed #cultuur #bezoekerservaring #VoiceStamp",
    },
    "Zorg & welzijn": {
        "wie": "een praktijk", "klant": "patiënten", "ding": "wachtkamer",
        "plekken": ["in de wachtkamer", "bij een behandelruimte", "op de folder die mee naar huis gaat"],
        "inzicht": "Dezelfde uitleg geef je elke dag opnieuw. Een stem kan het één keer goed vertellen.",
        "voorbeeld": "De therapeut spreekt de uitleg bij een oefening in, zodat patiënten die thuis nog eens kunnen horen.",
        "tags": "#zorg #welzijn #patiëntervaring #fysiotherapie #VoiceStamp",
    },
    "Wellness & yoga": {
        "wie": "een wellnesszaak of studio", "klant": "gasten", "ding": "studio",
        "plekken": ["bij de ingang", "in de kleedkamer", "bij de studio"],
        "inzicht": "Rust is de reden dat gasten komen. Een scherm of een bord past daar niet.",
        "voorbeeld": "Een korte ademoefening in de stem van de docent, te scannen bij de mat.",
        "tags": "#wellness #yoga #spa #rust #VoiceStamp",
    },
    "Vastgoed": {
        "wie": "een makelaar of woningbeheerder", "klant": "bewoners", "ding": "woning",
        "plekken": ["bij de voordeur", "bij de meterkast", "in de gezamenlijke ruimte"],
        "inzicht": "Bij oplevering krijgt een bewoner een map. Die ligt een jaar later nog in de kast.",
        "voorbeeld": "De beheerder vertelt bij de meterkast hoe de installaties werken, in een stem die je kent.",
        "tags": "#vastgoed #wonen #makelaar #klantbeleving #VoiceStamp",
    },
    "Evenementen & locaties": {
        "wie": "een evenementenlocatie", "klant": "bezoekers", "ding": "locatie",
        "plekken": ["bij de ingang", "bij een zaal", "bij een stand"],
        "inzicht": "Een programmaboekje is na een uur kwijt. Wat bij de plek hoort, blijft hangen.",
        "voorbeeld": "De organisator vertelt bij de ingang wat er te doen is en waar je het beste begint.",
        "tags": "#evenementen #events #beurs #congres #VoiceStamp",
    },
    "Onderwijs": {
        "wie": "een school of opleiding", "klant": "studenten", "ding": "gebouw",
        "plekken": ["bij de ingang", "bij een lokaal", "op de flyer voor de open dag"],
        "inzicht": "Op de eerste dag is alles nieuw. Veel uitleg gaat verloren.",
        "voorbeeld": "Een docent vertelt bij het lokaal wat je er kunt verwachten, op de open dag of de eerste schooldag.",
        "tags": "#onderwijs #opleiding #onboarding #opendag #VoiceStamp",
    },
    "Werkgevers & onboarding": {
        "wie": "een werkgever", "klant": "nieuwe medewerkers", "ding": "werkplek",
        "plekken": ["bij de werkplek", "bij een machine", "bij de ingang"],
        "inzicht": "Een nieuwe medewerker krijgt in de eerste week veel informatie. Het meeste is een week later vergeten.",
        "voorbeeld": "Een collega vertelt bij de werkplek precies wat je daar moet weten, op het moment dat het nodig is.",
        "tags": "#onboarding #hr #werkgeverschap #medewerkerervaring #VoiceStamp",
    },
    "Dealers & mobiliteit": {
        "wie": "een dealer", "klant": "klanten", "ding": "voertuig",
        "plekken": ["op het voertuig", "bij de showroomplek", "in de map bij aflevering"],
        "inzicht": "Bij de aflevering van een voertuig is er veel uit te leggen en weinig tijd.",
        "voorbeeld": "De monteur vertelt bij aflevering hoe je het voertuig het beste onderhoudt, in zijn eigen stem.",
        "tags": "#dealer #fiets #camper #klantbeleving #VoiceStamp",
    },
}

FORMATEN = ["inzicht", "zo werkt het", "voorbeeld", "vraag", "uitnodiging"]

# Voorgestelde posttijden. Dit is een voorstel, geen onderzoek: pas aan op wat bij jouw publiek werkt.
STANDAARD = {"LinkedIn": {"dagen": [1, 3], "tijd": "08:30"},      # di, do
             "Instagram": {"dagen": [0, 2, 5], "tijd": "18:00"}}  # ma, wo, za

def _begin(s):
    return s["inzicht"]

def maak_tekst(sector, kanaal, formaat, animatie=None):
    s = SECTOREN[sector]
    plek = s["plekken"][0]
    plek2 = s["plekken"][1] if len(s["plekken"]) > 1 else plek
    if formaat == "inzicht":
        li = (f"{s['inzicht']}\n\nWat als {s['wie']} dat anders doet?\n\n"
              f"Wie het maakte. Waarom het bijzonder is. Hoe je het het beste gebruikt. In een stem, op de plek waar het ertoe doet.\n\n"
              f"Geen app. Geen account. Gewoon luisteren.\n\nMeer over VoiceStamp: {SITE}")
        ig = (f"{s['inzicht']}\n\nWat als jouw {s['ding']} zelf kon vertellen? \U0001F50A\n\nGeen app. Geen account. Gewoon luisteren.")
    elif formaat == "zo werkt het":
        li = ("Zo werkt VoiceStamp, in drie stappen:\n\n1. Je plaatst een eenvoudige stempel.\n2. Je klant scant hem met de camera.\n"
              "3. Hij hoort direct jouw boodschap, in jouw stem.\n\nGeen app. Geen account. Gewoon luisteren.\n\n"
              f"Hoe het werkt: {SITE}/how-it-works")
        ig = ("Scan. Luister. Klaar. \U0001F3A7\n\n1️⃣ Plaats een stempel\n2️⃣ Je klant scant hem\n3️⃣ Hij hoort jouw stem\n\n"
              "Geen app. Geen account.")
    elif formaat == "voorbeeld":
        li = (f"Een voorbeeld voor {s['wie']}:\n\n{s['voorbeeld']}\n\nEen VoiceStamp kan {plek}, {plek2} of op een andere plek die ertoe doet.\n\n"
              f"En dat is maar één voorbeeld. VoiceStamp geeft elk product, elke plek en elk moment een stem.\n\n{SITE}")
        ig = (f"Voorbeeld voor {s['wie']} \U0001F447\n\n{s['voorbeeld']}\n\nTe plaatsen {plek}. Geen app nodig.")
    elif formaat == "vraag":
        li = (f"Een vraag aan iedereen met {s['wie'].replace('een ', '', 1)}:\n\nWelk moment of welke plek bij jullie verdient een stem?\n\n"
              "Ik ben benieuwd. Laat het hieronder weten.")
        ig = (f"Welke plek bij jou verdient een stem? \U0001F4AC\n\nLaat het hieronder weten, ik lees alles.")
    else:  # uitnodiging
        li = (f"Benieuwd hoe VoiceStamp eruitziet bij {s['wie']}?\n\nIk laat het graag zien, in een kort gesprek of met een voorbeeld op maat.\n\n"
              f"{LINKS}\n\nStuur me gerust een bericht.")
        ig = ("Benieuwd hoe het klinkt? \U0001F3A7\n\nStuur me een bericht, dan laat ik het je horen.\n\n\U0001F517 voicestamp.nl (link in bio)")
    if kanaal == "LinkedIn":
        tekst = li
    else:
        tekst = ig + "\n\n" + s["tags"]
    beeld = animatie or (f"Animatie of foto: {s['wie']}, {plek}" if formaat in ("inzicht", "voorbeeld")
                         else "Korte video of animatie van het scannen" if formaat == "zo werkt het"
                         else "Eenvoudige tekstkaart in de huisstijl")
    return tekst, beeld

def maak_plan(sectoren, kanalen, weken=4, start=None, per_week=None):
    """Roteert over sectoren en formaten. Geeft een lijst dicts, gesorteerd op datum en tijd."""
    start = start or date.today()
    per_week = per_week or {}
    posts = []
    teller = {k: 0 for k in kanalen}
    maandag = start - timedelta(days=start.weekday())
    for w in range(weken):
        for kanaal in kanalen:
            cfg = STANDAARD.get(kanaal, STANDAARD["LinkedIn"])
            dagen = cfg["dagen"][: per_week.get(kanaal, len(cfg["dagen"]))]
            for d in dagen:
                dag = maandag + timedelta(days=7 * w + d)
                if dag < start:
                    continue
                n = teller[kanaal]; teller[kanaal] += 1
                sector = sectoren[n % len(sectoren)]
                formaat = FORMATEN[(n // len(sectoren) + n) % len(FORMATEN)]
                tekst, beeld = maak_tekst(sector, kanaal, formaat)
                posts.append({"gepland_op": dag.isoformat(), "tijd": cfg["tijd"], "kanaal": kanaal,
                              "sector": sector, "formaat": formaat, "tekst": tekst, "beeld": beeld,
                              "status": "concept"})
    posts.sort(key=lambda p: (p["gepland_op"], p["tijd"], p["kanaal"]))
    return posts
