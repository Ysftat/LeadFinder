"""Campagnes van de VoiceStamp Lead Finder: hoofdgroepen met subcategorieën.

Elke subcategorie zegt: welke OpenStreetMap-labels zoek je, hoe heet het type in de lijst en welke
mailtekst (segment) hoort erbij. Geen Streamlit nodig, dus los te testen.

Opbouw van een subcategorie:
    naam    tekst in de keuzelijst
    label   type zoals het in de leadlijst en het CRM komt te staan
    seg     mailsegment (zie SEGMENTEN onderaan)
    keys    {osm_sleutel: [waarden]}  een waarde "*" betekent: de sleutel hoeft alleen te bestaan
    naam_re (optioneel) de naam van de plek moet hierop passen (hoofdletterongevoelig)
    niet_re (optioneel) de naam mag hier juist NIET op passen
    ook     (optioneel) extra tags die ook moeten kloppen, bijv. {"mobile": "yes"}
    let_op  (optioneel) eerlijke kanttekening die in de app getoond wordt
"""
import re

CAMPAGNES = {
    "Verblijven & toerisme": [
        {"naam": "Campings & vakantieparken", "label": "Camping/vakantiepark", "seg": "natuurcamping",
         "keys": {"tourism": ["camp_site", "caravan_site", "chalet"]}},
        {"naam": "Hotels", "label": "Hotel", "seg": "hotel",
         "keys": {"tourism": ["hotel", "motel"]}},
        {"naam": "B&B's", "label": "B&B / guest house", "seg": "bnb",
         "keys": {"tourism": ["guest_house", "apartment"]}},
        {"naam": "Hostels", "label": "Hostel", "seg": "hotel",
         "keys": {"tourism": ["hostel"]}},
        {"naam": "Recreatieparken", "label": "Recreatiepark", "seg": "recreatie",
         "keys": {"leisure": ["resort", "water_park"]}},
        {"naam": "Natuurparken", "label": "Natuurpark", "seg": "recreatie",
         "keys": {"leisure": ["nature_reserve"], "boundary": ["national_park"]},
         "naam_re": r"natuurpark|nationaal park|nationaal landschap|natuurcentrum|bezoekerscentrum|landgoed",
         "let_op": "Natuurgebieden hebben zelden een eigen e-mailadres op de kaart. Verwacht hier minder resultaat."},
    ],
    "Eten, drinken & lokale producten": [
        {"naam": "Restaurants", "label": "Restaurant", "seg": "horeca",
         "keys": {"amenity": ["restaurant", "bistro"]}},
        {"naam": "Cafés & koffiebars", "label": "Café/koffiebar", "seg": "horeca",
         "keys": {"amenity": ["cafe"]}},
        {"naam": "Bakkerijen", "label": "Bakkerij", "seg": "streek",
         "keys": {"shop": ["bakery", "pastry"]}},
        {"naam": "Thee- en koffiemerken", "label": "Thee/koffie", "seg": "streek",
         "keys": {"shop": ["tea", "coffee"], "craft": ["coffee_roaster", "roaster"]},
         "let_op": "Merken zonder winkel staan niet op de kaart. Voor online merken is LinkedIn of Instagram beter."},
        {"naam": "Foodtruckondernemers", "label": "Foodtruck", "seg": "horeca",
         "keys": {"amenity": ["fast_food", "food_truck"]}, "ook": {"mobile": "yes"},
         "let_op": "Foodtrucks worden zelden op de kaart gezet. Verwacht weinig resultaat."},
    ],
    "Kunst, cultuur & erfgoed": [
        {"naam": "Musea", "label": "Museum", "seg": "erfgoed",
         "keys": {"tourism": ["museum", "gallery"]}},
    ],
    "Retail & winkels": [
        {"naam": "Kledingwinkels", "label": "Kledingwinkel", "seg": "winkel",
         "keys": {"shop": ["clothes", "boutique"]}},
        {"naam": "Meubelzaken", "label": "Meubelzaak", "seg": "winkel",
         "keys": {"shop": ["furniture"]}},
        {"naam": "Interieurwinkels", "label": "Interieurwinkel", "seg": "winkel",
         "keys": {"shop": ["interior_decoration", "houseware"]}},
        {"naam": "Conceptstores", "label": "Conceptstore", "seg": "winkel",
         "keys": {"shop": ["concept_store"]}},
        {"naam": "Boekhandels", "label": "Boekhandel", "seg": "winkel",
         "keys": {"shop": ["books"]}},
        {"naam": "Tuincentra", "label": "Tuincentrum", "seg": "winkel",
         "keys": {"shop": ["garden_centre"]}},
    ],
    "Bouw, wonen & vastgoed": [
        {"naam": "Makelaars", "label": "Makelaar", "seg": "vastgoed",
         "keys": {"office": ["estate_agent"]}},
        {"naam": "Woningcorporaties", "label": "Woningcorporatie", "seg": "vastgoed",
         "keys": {"office": ["*"]},
         "naam_re": r"woningcorporatie|woningbouwvereniging|woonstichting|woonbedrijf|woonstad|(^| )wonen( |$)|woonplus|woonzorg",
         "let_op": "Woningcorporaties hebben geen eigen label op de kaart. Ik zoek op de naam, dus controleer de lijst."},
        {"naam": "Vastgoedbeheerders", "label": "Vastgoedbeheerder", "seg": "vastgoed",
         "keys": {"office": ["property_management"]}},
    ],
    "Zorg & welzijn": [
        {"naam": "Tandartspraktijken", "label": "Tandarts", "seg": "zorg",
         "keys": {"amenity": ["dentist"], "healthcare": ["dentist"]}},
        {"naam": "Fysiotherapeuten", "label": "Fysiotherapeut", "seg": "zorg",
         "keys": {"healthcare": ["physiotherapist"], "amenity": ["physiotherapist"]}},
        {"naam": "Huisartsen", "label": "Huisarts", "seg": "zorg",
         "keys": {"amenity": ["doctors"], "healthcare": ["doctor"]}},
        {"naam": "Wellness", "label": "Wellness", "seg": "wellness",
         "keys": {"leisure": ["sauna"], "shop": ["massage", "beauty"], "amenity": ["public_bath"]}},
        {"naam": "Spa's", "label": "Spa", "seg": "wellness",
         "keys": {"leisure": ["spa"], "amenity": ["spa"]}},
        {"naam": "Yoga- en pilatesstudio's", "label": "Yoga/pilates", "seg": "wellness",
         "keys": {"sport": ["yoga", "pilates"]}},
    ],
    "Werkgevers & bedrijven": [
        {"naam": "Bedrijven met onboarding", "label": "Bedrijf", "seg": "werkgever",
         "keys": {"office": ["company"]},
         "let_op": "Op de kaart staat niet wie onboarding heeft. Dit zijn bedrijven in het algemeen. Voor HR-lijsten is het tabblad Bestand uploaden beter."},
        {"naam": "Fabrieken", "label": "Fabriek", "seg": "werkgever",
         "keys": {"man_made": ["works"]}},
        {"naam": "Kantoren", "label": "Kantoor", "seg": "werkgever",
         "keys": {"building": ["office"], "office": ["coworking"]}},
        {"naam": "Logistieke bedrijven", "label": "Logistiek bedrijf", "seg": "werkgever",
         "keys": {"industrial": ["logistics", "distribution_center"], "office": ["logistics"]}},
        {"naam": "Uitzendorganisaties", "label": "Uitzendbureau", "seg": "werkgever",
         "keys": {"office": ["employment_agency"]}},
    ],
    "Festivals, evenementen & entertainment": [
        {"naam": "Festivals", "label": "Festivalterrein/locatie", "seg": "evenement",
         "keys": {"amenity": ["events_venue"], "leisure": ["festival_ground"]},
         "naam_re": r"festival|evenemententerrein|openluchttheater",
         "let_op": "Evenementen zelf staan niet op de kaart. Dit zoekt de locaties waar ze plaatsvinden."},
        {"naam": "Concerten", "label": "Concertzaal/poppodium", "seg": "evenement",
         "keys": {"amenity": ["concert_hall", "music_venue", "nightclub"]},
         "let_op": "Evenementen zelf staan niet op de kaart. Dit zoekt de zalen en poppodia."},
        {"naam": "Beurzen", "label": "Beurslocatie", "seg": "evenement",
         "keys": {"amenity": ["exhibition_centre"]},
         "let_op": "Beurzen zelf staan niet op de kaart. Dit zoekt de beurslocaties."},
        {"naam": "Congressen", "label": "Congrescentrum", "seg": "evenement",
         "keys": {"amenity": ["conference_centre"]},
         "let_op": "Congressen zelf staan niet op de kaart. Dit zoekt de congrescentra."},
        {"naam": "Sportevenementen", "label": "Stadion/sportlocatie", "seg": "evenement",
         "keys": {"leisure": ["stadium"]},
         "let_op": "Sportevenementen zelf staan niet op de kaart. Dit zoekt de stadions."},
        {"naam": "Bedrijfsevenementen", "label": "Evenementenlocatie", "seg": "evenement",
         "keys": {"amenity": ["events_venue"]},
         "let_op": "Dit zoekt locaties voor bedrijfsevenementen, niet de evenementen zelf."},
        {"naam": "Theater", "label": "Theater", "seg": "entertainment",
         "keys": {"amenity": ["theatre"]}},
        {"naam": "Bioscopen", "label": "Bioscoop", "seg": "entertainment",
         "keys": {"amenity": ["cinema"]}},
        {"naam": "Escape rooms", "label": "Escape room", "seg": "entertainment",
         "keys": {"leisure": ["escape_game"]}},
        {"naam": "Pretparken", "label": "Pretpark", "seg": "attractie",
         "keys": {"tourism": ["theme_park", "zoo", "aquarium"]}},
    ],
    "Onderwijs & opleidingen": [
        {"naam": "Middelbare scholen", "label": "Middelbare school", "seg": "onderwijs",
         "keys": {"amenity": ["school"]},
         "naam_re": r"college|lyceum|gymnasium|scholengemeenschap|atheneum|vmbo|havo|vwo|praktijkonderwijs"},
        {"naam": "MBO", "label": "MBO", "seg": "onderwijs",
         "keys": {"amenity": ["college", "school"]},
         "naam_re": r"(^| )roc( |$)|(^| )mbo( |$)|vakschool|vakcollege|techniekcollege|regio ?college|albeda|mondriaan|deltion|landstede|noorderpoort|summa|aeres|scalda|friesland college|koning willem|da vinci|graafschap|rijn ijssel|horizon college|zadkine|nova college|gilde"},
        {"naam": "HBO", "label": "HBO", "seg": "onderwijs",
         "keys": {"amenity": ["university", "college"]},
         "naam_re": r"hogeschool|hbo|university of applied"},
        {"naam": "Universiteiten", "label": "Universiteit", "seg": "onderwijs",
         "keys": {"amenity": ["university"]},
         "niet_re": r"hogeschool|hbo|university of applied"},
        {"naam": "Kinderopvang", "label": "Kinderopvang", "seg": "onderwijs",
         "keys": {"amenity": ["kindergarten", "childcare"]}},
        {"naam": "Opleidingsinstituten", "label": "Opleidingsinstituut", "seg": "onderwijs",
         "keys": {"office": ["educational_institution"], "amenity": ["language_school", "music_school", "prep_school"]}},
        {"naam": "Trainingscentra", "label": "Trainingscentrum", "seg": "onderwijs",
         "keys": {"amenity": ["training"], "office": ["educational_institution"]},
         "naam_re": r"training|opleiding|academy|academie|instituut"},
    ],
    "Transport & mobiliteit": [
        {"naam": "Autodealers", "label": "Autodealer", "seg": "dealer",
         "keys": {"shop": ["car"]}},
        {"naam": "Fietsenwinkels", "label": "Fietsenwinkel", "seg": "dealer",
         "keys": {"shop": ["bicycle"]}},
        {"naam": "Camperdealers", "label": "Camperdealer", "seg": "dealer",
         "keys": {"shop": ["caravan", "motorhome"]}},
        {"naam": "Openbaar vervoer", "label": "Openbaar vervoer", "seg": "vervoer",
         "keys": {"railway": ["station"], "amenity": ["bus_station"]},
         "let_op": "Vervoerders staan als station op de kaart, zonder mailadres. Zoek de beslisser via LinkedIn."},
        {"naam": "Reisorganisaties", "label": "Reisorganisatie", "seg": "vervoer",
         "keys": {"shop": ["travel_agency"], "office": ["travel_agent"]}},
        {"naam": "Luchthavens", "label": "Luchthaven", "seg": "vervoer",
         "keys": {"aeroway": ["aerodrome"]},
         "let_op": "Er zijn maar weinig luchthavens. Dit zijn grote organisaties: zoek de beslisser via LinkedIn."},
    ],
    "Eerdere campagnes (extra)": [
        {"naam": "Streek (wijn, kaas, brouwers, boerderijwinkels)", "label": "Streekproducent", "seg": "streek",
         "keys": {"craft": ["winery", "brewery", "distillery", "cheese_making"],
                  "shop": ["wine", "cheese", "farm", "dairy", "greengrocer", "deli", "butcher",
                           "confectionery", "chocolate", "honey"]}},
        {"naam": "Cadeau, kunst en sieraden", "label": "Cadeau/kunst", "seg": "winkel",
         "keys": {"shop": ["gift", "art", "jewelry", "antiques", "pottery", "craft",
                           "musical_instrument", "watches"]}},
        {"naam": "Bloemisten", "label": "Bloemist", "seg": "bloemist",
         "keys": {"shop": ["florist"]}},
        {"naam": "Kastelen en landgoederen", "label": "Kasteel/landgoed", "seg": "erfgoed",
         "keys": {"historic": ["castle", "manor", "fort", "monastery"]}},
    ],
}

# ---------------------------------------------------------------- opzoeken
def alle_subs():
    return [s for subs in CAMPAGNES.values() for s in subs]

def label_naar_segment():
    """Type-label zoals het in het CRM staat -> mailsegment."""
    return {s["label"]: s["seg"] for s in alle_subs()}

def subs_van(groep, namen):
    return [s for s in CAMPAGNES[groep] if s["naam"] in namen]

# ---------------------------------------------------------------- Overpass
def clauses(subs):
    """Hashbare lijst (sleutel, waarde, naam_re, niet_re, ook) voor de zoekopdracht. Dubbele weglaten."""
    out, seen = [], set()
    for s in subs:
        ook = tuple(sorted((s.get("ook") or {}).items()))
        for k, vals in s["keys"].items():
            for v in vals:
                c = (k, v, s.get("naam_re") or "", s.get("niet_re") or "", ook)
                if c not in seen:
                    seen.add(c); out.append(c)
    return tuple(out)

def _esc(x):
    return x.replace("\\", "\\\\").replace('"', '\\"')

def overpass_stukken(cl):
    parts = []
    for k, v, ni, ne, ook in cl:
        f = f'["{k}"]' if v == "*" else f'["{k}"="{v}"]'
        if ni: f += f'["name"~"{_esc(ni)}",i]'
        if ne: f += f'["name"!~"{_esc(ne)}",i]'
        for ok, ov in ook: f += f'["{ok}"="{ov}"]'
        parts.append(f"nwr{f}(area.a);")
    return "".join(parts)

# ---------------------------------------------------------------- indelen
def past_bij(tags, sub):
    naam = tags.get("name") or ""
    if not any((v == "*" and k in tags) or tags.get(k) == v
               for k, vals in sub["keys"].items() for v in vals):
        return False
    if sub.get("naam_re") and not re.search(sub["naam_re"], naam, re.I):
        return False
    if sub.get("niet_re") and re.search(sub["niet_re"], naam, re.I):
        return False
    for k, v in (sub.get("ook") or {}).items():
        if tags.get(k) != v:
            return False
    return True

def deel_in(tags, subs):
    """Eerste subcategorie waar deze plek bij hoort, of None."""
    for s in subs:
        if past_bij(tags, s):
            return s
    return None

# ---------------------------------------------------------------- mailteksten voor nieuwe segmenten
_ZELFDE = ("Met VoiceStamp scant {wie} een eenvoudige stempel en hoort direct jullie boodschap, op het "
           "moment en de plek waar die relevant is.\n\n"
           "Geen app. Geen account. Gewoon luisteren.\n\n")

SEGMENTEN = {
 "zorg": {
  "onderwerp": "Wat als jullie uitleg ook te horen was, thuis?",
  "segwoord": "praktijk",
  "kern":
   "Wat als patiënten jullie uitleg niet alleen lezen, maar ook kunnen horen?\n\n"
   "Een welkom in de wachtkamer. Uitleg bij een behandeling of oefening. Een rustig woord voor het "
   "naar huis gaat.\n\n"
   + _ZELFDE.format(wie="een patiënt") +
   "Een VoiceStamp kan in de wachtkamer, bij een behandelruimte of op de folder die mee naar huis "
   "gaat. In jullie eigen stem.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal in jullie praktijk een uitleg, boodschap of "
   "moment een stem geven.\n\n"
   "Zo hoef je dezelfde uitleg niet elke keer opnieuw te geven, en gaat hij mee naar huis."},
 "wellness": {
  "onderwerp": "Wat als jullie plek ook te horen was?",
  "segwoord": "plek",
  "kern":
   "Wat als gasten jullie plek niet alleen zien en voelen, maar ook kunnen horen?\n\n"
   "Een welkom bij binnenkomst. Uitleg bij de ruimtes en de huisregels. Een korte ademoefening in "
   "jullie eigen stem.\n\n"
   + _ZELFDE.format(wie="een gast") +
   "Een VoiceStamp kan bij de ingang, in de kleedkamer of bij de studio. In jullie eigen stem.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal in jullie zaak een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je gasten iets mee dat op een bordje vaak verloren gaat, zonder de rust te verstoren."},
 "vastgoed": {
  "onderwerp": "Wat als een woning zelf zijn verhaal kon vertellen?",
  "segwoord": "organisatie",
  "kern":
   "Wat als een woning of gebouw niet alleen gezien wordt, maar ook iets kan vertellen?\n\n"
   "Het verhaal van het huis en de buurt. Uitleg bij oplevering of bij de installaties. Een welkom "
   "voor nieuwe bewoners.\n\n"
   + _ZELFDE.format(wie="een bewoner of kijker") +
   "Een VoiceStamp kan bij de voordeur, bij de meterkast of in de gezamenlijke ruimte. In jullie "
   "eigen stem.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan bij elk pand of elke woning een verhaal, uitleg of "
   "moment een stem geven.\n\n"
   "Zo geef je bewoners iets mee dat in een map vaak verloren gaat, en scheelt het vragen die nu "
   "telkens opnieuw komen."},
 "werkgever": {
  "onderwerp": "Nieuwe medewerkers sneller op weg, zonder extra app",
  "segwoord": "organisatie",
  "kern":
   "Wat als een nieuwe medewerker de belangrijkste info niet hoeft te lezen, maar gewoon kan "
   "horen?\n\n"
   "Een welkom van het team op de eerste dag. Uitleg bij een werkplek, precies waar iemand het nodig "
   "heeft. Een collega die op het juiste moment zijn kennis deelt.\n\n"
   + _ZELFDE.format(wie="iemand") +
   "Een VoiceStamp kan bij de werkplek, bij een machine of bij de ingang. In jullie eigen stem, of "
   "die van een collega.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal waar kennis of een boodschap op het juiste "
   "moment moet landen een stem geven.\n\n"
   "Zo staat de uitleg klaar, telkens hetzelfde, ook als het druk is."},
 "evenement": {
  "onderwerp": "Wat als jullie evenement ook te horen was?",
  "segwoord": "locatie",
  "kern":
   "Wat als bezoekers jullie locatie niet alleen zien, maar ook kunnen horen?\n\n"
   "Een welkom bij binnenkomst. Het programma van de dag. Uitleg bij een stand of een zaal, precies "
   "op het moment dat het nodig is.\n\n"
   + _ZELFDE.format(wie="een bezoeker") +
   "Een VoiceStamp kan bij de ingang, bij een zaal of bij een stand. In de stem van de organisator "
   "of de maker.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal op jullie locatie een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je bezoekers iets mee dat op een bordje of in een programmaboekje vaak verloren gaat."},
 "entertainment": {
  "onderwerp": "Wat als jullie voorstelling ook voor de deur te horen was?",
  "segwoord": "zaak",
  "kern":
   "Wat als bezoekers jullie programma niet alleen lezen, maar ook kunnen horen?\n\n"
   "Een woord van de maker voor de voorstelling begint. Het verhaal achter een film of stuk. Een "
   "welkom bij binnenkomst.\n\n"
   + _ZELFDE.format(wie="een bezoeker") +
   "Een VoiceStamp kan bij de kassa, op het affiche of bij de ingang van een zaal. In de stem van de "
   "maker of van jullie team.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal in jullie zaak een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je bezoekers iets mee dat op een affiche vaak verloren gaat, ook als de zaal vol zit."},
 "onderwijs": {
  "onderwerp": "Wat als jullie school ook te horen was?",
  "segwoord": "organisatie",
  "kern":
   "Wat als nieuwe studenten, leerlingen of ouders jullie uitleg niet alleen lezen, maar ook kunnen "
   "horen?\n\n"
   "Een welkom op de eerste dag. Een rondleiding op de open dag, in de stem van een docent. Uitleg "
   "bij een praktijkruimte of een regel.\n\n"
   + _ZELFDE.format(wie="iemand") +
   "Een VoiceStamp kan bij de ingang, bij een lokaal of op de flyer voor de open dag. In de stem van "
   "een docent of student.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal op school een uitleg, verhaal of moment een "
   "stem geven.\n\n"
   "Zo staat de uitleg klaar, telkens hetzelfde, ook als het druk is."},
 "dealer": {
  "onderwerp": "Wat als jullie producten iets konden vertellen?",
  "segwoord": "zaak",
  "kern":
   "Wat als een klant jullie voertuig niet alleen ziet, maar ook iets kan horen?\n\n"
   "Waarom dit model is gekozen. Hoe je het het beste gebruikt en onderhoudt. Een woord van de "
   "monteur bij aflevering.\n\n"
   + _ZELFDE.format(wie="een klant") +
   "Een VoiceStamp kan op het voertuig, bij de showroomplek of in de map bij aflevering. In de stem "
   "van de verkoper of de monteur.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan bij elk product een verhaal, uitleg of moment een "
   "stem geven.\n\n"
   "Zo geef je klanten iets mee dat in de showroom verteld wordt, maar anders niet meegaat naar huis."},
 "vervoer": {
  "onderwerp": "Wat als jullie reizigers iets konden horen, op de plek zelf?",
  "segwoord": "organisatie",
  "kern":
   "Wat als reizigers jullie informatie niet alleen lezen, maar ook kunnen horen?\n\n"
   "Een welkom bij de eerste stap. Uitleg bij een route of overstap. Een tip voor onderweg, op het "
   "moment dat het nodig is.\n\n"
   + _ZELFDE.format(wie="een reiziger") +
   "Een VoiceStamp kan op een perron, bij een gate of op een reisdocument. In een stem die past bij "
   "jullie.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal onderweg een verhaal, boodschap of moment "
   "een stem geven.\n\n"
   "Zo geef je reizigers iets mee dat op een bord vaak verloren gaat, zonder extra schermen."},
 "recreatie": {
  "onderwerp": "Wat als jullie plek zelf haar verhaal kon vertellen?",
  "segwoord": "plek",
  "kern":
   "Wat als bezoekers jullie plek niet alleen zien, maar ook kunnen horen?\n\n"
   "Een welkom bij aankomst. Het verhaal van het gebied. Een persoonlijke tip voor de mooiste "
   "route.\n\n"
   + _ZELFDE.format(wie="een bezoeker") +
   "Een VoiceStamp kan bij de ingang, langs een route of bij een bijzonder punt. In de stem van een "
   "gids of beheerder.\n\n"
   "En dat is maar één voorbeeld. VoiceStamp kan overal in jullie gebied een verhaal, boodschap of "
   "moment een stem geven.\n\n"
   "Zo geef je bezoekers iets mee dat op een bordje vaak verloren gaat, zonder extra drukte."},
}
