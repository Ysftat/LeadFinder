# VoiceStamp Lead Finder - handleiding

Een korte gids voor de hele app: leads vinden, mails, in-app mailen via Zoho, en het CRM.

## 1. Leads vinden
1. Kies links een **campagne** (Verblijf, Streek, Attracties, Erfgoed, Horeca, Winkels, Bloemisten).
2. Kies een **provincie** (of "Heel Nederland", maar dat is trager - houd het tabblad open).
3. Zet desgewenst de opties: websites verrijken (e-mail + socials + haak), MX-check, max 1 per
   domein, snelheid, en aantal per verzenddag.
4. Klik **Start**. Werk per provincie voor een snelle run. Dezelfde provincie opnieuw ophalen gaat
   direct (opgeslagen in cache).

## 2. Resultaten & mails
- Elke lead krijgt een **score** (balk), een **haak** (reden voor contact), een **kanaal-advies**
  en een **verzendbatch** (dagportie).
- Filter op score, alleen-met-e-mail of alleen-met-haak.
- Onder de tabel: kies een lead en **kopieer de mail** (Mail 1, Opvolgmail 2, Opvolgmail 3), of
  pak een **hele verzenddag** in één keer.
- Download alles als **Excel** of als **CSV** voor Smartlead/Instantly.

### Je mail aanpassen (optioneel)
Onder "Jouw merk & pitch" links kun je:
- de **aanhef** kiezen: "Beste team van [bedrijfsnaam]" (persoonlijk, scoort meestal beter) of
  "Beste team" (voor iedereen gelijk);
- de **productnaam** en het **links-blok** aanpassen;
- een **eigen pitch** invullen die de standaardtekst voor de hele campagne vervangt.
De mail eindigt altijd na "Met vriendelijke groet," - jouw eigen handtekening komt daaronder.

## 3. Direct mailen vanuit de app (Zoho)
Let op: dit werkt alleen zolang de app open staat in je browser. Er is geen onbewaakt versturen.
Er geldt een limiet van **15 mails per dag** (bewaakt je reputatie).

**Eenmalig: Zoho app-wachtwoord maken**
1. Log in op Zoho Mail > ga naar je account/beveiliging.
2. Zoek "App Passwords" / "App-specifieke wachtwoorden" en maak er één aan (bijv. naam "LeadFinder").
3. Kopieer dat wachtwoord. Gebruik dít, niet je gewone Zoho-wachtwoord.

**In de app**
1. Log eerst in bij **Account / CRM** (de mailfunctie verschijnt alleen als je bent ingelogd).
2. Klap **Mailen (Zoho)** open, vul in:
   - je Zoho-e-mailadres;
   - het app-wachtwoord;
   - de server: **smtp.zoho.eu** (meestal, voor Europese accounts) of **smtp.zoho.com**;
   - je afzendernaam;
   - je handtekening (komt onder "Met vriendelijke groet,").
3. Klik **Mail inloggen**. Bij een lead verschijnt nu **Verstuur via Zoho**.
4. Na versturen zet de app de status op **Gemaild** en logt het in de contacthistorie.

Werkt versturen niet? Meestal is het de verkeerde server (probeer de andere) of het gewone
wachtwoord in plaats van het app-wachtwoord.

**Uitloggen** bij Account / CRM wist meteen ook je Zoho-gegevens uit de sessie.

## 4. CRM
Vereist een gratis Supabase-koppeling (zie LEES_MIJ_crm.md).

- **Opslaan**: zoek leads en klik "Deze selectie opslaan in mijn CRM".
- **Importeren**: in het CRM-tabblad kun je je bestaande Excel importeren (kolommen als Naam,
  E-mail, Type, Status worden automatisch herkend, dubbele e-mails overgeslagen).
- **Dashboard**: bovenaan zie je de tellingen (Nieuw, Gemaild, Follow-up nodig, Reacties,
  Afspraken, Klanten) en de lijst **Vandaag opvolgen** (op basis van de datum volgende actie).
- **Bewerken**: per lead een Status, Volgende actie, Datum volgende actie, Notitie en (bij Do not
  contact) een reden. Klik "Wijzigingen opslaan".
- **Contacthistorie**: kies een lead en zie de tijdlijn (opgeslagen, gemaild, status gewijzigd).
  Je kunt zelf een notitie loggen.

## 5. Beslisser vinden, mailreeks en antwoorden (nieuw)

Dit is de Sellbird-aanpak in je eigen app. Je hebt er **sellbird_pipeline.py** voor nodig, in
dezelfde map als app.py, en je moet **supabase_setup.sql** opnieuw uitvoeren (veilig: het voegt alleen toe).

**Stap 1 - Beslisser en adres vinden** (Mijn CRM > Beslisser vinden & mailreeks klaarzetten)
- Kies leads (of "Eerste 25 zonder contactpersoon") en klik **Zoek eigenaar of manager + adres**.
- De app leest alleen de eigen pagina's van het bedrijf (home, over ons, team, contact), met
  robots.txt en pauzes. Geen LinkedIn.
- Elk adres krijgt een status: **published** (stond op hun eigen site), **verified** (gecontroleerd),
  **risky** (onzeker), **guess** (gok) of **invalid**. Een gok wordt nooit gebruikt; dan valt de
  mail terug op de gewone inbox met "Beste team van ...".
- Zelf gebouwde adressen (jan@, jan.jansen@) worden alleen betrouwbaar met een verificatiedienst.
  Zet `REOON_API_KEY` of `ZEROBOUNCE_API_KEY` in je secrets.toml (Reoon heeft een gratis pakket).

**Stap 2 - Reeks klaarzetten**
- Klik **Zet reeks klaar in Outbox**. Er komen 3 mails per lead: mail 1 vandaag, mail 2 na 4 dagen,
  mail 3 na 8 dagen, altijd op een werkdag. Er wordt nog niets verstuurd.
- Leads met een type dat de app niet herkent worden overgeslagen, zodat een winkel niet de
  campingtekst krijgt (of kies zelf een tekst bij "Type niet herkend").

**Stap 3 - Goedkeuren en versturen** (tabblad Outbox & antwoorden)
- Lees de drie mails per lead, pas ze aan en klik **Opslaan en hele reeks goedkeuren**.
- Voor goedkeuring en bij verzenden draait een controle zonder AI. Die blokkeert: ongeldig adres,
  gok-adres, lead die al gereageerd heeft of op Do not contact staat, niet ingevulde
  plaatshouders, domein zonder mailserver. Je krijgt een waarschuwing bij een lange mail,
  een link naar een ander domein of een ontbrekende afmeldregel.
- **Verstuur nu** verstuurt wat vandaag aan de beurt is, binnen je dagelijkse limiet, met een korte
  wisselende pauze tussen mails. Opvolgmails worden aan mail 1 gekoppeld (zelfde gesprek in de inbox
  van de ontvanger) en gaan nooit minder dan 3 dagen na de vorige.
- Dit werkt alleen terwijl de app open staat. Open hem elke werkdag even en klik op Verstuur nu.

**Stap 4 - Antwoorden**
- **Check mijn inbox** leest je inbox (alleen lezen) en herkent bounce, afmelding, afwijzing en reactie.
  Een reactie zet de lead op Reactie, een afmelding op Do not contact, een afwijzing op Afgewezen,
  een bounce markeert het adres als ongeldig. In alle gevallen stopt de rest van de reeks. Ook een
  antwoord van een andere persoon bij hetzelfde bedrijf telt.
- Lukt de inbox niet, plak het antwoord dan bij **Antwoord handmatig verwerken**.

**Let op: Zoho gratis.** Volgens meldingen van gebruikers heeft het gratis Zoho-pakket geen
POP/IMAP/SMTP meer. Dan werkt versturen en inbox lezen niet, en kan dat ook je
"535 Authentication Failed" verklaren. Controleer in Zoho > Instellingen > Mail > POP/IMAP.

## Belangrijk
- Resultaten worden niet vanzelf bewaard tenzij je ze downloadt of opslaat in het CRM.
- De app werkt alleen terwijl je browser-tabblad open staat.
- Zakelijk gebruik: benader relevante bedrijven, met afmeldmogelijkheid, binnen AVG en de
  Telecommunicatiewet (B2B). Verstuur gedoseerd.


## Campagnes kiezen

Links kies je eerst een **hoofdgroep** (bijvoorbeeld Zorg & welzijn) en daaronder de **categorieën** (Tandartspraktijken, Fysiotherapeuten, Huisartsen, enzovoort). Alles tegelijk zoeken kan: het is één zoekopdracht. De lijst met campagnes staat in `campagnes.py`. Zet dat bestand naast `app.py`.

Een paar branches zijn lastig te vinden op de kaart (foodtrucks, natuurparken, woningcorporaties, bedrijven met onboarding). Bij die categorieën staat in de zijbalk een korte waarschuwing. Evenementen zelf (festivals, beurzen, congressen) staan niet op de kaart: de app zoekt de locaties waar ze plaatsvinden.
