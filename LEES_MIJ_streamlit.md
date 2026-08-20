# VoiceStamp Lead Finder — webapp (Streamlit)

Een klikbare webversie: kies campagne en provincie, klik Start, download de Excel en de
Smartlead-CSV. Werkt in elke browser, ook op je telefoon.

## Eerst even lokaal testen (aanrader)
1. Installeer Python 3 (python.org) als je dat nog niet hebt.
2. In een terminal, in de map met `app.py`:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Je browser opent vanzelf op http://localhost:8501. Klaar.

## Gratis online zetten (zodat je 'm overal opent)
Je hebt een gratis GitHub-account nodig.

1. Maak op github.com een nieuwe (private of public) repository.
2. Zet daarin deze drie bestanden: `app.py`, `requirements.txt` en (optioneel) dit lees-mij.
   - Kan via de site: "Add file" -> "Upload files" -> sleep de bestanden erin -> "Commit".
3. Ga naar https://share.streamlit.io (Streamlit Community Cloud) en log in met GitHub.
4. Klik "New app", kies je repository, en als hoofdbestand `app.py`. Klik "Deploy".
5. Na een minuut heb je een webadres (bijv. voicestamp-leadfinder.streamlit.app) dat je
   overal kunt openen en delen.

## Wat de app doet
- Haalt campings/hotels/B&B's (of streek, attracties, erfgoed) uit OpenStreetMap.
- Vult ontbrekende e-mails aan via de websites en maakt per lead een concept-mail.
- Controleert e-mail (MX), bepaalt de "beste ingang", en deelt leads in verzenddagen.
- Levert een Excel (met concept-mails) en een CSV voor Smartlead/Instantly.

## Belangrijk om te weten
- **Scrapen vanuit de cloud** (Streamlit Community Cloud) wordt door sommige sites geblokkeerd.
  De OpenStreetMap-data komt altijd binnen; het aanvullen van e-mails lukt in de cloud wat
  minder vaak dan lokaal of in Colab. Voor maximale e-mail-dekking: draai 'm lokaal.
- Begin met één provincie en (bij het testen) een lage "Max sites", zodat je snel resultaat ziet.
- Verstuur gedoseerd (de kolom Verzendbatch helpt) en houd je aan B2B/AVG: relevante
  benadering met een duidelijke afmeldmogelijkheid.

## Later uitbreiden
Denk aan: opvolgmails (mail 2 en 3) als extra kolommen, voornaam-verrijking uit extra bronnen,
of inloggen zodat meerdere mensen 'm kunnen gebruiken. Vraag het gerust.
