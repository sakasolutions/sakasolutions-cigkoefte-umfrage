# Çiğ Köfte Umfrage

Bewertungs-Kiosk fürs iPad (quer): Geschmack, Schärfe, Bestellabsicht in drei Taps, danach 7 Sekunden „Danke“ und automatisch von vorn. Die Antworten landen in einem Google Sheet.

Live: https://sakasolutions.github.io/sakasolutions-cigkoefte-umfrage/

## Google Sheet verbinden (einmalig, ca. 5 Minuten)

1. Neues Google Sheet anlegen, z. B. „Çiğ Köfte Bewertungen“.
2. Im Sheet: **Erweiterungen → Apps Script**.
3. Den Inhalt von `Code.gs` löschen und den kompletten Inhalt von [`apps-script.gs`](apps-script.gs) einfügen. Speichern.
4. **Bereitstellen → Neue Bereitstellung** → Typ (Zahnrad) **Web-App**:
   - Ausführen als: **Ich**
   - Zugriff: **Jeder**
5. **Bereitstellen** klicken, Zugriff erlauben (Google warnt „App nicht verifiziert“: *Erweitert → Weiter zu …*, das ist dein eigenes Script).
6. Die **Web-App-URL** (endet auf `/exec`) kopieren und in `index.html` bei `var SCRIPT_URL = "";` eintragen.

Das Tabellenblatt „Bewertungen“ legt das Script beim ersten Eintrag selbst an.

## Am iPad

- Seite in Safari öffnen → Teilen → **Zum Home-Bildschirm**. Von dort gestartet läuft sie im Vollbild ohne Adressleiste.
- Einstellungen → Anzeige & Helligkeit → Automatische Sperre: **Nie**.
- Einstellungen → Bedienungshilfen → **Geführter Zugriff** an, am Stand 3× Seitentaste.
- Auswertung: oben links ca. 1,5 Sekunden auf „Çiğ Köfte“ drücken.
- Punkt neben „vegan“: grün = alles im Sheet, gelb = Antworten warten noch auf Internet (werden automatisch nachgeschickt).
