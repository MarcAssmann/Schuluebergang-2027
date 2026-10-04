# Schulübergang 2027 – weiterführende Schulen in Darmstadt und im Landkreis Darmstadt-Dieburg

Für die Eltern und Familien unserer 4. Klasse, die für das **Schuljahr 2027/2028** eine weiterführende Schule suchen.

Alle haben die Broschüre des Staatlichen Schulamts als PDF bekommen [Informationsveranstaltungen oder Tage der offenen Tür 2026/2027](https://marcassmann.github.io/Schuluebergang-2027/26-27_Mit%20Karten_Informationsveranstaltungen%20oder%20Tage%20der%20offenen%20Tür.pdf). Die Inhalte daraus gibt es hier in zwei zusätzlichen Formen:

- als **Excel-Tabelle**, um Schulen zu vergleichen und zu filtern, mit geschätzten Schulwegen
- als **Kalenderdateien**, damit man keinen Infoabend und keinen Tag der offenen Tür verpasst

> ⚠️ **Hinweis:** Die Tabelle und die Kalender wurden mit KI (Claude) aus dem PDF erzeugt. **Sie können deshalb Fehler enthalten.** Maßgeblich sind die Broschüre und die Homepages der Schulen. Bitte Termine vor dem Besuch dort noch einmal prüfen.

## 📊 Excel-Tabelle zur Schulauswahl

**Download:** [Schulauswahl_Darmstadt_2027-2028.xlsx](https://marcassmann.github.io/Schuluebergang-2027/Schulauswahl_Darmstadt_2027-2028.xlsx)

Die Datei hat drei Blätter:

- **Schulen**: eine Zeile pro Schule (38 Schulen) mit Schulform, Fremdsprachen, Oberstufe, G8/G9, Ganztag, Terminen und Kontaktdaten. Über die Pfeile in der Kopfzeile lässt sich filtern, z. B. „nur Gymnasien“, „nur Stadt Darmstadt“ oder „Latein als 2. Fremdsprache“.
- **Steckbriefe**: der vollständige Text jeder Schule aus der Broschüre, gut zum Lesen und Ausdrucken.
- **Anleitung**: Erklärungen zu den Spalten, Filterbeispiele und die Schulformen (Gym, KGS, IGS, MSS …) kurz erklärt.

**Geschätzte Schulwege:** Die grünen Spalten zeigen Strecke und Dauer **ab der Haltestelle „Darmstadt Lincoln-Siedlung“** (Tram 1/7/8) mit Bus & Bahn, Fahrrad, zu Fuß und Auto. Die Zeit für den eigenen Weg von zu Hause zur Haltestelle bitte dazurechnen. Grundlage:

- Bus & Bahn: späteste Verbindung mit Ankunft bis 7:50 Uhr laut Fahrplan an einem normalen Dienstag. Schulbusse und Verstärkerfahrten fehlen eventuell, daher bitte in der RMV-App gegenprüfen.
- Fahrrad und zu Fuß: Strecke laut OpenStreetMap, gerechnet im Kindertempo (12 km/h bzw. 4,5 km/h), ohne Ampeln und Steigungen.
- Auto: Fahrzeit bei freier Straße plus eine grobe Schätzung für den Berufsverkehr.

Die Tabelle enthält bewusst **keine Bewertungen** der Schulen.

## 📅 Kalender mit allen Terminen

Es gibt zwei Kalender. Bitte einen davon auswählen:

| Kalender | Inhalt | Link |
|---|---|---|
| **Stadt Darmstadt** | Termine der Schulen in der Stadt Darmstadt (inkl. Eberstadt) | https://marcassmann.github.io/Schuluebergang-2027/Schultermine_Stadt_Darmstadt_2026-2027.ics |
| **Alle Schulen** | Stadt Darmstadt und Landkreis Darmstadt-Dieburg | https://marcassmann.github.io/Schuluebergang-2027/Schultermine_alle_Schulen_2026-2027.ics |

Jeder Termin enthält Uhrzeit, Ort bzw. Adresse der Schule, Art der Veranstaltung (Infoabend, Tag der offenen Tür, Schnuppertag …), Hinweise wie „Anmeldung erforderlich“ und den Link zur Homepage der Schule. Ist in der Broschüre keine Endzeit angegeben, steht der Termin mit 2 Stunden im Kalender.

### So nutzt man die Kalender

Es gibt zwei Möglichkeiten:

- **Abonnieren (empfohlen):** Den Link oben kopieren und in der Kalender-App als *Kalender-Abonnement* bzw. *Kalender per URL* hinzufügen. Die Termine erscheinen dann als eigener Kalender, den man ein- und ausblenden kann. Korrekturen werden später automatisch übernommen (je nach App mit einigen Stunden Verzögerung).
- **Importieren:** Die Datei herunterladen und öffnen. Die Termine werden dann einmalig in einen bestehenden Kalender kopiert, spätere Korrekturen kommen aber nicht an.

Anleitungen der Hersteller:

- Google Kalender: [Kalender per URL hinzufügen](https://support.google.com/calendar/answer/37100?hl=de) (geht nur am Computer, danach erscheint der Kalender auch auf dem Handy)
- iPhone / iPad: [Kalender abonnieren](https://support.apple.com/de-de/guide/iphone/iph3d1110d4/ios) (Einstellungen → Apps → Kalender → Accounts → Account hinzufügen → Andere → Kalenderabo hinzufügen)
- Mac: [Abonnieren von Kalendern auf dem Mac](https://support.apple.com/de-de/guide/calendar/icl1022/mac)
- Outlook: [Kalender importieren oder abonnieren](https://support.microsoft.com/de-de/office/importieren-oder-abonnieren-eines-kalenders-in-outlook-com-oder-outlook-im-web-cff1429c-5af6-41ec-a5b4-74f2c278e98c)
- Android ohne Google Kalender: Viele Kalender-Apps können Abonnements; sonst hilft z. B. die App *ICSx⁵*.

## Bekannte Auffälligkeiten in der Broschüre

- **Justus-Liebig-Schule (LIO):** Der Tag der offenen Tür steht in der Broschüre am „16.01.2026“. Da der 16.01.2027 ein Samstag ist, ist vermutlich dieses Datum gemeint, und so steht er auch im Kalender.
- **Edith-Stein-Schule:** Die Postleitzahl ist in der Broschüre mit 65285 angegeben (vermutlich 64285). Sie wurde unverändert übernommen.
- Einige Termine aus der Broschüre lagen bei Erstellung schon in der Vergangenheit (September 2026).

Wer einen Fehler findet, sagt bitte Bescheid, dann wird er korrigiert.

---

<details>
<summary>Technisches (für Interessierte)</summary>

- `schulen_2026-2027.yaml`: strukturierte Daten aus dem PDF, die Quelle für alles andere
- `build_excel.py`: erzeugt die Excel-Tabelle
- `build_ics.py`: erzeugt die beiden Kalenderdateien
- `Entfernungen_der_Schulen_von_der_Haltestelle_Lincoln-Siedlung.yml`: Schulwege (OpenStreetMap/OSRM, Fahrplandaten über Transitous)

```sh
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python build_excel.py
.venv/bin/python build_ics.py
```

</details>
