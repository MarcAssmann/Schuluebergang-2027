# Schulübergang 2027

Private Recherche zur Wahl einer weiterführenden Schule (Übergang Klasse 4 → 5 zum Schuljahr 2027/2028) in der Stadt Darmstadt und im Landkreis Darmstadt-Dieburg. Sprache der Inhalte: Deutsch.

## Dateien

- `26-27_Mit Karten_Informationsveranstaltungen oder Tage der offenen Tür.pdf` – Quelle: Broschüre des Staatlichen Schulamts (Schulbeschreibungen S. 5–42, Termine S. 43–47).
- `schulen_2026-2027.yaml` – **maßgebliche strukturierte Datenquelle**, von Hand aus dem PDF extrahiert und geprüft. Weitere Ausgaben (Excel, Kalender) sollen hieraus erzeugt werden.
- `build_excel.py` – erzeugt aus der YAML `Schulauswahl_Darmstadt_2027-2028.xlsx` für die Eltern der Klasse (Blätter Schulen mit Filtern, Steckbriefe, Anleitung; Spaltenaufbau vom Nutzer per Hand in Excel festgelegt – Änderungen am Layout zuerst in Excel abstimmen und dann ins Skript übernehmen). Enthält bewusst keine persönlichen Schulwege/Bewertungen.
- `Schulauswahl_Darmstadt_2027-2028.xlsx` – kann manuell bearbeitet sein → vor dem Überschreiben mit `git status`/`git diff` prüfen, ob es uncommittete Änderungen gibt.

## YAML-Schema (`schulen_2026-2027.yaml`)

```yaml
meta:
  titel, untertitel, schuljahr_veranstaltungen, uebergang, region, quelle
  zeitzone: Europe/Berlin
  herausgeber: {name, dezernat, anschrift, telefon, email}
  schulform_abkuerzungen: {Gym, KGS, IGS, H/R, MSS, R, GO, priv.}

schulen:                      # 38 Einträge, alphabetisch wie im PDF
  - id: str                   # eindeutiger Kurzschlüssel, klein (z.B. aes, gbs, schuldorf)
    name: str
    abkuerzung: str | null    # aus dem Inhaltsverzeichnis; null bei Georg-Müller-Schule, Goetheschule, Schuldorf
    schulform_kurz: str       # z.B. "KGS+GO", "Gym", "IGS", "MSS", "R/Gym"
    schulform: str            # Langtext
    privat: bool
    traeger: str              # optional (ESS, PTID)
    anschrift: {strasse, plz (str), ort, zusatz?}
    weitere_anschriften: []   # optional (MPG)
    stadt_darmstadt: bool     # true = Stadt Darmstadt (inkl. Eberstadt), false = Landkreis
    telefon: str              # Schreibweise wie im PDF
    email: [str]              # immer Liste
    homepage: str             # immer mit https://
    schulleitung: [str]
    ganztag: [str]            # optional (nur GBS hat eigenen Abschnitt)
    schwerpunkte: [str]       # in einzelne Punkte zerlegt, Silbentrennung aufgelöst
    beschreibung: str         # optional, Fließtext (nur GUT)
    fremdsprachen: {erste, zweite, dritte}   # str | null
    bilinguale_angebote: str | null          # "/" im PDF → null
    seite: int                # Seite im PDF
    hinweis_extraktion: str   # optional, vermutlicher Fehler in der Vorlage
    termine_hinweis: str      # optional (z.B. GMS: noch keine Termine)
    termine:
      - datum: YYYY-MM-DD     # unquoted → wird als date geparst
        wochentag: str        # deutsch, ausgeschrieben; gegen datum validiert
        beginn: "HH:MM"       # IMMER in Anführungszeichen (YAML 1.1 sonst Sexagesimalzahl)
        ende: "HH:MM" | null  # null = keine Endzeit angegeben
        zeit_hinweis: str     # optional, z.B. "ab 10:00"
        art: str              # Bezeichnung wie im PDF
        typ: tag_der_offenen_tuer | informationsabend | schnuppertag | markt_fest
        ort: str | null
        online: bool          # optional
        anmeldung_erforderlich: bool   # optional
        hinweis: str          # optional
        hinweis_extraktion: str        # optional
```

## Bekannte Korrekturen/Auffälligkeiten gegenüber dem PDF

- LIO Tag der offenen Tür: PDF „16.01.2026“ → als 2027-01-16 erfasst (Samstag passt nur zu 2027).
- ESS: PLZ im PDF 65285 (vermutlich 64285), unverändert übernommen.
- PTID heißt im Terminteil „PTID Darmstadt“.
- Bereits vergangene Termine (Stand 2026-10-03): FWS 2026-09-12, SBS 2026-09-26.

## Umgebung / Hinweise

- **Keine Backup-Dateien anlegen** (kein `*.backup.*`, `*-bak-*` o.ä.) – weder von Hand noch in Skripten. Versionierung läuft über git.

- macOS; `pdftotext` (poppler) ist vorhanden. Python-Umgebung: venv unter `.venv/` mit `openpyxl`, `pyyaml`, `pdftotext` (siehe `requirements.txt`) → Skripte mit `.venv/bin/python` ausführen. Neu anlegen: `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt`.
- Nach Änderungen an der YAML validieren: parsebar, `id` eindeutig, `wochentag` passt zu `datum`, Zeiten sind Strings.
- Kalender-Export: Termine ohne `ende` sinnvoll behandeln (z.B. Standarddauer 2 h), Zeitzone Europe/Berlin.
