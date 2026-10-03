#!/usr/bin/env python3
"""Erzeugt aus schulen_2026-2027.yaml eine Excel-Arbeitsmappe für die Eltern der Klasse.

Blätter: Schulen (Übersicht mit Filtern), Steckbriefe (Volltext je Schule, druckbar), Anleitung.

Aufruf: .venv/bin/python build_excel.py [ausgabe.xlsx]
"""
import re
import sys
from datetime import date
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
QUELLE = BASE / "schulen_2026-2027.yaml"
ZIEL = BASE / "Schulauswahl_Darmstadt_2027-2028.xlsx"

HEUTE = date.today()

# Stichwortsuche in den Schwerpunkten → Ganztag-Spalten (leer = nicht erwähnt)
GANZTAG_RX = r"ganztag|betreuung|mittag|hausaufgaben|pakt für den nachmittag"

# ---------- Stil ----------
DUNKEL = "1F4E78"
FARBE_KOPF = {"info": "1F4E78", "detail": "595959"}
HELLBLAU = "DDEBF7"
WEISS = Font(color="FFFFFF", bold=True)
FETT = Font(bold=True)
DUENN = Side(style="thin", color="BFBFBF")
RAHMEN = Border(left=DUENN, right=DUENN, top=DUENN, bottom=DUENN)
OBEN_UMBRUCH = Alignment(vertical="top", wrap_text=True)
MITTE = Alignment(horizontal="center", vertical="top", wrap_text=True)
KOPF_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def fill(farbe: str) -> PatternFill:
    return PatternFill("solid", fgColor=farbe)


def adresse(a: dict) -> str:
    return f"{a['strasse']}, {a['plz']} {a['ort']}"


def ganztag_punkte(s: dict) -> list[str]:
    """Eigener Ganztag-Abschnitt (GBS) plus Schwerpunkte, die Ganztag/Betreuung erwähnen."""
    return s.get("ganztag", []) + [x for x in s["schwerpunkte"] if re.search(GANZTAG_RX, x.lower())]


def aufzaehlung(punkte: list[str]) -> str:
    return "\n".join(f"• {x}" for x in punkte)


def eigene_oberstufe(s: dict) -> bool:
    k = s["schulform_kurz"]
    return "Gym" in k or "GO" in k or "oberstufe" in s["schulform"].lower()


def g8_g9(s: dict) -> str:
    m = re.findall(r"G[89]", s["schulform"])
    return "/".join(dict.fromkeys(m))


def termin_zeit(t: dict) -> str:
    if t.get("zeit_hinweis"):
        return f"{t['zeit_hinweis']} Uhr"
    return f"{t['beginn']}–{t['ende']} Uhr" if t["ende"] else f"{t['beginn']} Uhr"


def termin_kurz(t: dict) -> str:
    txt = f"• {t['wochentag'][:2]} {t['datum']:%d.%m.%Y}, {termin_zeit(t)}: {t['art']}"
    if t.get("ort") and t["ort"] != "Schulgelände":
        txt += f" ({t['ort']})"
    if t["datum"] < HEUTE:
        txt += " – vorbei"
    return txt


def hinweise(s: dict) -> str:
    teile = [s.get("termine_hinweis"), s.get("hinweis_extraktion")]
    for t in s["termine"]:
        teile += [t.get("hinweis"), t.get("hinweis_extraktion")]
    return "; ".join(x for x in teile if x)


def zeilenhoehe(ws, r, sp) -> float:
    """Zeilenhöhe aus Textlänge und Spaltenbreite schätzen (Excel passt sie nicht selbst an)."""
    zeilen = 1
    for col, d in enumerate(sp, start=1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str):
            pro_zeile = max(1, int(d["breite"] * 1.3))
            zeilen = max(zeilen, sum(-(-max(1, len(z)) // pro_zeile) for z in v.split("\n")))
    return min(409, max(30, 15 * zeilen))


# ---------------------------------------------------------------- Blatt „Schulen“
def blatt_schulen(wb, schulen, steckbrief_zeile):
    ws = wb.create_sheet("Schulen")
    S, D = "info", "detail"
    sp = []

    def add(titel, gruppe, breite, wert, **kw):
        sp.append(dict(titel=titel, gruppe=gruppe, breite=breite, wert=wert, **kw))

    add("Schule", S, 30, lambda s: s["name"], fett=True)
    add("Homepage", S, 26, "homepage")
    add("Abk.", S, 8, lambda s: s["abkuerzung"] or "", mitte=True)
    add("Schulform", S, 11, lambda s: s["schulform_kurz"], mitte=True)
    add("Träger", S, 10, lambda s: "privat" if s["privat"] else "öffentlich", mitte=True)
    add("Lage", S, 13, lambda s: "Stadt Darmstadt" if s["stadt_darmstadt"] else "Landkreis", mitte=True)
    add("Ort", S, 16, lambda s: s["anschrift"]["ort"], ausgeblendet=True)
    add("Eigene Oberstufe (Abitur)", S, 11, lambda s: "Ja" if eigene_oberstufe(s) else "Nein", mitte=True)
    add("G8 / G9", S, 8, g8_g9, mitte=True)
    add("1. Fremdsprache", S, 18, lambda s: s["fremdsprachen"]["erste"] or "")
    add("2. Fremdsprache", S, 28, lambda s: s["fremdsprachen"]["zweite"] or "–")
    add("3. Fremdsprache", S, 28, lambda s: s["fremdsprachen"]["dritte"] or "–")
    add("Ganztag / Betreuung", S, 10, lambda s: "✓" if ganztag_punkte(s) else "", mitte=True)
    add("Ganztag / Betreuung (lt. Broschüre)", S, 45, lambda s: aufzaehlung(ganztag_punkte(s)))
    add("Besondere Schwerpunkte", S, 70, lambda s: aufzaehlung(s["schwerpunkte"]))
    add("Termine 2026/27", S, 48,
        lambda s: "\n".join(termin_kurz(t) for t in s["termine"]) or (s.get("termine_hinweis") or "–"))

    add("Schulform (ausführlich)", D, 30, lambda s: s["schulform"])
    add("Anschrift", D, 30, lambda s: adresse(s["anschrift"]))
    add("Bilinguale Angebote", D, 30, lambda s: s["bilinguale_angebote"] or "–")
    add("Telefon", D, 16, lambda s: s["telefon"])
    add("E-Mail", D, 30, "email")
    add("Schulleitung", D, 28, lambda s: "\n".join(s["schulleitung"]))
    add("Hinweise", D, 30, hinweise)
    add("Seite im PDF", D, 8, lambda s: s["seite"], mitte=True)

    for col, d in enumerate(sp, start=1):
        c = ws.cell(row=1, column=col, value=d["titel"])
        c.font, c.fill, c.alignment, c.border = WEISS, fill(FARBE_KOPF[d["gruppe"]]), KOPF_ALIGN, RAHMEN
        dim = ws.column_dimensions[get_column_letter(col)]
        dim.width = d["breite"]
        dim.hidden = d.get("ausgeblendet", False)
    ws.row_dimensions[1].height = 60

    for r, s in enumerate(schulen, start=2):
        for col, d in enumerate(sp, start=1):
            c = ws.cell(row=r, column=col)
            w = d["wert"]
            if w == "homepage":
                c.value = s["homepage"].removeprefix("https://")
                c.hyperlink = s["homepage"]
                c.font = Font(color="0563C1", underline="single")
            elif w == "email":
                c.value = "\n".join(s["email"])
                c.hyperlink = f"mailto:{s['email'][0]}"
                c.font = Font(color="0563C1", underline="single")
            else:
                c.value = w(s)
            c.alignment = MITTE if d.get("mitte") else OBEN_UMBRUCH
            c.border = RAHMEN
            if d.get("fett"):
                c.font = Font(bold=True, color="0563C1", underline="single")
                c.hyperlink = f"#'Steckbriefe'!A{steckbrief_zeile[s['id']]}"
        ws.row_dimensions[r].height = zeilenhoehe(ws, r, sp)

    letzte = 1 + len(schulen)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(sp))}{letzte}"
    ws.freeze_panes = "B2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    return ws


# ---------------------------------------------------------------- Blatt „Steckbriefe“
def blatt_steckbriefe(wb, schulen):
    """Volltext je Schule untereinander; liefert {id: Zeile der Überschrift}."""
    ws = wb.create_sheet("Steckbriefe")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 95
    zeilen = {}
    r = 1
    c = ws.cell(row=r, column=1, value="Steckbriefe der Schulen (Volltext aus der Broschüre)")
    c.font = Font(bold=True, size=14, color=DUNKEL)
    r += 2
    for s in schulen:
        zeilen[s["id"]] = r
        titel = s["name"] + (f" ({s['abkuerzung']})" if s["abkuerzung"] else "")
        for col in (1, 2):
            ws.cell(row=r, column=col).fill = fill(DUNKEL)
        c = ws.cell(row=r, column=1, value=titel)
        c.font = Font(bold=True, size=12, color="FFFFFF")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.row_dimensions[r].height = 20
        r += 1
        anschriften = [adresse(s["anschrift"])] + [
            f"{a.get('zusatz', '')}: {adresse(a)}".strip(": ") for a in s.get("weitere_anschriften", [])]
        felder = [
            ("Schulform", s["schulform"]),
            ("Träger", ("privat" + (f" ({s['traeger']})" if s.get("traeger") else "")) if s["privat"] else "öffentlich"),
            ("Anschrift", "\n".join(anschriften)),
            ("Telefon", s["telefon"]),
            ("E-Mail", ", ".join(s["email"])),
            ("Homepage", s["homepage"]),
            ("Schulleitung", "\n".join(s["schulleitung"])),
            ("Beschreibung", s.get("beschreibung")),
            ("Ganztag", aufzaehlung(s.get("ganztag", [])) or None),
            ("Schwerpunkte", aufzaehlung(s["schwerpunkte"])),
            ("1. Fremdsprache", s["fremdsprachen"]["erste"] or "–"),
            ("2. Fremdsprache", s["fremdsprachen"]["zweite"] or "–"),
            ("3. Fremdsprache", s["fremdsprachen"]["dritte"] or "–"),
            ("Bilinguale Angebote", s["bilinguale_angebote"] or "–"),
            ("Termine 2026/27", "\n".join(termin_kurz(t) for t in s["termine"]) or s.get("termine_hinweis")),
            ("Hinweise", hinweise(s) or None),
            ("Quelle", f"Broschüre, Seite {s['seite']}"),
        ]
        for label, wert in felder:
            if not wert:
                continue
            a = ws.cell(row=r, column=1, value=label)
            a.font, a.alignment, a.fill = FETT, OBEN_UMBRUCH, fill(HELLBLAU)
            b = ws.cell(row=r, column=2, value=wert)
            b.alignment = OBEN_UMBRUCH
            if label == "Homepage":
                b.hyperlink, b.font = wert, Font(color="0563C1", underline="single")
            # Zeilenhöhe grob schätzen (≈ 100 Zeichen pro Zeile bei Breite 95)
            n = sum(max(1, -(-len(z) // 100)) for z in str(wert).split("\n"))
            ws.row_dimensions[r].height = max(15, 15 * n)
            for col in (1, 2):
                ws.cell(row=r, column=col).border = RAHMEN
            r += 1
        z = ws.cell(row=r, column=2, value="→ zurück zur Übersicht")
        z.hyperlink = "#'Schulen'!A1"
        z.font = Font(color="0563C1", underline="single", size=9)
        r += 2
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws, zeilen


# ---------------------------------------------------------------- Blatt „Anleitung“
def blatt_anleitung(wb, meta, schulen, n_termine):
    ws = wb.create_sheet("Anleitung")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 95
    r = 1

    def zeile(b="", c="", stil=None, hoehe=None):
        nonlocal r
        cb, cc = ws.cell(row=r, column=2, value=b), ws.cell(row=r, column=3, value=c)
        cb.alignment = cc.alignment = OBEN_UMBRUCH
        if not c:  # Text ohne Wert in Spalte C läuft über die Breite, ohne Umbruch
            cb.alignment = Alignment(vertical="top")
        if stil == "titel":
            cb.font = Font(bold=True, size=16, color=DUNKEL)
        elif stil == "abschnitt":
            cb.font = Font(bold=True, size=12, color=DUNKEL)
        elif stil == "label":
            cb.font = FETT
        if hoehe:
            ws.row_dimensions[r].height = hoehe
        r += 1

    zeile(f"Weiterführende Schulen – Übergang {meta['uebergang']}", stil="titel", hoehe=24)
    zeile(meta["region"])
    zeile(f"{len(schulen)} Schulen, {n_termine} Termine (Infoabende, Tage der offenen Tür, Schnuppertage). "
          f"Stand der Tabelle: {HEUTE:%d.%m.%Y}.")
    zeile()
    zeile("So nutzen Sie die Tabelle", stil="abschnitt")
    zeile("Blatt „Schulen“", "Eine Zeile pro Schule. Blaue Spalten: Angaben aus der Broschüre. "
          "Graue Spalten ganz rechts: Anschrift, Kontaktdaten und weitere Details. "
          "Ein Klick auf den Schulnamen öffnet den Steckbrief.", "label", 30)
    zeile("Blatt „Steckbriefe“", "Der vollständige Text jeder Schule aus der Broschüre – gut zum Lesen und Ausdrucken.",
          "label")
    zeile()
    zeile("Filtern", stil="abschnitt")
    zeile("", "Klicken Sie auf den kleinen Pfeil in der Kopfzeile einer Spalte und wählen Sie die gewünschten Werte. "
          "Mehrere Filter lassen sich kombinieren. Zum Zurücksetzen: Daten → Filter löschen.", hoehe=30)
    zeile("Beispiele", "• Nur Gymnasien: Spalte „Schulform“ → „Gym“\n"
          "• Nur Schulen in der Stadt: Spalte „Lage“ → „Stadt Darmstadt“\n"
          "• Abitur an derselben Schule: Spalte „Eigene Oberstufe“ → „Ja“\n"
          "• Latein gewünscht: Spalte „2. Fremdsprache“ → Textfilter „enthält“ → Latein\n"
          "• Mit Ganztag/Betreuung: Spalte „Ganztag / Betreuung“ → „✓“",
          "label", 68)
    zeile("Ganztag / Betreuung", "Das ✓ und der Text daneben stammen aus einer Stichwortsuche in den Schwerpunkten "
          "(Ganztag, Betreuung, Mittagessen, Hausaufgaben). Ein leeres Feld heißt nur: in der Broschüre nicht "
          "erwähnt – nicht, dass es das Angebot nicht gibt. Bitte im Zweifel auf der Homepage nachsehen.", "label", 45)
    zeile()
    zeile("Schulformen", stil="abschnitt")
    for k, v in meta["schulform_abkuerzungen"].items():
        zeile(k, v, "label")
    zeile("Erläuterung", "• Gymnasium: führt direkt zum Abitur.\n"
          "• KGS: Haupt-, Real- und Gymnasialzweig unter einem Dach, in getrennten Klassen.\n"
          "• IGS: Alle Kinder lernen gemeinsam; Abschlüsse aller Bildungsgänge möglich, Kurse nach Leistungsniveau.\n"
          "• Mittelstufenschule (MSS): Haupt- und Realschulbildungsgang mit starkem Praxisbezug.\n"
          "• G8 / G9: Abitur nach 8 bzw. 9 Jahren Gymnasium.\n"
          "• Ohne eigene Oberstufe: Für das Abitur wird nach Klasse 10 an eine Oberstufe gewechselt.",
          "label", 95)
    zeile()
    zeile("Wichtige Hinweise", stil="abschnitt")
    q = meta["herausgeber"]
    zeile("Quelle", f"Broschüre „{meta['titel']} – {meta['untertitel']} {meta['schuljahr_veranstaltungen']}“, "
          f"{q['name']}, {q['anschrift']}.", "label", 30)
    zeile("Ohne Gewähr", "Die Angaben wurden von Hand aus der Broschüre übernommen. Termine können sich ändern – "
          "bitte vor dem Besuch auf der Homepage der Schule prüfen.", "label", 30)
    zeile("Auffälligkeiten", "• Justus-Liebig-Schule: In der Broschüre steht der Tag der offenen Tür am "
          "„16.01.2026“; da es ein Samstag ist, ist der 16.01.2027 gemeint.\n"
          "• Edith-Stein-Schule: PLZ in der Broschüre 65285 (vermutlich 64285).",
          "label", 45)
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def main():
    ziel = Path(sys.argv[1]) if len(sys.argv) > 1 else ZIEL
    daten = yaml.safe_load(QUELLE.read_text(encoding="utf-8"))
    schulen = daten["schulen"]

    wb = Workbook()
    wb.remove(wb.active)
    # Steckbriefe zuerst erzeugen (Zeilennummern für die Links), Reihenfolge danach korrigieren
    steckbriefe, zeilen = blatt_steckbriefe(wb, schulen)
    schul_ws = blatt_schulen(wb, schulen, zeilen)
    n_termine = sum(len(s["termine"]) for s in schulen)
    anleitung = blatt_anleitung(wb, daten["meta"], schulen, n_termine)
    wb._sheets = [schul_ws, steckbriefe, anleitung]
    wb.active = 0
    wb.properties.title = "Schulauswahl Darmstadt 2027/2028"
    wb.save(ziel)
    print(f"{ziel.name}: {len(schulen)} Schulen, {n_termine} Termine")


if __name__ == "__main__":
    main()
