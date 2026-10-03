#!/usr/bin/env python3
"""Erzeugt aus schulen_2026-2027.yaml eine iCalendar-Datei (RFC 5545) mit allen Terminen.

Aufruf: .venv/bin/python build_ics.py [ausgabe.ics]
"""
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import yaml

BASE = Path(__file__).resolve().parent
QUELLE = BASE / "schulen_2026-2027.yaml"
ZIEL = BASE / "Schultermine_Darmstadt_2026-2027.ics"

STANDARDDAUER = timedelta(hours=2)
TZID = "Europe/Berlin"
UID_DOMAIN = "schuluebergang-2027.local"

TYP_LABEL = {
    "tag_der_offenen_tuer": "Tag der offenen Tür",
    "informationsabend": "Informationsabend",
    "schnuppertag": "Schnuppertag",
    "markt_fest": "Markt/Fest",
}

# Orte, die nicht auf dem Schulgelände liegen (Adresse also nicht die der Schule)
EXTERNE_ORTE = {
    "Pfälzer Schloss": "Pfälzer Schloss, 64823 Groß-Umstadt",
}

VTIMEZONE = """BEGIN:VTIMEZONE
TZID:Europe/Berlin
X-LIC-LOCATION:Europe/Berlin
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE""".splitlines()


def esc(text: str) -> str:
    """Escaping für TEXT-Werte (RFC 5545, 3.3.11)."""
    return (str(text).replace("\\", "\\\\").replace(";", "\\;")
            .replace(",", "\\,").replace("\r\n", "\n").replace("\n", "\\n"))


def fold(line: str) -> list[str]:
    """Zeilen auf max. 75 Oktette falten, ohne UTF-8-Zeichen zu zerteilen."""
    out, cur, cur_len = [], "", 0
    limit = 75
    for ch in line:
        n = len(ch.encode("utf-8"))
        if cur_len + n > limit:
            out.append(cur)
            cur, cur_len, limit = " ", 1, 75
        cur += ch
        cur_len += n
    out.append(cur)
    return out


def adresse(a: dict) -> str:
    return f"{a['strasse']}, {a['plz']} {a['ort']}"


def schuladresse(s: dict) -> dict:
    """Adresse des Schulgebäudes (MPG: Verwaltung und Schulgebäude getrennt)."""
    for a in s.get("weitere_anschriften", []):
        if a.get("zusatz") == "Schulgebäude":
            return a
    return s["anschrift"]


def location(s: dict, t: dict) -> str:
    if t.get("online"):
        return "Online"
    ort = t.get("ort")
    if ort in EXTERNE_ORTE:
        return EXTERNE_ORTE[ort]
    teile = [s["name"]]
    if ort and ort != "Schulgelände":
        teile.append(ort)
    teile.append(adresse(schuladresse(s)))
    return ", ".join(teile)


def kurzname(s: dict) -> str:
    return f"{s['name']} ({s['abkuerzung']})" if s.get("abkuerzung") else s["name"]


def beschreibung(s: dict, t: dict, meta: dict) -> str:
    z = []
    z.append(f"{t['art']} – {kurzname(s)}")
    zeit = f"{t['wochentag']}, {t['datum']:%d.%m.%Y}, {t['beginn']}"
    zeit += f"–{t['ende']} Uhr" if t.get("ende") else " Uhr (Endzeit nicht angegeben)"
    z.append(zeit)
    if t.get("zeit_hinweis"):
        z.append(f"Zeit: {t['zeit_hinweis']}")
    if t.get("ort"):
        z.append(f"Ort: {t['ort']}")
    if t.get("anmeldung_erforderlich"):
        z.append("ANMELDUNG ERFORDERLICH")
    if t.get("hinweis"):
        z.append(f"Hinweis: {t['hinweis']}")
    if t.get("hinweis_extraktion"):
        z.append(f"Hinweis zur Datenerfassung: {t['hinweis_extraktion']}")

    z += ["", "SCHULE"]
    z.append(f"Schulform: {s['schulform']} ({s['schulform_kurz']})")
    if s.get("privat"):
        z.append("Privatschule" + (f", Träger: {s['traeger']}" if s.get("traeger") else ""))
    elif s.get("traeger"):
        z.append(f"Träger: {s['traeger']}")
    z.append("Lage: " + ("Stadt Darmstadt" if s["stadt_darmstadt"] else "Landkreis Darmstadt-Dieburg"))
    for a in [s["anschrift"], *s.get("weitere_anschriften", [])]:
        zusatz = f" ({a['zusatz']})" if a.get("zusatz") else ""
        z.append(f"Anschrift{zusatz}: {adresse(a)}")
    z.append(f"Telefon: {s['telefon']}")
    z.append("E-Mail: " + ", ".join(s["email"]))
    z.append(f"Homepage: {s['homepage']}")
    if s.get("schulleitung"):
        z.append("Schulleitung: " + "; ".join(s["schulleitung"]))

    fs = s.get("fremdsprachen") or {}
    z += ["", "FREMDSPRACHEN"]
    for key, label in (("erste", "1."), ("zweite", "2."), ("dritte", "3.")):
        z.append(f"{label} Fremdsprache: {fs.get(key) or '–'}")
    z.append(f"Bilinguale Angebote: {s.get('bilinguale_angebote') or '–'}")

    if s.get("beschreibung"):
        z += ["", "BESCHREIBUNG", s["beschreibung"].strip()]
    if s.get("schwerpunkte"):
        z += ["", "SCHWERPUNKTE"] + [f"• {p}" for p in s["schwerpunkte"]]
    if s.get("ganztag"):
        z += ["", "GANZTAG"] + [f"• {p}" for p in s["ganztag"]]
    if s.get("termine_hinweis"):
        z += ["", f"Hinweis Termine: {s['termine_hinweis']}"]
    if s.get("hinweis_extraktion"):
        z += ["", f"Hinweis zur Datenerfassung: {s['hinweis_extraktion']}"]

    z += ["", f"Quelle: {meta['herausgeber']['name']}, Broschüre „{meta['titel']}“ "
              f"{meta['schuljahr_veranstaltungen']}, S. {s['seite']}"]
    return "\n".join(z)


def zeitpunkt(d: date, hhmm: str) -> datetime:
    h, m = map(int, hhmm.split(":"))
    return datetime(d.year, d.month, d.day, h, m)


def vevent(s: dict, t: dict, meta: dict, dtstamp: str) -> list[str]:
    start = zeitpunkt(t["datum"], t["beginn"])
    ende = zeitpunkt(t["datum"], t["ende"]) if t.get("ende") else start + STANDARDDAUER
    fmt = "%Y%m%dT%H%M%S"
    summary = f"{t['art']} – {kurzname(s)}"
    if t.get("anmeldung_erforderlich"):
        summary += " (Anmeldung erforderlich)"
    kategorien = [
        TYP_LABEL.get(t["typ"], t["typ"]),
        s["schulform_kurz"],
        "Privatschule" if s.get("privat") else "Öffentliche Schule",
        "Stadt Darmstadt" if s["stadt_darmstadt"] else "Landkreis Darmstadt-Dieburg",
    ]
    if t.get("online"):
        kategorien.append("Online")
    kontakt = f"{s['name']}, Tel. {s['telefon']}, {', '.join(s['email'])}"
    lines = [
        "BEGIN:VEVENT",
        f"UID:{s['id']}-{start:%Y%m%dT%H%M}@{UID_DOMAIN}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART;TZID={TZID}:{start.strftime(fmt)}",
        f"DTEND;TZID={TZID}:{ende.strftime(fmt)}",
        f"SUMMARY:{esc(summary)}",
        f"LOCATION:{esc(location(s, t))}",
        f"DESCRIPTION:{esc(beschreibung(s, t, meta))}",
        "CATEGORIES:" + ",".join(esc(k) for k in kategorien),
        f"URL;VALUE=URI:{s['homepage']}",
        f"CONTACT:{esc(kontakt)}",
        "CLASS:PUBLIC",
        "STATUS:CONFIRMED",
        "TRANSP:OPAQUE",
        "END:VEVENT",
    ]
    return lines


def main() -> None:
    ziel = Path(sys.argv[1]) if len(sys.argv) > 1 else ZIEL
    data = yaml.safe_load(QUELLE.read_text(encoding="utf-8"))
    meta, schulen = data["meta"], data["schulen"]
    dtstamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    events = sorted(
        ((t["datum"], t["beginn"], s, t) for s in schulen for t in s.get("termine", [])),
        key=lambda x: (x[0], x[1], x[2]["name"]),
    )

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Schulübergang 2027//build_ics.py//DE",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{esc('Weiterführende Schulen – Infotermine ' + meta['schuljahr_veranstaltungen'])}",
        f"X-WR-CALDESC:{esc(meta['untertitel'] + ', ' + meta['region'] + ' (Übergang ' + meta['uebergang'] + ')')}",
        f"X-WR-TIMEZONE:{TZID}",
        *VTIMEZONE,
    ]
    for _, _, s, t in events:
        lines += vevent(s, t, meta, dtstamp)
    lines.append("END:VCALENDAR")

    out = "".join(f"{part}\r\n" for line in lines for part in fold(line))
    ziel.write_text(out, encoding="utf-8", newline="")
    print(f"{len(events)} Termine → {ziel.name}")


if __name__ == "__main__":
    main()
