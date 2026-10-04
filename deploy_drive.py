#!/usr/bin/env python3
"""Kopiert die Excel-Tabelle in den lokal eingebundenen Google-Drive-Ordner (Google Drive for Desktop).

Die Datei wird an Ort und Stelle überschrieben, damit Drive sie als neue Version derselben Datei
übernimmt und der Freigabelink erhalten bleibt. Den Upload erledigt der Drive-Client.

Aufruf: .venv/bin/python deploy_drive.py [zielordner]
"""
import shutil
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATEI = BASE / "Schulauswahl_Darmstadt_2027-2028.xlsx"
DRIVE = Path.home() / "Library/CloudStorage/GoogleDrive-marc.assmann@googlemail.com/My Drive"
ZIEL = DRIVE / "Shared" / "Schulübergang 2027"


def main():
    ziel = Path(sys.argv[1]) if len(sys.argv) > 1 else ZIEL
    if not DATEI.exists():
        sys.exit(f"{DATEI.name} fehlt – zuerst build_excel.py ausführen.")
    if not ziel.parent.is_dir():
        sys.exit(f"{ziel.parent} nicht gefunden – läuft Google Drive for Desktop?")
    ziel.mkdir(exist_ok=True)
    # copyfile schreibt in die bestehende Datei (kein Löschen/Umbenennen) → gleiche Drive-Datei-ID
    shutil.copyfile(DATEI, ziel / DATEI.name)
    print(f"{DATEI.name} → {ziel}")


if __name__ == "__main__":
    main()
