# -*- coding: utf-8 -*-
"""
PYRevit Script: Filesize
Zeigt Dateigröße an, egal ob lokal oder Cloud-Modell.
Optimiert für bessere Lesbarkeit, Fehlerbehandlung und Wiederverwendbarkeit.
"""

from pyrevit import revit, script
from System.IO import FileInfo
from System.Diagnostics import Stopwatch
from Autodesk.Revit.DB import ModelPathUtils
import os
import subprocess
import revitron
from collections import defaultdict

__title__ = 'Filesize'
__doc__ = 'Zeigt Dateigröße an, egal ob lokal oder Cloud-Modell.'


def get_local_file_size(path):
    """Gibt Dateigröße in Bytes zurück oder löst eine Exception aus."""
    fi = FileInfo(path)
    return fi.Length


def get_cloud_file_size(forge_project_id, item_urn, ps_script_path):
    """Ruft die Cloud-Modellgröße via PowerShell-Skript ab."""
    try:
        result = subprocess.check_output([
            "powershell", "-ExecutionPolicy", "Bypass",
            "-File", ps_script_path,
            forge_project_id, item_urn
        ], stderr=subprocess.STDOUT)
        # Entferne Zeilenumbrüche und konvertiere
        size_str = result.decode('utf-8').strip()
        if size_str and size_str.isdigit():
            return int(size_str)
    except subprocess.CalledProcessError as err:
        script.get_logger().error("PowerShell-Fehler: {err}")
    return None


def format_size(bytes_count):
    """Formatiert Bytes in lesbare MB-Form mit 2 Nachkommastellen."""
    mb = float(bytes_count) / (1024 ** 2)
    return "{mb:.2f} MB ({bytes_count} Bytes)"


def main():
    output = script.get_output()
    sw = Stopwatch()
    sw.Start()

    raw_path = revit.doc.PathName
    if not raw_path:
        output.print_md(
            "**Fehler:** Das Modell wurde noch nicht gespeichert. Bitte speichere es zuerst."
        )
        return

    # Versuche, User-Visible-Pfad zu ermitteln
    try:
        mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(raw_path)
        user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
    except Exception:
        user_path = raw_path

    size_bytes = None
    # Lokale Datei prüfen
    try:
        size_bytes = get_local_file_size(user_path)
    except Exception:
        # Kein lokaler Zugriff, versuche Cloud
        doc = revit.doc
        cloud_model_path = doc.GetCloudModelPath()
        if cloud_model_path:
            cfg = revitron.DocumentConfigStorage().get('revitron.history', defaultdict())
            forge_id = cfg.get('forgeprojectid', '').strip()
            item_urn = cfg.get('itemurn', '').strip()
            ps_path = os.path.join(os.path.dirname(__file__), 'get_model_size.ps1')

            # Prüfung, ob Konfigurationswerte vorhanden sind
            if not forge_id or not item_urn:
                output.print_md(
                    "**Fehler:** Forge Project ID oder Item URN nicht konfiguriert."
                )
                return
            if not os.path.isfile(ps_path):
                output.print_md(
                    "**Fehler:** PowerShell-Skript nicht gefunden: {ps_path}"
                )
                return

            size_bytes = get_cloud_file_size(forge_id, item_urn, ps_path)
            if size_bytes is None:
                output.print_md("**Fehler:** Konnte Cloud-Modellgröße nicht ermitteln.")
                return
        else:
            output.print_md("**Hinweis:** Kein Cloud-Modell erkannt und lokale Datei nicht zugänglich.")
            return
    size_mb     = float(size_bytes) / (1024**2)
    # Ausgabe
    #output.print_md("**Dateigröße des aktuellen Modells**")
    #output.print_md("- **Größe**: **{format_size(size_bytes)}**")
    output.print_md("**Dateigröße des aktuellen Modells**")
    output.print_md("- **Größe**: **{0:.2f} MB** ({1} Bytes)".format(size_mb, size_bytes))
    # Laufzeit
    sw.Stop()
    output.print_md("Scriptlaufzeit: {0}".format(sw.Elapsed))
    script.exit()


if __name__ == '__main__':
    main()
