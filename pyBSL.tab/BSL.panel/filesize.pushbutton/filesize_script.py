# -*- coding: utf-8 -*-

from pyrevit import revit, script
from System.IO import FileInfo
from System.Diagnostics import Stopwatch
from Autodesk.Revit.DB import ModelPathUtils
import os
import subprocess
import revitron
from collections import defaultdict

__title__ = 'Filesize'
__doc__ = 'Displays file size, regardless of whether it is local or cloud-based.'


def get_local_file_size(path):
    #Returns the file size in bytes or throws an exception.
    fi = FileInfo(path)
    return fi.Length


def get_cloud_file_size(forge_project_id, item_urn, ps_script_path):
    #Retrieves the cloud model size via PowerShell script.
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
        script.get_logger().error("PowerShell-Error: {err}")
    return None


def format_size(bytes_count):
    #Formats bytes into readable MB format with 2 decimal places.
    mb = float(bytes_count) / (1024 ** 2)
    return "{mb:.2f} MB ({bytes_count} Bytes)"


def main():
    output = script.get_output()
    sw = Stopwatch()
    sw.Start()

    raw_path = revit.doc.PathName
    if not raw_path:
        output.print_md(
            "**Error:** The model has not yet been saved. Please save it first."
        )
        return

    # Attempts to determine user-visible path
    try:
        mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(raw_path)
        user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
    except Exception:
        user_path = raw_path

    size_bytes = None
    # Check local file
    try:
        size_bytes = get_local_file_size(user_path)
    except Exception:
        # No local access, try cloud
        doc = revit.doc
        cloud_model_path = doc.GetCloudModelPath()
        if cloud_model_path:
            cfg = revitron.DocumentConfigStorage().get('revitron.history', defaultdict())
            forge_id = cfg.get('forgeprojectid', '').strip()
            item_urn = cfg.get('itemurn', '').strip()
            ps_path = os.path.join(os.path.dirname(__file__), 'get_model_size.ps1')

            # Check whether configuration values exist
            if not forge_id or not item_urn:
                output.print_md(
                    "**Error:** Forge Project ID or Item URN not configured."
                )
                return
            if not os.path.isfile(ps_path):
                output.print_md(
                    "**Error:** PowerShell script not found: {ps_path}"
                )
                return

            size_bytes = get_cloud_file_size(forge_id, item_urn, ps_path)
            if size_bytes is None:
                output.print_md("**Error:** Unable to determine cloud model size.")
                return
        else:
            output.print_md("**Note:** No cloud model detected and local file not accessible.")
            return
    size_mb     = float(size_bytes) / (1024**2)

    output.print_md("**File size of the current model**")
    output.print_md("- **Size**: **{0:.2f} MB** ({1} Bytes)".format(size_mb, size_bytes))

    sw.Stop()
    output.print_md("Script runtime: {0}".format(sw.Elapsed))
    script.exit()


if __name__ == '__main__':
    main()
