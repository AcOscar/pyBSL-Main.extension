# AI Agent Guidance for pyBSL-Main.extension

## What this repository is
- A `pyRevit` add-in extension for Autodesk Revit.
- The extension manifest is `extension.json`.
- User interface panels live under `pyBSL.tab/`.
- Each tool is a `*.pushbutton` directory containing a Python script (`script.py` or `*_script.py`) and optional supporting files.

## Primary languages and frameworks
- Python targeting the `pyRevit` / IronPython environment.
- Uses `pyrevit`, `Autodesk.Revit.DB`, `Autodesk.Revit.UI`, `revitron`, and WPF/XAML for UI.
- Some buttons contain helper libraries under nested `lib/` folders.

## Key conventions
- Do not reorganize the `pyBSL.tab/<panel>.panel/<button>.pushbutton/` directory structure without preserving the `pyRevit` button layout.
- Scripts are executed by pyRevit, not by a standard Python package manager.
- There is no existing build/test automation in the repo; changes should be validated in the Revit/pyRevit environment.
- Prefer minimal edits and preserve existing button scripts unless asked to refactor a specific tool.

## Useful starting points
- `extension.json` — extension metadata and enablement settings.
- `pyBSL.tab/` — main UI panel tree.
- `pyBSL.tab/BSL.panel/ShowRoom.pushbutton/ShowRooms_script.py` — representative sample of script style and Revit API usage.

## What not to do
- Do not introduce standard Python packaging files such as `setup.py`, `pyproject.toml`, or `requirements.txt` unless the user explicitly requests packaging support.
- Do not assume a modern CPython environment; keep compatibility with `pyRevit` and IronPython conventions.
- Do not move or rename `*.pushbutton` folders absent a clear extension author intent.

## Recommended next customization
- Add a `copilot-instructions.md` or `.github/copilot-instructions.md` if deeper repository-specific guidance is needed for prompt behavior, code styling, or Revit-specific patterns.
- Create a skill for `pyRevit` extension maintenance if the repository grows or if agents need to automate button creation and script scaffolding.
