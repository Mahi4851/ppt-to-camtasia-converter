# PowerPoint to editable Camtasia converter

Windows desktop application for converting PowerPoint (`.pptx`) presentations into an editable Camtasia project (`.tscproj`) with a graphical file picker.

## Features

- Browse and select a PowerPoint file using the desktop UI.
- Generate a Camtasia project with one group per slide.
- Export non-text visual elements as individual PNG layers where supported.
- Create editable text layers and preserve animation sequencing where possible.
- Save Camtasia project files and accompanying reports beside the selected PPTX.

## Requirements

- Windows
- Python 3.10 or newer
- Microsoft PowerPoint installed
- Camtasia installed to open the output project

## Quick start

1. Download or clone this repository.
2. Double-click `run_all.bat` to install the Python dependencies and open the UI.
3. Select a `.pptx` file using **Choose PPT**.
4. Start conversion and check the output beside the selected PowerPoint.

You can also install packages manually with `py -m pip install -r requirements.txt` and start the UI with `py scripts\ui_launcher.py`.

## Project files

- `scripts/ui_launcher.py` – file-picker interface and conversion launcher.
- `scripts/proper_one_click.py` – PowerPoint-to-Camtasia conversion engine.
- `run_all.bat` – Windows startup script.
- `requirements.txt` – Python dependencies.
- `README_FIRST.txt` – original setup instructions and conversion behavior.
- `input/README_INPUT.txt` – original input-folder notes.

## Notes and limitations

- PowerPoint COM automation requires Microsoft PowerPoint on Windows.
- Complex PowerPoint effects may need manual refinement in Camtasia.
- This repository intentionally **does not include** the original uploaded sample presentations, generated media, working files, or prebuilt Camtasia output. Provide your own PPTX.
- Only Python syntax was checked during packaging; end-to-end conversion was not tested here.

## Privacy

Before pushing changes, check that no confidential presentation files, generated output, credentials, or client assets are tracked by Git.
