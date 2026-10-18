PPT to Proper Editable Camtasia Converter - UI Version
=======================================================

How to run
----------

1. Double-click:
   run_all.bat

2. A converter window opens.

3. Click Browse and select any .pptx file from any folder on your computer.
   You no longer need to copy the PowerPoint into this project's input folder.

4. Click Convert to Camtasia.

5. The finished output is created beside the PowerPoint you selected.

Example
-------

Selected PowerPoint:
C:\Courses\Module 01.pptx

Same-location Camtasia project:
C:\Courses\Module 01.tscproj

Self-contained project folder:
C:\Courses\Module 01_CAMTASIA_PROJECT\

Packaged Camtasia project inside that folder:
C:\Courses\Module 01_CAMTASIA_PROJECT\Module 01.tscproj

ZIP package:
C:\Courses\Module 01_CAMTASIA_PROJECT.zip

The .tscproj beside the PPT uses the exact same base filename as the selected PowerPoint.
It safely references the media in the companion project folder. The packaged project remains self-contained.

What has NOT changed
--------------------

The conversion/output engine is intentionally preserved. The converter still:

- Creates one expandable Camtasia group per slide.
- Creates clean PNG slide backgrounds without slide-level shapes where PowerPoint allows it.
- Exports all visible non-text PowerPoint shapes as individual PNG layers, including colored bars, rounded boxes, borders, cards, icons, pictures, SmartArt, charts, and grouped artwork.
- Keeps editable text above those visual layers.
- Recreates editable text as Open Sans, 34 pt.
- Reads PowerPoint animation target/order/delay data and maps it to Camtasia clip start times where possible.
- Adds a Fade transition to every imported timeline item.
- Places the slide background at the bottom of each group.
- Orders visual and editable-text tracks by PowerPoint spatial reading order, keeping related elements together.
- Writes slide, track, animation, timing, editability, and QC reports.
- Fails closed if any expected visual PNG layer is missing.
- Includes the source PPT copy and editing guides in the project package.

Only these workflow items changed
---------------------------------

- Input is selected through a UI file picker instead of the input folder.
- Output is saved beside the selected PowerPoint.
- A Camtasia .tscproj with the same base filename is created directly beside the selected PowerPoint.
- A second self-contained copy remains inside the companion project folder for packaging/portability.
- The output project folder and ZIP use the PowerPoint name instead of FINAL_CAMTASIA_PROJECT.

Requirements
------------

- Windows
- Python 3.10+
- Microsoft PowerPoint installed
- Camtasia installed to open the final project

Important limitation
--------------------

The same original limitation still applies: this runner prioritizes proper timeline organization, editable text, accurate visible shape layers, clean backgrounds, separate image/person layers, and animation sequencing. Some advanced PowerPoint effects may still need manual refinement in Camtasia after conversion.
