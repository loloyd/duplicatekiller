# Duplicate Killer #

Duplicate Killer is an open-source byte-by-byte file comparison with match detection and duplicate file deletion features, developed in Visual Basic 6 by LoloyD of LD TechnoLogics http://loloyd.com/ and available at http://github.io/loloyd/duplicatekiller/.  Design considerations are on the most minimal requirements for VB6 as possible, hence the absence of explicit file pickers and the utilization of drag-drop functionalities.

It can run in Linux environments with WineTricks installed and appropriate libraries, particularly the vb6run (MS Visual Basic 7 runtime sp6) package, duly installed.

## Quick Usage Instructions ##

1. Drag and drop possible duplicate files into the "Target" list box.  Files will be checked here for duplicate matches against files found in the "Source" list box once the "Detect Duplicates" command button has been summoned.

1a. To remove listed files (delist) from the *"Target"* list box, make a selection of the desired files first then click on the Delete keyboard button.

2. Drag and drop potential comparison files into the *"Source"* list box.  Files listed here will serve as reference files for comparison only once the *"Detect Duplicates"* command button has been summoned.

2a. To remove listed files (delist) from the *"Source"* list box, make a selection of the desired files first then click on the Delete keyboard button.

3. Mark the *"Kill detected duplicates"* checkbox if the deletion of detected duplicate files from the *"Target"* list box is desired.  When marked, the *"Detect Duplicates"* command button changes its caption to *"Detect Duplicates and Kill"*.  When unmarked, the *"Detect Duplicates"* command button changes its caption to *"Detect Duplicates Only"*.

4. Mark the *"Remove empty directories"* checkbox if the deletion of empty directories is desired.

5. Click the *"Detect Duplicates and Kill"* or *"Detect Duplicates Only"* to start the detection of duplicate files in a byte-by-byte comparison mode.

## Known Limitations ##

Operable files are limited to 2 GB each (Windows limitation).

## Disclaimer ##

LoloyD, Loloy D and LD TechnoLogics is not responsible nor liable for any misuse, abuse, accidental file deletion, file corruption, filesystem corruption, filesystem breakdown, machine breakdown, mechanical breakdown, logical breakdown, mental breakdown arising from the utilization of this open-source software application which has been distributed on an AS-IS, WHERE-IS basis and which source has been made fully appreciable and auditable at http://loloyd.com/ and http://github.io/loloyd/duplicatekiller/.  This disclaimer also applies to any eventual derivatives of this software application and its codebase.

## Licensing Distribution ##

GNU-GPL or its lesser derivatives apply, or whichever license closest to GNU-GPL applies to VB6 source code and compiled code.

## Revision Information ##

2018 March 30 - initial release by Loloy D