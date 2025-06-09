# MEcode

This repository contains sample Excel VBA code.

## VBA Macro

`FileSearchCopy.bas` defines macros to search for files, copy them to a destination folder and list the copied files in the active worksheet. A temporary form lets you choose the file formats to search via check boxes.

### How to use
1. Import `FileSearchCopy.bas` into your Excel workbook (via `File` > `Import File` in the VBA editor).
2. Run the `AddSearchButton` macro to place a button in cell A2.
3. Click the inserted button and follow the prompts to search and copy files. When the form appears, check the file extensions you want to include and press **確定**.
