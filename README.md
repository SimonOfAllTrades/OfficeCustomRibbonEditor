# OfficeCustomRibbonEditor
PowerShell Script that allows addition/removal/changes to custom ribbons in Office applications

This PS script makes editing custom ribbons in Excel, Word, and Powerpoint easy.
Users may add and remove custom ribbons as well as change pre-existing ones.  The ribbon hosts macro shortcuts that make running them much more reliable and easy to nvaigate to.

This PS script works since Office applications are actually zipped folders wrapped with an executable.  Using Excel, for example, contains all the data for a workbook in directories in the zipped file.  Custom ribbons can be added by adding an additional folder as well as XML styling for the ribbon.  A few edits are needed in the reference file so that the ribbon is properly linked.  Afterwards, the zipped file is reanmed to .xlsm to re-wrap the zip with an executable.
