# Summer-2024-Internship-Macros
This repository contains the excel macros that I created to automate certain processes during my summer 2024 manufacturing engineering internship.

The following macros have been included:

**SetupSheetParser:** This macro utilizes if statements, loops, and Excel VBA functions such as VLOOKUP, IFERROR, TRIM, and INSTR to lookup the 5-character alphanumeric tool locations in a preset column in the external database "MasterToolingList.xlsx" and add the tool number if found. If the tool number is not found it will print "not found". The macro also indicates in the adjacent column whether the tool is obsolete or has been moved to a new location. The macro and MasterToolingList.xlsx must be simultaneously open for full functionality. Unfortunately, I was not able to obtain the master tooling list for  
