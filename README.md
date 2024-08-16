# Summer-2024-Internship-Macros
This repository contains Visual Basic macros that I created to automate certain processes in Excel during my summer 2024 manufacturing engineering internship.

The following files have been included:

**SetupSheetParser.xlsm:** This macro utilizes if statements, loops, and functions such as VLOOKUP, IFERROR, TRIM, and INSTR to lookup the 5-character alphanumeric tool locations in a preset column in the external database "MasterToolingList.xlsx" and add the tool number if found. If the tool number is not found it will print "not found". The macro also indicates in the adjacent column whether the tool is obsolete or has been moved to a new location. The macro and MasterToolingList.xlsx must be simultaneously open for full functionality. Unfortunately, I was not able to obtain the master tooling list from my internship company for security reasons, but it can be adjusted to work with other databases.

**Sort_by_Job.xlsm:** This macro utilizes recursion, FileSystemObjects, and functions such as TRIM, RIGHT, INSTR, INSTRREV, 
