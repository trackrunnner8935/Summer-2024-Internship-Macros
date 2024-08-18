# Summer-2024-Internship-Macros
This repository contains Visual Basic macros that I created to automate certain processes in Excel during my summer 2024 manufacturing engineering internship.

The following files have been included:

**SetupSheetParser.xlsm:** This macro utilizes if statements, loops, and functions such as VLOOKUP, IFERROR, TRIM, and INSTR to iterate through 5-character alphanumeric tool locations in a column in a user-selected setup sheet, lookup the locations in the external database "MasterToolingList.xlsx", and add a forward slash and the matching tool number if found. If the tool number is not found it will print "not found". The macro also indicates in the adjacent column whether the tool is obsolete or has been moved to a new location. The macro and MasterToolingList.xlsx must be simultaneously open for full functionality. Unfortunately, I was not able to obtain the master tooling list from my internship company for security reasons, but it can be adjusted to work with other databases.

**Sort_by_Job.xlsm:** This macro utilizes recursion to loop through files in a user-selected directory and FileSystemObjects, and functions such as TRIM, RIGHT, INSTR, and INSTRREV sort files into folders based on the suffix of the filename (without the file extension) following the first space in the filename. (i.e. S24 ACUTE SCREW and S45 ACUTE SCREW would both be moved to the folder ACUTE SCREW.) This can be adjusted to sort by a prefix or suffix with different delimiters.

**S69 QUARTEX BODY POLYAXIAL SCREW test.xlsx and S69 QUARTEX BODY POLYAXIAL SCREW test 2.xlsx:** These are test files for the previous two macros.


