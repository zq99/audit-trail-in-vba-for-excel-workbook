# Audit Trail in VBA for an Excel Project

This is a simple class that can be added to an Excel VBA project, that can track user changes to the workbook.

The changes are exported to a separate text file.

## Setup

Add the clsLogger class to a VBA project.

The ThisWookbook.bas file is the code need for the workbook VBA module.

If handles the operation of calling the logger class for the different workbook events triggered by the users actions when have the workbook open.
