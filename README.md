# Link-Retriever-for-Teams
A short AutoHotkey script that converts an Excel table into a series of text files and provides a lookup function.

The script has these components:

Creates and writes to a log file for debugging and tracking errors

Creates and writes to one text file for each column in the data table

Loops through the columns and rows of the table in left-right top-down order, and copies the data from each column into an individual text file. This only needs to be done once, or as often as the user needs to update the data.

Allows the user to press a hotkey which will prompt the user for a group ID, and then it finds the matching Teams link.

The user may choose to save the link to the clipboard.


To-dos:

Expand the lookup functionality

Add a GUI with settings features

Replace the text files with a single CSV

Remove/generalize the parts of the script that are specific to my use, generalize the functions, and separate the functions into individual AHKs.
