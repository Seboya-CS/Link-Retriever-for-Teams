# Link-Retriever-for-Teams
A short AutoHotkey script that converts an Excel table into a series of text files and provides a lookup function.

The script has these components:

Creates and writes to a log file for debugging and tracking errors

Creates and writes to one text file for each column in the data table

Loops through the columns and rows of the table in left-right top-down order, and copies the data from each column into an individual text file. This only needs to be done once, or as often as the user needs to update the data.

Allows the user to press a hotkey which will prompt the user for a group ID, and then it finds the matching Teams link.

The user may choose to save the link to the clipboard.


Instructions:

Download AutoHotkey V2, if not already on the machine: https://www.autohotkey.com/.

If using the portable version of AutoHotkey, download and extract the .zip: https://www.autohotkey.com/download/

In this repository are two versions of the script. The first version, "Retriever.ahk" allows for command-line arguments to specify the directories and file names of the various components used by the script. It also allows to specify the start column and row in the data table, the path to the Excel spreadsheet, and the worksheet name.

The second version,  "Retriever-no-cla.ahk" is the same script but the command line arguments are replaced with default options. The working directory is the local directory where the script exists. The user only needs to define the path to the Excel worksheet, which can be done by editing the .ahk document and inputting the path in line 4 where it says `xlPath := "replace"`. The path should be encompassed by double quotes.

Copy whichever version you prefer to a plain text file, then save it on your machine with an .ahk extension. To use, simply double-click the .ahk file, and if the machine prompts you for which program to use to open the file, find the AutoHotkey64.exe and select that.

When the script launches, it will set up the working environment either in the path defined by command-line argument, or in the directory where the script is located. The user must have write access to do so, or the script will exit.

There are two hotkeys used by the script. `Alt Shift S` directs the script to loop through the Excel file and copy each column into a text file. If the script encounters a problem, it will display a dialogue window and request the user's input. This only needs to be done once, to build the text files, but it can be done as often as the user needs in case the data changes.

When updating text files that already exist, the script will prompt the user if the user would like to overwrite the text files, or move them to an archive directory. The default option is to archive the files. This will be replaced with a static settings option in a future update.

`Alt Shift L` calls the lookup function, which prompts the user for a Group ID, then retrieves the link associated with that group ID, and displays it in a dialogue window with a button to copy it to clipboard.

The script will remain idle on the machine until it is closed by the user. It can be closed by right-clicking the icon in the system tray and selecting "Exit".


Command-line arguments:

If you want to make a batch file, the structure of the file is located in `/batch/`

All strings should be encompassed by double quotes. One or more spaces should separate each parameter. The first argument is a comma-separated list of relative paths (or absolute if storing them outside of workingDir) to the text files that will be used by the script. Do not include a slash at the beginning of a relative path. The text files do not need to exist, the script will create them. Number values do not need to be quoted.


To-dos:

Standardize user-input relative paths so it doesn't matter if it includes a leading slash or not.

Expand the lookup functionality

Add reminder to update links once per day

Add a GUI with settings options

Replace the text files with a single CSV

There's a hardcoded `11` value used in the script that needs to be replaced with a relative value.

Remove/generalize the parts of the script that are specific to my use, generalize the functions, and separate the functions into individual AHKs.

When the user retrieves a link, allow the option to have a persistent dialogue window that has the buttons "Copy to clipboard" "Copy and close" "Close". That way, if a link is needed frequently, the user does not need to keep on retrieving it.

Add the option to copy the group link + formatted group information like:

Group name: name

Group day: day

Group time: time

Group link: link
