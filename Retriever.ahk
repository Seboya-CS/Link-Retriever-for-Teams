#SingleInstance force
#Requires AutoHotkey >=2.0- <2.1

xlPath := A_Args[1]																; path to Excel spreadsheet
xlSheet := A_Args[2]															; worksheet name
logPath := A_Args[3] 															; path to log file
workingDir := CheckForSlash(A_Args[4])											; working directory
startRow := A_Args[5]															; start row of Excel table
startCol := A_Args[6]															; start col of Excel table
colFiles := A_Args[7]															; paths to text files, comma-separated
archiveDir := CheckForSlash(A_Args[8]) 											; path to archive directory


; +++++		Error codes		+++++++++++++++++++++++++++++++++++++++++++++++++++;
; 000 - Insufficient arguments												++;+
; 001 - Error accessing log file											++;+
; 002 - User does not have write access to log file							++;+
; 003 - Error creating archive directory									++;+
; 004 - Error moving the log file											++;+
; 005 - Error moving the column text files									++;+
; 006 - Problem getting a valid group ID from the input box					++;+
; 007 - Error retrieving the link, probably invalid group ID				++;+
; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++;


; =============================================================================
; ============		Initialization		=======================================
; =============================================================================

if A_Args.capacity < 8 {														; ensure all arguments are included
    MsgBox(ErrorCodes(000))
    ExitApp
}

SetWorkingDir workingDir														; set working directory

noLogBool := false																; initialize global variables
logOpenBool := false
firstLogOpenBool := true
noArchiveBool := false
notifyID := ""
logFile := ""

acol := StrSplit(colFiles, ",")													; convert list to array

; =============================================================================
; ============		Initialize log		=======================================
; =============================================================================
; This procedure checks to make sure the log is accessible and in a directory
; in which the user has write access. If a problem occurs, the user is prompted
; for input.

InitializeLog() {
	global noLogBool, logOpenBool, firstLogOpenBool, notifyID, logFile, 
	global logPath, workingDir, startRow, startCol, colFiles, archiveDir, acol	
	global xlPath, xlSheet, noArchiveBool
	try {																	
        logFile := FileOpen(logPath, "a")										; create file object
        if (!logFile) {
            throw Error("Unable to open log file", -1)							; if fails, throw error
        }
    } catch Error as err {
        if (MsgBox(ErrorCodes(001), , "YN") = "YES") {							; user prompted to try to create a new
			attribute := FileGetAttrib(logPath)										; log file, or proceed without
			if (!InStr(attribute, "W")) {										; if logPath does not have write access
				MsgBox(ErrorCodes(002))												; exit script
				ExitApp
			}
			try {
				archiveDir := CreateNewArchiveFolder()							; create a new archive to move the old
				FileMove(logPath, archiveDir)										; log into
			} catch Error as err {
				if (MsgBox(ErrorCodes(004), , "YN") = "Yes") {					; user prompted to exit app or continue
					ExitApp															; without logging
				} else {
					noLogBool := true											; global boolean for when proceeding
					return 0														; without logs
				}
			}
		} else {
			noLogBool := true
			return 0
		}
	}
    if firstLogOpenBool {
        logFile.Write("Log initialization success: " FormatTime(A_Now, 			; write some information to log
						"yyyy-MM-dd hh:mm:ss") "`r`n")
        firstLogOpenBool := false
    }
	logFile.close																; end function
    return true
}

; =============================================================================
; ============		Get worksheet info		===================================
; =============================================================================
; This procedure loops through the rows and columns in the Excel table. It loops
; left-right top-down. Each column is stored into a single string, and that
; string is written to the text file. Each column in the table is represented
; by a single text file. The row number is added to the left side of the string
; within each row, separated from the actual content with a single colon.
; This procedure also identifies endCol and endRow, which are undefined prior
; to the first time the script completes this procedure.
; I chose to use text files because:
; 1) It does not need to hold any information in memory
; 2) It's faster than opening and closing a new Excel instance every time
; 3) It's easier for me to write code using this approach compared
; 	 to using a CSV file. I'm going to switch this over to use a CSV in a 
;	 future update, since that is more efficient.

!+s:: { 
	global noLogBool, logOpenBool, firstLogOpenBool, notifyID, logFile, 
	global logPath, workingDir, startRow, startCol, colFiles, archiveDir, acol	
	global xlPath, xlSheet, noArchiveBool
	
	if (!MoveOrOverwriteFiles) {												; handle column text files if they
		WriteToLog("Error code 005. Writing over column files.")					; currently exist
	}
										   										; initialize Excel COM object
    xl := ComObject("Excel.Application")										; note: `ComObject("Excel.Application")`
    xl.Visible := False															; is used to initialize a new instance
    wb := xl.workbooks.open(xlpath)												; of Excel, which is what we want in
	ws := wb.worksheets(xlSheet)												; this case. for attaching to an 
    ws.activate																	; existing Excel instance, use
																				; `ComObjActive("Excel.Application")`
	
	colBool := true																; initialize some vars
	endRow := 0
	c := 0
	
    While (colBool) {															; loop columns
		c++
        col := startCol + c - 1													; standardize the column index
        columnData := ""														; reset columnData
		
		rowBool := true
		
        while (rowBool) {														; loop rows
            row := startRow + A_Index - 1										; standardize row index
            cellValue := ws.Cells(row, col).Value								; get cell value
			
			if (col = startCol && row != startRow) {
				finalStr := CheckForTypos(cellValue)							; remove typos from group IDs
			} else {
				finalStr := cellValue
			}
			
			if (endRow = 0 && ws.Cells(row + 1, col).Value = "") {				; if the cell in the next row is empty
				endRow := row														; then we are at the last row
				rowBool := false
			} else if (endRow > 0 && row = endRow) {							; after endRow is defined, simply test
				rowBool := false													; if row = endRow
			}
			
			if (rowBool) {														; if we are at endRow, no need to add
				columnData := columnData . row . ":" . finalStr . "`n"				; a new line
			} else {
				columnData := columnData . row . ":" . finalStr
			}
		}

        columnFile := FileOpen(acol[c], "w")									; the paths to the files are contained 
		columnFile.Write(columnData)												; in acol. 
		columnFile.Close														; if the file exists, it is overwritten.
		
		if (ws.Cells(startRow, col + 1).value = "") {							; if the cell in the next column is
			colBool := false														; empty then we are in the last
			endCol := col															; column
		}
   }
    
    wb.Close()																	; clean up. close Excel workbook,
    xl.Quit()																		; quit Excel application
    xl := ""

	If (notifyID != "") {														; handle notifications for typos in
		logFile := FileOpen(logPath, "a")											; group IDs
		logFile.Write(notifyID . "`n")
		logFile.Close
		MsgBox(notifyID)
		notifyID := ""
	}
}																				; finished

return

; =============================================================================
; ===========		Retrieve the link		===================================
; =============================================================================
; This procedure prompts the user for a group ID, then uses RegEx to match
; with the row number associated with that group ID. Then, uses RegEx to match
; with the link. The procedure asks the user if they want to add the link to
; clipboard.

!+L:: {
	global noLogBool, logOpenBool, firstLogOpenBool, notifyID, logFile, 
	global logPath, workingDir, startRow, startCol, colFiles, archiveDir, acol	
	global xlPath, xlSheet, noArchiveBool
	groupNum := ""
	c := 0
	While (!IsNumber(groupNum) &&  c <= 10) {									; we give the user 10 opportunities
		c++																			; to type a valid integer
		If (c < 3) {
			response := InputBox("Input `"-1`" to exit script. Group ID:")		; I chose to use "-1" as the exit cue
		} else {																	; in case someone accidentally hits
			response := InputBox("Enter the group ID using only numbers."			; cancel, the script proceeds
								" Non-numeric entries will be discarded.")
		}
		if (response.value = "-1") {
			ExitApp
		}
		groupNum := response.value												; store into convenient variable
	}
	If (c >= 10 && IsNumber(groupNum) = false) {								; if the user failed to input a valid
		MsgBox(ErrorCodes(006))														; integer, exit
		WriteToLog("Error code 006. Failed to acquire valid group ID")
		Exit
	}

	groups := FileRead(acol[1])													; read from group IDs file
	links := FileRead(acol[11])													; read from links file
	
; this matches one to three digits that are followed by (and not including in 
; the match) a colon and the group number. also, that entire string must be a 
; single word (no white space).
; this results in a match with just the row number
	pattern := "\b\d{1,3}(?=:" groupNum "\b)"									
	
	groupsPos := RegExMatch(groups, pattern, &match)							; perform the ReGex
	lineNum := match[0]															; match[0] is the actual matched string
	
; this pattern has two parts. left part:
; [new line] followed by [the row number that contains the data we are
; retrieving] followed by ":"
; right part:
; "https:" followed by [any number of any type of character except new lines]
; followed by (but not including in the match) a new line
; the right string must follow the left string, and the match does not include
; the left string in the match. This results in only getting the link.
	pattern := "\R" lineNum ":\Khttps:(.*)(?=\R)"
	linksPos := RegExMatch(links, pattern, &match2)								; perform ReGex again
	
	try {
		thelink := match2[0]													; this is in a try-catch to account for
	} catch Error as err {															; invalid group IDs
		MsgBox(ErrorCodes(007))
		WriteToLog("Error code 007. Failed to retrieve group ID: " lineNum)
		return 0
	}
	if (MsgBox(match2[0] "`r`n`r`nCopy link to clipboard?", , "YN") = "Yes") {	; prompt user if they want to add
		A_Clipboard := thelink														; the link to clipboard
	}
	return thelink
}

; =============================================================================
; ============		Move or overwrite files		===============================
; =============================================================================
; This procedure checks if the column text files already exist. If so, the
; procedure prompts the user whether to overwrite them, or archive them.

MoveOrOverwriteFiles() {
	global noLogBool, logOpenBool, firstLogOpenBool, notifyID, logFile, 
	global logPath, workingDir, startRow, startCol, colFiles, archiveDir, acol	
	global xlPath, xlSheet, noArchiveBool
	fileExistsBool := false
	writeOverBool := false
	fileLen := 0
	num := 0

	While (fileExistsBool = false && num <= 11) {								; loop through all of the column files
		num := A_Index
		try {
			fileText := FileRead(acol[num])
			fileLen := StrLen(fileText)
		} catch Error as err {													; no need for action if error
		}																
		If (fileLen >= 1) {														; if any files contain data, then they
			fileExistsBool := true													; exist. I chose this approach
		}																			; to account for empty files. No
	}																				; need to archive empty files.
	If (fileExistsBool) {
		response := InputBox('Submit "1" to write over the current files.'		; I chose input box so it can default to
					' Submit "2" to move the current files to the archive'			; move the files. Eventually I will
					' directory. Submit "0" to cancel and end the script.'			; replace this with a static
					, , , "2")														; settings option.
		If (response.value = "0") {
			ExitApp
		} else if (response.value = "2") {
			newDirPath := CreateNewArchiveFolder()
			for file in acol {
				try {
					FileMove(file, newDirPath)									; move the files
				} catch Error as err {
					if (MsgBox(ErrorCodes(005), , "YN") = "Yes") {				; if error, user is prompted to either
						ExitApp														; end the script or write over
					} else {														; the files.
						return false
					}
				}
			}				
		}
	}
	return true
}

; =============================================================================
; =============		Miscellaneous functions		===============================
; =============================================================================

; This procedure handles the notifications for any instances of typos in a 
; group ID
UpdateNotifyMsg(uChar, uRow, uUnchangedID)
{
	global notifyID
	If (notifyID= "") {
		notifyID :=  uChar . " was removed from row " . uRow . " . Group ID: " 
					. uUnchangedID
	} else {
		notifyID := (notifyID . "`n" . uChar . " was removed from row " . uRow 
		. ". Group ID: " . uUnchangedID)
	}
	return
}

; This procedure simply standardizes user-input file paths to always have a 
; slash at the end.
CheckForSlash(path)
{
	strL := StrLen(path)
	lastChar := SubStr(path, strL, 1)
	if (lastChar = "\") {
		return(path)
	} else {
		return(path . "\")
	}
}

; This procedure attempts to create a new archive directory. If there is a
; problem, it prompts the user for direction.
CreateNewArchiveFolder() {
	global noLogBool, logOpenBool, firstLogOpenBool, notifyID, logFile, 
	global logPath, workingDir, startRow, startCol, colFiles, archiveDir, acol	
	global xlPath, xlSheet, noArchiveBool
	
	success := false

	newDirName := "Archive-files-" FormatTime(A_Now, "yyyy-MM-dd hh mm")
	if (archiveDir = "\") {														; if the user did not pass a file path
		newDirPath := workingDir . "Archive\" . newDirName							; for the archive directory as an 
	} else {																		; argument, create one
		newDirPath := archiveDirectory . "Archive\" . newDirName
	}
	try {
		DirCreate newDirPath													; attempt to create the new directory
	} catch Error as err {
		if (MsgBox(ErrorCodes(003), , "YN") = "Yes") {							; prompt the user if they want to 
			ExitApp																	; exit the script or proceed
		} else {																	; without archives
			noArchiveBool := true
			WriteToLog("Error code 003. Proceeding without archives.")
		}
	}
	if (FileExist(newDirPath)) {
		return newDirPath
	} else {
		return false
	}
}

; This procedure handles writing to the log file.
WriteToLog(str) {
	if (!noLogBool) {															; if noLogBool = false
		logFile := FileOpen(logPath, "a")										; open log
		if (logFile.Pos > 0) {													; if log is not empty
			logFile.Write("`r`n" str)											; append with new line
		} else {
			logFile.Write(str)													; else, don't use a new line
		}
		logFile.Close
		return 1
	}
}

; This function checks the group IDs for erroneous characters, which tend to
; appear quite frequently. This function removes the errors and calls the
; procedure UpdateNotifyMsg so it can be logged. When the "Get worksheet info"
; procedure concludes, a Message Box will display any alterations that were made
; to group IDs on the spreadsheet.
CheckForTypos(str) {
	firstChar := SubStr(str, 1, 1)												; get first char of string
	strL := StrLen(str)															; get string length
	lastChar := SubStr(str, strL, 1)											; get last char of string
	if (IsNumber(firstChar) = false) {											; if either the first or last chars
		tempStr := SubStr(str, 2, strL - 1)											; are not numbers, then they are
		strL := StrLen(tempStr)														; removed
		UpdateNotifyMsg(firstChar, row, str)
	} else {
		tempStr := str
	}
	if (IsNumber(lastChar) = false) {
		finalStr := SubStr(tempStr, 1, strL - 1)
		updateNotifyMsg(lastChar, row, str)
	} else {
		finalStr := tempStr
	}
	return finalStr
}

; This procedure contains the error messages.
ErrorCodes(n) {
	Switch n {
		Case 000:
			str := (
				"E000: Insufficient arguments provided. Please provide the following in this order:`r`n1) xlPath - path"
				" to the Excel spreadsheet.`r`n2) xlSheet - name of worksheet that contains the table.`r`n3) logPath -"
				" path to the log txt file.`r`n4) workingDir - current working directory. The working directory should"
				" contain all auxiliary components used by the Retriever.`r`n5) startRow - the first row containing"
				" data to be used by the Retriever.`r`n6) startCol - the first column containing data to be used by the"
				" Retriever.`r`n7) colFiles - comma-separated list of paths to the text files, one for each data"
				" column.`r`n8) archiveDir - directory to archive logs and column files, if desired. Set to `"`" and"
				" the script will create a default archive directory.")
		Case 001:
			str := ("E001: There was an error accessing the log. Select `"Yes`" to attempt to create a new log file."
					" Select `"No`" to have the script proceed without logging. This may result in"
					" lost log information and/or lost Teams spreadsheet data.")
		Case 002:
			str := "E002: You do not have write access to the log path at this time. Exiting the script."
		Case 003:
			str := ("E003: There was an error creating the archive directory. Select `"Yes`" to cancel and end the"
					" script. Select `"No`" to have the script proceed without an archive. This may result in"
					" lost log information and/or lost Teams spreadsheet data.")
		Case 004:
			str := ("E004: There was an error when moving the log file. Select `"Yes`" to cancel and end the"
					" script. Select `"No`" to have the script proceed without logging. This may result in"
					" lost log information and/or lost Teams spreadsheet data.")
		Case 005:
			str := ("E005: There was an error moving the text files. Select `"Yes`" to cancel and end the"
					" script. Select `"No`" to write over the text files.")
		Case 006:
			str := ("E006: There was a problem getting the group ID from the input box. The script will"
					" now abort.")
		Case 007:
			str := ("There was an error when retrieving the link. Please check the group number"
					" and try again.")
	}
	return str
}
