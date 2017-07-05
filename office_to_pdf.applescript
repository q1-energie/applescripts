on open droppedItems
	
	set theOutputFolder to choose folder with prompt "Please select an output folder:"
	
	repeat with a from 1 to count of droppedItems
		set current to item a of droppedItems
		tell application "Finder"
			set was_hidden to extension hidden of current
			set extension hidden of current to true
			set currentWithout to displayed name of current
			set ext to name extension of current
			set extension hidden of current to was_hidden
		end tell
		set target to ((POSIX path of theOutputFolder) as string) & "/" & currentWithout & ".pdf"
		if (ext = "doc") or (ext = "docx") then
			saveWordAsPDF(current, target)
		else if (ext = "xls") or (ext = "xlsx") then
			saveExcelAsPDF(current, target)
		end if
	end repeat
end open

on saveExcelAsPDF(documentPath, PDFPath)
	set tFile to (POSIX path of documentPath) as POSIX file
	tell application "Microsoft Excel"
		set isRun to running
		set wkbk1 to open workbook workbook file name tFile
		alias PDFPath
		save workbook as wkbk1 filename PDFPath file format PDF file format
		close wkbk1 saving no
		if not isRun then quit
	end tell
end saveExcelAsPDF

on saveWordAsPDF(documentPath, PDFPath)
	set tFile to (POSIX path of documentPath) as POSIX file
	tell application "Microsoft Word"
		set isRun to running
		open tFile
		set theDocument to active document
		alias PDFPath
		save as active document file name PDFPath file format format PDF
		close theDocument saving no
		if not isRun then quit
	end tell
end saveWordAsPDF