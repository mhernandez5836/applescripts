# Convert Numbers to Excel

-- 1. SETUP PATHS
-- Make sure these paths end with a colon ":"
set sourceFolder to "Macintosh HD:Users:path:to:file"
set outputFolder to "Macintosh HD:Users:path:to:export:location"

tell application "System Events"
	-- Check if folders exist
	if not (exists folder sourceFolder) then
		display dialog "Source folder not found: " & return & sourceFolder buttons {"Cancel"} default button "Cancel"
		return
	end if
	if not (exists folder outputFolder) then
		display dialog "Output folder not found: " & return & outputFolder buttons {"Cancel"} default button "Cancel"
		return
	end if
	
	-- 2. GET FILES (NATIVELY)
	-- System Events handles Chinese characters perfectly here
	set filesToProcess to path of every file of folder sourceFolder whose name extension is "numbers"
end tell

if filesToProcess is {} then
	display dialog "No .numbers files found in source folder." buttons {"OK"} default button "OK"
	return
end if

-- 3. PROCESS FILES
tell application "Numbers" to activate

repeat with currentFile in filesToProcess
	
	-- Get the file name properly using Finder/System logic logic
	tell application "System Events"
		set fileAlias to currentFile as alias
		set originalName to name of fileAlias
		set fileExtension to name extension of fileAlias
	end tell
	
	-- Remove extension safely using text manipulation
	set baseName to text 1 thru -((length of fileExtension) + 2) of originalName
	set newFileName to baseName & ".xlsx"
	set finalHFSPath to outputFolder & newFileName
	
	tell application "Numbers"
		try
			-- Open the file
			set theDoc to open fileAlias
			
			-- Export
			export theDoc to file finalHFSPath as Microsoft Excel
			
			-- Close
			close theDoc saving no
			
		on error errMsg
			display notification "Error on: " & originalName subtitle errMsg
		end try
	end tell
end repeat

tell application "Numbers" to quit
display dialog "Conversion Complete!" buttons {"OK"} default button "OK"
