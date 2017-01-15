(*

(c) 2016 Jerod M. Price

*)

property dChar1 : ":;,'/|!@#$%^&*()_+=-" -- CleanName characters to be removed from file names
property dChar2 : " " -- CleanName sapce character to remove from file names
property rChar : "" -- CleanName reserved character to be removed from file names
property mydis : ":" -- myCleanURL
property myrep : "/" -- myCleanURL
property isOffice : {"xls", "xlsx", "doc", "docx", "ppt", "pptx", "pdf", "potx", "ttf", "rtf", "otf"} -- file suffix that are office documents
property isGraphic : {"jpg", "jpeg", "png", "tiff"} -- file suffix that are image documents
property oddballAttachment : "xxxx" -- reserved for special file types
property br : "<br>"
property td1 : "<td>"
property td2 : "</td>"
property tr1 : "<tr>"
property tr2 : "</tr>"
property table1 : "<table cellspacing=0 cellpadding=4 border=1 bordercolor=#F8F8BC bgcolor=#FFF8DC>"
property table2 : "</table>"
set firstROW to tr1 & td1 & "Document" & td2 & td1 & "Link" & td2 & td1 & "Size" & td2 & tr2
set myApp to (POSIX path of (path to applications folder as string) as string) & "Microsoft Outlook"

set nameAtt to "" as string
set dlClean to "" as string
set sizeAtt to 0.0
set extensionDigits to "" as string
set examination to false as boolean
set shoulddelete to false as boolean
set isdeleted to false as boolean
set wassaved to false as boolean
set anAttachment to {nameAtt, dlClean, sizeAtt, examination, shoulddelete, wassaved, isdeleted}
set attachmentList to {}
set selectedMessages to {}

-- a reference name for operating system notification messages
set ScriptTitle to "Outlook Attachment Removal"

-- make sure the Box.com sync folder is real and exists, create if not
set myHome to POSIX path of (path to home folder as string) as string
set myBox to "Box Sync" as string
set downloadPath to my checkForFolder(myHome, myBox) as string

-- make sure the attachment folder is real and exists, create if not
set myOLA to "OLAttachments" as string
set downloadFolder to myOLA & "-" & (year of (current date) as string)
set downloadString to my checkForFolder(downloadPath, downloadFolder) as string

-- main loop
tell application "Microsoft Outlook"
	
	-- get the currently selected message or messages
	set selectedMessages to selected objects
	
	-- if there are no messages selected, warn the user and then quit
	if selectedMessages is {} then
		tell application "Finder" to display notification "Please select a message first." with title ScriptTitle subtitle "Input not Available"
		return
	end if
	
	-- we one or more messaged handle the first and then work through each subsequent
	repeat with theMessage in selectedMessages
		
		if class of theMessage is incoming message then
			
			set attachmentList to {} -- empty array for housekeeping
			
			--get information from the message, and store it in variables
			set theName to subject of theMessage
			set fURL to exchange id of theMessage
			set theTime to time received of theMessage --set theCategory to category of theMessage
			set theShortTime to time string of theTime as string
			set theCleanTime to my CleanName(theShortTime) as string
			--set thePriority to priority of theMessage
			set theContent to content of theMessage
			
			-- configure some counters to help track iterations through various attachments and documents in an email
			set counterA to (count of (get attachments of theMessage))
			set counterB to 0
			set counterC to 0
			set counterD to 0
			set errorCount to 0
			set didWork to false
			set didPic to false
			repeat while ((counterA - counterB - counterC) > 0) and (counterD < (counterA * 2))
				
				-- counterD prevents runaway loop
				-- time to save and then delete the offensive attachments
				-- sort the array by isOffice at end, then delete from end forward
				-- end forward because the count is reset by MS OFFFICE when something changes
				-- if we started at the front, a failure will happen.
				
				try
					if (counterB = 0 and counterC = 0 and counterD = 0) then -- the first pass action
						set thisAttachment to first item in (get attachments of theMessage)
						set nameAtt to name of thisAttachment
						
						-- in future rev we need to make sure we don't accidentally clean away the suffix slash file type
						set cleannameAtt to my CleanName(nameAtt)
						
					else -- do the second and subsequent pass
						set thisAttachment to item (1 + counterC) in (get attachments of theMessage)
						set nameAtt to name of thisAttachment
						
						-- in future rev we need to make sure we don't accidentally clean away the suffix slash file type
						set cleannameAtt to my CleanName(nameAtt)
					end if
				end try
				
				if nameAtt is not in attachmentList then -- only handle attachments that we have not already handled, as outlook resets the attachment list array after one is deleted or modified we keep record of those touched in a list
					
					set sizeAtt to (my convertByteSize(((file size of thisAttachment) as integer), missing value, 2))
					set extensionDigits to my returnExtension(text -4 through -1 of nameAtt)
					
					if isOffice contains extensionDigits then -- do this action if the attachment of of type isOffice
						
						set didWork to true
						set theCleanAttachmentIdentifier to theCleanTime & "_" & cleannameAtt as string
						set dlClean to downloadString & theCleanAttachmentIdentifier as string
						
						--need to do this as a try just in case it didn't save
						--need to confirm it saved and check file exists so we don't delete if not saved
						save thisAttachment in file dlClean
						copy {nameAtt, dlClean, sizeAtt, true, true, false, false} to the end of attachmentList
						
						
						--after we are saved, then mark it as so in the list of files we are tracking
						set item 6 of the last item of attachmentList to true
						display notification "Saved " & nameAtt with title ScriptTitle subtitle "Success"
						--my updateFileWhereFromAttribute(dlClean, fURL)
						--need to do this as a try just in case it goes wrong
						delete thisAttachment
						
						
						set counterB to counterB + 1
						set counterD to counterD + 1
						set item 7 of the last item of attachmentList to true
						
						display notification "Deleted " & nameAtt & " from email " & theName with title ScriptTitle subtitle "Success"
						
						
					else if isGraphic contains extensionDigits then -- do this action if the attachment of of type isGraphic
						
						set didPic to true
						copy {nameAtt, "", sizeAtt, true, false, false, false} to the end of attachmentList
						set counterC to counterC + 1
						set counterD to counterD + 1
						
						
					else -- do this action when the attachment is of some other type
						copy {nameAtt, "", sizeAtt, true, false, false, false} to the end of attachmentList
						set counterC to counterC + 1
						set counterD to counterD + 1
						
						
					end if -- end attachment type was identified as Office, Graphic, or other
				end if -- end counters have run out
			end repeat -- end no more attachments to be found
			
			-- create a table of attachments
			-- format nicely with HTML
			-- pre-pend the table with the email content theContent
			
			
			-- need to do a case if didWork as a list, then case if didPic post some thumbs but push original to folder, then case other stuff...end
			
			
			if (didWork) then
				
				set HTMLList to ""
				repeat with a in attachmentList
					
					set deleted to (the seventh item in a)
					if deleted then
						set b to (the second item in a)
						
						set myWEBNAME to my myCleanURL(b)
						set myWEBNAME to "<a href='file://" & myWEBNAME & "'>here</a>"
						set HTMLList to HTMLList & tr1 & td1 & (the first item in a) & td2 & td1 & myWEBNAME & td2 & td1 & (the third item in a) & td2 & tr2
					end if
					
				end repeat
				
				set the category of theMessage to {category "Files Removed", category "OLA"}
				set HTMLList to table1 & firstROW & HTMLList & table2
				set content of theMessage to HTMLList & br & theContent
				
			else
				set the category of theMessage to {category "OLA"}
			end if
			
			
			set HTMLList to "" -- empty for housekeeping
			
			-- nearly there...do it again if more emails are selected
		end if -- end of incoming messages
	end repeat -- end because no more emails
	set attachmentList to {} -- empty array for housekeeping
	
end tell



-- use xattr to leave a link behind
on updateFileWhereFromAttribute(fpath, fURL)
	
	set myCommand to "xattr -w com.apple.metadata:kMDItemWhereFroms " & "outlook://" & fURL & " " & fpath
	do shell script myCommand
	
end updateFileWhereFromAttribute



(* Convert a size in bytes to a convenient larger unit size with suffix. The 'KBSize' parameter specifies the number of units in the next unit up (1024 or 1000; or 'missing value' for 1000 in Snow Leopard or later and 1024 otherwise). The 'decPlaces' parameter specifies to how many decimal places the result is to be rounded (but not padded). *)

on convertByteSize(byteSize, KBSize, decPlaces)
	
	if (KBSize is missing value) then set KBSize to 1000 + 24 * (((system attribute "sysv") < 4192) as integer)
	
	if (byteSize is 1) then
		set conversion to "1 byte" as Unicode text
	else if (byteSize < KBSize) then
		set conversion to (byteSize as Unicode text) & " bytes"
	else
		set conversion to "Oooh lots!" -- Default in case yottabytes isn't enough!
		set suffixes to {" K", " MB", " GB", " TB", " PB", " EB", " ZB", " YB"}
		set dpShift to ((10 ^ 0.5) ^ 2) * (10 ^ (decPlaces - 1)) -- (10 ^ decPlaces) convolutedly to try to shake out any floating-point errors.
		repeat with p from 1 to (count suffixes)
			if (byteSize < (KBSize ^ (p + 1))) then
				tell ((byteSize / (KBSize ^ p)) * dpShift) to set conversion to (((it div 0.5 - it div 1) / dpShift) as Unicode text) & item p of suffixes
				exit repeat
			end if
		end repeat
	end if
	
	return conversion
end convertByteSize

-- return the text right of the last decimal, aka file extention)
on returnExtension(aFileName)
	set dot to offset of "." in aFileName
	if dot > 0 then
		set theExtension to text (dot + 1) thru -1 of aFileName
		return theExtension as string
	end if
	return aFileName
end returnExtension

on path2URL(thepath)
	-- Needed to properly URL encode the path
	return do shell script "python -c \"import urllib, sys; print (urllib.quote(sys.argv[1]))\" " & quoted form of thepath
end path2URL

-- remove various characters in the list as set in the properties above
on CleanName(theName)
	set newName to ""
	repeat with i from 1 to length of theName
		--check if the character is in dChar1
		--replace it with the rChar if it is
		if ((character i of theName) is in dChar1) then
			set newName to newName & rChar
			--check if the character is in dChar2
			--remove it completely if it is
		else if ((character i of theName) is in dChar2) then
			set newName to newName & ""
			--if the character is not in either dChar1 or
			--dChar2, keep it in the file name
		else
			set newName to newName & character i of theName
		end if
	end repeat
	return newName
end CleanName

-- an additional clean up function to remove colons and replace with slashes, to convert from POSIX back to HTML references
on myCleanURL(theName)
	set newName to ""
	repeat with i from 1 to length of theName
		--check if the character is in dChar1
		--replace it with the rChar if it is
		if ((character i of theName) is in mydis) then
			set newName to newName & myrep
			--check if the character is in dChar2
			--remove it completely if it is
		else if ((character i of theName) is in mydis) then
			set newName to newName & ""
			--if the character is not in either dChar1 or
			--dChar2, keep it in the file name
		else
			set newName to newName & character i of theName
		end if
	end repeat
	return newName
end myCleanURL

-- make sure a folder is present in the parent, if it isn't then build one, return a reference to the folder
to checkForFolder(fParent, fName)
	tell application "System Events"
		if not (exists folder fName of folder fParent) then
			set output to path of (make new folder at end of folder fParent with properties {name:fName})
		else
			set output to (path of (folder fName of folder fParent))
		end if
	end tell
	return output
end checkForFolder


