# outlook-documents-applescript
Selectively handle attachments in Outlook for Mac using AppleScript.  Tidy and clean up.  Save server space.

(*

Outlook 2016

When an email is selected and then this script is run
* each email of one or more will be checked
* for each attachment in the email a comparison to a list of file types to save and delete is made
* for each attachment that is Office type document, it is cataloged, saved to downloadPath, and deleted from the email.
* the catalog of attachmentList holds a record for each attachment for processing information back to user
	anAttachment contains {nameAtt, extensionDigits, sizeAtt, examination, shoulddelete, wassaved, isdeleted}
Script is not optimized and probably uses more variables and lists than needed

Things to do:
* convert hte attachmentList into a tab delimited table of title, size, and status - complete
* convert the table to an HTML version - complete
* convert the HTML version to include links to the saved files - complete
* insert the hyperlinked HTML table into the body of the email using a sexy set of code - complete
* remove large images or photos in the same manner as above however leave some sort of thumbnail behind in the table linked to the original - concept
* if an image thumb remains, mark the email so that script doesn't address attachments more than one time, or perhaps offload all thumbnails to desktop such that HTML uses offline image - concept
Functions used are found in the various websites noted in the function, attributed therein
*)
