# outlook-documents-applescript
Selectively handle attachments in Outlook for Mac using AppleScript.  Tidy and clean up.  Save server space.
I use this to push all of my .office type documents to a sync folder (box.com) in the coded case.  THese are synced to the cloud.

(*

Outlook 2016

When an email is selected and then this script is run
* each email of one or more will be checked
* for each attachment in the email a comparison to a list of file types to save and delete is made
* for each attachment that is Office type document, it is cataloged, saved to downloadPath, and deleted from the email.
* the catalog of attachmentList holds a record for each attachment for processing information back to user
* anAttachment contains {nameAtt, extensionDigits, sizeAtt, examination, shoulddelete, wassaved, isdeleted}
* if a file is removed, its name and filesize will be included in a simple HTML table and pre-pended to the original email; the file name will link to the locally saved file.
* file names are saved with uniquestamp prefix, spaces are removed to promote uniqueness of files in a single deep folder which means I don't have to come up with some folder scheme to store attachments

* if a file isn't matching .office kind, it will remain - no touchy, no feely.

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
