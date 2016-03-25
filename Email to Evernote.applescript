(*

Advanced Apple Mail to Evernote
Version 1.0
https://github.com/scouture

// ATTRIBUTION
This script is forked from version 2.04 of "Apple Mail to Evernote" script by Veritrope.com
http://veritrope.com/code/apple-mail-to-evernote/

// TERMS OF USE:
This work is licensed under the Creative Commons Attribution-NonCommercial-ShareAlike 3.0 Unported License. 
To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/3.0/ or send a letter to Creative Commons, 444 Castro Street, Suite 900, Mountain View, California, 94041, USA.

// IMPORTANT LINKS:
-- Original Project Page: http://veritrope.com/code/apple-mail-to-evernote
-- GROWL (App Store Version) (Optional): http://bit.ly/GrowlApp
-- terminal-notifier (Optional): https://github.com/alloy/terminal-notifier/downloads
-- FastScripts (Optional): http://bit.ly/FastScripts
-- Alfred (Optional): http://www.alfredapp.com

// REQUIREMENTS:
THIS SCRIPT REQUIRES LION OR GREATER (OS X 10.7+) TO RUN WITHOUT MODIFICATION

// INSTALLATION:  
-- You can save this script to /Library/Scripts/Mail Scripts and launch it using the system-wide script menu from the Mac OS X menu bar. (The script menu can be activated using the AppleScript Utility application). 
-- To use, highlight the email messages you want to archive into Evernote and run this script file;
-- The "User Switches" below allow you to customize the way this script works.
-- You can save this script as a service and trigger it with a keyboard shortcut.

	(Optional but recommended)
	Easier Keyboard Shortcut with FastScripts
	-- Download and Install FastScripts here: 
	-- http://bit.ly/FastScripts
	Assign to Alfred keyword
	-- Download and install Alfred here:
	-- http://www.alfredapp.com
	

// CHANGELOG:
    	* 1.00 (February 16, 2013) 
	- Fork from v.2.0.4 of "Apple Mail to Evernote" script by Veritrope.com (http://veritrope.com/code/apple-mail-to-evernote/)
	- Made GROWL notifications optional
	- Added OSX notifications with "terminal-notifier"
	- Added the ability to turn off notifications
	- Added mail archiving and flagging
	- Code cleanup
*)

(* 
======================================
// USER SWITCHES 
======================================
*)

-- SET THIS TO "OFF" IF YOU WANT TO SKIP THE TAGGING/NOTEBOOK DIALOG
-- AND SEND ITEMS DIRECTLY INTO YOUR DEFAULT NOTEBOOK
property tagging_Switch : "ON"

-- IF YOU'VE DISABLED THE TAGGING/NOTEBOOK DIALOG,
-- TYPE THE NAME OF THE NOTEBOOK YOU WANT TO SEND ITEM TO
-- BETWEEN THE QUOTES IF IT ISN'T YOUR DEFAULT NOTEBOOK.
-- (EMPTY SENDS TO DEFAULT)
property EVnotebook : ""

-- IF TAGGING IS ON AND YOU'D LIKE TO CHANGE THE DEFAULT TAG,
-- TYPE IT BETWEEN THE QUOTES ("Email Message" IS DEFAULT)
property defaultTags : ""

-- SET THIS "ON" IF YOU WISH TO ACTIVATE ARCHIVING OF PROCESSED MESSAGES IN '<year> Archive' MAILBOX
property archiving : "ON"

-- SET THIS "ON" IF YOU WISH TO FLAG PROCESSED MESSAGES
property flagging : "ON"

-- SET THIS TO "GROWL", "OSX" OR "OFF". FOR OSX NOTIFICATIONS, YOU MUST INSTALL 'terminial-notifier.app' AND SET COMMAND PATH IN 'terminal_notifier_path' PROPERTY
property notifications : "OSX"
(* 
======================================
// OTHER PROPERTIES 
======================================
*)

-- Global properties
property successCount : 0
property growl_Running : "false"
property osxNotifications_Available : "false"
property myTitle : "Mail Item"
property theMessages : {}
property thisMessage : ""
property itemNum : "0"
property attNum : "0"
property errNum : "0"
property userTag : ""
property EVTag : {}
property multiHTML : ""
property theSourceItems : {}
property mySource : ""
property decode_Success : ""
property finalHTML : ""
property myHeaders : ""
property mysource_Paragraphs : {}
property base64_Raw : ""

-- Archive properties
property archive_mailbox_label : "Archive" -- Will generate "<year> <label>"
property archive_flag : 3

-- Notification properties
property terminal_notifier_path : "/usr/local/bin/terminal-notifier.app/Contents/MacOS/terminal-notifier"
property notificationAppName : "Apple Mail to Evernote"
property notificationAction : "com.apple.Mail"
property notificationIcon : "Mail"


(* 
======================================
// MAIN PROGRAM 
======================================
*)

--RESET ITEMS
set successCount to "0"
set errNum to "0"
set AppleScript's text item delimiters to ""

try
	-- Check for Growl
	if notifications is "GROWL" then
		-- Activate Grown
		my Growl_Check()
	end if
	
	-- Set up activites
	my item_Check()
	
	-- Check for selected messages
	if theMessages is not {} then
		
		-- Get messages count
		my item_Count(theMessages)
		
		-- Announce the export of items
		my process_Notification(itemNum, attNum)
		
		-- Process mail items for export
		my mail_Process(theMessages)
		
	else
		-- No messages selected
		set successCount to -1
	end if
	
	-- Show results notification
	my processed_Notification(successCount, errNum)
	
	-- Error handling
on error errText number errNum
	if growl_Running is true then
		
		if errNum is -128 then
			
			-- Failure notification for cancel
			notification("Failure Notification", "User Cancelled", "Failed to export!", notificationAppName, notificationAction, notificationIcon)
			
		else
			
			-- Failure notification for error
			notification("Failure Notification", "Import Failure", "Failed to export " & return & myTitle & "\"  due to the following error: " & return & errText, notificationAppName, notificationAction, notificationIcon)
		end if
		
		-- Non notification error message
	else if growl_Running is false and osxNotifications_Available is false then
		display dialog "Item Failed to Import: " & errNum & return & errText with icon 0
	end if
end try

(* 
======================================
// PREPARATORY SUBROUTINES 
=======================================
*)

-- App detect
on appIsRunning(appName)
	tell application "System Events" to (name of processes) contains appName
end appIsRunning

-- Set up activities
on item_Check()
	set myPath to (path to home folder)
	tell application "Mail"
		try
			set theMessages to selection
		end try
	end tell
end item_Check

-- Get count of items and attachments
on item_Count(theMessages)
	tell application "Mail"
		set itemNum to count of theMessages
		set attNum to 0
		repeat with theMessage in theMessages
			set attNum to attNum + (count of mail attachment of theMessage)
		end repeat
	end tell
end item_Count

(* 
======================================
// TAGGING AND NOTEBOOK SUBROUTINES
=======================================
*)

-- Tagging and notebook selection dialog
on tagging_Dialog()
	try
		display dialog "" & Â
			"Please Enter Your Tags Below:
(Multiple Tags Separated By Commas)" with title "Veritrope.com | Apple Mail to Evernote Export" default answer defaultTags buttons {"Create in Default Notebook", "Select Notebook from List", "Cancel"} default button "Create in Default Notebook" cancel button Â
			"Cancel" with icon path to resource "Evernote.icns" in bundle (path to application "Evernote")
		set dialogresult to the result
		set userInput to text returned of dialogresult
		set ButtonSel to button returned of dialogresult
		set theDelims to {","}
	on error number -128
		set errNum to -128
	end try
	
	-- Assemble tag list
	set theTags to my Tag_List(userInput, theDelims)
	
	-- Reset, final check and formating of tags
	set EVTag to {}
	set EVTag to my Tag_Check(theTags)
	
	-- Select Notebook
	if ButtonSel is "Select Notebook from List" then set EVnotebook to my Notebook_List()
end tagging_Dialog

-- Get Evernote's default Notebook
on default_Notebook()
	tell application "Evernote"
		set get_defaultNotebook to every notebook whose default is true
		if EVnotebook is "" then
			set EVnotebook to name of (item 1 of get_defaultNotebook) as text
		end if
	end tell
end default_Notebook

-- Tag selection subroutine
on Tag_List(userInput, theDelims)
	set oldDelims to AppleScript's text item delimiters
	set theList to {userInput}
	repeat with aDelim in theDelims
		set AppleScript's text item delimiters to aDelim
		set newList to {}
		repeat with anItem in theList
			set newList to newList & text items of anItem
		end repeat
		set theList to newList
	end repeat
	set AppleScript's text item delimiters to oldDelims
	return theList
end Tag_List

-- Creates tags if they don't exist
on Tag_Check(theTags)
	tell application "Evernote"
		set finalTags to {}
		repeat with theTag in theTags
			if (not (tag named theTag exists)) then
				try
					set makeTag to make tag with properties {name:theTag}
					set end of finalTags to makeTag
				end try
			else
				set end of finalTags to tag theTag
			end if
		end repeat
	end tell
	return finalTags
end Tag_Check

-- Evernote Notebook selection subroutine
on Notebook_List()
	tell application "Evernote"
		activate
		set listOfNotebooks to {} (*PREPARE TO GET EVERNOTE'S LIST OF NOTEBOOKS *)
		set EVNotebooks to every notebook (*GET THE NOTEBOOK LIST *)
		repeat with currentNotebook in EVNotebooks
			set currentNotebookName to (the name of currentNotebook)
			copy currentNotebookName to the end of listOfNotebooks
		end repeat
		set Folders_sorted to my simple_sort(listOfNotebooks) (*SORT THE LIST *)
		set SelNotebook to choose from list of Folders_sorted with title "Select Evernote Notebook" with prompt Â
			"Current Evernote Notebooks" OK button name "OK" cancel button name "New Notebook" (*USER SELECTION FROM NOTEBOOK LIST *)
		if (SelNotebook is false) then (*CREATE NEW NOTEBOOK OPTION *)
			set userInput to Â
				text returned of (display dialog "Enter New Notebook Name:" default answer "")
			set EVnotebook to userInput
		else
			set EVnotebook to item 1 of SelNotebook
		end if
	end tell
end Notebook_List

(* 
======================================
// UTILITY SUBROUTINES 
=======================================
*)

-- Extraction subroutine
on extractBetween(SearchText, startText, endText)
	set tid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to startText
	set endItems to text of text item -1 of SearchText
	set AppleScript's text item delimiters to endText
	set beginningToEnd to text of text item 1 of endItems
	set AppleScript's text item delimiters to tid
	return beginningToEnd
end extractBetween

-- Sort subroutine
on simple_sort(my_list)
	set the index_list to {}
	set the sorted_list to {}
	repeat (the number of items in my_list) times
		set the low_item to ""
		repeat with i from 1 to (number of items in my_list)
			if i is not in the index_list then
				set this_item to item i of my_list as text
				if the low_item is "" then
					set the low_item to this_item
					set the low_item_index to i
				else if this_item comes before the low_item then
					set the low_item to this_item
					set the low_item_index to i
				end if
			end if
		end repeat
		set the end of sorted_list to the low_item
		set the end of the index_list to the low_item_index
	end repeat
	return the sorted_list
end simple_sort

(* 
======================================
// PROCESS MAIL ITEMS SUBROUTINE
=======================================
*)

on mail_Process(theMessages)
	--CHECK DEFAULT NOTEBOOK
	my default_Notebook()
	tell application "Mail"
		try
			if tagging_Switch is "ON" then my tagging_Dialog()
			
			repeat with thisMessage in theMessages
				try
					-- Get message info
					set myTitle to the subject of thisMessage
					set myContent to the content of thisMessage
					set mySource to the source of thisMessage
					set ReplyAddr to the reply to of thisMessage
					set EmailDate to the date received of thisMessage
					set allRecipients to (every to recipient of item 1 of thisMessage)
					
					-- Assemble all to : resipients for header
					set toRecipients to ""
					repeat with allRecipient in allRecipients
						set toName to (name of allRecipient)
						set toEmail to (address of allRecipient)
						set toCombined to toName & space & "(" & toEmail & ")<br/>"
						set toRecipients to (toRecipients & toCombined as string)
					end repeat
					
					-- Create mail message URL
					set theRecipient to ""
					set ex to ""
					set MsgLink to ""
					try
						set theRecipient to ""
						set theRecipient to the address of to recipient 1 of thisMessage
						set MsgLink to "message://%3c" & thisMessage's message id & "%3e"
						if theRecipient is not "" then set ex to my extractBetween(ReplyAddr, "<", ">") -- extract the Address
					end try
					
					-- HTML email functions
					set theBoundary to my extractBetween(mySource, "boundary=\"", "\"")
					set theMessagestart to (return & "--" & theBoundary)
					set theMessageEnd to ("--" & theBoundary & return & "Content-Type:")
					set paraSource to paragraphs of mySource
					set myHeaderlines to paragraphs of (all headers of thisMessage as rich text)
					
					
					-- Get content type
					repeat with myHeaderline in myHeaderlines
						if myHeaderline starts with "Content-Type: " then
							set myHeaders to my extractBetween(myHeaderline, "Content-Type: ", ";")
						end if
					end repeat
					set cutSource to my stripHeader(paraSource, myHeaderlines)
					set evHTML to cutSource
				end try
				
				-- Make header template
				set the_Template to "
<table border=\"1\" width=\"100%\" cellspacing=\"0\" cellpadding=\"2\">
<tbody>

<tr BGCOLOR=\"#ffffff\">
<td valign=\"top\"><font color=\"#797979\"><strong>From: </strong>  </td>
<td valign=\"top\" ><a href=\"mailto:" & ex & "\">" & ex & "</a></td>
</tr>

<tr BGCOLOR=\"#ffffff\">
<td valign=\"top\"><font color=\"#797979\"><strong>Subject: </strong>  </td>
<td valign=\"top\" ><strong>" & myTitle & "</strong></td>
</tr>

<tr BGCOLOR=\"#ffffff\">
<td valign=\"top\"><font color=\"#797979\"><strong>Date / Time:  </strong></td>
<td valign=\"top\">" & EmailDate & "</td>
</tr>

<tr BGCOLOR=\"#ffffff\">
<td valign=\"top\"><font color=\"#797979\"><strong>To:</strong></td>
<td valign=\"top\">" & toRecipients & "</td>
</tr>

</tbody>
</table>
<hr />"
				
				-- Sent item to Evernote subroutine
				my make_Evernote(myTitle, EVTag, EmailDate, MsgLink, myContent, mySource, theBoundary, theMessagestart, theMessageEnd, myHeaders, thisMessage, evHTML, EVnotebook, the_Template)
				
				-- Run message post process subroutine
				my mail_post_Process(theMessages)
				
			end repeat
		end try
	end tell
end mail_Process

-- Archiving and flagging of processed emails
on mail_post_Process(theMessages)
	tell application "Mail"
		repeat with m in theMessages
			
			-- Flag message
			if flagging is "ON" then
				set flag index of m to archive_flag as integer
			end if
			
			-- Archive message
			if archiving is "ON" then
				set mb to mailbox of m
				set acc to account of mb
				set archive_mailbox to get (the (year of (current date)) as string) & " " & archive_mailbox_label
				log "here"
				log archive_mailbox
				try
					set archive to acc's mailbox archive_mailbox
				on error
					display alert "No '" & archive_mailbox & "' mailbox found for account '" & acc & "'."
					return
				end try
				
				try
					move m to archive
				on error
					display alert "Error"
					return
				end try
			end if
			
		end repeat
	end tell
end mail_post_Process


(* 
======================================
// MAKE ITEM IN EVERNOTE SUBROUTINE
=======================================
*)

on make_Evernote(myTitle, EVTag, EmailDate, MsgLink, myContent, mySource, theBoundary, theMessagestart, theMessageEnd, myHeaders, thisMessage, evHTML, EVnotebook, the_Template)
	
	tell application "Evernote"
		try
			-- Is it a text email?
			if myHeaders contains "text/plain" then
				set n to create note with html the_Template title myTitle notebook EVnotebook
				if EVTag is not {} then assign EVTag to n
				tell n to append text myContent
				set creation date of n to EmailDate
				set source URL of n to MsgLink
				
				-- Is it multipart alternative?
			else if myHeaders contains "multipart/alternative" then
				
				-- Check for Base64
				set base64Detect to my base64_Check(mySource)
				
				-- If message if Base64 encoded
				if base64Detect is true then
					set multiHTML to my extractBetween(mySource, "Content-Transfer-Encoding: base64", theBoundary)
					
					-- Strip out content-disposition, if necessary
					if multiHTML contains "Content-Disposition: inline" then set multiHTML to my extractBetween(multiHTML, "Content-Disposition: inline", theBoundary)
					if multiHTML contains "Content-Transfer-Encoding: 7bit" then set multiHTML to my extractBetween(multiHTML, "Content-Transfer-Encoding: 7bit", theBoundary)
					
					-- Decode Base64
					set baseHTML to do shell script "echo " & (quoted form of multiHTML) & "| openssl base64 -d"
					
					-- Make note in Evernote
					set n to create note with html the_Template title myTitle notebook EVnotebook
					if EVTag is not {} then assign EVTag to n
					tell n to append html baseHTML
					set creation date of n to EmailDate
					set source URL of n to MsgLink
				else
					
					-- If message is not Base64 encoded
					set finalHTML to my htmlFix(mySource, theBoundary, myContent)
					if decode_Success is true then
						
						-- Make note in Evernote
						set n to create note with html the_Template title myTitle notebook EVnotebook
						if EVTag is not {} then assign EVTag to n
						tell n to append html finalHTML
						set creation date of n to EmailDate
						set source URL of n to MsgLink
					else
						
						-- Make note in Evernote
						set n to create note with html the_Template title myTitle notebook EVnotebook
						if EVTag is not {} then assign EVTag to n
						tell n to append text myContent
						set creation date of n to EmailDate
						set source URL of n to MsgLink
					end if
				end if
				
				-- Is it multipart mixed?
			else if myHeaders contains "multipart" then
				if mySource contains "Content-Type: text/html" then
					
					-- Check for Base64
					set base64Detect to my base64_Check(mySource)
					
					-- If message is Base64 encoded
					if base64Detect is true then
						set baseHTML to my base64_Decode(mySource)
						
						-- Make note in Evernote
						set n to create note with html the_Template title myTitle notebook EVnotebook
						if EVTag is not {} then assign EVTag to n
						tell n to append html baseHTML
						set creation date of n to EmailDate
						set source URL of n to MsgLink
						
						-- If message is not Base64 encoded
					else if base64Detect is false then
						set finalHTML to my htmlFix(mySource, theBoundary, myContent)
						if decode_Success is true then
							
							-- Make note in Evernote
							set n to create note with html the_Template title myTitle notebook EVnotebook
							if EVTag is not {} then assign EVTag to n
							tell n to append html finalHTML
							set creation date of n to EmailDate
							set source URL of n to MsgLink
						else
							
							-- Make note in Evernote
							set n to create note with html the_Template title myTitle notebook EVnotebook
							if EVTag is not {} then assign EVTag to n
							tell n to append text myContent
							set creation date of n to EmailDate
							set source URL of n to MsgLink
						end if
					end if
					
				else if mySource contains "text/plain" then
					
					-- Make note in Evernote
					set n to create note with html the_Template title myTitle notebook EVnotebook
					if EVTag is not {} then assign EVTag to n
					tell n to append text myContent
					set creation date of n to EmailDate
					set source URL of n to MsgLink
					
				end if
				
				-- Multipart mixed
				
				-- Other types of HTML-encoding
			else
				
				-- Check for Base64
				set base64Detect to my base64_Check(mySource)
				
				-- If message is Base64 encoded
				if base64Detect is true then
					set finalHTML to my base64_Decode(mySource)
				else
					set multiHTML to my extractBetween(evHTML, "</head>", "</html>")
					set finalHTML to my htmlFix(multiHTML, theBoundary, myContent) as text
				end if
				
				-- Make note in Evernote
				set n to create note with html the_Template title myTitle notebook EVnotebook
				if EVTag is not {} then assign EVTag to n
				tell n to append html finalHTML
				set creation date of n to EmailDate
				set source URL of n to MsgLink
				
				-- End of message processing
			end if
			
			-- Start of attachment processing
			tell application "Mail"
				
				-- If attachment present, run attachment subroutine
				if thisMessage's mail attachments is not {} then my attachment_process(thisMessage, n)
			end tell
			
			-- Item has finished. Count as success
			set successCount to successCount + 1
		end try
	end tell
	log "successCount: " & successCount
end make_Evernote



(* 
======================================
// ATTACHMENT SUBROUTINES 
=======================================
*)

-- Folder exists?
on f_exists(ExportFolder)
	try
		set myPath to (path to home folder)
		get ExportFolder as alias
		set SaveLoc to ExportFolder
	on error
		tell application "Finder" to make new folder with properties {name:"Temp Export From Mail"}
	end try
end f_exists

-- Attachment processing
on attachment_process(thisMessage, n)
	tell application "Mail"
		
		-- Make sure text item delimiters are default
		set AppleScript's text item delimiters to ""
		
		-- Temp files processed on the Desktop
		set ExportFolder to ((path to desktop folder) & "Temp Export From Mail:") as string
		set SaveLoc to my f_exists(ExportFolder)
		
		-- Process attachments
		set theAttachments to thisMessage's mail attachments
		set attCount to 0
		repeat with theAttachment in theAttachments
			set theFileName to ExportFolder & theAttachment's name
			try
				save theAttachment in file theFileName
			end try
			tell application "Evernote"
				tell n to append attachment file theFileName
			end tell
			
			-- Silent delete of temp file
			set trash_Folder to path to trash folder from user domain
			do shell script "mv " & quoted form of POSIX path of theFileName & space & quoted form of POSIX path of trash_Folder
			
		end repeat
		
		-- Silent delete of temp folder
		set success to my trashfolder(SaveLoc)
		
	end tell
end attachment_process

-- Silent delete of temp folder
on trashfolder(SaveLoc)
	try
		set trashfolderpath to ((path to trash) as Unicode text)
		set srcfolderinfo to info for (SaveLoc as alias)
		set srcfoldername to name of srcfolderinfo
		set SaveLoc to quoted form of POSIX path of SaveLoc
		set counter to 0
		repeat
			if counter is equal to 0 then
				set destfolderpath to trashfolderpath & srcfoldername & ":"
			else
				set destfolderpath to trashfolderpath & srcfoldername & " " & counter & ":"
			end if
			try
				set destfolderalias to destfolderpath as alias
			on error
				exit repeat
			end try
			set counter to counter + 1
		end repeat
		set destfolderpath to quoted form of POSIX path of destfolderpath
		set command to "ditto " & SaveLoc & space & destfolderpath
		do shell script command
		-- this won't be executed if the ditto command errors
		set command to "rm -r " & SaveLoc
		do shell script command
		return true
	on error
		return false
	end try
end trashfolder

(* 
======================================
// HTML CLEANUP SUBROUTINES 
=======================================
*)

-- Header strip
on stripHeader(paraSource, myHeaderlines)
	
	-- Find the last non-empty header line
	set lastheaderline to ""
	set n to count (myHeaderlines)
	repeat while (lastheaderline = "")
		set lastheaderline to item n of myHeaderlines
		set n to n - 1
	end repeat
	
	-- Compare header to source
	set sourcelength to (count paraSource)
	repeat with n from 1 to sourcelength
		if (item n of paraSource is equal to "") then exit repeat
	end repeat
	
	-- Strip out headers
	set cutSourceItems to (items (n + 1) thru sourcelength of paraSource)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to return
	set cutSource to (cutSourceItems as text)
	set AppleScript's text item delimiters to oldDelims
	
	return cutSource
	
end stripHeader

-- Base64 check
on base64_Check(mySource)
	set base64Detect to false
	set base64MsgStr to "Content-Transfer-Encoding: base64"
	set base64ContentType to "Content-Type: text"
	set base64MsgOffset to offset of base64MsgStr in mySource
	set base64ContentOffset to offset of base64ContentType in mySource
	set base64Offset to base64MsgOffset - base64ContentOffset as real
	set theOffset to base64Offset as number
	if theOffset is not greater than or equal to 50 then
		if theOffset is greater than -50 then set base64Detect to true
	end if
	return base64Detect
end base64_Check

-- Base64 decode
on base64_Decode(mySource)
	
	-- Use TID to quickly isolate Base64 data
	set oldDelim to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "Content-Type: text/html"
	set base64_Raw to second text item of mySource
	set AppleScript's text item delimiters to linefeed & linefeed
	set base64_Raw to second text item of base64_Raw
	set AppleScript's text item delimiters to "-----"
	set multiHTML to first text item of base64_Raw
	set AppleScript's text item delimiters to oldDelim
	
	-- Decode Base64
	set baseHTML to do shell script "echo " & (quoted form of multiHTML) & "| openssl base64 -d"
	
	return baseHTML
end base64_Decode


-- HTML fix
on htmlFix(evHTML, theBoundary, myContent)
	
	set oldDelims to AppleScript's text item delimiters
	set multiHTML to evHTML as string
	
	-- Test for / strip out header
	set paraSource to paragraphs of multiHTML
	if item 1 of paraSource contains "Received:" then
		set myHeaderlines to (item 1 of paraSource)
		set multiHTML to my stripHeader(paraSource, myHeaderlines)
	end if
	
	-- Trim ending
	if multiHTML contains "</html>" then
		set multiHTML to my extractBetween(multiHTML, "Content-Type: text/html", "</html>")
	else
		set multiHTML to my extractBetween(multiHTML, "Content-Type: text/html", theBoundary)
	end if
	set paraSource to paragraphs of multiHTML
	
	-- Test for / strip out leading semi-colon
	if item 1 of paraSource contains ";" then
		set myHeaderlines to (item 1 of paraSource)
		set multiHTML to my stripHeader(paraSource, myHeaderlines)
		set paraSource to paragraphs of multiHTML
	end if
	
	-- Test for empty line / clean subsequent encoding info, if necessary
	if item 1 of paraSource is "" then
		
		-- Test for / strip out content-transfer-encoding
		if item 2 of paraSource contains "Content-Transfer-Encoding" then
			set myHeaderlines to (item 2 of paraSource)
			set multiHTML to my stripHeader(paraSource, myHeaderlines)
			set paraSource to paragraphs of multiHTML
		end if
		-- Test for / strip out charset
		if item 2 of paraSource contains "charset" then
			set myHeaderlines to (item 2 of paraSource)
			set multiHTML to my stripHeader(paraSource, myHeaderlines)
			set paraSource to paragraphs of multiHTML
		end if
	end if
	
	-- Test for / strip out content-transfer-encoding
	if item 1 of paraSource contains "Content-Transfer-Encoding" then
		set myHeaderlines to (item 1 of paraSource)
		set multiHTML to my stripHeader(paraSource, myHeaderlines)
		set paraSource to paragraphs of multiHTML
	end if
	
	-- Test for / strip out charset
	if item 1 of paraSource contains "charset" then
		set myHeaderlines to (item 1 of paraSource)
		set multiHTML to my stripHeader(paraSource, myHeaderlines)
		set paraSource to paragraphs of multiHTML
	end if
	
	-- Clean content
	set AppleScript's text item delimiters to theBoundary
	set theSourceItems to text items of multiHTML
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "&#" & "37;" as string
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "="
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "%"
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%\""
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "=\""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%" & (ASCII character 13)
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%%"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "%"
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%" & (ASCII character 10)
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%0A"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%09"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%C2%A0"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "&nbsp;"
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "%20"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to " "
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to (ASCII character 10)
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to ""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "="
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "&#" & "61;" as string
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "$"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "&#" & "36;" as string
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "'"
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "&apos;"
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to "\""
	set theSourceItems to text items of theEncoded
	set AppleScript's text item delimiters to "\\\""
	set theEncoded to theSourceItems as text
	
	set AppleScript's text item delimiters to oldDelims
	
	set trimHTML to my extractBetween(theEncoded, "</head>", "</html>")
	
	set theHTML to myContent
	
	try
		set decode_Success to false
		
		-- UTF-8 conversion
		set NewEncodedText to do shell script "echo " & quoted form of trimHTML & " | iconv -t UTF-8 "
		set the_UTF8Text to quoted form of NewEncodedText
		
		-- URL decode conversion
		set theDecodeScript to "php -r \"echo urldecode(" & the_UTF8Text & ");\"" as text
		set theDecoded to do shell script theDecodeScript
		
		-- Fix for apostrophe / percent / equal issues
		set AppleScript's text item delimiters to "&apos;"
		set theSourceItems to text items of theDecoded
		set AppleScript's text item delimiters to "'"
		set theDecoded to theSourceItems as text
		
		set AppleScript's text item delimiters to "&#" & "37;" as string
		set theSourceItems to text items of theDecoded
		set AppleScript's text item delimiters to "%"
		set theDecoded to theSourceItems as text
		
		set AppleScript's text item delimiters to "&#" & "61;" as string
		set theSourceItems to text items of theDecoded
		set AppleScript's text item delimiters to "="
		set theDecoded to theSourceItems as text
		
		--RETURN THE VALUE
		set finalHTML to theDecoded
		set decode_Success to true
		return finalHTML
	end try
	
end htmlFix

(* 
======================================
// NOTIFICATIONS SUBROUTINES
=======================================
*)

-- Check for Growl and initialize 
on Growl_Check()
	if appIsRunning("Growl") then
		set growl_Running to true
		tell application "GrowlHelperApp"
			set allNotificationsFiles to {"Import Notification", "Success Notification", "Failure Notification"}
			set enabledNotificationsFiles to {"Import Notification", "Success Notification", "Failure Notification"}
			register as application Â
				notificationAppName all notifications allNotificationsFiles Â
				default notifications enabledNotificationsFiles Â
				icon of application notificationIcon
		end tell
	end if
end Growl_Check

-- Check for presence of terminal-notifier.app
on osxNotifications_Check()
	tell application "System Events"
		if exists file terminal_notifier_path then
			set osxNotifications_Available to true
			return true
		else
			set osxNotifications_Available to false
			return false
		end if
	end tell
end osxNotifications_Check

-- Annouce the count of total items to export
on process_Notification(itemNum, attNum)
	set attPlural to ""
	if attNum = 0 then
		set attNum to "No"
	else if attNum > 1 then
		set attPlural to "s"
	end if
	
	set Plural_Test to (itemNum) as number
	if Plural_Test is greater than 1 then
		notification("Import Notification", "Import To Evernote Started", "Now Processing " & itemNum & " Items with " & attNum & " attachment" & attPlural & ".", notificationAppName, notificationAction, notificationIcon)
		
	else
		notification("Import Notification", "Import To Evernote Started", "Now Processing " & itemNum & " Item With " & attNum & " Attachment" & attPlural & ".", notificationAppName, notificationAction, notificationIcon)
		
	end if
	
end process_Notification

-- Results notification
on processed_Notification(successCount, errNum)
	
	-- Notification failure : user canceled	
	if errNum is -128 then
		notification("Failure Notification", "User Cancelled", "Failed to export!", notificationAppName, notificationAction, notificationIcon)
	end if
	
	set Plural_Test to (successCount) as number
	
	-- Notification failure : no items selected in Mail	
	if Plural_Test is -1 then
		notification("Failure Notification", "Import Failure", "No Items Selected In Apple Mail!", notificationAppName, notificationAction, notificationIcon)
		
		-- Notification failure : no items exported from Mail	
	else if Plural_Test is 0 then
		notification("Failure Notification", "Import Failure", "No Items Exported From Mail!", notificationAppName, notificationAction, notificationIcon)
		
		-- Notification success
	else if Plural_Test is equal to 1 then
		notification("Success Notification", "Import Success", "Successfully Exported " & itemNum & " Item to the " & EVnotebook & " Notebook in Evernote", notificationAppName, notificationAction, notificationIcon)
		
		-- Notification success
	else if Plural_Test is greater than 1 then
		notification("Success Notification", "Import Success", "Successfully Exported " & itemNum & " Items to the " & EVnotebook & " Notebook in Evernote", notificationAppName, notificationAction, notificationIcon)
	end if
	set itemNum to "0"
	
end processed_Notification


-- Trigger OSX notification
on terminal_notification(notificationTitle, notificationMessage, notoficationAction)
	if osxNotifications_Check() is true then
		if notoficationAction is not "" then
			set action to " -activate '" & notoficationAction & "'"
		else
			set action to ""
		end if
		do shell script terminal_notifier_path & " -title '" & notificationTitle & "' -message '" & notificationMessage & "'" & action
	end if
end terminal_notification


-- Global notification function
on notification(nName, nTitle, nMessage, nAppName, nAction, nIcon)
	
	if notifications is "GROWL" then
		if growl_Running is true then
			tell application "GrowlHelperApp"
				notify with name nName title nTitle description nMessage application name nAppName icon of application nIcon
			end tell
		end if
		
	else if notifications is "OSX" then
		terminal_notification(nTitle, nMessage, nAction)
	end if
	
end notification