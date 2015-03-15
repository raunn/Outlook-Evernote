--
--  AppDelegate.applescript
--  Outlook-Evernote
--
--  Created by Raun Nohavitza on 3/15/15.
--  Copyright (c) 2015 Raun Nohavitza. All rights reserved.
--

script AppDelegate
	property parent : class "NSObject"
	
	-- IBOutlets
	property theWindow : missing value
    property tagInputField: missing value
    property titleInputField: missing value
    property defaultButton: missing value
    property cabinetButton: missing value
    property cancelButton: missing value
    property mergeCheckbox: missing value

    set visible of theWindow to false
    (*
     ======================================
     // USER SWITCHES
     ======================================
     *)

    --SET THIS TO "OFF" IF YOU WANT TO SKIP THE TAGGING/NOTEBOOK DIALOG
    --AND SEND ITEMS DIRECTLY INTO YOUR DEFAULT NOTEBOOK
    property tagging_Switch : "ON"

    --IF YOU'VE DISABLED THE TAGGING/NOTEBOOK DIALOG,
    --TYPE THE NAME OF THE NOTEBOOK YOU WANT TO SEND ITEM TO
    --BETWEEN THE QUOTES IF IT ISN'T YOUR DEFAULT NOTEBOOK.
    --(EMPTY SENDS TO DEFAULT)
    property EVnotebook : ""

    --IF TAGGING IS ON AND YOU'D LIKE TO CHANGE THE DEFAULT TAG,
    --TYPE IT BETWEEN THE QUOTES (ITEM TYPE IS DEFAULT)
    property defaultTag : ""
    property defaultTagPrefix : "@Work, "

    --SOME EMAILS USING THE SRC="CID:..." TAG FOR EMBEDDED IMAGES
    --GENERATE A "There is no application set to open the URL cid:(filename)"
    --ERROR WHEN SENDING TO EVERNOTE. SETTING THIS PROPERTY TO
    --ON WILL STRIP OUT THOSE TAGS AND AVOID THE ERROR.
    --property stripEmbeddedImages : "OFF"
    property stripEmbeddedImages : "ON"
    (*
     ======================================
     // OTHER PROPERTIES
     ======================================
     *)
    property successCount : 0
    property account_Type : "free"
    property myTitle : "Item"
    property theAttachments : ""
    property thisMessage : ""
    property itemNum : "0"
    property attNum : "0"
    property errNum : "0"
    property userTag : ""
    property EVTag : {}
    property the_class : ""
    property list_Props : {}
    property SaveLoc : ""
    property selectedItem : {}
    property t_List : {}
    property c_List : {}
    property selectedItems : {}
    property titleEditable : false
    property mergeContent : ""
    property theNote : {}

	on applicationWillFinishLaunching_(aNotification)
		-- Insert code here to initialize your application before any files are opened
        --display dialog "Yo"
        set successCount to "0"
        set AppleScript's text item delimiters to ""
        set selectedItems to {}
        set ExportFolder to ""
        set SaveLoc to ""
        
        --SET UP ACTIVITIES
        set selectedItems to my item_Check()
            
        --MESSAGES SELECTED?
        if selectedItems is not missing value then
            --GET FILE COUNT
            my item_Count(selectedItems, the_class)
            
            --ANNOUNCE THE EXPORT OF ITEMS
            my process_Items(itemNum, attNum)
                
            --CHECK EVERNOTE ACCOUNT TYPE
            my account_Check()
        else
            set successCount to -1
            my notify_results(successCount)
            quit
        end if
        
        tagInputField's setStringValue_(defaultTag)
        if itemNum is 1 then
            tell application id "com.microsoft.Outlook"
                titleInputField's setStringValue_((subject of the selectedItems as text))
            end tell
            mergeCheckbox's setEnabled_(false)
            titleInputField's setEditable_(true)
            titleInputField's setEnabled_(true)
            theWindow's makeFirstResponder_(titleInputField)
        else
            titleInputField's setStringValue_("< multiple >")
            theWindow's makeFirstResponder_(tagInputField)
        end if
        set visible of theWindow to true
    end applicationWillFinishLaunching_
    
    on applicationShouldTerminate_(sender)
        -- Insert code here to do any housekeeping before your application quits
        --display dialog "bye"
        return current application's NSTerminateNow
        --quit
    end applicationShouldTerminate_

    on defaultButtonClick_(sender)
        -- Clicked Default Notebook button
        --display dialog "def"
        myWinReturn(selectedItems)
    end defaultButtonClick_
    
    on cabinetButtonClick_(sender)
        -- Clicked Cabinet button
        --display dialog "cab"
        set EVnotebook to "Cabinet"
        myWinReturn(selectedItems)
    end cabinetButtonClick_
    
    on cancelButtonClick_(sender)
        -- Clicked Cancel button
        --display dialog "bye"
        set notifTitle to "Outlook to Evernote"
        set notifSubtitle to "Failure Notification"
        set notifMessage to "User Cancelled - Failed to export!"
        display notification notifMessage with title notifTitle subtitle notifSubtitle
        delay 0.2
        quit
    end cancelButtonClick_
    
    on myWinReturn(selectedItems)
        try
            -- Process the tags from the window
            my tagging_Dialog()
            
            if selectedItems is not missing value then
                
                --PROCESS ITEMS FOR EXPORT
                my item_Process(selectedItems)
                
                --DELETE TEMP FOLDER IF IT EXISTS
                set success to my trashfolder(SaveLoc)
                
                --NO ITEMS SELECTED
            else
                set successCount to -1
            end if
            
            --Notify RESULTS
            my notify_results(successCount)
            
            --Raise Outlook window
            tell application id "com.microsoft.Outlook"
                activate
                set visible of first window to true
            end tell
            
            -- ERROR HANDLING
            on error errText number errNum
            
            if errNum is -128 then
                set notifTitle to "Outlook to Evernote"
                set notifSubtitle to "Failure Notification"
                set notifMessage to "User Cancelled - Failed to export!"
                else
                set notifTitle to "Outlook to Evernote"
                set notifSubtitle to "Failure Notification"
                set notifMessage to "Import Failure - " & errText
            end if
            
            display notification notifMessage with title notifTitle subtitle notifSubtitle
            
        end try
        --display dialog "end"
        delay 0.2
        quit

    end myWinReturn
	
    
    --CHECK ACCOUNT TYPE
    on account_Check()
        tell application "Evernote"
            set account_Info to (properties of account info 1)
            set account_Type to (account type of account_Info) as text
            if EVnotebook is "" then set EVnotebook to my default_Notebook()
        end tell
    end account_Check
    
    --SET UP ACTIVITIES
    on item_Check()
        --set myPath to (path to home folder)
        tell application id "com.microsoft.Outlook"
            try -- GET MESSAGES
                set selectedItems to selection
                set raw_Class to (class of selectedItems)
                if raw_Class is list then
                    set classList to {}
                    repeat with selectedItem in selectedItems
                        copy class of selectedItem to end of classList
                    end repeat
                    if classList contains task then
                        set the_class to "Task"
                    else
                        set raw_Class to (class of item 1 of selectedItems)
                    end if
                end if
                if raw_Class is calendar event then set the_class to "Calendar"
                if raw_Class is note then set the_class to "Note"
                if raw_Class is task then set the_class to "Task"
                if raw_Class is contact then set the_class to "Contact"
                if raw_Class is incoming message then set the_class to "Message"
                if raw_Class is text then set the_class to "Text"
                if defaultTag is "" then set defaultTag to defaultTagPrefix & the_class
            end try
            return selectedItems
        end tell
    end item_Check
    
    --GET COUNT OF ITEMS AND ATTACHMENTS
    on item_Count(selectedItems)
        tell application id "com.microsoft.Outlook"
            if the_class is not "Text" then
                set itemNum to count of selectedItems
                set attNum to 0
                try
                    repeat with selectedMessage in selectedItems
                        set attNum to attNum + (count of attachment of selectedMessage)
                    end repeat
                end try
            else
                set itemNum to 1
            end if
        end tell
    end item_Count
    
    (*
     ======================================
     // PROCESS OUTLOOK ITEMS SUBROUTINE
     ======================================
     *)
    
    on item_Process(selectedItems)
        tell application id "com.microsoft.Outlook"
            
            --TAGGING?
            if tagging_Switch is "ON" then my tagging_Dialog()
            
            --TEXT CLIPPING?
            if (class of selectedItems) is text then
                set EVTitle to "Text Clipping from Outlook"
                set theContent to selectedItems
                --CREATE IN EVERNOTE
                tell application "Evernote"
                    set theNote to create note with text theContent title EVTitle notebook EVnotebook
                    if EVTag is not {} then assign EVTag to theNote
                end tell
                
                --ITEM HAS FINISHED -- COUNT IT AS A SUCCESS!
                set successCount to 1
            else
                -- GET OUTLOOK ITEM INFORMATION
                -- CREATE THENOTE IF MERGED HERE
                --if titleEditable is true or titleEditable is 1 then
                if (state of mergeCheckbox as text) is "1" then
                    set titleInput to titleInputField's stringValue() as text
                    set EVTitle to titleInput
                    tell application "Evernote"
                        set theNote to create note with html "<!-- Merged Note -->" title EVTitle notebook EVnotebook
                    end tell
                end if
                repeat with selectedItem in selectedItems
                    try
                        set theAttachments to attachments of selectedItem
                        set raw_Attendees to attendees of selectedItem
                    end try
                    
                    try
                        set t_List to {}
                        set c_List to {}
                        
                        --LOOK FOR "TO: RECIPIENTS" AND MAKE LIST
                        set t_Recipients to (to recipients of selectedItem)
                        set t_Count to (count of t_Recipients)
                        set t_Mult to ", "
                        repeat with t_Recipient in t_Recipients
                            set t_Completed to false
                            if t_Count is 1 then set t_Mult to ""
                            set t_Address to (email address of t_Recipient)
                            try
                                set t_Name to (name of t_Address)
                                set t_List to t_List & {t_Name & " (" & (address of t_Address) & ")" & t_Mult} as string
                                set t_Completed to true
                            end try
                            if t_Completed is false then
                                set t_List to t_List & {(address of t_Address) & t_Mult} as string
                            end if
                            set t_Count to (t_Count - 1)
                        end repeat
                        
                        
                        
                        --LOOK FOR "CC: RECIPIENTS" AND MAKE LIST
                        set c_Recipients to (cc recipients of selectedItem)
                        set c_Count to (count of c_Recipients)
                        set c_Mult to ", "
                        repeat with c_Recipient in c_Recipients
                            set c_Completed to false
                            if c_Count is 1 then set c_Mult to ""
                            set c_Address to (email address of c_Recipient)
                            try
                                set c_Name to (name of c_Address)
                                set c_List to c_List & {c_Name & " (" & (address of c_Address) & ")" & c_Mult} as string
                                set c_Completed to true
                            end try
                            if c_Completed is false then
                                set c_List to c_List & {(address of c_Address) & c_Mult} as string
                            end if
                            set c_Count to (c_Count - 1)
                        end repeat
                        
                    end try
                    
                    set selectedItem to (properties of selectedItem)
                    set the_vCard to {}
                    set the_notes to ""
                    
                    --WHAT KIND OF ITEM IS IT?
                    if the_class is "Calendar" then
                        
                        (* // CALENDAR ITEM *)
                        
                        --PREPARE THE TEMPLATE
                        --LEFT SIDE (FORM FIELDS)
                        set l_1 to "Event:"
                        set l_2 to "Start Time:"
                        set l_3 to "End Time:"
                        set l_4 to "Location:"
                        set l_5 to "Notes:"
                        
                        --RIGHT SIDE (DATA FIELDS)
                        set r_1 to (subject of selectedItem)
                        set r_2 to (start time of selectedItem)
                        set r_3 to (end time of selectedItem)
                        set the_Location to (location of selectedItem)
                        if the_Location is missing value then set the_Location to "None"
                        set r_4 to the_Location
                        
                        --THE NOTES
                        set the_notes to ""
                        set item_Created to (current date)
                        try
                            set the_notes to (content of selectedItem)
                        end try
                        if the_notes is missing value then set the_notes to ""
                        
                        --ADD ORGANIZER / ATTENDEE INFO IF IT'S A MEETING
                        if (count of raw_Attendees) > 0 then
                            set the_Organizer to "<strong>Organized By: </strong><br/>" & (organizer of selectedItem) & "<br/><br/>"
                            set the_Attendees to "<strong>Invited Attendees: </strong><br/>"
                            repeat with raw_Attendee in raw_Attendees
                                
                                --GET ATTENDEE DATA
                                set raw_EmailAttendee to (email address of raw_Attendee)
                                set attend_Name to (name of raw_EmailAttendee) as text
                                set raw_Status to (status of raw_Attendee)
                                
                                --COERCE STATUS TO TEXT
                                if raw_Status contains not responded then
                                    set attend_Status to "Not Responded"
                                    else if raw_Status contains accepted then
                                    set attend_Status to "Accepted"
                                    else if raw_Status contains declined then
                                    set attend_Status to "Declined"
                                    else if raw_Status contains tentatively accepted then
                                    set attend_Status to "Tentatively Accepted"
                                end if
                                
                                --COMPILE THE ATTENDEE DATA
                                set attend_String to attend_Name & " (" & attend_Status & ")<br/>"
                                set the_Attendees to the_Attendees & attend_String
                            end repeat
                            set the_notes to (the_Organizer & the_Attendees & the_notes)
                            set raw_Attendees to ""
                        end if
                        
                        --ASSEMBLE THE TEMPLATE
                        set theContent to my make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
                        
                        --EXPORT VCARD DATA
                        try
                            set vcard_data to (icalendar data of selectedItem)
                            set vcard_extension to ".ics"
                            set the_vCard to my write_File(r_1, vcard_data, vcard_extension)
                        end try
                        
                        set theHTML to true
                        set EVTitle to r_1
                        
                    (* // NOTE ITEM *)
                    else if the_class is "note" then
                        
                        --PREPARE THE TEMPLATE
                        --LEFT SIDE (FORM FIELDS)
                        set l_1 to "Note:"
                        set l_2 to "Creation Date:"
                        set l_3 to "Category"
                        set l_4 to ""
                        set l_5 to "Notes:"
                        
                        --RIGHT SIDE (DATA FIELDS)
                        set r_1 to name of selectedItem
                        set item_Created to creation date of selectedItem
                        set r_2 to (item_Created as text)
                        
                        --GET CATEGORY INFO
                        set the_Cats to (category of selectedItem)
                        set list_Cats to {}
                        set count_Cat to (count of the_Cats)
                        repeat with the_Cat in the_Cats
                            set cat_Name to (name of the_Cat as text)
                            copy cat_Name to the end of list_Cats
                            if count_Cat > 1 then
                                copy ", " to the end of list_Cats
                                set count_Cat to (count_Cat - 1)
                                else
                                set count_Cat to (count_Cat - 1)
                            end if
                        end repeat
                        
                        set r_3 to list_Cats
                        set r_4 to ""
                        
                        set item_Created to creation date of selectedItem
                        
                        --THE NOTES
                        try
                            set the_notes to content of selectedItem
                        end try
                        if the_notes is missing value then set the_notes to ""
                        
                        --ASSEMBLE THE TEMPLATE
                        set theContent to my make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
                        
                        --EXPORT VCARD DATA
                        set vcard_data to (icalendar data of selectedItem)
                        set vcard_extension to ".ics"
                        set the_vCard to my write_File(r_1, vcard_data, vcard_extension)
                        
                        set theHTML to true
                        set EVTitle to r_1
                        
                    (* // CONTACT ITEM *)
                    else if the_class is "contact" then
                        
                        --PREPARE THE TEMPLATE
                        --LEFT SIDE (FORM FIELDS)
                        set l_1 to "Name:"
                        set l_2 to "Email:"
                        set l_3 to "Phone:"
                        set l_4 to "Address:"
                        set l_5 to "Notes:"
                        
                        --GET EMAIL INFO
                        try
                            set list_Addresses to {}
                            set email_Items to (email addresses of selectedItem)
                            repeat with email_Item in email_Items
                                set the_Type to (type of email_Item as text)
                                set addr_Item to (address of email_Item) & " (" & my TITLECASE(the_Type) & ")<br />" as text
                                copy addr_Item to the end of list_Addresses
                            end repeat
                        end try
                        
                        --GET PHONE INFO AND ENCODE TELEPHONE LINK
                        try
                            set list_Phone to {}
                            if business phone number of selectedItem is not missing value then
                                set b_Number to (business phone number of selectedItem)
                                set b_String to "<strong>Work: </strong><a href=\"tel:\\" & b_Number & "\">" & b_Number & "</a><br /><br />"
                                copy b_String to end of list_Phone
                            end if
                            if home phone number of selectedItem is not missing value then
                                set h_Number to (home phone number of selectedItem)
                                set h_String to "<p><strong>Home: </strong><a href=\"tel:\\" & h_Number & "\">" & h_Number & "<br /><br />"
                                copy h_String to end of list_Phone
                            end if
                            if mobile number of selectedItem is not missing value then
                                set m_Number to (mobile number of selectedItem)
                                set m_String to "<p><strong>Mobile: </strong><a href=\"tel:\\" & m_Number & "\">" & m_Number & "<br /><br />"
                                copy m_String to end of list_Phone
                            end if
                        end try
                        
                        --GET ADDRESS INFO
                        try
                            set list_Addr to {}
                            
                            (*BUSINESS *)
                            if business street address of selectedItem is not missing value then
                                set b_Str to (business street address of selectedItem)
                                set b_gStr to my encodedURL(b_Str)
                                if (business city of selectedItem) is not missing value then
                                    set b_Cit to (business city of selectedItem)
                                    set b_gCit to my encodedURL(b_Cit)
                                    else
                                    set b_Cit to ""
                                    set b_gCit to ""
                                end if
                                if (business state of selectedItem) is not missing value then
                                    set b_Sta to (business state of selectedItem)
                                    set b_gSta to my encodedURL(b_Sta)
                                    else
                                    set b_Sta to ""
                                    set b_gSta to ""
                                end if
                                if (business zip of selectedItem) is not missing value then
                                    set b_Zip to (business zip of selectedItem)
                                    set b_gZip to my encodedURL(b_Zip)
                                    else
                                    set b_Zip to ""
                                    set b_gZip to ""
                                end if
                                if (business country of selectedItem) is not missing value then
                                    set b_Cou to (business country of selectedItem)
                                    set b_gCou to my encodedURL(b_Cou)
                                    else
                                    set b_Cou to ""
                                    set b_gCou to ""
                                end if
                                set b_Addr to b_Str & "<br/>" & b_Cit & ", " & b_Sta & "  " & b_Zip & "<br/>" & b_Cou
                                
                                --GOOGLE MAPS LOCATION IN URL
                                set b_gString to b_gStr & "," & b_gCit & "," & b_gSta & "," & b_gZip & "," & b_gCou
                                set b_GMAP to "http://maps.google.com/maps?q=" & b_gString
                                set b_String to "<strong>Work: </strong><br /><a href=\"" & b_GMAP & "\">" & b_Addr & "</a><br /><br />"
                                copy b_String to end of list_Addr
                            end if
                            
                            (*HOME *)
                            if home street address of selectedItem is not missing value then
                                set h_Str to (home street address of selectedItem)
                                set h_gStr to my encodedURL(h_Str)
                                if (home city of selectedItem) is not missing value then
                                    set h_Cit to (home city of selectedItem)
                                    set h_gCit to my encodedURL(h_Cit)
                                    else
                                    set h_Cit to ""
                                    set h_gCit to ""
                                end if
                                if (home state of selectedItem) is not missing value then
                                    set h_Sta to (home state of selectedItem)
                                    set h_gSta to my encodedURL(h_Sta)
                                    else
                                    set h_Sta to ""
                                    set h_gSta to ""
                                end if
                                if (home zip of selectedItem) is not missing value then
                                    set h_Zip to (home zip of selectedItem)
                                    set h_gZip to my encodedURL(h_Zip)
                                    else
                                    set h_Zip to ""
                                    set h_gZip to ""
                                end if
                                if (home country of selectedItem) is not missing value then
                                    set h_Cou to (home country of selectedItem)
                                    set h_gCou to my encodedURL(h_Cou)
                                    else
                                    set h_Cou to ""
                                    set h_gCou to ""
                                end if
                                set h_Addr to h_Str & "<br/>" & h_Cit & ", " & h_Sta & "  " & h_Zip & "<br/>" & h_Cou
                                
                                --GOOGLE MAPS LOCATION IN URL
                                set h_gString to h_gStr & "," & h_gCit & "," & h_gSta & "," & h_gZip & "," & h_gCou
                                set h_GMAP to "http://maps.google.com/maps?q=" & h_gString
                                set h_String to "<strong>Home: </strong><br /><a href=\"" & h_GMAP & "\">" & h_Addr & "</a><br />"
                                copy h_String to end of list_Addr
                            end if
                        end try
                        
                        --RIGHT SIDE (DATA FIELDS)
                        set r_1 to (display name of selectedItem)
                        set r_2 to (list_Addresses as string)
                        set r_3 to (list_Phone as text)
                        set r_4 to (list_Addr as text)
                        
                        --EXPORT VCARD DATA
                        set vcard_data to (vcard data of selectedItem)
                        set vcard_extension to ".vcf"
                        set item_Created to (current date)
                        
                        --THE NOTES
                        try
                            set the_notes to plain text note of selectedItem
                            if the_notes is missing value then set the_notes to ""
                        end try
                        
                        --ASSEMBLE THE TEMPLATE
                        set theContent to my make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
                        set the_vCard to my write_File(r_1, vcard_data, vcard_extension)
                        
                        set theHTML to true
                        set EVTitle to r_1
                        
                    (* // TASK ITEM *)
                    else if the_class is "task" then
                        
                        --PREPARE THE TEMPLATE
                        --LEFT SIDE (FORM FIELDS)
                        set l_1 to "Note:"
                        set l_2 to "Priority:"
                        set l_3 to "Due Date:"
                        set l_4 to "Status:"
                        set l_5 to "Notes:"
                        
                        --RIGHT SIDE (DATA FIELDS)
                        set propClass to (class of selectedItem) as text
                        if propClass is "incoming message" then
                            set r_1 to (subject of selectedItem)
                            else
                            set r_1 to (name of selectedItem)
                        end if
                        set the_Priority to (priority of selectedItem)
                        if the_Priority is priority normal then set r_2 to "Normal"
                        if the_Priority is priority high then set r_2 to "High"
                        if the_Priority is priority low then set r_2 to "Low"
                        
                        set r_3 to (due date of selectedItem)
                        set item_Created to (current date)
                        
                        --TODO?
                        try
                            set todo_Flag to (todo flag of selectedItem) as text
                            set r_4 to my TITLECASE(todo_Flag)
                        end try
                        
                        --THE NOTES
                        try
                            if content of selectedItem is missing value then
                                set the_notes to plain text content of selectedItem
                                else
                                set the_notes to content of selectedItem
                            end if
                            
                        end try
                        if the_notes is missing value then set the_notes to ""
                        
                        --ASSEMBLE THE TEMPLATE
                        set theContent to my make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
                        
                        --EXPORT VCARD DATA
                        if propClass is not "incoming message" then
                            set vcard_extension to ".ics"
                            set vcard_data to (icalendar data of selectedItem)
                            set the_vCard to my write_File(r_1, vcard_data, vcard_extension)
                        end if
                        
                        set theHTML to true
                        set EVTitle to r_1
                        
                    (* // MESSAGE ITEM *)
                    else
                        --PREPARE THE TEMPLATE
                        --LEFT SIDE (FORM FIELDS)
                        set l_1 to "From: / To: / CC: "
                        set l_2 to "Subject:"
                        set l_3 to "Date:"
                        set l_4 to "Category:"
                        set l_5 to "Email Contents:"
                        
                        --GET EMAIL INFO
                        set the_Sender to (sender of selectedItem)
                        set s_Name to (address of the_Sender)
                        set s_Address to (address of the_Sender)
                        
                        --REPLACE WITH NAME, IF AVAILABLE
                        try
                            set s_Name to (name of the_Sender)
                        end try
                        
                        set sender_Link to "<a href=\"mailto:" & s_Address & "\">" & s_Name & " (" & s_Address & ")</a>"
                        
                        --GET CATEGORY INFO
                        set the_Cats to (category of selectedItem)
                        set list_Cats to {}
                        set count_Cat to (count of the_Cats)
                        repeat with the_Cat in the_Cats
                            set cat_Name to (name of the_Cat as text)
                            copy cat_Name to the end of list_Cats
                            if count_Cat > 1 then
                                copy ", " to the end of list_Cats
                                set count_Cat to (count_Cat - 1)
                                else
                                set count_Cat to (count_Cat - 1)
                            end if
                        end repeat
                        
                        --RIGHT SIDE (DATA FIELDS)
                        --set r_1 to "From: " & sender_Link & "<hr/>To: " & t_List & "<hr/>CC: " & c_List
                        
                        set r_1 to "<table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"table-layout:fixed;width:100%;border-style:hidden;border-width:0px\">
                        <tr style=\"\">
                        <td align=\"center\" valign=\"middle\" height=\"15\" width=\"100%\" style=\"width:100%;text-align:center;border-bottom-style:solid;border-bottom-color:rgb(0, 0, 0);border-bottom-width:1px;\"><b>" & "From: " & sender_Link & "</b></td>
                        </tr>
                        <tr style=\"\">
                        <td align=\"center\" valign=\"middle\" height=\"15\" width=\"100%\" style=\"width:100%;text-align:center;border-bottom-style:solid;border-bottom-color:rgb(0, 0, 0);border-bottom-width:1px;\"><b>" & "To: " & t_List & "</b></td>
                        </tr>
                        <tr style=\"\">
                        <td align=\"center\" valign=\"middle\" height=\"15\" width=\"100%\" style=\"width:100%;text-align:center;border-style:hidden;border-width:0px;\"><b>" & "CC: " & c_List & "</b></td>
                        </tr>
                        </table>"

                        
                        
                        set m_Sub to (subject of selectedItem)
                        if m_Sub is missing value then
                            set r_2 to "<No Subject>"
                        else
                            set r_2 to {subject of selectedItem}
                        end if
                        set r_3 to (time sent of selectedItem)
                        set r_4 to list_Cats
                        
                        set theID to id of selectedItem as string
                        set item_Created to r_3
                        set EVTitle to r_2
                        
                        --PROCESS EMAIL CONTENT
                        set m_Content to content of selectedItem
                        set theHTML to has html of selectedItem
                        --set m_Content to plain text content of selectedItem
                        --set theHTML to false
                        
                        
                        --IF PLAINTEXT EMAIL CONTENT…
                        if theHTML is false then
                            set theContent to "Name: " & s_Name & return & "Subject: " & r_2 & return & "Sent: " & r_3 & return & return & return & return & m_Content
                            --IF HTML EMAIL CONTENT…
                        else
                            set the_notes to m_Content
                            --ASSEMBLE THE TEMPLATE
                            set theContent to my make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
                            
                            if stripEmbeddedImages is "ON" then
                                --REMOVE ANY EMBEDDED IMAGE RERENCES
                                set theContent to my stripCID(theContent)
                            end if
                        end if
                    end if
                    
                    --CREATE NOTE IN EVERNOTE (FINALLY!)
                    
                    --if titleEditable is false or titleEditable is 0 then
                    if (state of mergeCheckbox as text) is "0" then
                        set titleInput to titleInputField's stringValue() as text
                        if titleInput is not "< multiple >" then
                            set EVTitle to titleInput
                        end if
                        if theHTML is true then
                            tell application "Evernote"
                                set theNote to create note with html theContent title EVTitle notebook EVnotebook
                                if EVTag is not {} then assign EVTag to theNote
                                set creation date of theNote to item_Created
                            
                                --ATTACH VCARD (IF PRESENT)
                                if the_vCard is not {} then tell theNote to append attachment file the_vCard
                            end tell
                        else
                            tell application "Evernote"
                                set theNote to create note with text theContent title EVTitle notebook EVnotebook
                                if EVTag is not {} then assign EVTag to theNote
                                set creation date of theNote to item_Created
                            
                                --ATTACH VCARD (IF PRESENT)
                                if the_vCard is not {} then tell theNote to append attachment file the_vCard
                            end tell
                        end if
                    else
                        set mergeContent to mergeContent & theContent & "<hr />"
                        tell application "Evernote"
                            if the_vCard is not {} then tell theNote to append attachment file the_vCard
                        end tell
                    --display dialog "would merge"
                    end if
                    
                    --IF ATTACHMENTS PRESENT, RUN ATTACHMENT SUBROUTINE
                    if theAttachments is not {} then my message_Attach(theAttachments, selectedItem, theNote)
                    
                    --ITEM HAS FINISHED! COUNT IT AS A SUCCESS AND RESET ATTACHMENTS!
                    set successCount to successCount + 1
                    set theAttachments to {}
                end repeat
                if (state of mergeCheckbox as text) is "1" then
                --if titleEditable is true or titleEditable is 1 then
                    tell application "Evernote"
                        if EVTag is not {} then assign EVTag to theNote
                        tell theNote to append html mergeContent
                    end tell
                    set successCount to 1
                end if
                -- SEND HERE IF MERGED
            end if
        end tell
    end item_Process
    
    
    (* 
     ======================================
     // UTILITY SUBROUTINES 
     ======================================
     *)
    --URL ENCODE
    on encodedURL(the_Word)
        set scpt to "php -r 'echo urlencode(\"" & the_Word & "\");'"
        return do shell script scpt
    end encodedURL
    
    --TITLECASE
    on TITLECASE(txt)
        return do shell script "python -c \"import sys; print unicode(sys.argv[1], 'utf8').title().encode('utf8')\" " & quoted form of txt
    end TITLECASE
    
    --SORT SUBROUTINE
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
    
    --REPLACE
    on replaceString(theString, theOriginalString, theNewString)
        set theNum to 0
        set {od, AppleScript's text item delimiters} to {AppleScript's text item delimiters, theOriginalString}
        set theStringParts to text items of theString
        if (count of theStringParts) is greater than 1 then
            set theString to text item 1 of theStringParts as string
            repeat with eachPart in items 2 thru -1 of theStringParts
                set theString to theString & theNewString & eachPart as string
                set theNum to theNum + 1
            end repeat
        end if
        set AppleScript's text item delimiters to od
        return theString
    end replaceString
    
    --REMOVE EMBEDDED IMAGE REFERENCES
    on stripCID(theContent)
        set theCommandString to "echo " & quoted form of theContent & " | sed 's/\"cid:.*\"/\"\"/'"
        set theResult to do shell script theCommandString
        return theResult
    end stripCID
    
    (* 
     ======================================
     // TAGGING SUBROUTINES
     ======================================
     *)
    
    --TAGGING
    on tagging_Dialog()
        set userInput to tagInputField's stringValue() as text
        set theDelims to {","}
        set userTag to my Tag_List(userInput, theDelims)
        
        --RESET, FINAL CHECK, AND FORMATTING OF TAGS
        set EVTag to {}
        set EVTag to my Tag_Check(userTag)
    end tagging_Dialog
    
    --TAG SELECTION SUBROUTINE
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
    
    --CREATES TAGS IF THEY DON'T EXIST
    on Tag_Check(theTags)
        tell application "Evernote"
            set finalTags to {}
            repeat with theTag in theTags
                
                -- TRIM LEADING SPACE, IF ANY
                if (the character 1 of theTag is " ") then set theTag to text 2 thru end of theTag as text
                
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
    
    (* 
     ======================================
     // NOTEBOOK SUBROUTINES
     ======================================
     *)
    
    --GET EVERNOTE'S DEFAULT NOTEBOOK
    on default_Notebook()
        tell application "Evernote"
            set get_defaultNotebook to every notebook whose default is true
            if EVnotebook is "" then
                set EVnotebook to name of (item 1 of get_defaultNotebook) as text
            end if
        end tell
    end default_Notebook
    
    --EVERNOTE NOTEBOOK SELECTION SUBROUTINE 
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
            set SelNotebook to choose from list of Folders_sorted with title "Select Evernote Notebook" with prompt ¬
            "Current Evernote Notebooks" OK button name "OK" cancel button name "New Notebook" (*USER SELECTION FROM NOTEBOOK LIST *)
            if (SelNotebook is false) then (*CREATE NEW NOTEBOOK OPTION *)
                set userInput to ¬
                text returned of (display dialog "Enter New Notebook Name:" default answer "")
                set EVnotebook to userInput
                else
                set EVnotebook to item 1 of SelNotebook
            end if
        end tell
    end Notebook_List
    
    (* 
     ======================================
     // ATTACHMENT SUBROUTINES 
     =======================================
     *)
    
    --CLEAN TITLE FOR FILENAME
    on clean_Title(rawFileName)
        set previousDelimiter to AppleScript's text item delimiters
        set potentialName to rawFileName
        set legalName to {}
        set illegalCharacters to {".", ",", "/", ":", "[", "]"}
        repeat with thisCharacter in the characters of potentialName
            set thisCharacter to thisCharacter as text
            if thisCharacter is not in illegalCharacters then
                set the end of legalName to thisCharacter
                else
                set the end of legalName to "_"
            end if
        end repeat
        return legalName
    end clean_Title
    
    --WRITE THE FILE
    on write_File(r_1, vcard_data, vcard_extension)
        set ExportFolder to ((path to desktop folder) & "Temp Export From Outlook:") as string
        set SaveLoc to my f_exists(ExportFolder)
        set FileName to (my clean_Title(r_1) & vcard_extension)
        set theFileName to (ExportFolder & FileName)
        try
            open for access file theFileName with write permission
            write vcard_data to file theFileName as string
            close access file theFileName
            return theFileName
            
            on error errorMessage
            log errorMessage
            try
                close access file theFileName
            end try
        end try
    end write_File
    
    --FOLDER EXISTS
    on f_exists(ExportFolder)
        try
            set myPath to (path to home folder)
            get ExportFolder as alias
            set SaveLoc to ExportFolder
            on error
            tell application "Finder" to make new folder with properties {name:"Temp Export From Outlook"}
        end try
    end f_exists
    
    --ATTACHMENT PROCESSING
    on message_Attach(theAttachments, selectedItem, theNote)
        
        tell application id "com.microsoft.Outlook"
            --MAKE SURE TEXT ITEM DELIMITERS ARE DEFAULT
            set AppleScript's text item delimiters to ""
            
            --TEMP FILES PROCESSED ON THE DESKTOP
            set ExportFolder to ((path to desktop folder) & "Temp Export From Outlook:") as string
            set SaveLoc to my f_exists(ExportFolder)
            
            --PROCESS THE ATTCHMENTS
            set attCount to 0
            repeat with theAttachment in theAttachments
                
                set theFileName to ExportFolder & theAttachment's name
                try
                    save theAttachment in file theFileName
                end try
                tell application "Evernote"
                    tell theNote to append attachment file theFileName
                end tell
                
                --SILENT DELETE OF TEMP FILE
                set trash_Folder to path to trash folder from user domain
                do shell script "mv " & quoted form of POSIX path of theFileName & space & quoted form of POSIX path of trash_Folder
                
            end repeat
        end tell
        
    end message_Attach
    
    --SILENT DELETE OF TEMP FOLDER (THANKS MARTIN MICHEL!)
    on trashfolder(SaveLoc)
        try
            set trashfolderpath to ((path to trash) as Unicode text)
            set srcfolderinfo to info for (SaveLoc as alias)
            set srcfoldername to name of srcfolderinfo
            set SaveLoc to (SaveLoc as alias)
            set SaveLoc to (quoted form of POSIX path of SaveLoc)
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
     // NOTIFICATION SUBROUTINES
     ======================================
     *)
    
    --ANNOUNCE THE COUNT OF TOTAL ITEMS TO EXPORT
    on process_Items(itemNum, attNum)
        
        set attPlural to "s"
        
        if attNum = 0 then
            set attNum to "No"
            else if attNum is 1 then
            set attPlural to ""
        end if
        
        set notifTitle to "Import to Evernote"
        set notifSubtitle to "Started - Processing " & itemNum & " " & " Item(s)"
        set notifText to "Including " & attNum & " Attachment" & attPlural
        
        display notification notifText with title notifTitle subtitle notifSubtitle
        
    end process_Items
    
    --NOTIFY RESULTS
    on notify_results(successCount)
        if EVnotebook is "" then set EVnotebook to "Default"
        
        set notifTitle to ""
        set notifSubtitle to ""
        set notifMessage to ""
        
        set Plural_Test to (successCount) as number
        
        if Plural_Test is -1 then
            set notifTitle to "Outlook to Evernote"
            set notifSubtitle to "Failure Notification"
            set notifMessage to "Import failure - No Items Selected in Outlook!"
            
        else if Plural_Test is 0 then
            set notifTitle to "Outlook to Evernote"
            set notifSubtitle to "Failure Notification"
            set notifMessage to "No Items Exported From Outlook!"
            
        else
            set notifTitle to "Outlook to Evernote"
            set notifSubtitle to "Success Notification"
            set notifMessage to "Exported " & successCount & " item(s) to " & EVnotebook & " notebook"
            
        end if
        
        display notification notifMessage with title notifTitle subtitle notifSubtitle
        
        set itemNum to "0"
        set EVnotebook to ""
        
    end notify_results
    
    (* 
     ======================================
     // TEMPLATE SUBROUTINES
     ======================================
     *)
    on make_Template(l_1, l_2, l_3, l_4, l_5, r_1, r_2, r_3, r_4, the_notes)
        --MAKE TASK TEMPLATE
        set the_Template to "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>
        <!DOCTYPE en-note SYSTEM \"http://xml.evernote.com/pub/enml2.dtd\">
        <en-note>
        <div>
        <table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"table-layout:fixed;width:100%;border-style:solid;border-color:rgb(0, 0, 0);border-collapse:collapse;border-width:1px;\">
        <tr style=\"\">
        <td align=\"center\" valign=\"middle\" height=\"15\" width=\"221\" style=\"width:221px;text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & l_1 & "</b></td>
        <td align=\"center\" valign=\"middle\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & r_1 & "</b></td>
        </tr>
        <tr style=\"\">
        <td align=\"center\" valign=\"middle\" height=\"15\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & l_2 & "</b></td>
        <td align=\"center\" valign=\"middle\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & r_2 & "</b></td>
        </tr>
        <tr style=\"\">
        <td align=\"center\" valign=\"middle\" height=\"15\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & l_3 & "</b></td>
        <td align=\"center\" valign=\"middle\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & r_3 & "</b></td>
        </tr>
        <tr style=\"\">
        <td align=\"center\" valign=\"middle\" height=\"15\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & l_4 & "</b></td>
        <td align=\"center\" valign=\"middle\" style=\"text-align:center;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\"><b>" & r_4 & "</b></td>
        </tr>
        <tr style=\"\">
        <td colspan=\"2\" style=\"padding:5px;border-style:solid;border-color:rgb(0, 0, 0);border-width:1px;\">
        <b>" & l_5 & "</b>
        <br /><br />" & the_notes & "
        </td></tr>
        </table>
        </div>
        <div><br/></div>
        </en-note>"
    end make_Template
        

end script