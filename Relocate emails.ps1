function get-folder {
    # this function return the folder object from the given folder path
    # Folder path is in this format: 'folder/subfolder1/subfolder2/...'
    param (
        $Folderpath
    )
    $Folder_arr = $Folderpath.Split("\")
    $folder = $main_Mailbox # this is a global variable that refer to the mailbox that you're working with
    foreach ($folder_name in $Folder_arr){
        $folder = $folder.Folders.Item($folder_name)
    }
    $folder
}

### SET UP VARIABLES 
$mailbox_name = 'example@email.com' # If this is a shared mailbox, check your Outlook for the mailbox's name
$intruction_path = 'path\Instruction.csv'
$record_path = 'path\Email relocation record.csv'
$matched_Emails = @()


########## GET ALL NEW EMAILS TODAY ####
#Access Outlook application
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")

# Access the mailbox
$main_Mailbox = $namespace.Folders.Item($mailbox_name)
$inbox = $main_Mailbox.Folders.Item('Inbox') # Assumming your mailbox have similar default folders as most Outlook users
$newEmails = $inbox.items|where-object SentOn -ge (get-date).Date

############ MOVING THE MATCHING EMAILS ####
# Import instruction
$instruction = Import-Csv $intruction_path
# Matching emails with the instruction list
foreach ($email in $newEmails){
    foreach ($row in $instruction){
        $emailsubject = $email.Subject
        $rowsubject = $row.Subject
        if ($emailsubject -like "*$rowsubject*"){
            $dtsn_path = $row.DestinationFolder
            # Move emails
            $dstn = get-folder($dtsn_path)
            $email.move($dstn)
            # Add the moved email to list to record later
            $email| add-member -MemberType NoteProperty -Name "Folderpath" -Value $dtsn_path
            $timestamp = Get-Date
            $email| add-member -MemberType NoteProperty -Name "MovedOn" -Value $timestamp
            $matched_Emails += $email
        } 
    }
}
# Export the records
$matched_Emails|Select-Object -Property Subject, SenderName, SentOn, Folderpath, MovedOn|
    Export-Csv -path $record_path -NoTypeInformation -Append


