on run argv
  try
    set attachmentPath to item 1 of argv
    if attachmentPath is not "" then
      set attachmentFile to attachmentPath as POSIX file
    else
      set attachmentFile to missing value
    end if
  on error
    set attachmentFile to missing value
  end try

  set messageSubject to item 2 of argv
  set messageBody to item 3 of argv
  set recipientName to item 4 of argv
  set recipientEmail to item 5 of argv

  tell application "Microsoft Outlook"
    activate
    set outgoingMessage to make new outgoing message with properties {subject:messageSubject, plain text content:messageBody}

    if attachmentFile is not missing value then
      tell outgoingMessage
        make new attachment with properties {file:attachmentFile}
      end tell
    end if

    make new recipient at outgoingMessage with properties {email address:{name:recipientName, address:recipientEmail}}
    send outgoingMessage
  end tell
end run
