tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {subject:"Subject", content:"This is the body of the email"}
    make new recipient at newMessage with properties {email address:{address:"zahrakheiroddin1997@gmail.com"}}
    send newMessage
end tell
