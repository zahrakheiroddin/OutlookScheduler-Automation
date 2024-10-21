import subprocess
from datetime import datetime

# Function to trigger AppleScript
def send_via_outlook(subject, body):
    script = f'''
    tell application "Microsoft Outlook"
        set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
        make new recipient at newMessage with properties {{email address:{{address:"zahrakheiroddin1997@yahoo.com"}}}}
        send newMessage
    end tell
    '''
    subprocess.run(['osascript', '-e', script])

# Get the current date, time, and day of the week
now = datetime.now()
current_time = now.strftime("%Y-%m-%d %H:%M:%S")

# Ask the user for six tasks
tasks = []
print("Enter your tasks for today:")
for i in range(1, 7):
    task = input(f"Task {i}: ")
    tasks.append(task)

# Format the tasks into a string
tasks_formatted = "\n".join(f"{i}. {task}" for i, task in enumerate(tasks, start=1))

# Combine email body
subject = f"{current_time} > Hi Zahra"
body = f"""Date and Time: {current_time}

Hi Zahra,

Here’s what I suggest for today’s tasks:
{tasks_formatted}

With love,
WIM
"""

# Optionally, save the email content to a file
filename = f"tasks_{now.strftime('%Y%m%d_%H%M%S')}.txt"
with open(filename, 'w') as file:
    file.write(body)

print(f"Email content saved to {filename}")

# Send email using Outlook
send_via_outlook(subject, body)

print("Email sent successfully via Outlook!")
