# Automated Outlook Email Sender with Python and AppleScript

This project demonstrates how to automatically send emails using Outlook on macOS by combining **Python** and **AppleScript**. The Python script collects tasks from the user, generates the email content, and then calls an AppleScript to send the email via Outlook.

## Features

- Collects 6 tasks from the user and formats them into an email body.
- Generates a timestamped email subject and body.
- Uses AppleScript to automate the email sending process via Microsoft Outlook.
- Optionally saves the email content in a text file for reference.

## Requirements

- **macOS** with **Microsoft Outlook** installed.
- **Python 3** installed.
- Basic knowledge of using **AppleScript** for macOS automation.

## Getting Started

### Step 1: Clone the repository

```bash
git clone https://github.com/yourusername/outlook-email-automation.git
cd outlook-email-automation
```

### Step 2: Create the AppleScript

You will need to create an AppleScript file that will handle sending the email via Outlook.

1. Open **Script Editor** on macOS.
2. Copy and paste the following AppleScript code into the editor:

```applescript
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {subject:"Subject", content:"This is the body of the email"}
    make new recipient at newMessage with properties {email address:{address:"zahrakheiroddin1997@gmail.com"}}
    send newMessage
end tell
```

3. Save the file as `send_outlook_email.scpt` on your desktop or in any directory you prefer.

### Step 3: Install Required Python Libraries

Ensure that you have Python installed on your macOS. If not, you can download it from [python.org](https://www.python.org/).

There are no additional dependencies required for this script.

### Step 4: Run the Python Script

The Python script will generate the email content and trigger the AppleScript to send the email via Outlook.

1. Open the terminal and navigate to the folder containing the Python script.

```bash
cd path/to/your/script
```

2. Run the script:

```bash
python send_email.py
```

3. The script will ask you to input six tasks. After that, it will:
   - Generate the email content.
   - Save the email content in a timestamped text file.
   - Use the AppleScript to send the email via Microsoft Outlook.

### Example Output

- **Email Content Example:**

```
Date and Time: 2024-10-17 10:45:00

Hi Zahra,

Here’s what I suggest for today’s tasks:
1. Task 1
2. Task 2
3. Task 3
4. Task 4
5. Task 5
6. Task 6

With love,
WIM
```

- **Saved File:**

The email content will be saved in a file named `tasks_YYYYMMDD_HHMMSS.txt` in the script directory.

## Customization

### Modify the Email Body

You can modify the email body format in the Python script to include more personalized content or change the formatting of the tasks.

### Update the AppleScript

If you'd like to send emails to multiple recipients or change the email format, you can modify the AppleScript to suit your needs.

## Troubleshooting

- Ensure that **Outlook** is installed and set up on your Mac.
- Verify that the `osascript` command works from the terminal. This command is used to run the AppleScript.
- Check that the correct recipient email address is provided in both the Python script and AppleScript.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Contributing

Feel free to open issues or submit pull requests if you’d like to contribute to this project.
