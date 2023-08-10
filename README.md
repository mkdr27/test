import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

# Calculate the date range for today
today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
tomorrow = today + timedelta(days=1)

today_emails = inbox.Items.Restrict("[ReceivedTime] >= '" + today.strftime('%m/%d/%Y %H:%M %p') + "' AND [ReceivedTime] < '" + tomorrow.strftime('%m/%d/%Y %H:%M %p') + "'")

print(f"Total emails received today: {len(today_emails)}")

for email in today_emails:
    print(f"\nSubject: {email.Subject}")
    print(f"Received Time: {email.ReceivedTime}")

    # Accessing subthreads might not be straightforward using COM objects.
    # You might need to track the email thread using ConversationID and find related emails.

print("\nScript completed.")
