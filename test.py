import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

# Calculate the date range for the last 7 days
end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
start_date = end_date - timedelta(days=7)

# Formulate the date filter for the Restrict method
date_filter = "[ReceivedTime] >= '" + start_date.strftime('%m/%d/%Y %H:%M %p') + "' AND [ReceivedTime] <= '" + end_date.strftime('%m/%d/%Y %H:%M %p') + "'"

recent_emails = inbox.Items.Restrict(date_filter)

for email in recent_emails:
    subject = email.Subject
    unique_id = email.EntryID
    from_address = email.SenderEmailAddress
    to_address = email.To

    print("Subject:", subject)
    print("Unique ID:", unique_id)
    print("From Address:", from_address)
    print("To Address:", to_address)
    print("-" * 50)

print("Script completed.")
