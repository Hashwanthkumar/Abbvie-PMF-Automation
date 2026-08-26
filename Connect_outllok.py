import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Inbox
inbox = outlook.GetDefaultFolder(6)

# PRM folder
prm_folder = inbox.Folders.Item("PRM")

# Abbvie PF folder
Abbvie_folder = prm_folder.Folders.Item("Abbvie PF")

# Read emails
messages = Abbvie_folder.Items
messages.Sort("[ReceivedTime]", True)

for mail in messages:
    try:
        if "[prod] IQ6-Abbvie PowerFlow Monitoring" in str(mail.Subject):

            print("Subject:", mail.Subject)
            print("Received:", mail.ReceivedTime)

            # Save HTML body
            with open("email.html", "w", encoding="utf-8") as f:
                f.write(mail.HTMLBody)

            print("Email saved successfully!")
            break

    except Exception as e:
        print("Error:", e)