    import win32com.client
    from datetime import datetime, timedelta
    import pandas as pd

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access "Sent Items"
    sent_folder = outlook.GetDefaultFolder(5)
    messages = sent_folder.Items
    messages.Sort("[SentOn]", True)

    def clean_text(text):
        if isinstance(text, str):
            return text.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
        return ""

    # Filter by date
    cutoff_date = datetime.now() - timedelta(days=2*365)

    # Prepare data
    email_data = []

    for message in messages:
        try:
            sent_time = message.SentOn
            if sent_time.replace(tzinfo=None) >= cutoff_date:
                # Recipients (To, CC, BCC)
                recipients = message.Recipients
                to_emails = []
                for r in recipients:
                    try:
                        addr_entry = r.AddressEntry
                        if addr_entry.Type == "SMTP":
                            to_emails.append(addr_entry.Address)
                        else:
                            # For internal Outlook addresses, try to get SMTP address
                            to_emails.append(addr_entry.GetExchangeUser().PrimarySmtpAddress)
                    except:
                        to_emails.append(r.Address)


                email_data.append({
                    "To Emails": clean_text(", ".join(to_emails)),
                    "Subject": clean_text(message.Subject),
                    "SentOn": sent_time.strftime("%Y-%m-%d %H:%M:%S"),
                    "Body Preview": clean_text(message.Body[:100])
                })

        except AttributeError:
            continue

    # Save to Excel
    df = pd.DataFrame(email_data)
    df.to_csv("outbox_emails.csv")
    print(f"Exported {len(df)} sent emails.")
