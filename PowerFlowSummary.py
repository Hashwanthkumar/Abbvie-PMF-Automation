import win32com.client
import pandas as pd
import json
import re

from datetime import datetime, timedelta, timezone

# ====================================================
# CONNECT TO OUTLOOK
# ====================================================

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Inbox -> PRM -> Abbvie PF
inbox = outlook.GetDefaultFolder(6)
prm_folder = inbox.Folders.Item("PRM")
Abbvie_folder = prm_folder.Folders.Item("Abbvie PF")

messages = Abbvie_folder.Items
messages.Sort("[ReceivedTime]", True)

# ====================================================
# PREVIOUS DAY RANGE
# Example:
# 26-Jul-2026 12:00:00 AM
# to
# 26-Jul-2026 11:59:59 PM
# ====================================================

today = datetime.now(timezone.utc).date()

start_time = datetime.combine(
    today - timedelta(days=1),
    datetime.min.time(),
    tzinfo=timezone.utc
)

end_time = datetime.combine(
    today,
    datetime.min.time(),
    tzinfo=timezone.utc
)

report_date = (today - timedelta(days=1)).strftime("%d-%b-%Y")

print(f"\nGenerating report for {report_date}")

all_failed_flows = []

# ====================================================
# READ EMAILS
# ====================================================

for mail in messages:

    try:

        if "[prod] IQ6-Abbvie PowerFlow Monitoring" in str(mail.Subject):

            if start_time <= mail.ReceivedTime < end_time:

                tables = pd.read_html(mail.HTMLBody)

                # Failed Flow Details Table
                if len(tables) >= 3:

                    failed_df = tables[2]

                    all_failed_flows.append(failed_df)

    except Exception as e:
        print("Error:", e)

# ====================================================
# NO FAILED FLOWS FOUND
# ====================================================

if not all_failed_flows:

    outlook_app = win32com.client.Dispatch(
        "Outlook.Application"
    )

    mail = outlook_app.CreateItem(0)

    mail.To = "your_email@iqvia.com"

    mail.Subject = (
        f"[IQ6-Abbvie] Top Failed Flows Report | "
        f"{report_date} | Failures > 5"
    )

    mail.Body = (
        f"No failed flows were found for {report_date}."
    )

    mail.Send()

    print("No failures found.")
    quit()

# ====================================================
# COMBINE DATA
# ====================================================

final_df = pd.concat(
    all_failed_flows,
    ignore_index=True
)

# Remove dummy rows
final_df = final_df[
    final_df["Flow Name"]
    != "No failed flows in lookback window."
]

# Convert Failed Runs to numeric
final_df["Failed Runs"] = pd.to_numeric(
    final_df["Failed Runs"],
    errors="coerce"
).fillna(0)

# ====================================================
# AGGREGATE BY FLOW
# ====================================================

summary_df = final_df.groupby(
    "Flow Name"
).agg(
    {
        "Failed Runs": "sum",
        "Error Message": "first"
    }
).reset_index()

# Only flows with + failures
summary_df = summary_df[
    summary_df["Failed Runs"] >= 5
]

# Sort descending
summary_df = summary_df.sort_values(
    by="Failed Runs",
    ascending=False
)

# ====================================================
# BUILD REPORT
# ====================================================

summary_text = f"""
PowerFlow Monitoring Report

Report Date : {report_date}

Time Window :
12:00 AM to 11:59 PM

Flows with more than 5 failures

"""

print("\nTOP FAILED FLOWS\n")

for _, row in summary_df.iterrows():

    flow_name = row["Flow Name"]
    failed_count = int(row["Failed Runs"])

    error_message = str(row["Error Message"])

    # Extract readable message
    try:

        cleaned = error_message.replace(
            "'",
            '"'
        )

        json_data = json.loads(cleaned)

        if "message" in json_data:

            error_message = json_data["message"]

    except:

        match = re.search(
            r'"message"\s*:\s*"([^"]+)"',
            error_message
        )

        if match:

            error_message = match.group(1)

    print(
        f"{flow_name} : {failed_count}"
    )

    summary_text += f"""
Flow Name    : {flow_name}
Failed Count : {failed_count}
Error        : {error_message}

------------------------------------------------------------

"""

total_failed_runs = int(
    summary_df["Failed Runs"].sum()
)

summary_text += f"""

Total Failed Runs (5+ Failure Flows) : {total_failed_runs}

"""

# ====================================================
# SEND EMAIL
# ====================================================

outlook_app = win32com.client.Dispatch(
    "Outlook.Application"
)

mail = outlook_app.CreateItem(0)

# UPDATE RECIPIENTS
mail.To = "hashwanth.kumar@iqvia.com"

# Optional
# mail.CC = "team@iqvia.com"

mail.Subject = (
    f"[IQ6-Abbvie] Top Failed Flows Report | "
    f"{report_date} | Failures > 5"
)

mail.Body = summary_text

mail.Send()

print("\nEmail sent successfully!")

print(
    f"\nTotal Failed Runs "
    f"(5+ Failure Flows): {total_failed_runs}"
)