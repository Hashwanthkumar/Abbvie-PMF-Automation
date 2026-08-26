import pandas as pd
from datetime import datetime, timedelta

# Read Excel file
df = pd.read_excel("FailedFlows.xlsx")

# Convert Date column
df["Date"] = pd.to_datetime(df["Date"])

# Last 24 Hours
last_24_hours = datetime.now() - timedelta(hours=24)

# Filter last 24 hours data
df = df[df["Date"] >= last_24_hours]

# Remove unwanted rows
df = df[
    df["Flow Name"] != "No failed flows in lookback window."
]

# Convert Failed Runs to numeric
df["Failed Runs"] = pd.to_numeric(
    df["Failed Runs"],
    errors="coerce"
).fillna(0)

# Group by Flow Name
summary = df.groupby("Flow Name")["Failed Runs"].sum()

print("\nPowerFlow Failure Summary\n")

for flow, count in summary.items():
    print(f"{flow} : {int(count)}")

print("\n--------------------------------")
print("Total Failed Runs:", int(summary.sum()))
print("--------------------------------")