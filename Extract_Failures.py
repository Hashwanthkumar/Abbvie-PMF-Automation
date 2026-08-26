import pandas as pd
from datetime import datetime

tables = pd.read_html("email.html")

failed_df = tables[2]

failed_df["Date"] = datetime.now()

file_name = "FailedFlows.xlsx"

try:
    existing = pd.read_excel(file_name)

    final_df = pd.concat(
        [existing, failed_df],
        ignore_index=True
    )

except:
    final_df = failed_df

final_df.to_excel(file_name, index=False)

print("Data saved to FailedFlows.xlsx")