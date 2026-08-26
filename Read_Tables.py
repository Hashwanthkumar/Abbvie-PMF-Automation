import pandas as pd

tables = pd.read_html("email.html")

print(f"Total tables found: {len(tables)}")

for i, table in enumerate(tables):
    print(f"\n--- TABLE {i} ---")
    print(table.head())