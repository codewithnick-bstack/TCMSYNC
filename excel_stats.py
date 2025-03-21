import pandas as pd
from tabulate import tabulate
from termcolor import colored

# ✅ Constants
EXCEL_FILE = "Live ++ Automation Data.xlsx"
SHEET_NAME = "TCM Bug capture automation trac"

# ✅ 1. Read Excel Sheet
def read_excel(file, sheet):
    print("📊 Reading Excel sheet...")
    df = pd.read_excel(file, sheet_name=sheet, engine='openpyxl')

    # Normalize column names and lowercase everything
    df.columns = df.columns.str.strip()
    df['Status'] = df['Status'].str.lower()
    df['Automated'] = df['Automated'].str.lower()

    print(f"✅ Loaded {len(df)} test cases from Excel.")
    return df

# ✅ 2. Generate Stats
def generate_stats(df):
    print("\n📊 Generating Excel Stats...\n")

    total_cases = len(df)

    # Automation Stats
    automation_stats = df['Automated'].value_counts().reset_index()
    automation_stats.columns = ['Automation', 'Count']

    # Status Stats
    status_stats = df['Status'].value_counts().reset_index()
    status_stats.columns = ['Status', 'Count']

    # Print Stats
    print(colored(f"\n✅ Total Test Cases: {total_cases}", "cyan"))

    print("\n⚙️ Automation Stats:")
    print(tabulate(automation_stats, headers=["Automation", "Count"], tablefmt="grid", showindex=False))

    print("\n📌 Status Stats:")
    print(tabulate(status_stats, headers=["Status", "Count"], tablefmt="grid", showindex=False))

# ✅ 3. Main Execution
def main():
    # Step 1: Read Excel
    excel_df = read_excel(EXCEL_FILE, SHEET_NAME)

    # Step 2: Generate and display stats
    generate_stats(excel_df)

if __name__ == "__main__":
    main()
