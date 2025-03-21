import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from termcolor import colored
from tabulate import tabulate
from dotenv import load_dotenv

# ✅ Load Environment Variables
load_dotenv()

# ✅ Constants
EXCEL_FILE = os.getenv('EXCEL_FILE')
SHEET_NAME = os.getenv('SHEET_NAME')
USERNAME = os.getenv("USERNAME")
ACCESS_KEY = os.getenv("ACCESS_KEY")
PROJECT_ID = os.getenv("PROJECT_ID")
FOLDER_IDS = list(map(int, os.getenv("FOLDER_IDS").split(',')))

API_URL = f"https://test-management.browserstack.com/api/v2/projects/{PROJECT_ID}/test-cases"

# ✅ 1. Read Excel Sheet
def read_excel(file, sheet):
    print("📊 Reading Excel sheet...")
    df = pd.read_excel(file, sheet_name=sheet, engine='openpyxl')

    # Normalize column names
    df.columns = df.columns.str.strip()

    print(f"✅ Loaded {len(df)} test cases from Excel.")
    return df

# ✅ 2. Fetch Test Cases from Multiple Folders
def fetch_test_cases():
    print("\n🔗 Fetching test cases from multiple folders...")

    all_test_cases = []

    for folder_id in FOLDER_IDS:
        print(f"📂 Fetching test cases from folder ID: {folder_id}")
        page = 1

        while True:
            params = {"folder_id": folder_id, "p": page, "status": 'active'}
            response = requests.get(API_URL, params=params, auth=HTTPBasicAuth(USERNAME, ACCESS_KEY))

            if response.status_code != 200:
                print(f"❌ Failed to fetch data from folder {folder_id}: {response.status_code}")
                print(response.text)
                break

            data = response.json()
            all_test_cases.extend(data.get("test_cases", []))

            if not data["info"]["next"]:
                break

            page += 1
        print(f"✅ Fetched {len(all_test_cases)} test cases from folder {folder_id}.")

    print(f"✅ Fetched {len(all_test_cases)} test cases from {len(FOLDER_IDS)} folders.")
    return all_test_cases

# ✅ 3. Convert API Data to DataFrame
def api_to_dataframe(api_data):
    """Convert API response to a Pandas DataFrame."""
    print("\n🔄 Converting API data to DataFrame...")
    data = []

    for tc in api_data:
        data.append({
            "Test Case ID": tc.get("identifier"),
            "Priority": tc.get("priority").lower(),
            "Automated": tc.get("automation_status").lower().replace('_', ' '),
            "Case Type": tc.get("case_type", "N/A").lower(),
            "Status": tc.get("status", "N/A").lower(),
        })

    df = pd.DataFrame(data)
    print(f"✅ Converted {len(df)} test cases to DataFrame.")
    return df

# ✅ 4. Compare Excel and API Data
def compare_data(excel_df, api_df):
    print("\n🔎 Comparing Excel and API data...")

    # Normalize column names and lowercase everything
    excel_df.columns = excel_df.columns.str.strip()
    api_df.columns = api_df.columns.str.strip()

    # Lowercase Excel data
    excel_df['Status'] = excel_df['Status'].str.lower()
    excel_df['Automated'] = excel_df['Automated'].str.lower()

    # Identify missing cases
    excel_ids = set(excel_df['Test Case ID'])
    api_ids = set(api_df['Test Case ID'])

    missing_in_excel = api_df[api_df['Test Case ID'].isin(api_ids - excel_ids)]
    missing_in_api = excel_df[excel_df['Test Case ID'].isin(excel_ids - api_ids)]

    # Find mismatched test cases
    differences = []
    common_ids = excel_ids & api_ids

    for case_id in common_ids:
        excel_row = excel_df[excel_df['Test Case ID'] == case_id].iloc[0].to_dict()
        api_row = api_df[api_df['Test Case ID'] == case_id].iloc[0].to_dict()

        diff = {}

        # Compare Status (Excel → API Priority)
        if excel_row['Status'] != api_row['Priority']:
            diff['Status'] = f"Excel: {excel_row['Status']} | API: {api_row['Priority']}"

        # Compare Automated field
        if excel_row['Automated'] != api_row['Automated']:
            diff['Automated'] = f"Excel: {excel_row['Automated']} | API: {api_row['Automated']}"

        if diff:
            differences.append({
                "Test Case ID": case_id,
                "Case Type": api_row["Case Type"],
                "Priority": api_row["Priority"],
                **diff
            })

    # ✅ Print Results in Table Format
    print("\n✅ Results:\n")

    # Print missing in Excel (Green)
    if not missing_in_excel.empty:
        print(colored("\n✅ Test Cases in API but Missing in Excel:", "green"))
        table = missing_in_excel[['Test Case ID', 'Case Type', 'Priority', 'Status', 'Automated']].values.tolist()
        print(colored(tabulate(table, headers=["Test Case ID", "Case Type", "Priority", "Status", "Automated"], tablefmt="grid"), "green"))

    # Print missing in API (Red)
    if not missing_in_api.empty:
        print(colored("\n❌ Test Cases in Excel but Missing in API:", "red"))
        table = missing_in_api[['Test Case ID', 'Status', 'Automated']].values.tolist()
        print(colored(tabulate(table, headers=["Test Case ID", "Status", "Automated"], tablefmt="grid"), "red"))

    # Print mismatched cases (Yellow)
    if differences:
        print(colored("\n⚠️ Mismatched Test Cases:", "yellow"))
        table = [[d["Test Case ID"], d["Case Type"], d["Priority"], d.get("Status", ""), d.get("Automated", "")] for d in differences]
        print(colored(tabulate(table, headers=["Test Case ID", "Case Type", "Priority", "Status (Excel | API)", "Automated (Excel | API)"], tablefmt="grid"), "yellow"))
    else:
        print(colored("\n✅ No mismatched test cases!", "green"))

    # comparison statistics
    print(f"\n✅ Total Test Cases in Excel: {len(excel_df)}")
    print(f"✅ Total Test Cases in API: {len(api_df)}")
    print(f"✅ Total Test Cases in Excel but Missing in API: {len(missing_in_api)}")
    print(f"✅ Total Test Cases in API but Missing in Excel: {len(missing_in_excel)}")
    print(f"✅ Total Mismatched Test Cases: {len(differences)}")

# ✅ 5. Generate Stats and Pivot Tables
def generate_stats(excel_df, api_df):
    print("\n📊 Generating detailed stats...")

    # 🔥 Excel Stats
    print("\n📊 Excel Stats:")
    excel_pivot = pd.pivot_table(excel_df, index='Automated', columns='Status', aggfunc='size', fill_value=0)
    print("\nExcel Pivot Table:")
    print(tabulate(excel_pivot, headers='keys', tablefmt='grid'))

    # 🔥 API Stats
    print("\n📊 API Stats:")
    print(api_df.describe(include='all'))
    api_pivot = pd.pivot_table(api_df, index='Automated', columns='Priority', aggfunc='size', fill_value=0)
    print("\nAPI Pivot Table:")
    print(tabulate(api_pivot, headers='keys', tablefmt='grid'))

# ✅ 6. Main Execution
def main():
    excel_df = read_excel(EXCEL_FILE, SHEET_NAME)
    test_cases = fetch_test_cases()
    api_df = api_to_dataframe(test_cases)

    compare_data(excel_df, api_df)
    generate_stats(excel_df, api_df)

if __name__ == "__main__":
    main()
