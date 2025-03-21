# Test Case Management Tool

## Description
This project is a Test Case Management Tool that reads test cases from an Excel file, fetches additional test cases from an API, compares the two datasets, and generates detailed statistics. It helps in managing and validating test cases efficiently.

## Installation
To set up the project, follow these steps:

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a .env file in the root directory and set the following environment variables:
   ```
   EXCEL_FILE=<path_to_your_excel_file>
   SHEET_NAME=<your_sheet_name>
   USERNAME=<your_api_username>
   ACCESS_KEY=<your_api_access_key>
   PROJECT_ID=<your_project_id>
   FOLDER_IDS=<comma_separated_folder_ids>
   ```

## Usage
To run the project, execute the following command:
```bash
python main.py
```
This will read the test cases from the specified Excel file, fetch additional test cases from the API, compare the data, and generate statistics.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License
This project is licensed under the MIT License.
