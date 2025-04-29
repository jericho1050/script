
# Lead Enrichment Automation

A Python utility to automate enrichment of lead data (email, LinkedIn, website, address) in an Excel spreadsheet using the Lead411 API.

## Features

- Moves `KNM_LIST.xlsx` from `~/Downloads` to the working directory  
- Loads spreadsheet into a pandas DataFrame  
- Authenticates with Lead411 and searches companies & employees  
- Unlocks contact details (email, LinkedIn URL, company website, address)  
- Updates missing fields and saves back to `KNM_LIST.xlsx`  
- Prints a summary of all changes  

## Prerequisites

- Python 3.8 or higher  
- A Lead411 account with API credentials  

## Installation

1. Clone this repository  
   ```bash
   git clone https://github.com/your-org/python_assessment_va.git
   cd python_assessment_va
   ```
2. Create and activate a virtual environment  
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
3. Install dependencies  
   ```bash
   pip install pandas python-dotenv requests openpyxl
   ```

## Configuration

Create a `.env` file in the project root with your Lead411 credentials:

```dotenv
LEAD411_EMAIL=your_email@example.com
LEAD411_PASSWORD=your_password
```

## Usage

1. Place `KNM_LIST.xlsx` in your `Downloads` folder.  
2. Run the script:

    ```bash
    python main.py
    ```

3. The script will move the file, enrich missing fields, and overwrite `KNM_LIST.xlsx` in the project directory.  

## Environment Variables

- `LEAD411_EMAIL` – Lead411 account email  
- `LEAD411_PASSWORD` – Lead411 account password  

## Project Structure

```text
.
├── .env                   # API credentials
├── KNM_LIST.xlsx          # Input/output Excel file
├── main.py                # Enrichment script
├── diagram.md             # Flowchart (Mermaid)
├── diagram2.md            # Sequence diagram (Mermaid)
└── readme.md              # This file
```

## License

This project is released under the MIT License.  
