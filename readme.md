**# Lead Enrichment Automation

This project automates the enrichment of lead data in an Excel file by leveraging multiple APIs (Hunter.io, Lead411, SerpAPI, and Google Gemini). It fills in missing email addresses and LinkedIn profile URLs for contacts, using fuzzy matching and LLM verification for accuracy.

## Features

- **Automatic Excel File Handling:** Moves the target Excel file (`KNM_LIST.xlsx`) from your Downloads folder to the working directory if needed.
- **Flexible API Usage:** Supports enrichment via Hunter.io, Lead411, and SerpAPI, using whichever credentials are available.
- **Fuzzy Matching:** Uses fuzzy string matching to improve the reliability of LinkedIn profile matches.
- **LLM Verification:** Optionally verifies borderline LinkedIn matches using Google Gemini.
- **Safe Data Updates:** Only updates missing fields and saves changes to the Excel file.
- **Detailed Logging:** Prints a summary of all updates and actions taken.

## How It Works

1. **Initialization:**  
   Loads environment variables and checks for the Excel file.
2. **Data Loading:**  
   Reads the Excel file into a pandas DataFrame.
3. **Row-by-Row Enrichment:**  
   For each contact:
   - If email is missing, tries Hunter.io and/or Lead411.
   - If LinkedIn is missing, tries Lead411 and/or SerpAPI (with Gemini verification for borderline cases).
   - Updates the DataFrame if new data is found.
4. **Saving Results:**  
   Writes the updated DataFrame back to the Excel file and prints a summary.

## Flowchart

```mermaid
flowchart TD
    A[Start Script] --> B{Check for Excel file in Downloads}
    B -- Found and moved --> C[Move file to current directory]
    B -- Already exists --> D[File already in current directory]
    B -- Not found --> E[Exit: File not found]
    C --> F[Initialize PersonDetailsScraperPandas]
    D --> F
    F --> G[Load API credentials from environment]
    G --> H{Any API credentials found?}
    H -- None --> I[Exit: No API credentials]
    H -- Found --> J[Initialize APIs: Lead411, Hunter.io, SerpAPI]
    J --> K[Run enrich_leads]
    K --> L{For each row in Excel}
    L -- Missing data --> M[Try Hunter.io for email]
    M --> N[Try Lead411 for email/LinkedIn]
    N --> O[Try SerpAPI for LinkedIn]
    O --> P[Update DataFrame if found]
    P --> Q[Delay to avoid rate limiting]
    Q --> L
    L -- All rows processed --> R[Save updated Excel file]
    R --> S[Print summary]
    S --> T[End Script]

    click A call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L1")
    click B call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L40")
    click C call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L41")
    click D call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L43")
    click E call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L45")
    click F call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L533")
    click K call linkCallback("/Users/jerichowenzel/Desktop/python_assessment_va/main.py#L562")
```

## Sequence Diagram

```mermaid
sequenceDiagram
    participant Script as mainpy
    participant FS as FileSystem
    participant Pandas as ExcelReader
    participant DF as DataFrame
    participant Hunter as Hunterio
    participant SerpAPI as SerpAPI
    participant Gemini as GeminiFlash

    Note over Script,FS: Initialization and setup
    Script->>FS: load .env (dotenv.load_dotenv)
    FS-->>Script: env vars loaded
    Script->>FS: move KNM_LIST.xlsx from Downloads to CWD
    FS-->>Script: file moved or exists

    Note over Script,Pandas: Load Excel data
    Script->>Pandas: read KNM_LIST.xlsx
    Pandas-->>Script: DataFrame loaded

    loop For each row in DataFrame
        Script->>DF: check if Email or LinkedIn missing
        alt Email missing
            alt Hunterio key provided
                Note right of Script: attempt email enrichment
                Script->>Hunter: GET /email-finder
                Hunter-->>Script: {email, score}
            else No Hunter key
                Note right of Script: skip Hunterio
            end
        else Email present
            Note right of Script: skip email enrichment
        end

        alt LinkedIn missing
            alt SerpAPI key provided
                Note right of Script: attempt LinkedIn enrichment
                Script->>SerpAPI: search "Name Company site:linkedin.com/in"
                SerpAPI-->>Script: organic_results
                loop For each result
                    Script->>Script: compute fuzzy scores
                    alt Combined score borderline
                        Script->>Gemini: verify(profile_info, score)
                        Gemini-->>Script: Yes or No
                        alt Yes
                            Note right of Script: accept URL
                        else No
                            Note right of Script: discard URL
                        end
                    else Strong match
                        Note right of Script: accept URL
                    end
                end
                Note right of Script: select best LinkedIn URL
            else No SerpAPI key
                Note right of Script: skip SerpAPI
            end
        else LinkedIn present
            Note right of Script: skip LinkedIn enrichment
        end

        Script->>DF: update DataFrame
        DF-->>Script: row updated?
        alt Updated
            Note right of Script: increment counters
        else Not updated
            Note right of Script: log no-data-found
        end

        Note right of Script: sleep 1s
    end

    Note over Script,FS: Finalize and save
    Script->>FS: save DataFrame to XLSX
    FS-->>Script: file written
    Note right of Script: print summary of updates
```

## Requirements

- Python 3.8+
- See [requirements.txt](requirements.txt) for dependencies.

## Setup

1. **Clone the repository and install dependencies:**
    ```sh
    pip install -r requirements.txt
    ```

2. **Prepare your `.env` file** with the following keys (see example below):
    ```
    HUNTER_API_KEY=your_hunter_api_key
    SERPAPI_API_KEY=your_serpapi_api_key
    GOOGLE_API_KEY=your_google_gemini_api_key
    LEAD411_EMAIL=your_lead411_email
    LEAD411_PASSWORD=your_lead411_password
    ```

3. **Place your `KNM_LIST.xlsx` file** in your Downloads folder (the script will move it automatically).

## Usage

Run the script:

```sh
python main.py
```

The script will:
- Move `KNM_LIST.xlsx` to the current directory if needed.
- Load API credentials from `.env`.
- Enrich missing emails and LinkedIn URLs.
- Save the updated Excel file.
- Print a summary of updates.

## File Structure

- `main.py`: Main script with all enrichment logic.
- `requirements.txt`: Python dependencies.
- `readme.md`: This documentation file with embedded diagrams.
- `KNM_LIST.xlsx`: Your input Excel file (not included in repo).

## Visuals

### Flowchart

(See Flowchart section above)

### Sequence Diagram

(See Sequence Diagram section above)

---

**Note:**
- The script requires at least one API credential to run (Hunter.io, Lead411, or SerpAPI).
- For best results, provide as many API keys as possible in your `.env` file.
**