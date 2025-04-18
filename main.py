import os
import sys
from pathlib import Path
import shutil
import pandas as pd
from dotenv import load_dotenv
import requests
import re
import time
from typing import Optional, Dict, Any, List
from serpapi.google_search import GoogleSearch # Adjusted import for SerpAPI
from thefuzz import fuzz # Added for fuzzy matching
import google.generativeai as genai # Added for LLM verification

load_dotenv()

# Configure Gemini API
gemini_api_key = os.getenv("GOOGLE_API_KEY")
gemini_model = "gemini-2.5-flash-preview-04-17"
if not gemini_api_key:
    print("Warning: GOOGLE_API_KEY environment variable not set. Gemini verification will be skipped.")
else:
    try:
        genai.configure(api_key=gemini_api_key)
        gemini_model = genai.GenerativeModel('gemini-2.5-flash-preview-04-17')
        print("Gemini Flash model configured successfully.")
    except Exception as e:
        print(f"Error configuring Gemini: {e}. Gemini verification will be skipped.")

# Constants for matching thresholds
MIN_NAME_SCORE = 85
MIN_COMPANY_SCORE = 70
MIN_COMBINED_SCORE = 80
MIN_DOMAIN_SCORE = 60
HIGH_NAME_SCORE = 95
LOWER_COMPANY_SCORE = 60
GEMINI_THRESHOLD_LOW = 80
GEMINI_THRESHOLD_HIGH = 95

# the xlsx file to fill
XLSX_FILE = 'KNM_LIST.xlsx'

source_path = Path.home() / 'Downloads' / XLSX_FILE
destination_path = Path.cwd() / XLSX_FILE

# Check if the source file exists and the destination doesn't already exist
if source_path.exists() and not destination_path.exists():
    # Move the file
    shutil.move(source_path, destination_path)
    print(f"Moved '{XLSX_FILE}' from Downloads to the current directory.")
elif destination_path.exists():
    print(f"'{XLSX_FILE}' already exists in the current directory.")
else:
    print(f"'{XLSX_FILE}' not found in Downloads directory.")
    sys.exit(1)

class Lead411Enricher:
    """Handles interaction with the Lead411 API for contact enrichment."""
    BASE_URL = "https://api.lead411.com/v1"

    def __init__(self, email: str, password: str):
        self.email: str = email
        self.password: str = password
        self._token: Optional[str] = None

    def _get_token(self) -> Optional[str]:
        """
        Authenticates with Lead411 API to get/refresh the access token.
        Returns the token string or None if authentication fails.
        """
        # Only check if token exists
        if self._token:
            return self._token

        print("   Authenticating with Lead411...")
        auth_url = f"{self.BASE_URL}/authenticate_user"
        payload = {"email": self.email, "password": self.password}
        headers = {"Accept": "application/json", "Content-Type": "application/json"}
        try:
            response = requests.post(auth_url, json=payload, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()
            
            if data.get("success") and data.get("token"):
                self._token = data["token"]
                print("   Lead411 Authentication successful.")
                return self._token
            else:
                error_message = data.get('message', 'Unknown error')
                print(f"   Lead411 Authentication failed: {error_message}")
                self._token = None
                return None
        except requests.exceptions.Timeout:
            print("   Lead411 Authentication request timed out.")
            self._token = None
            return None
        except requests.exceptions.RequestException as e:
            error_details = ""
            try:
                error_details = response.text
            except:
                pass
            print(f"   Lead411 Authentication request error: {e} - Response: {error_details}")
            self._token = None
            return None

    def _search_contact(self, first_name: str, last_name: str, 
                       company_name: Optional[str], company_website: Optional[str]) -> Optional[int]:
        """
        Searches for a contact using Lead411 API.
        Returns the employee_id if an exact match is found, otherwise None.
        """
        token = self._get_token()
        if not token:
            print("   Skipping search: Authentication token not available.")
            return None

        search_url = f"{self.BASE_URL}/search/searchUsingJSON"
        # Token goes in params, not headers for this endpoint
        headers = {"Content-Type": "application/json"} 
        params = {"token": token}
        
        # Use company website or name as the search string
        search_query = company_website or company_name
        if not search_query:
            print("   Skipping search: Missing company website or name.")
            return None

        search_criteria = {
            "first_name": first_name,
            "last_name": last_name,
            "user_search_string": search_query, 
            "page_number": 1,
            "records_per_page": 5  # Limit results to find the best match quickly
        }
        
        print(f"   Searching Lead411 for {first_name} {last_name} at {search_query}...")
        try:
            # Pass token in params, search criteria in json body
            response = requests.post(search_url, params=params, json=search_criteria, headers=headers, timeout=20)
            response.raise_for_status()
            data = response.json()

            if data.get("success") and data.get("employees"):
                # Find the best match (simple exact name match here)
                for employee in data["employees"]:
                    # Case-insensitive comparison for robustness
                    emp_fname = employee.get("first_name", "").strip().lower()
                    emp_lname = employee.get("last_name", "").strip().lower()
                    target_fname = first_name.strip().lower()
                    target_lname = last_name.strip().lower()

                    if emp_fname == target_fname and emp_lname == target_lname:
                        employee_id = employee.get('employee_id')
                        if employee_id:
                            print(f"   Found potential match with ID: {employee_id}")
                            return int(employee_id)
                        else:
                            print("   Found matching employee record but missing employee_id.")
                
                print("   No exact name match found in search results.")
                return None
            else:
                error_message = data.get('message', 'Search endpoint returned failure or no employees')
                print(f"   Lead411 search failed or returned no employees: {error_message}")
                return None
        except requests.exceptions.Timeout:
            print("   Lead411 Search request timed out.")
            return None
        except requests.exceptions.RequestException as e:
            print(f"   Lead411 Search request error: {e}")
            return None

    def _unlock_contact(self, employee_id: int) -> Optional[Dict[str, Any]]:
        """
        Unlocks contact details (email, linkedin) for a given employee ID.
        Returns a dictionary with details or None if unlock fails.
        """
        token = self._get_token()
        if not token or not employee_id:
            print("   Skipping unlock: Authentication token not available.")
            return None

        unlock_url = f"{self.BASE_URL}/employee/unlock_employee_record"
        headers = {"Content-Type": "application/json", "token": token}
        payload = {"employee_id": employee_id}

        print(f"   Unlocking Lead411 contact details for ID: {employee_id}...")
        try:
            response = requests.post(unlock_url, json=payload, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()

            if data.get("success") and data.get("employee"):
                print("   Unlock successful.")
                employee_data = data["employee"]
                unlocked_details = {
                    "email": employee_data.get("email"),
                    "linkedin_url": employee_data.get("linkedin_url")
                }
                return unlocked_details
            else:
                error_message = data.get('message', 'Unlock endpoint returned failure')
                print(f"   Lead411 Unlock failed: {error_message}")
                return None
        except requests.exceptions.Timeout:
            print("   Lead411 Unlock request timed out.")
            return None
        except requests.exceptions.RequestException as e:
            print(f"   Lead411 Unlock request error: {e}")
            return None

    def find_contact_details(self, first_name: str, last_name: str, 
                           company_name: Optional[str], company_website: Optional[str]) -> Optional[Dict[str, Any]]:
        """
        Orchestrates search and unlock to find contact details.
        Returns a dictionary {'email': ..., 'linkedin_url': ...} or None.
        """
        employee_id = self._search_contact(first_name, last_name, company_name, company_website)
        if employee_id:
            # Short delay before unlocking might be polite to the API
            time.sleep(0.5) 
            return self._unlock_contact(employee_id)
        return None




class PersonDetailsScraperPandas:
    """Reads an Excel file, orchestrates lead enrichment, and saves changes."""

    def __init__(self, filename: str = XLSX_FILE):
        self.filepath: Path = Path.cwd() / filename
        self.df: Optional[pd.DataFrame] = None
        self.rows_as_dicts: List[Dict] = []

        if not self.filepath.exists():
            raise FileNotFoundError(f"Input Excel file not found at {self.filepath}")
        
        self._load_data()

    def _load_data(self) -> None:
        """Loads data from the Excel file into a pandas DataFrame."""
        try:
            # header=0 means the FIRST row (index 0) is the header
            self.df = pd.read_excel(self.filepath, header=0) 
            # Print the actual column names found by pandas
            print(f"DataFrame columns found: {self.df.columns.tolist()}")
            
            # Check if essential columns exist before proceeding
            required_columns = ['Contact Name', 'Company Name', 'Company Website', 'Email', 'Linkedin Profile Link']
            missing_cols = [col for col in required_columns if col not in self.df.columns]
            if missing_cols:
                print(f"Warning: The following required columns were not found in the Excel file: {missing_cols}")
                print("Please ensure the first row of your Excel file contains the correct headers.")
                # Decide if you want to raise an error or try to continue
                # raise ValueError(f"Missing required columns: {missing_cols}")
            
            self.rows_as_dicts = self.df.to_dict(orient='records')
            print(f"Successfully loaded {len(self.rows_as_dicts)} rows into DataFrame.")
            
            # Optional: Print first few rows for verification
            # print("Sample rows (first 5):")
            # for row in self.rows_as_dicts[:5]:
            #     print(row)
        except Exception as e:
            # Provide more context in the error message
            raise IOError(f"Failed to read or process Excel file {self.filepath}: {e}")

    def get_rows(self) -> List[Dict]:
        """Returns the loaded data as a list of dictionaries."""
        return self.rows_as_dicts

    # Update enrich_leads signature and logic
    def enrich_leads(self,
                   lead411_enricher: Optional[Lead411Enricher] = None,
                   hunter_api_key: Optional[str] = None,
                   serpapi_api_key: Optional[str] = None) -> None: # Added serpapi_api_key
        """
        Enriches lead data using various APIs based on what's provided.

        Args:
            lead411_enricher: Optional Lead411Enricher instance for using Lead411 API
            hunter_api_key: Optional Hunter.io API key for using Hunter.io API
            serpapi_api_key: Optional SerpAPI API key for Google Search
        """
        if self.df is None:
            print("Error: DataFrame not loaded. Cannot enrich leads.")
            return

        if not lead411_enricher and not hunter_api_key and not serpapi_api_key: # Added serpapi_api_key check
            print("Error: No enrichment service provided. Need Lead411, Hunter API key, or SerpAPI key.")
            return

        print("\\nStarting enrichment process...")
        if lead411_enricher:
            print("- Using Lead411 API")
        if hunter_api_key:
            print("- Using Hunter.io API")
        if serpapi_api_key: # Added SerpAPI check
            print("- Using SerpAPI for LinkedIn search")

        # Tracking enrichment statistics
        updates_made = 0
        email_updates = 0
        linkedin_updates = 0
        enriched_rows = []

        for idx, row_data in self.df.iterrows():
            # Use .get with default to handle potentially missing columns gracefully
            email_missing = pd.isna(row_data.get('Email'))
            linkedin_missing = pd.isna(row_data.get('Linkedin Profile Link'))

            if not (email_missing or linkedin_missing):
                continue  # Skip if nothing is missing

            contact_name = row_data.get('Contact Name')
            company_name = row_data.get('Company Name')
            company_website = row_data.get('Company Website')

            # Basic validation
            if pd.isna(contact_name) or (pd.isna(company_name) and pd.isna(company_website)):
                print(f"\\nSkipping row index {idx}: Missing Contact Name or both Company Name/Website.")
                continue

            # Split name for API calls
            name_parts = str(contact_name).split()  # Ensure it's a string
            first_name = name_parts[0] if name_parts else None
            last_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else None

            if not first_name or not last_name:
                print(f"\\nSkipping row index {idx}: Could not extract First/Last name from '{contact_name}'.")
                continue

            print(f"\\nEnriching row index {idx}: {contact_name} at {company_name or company_website}")

            found_email = None
            found_linkedin = None

            # --- Try Hunter.io first if API key is provided ---
            if email_missing and hunter_api_key:
                print(f"   Trying Hunter.io API for email...")
                domain = None
                if pd.notna(company_website):
                    # Improved regex to handle various URL formats
                    match = re.search(r'^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/?#\n]+)', str(company_website))
                    if match:
                        domain = match.group(1)
                if first_name and last_name and (domain or company_name):
                    try:
                        params = {'api_key': hunter_api_key, 'first_name': first_name, 'last_name': last_name}
                        if domain: params['domain'] = domain
                        elif company_name: params['company'] = str(company_name) # Ensure company name is string
                        print(f"   Querying Hunter.io Email Finder...")
                        response = requests.get("https://api.hunter.io/v2/email-finder", params=params, timeout=15)
                        response.raise_for_status()
                        data = response.json()
                        if data.get('data') and data['data'].get('email'):
                            found_email = data['data']['email']
                            score = data['data'].get('score', 0)
                            print(f"   Found email via Hunter.io with confidence score: {score}/100")
                        else:
                            print(f"   Hunter.io found no email for {contact_name}.")
                            if data.get('errors'): print(f"   Hunter.io API Errors: {data['errors']}")
                    except requests.exceptions.RequestException as e: print(f"   Hunter.io API Request Error: {e}")
                    except Exception as e: print(f"   Error during Hunter.io enrichment: {e}")
                else: print("   Skipping Hunter.io: Insufficient data (need domain or company name).")


            # --- Try Lead411 if needed ---
            # Check if email is still missing OR if linkedin is missing
            needs_lead411 = (email_missing and not found_email) or linkedin_missing
            if lead411_enricher and needs_lead411:
                print(f"   Trying Lead411 API...")
                try:
                    details = lead411_enricher.find_contact_details(
                        first_name, last_name,
                        str(company_name) if pd.notna(company_name) else None, # Ensure string
                        str(company_website) if pd.notna(company_website) else None # Ensure string
                    )
                    if details:
                        # Only update email if it's still missing
                        if email_missing and not found_email and details.get("email"):
                            found_email = details.get("email")
                            print("   Found Email via Lead411.")
                        # Update LinkedIn if it's missing AND Lead411 provided it
                        if linkedin_missing:
                            found_linkedin = details.get("linkedin_url")
                            if found_linkedin:
                                print(f"   Lead411 found LinkedIn: {found_linkedin}")
                            else:
                                print("   Lead411 did not provide a LinkedIn URL.")
                        # Simplified message
                        print(f"   Lead411 results processed.")
                    else:
                        print(f"   Lead411 could not find details for {contact_name}.")
                except Exception as e:
                    print(f"   Error during Lead411 enrichment: {e}")


            # --- Try SerpAPI for LinkedIn if still missing ---
            # Check if LinkedIn is missing AND we haven't found it yet AND the API key is provided
            if linkedin_missing and not found_linkedin and serpapi_api_key:
                print(f"   Trying SerpAPI for LinkedIn...") # Removed key check message from here
                # No need for an inner 'else' block if the outer 'if' handles the key check

                try:
                    company_identifier = str(company_name or company_website or '')
                    if company_identifier:
                        search_query = f'"{first_name} {last_name}" "{company_identifier}" site:linkedin.com/in'
                        params = {
                            "engine": "google",
                            "q": search_query,
                            "api_key": serpapi_api_key,
                            "num": 10,
                            "start": 0
                        }
                        print(f"   Querying SerpAPI with: '{search_query}'")
                        search = GoogleSearch(params)
                        results = search.get_dict() # Execute the search and get results

                        # Check if 'organic_results' key exists and is not empty
                        organic_results = results.get("organic_results")
                        if not organic_results:
                            print("   SerpAPI returned no organic results.")
                            # Optional: Log the full response if needed for debugging
                            # print(f"   SerpAPI Full Response (for debugging): {results}")
                        else:
                            print(f"   SerpAPI Raw Results Count: {len(organic_results)}")

                            best_match_url = None
                            highest_score = -1

                            target_name_lower = f"{first_name} {last_name}".lower()
                            target_company_lower = company_identifier.lower()

                            for i, result in enumerate(organic_results): # Iterate directly over the list
                                link = result.get("link", "")
                                title = result.get("title", "").lower()
                                snippet = result.get("snippet", "").lower()
                                print(f"      [{i+1}] Checking: Title='{result.get('title', '')}' Link='{link}'")

                                # ... existing URL filtering ...
                                if "linkedin.com/in/" not in link or any(x in link for x in ["/company/", "/pub/", "/jobs/", "/sales/people/", "/posts/"]):
                                    print(f"         Skipping link: Not a valid profile URL or excluded path.")
                                    continue

                                # --- Fuzzy Matching ---
                                name_score_title = fuzz.partial_ratio(target_name_lower, title)
                                name_score_snippet = fuzz.partial_ratio(target_name_lower, snippet)
                                name_score = max(name_score_title, name_score_snippet)

                                # Calculate company_score using token_set_ratio on company name
                                company_score = max(
                                    fuzz.token_set_ratio(target_company_lower, title),
                                    fuzz.token_set_ratio(target_company_lower, snippet)
                                )

                                # Extract domain identifier (if available) and calculate domain_score
                                domain_identifier = ''
                                if pd.notna(company_website):
                                    dm = re.search(r'^(?:https?://)?(?:www\.)?([^:/?#\n]+)', company_website)
                                    if dm:
                                        domain_identifier = dm.group(1).lower()
                                domain_score = 0
                                if domain_identifier:
                                    domain_score = max(
                                        fuzz.partial_ratio(domain_identifier, link.lower()),
                                        fuzz.partial_ratio(domain_identifier, title),
                                        fuzz.partial_ratio(domain_identifier, snippet)
                                    )

                                # Combined confidence score with domain weight
                                combined_score = (name_score * 0.5) + (company_score * 0.3) + (domain_score * 0.2)
                                print(f"         Scores: Name={name_score}, Company={company_score}, Domain={domain_score}, Combined={combined_score:.2f}")

                                # Selection thresholds include domain check
                                # Define high-name and lower-company thresholds
                                # Dynamic selection logic
                                if (name_score >= HIGH_NAME_SCORE and combined_score >= MIN_COMBINED_SCORE):
                                    # High confidence on name alone
                                    highest_score = combined_score
                                    best_match_url = link
                                    print(f"         >>> High-name fallback selected! Score: {combined_score:.2f}, URL: {link}")
                                elif (name_score >= MIN_NAME_SCORE and company_score >= MIN_COMPANY_SCORE and
                                      (not domain_identifier or domain_score >= MIN_DOMAIN_SCORE) and
                                      combined_score >= MIN_COMBINED_SCORE):
                                    if (combined_score > highest_score):
                                        highest_score = combined_score
                                        best_match_url = link
                                        print(f"         >>> New best match found! Score: {highest_score:.2f}, URL: {link}")
                                elif (domain_identifier and domain_score >= 80 and
                                      name_score >= MIN_NAME_SCORE and combined_score >= MIN_COMBINED_SCORE):
                                    # Domain-heavy fallback
                                    highest_score = combined_score
                                    best_match_url = link
                                    print(f"         >>> Domain-fallback selected! Score: {combined_score:.2f}, URL: {link}")
                                else:
                                    print(f"         Skipping link: did not meet dynamic thresholds (Name: {name_score}, Company: {company_score}, Domain: {domain_score}, Combined: {combined_score:.2f}).")


                            # After checking all results, we may want to verify borderline matches with Gemini
                            if best_match_url and GEMINI_THRESHOLD_LOW <= highest_score <= GEMINI_THRESHOLD_HIGH:
                                # This is a borderline match - use Gemini to verify
                                print(f"   Borderline confidence score ({highest_score:.2f}): Using Gemini for verification...")
                                
                                # Extract the best match's title/snippet for context
                                best_match_result = next((r for r in organic_results if r.get("link") == best_match_url), None)
                                if best_match_result:
                                    best_title = best_match_result.get("title", "")
                                    best_snippet = best_match_result.get("snippet", "")
                                    linkedin_context = f"{best_title} - {best_snippet}"
                                    
                                    # Use Gemini to verify this match
                                    is_verified = verify_with_gemini(
                                        f"{first_name} {last_name}",
                                        company_identifier,
                                        linkedin_context,
                                        highest_score
                                    )
                                    
                                    if is_verified:
                                        print(f"   >>> Gemini verified the LinkedIn match: {best_match_url}")
                                    else:
                                        print(f"   >>> Gemini rejected the match. Discarding LinkedIn URL.")
                                        best_match_url = None  # Discard the match if Gemini says no
                            
                            # Update found_linkedin based on final decision
                            if best_match_url:
                                found_linkedin = best_match_url
                                print(f"   >>> Final selected LinkedIn URL: {found_linkedin} (Score: {highest_score:.2f})")
                            else:
                                print("   No suitable LinkedIn URL found that meets confidence thresholds.")

                    else: # if not company_identifier
                        print("   Skipping SerpAPI: Missing company name/website for query.")

                except requests.exceptions.RequestException as req_err:
                    # Catch potential network errors from the underlying request
                    print(f"   Error during SerpAPI request: {req_err}")
                except Exception as e:
                    # Catch other potential errors during SerpAPI interaction or result processing
                    print(f"   An unexpected error occurred during SerpAPI processing: {e}")
                    import traceback
                    traceback.print_exc() # Keep traceback for unexpected errors

            # --- Update DataFrame ---
            row_updated = False
            email_updated = False
            linkedin_updated = False

            # Update Email if it was missing and we found one
            if email_missing and pd.notna(found_email):
                self.df.loc[idx, 'Email'] = found_email
                print(f"   Updated Email: {found_email}")
                row_updated = True
                email_updated = True
                email_updates += 1

            # Update LinkedIn if it was missing and we found one (from Lead411 or SerpAPI)
            if linkedin_missing and pd.notna(found_linkedin):
                self.df.loc[idx, 'Linkedin Profile Link'] = found_linkedin
                print(f"   Updated LinkedIn: {found_linkedin}")
                row_updated = True
                linkedin_updated = True
                linkedin_updates += 1

            if row_updated:
                updates_made += 1
                # Track which row was updated and what was updated
                enriched_rows.append({
                    'index': idx,
                    'contact_name': contact_name,
                    'company': company_name or company_website,
                    'email_updated': email_updated,
                    'found_email': found_email if email_updated else None,
                    'linkedin_updated': linkedin_updated,
                    'found_linkedin': found_linkedin if linkedin_updated else None
                })
            # Only print failure message if we actually tried to find something and failed
            elif (email_missing or linkedin_missing):
                tried_email = email_missing and (hunter_api_key or lead411_enricher)
                tried_linkedin = linkedin_missing and (lead411_enricher or serpapi_api_key)
                if tried_email or tried_linkedin:
                    print(f"   Could not find missing data for {contact_name} using available services.")

            # Add a delay to avoid rate limiting (consider if SerpAPI needs its own delay)
            time.sleep(1) # Keep the general delay

        # Print summary of updates
        print(f"\\nEnrichment complete. Updated {updates_made} rows total:")
        print(f"- Email addresses filled: {email_updates}")
        print(f"- LinkedIn URLs filled: {linkedin_updates}")

        if enriched_rows:
            print("\\nDetails of enriched rows:")
            for row in enriched_rows:
                details = []
                if row['email_updated']:
                    details.append(f"Email: {row['found_email']}")
                if row['linkedin_updated']:
                    details.append(f"LinkedIn: {row['found_linkedin']}")
                if details: # Only print if something was updated
                    print(f"Row {row['index']} - {row['contact_name']} at {row['company']}: {', '.join(details)}")

        if updates_made > 0:
            self.save_changes()

    def save_changes(self, output_path: Optional[str] = None) -> None:
        """Saves the current state of the DataFrame back to an Excel file."""
        if self.df is None:
            print("Error: No DataFrame loaded to save.")
            return

        save_path = Path(output_path) if output_path else self.filepath
        print(f"\\nSaving updated DataFrame to '{save_path}'...")
        try:
            # Ensure the directory exists if saving to a new path
            save_path.parent.mkdir(parents=True, exist_ok=True)
            self.df.to_excel(save_path, index=False)
            print(f"DataFrame successfully saved.")
        except Exception as e:
            print(f"Error saving DataFrame to {save_path}: {e}")

def verify_with_gemini(person_name, company_name, linkedin_profile_title, confidence_score):
    """
    Uses Gemini Flash to verify if a LinkedIn profile likely belongs to the target person at the target company.
    
    Args:
        person_name: The name of the person we're searching for
        company_name: The name of the company the person works at
        linkedin_profile_title: The title or snippet from the LinkedIn profile
        confidence_score: The fuzzy matching confidence score (0-100)
        
    Returns:
        Boolean: True if verified as likely match, False otherwise
    """
    if not gemini_model:
        print("   Gemini verification skipped: model not available")
        return False  # Skip if Gemini is not configured
    
    prompt = f"""
    Analyze whether this LinkedIn profile likely belongs to the target person:
    
    Target Person: "{person_name}"
    Target Company: "{company_name}"
    LinkedIn Profile Info: "{linkedin_profile_title}"
    Fuzzy Match Score: {confidence_score}/100
    
    Based on the information, does this LinkedIn profile likely belong to the target person at the target company?
    Consider name variations, company name variations (abbreviations, omitted legal suffixes like Inc, Ltd, LLC),
    and job titles/roles. 
    
    Answer with only "Yes" or "No".
    """
    
    try:
        response = gemini_model.generate_content(prompt)
        # Parse the response (expecting a simple "Yes" or "No")
        answer = response.text.strip().lower()
        is_verified = "yes" in answer.lower()
        print(f"   Gemini verification for '{person_name}' at '{company_name}': {answer} -> {'Verified ✓' if is_verified else 'Not Verified ✗'}")
        return is_verified
    except Exception as e:
        print(f"   Error during Gemini verification: {e}")
        return False  # Treat errors as non-verification


if __name__ == "__main__":
    print("Script started.")
    try:
        # Initialize the scraper (loads data)
        scraper = PersonDetailsScraperPandas(filename=XLSX_FILE)

        # Get API credentials from environment
        lead411_email = os.getenv("LEAD411_EMAIL")
        lead411_password = os.getenv("LEAD411_PASSWORD")
        hunter_key = os.getenv("HUNTER_API_KEY")
        serpapi_key = os.getenv("SERPAPI_API_KEY") # Standardize env variable name

        # Initialize services based on available credentials
        lead411_enricher = None
        if lead411_email and lead411_password:
            print("Lead411 credentials found.")
            lead411_enricher = Lead411Enricher(lead411_email, lead411_password)
        else:
            print("Lead411 credentials not found. Skipping Lead411 API.")

        if not hunter_key:
            print("Hunter.io API key not found. Skipping Hunter.io API.")
        else:
            print("Hunter.io API key found.")

        if not serpapi_key: # Check the correct variable
            print("SerpAPI API key (SERPAPI_API_KEY) not found in .env. Skipping SerpAPI LinkedIn search.")
        else:
            print("SerpAPI API key found.")

        if not lead411_enricher and not hunter_key and not serpapi_key: # Check correct variable
            print("Error: No API credentials found. Please set at least one of:")
            print("- LEAD411_EMAIL and LEAD411_PASSWORD for Lead411 API")
            print("- HUNTER_API_KEY for Hunter.io API")
            print("- SERPAPI_API_KEY for SerpAPI") # Use standardized env var name in message
            sys.exit(1)

        # Run the enrichment process with available services
        scraper.enrich_leads(
            lead411_enricher=lead411_enricher,
            hunter_api_key=hunter_key,
            serpapi_api_key=serpapi_key # Pass the correct variable name
        )

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except IOError as e:
        print(f"Error reading/processing file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        # Print traceback for detailed debugging
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print("Script finished.")