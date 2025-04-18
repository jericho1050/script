import os
import sys
from pathlib import Path
import shutil
import pandas as pd
from dotenv import load_dotenv
import requests
import re
import time
import json
from typing import Optional, Dict, Any, List

load_dotenv()

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
        self._authenticate() # Authenticate on initialization

    def _authenticate(self) -> None:
        """Authenticates with Lead411 API to get the access token."""
        print("   Authenticating with Lead411...")
        auth_url = f"{self.BASE_URL}/authenticate_user"
        # Lead411 auth uses query parameters, not JSON body
        params = {"email": self.email, "password": self.password}
        headers = {"Accept": "application/json"} # Keep Accept header
        try:
            # Use POST method as per documentation
            response = requests.post(auth_url, params=params, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()

            if data.get("status") == "success" and data.get("token"):
                self._token = data["token"]
                print("   Lead411 Authentication successful.")
            else:
                error_message = data.get('message', 'Unknown authentication error')
                print(f"   Lead411 Authentication failed: {error_message}")
                self._token = None
                # Optionally raise an error if authentication is critical
                # raise ConnectionError(f"Lead411 Authentication failed: {error_message}")
        except requests.exceptions.Timeout:
            print("   Lead411 Authentication request timed out.")
            self._token = None
            # raise ConnectionError("Lead411 Authentication request timed out.")
        except requests.exceptions.RequestException as e:
            error_details = ""
            try:
                # Try to get more details from the response if available
                error_details = f" Status Code: {response.status_code}, Response: {response.text}"
            except:
                pass
            print(f"   Lead411 Authentication request error: {e}{error_details}")
            self._token = None
            # raise ConnectionError(f"Lead411 Authentication request error: {e}{error_details}")

    def _get_token(self) -> Optional[str]:
        """Returns the current token, re-authenticating if necessary (though currently auth is in __init__)."""
        # Simple getter for now, assumes __init__ handled auth
        # Could add token expiry check and re-auth logic here if needed later
        if not self._token:
            print("   Warning: Lead411 token not available. Authentication might have failed.")
        return self._token

    def _search_contact(self, first_name: str, last_name: str, title: Optional[str],
                       company_name: Optional[str], company_website: Optional[str],
                       state: Optional[str] = None) -> Optional[int]:
        """
        Searches for a contact using Lead411 API.
        Returns the employee_id if a suitable match is found, otherwise None.
        Uses the company-specific employee search endpoint.
        """
        token = self._get_token()
        if not token:
            print("   Skipping search: Authentication token not available.")
            return None

        # First, search for the company to get its ID
        company_id = self._get_company_id(company_name, company_website, token)
        if not company_id:
            print("   Could not find company ID.")
            return None

        # Now search for employees within that company
        search_url = f"{self.BASE_URL}/company/searchCompanyEmployees"
        headers = {"Accept": "application/json"}
        
        # Use both first and last name for better matching
        search_name = f"{first_name} {last_name}".strip()
        
        params = {
            "token": token,
            "company_id": company_id,
            "search_name": search_name
        }

        print(f"   Searching Lead411 for {search_name} at company ID {company_id}...")
        try:
            response = requests.post(search_url, params=params, headers=headers, timeout=20)
            response.raise_for_status()
            data = response.json()

            # Debug: Print the response structure
            print(f"   Response structure: {data.keys() if isinstance(data, dict) else 'Not a dict'}")

            # Check response structure
            if "companyEmployeesData" in data and "companyEmployees" in data["companyEmployeesData"]:
                employees = data["companyEmployeesData"]["companyEmployees"]
                if not employees:
                    print("   Lead411 search successful but returned no employee matches.")
                    return None

                # Find the best match
                target_fname_lower = first_name.strip().lower()
                target_lname_lower = last_name.strip().lower()

                for employee in employees:
                    emp_fname = employee.get("first_name", "").strip().lower()
                    emp_lname = employee.get("last_name", "").strip().lower()

                    # Exact match check
                    if emp_fname == target_fname_lower and emp_lname == target_lname_lower:
                        employee_id = employee.get('employee_id')
                        if employee_id:
                            print(f"   Found exact match with ID: {employee_id}")
                            return int(employee_id)

                print("   No exact name match found in search results.")
                return None

            else:
                error_message = data.get('message', 'Search endpoint returned failure or unexpected structure')
                print(f"   Lead411 search failed or returned unexpected data: {error_message}")
                print(f"   Response data:\n{json.dumps(data, indent=2)}")  # Pretty print the response
                return None

        except requests.exceptions.Timeout:
            print("   Lead411 Search request timed out.")
            return None
        except requests.exceptions.RequestException as e:
            error_details = ""
            try:
                error_details = f" Status Code: {response.status_code}, Response: {response.text}"
            except:
                pass
            print(f"   Lead411 Search request error: {e}{error_details}")
            return None
        except ValueError:
            print(f"   Lead411 Search response was not valid JSON. Response: {response.text}")
            return None

    def _get_company_id(self, company_name: Optional[str], company_website: Optional[str], token: str) -> Optional[int]:
        """
        Searches for a company to get its ID using Lead411 API.
        Returns the company ID if found, otherwise None.
        """
        search_url = f"{self.BASE_URL}/search/searchUsingJSON"
        headers = {"Accept": "application/json"}

        # Prepare parameters for the search
        params = {
            "token": token,
            "page": 1,
            "per_page": 1,  # We only need one result
            "companyResults": "companyResultsAll",
            "country_code": "US"  # Focus on US companies
        }

        # Prioritize company website for searching, fallback to company name
        search_query = None
        if company_website:
            # Extract domain if possible, otherwise use the full string
            match = re.search(r'^(?:https?:\/\/)?(?:www\.)?([^:\/?#\n]+)', company_website)
            if match:
                search_query = match.group(1)
            else:
                search_query = company_website
            params["user_search_string"] = search_query
        elif company_name:
            search_query = company_name
            params["user_search_string"] = search_query

        if not search_query:
            print("   Skipping company search: Missing company website or name.")
            return None

        print(f"   Searching Lead411 for company: {search_query}...")
        try:
            response = requests.post(search_url, params=params, headers=headers, timeout=20)
            response.raise_for_status()
            data = response.json()

            if "AllResults" in data and data["AllResults"]:
                company = data["AllResults"][0]
                company_id = company.get("company_id")
                if company_id:
                    print(f"   Found company ID: {company_id}")
                    return int(company_id)

            print("   No company found matching the search criteria.")
            return None

        except requests.exceptions.Timeout:
            print("   Lead411 Company Search request timed out.")
            return None
        except requests.exceptions.RequestException as e:
            error_details = ""
            try:
                error_details = f" Status Code: {response.status_code}, Response:\n{json.dumps(response.json(), indent=2)}"
            except:
                pass
            print(f"   Lead411 Company Search request error: {e}{error_details}")
            return None
        except ValueError:
            print(f"   Lead411 Company Search response was not valid JSON. Response: {response.text}")
            return None

    def _unlock_contact(self, employee_id: int) -> Optional[Dict[str, Any]]:
        """
        Unlocks contact details (email, linkedin) for a given employee ID.
        Returns a dictionary with details or None if unlock fails.
        Uses the 'Unlock Employee Data Email and Phone' endpoint.
        """
        token = self._get_token()
        if not token:
            print("   Skipping unlock: Authentication token not available.")
            return None
        if not employee_id:
             print("   Skipping unlock: Invalid employee_id provided.")
             return None

        unlock_url = f"{self.BASE_URL}/employee/unlock_employee_record"
        headers = {"Accept": "application/json"} # Only Accept header needed
        params = {
            "token": token,
            "employee_id": employee_id
        }

        print(f"   Unlocking Lead411 contact details for ID: {employee_id}...")
        try:
            # Use POST method as per documentation
            response = requests.post(unlock_url, params=params, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()

            # --- Debug: Print the full response immediately ---
            print(f"   Raw Unlock Response Data:\n{json.dumps(data, indent=2)}")
            # --- End Debug ---

            # Debug: Print the response structure keys
            print(f"   Unlock response structure (keys): {data.keys() if isinstance(data, dict) else 'Not a dict'}")

            # Check response structure based on the actual observed response
            if data.get("status") == "success":
                print("   Unlock status is 'success'. Extracting details...")
                # Extract email from top level
                found_email = data.get("email")
                
                # Extract LinkedIn URL from nested employee_details
                employee_details = data.get("employee_details") # Get the nested dictionary
                found_linkedin = None
                if isinstance(employee_details, dict):
                    found_linkedin = employee_details.get("linkedin")

                # Extract Company Website and Address from company_data
                company_data = data.get("company_data")
                found_company_website = None
                found_address1 = None
                found_address2 = None
                if isinstance(company_data, dict):
                    found_company_website = company_data.get("URL")
                    found_address1 = company_data.get("address1")
                    found_address2 = company_data.get("address2")
                    # Potentially add city, state, zip if needed later
                    # found_city = company_data.get("city")
                    # found_state = company_data.get("region_code")
                    # found_zip = company_data.get("zip")
                
                # Prepare the result dictionary
                unlocked_details = {}
                if found_email:
                    unlocked_details["email"] = found_email
                if found_linkedin:
                    unlocked_details["linkedin_url"] = found_linkedin # Use the key expected by enrich_leads
                if found_company_website:
                    unlocked_details["company_website"] = found_company_website
                # Combine address fields
                address_parts = [addr for addr in [found_address1, found_address2] if pd.notna(addr) and addr.strip()]
                if address_parts:
                    unlocked_details["address"] = ", ".join(address_parts)
                
                if unlocked_details:
                    print(f"   Found details: {json.dumps(unlocked_details, indent=2)}")
                    return unlocked_details
                else:
                    print("   Unlock successful, but no email, linkedin, website or address found in the response structure.")
                    # Already printed the full response above
                    return None
            else:
                error_message = data.get('message', 'Unlock endpoint returned failure status or unexpected structure')
                print(f"   Lead411 Unlock failed: {error_message}")
                # Already printed the full response above
                return None

        except requests.exceptions.Timeout:
            print("   Lead411 Unlock request timed out.")
            return None
        except requests.exceptions.RequestException as e:
            error_details = ""
            try:
                error_details = f" Status Code: {response.status_code}, Response:\n{json.dumps(response.json(), indent=2)}"
            except:
                pass
            print(f"   Lead411 Unlock request error: {e}{error_details}")
            return None
        except ValueError: # Catches JSONDecodeError
             print(f"   Lead411 Unlock response was not valid JSON. Response: {response.text}")
             return None

    def find_contact_details(self, first_name: str, last_name: str, title: Optional[str],
                           company_name: Optional[str], company_website: Optional[str],
                           state: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """
        Orchestrates search and unlock using Lead411 to find contact details.
        Returns a dictionary {'email': ..., 'linkedin_url': ...} or None.
        """
        employee_id = self._search_contact(first_name, last_name, title, company_name, company_website, state)
        if employee_id:
            # Short delay before unlocking might be polite to the API
            time.sleep(0.5)
            return self._unlock_contact(employee_id)
        else:
            print(f"   Could not find a suitable employee ID for {first_name} {last_name} via Lead411 search.")
            return None


class PersonDetailsScraperPandas:
    """Reads an Excel file, orchestrates lead enrichment using Lead411, and saves changes."""

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
            # Define desired data types for specific columns
            dtypes = {
                'Email': str,
                'Linkedin Profile Link': str
                # Add other columns if they also need specific types
            }
            # Load Excel, specifying dtypes. Use 'object' if 'str' causes issues.
            self.df = pd.read_excel(self.filepath, header=0, dtype=dtypes)
            print(f"DataFrame columns found: {self.df.columns.tolist()}")

            # # --- Ensure correct dtypes for string columns (Now handled by dtype in read_excel) ---
            # if 'Email' in self.df.columns:
            #     self.df['Email'] = self.df['Email'].astype(object)
            # if 'Linkedin Profile Link' in self.df.columns:
            #     self.df['Linkedin Profile Link'] = self.df['Linkedin Profile Link'].astype(object)
            # # --- End dtype correction ---

            required_columns = ['Contact Name', 'Company Name', 'Company Website', 'Email', 'Linkedin Profile Link']
            missing_cols = [col for col in required_columns if col not in self.df.columns]
            if missing_cols:
                print(f"Warning: The following required columns were not found in the Excel file: {missing_cols}")
                print("Please ensure the first row of your Excel file contains the correct headers.")
                # Consider raising ValueError if columns are absolutely essential

            self.rows_as_dicts = self.df.to_dict(orient='records')
            print(f"Successfully loaded {len(self.rows_as_dicts)} rows into DataFrame.")
        except Exception as e:
            raise IOError(f"Failed to read or process Excel file {self.filepath}: {e}")

    def get_rows(self) -> List[Dict]:
        """Returns the loaded data as a list of dictionaries."""
        return self.rows_as_dicts

    def enrich_leads(self, lead411_enricher: Lead411Enricher) -> None:
        """
        Enriches lead data using the provided Lead411Enricher instance.

        Args:
            lead411_enricher: Lead411Enricher instance for using Lead411 API
        """
        if self.df is None:
            print("Error: DataFrame not loaded. Cannot enrich leads.")
            return

        if not lead411_enricher:
            print("Error: Lead411Enricher instance not provided. Cannot enrich leads.")
            return

        print("\nStarting enrichment process using Lead411 API...")

        updates_made = 0
        email_updates = 0
        linkedin_updates = 0
        website_updates = 0 # New counter
        address_updates = 0 # New counter
        enriched_rows = []

        for idx, row_data in self.df.iterrows():
            # Check which fields are missing
            email_missing = pd.isna(row_data.get('Email'))
            linkedin_missing = pd.isna(row_data.get('Linkedin Profile Link'))
            website_missing = pd.isna(row_data.get('Company Website'))
            address_missing = pd.isna(row_data.get('Address')) # Assuming 'Address' column exists
            suite_missing = pd.isna(row_data.get('Suite')) # Check for Suite separately if needed

            # Skip if all target fields are already filled
            if not (email_missing or linkedin_missing or website_missing or address_missing or suite_missing):
                continue

            contact_name = row_data.get('Contact Name')
            company_name = row_data.get('Company Name')
            company_website = row_data.get('Company Website') # Get current value if exists

            if pd.isna(contact_name) or (pd.isna(company_name) and pd.isna(company_website)):
                print(f"\nSkipping row index {idx}: Missing Contact Name or both Company Name/Website.")
                continue

            name_parts = str(contact_name).split()
            first_name = name_parts[0] if name_parts else None
            last_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else None

            if not first_name or not last_name:
                print(f"\nSkipping row index {idx}: Could not extract First/Last name from '{contact_name}'.")
                continue

            print(f"\nEnriching row index {idx}: {contact_name} at {company_name or company_website}")

            found_email = None
            found_linkedin = None
            found_company_website = None # Initialize found values
            found_address = None         # Initialize found values
            title = str(row_data.get('Title')) if pd.notna(row_data.get('Title')) else None # Get title

            # --- Try Lead411 ---
            print(f"   Trying Lead411 API...")
            try:
                details = lead411_enricher.find_contact_details(
                    first_name, last_name, title,
                    str(company_name) if pd.notna(company_name) else None,
                    str(company_website) if pd.notna(company_website) else None,
                    str(row_data.get('State')) if pd.notna(row_data.get('State')) else None
                )
                
                if details:
                    if email_missing and details.get("email"):
                        found_email = details.get("email")
                        print(f"   Found Email via Lead411: {found_email}")
                        
                    if linkedin_missing and details.get("linkedin_url"):
                        found_linkedin = details.get("linkedin_url")
                        print(f"   Found LinkedIn via Lead411: {found_linkedin}")
                        
                    if website_missing and details.get("company_website"):
                        found_company_website = details.get("company_website")
                        print(f"   Found Company Website via Lead411: {found_company_website}")
                        
                    # Use found address for either 'Address' or 'Suite' if missing
                    if (address_missing or suite_missing) and details.get("address"):
                        found_address = details.get("address")
                        print(f"   Found Address via Lead411: {found_address}")

                    # Log if contact found but missing data wasn't retrieved
                    # (Refined logging check)
                    needed_but_not_found = []
                    if email_missing and not found_email: needed_but_not_found.append('Email')
                    if linkedin_missing and not found_linkedin: needed_but_not_found.append('LinkedIn')
                    if website_missing and not found_company_website: needed_but_not_found.append('Website')
                    if (address_missing or suite_missing) and not found_address: needed_but_not_found.append('Address/Suite')
                    
                    if needed_but_not_found and any([found_email, found_linkedin, found_company_website, found_address]):
                        print(f"   Lead411 found contact but did not return missing: {', '.join(needed_but_not_found)}.")

            except Exception as e:
                print(f"   Error during Lead411 enrichment orchestration: {e}")
                import traceback
                traceback.print_exc()

            # --- Update DataFrame ---            
            row_updated = False
            email_updated_this_row = False
            linkedin_updated_this_row = False
            website_updated_this_row = False # New flag
            address_updated_this_row = False # New flag

            if email_missing and pd.notna(found_email):
                self.df.at[idx, 'Email'] = found_email
                row_updated = True
                email_updated_this_row = True
                email_updates += 1

            if linkedin_missing and pd.notna(found_linkedin):
                self.df.at[idx, 'Linkedin Profile Link'] = found_linkedin
                row_updated = True
                linkedin_updated_this_row = True
                linkedin_updates += 1
                
            if website_missing and pd.notna(found_company_website):
                self.df.at[idx, 'Company Website'] = found_company_website
                row_updated = True
                website_updated_this_row = True
                website_updates += 1
                
            # Update 'Address' or 'Suite' based on what's missing and found
            if address_missing and pd.notna(found_address):
                 self.df.at[idx, 'Address'] = found_address
                 row_updated = True
                 address_updated_this_row = True
            elif suite_missing and pd.notna(found_address):
                # If Address exists but Suite is missing, put found address in Suite
                # This assumes 'Suite' column exists and is the target if 'Address' is already filled
                if 'Suite' in self.df.columns:
                    self.df.at[idx, 'Suite'] = found_address
                    row_updated = True
                    address_updated_this_row = True # Count as an address update
                else:
                    print("   Warning: 'Suite' column not found, cannot update.")
            
            # Increment address counter only once per row if an update happened
            if address_updated_this_row:
                address_updates += 1

            if row_updated:
                updates_made += 1
                enriched_rows.append({
                    'index': idx,
                    'contact_name': contact_name,
                    'company': company_name or company_website,
                    'email_updated': email_updated_this_row,
                    'found_email': found_email if email_updated_this_row else None,
                    'linkedin_updated': linkedin_updated_this_row,
                    'found_linkedin': found_linkedin if linkedin_updated_this_row else None,
                    'website_updated': website_updated_this_row,
                    'found_website': found_company_website if website_updated_this_row else None,
                    'address_updated': address_updated_this_row,
                    'found_address': found_address if address_updated_this_row else None
                })
            elif (email_missing or linkedin_missing or website_missing or address_missing or suite_missing):
                 print(f"   Could not find missing data for {contact_name} using Lead411.")

            time.sleep(1) # API delay

        # Print summary of updates
        print(f"\nEnrichment complete. Updated {updates_made} rows total using Lead411:")
        print(f"- Email addresses filled: {email_updates}")
        print(f"- LinkedIn URLs filled: {linkedin_updates}")
        print(f"- Company Websites filled: {website_updates}") # New summary line
        print(f"- Addresses/Suites filled: {address_updates}") # New summary line

        if enriched_rows:
            print("\nDetails of enriched rows:")
            for row in enriched_rows:
                details = []
                if row['email_updated']: details.append(f"Email: {row['found_email']}")
                if row['linkedin_updated']: details.append(f"LinkedIn: {row['found_linkedin']}")
                if row['website_updated']: details.append(f"Website: {row['found_website']}")
                if row['address_updated']: details.append(f"Address/Suite: {row['found_address']}")
                if details:
                    print(f"Row {row['index']} - {row['contact_name']} at {row['company']}: {', '.join(details)}")

        if updates_made > 0:
            self.save_changes()
        else:
            print("\nNo changes were made to the file.")

    def save_changes(self, output_path: Optional[str] = None) -> None:
        """Saves the current state of the DataFrame back to an Excel file."""
        if self.df is None:
            print("Error: No DataFrame loaded to save.")
            return

        save_path = Path(output_path) if output_path else self.filepath
        print(f"\nSaving updated DataFrame to '{save_path}'...")
        try:
            save_path.parent.mkdir(parents=True, exist_ok=True)
            self.df.to_excel(save_path, index=False)
            print(f"DataFrame successfully saved.")
        except Exception as e:
            print(f"Error saving DataFrame to {save_path}: {e}")


if __name__ == "__main__":
    print("Script started.")
    try:
        # Initialize the scraper (loads data)
        scraper = PersonDetailsScraperPandas(filename=XLSX_FILE)

        # Get Lead411 API credentials from environment
        lead411_email = os.getenv("LEAD411_EMAIL")
        lead411_password = os.getenv("LEAD411_PASSWORD")
        # Removed Hunter and SerpAPI key loading

        # Initialize Lead411 enricher
        lead411_enricher = None
        if lead411_email and lead411_password:
            print("Lead411 credentials found.")
            try:
                lead411_enricher = Lead411Enricher(lead411_email, lead411_password)
                # Check if authentication was successful
                if lead411_enricher._get_token() is None:
                    print("Error: Lead411 authentication failed. Please check credentials.")
                    sys.exit(1)
            except ConnectionError as e: # Catch auth errors if raised
                 print(f"Error initializing Lead411: {e}")
                 sys.exit(1)
        else:
            print("Error: Lead411 credentials (LEAD411_EMAIL, LEAD411_PASSWORD) not found in .env file.")
            print("Cannot proceed without Lead411 credentials.")
            sys.exit(1)

        # Removed checks for Hunter/SerpAPI keys

        # Run the enrichment process ONLY with Lead411
        scraper.enrich_leads(lead411_enricher=lead411_enricher)

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except IOError as e:
        print(f"Error reading/processing file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print("Script finished.")