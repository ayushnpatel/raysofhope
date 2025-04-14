import logging
import os
import sys
from dataclasses import dataclass, fields, field
from datetime import datetime
from typing import List, Optional, Dict, Union

import pandas as pd

# Get a logger for this step
logger = logging.getLogger(__name__)

@dataclass
class BillingData:
    """Represents a single row of billing data from the Excel sheet."""
    patient_last_name: str
    patient_first_name: str
    patient_birth_date: str
    insured_id: str
    diagnosis_code: str
    prior_auth_num: str
    procedures_services: str
    tier: str
    rate: float
    units: int
    dollars: float
    approval_start_date_str: str = field(repr=False) # Store original string, hide from repr
    approval_end_date: str
    units_approved: Optional[int]
    # Parsed date field for internal use
    approval_start_date: datetime = field(init=False, repr=False) # Don't include in __init__, hide from repr

    def __post_init__(self):
        """Parse the date string after initialization."""
        try:
            # Attempt to parse various date formats
            # Pandas to_datetime is generally good at this
            parsed_date = pd.to_datetime(self.approval_start_date_str, errors='coerce')
            if pd.isna(parsed_date):
                 raise ValueError(f"Could not parse date: {self.approval_start_date_str}")
            self.approval_start_date = parsed_date.to_pydatetime()
        except ValueError as e:
            logger.error(f"Error parsing date '{self.approval_start_date_str}' for {self.patient_first_name} {self.patient_last_name}: {e}")
            # Assign a default minimum date to handle errors gracefully during comparison
            # Or you might want to raise the error or skip the record entirely
            self.approval_start_date = datetime.min
            logger.warning(f"Assigning minimum date to record for {self.patient_first_name} {self.patient_last_name} due to parsing error.")


    @classmethod
    def get_expected_columns(cls) -> List[str]:
        """Returns the list of expected column names based on the dataclass fields."""
        expected_map = {
            'patient_last_name': "PATIENT'S LAST NAME",
            'patient_first_name': "PATIENT'S FIRST NAME",
            'patient_birth_date': "PATIENT'S BIRTH DATE",
            'insured_id': "INSURED'S I.D. #",
            'diagnosis_code': "DIAGNOSIS CODE",
            'prior_auth_num': "PRIOR AUTH #",
            'procedures_services': "PROCEDURES,SERVICES",
            'tier': "Tier",
            'rate': "RATE",
            'units': "UNITS",
            'dollars': "$$$",
            'approval_start_date_str': "APPROVAL START DATE", # Map to original string field
            'approval_end_date': "APPROVAL END DATE",
            'units_approved': "UNITS",
        }
        # Get fields excluding the internally calculated 'approval_start_date'
        return [expected_map[f.name] for f in fields(cls) if f.init]


def _parse_and_format_birth_date(date_input: Union[str, datetime, None]) -> str:
    """Attempts to parse a date string or datetime object from various formats
    and returns it consistently formatted as mm/dd/yyyy.

    Handles common Excel/pandas formats like datetime objects, mm/dd/yyyy, mm/dd/yy.

    Args:
        date_input: The date string, datetime object, or None to parse.

    Returns:
        The date string formatted as mm/dd/yyyy, or an empty string
        if input is None or parsing fails.
    """
    if date_input is None:
        return ""

    parsed_date: Optional[datetime] = None

    if isinstance(date_input, datetime):
        parsed_date = date_input
    elif isinstance(date_input, str):
        date_str = date_input.strip()
        if not date_str:
            return ""
        # Add more formats if needed, prioritize expected ones
        formats_to_try = ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]
        for fmt in formats_to_try:
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                break # Success
            except ValueError:
                continue # Try the next format
    else:
        # Handle other potential types if necessary, like pd.Timestamp
        try:
            # Attempt conversion if it's something convertible like pd.Timestamp
            parsed_date = pd.to_datetime(date_input).to_pydatetime()
        except Exception:
             logger.warning(f"Could not handle birth date input type '{type(date_input)}': {date_input}. Returning empty string.")
             return ""


    if parsed_date:
        return parsed_date.strftime("%m/%d/%Y")
    else:
        # Log warning only if it was a string we failed to parse
        if isinstance(date_input, str):
            logger.warning(f"Could not parse birth date string '{date_input}' using known formats. Returning empty string.")
        elif not isinstance(date_input, datetime): # Avoid logging if it was already a datetime
             logger.warning(f"Failed to parse birth date input: {date_input}. Returning empty string.")
        return ""


def _find_excel_file(file_path_arg: Optional[str]) -> str:
    """Determines the correct Excel file path to use.

    Priority:
    1. Command-line argument (`file_path_arg`)
    2. EXCEL_FILE_PATH environment variable
    3. Default: '../billing.xlsx' relative to the script's parent directory.

    Raises:
        FileNotFoundError: If the final determined file path does not exist.

    Returns:
        The validated path to the Excel file.
    """
    if file_path_arg:
        logger.info(f"Using Excel file path from command line: {file_path_arg}")
        if os.path.exists(file_path_arg):
            return file_path_arg
        else:
            logger.warning(f"File specified via command line not found: {file_path_arg}")
            # Fall through to try default

    env_path = os.getenv("EXCEL_FILE_PATH")
    if env_path:
        logger.info(f"Using Excel file path from environment variable: {env_path}")
        if os.path.exists(env_path):
            return env_path
        else:
            logger.warning(f"File specified via environment variable not found: {env_path}")
            # Fall through to try default

    # Default path: ../billing.xlsx relative to *this* file's location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_path = os.path.abspath(os.path.join(script_dir, '..', 'billing.xlsx'))
    logger.info(f"Trying default Excel file path: {default_path}")
    if os.path.exists(default_path):
        return default_path
    else:
        raise FileNotFoundError(f"Default billing file not found at {default_path}. Please provide a valid path via CLI argument or EXCEL_FILE_PATH environment variable.")


def parse_excel_step(file_path_arg: Optional[str] = None) -> List[BillingData]:
    """Parses the billing data from the specified Excel file,
       deduplicating by patient name based on the latest approval start date.

    Args:
        file_path_arg: Optional path to the Excel file from command line.

    Returns:
        A deduplicated list of BillingData objects.

    Raises:
        FileNotFoundError: If the Excel file cannot be found.
        ValueError: If the Excel file is missing required columns.
    """
    logger.info("--- Step: Parse Excel --- ")

    try:
        excel_file_path = _find_excel_file(file_path_arg)
        logger.info(f"Reading Excel file: {excel_file_path}")
        # Specify header row (6th row = index 5), let pandas detect columns
        # Explicitly set dtype for Insured ID to prevent float conversion
        df = pd.read_excel(
            excel_file_path,
            engine='openpyxl',
            header=5,
            dtype={"INSURED'S I.D. #": str} # Ensure ID is read as string
        )

        expected_columns = BillingData.get_expected_columns()
        actual_columns = [str(col).strip() for col in df.columns]
        logger.debug(f"Expected columns: {expected_columns}")
        logger.debug(f"Actual columns found: {actual_columns}")

        # Ensure all required columns are present before proceeding
        # Note: We check based on the names needed for dataclass initialization (`f.init`)
        required_init_cols = BillingData.get_expected_columns()
        missing_columns = [col for col in required_init_cols if col not in actual_columns]
        if missing_columns:
            raise ValueError(f"Excel file '{excel_file_path}' is missing required columns: {', '.join(missing_columns)}")

        all_parsed_rows: List[BillingData] = [] # Store all initially parsed rows
        logger.info(f"Parsing {len(df)} rows from Excel file...")
        for index, row in df.iterrows():
            excel_row_num = index + 7 # header=5 means data starts at row 7 (index 0)
            try:
                # Pre-filter check: Check Col A for '#' and Col B (Lat Name) for content
                # Access by position (iloc) as first col might not have a reliable name
                col_a_value = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                col_b_value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else "" # Assumes Col B is Last Name

                if not  col_a_value or not col_b_value:
                    logger.debug(f"Skipping Excel row {excel_row_num} based on pre-filter (Col A no '#' or Col B empty). Col A: '{col_a_value[:20]}...', Col B: '{col_b_value[:20]}...'")
                    continue

                logger.debug(f"Processing Excel row {excel_row_num}: {row.to_dict()}")

                patient_last_name = col_b_value # Use the already fetched value
                patient_first_name = str(row.get("PATIENT'S FIRST NAME", ""))
                approval_start_date_str = str(row.get("APPROVAL START DATE", ""))

                # Check other essential fields after pre-filter passed
                if not patient_last_name or not approval_start_date_str:
                    logger.warning(f"Skipping Excel row {excel_row_num} due to missing essential field (Last Name or Approval Start Date). Last Name: '{patient_last_name}', Approval Start: '{approval_start_date_str}'")
                    continue

                # --- Start creating BillingData ---
                # Parse and format birth date
                # Pass the raw value from the row directly
                raw_birth_date_input = row.get("PATIENT'S BIRTH DATE")
                formatted_birth_date = _parse_and_format_birth_date(raw_birth_date_input)

                data = BillingData(
                    patient_last_name=patient_last_name,
                    patient_first_name=patient_first_name,
                    patient_birth_date=formatted_birth_date, # Use formatted date
                    insured_id=str(row.get("INSURED'S I.D. #", "")),
                    diagnosis_code=str(row.get("DIAGNOSIS CODE", "")),
                    prior_auth_num=str(row.get("PRIOR AUTH #")) if pd.notna(row.get("PRIOR AUTH #")) else None,
                    procedures_services=str(row.get("PROCEDURES,SERVICES", "")),
                    tier=str(row.get("Tier")) if pd.notna(row.get("Tier")) else None,
                    rate=float(row.get("RATE")) if pd.notna(row.get("RATE")) else None,
                    units=int(row.get("UNITS", 0)) if pd.notna(row.get("UNITS")) else 0, # Default to 0 if missing/NaN
                    dollars=float(row.get("$$$", 0.0)) if pd.notna(row.get("$$$")) else 0.0, # Default to 0.0
                    approval_start_date_str=approval_start_date_str,
                    approval_end_date=str(row.get("APPROVAL END DATE", "")),
                    units_approved = None # Placeholder - Re-evaluate based on actual columns/logic
                )
                # Only append if date parsing didn't result in datetime.min
                if data.approval_start_date != datetime.min:
                    all_parsed_rows.append(data)
                else:
                    logger.warning(f"Skipping Excel row {excel_row_num} for {data.patient_first_name} {data.patient_last_name} in final list due to date parsing error.")

            except KeyError as e:
                 # This block might be less likely now with .get(), but kept for safety
                 logger.error(f"Error parsing Excel row {excel_row_num}: Unexpected KeyError '{e}'. Row data: {row.to_dict()}")
                 logger.warning(f"Skipping Excel row {excel_row_num} due to unexpected KeyError.")
            except (TypeError, ValueError, IndexError) as e: # Catch other parsing errors
                logger.error(f"Error parsing Excel row {excel_row_num}: {e}. Row data: {row.to_dict()}")
                logger.warning(f"Skipping Excel row {excel_row_num} due to parsing error.")

        logger.info(f"Successfully parsed {len(all_parsed_rows)} rows with valid data.")

        # --- Deduplication Logic ---
        if not all_parsed_rows:
            logger.warning("No valid data rows found or parsed from the Excel file.")
            return [] # Return empty list

        logger.info("Deduplicating records based on patient name and latest approval start date...")
        latest_records: Dict[tuple[str, str], BillingData] = {}

        for record in all_parsed_rows:
            key = (record.patient_first_name.strip().upper(), record.patient_last_name.strip().upper())
            if key not in latest_records or record.approval_start_date > latest_records[key].approval_start_date:
                latest_records[key] = record

        deduplicated_list = list(latest_records.values())
        logger.info(f"Reduced {len(all_parsed_rows)} records to {len(deduplicated_list)} unique patient records by latest approval date.")
        # --- End Deduplication ---

        logger.info("--- Step: Parse Excel Complete ---")
        return deduplicated_list # Return the deduplicated list

    except FileNotFoundError as e:
        logger.error(f"Excel file parsing failed: {e}")
        sys.exit(1)
    except ValueError as e:
        logger.error(f"Excel file validation failed: {e}")
        sys.exit(1)
    except Exception as e:
        logger.exception("An unexpected error occurred during Excel parsing:")
        sys.exit(1) 