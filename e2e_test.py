# e2e_test.py (Modified for custom local path)
import argparse  # Import argparse
import json
import uuid
from pathlib import Path

import requests  # pip install requests

# --- Configuration ---
BASE_URL = "https://shark-app-anlu9.ondigitalocean.app"  # Point to the deployed App Platform URL
SETUP_PASSWORD = "TestPassword_123!"
# Where to save the downloaded Excel file *locally*
LOCAL_OUTPUT_FILENAME = "downloaded_statement.xlsx"
# ---------------------

session = requests.Session()


def run_flow(local_pdf_path: Path):
    """Runs the full E2E flow using the provided local PDF path."""
    token = None
    api_key = None
    analyze_data = None

    # Use the directory of the script or current dir for output
    output_file_path = Path(__file__).parent / LOCAL_OUTPUT_FILENAME
    if not local_pdf_path.is_file():
        print(f"[!!!] Error: Local PDF file not found at: {local_pdf_path}")
        return

    # Use a unique email for each test run based on PDF name or random
    base_email_name = local_pdf_path.stem.replace(
        " ", "_"
    )  # Use PDF name for email base
    TEST_EMAIL = f"test_e2e_{base_email_name}_{uuid.uuid4().hex[:6]}@example.com"
    print(f"[*] Using test email: {TEST_EMAIL}")

    try:
        # 1. Register User
        print(f"[*] Registering user: {TEST_EMAIL}...")
        reg_payload = {"email": TEST_EMAIL, "password": SETUP_PASSWORD}
        response = session.post(f"{BASE_URL}/auth/register", json=reg_payload)
        response.raise_for_status()  # Raise exception for bad status codes (4xx or 5xx)
        user_info = response.json()
        print(f"[+] Registration successful (User ID: {user_info.get('id')})")

        # 2. Login
        print(f"[*] Logging in as: {TEST_EMAIL}...")
        login_payload = {"username": TEST_EMAIL, "password": SETUP_PASSWORD}
        # Login requires form data
        response = session.post(f"{BASE_URL}/auth/jwt/login", data=login_payload)
        response.raise_for_status()
        token = response.json().get("access_token")
        if not token:
            raise ValueError("Login failed: No access token.")
        print("[+] Login successful.")

        # 3. Get API Key
        print("[*] Requesting API Key...")
        headers = {"Authorization": f"Bearer {token}"}
        response = session.post(f"{BASE_URL}/users/api-key", headers=headers)
        response.raise_for_status()
        api_key = response.json().get("api_key")
        if not api_key:
            raise ValueError("API Key request failed.")
        print("[+] API Key received.")
        auth_headers = {"Authorization": f"Bearer {token}", "X-API-Key": api_key}

        # 4. Upload Local PDF
        print(f"[*] Uploading local PDF: {local_pdf_path}...")
        # httpie equivalent: files@'path/to/local.pdf'
        files_payload = {
            "files": (
                local_pdf_path.name,
                open(local_pdf_path, "rb"),
                "application/pdf",
            )
        }
        # Use multipart/form-data for file uploads, pass auth headers
        response = session.post(
            f"{BASE_URL}/files/upload-pdf",
            files=files_payload,  # Use files= parameter
            headers=auth_headers,  # Only auth needed, Content-Type set by requests
        )
        response.raise_for_status()
        upload_result = response.json()
        internal_paths = upload_result.get("internal_paths")
        if not internal_paths:
            raise ValueError("Upload failed: No internal paths returned.")
        internal_path = internal_paths[
            0
        ]  # Assuming single file upload for simplicity now
        print(f"[+] Upload successful. Internal path: {internal_path}")

        # 5. Analyze Statements using Internal Path
        print(f"[*] Analyzing statement using internal path: {internal_path}...")
        # Adapt bank name/dates if needed, assuming HDFC for now
        analyze_payload = {
            "bank_names": ["HDFC"],
            "pdf_paths": [internal_path],  # Use the path returned by upload endpoint
            "passwords": [""],
            "start_date": ["01-04-2023"],
            "end_date": ["01-04-2024"],
            "ca_id": f"E2E_{local_pdf_path.stem[:10]}",
        }
        # Add Content-Type for JSON payload
        analyze_headers = auth_headers.copy()
        analyze_headers["Content-Type"] = "application/json"
        response = session.post(
            f"{BASE_URL}/statements/",
            json=analyze_payload,
            headers=analyze_headers,
            timeout=300,
        )  # Increased timeout for analysis
        response.raise_for_status()
        analyze_data = response.json().get("data")
        if not analyze_data:
            raise ValueError("Analysis did not return 'data' field.")
        print("[+] Analysis successful.")

        # 6. Export Excel
        print("[*] Requesting Excel Export...")
        tx_sheet = analyze_data.get("Transactions")
        name_sheet = analyze_data.get("Name Acc No")
        if not tx_sheet or not name_sheet:
            raise ValueError("Required sheets not found.")

        export_payload = {
            "transaction_data": tx_sheet,
            "name_n_num": name_sheet,
            "case_name": f"E2E_{local_pdf_path.stem[:10]}_Export",
        }
        export_headers = auth_headers.copy()
        export_headers["Content-Type"] = "application/json"
        response = session.post(
            f"{BASE_URL}/export/",
            json=export_payload,
            headers=export_headers,
            stream=True,
        )
        response.raise_for_status()
        print("[+] Export request successful (Status Code:", response.status_code, ")")

        # 7. Save Excel File Locally
        print(f"[*] Saving downloaded Excel file to: {output_file_path}")
        with open(output_file_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"[+] Excel file saved successfully to {output_file_path}!")

    except requests.exceptions.RequestException as e:
        print(f"\n[!!!] HTTP Request Error: {e}")
        if e.response is not None:
            print(f"      Status Code: {e.response.status_code}")
            try:
                print(f"      Response Body: {e.response.json()}")
            except json.JSONDecodeError:
                print(f"      Response Body: {e.response.text}")
    except Exception as e:
        print(f"\n[!!!] An unexpected error occurred: {e}")

    finally:
        print("\n[*] End-to-end test script finished.")


if __name__ == "__main__":
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Run E2E test for CA Offline Suite API."
    )
    parser.add_argument(
        "pdf_file",  # Make the PDF file path a required positional argument
        type=str,
        help="Path to the local PDF file to upload and process.",
    )
    args = parser.parse_args()

    # Convert string path to Path object
    local_pdf_path_arg = Path(args.pdf_file)

    run_flow(local_pdf_path_arg)  # Pass the Path object to the function
