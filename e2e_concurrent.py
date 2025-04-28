# concurrency_test.py (Modified)
import json
import threading
import time
import uuid
from pathlib import Path

import requests  # pip install requests

# --- Configuration ---
BASE_URL = "http://localhost:7500"
SETUP_EMAIL = f"test_concurrency_{uuid.uuid4().hex[:8]}@example.com"
SETUP_PASSWORD = "TestPassword_123!"
CONTAINER_PDF_PATH = "/app/backend/pdfs/Manish Hdfc.pdf"
NUM_REQUESTS = 4
REQUEST_TIMEOUT = 300
# Directory to save downloaded Excel files locally
OUTPUT_DIR = Path("./concurrent_outputs")
# ---------------------

session = requests.Session()
common_headers = {}

# Ensure output directory exists
OUTPUT_DIR.mkdir(exist_ok=True)


def setup_auth():
    """Registers a user, logs in, gets API key, returns headers."""
    global common_headers
    token = None
    api_key = None
    try:
        print("[Setup] Registering user...")
        reg_payload = {"email": SETUP_EMAIL, "password": SETUP_PASSWORD}
        response = session.post(f"{BASE_URL}/auth/register", json=reg_payload)
        response.raise_for_status()

        print("[Setup] Logging in...")
        login_payload = {"username": SETUP_EMAIL, "password": SETUP_PASSWORD}
        response = session.post(f"{BASE_URL}/auth/jwt/login", data=login_payload)
        response.raise_for_status()
        token = response.json().get("access_token")

        print("[Setup] Getting API Key...")
        headers = {"Authorization": f"Bearer {token}"}
        response = session.post(f"{BASE_URL}/users/api-key", headers=headers)
        response.raise_for_status()
        api_key = response.json().get("api_key")

        if not token or not api_key:
            raise ValueError("Failed to get token or API key during setup.")

        common_headers = {
            "Authorization": f"Bearer {token}",
            "X-API-Key": api_key,
            # Content-Type for analyze payload
            # "Content-Type": "application/json",
        }
        print("[Setup] Authentication headers ready.")
        return True

    except requests.exceptions.RequestException as e:
        print(f"[Setup Error] HTTP Request Error: {e}")
        if e.response is not None:
            print(f"      Status Code: {e.response.status_code}")
            try:
                print(f"      Response Body: {e.response.json()}")
            except json.JSONDecodeError:
                print(f"      Response Body: {e.response.text}")
        return False
    except Exception as e:
        print(f"[Setup Error] Unexpected error: {e}")
        return False


def run_analysis_and_export(thread_id: int, results: list):
    """Sends analysis, then export request, saves file, records result."""
    start_time = time.monotonic()
    status = "FAILED"
    status_code_analyze = None
    status_code_export = None
    error = None
    analyze_duration = None
    export_duration = None
    analyze_data = None
    local_excel_path = None

    # Use a non-session requests object for thread safety potentially
    thread_session = requests  # Or requests.Session() if needed

    try:
        # --- 1. Analyze Statements ---
        analyze_start = time.monotonic()
        print(f"[Thread-{thread_id}] Sending analysis request...")
        analyze_payload = {
            "bank_names": ["HDFC"],
            "pdf_paths": [CONTAINER_PDF_PATH],
            "passwords": [""],
            "start_date": ["01-04-2023"],
            "end_date": ["01-04-2024"],
            "ca_id": f"CONCURRENCY_{thread_id}",
        }
        # Need to explicitly set Content-Type for POST with JSON
        analyze_headers = common_headers.copy()
        analyze_headers["Content-Type"] = "application/json"

        response = thread_session.post(
            f"{BASE_URL}/statements/",
            json=analyze_payload,
            headers=analyze_headers,
            timeout=REQUEST_TIMEOUT,
        )
        status_code_analyze = response.status_code
        response.raise_for_status()
        analyze_data = response.json().get("data")
        if not analyze_data:
            raise ValueError("Analysis did not return 'data' field.")
        analyze_duration = time.monotonic() - analyze_start
        print(
            f"[Thread-{thread_id}] Analysis successful ({status_code_analyze}, {analyze_duration:.2f}s)."
        )

        # --- 2. Export Excel ---
        tx_sheet = analyze_data.get("Transactions")
        name_sheet = analyze_data.get("Name Acc No")
        if not tx_sheet or not name_sheet:
            raise ValueError("Required sheets not found for export.")

        export_start = time.monotonic()
        print(f"[Thread-{thread_id}] Sending export request...")
        export_payload = {
            "transaction_data": tx_sheet,
            "name_n_num": name_sheet,
            "case_name": f"Concurrency_Export_{thread_id}",
        }
        # Use the same auth headers
        export_headers = common_headers.copy()
        export_headers["Content-Type"] = "application/json"

        response = thread_session.post(
            f"{BASE_URL}/export/",
            json=export_payload,
            headers=export_headers,
            stream=True,  # Important for file download
            timeout=120,  # Timeout for export
        )
        status_code_export = response.status_code
        response.raise_for_status()
        export_duration = time.monotonic() - export_start
        print(
            f"[Thread-{thread_id}] Export successful ({status_code_export}, {export_duration:.2f}s)."
        )

        # --- 3. Save Excel File Locally ---
        local_excel_path = OUTPUT_DIR / f"concurrent_download_{thread_id}.xlsx"
        print(f"[Thread-{thread_id}] Saving Excel file to: {local_excel_path}")
        with open(local_excel_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"[Thread-{thread_id}] Excel file saved.")
        status = "SUCCESS"  # Mark overall success only if file is saved

    except requests.exceptions.RequestException as e:
        error = f"HTTP Error: {e}"
        if e.response is not None:
            error += f" (Status: {e.response.status_code})"
            try:
                error += f" Body: {e.response.json()}"
            except:
                error += f" Body: {e.response.text[:100]}..."
        print(f"[Thread-{thread_id}] Request failed: {error}")
    except Exception as e:
        error = f"Unexpected Error: {e}"
        print(f"[Thread-{thread_id}] Request failed: {error}")
    finally:
        end_time = time.monotonic()
        results.append(
            {
                "thread_id": thread_id,
                "status": status,
                "status_code_analyze": status_code_analyze,
                "status_code_export": status_code_export,
                "duration_analyze": analyze_duration,
                "duration_export": export_duration,
                "duration_total": end_time - start_time,
                "error": error,
                "output_file": str(local_excel_path)
                if local_excel_path and status == "SUCCESS"
                else None,
            }
        )


# --- Main Execution ---
if __name__ == "__main__":
    print("--- Starting Concurrency Test (with Export) ---")

    if not setup_auth():
        print("\n[!!!] Setup failed. Cannot proceed with concurrency test.")
        exit(1)

    threads = []
    results = []
    start_overall = time.monotonic()

    print(f"\n[*] Launching {NUM_REQUESTS} concurrent analysis & export requests...")
    for i in range(NUM_REQUESTS):
        thread = threading.Thread(target=run_analysis_and_export, args=(i + 1, results))
        threads.append(thread)
        thread.start()
        time.sleep(0.1)

    print("[*] Waiting for all requests to complete...")
    for thread in threads:
        thread.join()

    end_overall = time.monotonic()
    print("\n--- Concurrency Test Results (with Export) ---")

    successful_requests = sum(1 for r in results if r["status"] == "SUCCESS")
    failed_requests = NUM_REQUESTS - successful_requests

    print(f"Total Requests Sent: {NUM_REQUESTS}")
    print(f"Successful Requests: {successful_requests}")
    print(f"Failed Requests    : {failed_requests}")
    print(f"\nTotal Time Elapsed : {end_overall - start_overall:.2f} seconds")

    if results:
        total_analyze_dur = sum(
            r["duration_analyze"] for r in results if r["duration_analyze"]
        )
        total_export_dur = sum(
            r["duration_export"] for r in results if r["duration_export"]
        )
        avg_total_dur = sum(r["duration_total"] for r in results) / len(results)
        max_total_dur = max(r["duration_total"] for r in results)

        if successful_requests > 0:
            print(
                f"Avg Analyze Duration : {(total_analyze_dur / successful_requests):.2f} seconds"
            )
            print(
                f"Avg Export Duration  : {(total_export_dur / successful_requests):.2f} seconds"
            )
        print(f"Average Total Duration : {avg_total_dur:.2f} seconds (per thread)")
        print(f"Max Total Duration   : {max_total_dur:.2f} seconds (slowest thread)")

    if failed_requests > 0:
        print("\n[!!!] Some requests failed:")
        for r in results:
            if r["status"] == "FAILED":
                print(f"  - Thread {r['thread_id']}: {r['error']}")
    else:
        print("\n[*] All concurrent requests completed successfully.")
        print(f"[*] Find the downloaded Excel files in: {OUTPUT_DIR.resolve()}")

    print("\n--- End Concurrency Test (with Export) ---")
