import json
import math

import requests

url = "http://localhost:7500/statements/"

payload = {
    "bank_names": ["HDFC"],
    "pdf_paths": [
        "/Users/sanchaythalnerkar/Documents/ca-offline-suite-uat 3/backend/pdfs/Manish Hdfc.pdf"
    ],
    "passwords": [""],
    "start_date": ["01-04-2023"],
    "end_date": ["01-04-2024"],
    "ca_id": "ADFC_1234",
}

try:
    print("Sending request to:", url)
    response = requests.post(url, json=payload)
    print(f"Status Code: {response.status_code}")
    print("\nResponse:")
    print(json.dumps(response.json(), indent=2))

    # After the analyze-statements call, if successful, call /excel-download/ endpoint
    if response.status_code == 200:
        resp_json = response.json()
        # The "data" field returned by analyze-statements is a JSON string; parse it to dict
        try:
            sheets_data = json.loads(resp_json.get("data", {}))
        except (ValueError, TypeError):
            print("Could not parse sheets data returned by /analyze-statements/")
            sheets_data = {}

        # Extract the sheets required for the next endpoint
        transaction_data = sheets_data.get("Transactions", [])
        name_n_num = sheets_data.get("Name Acc No", [])

        if transaction_data and name_n_num:
            # Sanitize any NaN/Inf values so that requests' json encoder does not raise
            def _sanitize(obj):
                if isinstance(obj, float):
                    return obj if math.isfinite(obj) else None
                elif isinstance(obj, list):
                    return [_sanitize(item) for item in obj]
                elif isinstance(obj, dict):
                    return {k: _sanitize(v) for k, v in obj.items()}
                else:
                    return obj

            excel_payload = _sanitize(
                {
                    "transaction_data": transaction_data,
                    "name_n_num": name_n_num,
                    "case_name": "Test_Case_001",
                }
            )
            excel_url = "http://localhost:7500/excel-download/"
            print(f"\nSending request to: {excel_url}")
            excel_resp = requests.post(excel_url, json=excel_payload)
            print(f"Status Code: {excel_resp.status_code}")
            try:
                print(json.dumps(excel_resp.json(), indent=2))
            except Exception:
                print(excel_resp.text)
        else:
            print(
                "Required sheets not present in the response; skipping /excel-download/ test."
            )

except requests.exceptions.ConnectionError:
    print(
        "Error: Could not connect to the server. Make sure the FastAPI server is running on port 4000"
    )
except Exception as e:
    print(f"Error occurred: {str(e)}")
