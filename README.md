# CA Offline Suite - Bank Statement Analyzer

This project provides a FastAPI backend service to analyze bank statements provided as PDF files. It extracts transaction data, categorizes it, and can generate output artifacts like Excel files.

## Project Structure

The core logic resides within the `backend/` directory:

*   `main.py`: The main FastAPI application entry point, defining API routes.
*   `utils.py`: Utility functions for file paths.
*   `common_functions.py`: Shared parsing and categorization logic.
*   `code_for_extraction.py` / `old_bank_extractions.py`: Core PDF data extraction logic.
*   `tax_professional/banks/`: Bank-specific statement handling logic.
*   `models/`: Machine learning models used for analysis.
*   `requirements.txt`: Python package dependencies.
*   `test_endpoint.py`: Integration tests for the API endpoints.

Sample PDF files can be found in `backend/pdfs/`. Output files (like Excel reports) are typically saved to the `output/` directory in the project root by default (configured via `backend/utils.get_saved_excel_dir()`).

## Setup

1.  **Clone the repository:**
    ```bash
    # If you haven't already
    git clone <repository_url>
    cd ca-offline-suite-uat 3
    ```

2.  **Create and activate a virtual environment (recommended):**
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    # On Windows use: .venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r backend/requirements.txt
    ```

## Running the Application

To start the FastAPI development server:

```bash
uvicorn backend.main:app --reload --host 0.0.0.0 --port 8000
```

This will make the API accessible at `http://localhost:8000` (or your machine's IP address on port 8000). The `--reload` flag automatically restarts the server when code changes are detected.

## Testing

Integration tests are located in `backend/test_endpoint.py`. You can run them using a test runner like `pytest`:

```bash
pytest backend/test_endpoint.py
``` 
