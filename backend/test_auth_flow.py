import os
import signal
import subprocess
import sys
import threading
from pathlib import Path

import requests
from dotenv import load_dotenv

HOST = "127.0.0.1"
PORT = 7500
BASE = f"http://{HOST}:{PORT}"


def run_server() -> subprocess.Popen:
    """Launch uvicorn in a subprocess and return the Popen object."""
    env = os.environ.copy()
    # ensure .env variables (incl. SENTRY_DSN) are propagated
    ENV_FILE = Path(__file__).resolve().parent / ".env"
    if ENV_FILE.exists():
        load_dotenv(dotenv_path=ENV_FILE)
        with ENV_FILE.open() as f:
            for line in f:
                if line.strip() and not line.strip().startswith("#") and "=" in line:
                    k, v = [p.strip() for p in line.split("=", 1)]
                    env.setdefault(k, v)
    # ensure reload off for test speed
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "uvicorn",
            "backend.main:app",
            "--port",
            str(PORT),
            "--host",
            HOST,
        ],
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    # wait for banner
    started = False
    while True:
        line = proc.stdout.readline()
        if not line:
            break
        print(line.strip())
        if "Application startup complete" in line:
            started = True
            break
    if not started:
        raise RuntimeError("Server failed to start")
    return proc


# continuously pipe server output
def stream_logs(proc: subprocess.Popen):
    def _reader():
        for line in proc.stdout:
            print("[SERVER]", line.rstrip())

    t = threading.Thread(target=_reader, daemon=True)
    t.start()


def stop_server(proc: subprocess.Popen) -> None:
    proc.send_signal(signal.SIGINT)
    proc.wait(timeout=5)


def main() -> None:
    # kill any existing instance
    os.system("pkill -f 'uvicorn .*backend.main:app' || true")

    proc = run_server()
    stream_logs(proc)
    try:
        print("\n=== Register ===")
        r = requests.post(
            f"{BASE}/auth/register",
            json={"email": "flow@ex.com", "password": "Pass123!"},
            timeout=10,
        )
        if r.status_code not in (200, 201):
            print(r.status_code, r.text)
        if r.status_code == 400 and "REGISTER_USER_ALREADY_EXISTS" in r.text:
            print("User already exists, continuing…")
        else:
            print(r.status_code, r.text)

        print("\n=== Login ===")
        r = requests.post(
            f"{BASE}/auth/jwt/login",
            data={"username": "flow@ex.com", "password": "Pass123!"},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=10,
        )
        print(r.status_code, r.text)
        token = r.json()["access_token"]

        print("\n=== Get me (token) ===")
        r = requests.get(
            f"{BASE}/users/me", headers={"Authorization": f"Bearer {token}"}, timeout=10
        )
        print(r.status_code, r.text)

        print("\n=== Create API key ===")
        r = requests.post(
            f"{BASE}/users/api-key",
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
        )
        print(r.status_code, r.text)
        api_key = r.json()["api_key"]
        print("Generated key (last4)", api_key[-4:])

        headers = {"Authorization": f"Bearer {token}", "X-API-Key": api_key}

        print("\n=== Analyze (new router) ===")
        sample_pdf = Path(__file__).resolve().parent / "pdfs" / "Manish Hdfc.pdf"
        if not sample_pdf.exists():
            print(f"Sample PDF not found at {sample_pdf}")
            return
        analyze_payload = {
            "bank_names": ["HDFC"],
            "pdf_paths": [str(sample_pdf.resolve())],
            "passwords": [""],
            "start_date": ["01-04-2023"],
            "end_date": ["01-04-2024"],
            "ca_id": "CASE_001",
        }
        analyze_resp = requests.post(
            f"{BASE}/statements/", json=analyze_payload, headers=headers, timeout=300
        )
        print(analyze_resp.status_code)
        if not analyze_resp.ok:
            print("Analyze error:", analyze_resp.text)
        sheets = analyze_resp.json().get("data", {}) if analyze_resp.ok else {}

        tx = sheets.get("Transactions", [])
        name_num = sheets.get("Name Acc No", [])
        if tx and name_num:
            print("\n=== Export ===")
            excel_payload = {
                "transaction_data": tx,
                "name_n_num": name_num,
                "case_name": "Flow_Case",
            }
            exp_resp = requests.post(
                f"{BASE}/export/", json=excel_payload, headers=headers, timeout=120
            )
            print(exp_resp.status_code, exp_resp.headers.get("content-disposition"))
        else:
            print("Analyze did not return required sheets; skipping export.")

        # ----------------------------------------------
        # Prometheus metrics snapshot
        # ----------------------------------------------
        try:
            print("\n=== Metrics snapshot ===")
            metrics_txt = requests.get(f"{BASE}/metrics", timeout=10).text.splitlines()[
                :10
            ]
            print("\n".join(metrics_txt))
        except Exception as e:
            print("Metrics fetch error:", e)

        # ----------------------------------------------
        # Send test message to Sentry (if DSN present)
        # ----------------------------------------------
        if os.getenv("SENTRY_DSN"):
            print("\n=== Sentry test message ===")
            try:
                import sentry_sdk

                sentry_sdk.init(dsn=os.getenv("SENTRY_DSN"))
                sentry_sdk.capture_message("Auth flow E2E test message")
                print("Message queued ✔ – check your Sentry dashboard")
            except Exception as e:
                print("Sentry error:", e)
        else:
            print("SENTRY_DSN not set – skipping Sentry test")

    finally:
        stop_server(proc)


if __name__ == "__main__":
    main()
