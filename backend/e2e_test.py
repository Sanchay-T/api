import json
import os
import signal
import subprocess
import sys
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Helper – load environment variables from backend/.env so SENTRY_DSN is picked
# up by both this helper script *and* the Uvicorn subprocess.  Prefer python‑dotenv
# for robust parsing.
# ---------------------------------------------------------------------------
ENV_FILE = Path(__file__).parent / ".env"
if ENV_FILE.exists():
    load_dotenv(dotenv_path=ENV_FILE)
    # Extra safety: strip whitespace around any "key = value" formats
    with ENV_FILE.open() as f:
        for line in f:
            if line.strip() and not line.strip().startswith("#") and "=" in line:
                k, v = [part.strip() for part in line.split("=", 1)]
                os.environ.setdefault(k, v)


# ---------------------------------------------------------------------------
# Start the FastAPI server in a subprocess (so we can tail its logs).
# ---------------------------------------------------------------------------
cmd = [
    sys.executable,
    "-m",
    "uvicorn",
    "backend.main:app",
    "--port",
    "7500",
]
print("[E2E] Launching API server …")
proc = subprocess.Popen(
    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1
)

# Tail until startup is done (non‑blocking print of first ~40 lines or until app ready)
ready = False
for i, line in enumerate(proc.stdout, 1):
    print("[SERVER]", line.rstrip())
    if "Application startup complete" in line:
        ready = True
        break
    if i > 40:  # safety
        break

if not ready:
    print("[E2E] Server did not start correctly – aborting.")
    proc.terminate()
    sys.exit(1)

# Small grace to ensure port is open
time.sleep(1)

BASE = "http://localhost:7500"

# ---------------------------------------------------------------------------
# 1.  Analyze a sample PDF via /statements/
# ---------------------------------------------------------------------------
url_statements = f"{BASE}/statements/"
print(f"[E2E] POST {url_statements}")

payload = {
    "bank_names": ["HDFC"],
    "pdf_paths": [str(Path(__file__).parent / "pdfs/Manish Hdfc.pdf")],
    "passwords": [""],
    "start_date": ["01-04-2023"],
    "end_date": ["01-04-2024"],
    "ca_id": "E2E_001",
}
resp = requests.post(url_statements, json=payload)
print("[E2E] /statements →", resp.status_code)
print(json.dumps(resp.json(), indent=2)[:500], "…\n")

# ---------------------------------------------------------------------------
# 2.  Prometheus metrics – pull first 10 lines
# ---------------------------------------------------------------------------
metrics_url = f"{BASE}/metrics"
print(f"[E2E] GET {metrics_url}")
metrics_txt = requests.get(metrics_url).text.splitlines()[:10]
print("\n".join(metrics_txt), "\n")

# ---------------------------------------------------------------------------
# 3.  Sentry – emit a test message directly (only if DSN present)
# ---------------------------------------------------------------------------
if os.getenv("SENTRY_DSN"):
    print("[E2E] Sending test message to Sentry …")
    try:
        import sentry_sdk

        sentry_sdk.init(dsn=os.getenv("SENTRY_DSN"))
        sentry_sdk.capture_message("E2E Test message from backend/e2e_test.py")
        print("[E2E] Sentry message queued ✔")
    except Exception as e:
        print("[E2E] Sentry sdk error:", e)
else:
    print("[E2E] SENTRY_DSN not set – skipping Sentry test")

# ---------------------------------------------------------------------------
# Cleanup – terminate server subprocess
# ---------------------------------------------------------------------------
print("[E2E] Shutting down server …")
proc.send_signal(signal.SIGINT)
proc.wait(timeout=10)
print("[E2E] Done.")
