import asyncio
import signal
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path

import httpx

HOST = "127.0.0.1"
PORT = 7500
BASE = f"http://{HOST}:{PORT}"
SAMPLE_PDF = Path(__file__).resolve().parent / "pdfs" / "Manish Hdfc.pdf"


def run_server() -> subprocess.Popen:
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
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    def _reader():
        for line in proc.stdout:
            print("[SERVER]", line.rstrip())

    threading.Thread(target=_reader, daemon=True).start()

    # wait until /health responds
    ready = False
    for _ in range(120):
        if proc.poll() is not None:
            raise RuntimeError("Server crashed on startup")
        try:
            resp = httpx.get(f"{BASE}/health", timeout=1.0)
            if resp.status_code == 200:
                ready = True
                break
        except Exception:
            pass
        time.sleep(0.25)
    if not ready:
        raise RuntimeError("Server failed to become ready in time")
    return proc


def stop_server(proc: subprocess.Popen):
    proc.send_signal(signal.SIGINT)
    proc.wait(timeout=5)


async def user_flow(idx: int):
    async with httpx.AsyncClient(timeout=120) as client:
        email = f"load{idx}_{uuid.uuid4().hex[:6]}@ex.com"
        pwd = "Pass123!"
        # register (ignore if exists)
        await client.post(
            f"{BASE}/auth/register", json={"email": email, "password": pwd}
        )
        # login
        r = await client.post(
            f"{BASE}/auth/jwt/login",
            data={"username": email, "password": pwd},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        token = r.json()["access_token"]
        headers = {"Authorization": f"Bearer {token}"}
        # api key
        r = await client.post(f"{BASE}/users/api-key", headers=headers)
        api_key = r.json()["api_key"]
        headers["X-API-Key"] = api_key

        # analyze
        analyze_payload = {
            "bank_names": ["HDFC"],
            "pdf_paths": [str(SAMPLE_PDF.resolve())],
            "passwords": [""],
            "start_date": ["01-04-2023"],
            "end_date": ["01-04-2024"],
            "ca_id": f"CASE_{idx}",
        }
        t0 = time.perf_counter()
        r = await client.post(
            f"{BASE}/statements/", json=analyze_payload, headers=headers
        )
        t_analyze = time.perf_counter() - t0
        if r.status_code != 200:
            print(f"user-{idx} analyze failed {r.status_code}: {r.text[:200]}")
            return
        sheets = r.json()["data"]
        # export
        excel_payload = {
            "transaction_data": sheets.get("Transactions", []),
            "name_n_num": sheets.get("Name Acc No", []),
            "case_name": f"Load_Case_{idx}",
        }
        t0 = time.perf_counter()
        r = await client.post(f"{BASE}/export/", json=excel_payload, headers=headers)
        t_export = time.perf_counter() - t0
        status = "✅" if r.status_code == 200 else "❌"
        print(f"user-{idx} analyze={t_analyze:.1f}s export={t_export:.1f}s {status}")


async def main():
    proc = run_server()
    start = time.perf_counter()
    try:
        await asyncio.gather(*(user_flow(i) for i in range(1, 4)))
        print("TOTAL wall-time %.1fs" % (time.perf_counter() - start))
    finally:
        stop_server(proc)


if __name__ == "__main__":
    asyncio.run(main())
