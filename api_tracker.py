"""
Módulo de tracking de uso para reportar sesiones al API de bizland.
"""

import requests
import socket
import os
import getpass
from datetime import datetime, timezone


API_URL = "https://9ut2tgqp5k.execute-api.us-east-1.amazonaws.com/prod/runs"
API_KEY = "86k04Dj2crQ_kAYM2yRLpQSoamcZyyliqyiV25Tp_pI"
PROCESS_ID = "log-analyzer-vl550"
COMPANY = "bizland"


def _now_utc():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _user_id():
    try:
        return f"{getpass.getuser()}@{socket.gethostname()}"
    except Exception:
        return "unknown"


def report_session(start_time, end_time, status="success", records=0, details=None):
    """Envía el reporte de sesión al API."""
    payload = {
        "process_id": PROCESS_ID,
        "company": COMPANY,
        "executed_by": _user_id(),
        "records_processed": records,
        "start_time": start_time,
        "end_time": end_time,
        "status": status,
        "notes": f"Sesión Log Analyzer - {socket.gethostname()}",
    }
    if details:
        payload["details"] = details

    headers = {
        "Content-Type": "application/json",
        "x-api-key": API_KEY,
    }

    try:
        resp = requests.post(API_URL, json=payload, headers=headers, timeout=10)
        resp.raise_for_status()
        return True, resp.status_code
    except Exception as e:
        return False, str(e)
