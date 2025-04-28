import json
import logging
import time
import traceback
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel, Field

from backend.api.api_key import get_api_key
from backend.auth import fastapi_users

# extraction helper from legacy code
from backend.tax_professional.banks.CA_Statement_Analyzer import (
    start_extraction_add_pdf,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/statements", tags=["statements"])

current_active_user = fastapi_users.current_user(active=True)


# ---------- Pydantic schemas (copied from main for now) ---------- #


class Bounds(BaseModel):
    start: float
    end: float


class ColumnData(BaseModel):
    index: int
    bounds: Bounds
    column_type: Optional[str] = Field(None, alias="type")


class BankStatementRequest(BaseModel):
    bank_names: List[str]
    pdf_paths: List[str]
    passwords: Optional[List[str]] = []
    start_date: List[str]
    end_date: List[str]
    ca_id: str
    whole_transaction_sheet: Optional[List[dict]] = None
    aiyazs_array_of_array: Optional[List[List[ColumnData]]] = None


# ---------- Service wrapper ---------- #


async def _analyze_service(req: BankStatementRequest):
    start_total = time.time()
    logger.info("⚡ analyze_service received request (banks=%s)", req.bank_names)

    def progress_tracker(current: int, total: int, info: str) -> None:
        logger.info("%s (%s/%s)", info, current, total)

    progress_data = {
        "progress_func": progress_tracker,
        "current_progress": 10,
        "total_progress": 100,
    }

    result = await run_in_threadpool(
        start_extraction_add_pdf,
        req.bank_names,
        req.pdf_paths,
        req.passwords or [],
        req.start_date,
        req.end_date,
        req.ca_id,
        progress_data,
        whole_transaction_sheet=req.whole_transaction_sheet,
        aiyazs_array_of_array=req.aiyazs_array_of_array or [],
    )

    logger.info("✅ analyze_service finished in %.2fs", time.time() - start_total)
    return result


def _sanitize(obj):
    import math

    if isinstance(obj, float):
        return obj if math.isfinite(obj) else None
    if isinstance(obj, list):
        return [_sanitize(i) for i in obj]
    if isinstance(obj, dict):
        return {k: _sanitize(v) for k, v in obj.items()}
    return obj


# ---------- Endpoint ---------- #


@router.post("/", summary="Analyze uploaded bank statements")
async def analyze_endpoint(
    req: BankStatementRequest,
    _: current_active_user = Depends(current_active_user),
    __: str = Depends(get_api_key),
):
    try:
        result = await _analyze_service(req)
        sheets_raw = result.get("sheets_in_json")
        try:
            sheets_data = (
                json.loads(sheets_raw) if isinstance(sheets_raw, str) else sheets_raw
            )
        except Exception:
            sheets_data = {}
        clean_data = _sanitize(sheets_data)
        return {
            "data": clean_data,
            "pdf_paths_not_extracted": result.get("pdf_paths_not_extracted"),
            "success_page_number": result.get("success_page_number"),
            "missing_months_list": result.get("missing_months_list"),
        }
    except Exception as e:
        tb = traceback.format_exc()
        logger.error("Analyze crashed: %s", tb)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e) + "\n" + tb
        )
