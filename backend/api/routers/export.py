import logging

import pandas as pd
from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse
from pydantic import BaseModel

from backend.api.api_key import get_api_key
from backend.auth import fastapi_users
from backend.tax_professional.banks.CA_Statement_Analyzer import save_to_excel

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/export", tags=["export"])

current_active_user = fastapi_users.current_user(active=True)


class ExcelDownloadRequest(BaseModel):
    transaction_data: list[dict]
    name_n_num: list[dict]
    case_name: str


async def _export_service(req: ExcelDownloadRequest) -> str:
    """Run heavy save_to_excel in threadpool and return file path."""
    transaction_df = pd.DataFrame(req.transaction_data)
    name_num_df = pd.DataFrame(req.name_n_num)
    file_path = await run_in_threadpool(
        save_to_excel, transaction_df, name_num_df, req.case_name
    )
    return file_path


@router.post("/", summary="Generate Excel workbook from extracted data")
async def export_endpoint(
    req: ExcelDownloadRequest,
    _: current_active_user = Depends(current_active_user),
    __: str = Depends(get_api_key),
):
    try:
        path = await _export_service(req)
        if not path:
            raise HTTPException(
                status_code=500, detail="save_to_excel returned empty path"
            )
        # Construct a more descriptive filename using the case_name
        # Ensure case_name is filesystem-safe (replace spaces, etc.)
        safe_case_name = req.case_name.replace(" ", "_").replace("/", "-")
        download_filename = f"{safe_case_name}_Bank_Statement_Analysis.xlsx"
        logger.info("Serving export file '%s' as '%s'", path, download_filename)
        return FileResponse(path=path, filename=download_filename)
    except Exception as e:
        logger.exception("export failed")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        )
