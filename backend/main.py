import logging
import os
from typing import List, Optional

import pandas as pd
import uvicorn
from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

# import matplotlib
# matplotlib.use('Agg')
# from findaddy.exceptions import ExtractionError
from backend.utils import get_saved_pdf_dir

TEMP_SAVED_PDF_DIR = get_saved_pdf_dir()

import uuid

from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from prometheus_fastapi_instrumentator import Instrumentator
from pydantic import Field

from backend import auth
from backend.api import api_key as api_key_router
from backend.api.routers import analyze as analyze_router
from backend.api.routers import export as export_router
from backend.api.routers import files as files_router
from backend.core.config import settings
from backend.core.logging import init_logging
from backend.core.observability import init_sentry
from backend.core.request_id import request_id_ctx_var

# If you have other custom imports:
from backend.tax_professional.banks.CA_Statement_Analyzer import (
    individual_summary,
    refresh_category_all_sheets,
    save_to_excel,
)

init_logging()

# Initialise Sentry (no‑op if DSN missing)
init_sentry()

logging.getLogger(__name__).info("✔ main module logging initialised")

app = FastAPI(title="Bank Statement Analyzer API")
app.include_router(auth.router)
app.include_router(api_key_router.router)
app.include_router(analyze_router.router)
app.include_router(export_router.router)
app.include_router(files_router.router)

# expose /metrics
Instrumentator().instrument(app).expose(app, endpoint="/metrics")

logger = logging.getLogger(__name__)
logger.info(f"Temp directory python : {TEMP_SAVED_PDF_DIR}")


@app.middleware("http")
async def add_request_id(request: Request, call_next):
    req_id = str(uuid.uuid4())
    request_id_ctx_var.set(req_id)
    response = await call_next(request)
    response.headers["X-Request-ID"] = req_id
    return response


class Transaction(BaseModel):
    id: int
    statementId: Optional[str] = None
    date: Optional[str] = None
    description: Optional[str] = None
    amount: Optional[float] = None
    category: Optional[str] = None
    transaction_type: Optional[str] = None
    balance: Optional[float] = None
    bank: Optional[str] = None
    entity: Optional[str] = None


class Bounds(BaseModel):
    start: float
    end: float


class ColumnData(BaseModel):
    index: int
    bounds: Bounds
    column_type: Optional[str] = Field(None, alias="type")


class EditPdfRequest(BaseModel):
    bank_names: List[str]
    pdf_paths: List[str]
    passwords: Optional[List[str]] = []  # Optional field, defaults to empty list
    start_dates: List[str]
    end_dates: List[str]
    aiyazs_array_of_array: List[List[ColumnData]]
    whole_transaction_sheet: Optional[List[Transaction]] = None
    ca_id: str


class BankStatementRequest(BaseModel):
    bank_names: List[str]
    pdf_paths: List[str]
    passwords: Optional[List[str]] = []  # Optional field, defaults to empty list
    start_date: List[str]
    end_date: List[str]
    ca_id: str
    whole_transaction_sheet: Optional[List[dict]] = None
    aiyazs_array_of_array: Optional[List[List[ColumnData]]] = None


class EditCategoryRequest(BaseModel):
    transaction_data: List[dict]
    new_categories: List[dict]
    eod_data: List[dict]


class ExcelDownloadRequest(BaseModel):
    transaction_data: List[dict]
    name_n_num: List[dict]
    case_name: str


class DummyRequest(BaseModel):
    data: str


class InvididualSummaryRequest(BaseModel):
    transactions_data: List[dict]


@app.get("/", response_class=HTMLResponse)
async def root():
    return "<h1>Yes, I am alive!</h1>"


@app.post("/")
async def root(data: str = Body(...)):
    print("Received data in root : ", data)
    return {"message": "Bank Statement Analyzer API"}


# @app.post("/analyze-statements/")
# async def analyze_bank_statements(request: BankStatementRequest):
#     try:
#         start_total = time.time()
#
#         logger.info(f"Received request with banks: {request.bank_names}")
#         print("Start Date : ", request.start_date)
#         print("End Date : ", request.end_date)
#         print("PDF Paths : ", request.pdf_paths)
#
#         # Create a progress tracking function
#         def progress_tracker(current: int, total: int, info: str) -> None:
#             logger.info(f"{info} ({current}/{total})")
#
#         progress_data = {
#             "progress_func": progress_tracker,
#             "current_progress": 10,
#             "total_progress": 100,
#         }
#
#         # Validate passwords length if provided
#         if request.passwords and len(request.passwords) != len(request.pdf_paths):
#             raise HTTPException(
#                 status_code=400,
#                 detail=(
#                     f"Number of passwords ({len(request.passwords)}) "
#                     f"must match number of PDFs ({len(request.pdf_paths)})"
#                 ),
#             )
#
#         logger.info("Initializing CABankStatement")
#         # Pass empty list if no passwords
#
#         bank_names = request.bank_names
#         pdf_paths = request.pdf_paths
#         passwords = request.passwords if request.passwords else []
#         start_date = request.start_date if request.start_date else []
#         end_date = request.end_date if request.end_date else []
#         CA_ID = request.ca_id
#         progress_data = progress_data
#
#         ner_results = {"Name": [], "Acc Number": []}
#
#         # Process PDFs with NER
#         start_ner = time.time()
#         person_count = 0
#         for pdf in pdf_paths:
#             person_count += 1
#             # result = pdf_to_name_and_accno(pdf)
#             fetched_name = None
#             fetched_acc_num = None
#
#             name_entities = extract_entities(pdf)
#             acc_number_ifsc = extract_accno_ifsc(pdf)
#
#             print("name_entities:- ", name_entities)
#
#             fetched_acc_num = acc_number_ifsc["acc"]
#
#             if name_entities:
#                 for entity in name_entities:
#                     if fetched_name == None:
#                         fetched_name = entity
#
#             if fetched_name:
#                 ner_results["Name"].append(fetched_name)
#             else:
#                 ner_results["Name"].append(f"Statement {person_count}")
#
#             if fetched_acc_num:
#                 ner_results["Acc Number"].append(fetched_acc_num)
#             else:
#                 ner_results["Acc Number"].append("XXXXXXXXXXX")
#         print("Ner results", ner_results)
#         end_ner = time.time()
#         print("Time taken to process NER", end_ner - start_ner)
#
#         start_extraction = time.time()
#
#         logger.info("Starting extraction")
#         whole_transaction_sheet = request.whole_transaction_sheet or None
#         temp_aiyaz_array_of_array = []
#         if request.aiyazs_array_of_array:
#             for statement in request.aiyazs_array_of_array:
#                 temp_aiyaz_array = []
#                 for col in statement:
#                     temp_aiyaz_array.append(col.model_dump())
#                 temp_aiyaz_array_of_array.append(temp_aiyaz_array)
#
#         if whole_transaction_sheet is not None:
#             whole_transaction_sheet = pd.DataFrame(whole_transaction_sheet)
#             print("whole_transaction_sheet", whole_transaction_sheet.head())
#             whole_transaction_sheet["Value Date"] = pd.to_datetime(
#                 whole_transaction_sheet["Value Date"], format="%d-%m-%Y"
#             )
#
#         result = start_extraction_add_pdf(
#             bank_names,
#             pdf_paths,
#             passwords,
#             start_date,
#             end_date,
#             CA_ID,
#             progress_data,
#             whole_transaction_sheet=whole_transaction_sheet,
#             aiyazs_array_of_array=temp_aiyaz_array_of_array,
#         )
#         end_extraction = time.time()
#         end_total = time.time()
#         total_time = end_total - start_total
#
#         print(
#             "Time taken for extraction:", end_extraction - start_extraction, "seconds"
#         )
#
#         print("RESULT GENERATED")
#         logger.info("Extraction completed successfully")
#         logger.info("Result = ", result)
#         return {
#             "status": "success",
#             "message": "Bank statements analyzed successfully",
#             "data": result["sheets_in_json"],
#             "pdf_paths_not_extracted": result["pdf_paths_not_extracted"],
#             "ner_results": ner_results,
#             "success_page_number": result["success_page_number"],
#             "missing_months_list": result["missing_months_list"],
#             "processing_times": {
#                 "ner_processing": end_ner - start_ner,
#                 "extraction": end_extraction - start_extraction,
#                 "total": total_time,
#             },
#         }
#
#     except Exception as e:
#         print(e)
#         logger.error(f"Error processing bank statements: {str(e)}")
#         return {
#             "status": "failed",
#             "message": str(e),
#         }
#         # raise HTTPException(
#         #     status_code=500, detail=f"Error processing bank statements: {str(e)}"
#         # )


# @app.post("/column-rectify-add-pdf/")
# async def column_rectify_add_pdf(request:EditPdfRequest):
#     print("Received request data:", request)
#     try:

#         # # Create a progress tracking function
#         def progress_tracker(current: int, total: int, info: str) -> None:
#             logger.info(f"{info} ({current}/{total})")

#         progress_data = {
#         "progress_func": progress_tracker,
#         "current_progress": 10,
#         "total_progress": 100,
#         }

#         # Validate passwords length if provided
#         if request.passwords and len(request.passwords) != len(request.pdf_paths):
#             raise HTTPException(
#                 status_code=400,
#                 detail=(
#                     f"Number of passwords ({len(request.passwords)}) "
#                     f"must match number of PDFs ({len(request.pdf_paths)})"
#                 ),
#             )

#         temp_aiyaz_array_of_array = []
#         for statement in request.aiyazs_array_of_array:
#             temp_aiyaz_array = []
#             for col in statement:
#                 temp_aiyaz_array.append(col.model_dump())
#             temp_aiyaz_array_of_array.append(temp_aiyaz_array)


#         bank_names = request.bank_names
#         pdf_paths = request.pdf_paths
#         passwords =  request.passwords if request.passwords else []
#         start_date = request.start_dates if request.start_dates else []
#         end_date = request.end_dates if request.end_dates else []
#         CA_ID = request.ca_id
#         progress_data = progress_data
#         aiyazs_array_of_array = temp_aiyaz_array_of_array
#         whole_transaction_sheet = request.whole_transaction_sheet


#         ner_results = {
#                 "Name": [],
#                 "Acc Number": []
#             }

#         # Process PDFs with NER
#         start_ner = time.time()
#         person_count = 0
#         for pdf in pdf_paths:
#             person_count+=1
#             # result = pdf_to_name_and_accno(pdf)
#             fetched_name = None
#             fetched_acc_num = None

#             name_entities = extract_entities(pdf)
#             acc_number_ifsc = extract_accno_ifsc(pdf)

#             print("name_entities:- ",name_entities)

#             fetched_acc_num=acc_number_ifsc["acc"]

#             if name_entities:
#                 for entity in name_entities:
#                     if fetched_name==None:
#                         fetched_name=entity

#             if fetched_name:
#                 ner_results["Name"].append(fetched_name)
#             else:
#                 ner_results["Name"].append(f"Statement {person_count}")

#             if fetched_acc_num:
#                 ner_results["Acc Number"].append(fetched_acc_num)
#             else:
#                 ner_results["Acc Number"].append("XXXXXXXXXXX")
#         print("Ner results", ner_results)
#         end_ner = time.time()
#         print("Time taken to process NER", end_ner-start_ner)


#         logger.info("Starting extraction")
#         result = start_extraction_edit_pdf(bank_names=bank_names,pdf_paths= pdf_paths,passwords= passwords,start_dates= start_date,end_dates= end_date,CA_ID= CA_ID, progress_data=progress_data,aiyazs_array_of_array=aiyazs_array_of_array,whole_transaction_sheet=whole_transaction_sheet)

#         print("RESULT GENERATED")
#         logger.info("Result = ", result["sheets_in_json"])
#         logger.info("Result pdf_paths_not_extracted= ", result["pdf_paths_not_extracted"])
#         logger.info("Extraction completed successfully")
#         return {
#             "status": "success",
#             "message": "Bank statements analyzed successfully",
#             "data": result["sheets_in_json"],
#             "pdf_paths_not_extracted": result["pdf_paths_not_extracted"],
#             "ner_results": ner_results,
#         }

#     except Exception as e:

#         print(e)
#         logger.error(f"Error processing bank statements: {str(e)}")
#         raise HTTPException(
#             status_code=500, detail=f"Error processing bank statements: {str(e)}"
#         )


@app.post("/refresh/")
async def refresh(request: BankStatementRequest):
    pass


@app.post("/add-pdf/")
async def add_pdf(request: BankStatementRequest):
    pass


@app.get("/health")
async def health_check():
    return {"status": "healthy"}


@app.post("/edit-category/")
async def edit_category(request: EditCategoryRequest):
    try:
        transaction_data = request.transaction_data
        new_categories = request.new_categories
        eod_data = request.eod_data
        print("New Categories : ", new_categories)
        print("Transaction Data : ", transaction_data)
        print("EOD Data : ", eod_data)
        logger.info(f"Received request with new categories: {new_categories}")
        logger.info(f"Received request with transaction data: {transaction_data[0]}")
        logger.info(f"Received request with eod data: {eod_data}")

        # convert transaction_data to df
        transaction_df = pd.DataFrame(transaction_data)
        print("Transactions : ", transaction_df.head())
        transaction_df["Value Date"] = pd.to_datetime(
            transaction_df["Value Date"], format="%d-%m-%Y"
        )
        eod_df = pd.DataFrame(eod_data)
        print("Transactions : ", transaction_df.head())
        # print(eod_df.head())
        print("new categories : ", new_categories)

        data = refresh_category_all_sheets(transaction_df, eod_df, new_categories)
        print(data)

        return data

    except Exception as e:
        logger.error(f"Error processing bank statements: {str(e)}")
        raise HTTPException(
            status_code=500, detail=f"Error processing bank statements: {str(e)}"
        )


@app.post("/excel-download/")
async def excel_download(request: ExcelDownloadRequest):
    try:
        transaction_data = request.transaction_data
        case_name = request.case_name
        name_n_num_data = request.name_n_num
        logger.info(f"Received request with transaction data: {transaction_data[0]}")
        logger.info(f"Received request with case name: {case_name}")

        # convert transaction_data to df
        transaction_df = pd.DataFrame(transaction_data)
        name_n_num_df = pd.DataFrame(name_n_num_data)

        print("Transactions : \n", transaction_df.head())
        print("Name and Number : \n", name_n_num_df.head())

        file_path = save_to_excel(transaction_df, name_n_num_df, case_name)
        print("Python data : ", file_path)

        if not os.path.exists(file_path):
            raise HTTPException(
                status_code=404, detail="Something went wrong while generating the file"
            )

        return file_path
    except Exception as e:
        logger.error(f"Error processing bank statements: {str(e)}")
        raise HTTPException(status_code=500, detail=f"{str(e)}")


@app.post("/individual-summary/")
async def individual_summary_api(request: InvididualSummaryRequest):
    try:
        logger.info(f"Received request with data: {request.transactions_data}")

        transaction_df = pd.DataFrame(request.transactions_data)
        transaction_df["Value Date"] = pd.to_datetime(
            transaction_df["Value Date"], format="%d-%m-%Y"
        )
        print(transaction_df.head(10))
        data = individual_summary(transaction_df)
        print(data)

        return data

    except Exception as e:
        print(e)
        logger.error(f"Error processing bank statements: {str(e)}")
        raise HTTPException(
            status_code=500, detail=f"Error processing bank statements: {str(e)}"
        )


## NOTE: Database schema is now managed by Alembic migrations.
## Running metadata.create_all() on startup is removed to prevent
## race conditions with multiple workers and ensure Alembic is the
## single source of truth for the schema.
#
# @app.on_event("startup")
# async def _startup() -> None:
#     logger.info("⚙ Initialising database schema via create_all (DEV ONLY) …")
#     async with engine.begin() as conn:
#         await conn.run_sync(UserBase.metadata.create_all)
#     logger.info("✔ database schema synced via create_all (tables=%s)", list(UserBase.metadata.tables))


if __name__ == "__main__":
    # Optionally use environment variables for host/port. Falls back to settings if none provided.
    host = os.getenv("API_HOST", settings.api_host)
    port = int(os.getenv("API_PORT", settings.api_port))

    # uds_path = "/tmp/bank_statement_analyzer.sock"

    # Clean up any old socket
    # if os.path.exists(uds_path):
    # os.remove(uds_path)

    # Start the FastAPI server on the Unix socket
    # uvicorn.run("main:app", uds=uds_path, log_level="info", reload=False)

    # IMPORTANT: reload=False for production usage
    # import time
    # time.sleep(8)
    uvicorn.run(app, host=host, port=port, reload=False)


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    print("Validation Error:", exc.errors())
    return JSONResponse(
        status_code=422,
        content={"detail": exc.errors()},
    )
