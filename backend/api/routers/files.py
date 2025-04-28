import logging
import os
import shutil
import uuid
from pathlib import Path
from typing import List

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from backend.api.api_key import get_api_key
from backend.auth import User, fastapi_users
from backend.models.api_key import APIKey

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/files", tags=["Files"])

# Define a directory for temporary uploads INSIDE the container
# Ensure this path is accessible and writable within the container environment.
# /tmp is usually a safe bet.
UPLOAD_DIR = Path("/tmp/user_uploads")


@router.on_event("startup")
async def _startup() -> None:
    # Ensure the upload directory exists when the app starts
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("✔ Upload directory ensured at: %s", UPLOAD_DIR)


@router.post("/upload-pdf")
async def upload_pdf_file(
    files: List[UploadFile] = File(..., description="PDF file(s) to upload"),
    # Protect endpoint with authentication
    user: User = Depends(fastapi_users.current_user(active=True)),
    api_key: APIKey = Depends(get_api_key),
) -> dict:
    """
    Uploads one or more PDF files.

    The file(s) are saved temporarily on the server. The response includes
    a list of the internal paths where the files were saved. These paths
    should be used in the `pdf_paths` field when calling the
    `/statements/` analysis endpoint.

    **Note:** Files are stored temporarily and may be deleted if the server restarts.
    Use the returned paths promptly after upload.
    """
    logger.info("Received %d file(s) for upload from user %s", len(files), user.id)
    saved_paths = []
    try:
        for file in files:
            if not file.filename:
                # Should not happen with standard clients, but handle defensively
                logger.warning("Received file upload with no filename.")
                continue
            if not file.filename.lower().endswith(".pdf"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid file type: '{file.filename}'. Only PDF files are allowed.",
                )

            # Create a unique filename to avoid collisions
            unique_suffix = uuid.uuid4().hex
            # Ensure filename is safe for the filesystem
            safe_filename = f"{unique_suffix}_{os.path.basename(file.filename)}"
            destination_path = UPLOAD_DIR / safe_filename

            logger.debug(
                "Saving uploaded file '%s' to '%s'", file.filename, destination_path
            )
            try:
                with open(destination_path, "wb") as buffer:
                    # shutil.copyfileobj is efficient for large files
                    shutil.copyfileobj(file.file, buffer)
            finally:
                # Ensure the file object is closed even if copy fails
                await file.close()

            saved_paths.append(str(destination_path))
            logger.info(
                "Successfully saved '%s' to %s", file.filename, destination_path
            )

        if not saved_paths:
            raise HTTPException(
                status_code=400, detail="No valid PDF files were uploaded."
            )

        return {"internal_paths": saved_paths}

    except HTTPException:
        # Re-raise HTTP exceptions directly
        raise
    except Exception as e:
        logger.exception("Error saving uploaded file for user %s", user.id, exc_info=e)
        # Clean up any partially saved files if an error occurs? (Optional)
        raise HTTPException(
            status_code=500, detail=f"Internal server error during file upload: {e}"
        )
