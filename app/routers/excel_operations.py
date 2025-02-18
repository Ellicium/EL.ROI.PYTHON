import logging
import os

from fastapi.responses import FileResponse
from fastapi import APIRouter, Depends
from fastapi.logger import logger
from fastapi.responses import JSONResponse
from fastapi import FastAPI, File, UploadFile
from io import BytesIO


# from ..config.auth_middleware import verify_and_decode_token
from ..config.logger_config import get_logger
from ..schemas.data_schema import read_excel_file
from ..services.data_service import modify_excel_fields
# upload_bom_function,upload_po_function,upload_cost_breakdown_function,upload_material_breakdown_function

router = APIRouter()
logger = get_logger()


@router.post("/data/excel_transformations", tags=["data"])
async def excel_operation(file: UploadFile = File(...)
):
    try:

           # Read the Excel file
        contents = await file.read()

        # Convert binary to a file-like object
        excel_file = BytesIO(contents)
        
        EXCEL_FILE_PATH=modify_excel_fields(excel_file)

        false_response={'response_code':400,'Message':'Excel Not Updated Please Try Again'}
                
        if EXCEL_FILE_PATH==0:
            return JSONResponse(false_response,status_code=false_response['response_code'])


        return FileResponse(
            path=EXCEL_FILE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="downloaded_excel.xlsx"
        )
     
    except Exception as e:
        logger.error("print_input failed", exc_info=e)
        return JSONResponse(f"error {e}", status_code=500)


