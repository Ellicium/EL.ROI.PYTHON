import logging
import os

from fastapi.responses import FileResponse
from fastapi import APIRouter, Depends
from fastapi.logger import logger
from fastapi.responses import JSONResponse
from fastapi import FastAPI, File, UploadFile, Form, BackgroundTasks

from io import BytesIO
from typing import Optional
import requests

# from ..config.auth_middleware import verify_and_decode_token
from ..config.logger_config import get_logger
from ..schemas.data_schema import read_excel_file,user_input
from ..services.data_service import modify_excel_fields,update_user_values, modify_excel_fields_v2
# upload_bom_function,upload_po_function,upload_cost_breakdown_function,upload_material_breakdown_function

router = APIRouter()
# # logger = get_logger()

@router.post("/data/excel_transformations", tags=["data"])
async def excel_operation(file: UploadFile = File(...)
):
    try:

           # Read the Excel file
        contents = await file.read()

        # Convert binary to a file-like object
        excel_file = BytesIO(contents)
        

        EXCEL_FILE_PATH, return_dictt = modify_excel_fields(excel_file)

        if EXCEL_FILE_PATH == 0:
            return JSONResponse(
                content={"response_code": 400, "message": "Excel Not Updated. Please Try Again"},
                status_code=400
            )

        # Attach extra data to response
        response = FileResponse(
            path=EXCEL_FILE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="downloaded_excel.xlsx"
        )

        # Add extra data in headers
        for key, value in return_dictt.items():
            response.headers[key] = str(value)  # Convert values to string for headers

        return response

        # EXCEL_FILE_PATH=modify_excel_fields(excel_file)

        # false_response={'response_code':400,'Message':'Excel Not Updated Please Try Again'}
                
        # if EXCEL_FILE_PATH==0:
        #     return JSONResponse(false_response,status_code=false_response['response_code'])


        # return FileResponse(
        #     path=EXCEL_FILE_PATH,
        #     media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #     filename="downloaded_excel.xlsx"
        # )
     
    except Exception as e:
        logger.error("print_input failed", exc_info=e)
        return JSONResponse(f"error {e}", status_code=500)







@router.post("/data/user_input", tags=["data"])
async def excel_operation(apipostschema:user_input
):
    try:
        
        # EXCEL_FILE_PATH, retun_dictt=update_user_values(apipostschema)

        # false_response={'response_code':400,'Message':'Excel Not Updated Please Try Again'}
                
        # if EXCEL_FILE_PATH==0:
        #     return JSONResponse(false_response,status_code=false_response['response_code'])


        # return FileResponse(
        #     path=EXCEL_FILE_PATH,
        #     media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #     filename="downloaded_excel.xlsx"
        # )
     
        EXCEL_FILE_PATH, return_dictt = update_user_values(apipostschema)

        if EXCEL_FILE_PATH == 0:
            return JSONResponse(
                content={"response_code": 400, "message": "Excel Not Updated. Please Try Again"},
                status_code=400
            )

        # Attach extra data to response
        response = FileResponse(
            path=EXCEL_FILE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="downloaded_excel.xlsx"
        )

        # Add extra data in headers
        for key, value in return_dictt.items():
            response.headers[key] = str(value)  # Convert values to string for headers

        return response

        #    # Read the Excel file
        # contents = await file.read()

        # # Convert binary to a file-like object
        # excel_file = BytesIO(contents)
        
        # EXCEL_FILE_PATH=modify_excel_fields(excel_file)

        # false_response={'response_code':400,'Message':'Excel Not Updated Please Try Again'}
                
        # if EXCEL_FILE_PATH==0:
        #     return JSONResponse(false_response,status_code=false_response['response_code'])


        # return FileResponse(
        #     path=EXCEL_FILE_PATH,
        #     media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #     filename="downloaded_excel.xlsx"
        # )
     
    except Exception as e:
        logger.error("print_input failed", exc_info=e)
        return JSONResponse(f"error {e}", status_code=500)





@router.post("/data/excel_transformations_v2", tags=["data"])
async def excel_operation_v2(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    monthly_vm_cost: Optional[float] = Form(None),
    monthly_fte_cost: Optional[float] = Form(None),
    monthly_seat_cost: Optional[float] = Form(None),
    monthly_fte_other_cost: Optional[float] = Form(None),
    monthly_runner_licence_cost: Optional[float] = Form(None),
    document_automation_cost: Optional[float] = Form(None),
    monthly_creater_licence_cost: Optional[float] = Form(None),
    support_cost_per_resource: Optional[float] = Form(None),
    monthly_developer_cost_per_resource: Optional[float] = Form(None),
):
    try:
           # Read the Excel file
        contents = await file.read()

        # Convert binary to a file-like object
        excel_file = BytesIO(contents)
        

        EXCEL_FILE_PATH, return_dictt = modify_excel_fields_v2(excel_file,monthly_vm_cost,monthly_fte_cost,monthly_seat_cost,monthly_fte_other_cost,monthly_runner_licence_cost,document_automation_cost,monthly_creater_licence_cost,support_cost_per_resource,monthly_developer_cost_per_resource)

        if EXCEL_FILE_PATH == 0:
            return JSONResponse(
                content={"response_code": 400, "message": "Excel Not Updated. Please Try Again"},
                status_code=400
            )

        # Attach extra data to response
        response = FileResponse(
            path=EXCEL_FILE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="downloaded_excel.xlsx"
        )

        # Add extra data in headers
        for key, value in return_dictt.items():
            response.headers[key] = str(value)  # Convert values to string for headers

        
        background_tasks.add_task(os.remove, os.path.join(os.getcwd(), EXCEL_FILE_PATH))

        print(os.path.join(os.getcwd(), EXCEL_FILE_PATH))

        return response

        # EXCEL_FILE_PATH=modify_excel_fields(excel_file)

        # false_response={'response_code':400,'Message':'Excel Not Updated Please Try Again'}
                
        # if EXCEL_FILE_PATH==0:
        #     return JSONResponse(false_response,status_code=false_response['response_code'])


        # return FileResponse(
        #     path=EXCEL_FILE_PATH,
        #     media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #     filename="downloaded_excel.xlsx"
        # )
     
    except Exception as e:
        logger.error("print_input failed", exc_info=e)
        return JSONResponse(f"error {e}", status_code=500)






