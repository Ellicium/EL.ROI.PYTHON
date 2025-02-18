from fastapi import FastAPI, Form, Request, Response, Depends
# from .routers import genAI_router, product_routes
# from .config.auth_middleware import verify_and_decode_token
from .routers import excel_operations
# default_routes, data_routes,configure_routes,tagcommodity_routes,shouldcost_routes,currency_update_routes
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    debug=True,
    title="ROI Application",
    description="This application automate ROI for RPA",
    version="0.0.1",
)

# Enable CORS for all routes
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Replace with the actual origin of your frontend
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


app.include_router(excel_operations.router)
