from typing import List, Optional

from pydantic import BaseModel

class read_excel_file(BaseModel):
    file_path: Optional[str] = None

class user_input(BaseModel):
    usecase: Optional[str] = None
    frequency: Optional[str] = None
    manual_mht_aht:Optional[float] = None
    avg_transaction_annual:Optional[float] = None
    number_of_vms:Optional[float] = None
    is_ocr: Optional[str] = None
    already_in_prod: Optional[str] = None
    complexivity: Optional[str] = None
    monthly_vm_cost:Optional[float] = None
    monthly_fte_cost:Optional[float] = None
    monthly_seat_cost:Optional[float] = None
    monthly_fte_other_cost:Optional[float] = None
    monthly_runner_licence_cost:Optional[float] = None
    document_automation_cost:Optional[float] = None
    monthly_creater_licence_cost:Optional[float] = None
    support_cost_per_resource:Optional[float] = None
    