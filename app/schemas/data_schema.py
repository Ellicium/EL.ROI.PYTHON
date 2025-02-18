from typing import List, Optional

from pydantic import BaseModel

class read_excel_file(BaseModel):
    file_path: Optional[str] = None
