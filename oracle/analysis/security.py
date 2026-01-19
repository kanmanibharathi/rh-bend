from slowapi import Limiter
from slowapi.util import get_remote_address
from fastapi import UploadFile, HTTPException

limiter = Limiter(key_func=get_remote_address)

def validate_file(file: UploadFile):
    """Validate uploaded file size and extension"""
    if not file.filename.endswith('.csv'):
        raise HTTPException(status_code=400, detail="Only CSV files are allowed.")
    
    # Check file size (max 5MB)
    MAX_SIZE = 5 * 1024 * 1024
    size = 0
    # Read in chunks to check size without loading entirely into memory
    for chunk in file.file:
        size += len(chunk)
        if size > MAX_SIZE:
             raise HTTPException(status_code=400, detail="File too large. Max 5MB.")
    file.file.seek(0) # Reset file pointer
    return True
