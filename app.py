from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import json
import tempfile
import os
import logging
from pathlib import Path
from pydantic import BaseModel
from typing import Dict, Any
from scripts.generate_ppt import generate_competitive_analysis_ppt

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Get allowed origins from environment variable
# 如果環境變數未設置，默認允許 localhost:3000
ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "http://localhost:3000").split(",")

# Configure storage
STORAGE_DIR = os.getenv("STORAGE_DIR", "generated_ppts")
os.makedirs(STORAGE_DIR, exist_ok=True)

app = FastAPI(
    title="PPT Generator API",
    description="API for generating competitive analysis PowerPoint presentations",
    version="1.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class GenerationResponse(BaseModel):
    message: str
    file_path: str

@app.post("/generate-ppt/", response_model=GenerationResponse)
async def generate_ppt(input_file: UploadFile = File(...)):
    try:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded JSON file
            temp_input_path = os.path.join(temp_dir, "input.json")
            content = await input_file.read()
            
            # Validate JSON content
            try:
                json_content = json.loads(content)
            except json.JSONDecodeError:
                raise HTTPException(status_code=400, detail="Invalid JSON file")
            
            # Write content to temporary file
            try:
                with open(temp_input_path, "wb") as f:
                    f.write(content)
            except Exception as e:
                logger.error(f"Error writing temporary file: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Error writing temporary file: {str(e)}")
            
            # Generate output filename
            timestamp = Path(input_file.filename).stem
            output_filename = f"{timestamp}_analysis.pptx"
            output_path = os.path.join(STORAGE_DIR, output_filename)
            
            # Generate PPT
            try:
                generate_competitive_analysis_ppt(temp_input_path, output_path)
                logger.info(f"Generated PPT at: {output_path}")
            except Exception as e:
                logger.error(f"Error generating PPT: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Error generating PPT: {str(e)}")
            
            # Verify file exists
            if not os.path.exists(output_path):
                logger.error(f"Generated file not found at: {output_path}")
                raise HTTPException(status_code=500, detail="Generated file not found")
            
            # Return only the filename in the response
            return GenerationResponse(
                message="PPT generated successfully",
                file_path=output_filename  # Only return the filename
            )
            
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{filename}")
async def download_ppt(filename: str):
    try:
        file_path = os.path.join(STORAGE_DIR, filename)
        logger.info(f"Attempting to download file: {file_path}")
        
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            raise HTTPException(status_code=404, detail="File not found")
            
        return FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=filename,
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"Error during file download: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error downloading file: {str(e)}")

@app.get("/")
async def root():
    return {"message": "Welcome to PPT Generator API. Use /docs for API documentation."} 