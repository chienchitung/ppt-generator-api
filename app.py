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
ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "*").split(",")
if "*" in ALLOWED_ORIGINS:
    ALLOWED_ORIGINS = ["*"]
logger.info(f"Configured ALLOWED_ORIGINS: {ALLOWED_ORIGINS}")

# Configure storage
STORAGE_DIR = os.getenv("STORAGE_DIR", "generated_ppts")
os.makedirs(STORAGE_DIR, exist_ok=True)
logger.info(f"Storage directory configured at: {STORAGE_DIR}")

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
    expose_headers=["*"],
    max_age=86400,
)

class GenerationResponse(BaseModel):
    message: str
    file_path: str

@app.post("/generate-ppt/", response_model=GenerationResponse)
async def generate_ppt(input_file: UploadFile = File(...)):
    logger.info(f"Received file: {input_file.filename}")
    try:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded JSON file
            temp_input_path = os.path.join(temp_dir, "input.json")
            content = await input_file.read()
            
            # Validate JSON content
            try:
                json_content = json.loads(content)
                logger.info("JSON content validated successfully")
                # Log the structure of the JSON content
                logger.info(f"JSON content structure: {json.dumps({k: type(v).__name__ for k, v in json_content.items()})}")
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON content received: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid JSON file: {str(e)}")
            
            # Write content to temporary file
            try:
                with open(temp_input_path, "wb") as f:
                    f.write(content)
                logger.info(f"Content written to temporary file: {temp_input_path}")
            except Exception as e:
                logger.error(f"Error writing temporary file: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Error writing temporary file: {str(e)}")
            
            # Generate output filename
            output_filename = f"data_analysis.pptx"
            output_path = os.path.join(STORAGE_DIR, output_filename)
            logger.info(f"Output path set to: {output_path}")
            
            # Generate PPT
            try:
                # Ensure the storage directory exists
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                # Generate the PPT
                generate_competitive_analysis_ppt(temp_input_path, output_path)
                logger.info(f"Generated PPT at: {output_path}")
            except Exception as e:
                logger.error(f"Error generating PPT: {str(e)}", exc_info=True)
                if hasattr(e, '__traceback__'):
                    import traceback
                    logger.error(f"Full traceback: {''.join(traceback.format_tb(e.__traceback__))}")
                raise HTTPException(status_code=500, detail=f"Error generating PPT: {str(e)}")
            
            # Verify file exists and has content
            if not os.path.exists(output_path):
                logger.error(f"Generated file not found at: {output_path}")
                raise HTTPException(status_code=500, detail="Generated file not found")
            
            if os.path.getsize(output_path) == 0:
                logger.error(f"Generated file is empty: {output_path}")
                raise HTTPException(status_code=500, detail="Generated file is empty")
            
            response_data = GenerationResponse(
                message="PPT generated successfully",
                file_path=output_filename
            )
            logger.info(f"Returning response: {response_data}")
            return response_data
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error in generate_ppt: {str(e)}", exc_info=True)
        if hasattr(e, '__traceback__'):
            import traceback
            logger.error(f"Full traceback: {''.join(traceback.format_tb(e.__traceback__))}")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")

@app.get("/download/{filename}")
async def download_ppt(filename: str):
    logger.info(f"Download requested for file: {filename}")
    try:
        file_path = os.path.join(STORAGE_DIR, filename)
        logger.info(f"Full file path: {file_path}")
        
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            raise HTTPException(status_code=404, detail="File not found")
            
        logger.info("Sending file response")
        response = FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=filename
        )
        # 添加 CORS headers
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "*"
        return response
    except Exception as e:
        logger.error(f"Error during file download: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error downloading file: {str(e)}")

@app.get("/")
async def root():
    return {"message": "Welcome to PPT Generator API. Use /docs for API documentation."} 