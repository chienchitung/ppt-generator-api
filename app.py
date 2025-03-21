from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
import json
import tempfile
import os
from pathlib import Path
from pydantic import BaseModel
from typing import Dict, Any
from scripts.generate_ppt import generate_competitive_analysis_ppt

app = FastAPI(
    title="PPT Generator API",
    description="API for generating competitive analysis PowerPoint presentations",
    version="1.0.0"
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
            with open(temp_input_path, "wb") as f:
                f.write(content)
            
            # Generate output path
            output_dir = "generated_ppts"
            os.makedirs(output_dir, exist_ok=True)
            timestamp = Path(input_file.filename).stem
            output_path = os.path.join(output_dir, f"{timestamp}_analysis.pptx")
            
            # Generate PPT
            generate_competitive_analysis_ppt(temp_input_path, output_path)
            
            return GenerationResponse(
                message="PPT generated successfully",
                file_path=output_path
            )
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{filename}")
async def download_ppt(filename: str):
    file_path = os.path.join("generated_ppts", filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=filename)

@app.get("/")
async def root():
    return {"message": "Welcome to PPT Generator API. Use /docs for API documentation."} 