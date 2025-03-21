import pytest
from fastapi.testclient import TestClient
from app import app
import json
import os

client = TestClient(app)

def test_root():
    response = client.get("/")
    assert response.status_code == 200
    assert response.json() == {"message": "Welcome to PPT Generator API. Use /docs for API documentation."}

def test_generate_ppt_invalid_json():
    # Test with invalid JSON
    response = client.post(
        "/generate-ppt/",
        files={"input_file": ("test.json", "invalid json content", "application/json")}
    )
    assert response.status_code == 400
    assert "Invalid JSON file" in response.json()["detail"]

def test_generate_ppt_valid_json():
    # Create a valid test JSON
    test_data = {
        "title": "Test Analysis",
        "date": "2024-03-21",
        "apps": [
            {
                "name": "Test App",
                "ratings": {
                    "ios": 4.5,
                    "android": 4.3
                },
                "reviews": {
                    "stats": {
                        "positive": 85,
                        "negative": 15
                    },
                    "analysis": {
                        "advantages": ["Good UI"],
                        "improvements": ["Slow loading"],
                        "summary": "Overall good"
                    }
                },
                "features": {
                    "core": ["Feature 1"],
                    "advantages": ["Advantage 1"],
                    "improvements": ["Improvement 1"]
                },
                "uxScores": {
                    "memberlogin": 90,
                    "search": 85,
                    "product": 88,
                    "checkout": 92,
                    "service": 87,
                    "other": 86
                },
                "uxAnalysis": {
                    "strengths": ["Strength 1"],
                    "improvements": ["Improvement 1"],
                    "summary": "Good UX"
                }
            }
        ]
    }
    
    response = client.post(
        "/generate-ppt/",
        files={"input_file": ("test.json", json.dumps(test_data), "application/json")}
    )
    assert response.status_code == 200
    assert "file_path" in response.json()
    
    # Clean up generated file
    file_path = response.json()["file_path"]
    if os.path.exists(file_path):
        os.remove(file_path)

def test_download_nonexistent_file():
    response = client.get("/download/nonexistent.pptx")
    assert response.status_code == 404
    assert "File not found" in response.json()["detail"] 