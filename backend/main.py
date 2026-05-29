from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os
import asyncio
from fastapi import File, UploadFile, Form
import json

# Services
from backend.services.ppt_service import generate_slide_content, create_ppt
from fastapi.responses import JSONResponse
import base64

app = FastAPI(title="AI Generator API 🚀")

# =========================
# 📌 Request Model
# =========================
class PPTRequest(BaseModel):
    topic: str
    num_slides: int = 5
    tone: str = "Professional"


class PPTBuildRequest(BaseModel):
    topic: str
    tone: str = "Professional"
    slide_data: dict


# =========================
# 🏠 Health Check
# =========================
@app.get("/")
def home():
    return {"message": "FastAPI backend running 🚀"}


# =========================
# 📊 Generate PPT Endpoint
# =========================
@app.post("/generate-ppt")
async def generate_ppt(
    topic: str = Form(...),
    num_slides: int = Form(5),
    tone: str = Form("Professional"),
    logo: UploadFile = File(None),
    content_image: UploadFile = File(None),
):
    try:
        # Save logo if provided
        logo_path = None
        if logo:
            os.makedirs("uploads", exist_ok=True)
            logo_path = f"uploads/{logo.filename}"

            with open(logo_path, "wb") as f:
                f.write(await logo.read())

        content_image_path = None
        if content_image:
            os.makedirs("uploads", exist_ok=True)
            content_image_path = f"uploads/{content_image.filename}"
            with open(content_image_path, "wb") as f:
                f.write(await content_image.read())

        # Generate content (bounded timeout so request doesn't hang forever)
        try:
            slide_data = await asyncio.wait_for(
                asyncio.to_thread(generate_slide_content, topic, num_slides, tone, logo_path),
                timeout=150,
            )
        except asyncio.TimeoutError:
            raise HTTPException(
                status_code=504,
                detail="Slide generation timed out after 150 seconds. Try fewer slides or retry.",
            )

        # Build PPT (also bounded)
        try:
            file_path = await asyncio.wait_for(
                asyncio.to_thread(create_ppt, slide_data, topic, logo_path, tone, content_image_path),
                timeout=90,
            )
        except asyncio.TimeoutError:
            raise HTTPException(
                status_code=504,
                detail="PPT rendering timed out after 90 seconds. Please retry.",
            )

        # Create PPT with logo
        from fastapi.responses import JSONResponse
        import base64

        # Read file
        with open(file_path, "rb") as f:
            ppt_bytes = f.read()

        # Convert to base64
        ppt_base64 = base64.b64encode(ppt_bytes).decode()

        # Return slides + file
        return JSONResponse({
            "slides": slide_data["slides"],
            "ppt_base64": ppt_base64,
            "usage": slide_data.get("usage"),
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate-outline")
async def generate_outline(
    topic: str = Form(...),
    num_slides: int = Form(5),
    tone: str = Form("Professional"),
):
    try:
        try:
            slide_data = await asyncio.wait_for(
                asyncio.to_thread(generate_slide_content, topic, num_slides, tone),
                timeout=150,
            )
        except asyncio.TimeoutError:
            raise HTTPException(
                status_code=504,
                detail="Slide generation timed out after 150 seconds. Try fewer slides or retry.",
            )

        return JSONResponse(slide_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/build-ppt")
async def build_ppt(
    topic: str = Form(...),
    tone: str = Form("Professional"),
    slides_json: str = Form(...),
    logo: UploadFile = File(None),
    content_image: UploadFile = File(None),
):
    try:
        try:
            slide_data = json.loads(slides_json)
        except json.JSONDecodeError as e:
            raise HTTPException(status_code=400, detail=f"Invalid slides_json: {e}")

        logo_path = None
        if logo:
            os.makedirs("uploads", exist_ok=True)
            logo_path = f"uploads/{logo.filename}"
            with open(logo_path, "wb") as f:
                f.write(await logo.read())

        content_image_path = None
        if content_image:
            os.makedirs("uploads", exist_ok=True)
            content_image_path = f"uploads/{content_image.filename}"
            with open(content_image_path, "wb") as f:
                f.write(await content_image.read())

        try:
            file_path = await asyncio.wait_for(
                asyncio.to_thread(create_ppt, slide_data, topic, logo_path, tone, content_image_path),
                timeout=90,
            )
        except asyncio.TimeoutError:
            raise HTTPException(
                status_code=504,
                detail="PPT rendering timed out after 90 seconds. Please retry.",
            )

        with open(file_path, "rb") as f:
            ppt_bytes = f.read()

        ppt_base64 = base64.b64encode(ppt_bytes).decode()
        return JSONResponse({
            "slides": slide_data.get("slides", []),
            "ppt_base64": ppt_base64,
        })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))