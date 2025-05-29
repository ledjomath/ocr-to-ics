from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from ocr_utils import process_pdf
import shutil

app = FastAPI()

@app.post("/upload/")
async def upload_pdf(file: UploadFile = File(...)):
    with open("input.pdf", "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    xlsx_path, ics_path = process_pdf("input.pdf")
    return {"xlsx": xlsx_path, "ics": ics_path}

@app.get("/download/xlsx")
async def get_xlsx():
    return FileResponse("planning_juillet25.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/download/ics")
async def get_ics():
    return FileResponse("planning_juillet25.ics", media_type="text/calendar")