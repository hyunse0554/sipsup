
from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import JSONResponse, FileResponse
import msoffcrypto
import pandas as pd
import io
import os

app = FastAPI()

PASSWORD = "1234"  # 고정 암호

@app.post("/search")
async def search_excel(file: UploadFile, dong: str = Form(...), ho: str = Form(...)):
    try:
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file.file)
        office_file.load_key(password=PASSWORD)
        office_file.decrypt(decrypted)

        df = pd.read_excel(decrypted)

        filtered = df[(df['동'] == dong) & (df['호수'] == ho)]

        if filtered.empty:
            return JSONResponse({"message": "검색 결과가 없습니다."}, status_code=404)

        output = io.BytesIO()
        filtered.to_excel(output, index=False)
        output.seek(0)

        result_path = "result.xlsx"
        with open(result_path, "wb") as f:
            f.write(output.read())

        return FileResponse(result_path, filename="검색결과.xlsx")

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
