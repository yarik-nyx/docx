from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from docxtpl import DocxTemplate
import json
import io
import uvicorn

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"]
)



@app.post("/fill_template")
async def fill_template(
    docx_file: UploadFile = File(..., description="DOCX шаблон"),
    data: str  = Body(..., example={"name": "Иван Иванов"})
):
    try:
        # Читаем бинарные данные файла
        file_content = await docx_file.read()
        
        # Создаем временный файл в памяти
        template_stream = io.BytesIO(file_content)
        doc = DocxTemplate(template_stream)
        
        json_data = json.loads(data)

        # Заполняем шаблон данными из JSON
        doc.render(json_data)
        
        # Готовим файл для отдачи
        # output_stream = io.BytesIO()
        # doc.save(output_stream)
        # output_stream.seek(0)
        
        doc.save("filled_contract.docx")
        return FileResponse("filled_contract.docx", filename="filled_contract.docx")
        # return StreamingResponse(
        #     output_stream,
        #     media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        #     headers={"Content-Disposition": f"attachment; filename=result.docx"}
        # )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    


if __name__ == "__main__":
    uvicorn.run("main:app", host = "127.0.0.1", port=10000, reload=True)