#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
ClickAI Product Normalizer - FastAPI Server
=============================================================================
Dịch vụ Web API tiếp nhận file sản phẩm từ Click AI Workflow, chuẩn hóa dữ liệu
và trả về link tải file Excel hoàn chỉnh theo chuẩn 16 cột ClickAI.
=============================================================================
"""

import os
import shutil
import uuid
import time
from pathlib import Path
from typing import Optional
import urllib.parse

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import requests

import re
from product_mapper_local import process_product_file, remove_accents

app = FastAPI(
    title="ClickAI Product Normalizer API",
    description="API chuẩn hóa dữ liệu sản phẩm từ mọi sàn TMĐT sang 16 cột tiêu chuẩn ClickAI",
    version="1.0.0"
)

# Cho phép CORS để Click AI và web clients có thể gọi trực tiếp
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thư mục lưu trữ file tạm
STORAGE_DIR = Path(os.getenv("STORAGE_DIR", "/tmp/clickai_jobs"))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
FILE_RETENTION_SECONDS = int(os.getenv("RETENTION_SECONDS", 7200))  # Tự xóa sau 2 giờ


class ConvertUrlRequest(BaseModel):
    file_url: str
    custom_filename: Optional[str] = None


def cleanup_old_jobs():
    """Tự động dọn dẹp các file tạm cũ hơn 2 giờ để tránh đầy ổ cứng."""
    now = time.time()
    try:
        for job_folder in STORAGE_DIR.iterdir():
            if job_folder.is_dir():
                folder_age = now - job_folder.stat().st_mtime
                if folder_age > FILE_RETENTION_SECONDS:
                    shutil.rmtree(job_folder, ignore_errors=True)
    except Exception as e:
        print(f"[!] Lỗi khi dọn dẹp file tạm: {e}")


def get_base_url(request: Request) -> str:
    """Xác định Base URL của server (hỗ trợ cả trường hợp đứng sau Reverse Proxy Nginx)."""
    env_base = os.getenv("BASE_URL")
    if env_base:
        return env_base.rstrip("/")
    forwarded_proto = request.headers.get("x-forwarded-proto", request.url.scheme)
    host = request.headers.get("x-forwarded-host", request.headers.get("host", str(request.base_url.netloc)))
    return f"{forwarded_proto}://{host}"


@app.get("/")
@app.get("/health")
@app.get("/api/v1/health")
def health_check():
    """Kiểm tra trạng thái server."""
    return {
        "status": "online",
        "service": "ClickAI Product Normalizer API",
        "version": "1.0.0",
        "timestamp": int(time.time())
    }


@app.post("/api/convert")
async def convert_file(
    request: Request,
    background_tasks: BackgroundTasks,
    payload: Optional[ConvertUrlRequest] = None,
    file: Optional[UploadFile] = File(None)
):
    """
    Endpoint chuẩn hóa dữ liệu sản phẩm:
    Hỗ trợ 2 cách truyền file:
    1. Gửi JSON body: {"file_url": "https://..."} (Khuyên dùng trên Click AI Workflow)
    2. Upload file trực tiếp (multipart/form-data)
    """
    background_tasks.add_task(cleanup_old_jobs)
    job_id = str(uuid.uuid4())[:12]
    job_dir = STORAGE_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = None

    try:
        # Trường hợp 1: Nhận file qua file_url
        if payload and payload.file_url:
            file_url = payload.file_url.strip()
            print(f"[*] Nhận yêu cầu tải file từ URL: {file_url}")
            
            # Lấy tên file từ URL
            parsed_url = urllib.parse.urlparse(file_url)
            url_filename = os.path.basename(parsed_url.path)
            if not url_filename or "." not in url_filename:
                url_filename = "input_products.xlsx"

            input_path = job_dir / url_filename

            # Tải file về máy chủ (giới hạn timeout 30s)
            response = requests.get(file_url, stream=True, timeout=30)
            if response.status_code != 200:
                raise HTTPException(
                    status_code=400,
                    detail=f"Không thể tải file từ URL được cung cấp (Mã HTTP: {response.status_code})"
                )

            with open(input_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

        # Trường hợp 2: Nhận file upload trực tiếp
        elif file:
            filename = file.filename or "input_products.xlsx"
            input_path = job_dir / filename
            with open(input_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

        else:
            raise HTTPException(
                status_code=400,
                detail="Vui lòng cung cấp 'file_url' qua JSON hoặc upload file đính kèm!"
            )

        # Kiểm tra dung lượng file (tối đa 50MB)
        if input_path.stat().st_size > 50 * 1024 * 1024:
            raise HTTPException(status_code=400, detail="Dung lượng file vượt quá giới hạn cho phép (tối đa 50MB).")

        # Tiến hành chuẩn hóa dữ liệu
        output_filename = "ClickAI_Product_Processed.xlsx"
        output_path = job_dir / output_filename

        result_path = process_product_file(input_path, output_path)

        # Nếu có tham số return_file=true hoặc gọi /api/convert-file: Trả trực tiếp file Excel
        if request.query_params.get("return_file") == "true":
            return FileResponse(
                path=result_path,
                filename=output_filename,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename={output_filename}"}
            )

        base_url = get_base_url(request)
        download_url = f"{base_url}/api/download/{job_id}"

        return JSONResponse(
            status_code=200,
            content={
                "success": True,
                "job_id": job_id,
                "filename": output_filename,
                "download_url": download_url,
                "message": "Chuẩn hóa dữ liệu sản phẩm thành công theo 16 cột ClickAI!"
            }
        )

    except HTTPException:
        raise
    except ValueError as ve:
        # Lỗi về dữ liệu (ví dụ: không tìm thấy cột sản phẩm hợp lệ)
        return JSONResponse(
            status_code=422,
            content={
                "success": False,
                "error": str(ve),
                "error_type": "DataValidationError"
            }
        )
    except Exception as e:
        print(f"[LỖI XỬ LÝ]: {e}")
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": f"Lỗi hệ thống trong quá trình xử lý: {str(e)}",
                "error_type": "InternalServerError"
            }
        )


@app.post("/api/convert-file")
async def convert_and_return_file(
    request: Request,
    background_tasks: BackgroundTasks,
    payload: Optional[ConvertUrlRequest] = None,
    file: Optional[UploadFile] = File(None)
):
    """
    Endpoint chuẩn hóa và trả về TRỰC TIẾP file Excel (.xlsx) trong response.
    Dành riêng cho Click AI Workflow bắt lấy biến `files: Array[File]`.
    """
    background_tasks.add_task(cleanup_old_jobs)
    job_id = str(uuid.uuid4())[:12]
    job_dir = STORAGE_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = None
    try:
        if payload and payload.file_url:
            file_url = payload.file_url.strip()
            parsed_url = urllib.parse.urlparse(file_url)
            url_filename = os.path.basename(parsed_url.path) or "input_products.xlsx"
            input_path = job_dir / url_filename
            response = requests.get(file_url, stream=True, timeout=30)
            if response.status_code != 200:
                raise HTTPException(status_code=400, detail=f"Không thể tải file từ URL: {response.status_code}")
            with open(input_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
        elif file:
            filename = file.filename or "input_products.xlsx"
            input_path = job_dir / filename
            with open(input_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
        else:
            raise HTTPException(status_code=400, detail="Vui lòng cung cấp file_url qua JSON hoặc upload file đính kèm!")

        # Tạo tên file đẹp theo tên file gốc của khách hàng (VD: Processed_Bao_cao_ton_kha_dung.xlsx)
        orig_stem = input_path.stem
        clean_stem = re.sub(r'[^a-zA-Z0-9_\-]', '_', remove_accents(orig_stem)).strip('_')
        if not clean_stem or clean_stem.lower() in ["input_products", "file", "tmp"]:
            clean_stem = "ClickAI_Product"
        output_filename = f"Processed_{clean_stem}.xlsx"

        output_path = job_dir / output_filename
        result_path = process_product_file(input_path, output_path)

        return FileResponse(
            path=result_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"; filename*=UTF-8\'\'{output_filename}'
            }
        )
    except HTTPException:
        raise
    except ValueError as ve:
        raise HTTPException(status_code=422, detail=str(ve))
    except Exception as e:
        print(f"[LỖI XỬ LÝ]: {e}")
        raise HTTPException(status_code=500, detail=f"Lỗi chuẩn hóa file: {str(e)}")


@app.get("/api/download/{job_id}")
def download_processed_file(job_id: str):
    """Endpoint tải file kết quả đã xử lý."""
    target_file = STORAGE_DIR / job_id / "ClickAI_Product_Processed.xlsx"
    if not target_file.exists():
        raise HTTPException(
            status_code=404,
            detail="File không tồn tại hoặc đã hết hạn tải về (file được lưu tối đa 2 giờ)."
        )

    return FileResponse(
        path=target_file,
        filename="ClickAI_Product_Processed.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    print(f"[*] Khởi động ClickAI Product Normalizer Server tại port {port}...")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
