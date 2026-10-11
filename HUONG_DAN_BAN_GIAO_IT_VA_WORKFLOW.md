# 📘 TÀI LIỆU BÀN GIAO TRIỂN KHAI CHO IT & HƯỚNG DẪN KẾT NỐI WORKFLOW CLICK AI

Tài liệu này được biên soạn để bạn có thể **gửi trực tiếp cho bộ phận IT** để đưa dịch vụ chuẩn hóa dữ liệu lên máy chủ riêng của Click AI, đồng thời hướng dẫn bạn **kết nối với Workflow của Click AI**.

---

## 🎯 1. TỔNG QUAN HỆ THỐNG

Dịch vụ **ClickAI Product Normalizer** là một Microservice API xây dựng bằng **FastAPI (Python 3.11)**.
- **Nhiệm vụ:** Nhận URL file sản phẩm (hỗ trợ mọi định dạng: Excel `.xlsx`/`.xls`/`.xlsm`, `.csv`, `.tsv`, `.txt`, `.xml`, `.json`, `.jsonl`) từ Workflow -> Tự động nhận diện cấu trúc, cột, sheet -> Chuẩn hóa về **đúng 16 cột tiêu chuẩn của ClickAI** -> Định dạng bảng tính (Header Navy, border, zebra styling, auto width) -> Cung cấp link download trực tiếp cho khách hàng.
- **Ưu điểm khi chạy trên Server riêng:**
  - Tốc độ xử lý siêu nhanh: **1 - 3 giây** (không bị độ trễ chờ đợi như GitHub Actions).
  - Khách hàng tải file trực tiếp, không yêu cầu đăng nhập.
  - Tự động dọn dẹp file tạm sau 2 giờ (chống đầy ổ cứng).
  - Chịu tải tốt nhiều người convert cùng lúc.

---

## 💻 2. HƯỚNG DẪN CHO IT TRIỂN KHAI LÊN SERVER CLICK AI

IT chỉ cần chọn **1 trong 2 cách** sau để chạy dịch vụ:

### CÁCH 1: Triển khai bằng Docker (Khuyên dùng - Nhanh nhất & Tự động chạy lại)

1. **Copy thư mục mã nguồn lên Server Click AI** (hoặc clone từ Git).
2. **Vào thư mục dự án và chạy 1 câu lệnh duy nhất:**
   ```bash
   docker compose up -d --build
   ```
3. Kiểm tra log dịch vụ:
   ```bash
   docker compose logs -f
   ```
   *Dịch vụ sẽ tự động chạy tại cổng `8000` và tự khởi động lại nếu server restart (`restart: always`).*

---

### CÁCH 2: Chạy trực tiếp bằng Python & Quản lý bằng Systemd

Nếu server không dùng Docker, IT có thể chạy bằng môi trường Python máy chủ:

1. **Cài đặt thư viện:**
   ```bash
   pip install -r requirements.txt
   ```
2. **Khởi động server:**
   ```bash
   uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4
   ```
3. **Cấu hình chạy nền liên tục bằng Systemd (`/etc/systemd/system/clickai-mapper.service`):**
   ```ini
   [Unit]
   Description=ClickAI Product Mapper API Service
   After=network.target

   [Service]
   User=root
   WorkingDirectory=/path/to/project
   ExecStart=/usr/local/bin/uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```
   Kích hoạt dịch vụ:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable clickai-mapper
   sudo systemctl start clickai-mapper
   ```

---

### CẤU HÌNH NGINX & SSL (Nếu muốn gán Domain riêng)

IT có thể trỏ domain (ví dụ `mapper.clickai.vn` hoặc dùng IP server) qua cấu hình Nginx mẫu:

```nginx
server {
    server_name mapper.clickai.vn;

    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

---

## 📡 3. THÔNG SỐ KỸ THUẬT API (API SPECIFICATION)

### 3.1. Kiểm tra trạng thái máy chủ (Health Check)
- **Method:** `GET`
- **URL:** `http://<IP_HOẶC_DOMAIN>:8000/`
- **Response:**
  ```json
  {
    "status": "online",
    "service": "ClickAI Product Normalizer API",
    "version": "1.0.0"
  }
  ```

---

### 3.2. Chuẩn hóa File (Convert Endpoint)
- **Method:** `POST`
- **URL:** `http://<IP_HOẶC_DOMAIN>:8000/api/convert`
- **Header:** `Content-Type: application/json`
- **Request Body:**
  ```json
  {
    "file_url": "https://domain.com/link_file_khach_gui.xlsx"
  }
  ```
- **Response Thành Công (HTTP 200):**
  ```json
  {
    "success": true,
    "job_id": "9a7d3f1b4c2e",
    "filename": "ClickAI_Product_Processed.xlsx",
    "download_url": "http://<IP_HOẶC_DOMAIN>:8000/api/download/9a7d3f1b4c2e",
    "message": "Chuẩn hóa dữ liệu sản phẩm thành công theo 16 cột ClickAI!"
  }
  ```
- **Response Khi File Không Hợp Lệ (HTTP 422):**
  ```json
  {
    "success": false,
    "error": "Không nhận diện được cột sản phẩm hợp lệ trong file (cần tối thiểu cột 'Tên sản phẩm', 'Mã SKU' hoặc 'Giá bán')...",
    "error_type": "DataValidationError"
  }
  ```

---

### 3.3. Tải File Kết Quả (Download Endpoint)
- **Method:** `GET`
- **URL:** `http://<IP_HOẶC_DOMAIN>:8000/api/download/{job_id}`
- **Kết quả:** Trả về trực tiếp file Excel `.xlsx` với định dạng tải về ngay.

---

## 🤖 4. CÁCH KẾT NỐI VỚI WORKFLOW CLICK AI

Trong giao diện thiết kế Workflow trên Click AI:

### BƯỚC 1: Node Nhận Tin Nhắn & File Từ Khách Hàng (User Input)
- Khách hàng tải file lên cửa sổ Chat của Agent.
- Lấy biến URL của file: giả sử là `{{file_url}}`.

### BƯỚC 2: Node HTTP Request (Gọi về Server API)
- **Cấu hình Node:**
  - **Method:** `POST`
  - **URL:** `http://<IP_HOẶC_DOMAIN_CỦA_SERVER>:8000/api/convert`
  - **Headers:** 
    - `Content-Type`: `application/json`
  - **Body (JSON):**
    ```json
    {
      "file_url": "{{file_url}}"
    }
    ```
  - **Timeout:** `30s`

### BƯỚC 3: Node Phân Nhánh Điều Kiện (Condition Node)
- **Nếu `{{http_response.success}} == true`**:
  - Chuyển sang Node Trả Lời Thành Công.
- **Nếu `{{http_response.success}} == false`**:
  - Chuyển sang Node Báo Lỗi: *"Xin lỗi, file bạn gửi không hợp lệ: {{http_response.error}}. Vui lòng kiểm tra lại file!"*

### BƯỚC 4: Node Tin Nhắn Trả Lời Của Agent (LLM / Message)
- Agent gửi tin nhắn kèm link tải:
  > 🎉 Dữ liệu sản phẩm của bạn đã được chuẩn hóa thành công sang **16 cột tiêu chuẩn của ClickAI**!
  > 
  > 📥 **[Bấm vào đây để tải file kết quả Excel]({{http_response.download_url}})**
  > *(Đường dẫn có hiệu lực trong 2 giờ)*
