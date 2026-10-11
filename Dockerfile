# Sử dụng image Python 3.11 slim nhẹ và bảo mật
FROM python:3.11-slim

# Thiết lập thư mục làm việc trong container
WORKDIR /app

# Cài đặt các gói hệ thống cần thiết
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy và cài đặt thư viện phụ thuộc
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn vào container
COPY . .

# Tạo thư mục lưu file tạm
RUN mkdir -p /tmp/clickai_jobs

# Mở cổng 8000
EXPOSE 8000

# Chạy ứng dụng bằng Uvicorn với nhiều worker để phục vụ đồng thời
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]
