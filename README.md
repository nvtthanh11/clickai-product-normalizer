# 🚀 ClickAI Product Data Normalizer & GitHub Actions Workflow

Bộ công cụ tự động hóa chuẩn hóa dữ liệu sản phẩm cho **ClickAI**, hỗ trợ chuyển đổi từ mọi định dạng file Excel/CSV từ các sàn TMĐT (Shopee, Lazada, TikTok Shop, KiotViet, Sapo, Haravan, WooCommerce, Shopify,...) về đúng 16 cột tiêu chuẩn của ClickAI.

---

## 📋 16 Cột Tiêu Chuẩn Của ClickAI

Script sẽ tự động nhận diện và sắp xếp dữ liệu vào đúng 16 cột theo thứ tự:

1. `Tên sản phẩm`
2. `Tùy chọn giá`
3. `Mã SKU`
4. `Danh mục`
5. `Giá bán`
6. `Giá vốn`
7. `Tồn kho` *(mặc định: `0` nếu trống)*
8. `Đơn vị tính`
9. `Tiền tệ` *(mặc định: `"VND"` nếu trống)*
10. `Thuế (%)` *(mặc định: `0` nếu trống)*
11. `Chi phí trực tiếp` *(mặc định: `0` nếu trống)*
12. `Chu kỳ thanh toán` *(mặc định: `"Một lần"` nếu trống)*
13. `Xuất xứ`
14. `Mô tả`
15. `Link hình ảnh`
16. `Link sản phẩm`

---

## 📂 Cấu Trúc Thư Mục Repo

```text
├── .github/
│   └── workflows/
│       └── map_products.yml       # Quy trình GitHub Actions tự động
├── product_mapper_local.py        # Script Python chuẩn hóa đa năng (Excel, CSV, XML, JSON)
├── main.py                        # FastAPI Web Service kết nối ClickAI Workflow
├── Dockerfile & docker-compose    # Đóng gói triển khai IT Server riêng
├── requirements.txt               # Thư viện phụ thuộc
├── .gitignore                     # Bỏ qua file rác / file kết quả tạm
├── sample_products.csv            # File dữ liệu mẫu CSV
└── Báo cáo tồn khả dụng theo kho.xml # File dữ liệu mẫu XML MISA/ERP
```

---

## 💻 Cách Sử Dụng Trên Máy Tính Cá Nhân (Local)

### 1. Cài đặt thư viện:
```bash
pip install -r requirements.txt
```

### 2. Chạy script:
- **Tự động tìm file sản phẩm trong thư mục:**
  ```bash
  python product_mapper_local.py
  ```
- **Chỉ định cụ thể file nguồn:**
  ```bash
  python product_mapper_local.py --input "du_lieu_san_pham.xlsx"
  ```
- **Tùy chỉnh tên file đầu ra:**
  ```bash
  python product_mapper_local.py -i "du_lieu_san_pham.csv" -o "Ket_Qua_ClickAI.xlsx"
  ```

File kết quả sẽ được tạo tại: **`ClickAI_Product_Processed.xlsx`** (đã định dạng sẵn màu sắc, font chữ, độ rộng cột, zebra striping chuyên nghiệp).

---

## ☁️ Cách Sử Dụng Trên GitHub (GitHub Actions)

1. Đẩy (push) mã nguồn cùng file sản phẩm của bạn lên repository GitHub.
2. Vào tab **Actions** trên GitHub repository của bạn.
3. Chọn workflow **"Map Products to ClickAI Standard"** ở danh sách bên trái.
4. Bấm **Run workflow**:
   - Bạn có thể nhập tên file (ví dụ: `my_products.xlsx`) hoặc bỏ trống để GitHub tự động dò tìm file.
5. Sau khi quy trình chạy hoàn tất (chỉ mất ~20-30 giây), cuộn xuống mục **Artifacts** ở cuối trang chạy để tải file **`ClickAI_Product_Processed.xlsx`** về máy.
