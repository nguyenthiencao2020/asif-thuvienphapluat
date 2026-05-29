# Thư Viện Pháp Luật – Đăng Ký Doanh Nghiệp

Ứng dụng web tĩnh tra cứu pháp luật doanh nghiệp + **điền biểu mẫu TT 68/2025
và xuất file Word (.docx)** đúng theo bản gốc của Bộ Tài chính.

## Tính năng tự điền biểu mẫu (Hướng C – điền vào mẫu gốc)

Khi người dùng nhập thông tin trong form và bấm **"Tạo văn bản Word"**, hệ thống:

1. Lấy file biểu mẫu gốc trong `templates/<id>.docx` (đã chèn sẵn placeholder
   dạng `{ten_field}` đúng vị trí các dòng "………" của mẫu gốc).
2. Dùng [docxtemplater](https://docxtemplater.com/) + [PizZip] (chạy **hoàn toàn
   phía trình duyệt**, nạp qua CDN) để thay placeholder bằng dữ liệu người nhập.
3. Tải về file `.docx` đã điền – giữ **nguyên định dạng bản gốc**.

Thông tin dùng chung (họ tên, CCCD, ngày sinh, cơ quan…) được lưu `localStorage`
và **tự điền chéo** sang các mẫu khác.

> Tính năng cần chạy qua web (http/https hoặc Vercel), không chạy khi mở trực
> tiếp `index.html` bằng `file://` (do trình duyệt chặn `fetch`).

## Các mẫu đã hỗ trợ điền + xuất Word

| ID | Biểu mẫu |
|----|----------|
| pl1-1 | Giấy đề nghị ĐKDN – Doanh nghiệp tư nhân |
| pl1-2 | Giấy đề nghị ĐKDN – Công ty TNHH một thành viên |
| pl1-3 | Giấy đề nghị ĐKDN – Công ty TNHH hai thành viên trở lên |
| pl1-4 | Giấy đề nghị ĐKDN – Công ty cổ phần |
| pl1-5 | Giấy đề nghị ĐKDN – Công ty hợp danh |
| pl1-6 | Danh sách thành viên công ty TNHH 2TV (tự điền bảng) |
| pl1-7 | Danh sách cổ đông sáng lập (tự điền bảng) |
| pl1-13 | Thay đổi người đại diện theo pháp luật |
| pl1-27 | Thông báo tạm ngừng / tiếp tục kinh doanh |
| pl2-1 | Giấy đề nghị đăng ký hộ kinh doanh |

## Sinh lại template

Các file `templates/*.docx` được sinh tự động từ `Thư Viện Pháp Luật/Phu luc 68_2025.docx`:

```bash
pip install python-docx
python3 scripts/build_templates.py
```

Script tách từng "Mẫu số N" ra file riêng (giữ nguyên định dạng), chèn
placeholder docxtemplater vào các dòng cần điền (`MAPS`) và tạo vòng lặp cho
bảng động – ngành nghề, danh sách thành viên/cổ đông (`TABLE_LOOPS`).

## Deploy lên Vercel

Đây là site tĩnh – import repo vào Vercel là chạy ngay (đã có `vercel.json`).
Không cần backend / serverless.
