## ĐỀ THI THỬ

**4.1  Tổ chức thư mục bằng GUI**
Logo Van Phong Pham.PNG --- >logo_van_phong_pham.png |
                                                     | ---> Chứa trong thư mục gốc
Nhap Hang.txt --- > nhap_hang.txt                    |

Don Ban Thang 10.xlsx ---> don_ban_thang_10.xlsx | ---> Chứa trong thư mục data
                                                 
Shop.log ---> shop.log                           | ---> Chứa trong thư mục log

**4.2. Thao tác cơ bản bằng CLI**

|Yêu cầu|macOS (Terminal)|
|-|-|
|Đường dẫn tuyệt đối|pwd|
|Liệt kê cả file ẩn|ls -a|
|Di chuyển nhap_hang.txt, logo|mv nhap_hang.txt docs/ mv logo_van_phong_pham.png media/|
|Sao chép logo|cp media/logo_van_phong_pham.png backup/|
|Xoá file tạm|rm temp_131.tmp|
|Vào logs, liệt kê, quay lại|cd logs ls cd ..|
|Liệt kê toàn bộ dự án|ls -aR|

**4.3. Lọc log thủ công**
+ Đọc file: cat logs/shop.log.
+ Lọc nội dung: grep "ERROR"

---> Lệnh hoàn chỉnh: cat logs/shop.log.| grep "ERROR" > logs/daily_report.txt.

**4.4. Thao tác với tài liệu Markdown**
  - Heading 1: tên dự án.
  - Một mô tả ngắn có cả in đậm và in nghiêng.
  - Bảng Thư mục | Nội dung, ít nhất 3 dòng dữ liệu đúng.
  - Code block chứa lệnh PowerShell/Terminal liệt kê nội dung dự án.
  - Task list ít nhất 2 công việc, có ít nhất 1 mục đã hoàn thành.
  - Blockquote nhắc không upload .env.

**4.5. Hoàn thiện script có khung sẵn**

echo "=== BAO CAO LOI ===" > logs/auto_report.txt cat logs/shop.log | grep "ERROR" >> logs/auto_report.txt
---> bash auto_report.sh

**4.6. Đếm số dòng lỗi**

sử dụng lệnh wc - l kết hợp với dấu ( "|" ) 

**4.7. Báo cáo tổng hợp**
```
echo "=== TONG HOP LOG ===" > logs/tong_hop.txt cat logs/shop.log | grep "WARNING" >> logs/tong_hop.txt cat logs/shop.log | grep "ERROR" | wc -l >> logs/tong_hop.txt
```

**4.8. Excel**

ten_sp (E2):    =VLOOKUP(C2,SanPham!$A$2:$C$8,2,FALSE)
gia_ban (F2):   =VLOOKUP(C2,SanPham!$A$2:$C$8,3,FALSE)
tien_ban (G2):  =D2*F2