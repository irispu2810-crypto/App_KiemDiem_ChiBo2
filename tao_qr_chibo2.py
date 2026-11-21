# -*- coding: utf-8 -*-
"""
Tự động dò IP LAN và sinh Mã QR Code độ nét cao chuẩn VietinBank
Cho Ứng Dụng Kiểm Điểm Đảng Viên Chi Bộ 2 - VietinBank Bảo Lộc
"""

import socket
import qrcode
from PIL import Image, ImageDraw, ImageFont

def get_lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = socket.gethostbyname(socket.gethostname())
    finally:
        s.close()
    return ip

lan_ip = get_lan_ip()
port = 8080
url = f"http://{lan_ip}:{port}/app_kiem_diem_chi_bo_2.html"

print("=" * 70)
print("  TỰ ĐỘNG SINH MÃ QR CODE TRUY CẬP CHI BỘ 2 - VIETINBANK BẢO LỘC")
print("=" * 70)
print(f"Địa chỉ IP mạng nội bộ (LAN): {lan_ip}")
print(f"Cổng dịch vụ (Port):           {port}")
print(f"Đường link truy cập trực tiếp: {url}")
print("=" * 70)

# 1. Sinh mã QR chất lượng cao
qr = qrcode.QRCode(
    version=None,
    error_correction=qrcode.constants.ERROR_CORRECT_H,
    box_size=12,
    border=2,
)
qr.add_data(url)
qr.make(fit=True)

# Màu sắc thương hiệu VietinBank: Navy #002D62
qr_img = qr.make_image(fill_color='#002D62', back_color='white').convert('RGB')
qr_w, qr_h = qr_img.size

# 2. Tạo khung Card đồ họa 3D đẹp mắt cho ảnh QR
card_w = qr_w + 120
card_h = qr_h + 200

card = Image.new('RGB', (card_w, card_h), '#F8FAFC')
draw = ImageDraw.Draw(card)

# Vẽ dải banner trên cùng: VietinBank Navy
draw.rectangle([(0, 0), (card_w, 80)], fill='#002D62')
# Dải viền đỏ cờ Đảng
draw.rectangle([(0, 80), (card_w, 86)], fill='#DA251D')

# Dán mã QR vào giữa khung
card.paste(qr_img, ((card_w - qr_w) // 2, 110))

# Viền ngoài của khung
draw.rectangle([(0, 0), (card_w - 1, card_h - 1)], outline='#CBD5E1', width=2)

# Lưu các file ảnh QR
card.save('Ma_QR_ChiBo2_VietinBank.png')
qr_img.save('qr_share.png')
qr_img.save('qr_lan_app.png')

print("✅ Đã tạo thành công:")
print("   - Ma_QR_ChiBo2_VietinBank.png (Khung ảnh sang trọng để gửi Zalo/in ấn)")
print("   - qr_share.png (Dùng cho Cổng chia sẻ Portal)")
print("   - qr_lan_app.png (Dùng nhúng trong ứng dụng)")
print("=" * 70)
print("MÃ QR CODE ASCII (Quét trực tiếp từ màn hình bên dưới):")
qr.print_ascii(invert=True)
print("=" * 70)
