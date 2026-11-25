# -*- coding: utf-8 -*-
"""
Tự động dò IP LAN và sinh Mã QR Code độ nét cao chuẩn VietinBank
Bao gồm cả 2 phiên bản mới nhất:
1. Online 24/7 trên GitHub Pages (hoạt động kể cả khi tắt máy chủ)
2. Mạng nội bộ LAN / Wi-Fi phòng họp VietinBank Bảo Lộc
"""

import socket
import os
import shutil
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
url_lan = f"http://{lan_ip}:{port}/?v=2026.11.25.16"
url_github = "https://irispu2810-crypto.github.io/App_KiemDiem_ChiBo2/?v=2026.11.25.16"

print("=" * 72)
print("  ⭐ TỰ ĐỘNG SINH BỘ MÃ QR CODE PHIÊN BẢN MỚI NHẤT - CHI BỘ 2 ⭐")
print("=" * 72)
print(f"1. Đường link Online 24/7 (GitHub Pages): {url_github}")
print(f"2. Đường link Mạng Nội bộ (LAN Wi-Fi):     {url_lan}")
print("=" * 72)

# Load fonts
try:
    font_title = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 20)
    font_sub = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 15)
    font_regular = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 13)
    font_small = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 12)
    font_mono = ImageFont.truetype('C:/Windows/Fonts/cour.ttf', 13)
except Exception:
    font_title = font_sub = font_regular = font_small = font_mono = ImageFont.load_default()

def create_qr_card(url, title, badge_text, badge_color, card_filename, qr_filename, note_text):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color='#002D62', back_color='white').convert('RGB')
    qr_w, qr_h = qr_img.size
    
    # Lưu ảnh QR thuần
    qr_img.save(qr_filename)

    card_w = 600
    card_h = qr_h + 340
    card = Image.new('RGB', (card_w, card_h), '#FFFFFF')
    draw = ImageDraw.Draw(card)

    # Dải Header VietinBank & Cờ Đảng
    draw.rectangle([(0, 0), (card_w, 95)], fill='#002D62')
    draw.rectangle([(0, 95), (card_w, 100)], fill='#DA251D')
    draw.rectangle([(0, 100), (card_w, 103)], fill='#DAA520')

    t1 = 'ĐẢNG BỘ NGÂN HÀNG TMCP CÔNG THƯƠNG VIỆT NAM'
    t2 = 'CHI BỘ 2 - CHI NHÁNH BẢO LỘC'
    bbox1 = draw.textbbox((0, 0), t1, font=font_small)
    bbox2 = draw.textbbox((0, 0), t2, font=font_title)
    draw.text(((card_w - (bbox1[2] - bbox1[0])) // 2, 18), t1, fill='#E2E8F0', font=font_small)
    draw.text(((card_w - (bbox2[2] - bbox2[0])) // 2, 42), t2, fill='#FFFFFF', font=font_title)

    # Huy hiệu Badge
    bbox_b = draw.textbbox((0, 0), badge_text, font=font_sub)
    bw = bbox_b[2] - bbox_b[0] + 30
    bh = 32
    bx = (card_w - bw) // 2
    by = 120
    draw.rounded_rectangle([(bx, by), (bx + bw, by + bh)], radius=8, fill=badge_color)
    draw.text(((card_w - (bbox_b[2] - bbox_b[0])) // 2, by + 7), badge_text, fill='#FFFFFF', font=font_sub)

    # Đặt ảnh QR ở giữa
    qr_x = (card_w - qr_w) // 2
    qr_y = 170
    draw.rounded_rectangle([(qr_x - 6, qr_y - 6), (qr_x + qr_w + 6, qr_y + qr_h + 6)], radius=12, fill='#F8FAFC', outline='#CBD5E1', width=2)
    card.paste(qr_img, (qr_x, qr_y))

    # Khung hiển thị URL
    uy = qr_y + qr_h + 20
    bbox_u = draw.textbbox((0, 0), url, font=font_mono)
    uw = min(card_w - 40, bbox_u[2] - bbox_u[0] + 24)
    ux = (card_w - uw) // 2
    draw.rounded_rectangle([(ux, uy), (ux + uw, uy + 30)], radius=6, fill='#F1F5F9', outline='#94A3B8', width=1)
    draw.text(((card_w - (bbox_u[2] - bbox_u[0])) // 2, uy + 7), url, fill='#0F172A', font=font_mono)

    # Hướng dẫn thao tác
    iy = uy + 42
    inst = 'Mở Camera điện thoại (Zalo / iPhone / Android) quét mã để vào ứng dụng'
    bbox_i = draw.textbbox((0, 0), inst, font=font_regular)
    draw.text(((card_w - (bbox_i[2] - bbox_i[0])) // 2, iy), inst, fill='#334155', font=font_regular)

    bbox_n = draw.textbbox((0, 0), note_text, font=font_small)
    draw.text(((card_w - (bbox_n[2] - bbox_n[0])) // 2, iy + 24), note_text, fill='#64748B', font=font_small)

    # Viền bao ngoài
    draw.rectangle([(0, 0), (card_w - 1, card_h - 1)], outline='#CBD5E1', width=2)

    card.save(card_filename)
    return qr

# 1. Sinh Mã QR Online 24/7 (GitHub Pages)
qr_git = create_qr_card(
    url_github,
    'ỨNG DỤNG KIỂM ĐIỂM ĐẢNG VIÊN 2026',
    '🌐 TRUY CẬP TRỰC TUYẾN 24/7 (KỂ CẢ TẮT MÁY CHỦ)',
    '#059669',
    'Ma_QR_GitHub_ChiBo2_VietinBank.png',
    'qr_github.png',
    'User: Tên viết liền không dấu (VD: nguyenthithao) • Mật khẩu: 123456'
)

# 2. Sinh Mã QR LAN Mạng Nội Bộ
qr_lan_obj = create_qr_card(
    url_lan,
    'ỨNG DỤNG KIỂM ĐIỂM ĐẢNG VIÊN 2026',
    '🏢 MẠNG NỘI BỘ LAN / WI-FI HỘI TRƯỜNG',
    '#0284C7',
    'Ma_QR_ChiBo2_VietinBank.png',
    'qr_lan_app.png',
    'Kết nối cùng mạng Wi-Fi VietinBank Bảo Lộc để đồng bộ tức thì'
)

shutil.copy('Ma_QR_ChiBo2_VietinBank.png', 'Ma_QR_LAN_ChiBo2_VietinBank.png')
shutil.copy('qr_lan_app.png', 'qr_share.png')

print("✅ Đã tạo thành công bộ mã QR Code phiên bản mới nhất:")
print("   [1] Ma_QR_GitHub_ChiBo2_VietinBank.png -> Bản Online 24/7 (Gửi Zalo/In ấn)")
print("   [2] Ma_QR_ChiBo2_VietinBank.png        -> Bản Mạng nội bộ LAN Wi-Fi")
print("   [3] qr_github.png & qr_lan_app.png     -> Nhúng trực tiếp trong Web App & Portal")
print("=" * 72)
