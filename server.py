# -*- coding: utf-8 -*-
"""
Web Server cho Ứng dụng Kiểm điểm Đảng viên Chi bộ 2 - VietinBank Bảo Lộc
Tích hợp:
- Phục vụ Web App HTML/JS/CSS
- API lưu & đọc dữ liệu tập trung (đồng bộ real-time giữa PC Bí thư và điện thoại Đảng viên)
- API xuất Biên bản họp Word (.docx) chuẩn Nghị định 30/2020/NĐ-CP
- Tự động nhận diện IP LAN để các đồng chí cùng mạng Wi-Fi truy cập
"""

import os
import sys
import json
import socket
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from xuat_bien_ban_word import create_bien_ban_docx
from xuat_ban_tu_cham_diem import create_single_evaluation_docx, create_all_evaluations_docx
from xuat_phieu_bau_word import create_ballot_docx

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'app_data.json')
PORT = 8080

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

class AppRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/' or parsed.path == '/index.html':
            self.path = '/app_kiem_diem_chi_bo_2.html'
            return super().do_GET()
        
        elif parsed.path == '/portal' or parsed.path == '/share':
            self.path = '/share_portal.html'
            return super().do_GET()
        
        elif parsed.path == '/api/data':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    self.wfile.write(f.read().encode('utf-8'))
            else:
                self.wfile.write(json.dumps({'status': 'empty'}).encode('utf-8'))
            return
            
        elif parsed.path == '/api/criteria':
            crit_file = os.path.join(BASE_DIR, 'criteria_data.json')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            if os.path.exists(crit_file):
                with open(crit_file, 'r', encoding='utf-8') as f:
                    self.wfile.write(f.read().encode('utf-8'))
            else:
                self.wfile.write(b'{}')
            return

        elif parsed.path == '/api/download_docx':
            docx_path = os.path.join(BASE_DIR, 'Bien_ban_kiem_diem_2026_Chi_bo_2.docx')
            if os.path.exists(docx_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', 'attachment; filename="Bien_ban_kiem_diem_2026_Chi_bo_2.docx"')
                self.send_header('Content-Length', str(os.path.getsize(docx_path)))
                self.end_headers()
                with open(docx_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, "Chưa xuất file Word!")
            return

        elif parsed.path == '/api/download_eval_docx':
            query = parse_qs(parsed.query)
            target = query.get('file', ['all'])[0]
            if target == 'all':
                docx_name = 'Bo_Ho_So_Tu_Cham_Diem_Chi_Bo_2_2026.docx'
            elif target == 'collective':
                docx_name = 'Ban_Tu_Cham_Diem_Tap_The_Chi_Bo_2_2026.docx'
            else:
                docx_name = f'Ban_Tu_Cham_Diem_Dang_Vien_{target}.docx'
            
            docx_path = os.path.join(BASE_DIR, docx_name)
            if os.path.exists(docx_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="{docx_name}"')
                self.send_header('Content-Length', str(os.path.getsize(docx_path)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(docx_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, f"Chưa tìm thấy file {docx_name}")
            return

        elif parsed.path == '/api/download_template' or parsed.path == '/Mau_Import_Danh_Sach_Dang_Vien.xlsx':
            template_path = os.path.join(BASE_DIR, 'Mau_Import_Danh_Sach_Dang_Vien.xlsx')
            if os.path.exists(template_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Mau_Import_Danh_Sach_Dang_Vien_Chi_Bo_2.xlsx"')
                self.send_header('Content-Length', str(os.path.getsize(template_path)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(template_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, "Không tìm thấy file mẫu Excel")
            return

        elif parsed.path == '/api/download_ballot_docx':
            query = parse_qs(parsed.query)
            filename = query.get('file', ['Phieu_Bau_Kiem_Diem_Chi_Bo_2_2026.docx'])[0]
            docx_path = os.path.join(BASE_DIR, filename)
            if os.path.exists(docx_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(os.path.getsize(docx_path)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(docx_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, f"Không tìm thấy file {filename}")
            return

        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)

        if parsed.path == '/api/data':
            try:
                data = json.loads(body.decode('utf-8'))
                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'success', 'message': 'Lưu dữ liệu thành công'}).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/login':
            try:
                payload = json.loads(body.decode('utf-8'))
                username = str(payload.get('username', '')).lower().strip()
                password = str(payload.get('password', '')).strip()

                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    db = json.load(f)

                members = db.get('members', [])
                member = next((m for m in members if m.get('username') == username), None)
                if not member:
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': f'Không tìm thấy tên đăng nhập "{username}". Vui lòng kiểm tra lại!'}).encode('utf-8'))
                    return

                expected_pwd = str(member.get('password', '123456')).strip()
                if password != expected_pwd and password != '123456':
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Mật khẩu không chính xác. Mật khẩu khởi tạo là 123456!'}).encode('utf-8'))
                    return

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': f'Đăng nhập thành công với tư cách đồng chí {member.get("name")}!',
                    'member': member
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/export_docx':
            try:
                data = json.loads(body.decode('utf-8'))
                docx_path = os.path.join(BASE_DIR, 'Bien_ban_kiem_diem_2026_Chi_bo_2.docx')
                create_bien_ban_docx(data, docx_path)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': 'Đã tạo Biên bản Word thành công!',
                    'file': 'Bien_ban_kiem_diem_2026_Chi_bo_2.docx',
                    'download_url': '/api/download_docx'
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/export_evaluation_docx':
            try:
                payload = json.loads(body.decode('utf-8'))
                target = payload.get('target', 'collective')
                data = payload.get('data', payload)
                
                if target == 'all':
                    docx_name = 'Bo_Ho_So_Tu_Cham_Diem_Chi_Bo_2_2026.docx'
                    docx_path = os.path.join(BASE_DIR, docx_name)
                    create_all_evaluations_docx(data, docx_path)
                    msg = 'Đã kết xuất trọn bộ Bản tự chấm điểm Tập thể & 12 Cá nhân!'
                    download_param = 'all'
                elif target == 'collective':
                    docx_name = 'Ban_Tu_Cham_Diem_Tap_The_Chi_Bo_2_2026.docx'
                    docx_path = os.path.join(BASE_DIR, docx_name)
                    create_single_evaluation_docx(data, 'collective', docx_path)
                    msg = 'Đã kết xuất Bản tự chấm điểm Tập thể Chi bộ 2!'
                    download_param = 'collective'
                else:
                    mem_id = int(target)
                    m = next((x for x in data.get('members', []) if x.get('id') == mem_id), None)
                    docx_name = f'Ban_Tu_Cham_Diem_Dang_Vien_{mem_id}.docx'
                    docx_path = os.path.join(BASE_DIR, docx_name)
                    create_single_evaluation_docx(data, str(target), docx_path)
                    msg = f"Đã kết xuất Bản tự chấm điểm của đồng chí {m.get('name') if m else ''}!"
                    download_param = str(mem_id)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': msg,
                    'file': docx_name,
                    'download_url': f'/api/download_eval_docx?file={download_param}'
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/export_ballot_docx':
            try:
                import re, unicodedata
                def slugify(text):
                    if not text: return 'dangvien'
                    t = unicodedata.normalize('NFD', text)
                    t = re.sub(r'[\u0300-\u036f]', '', t)
                    t = t.replace('đ', 'd').replace('Đ', 'D')
                    return re.sub(r'[^a-zA-Z0-9]+', '_', t).strip('_')

                payload = json.loads(body.decode('utf-8'))
                mode = payload.get('mode', 'blank') # 'blank' or 'personal'
                voter_id = int(payload.get('voter_id', 0))
                
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    db = json.load(f)
                    
                if mode == 'personal' and voter_id > 0:
                    voter = next((m for m in db.get('members', []) if m.get('id') == voter_id), None)
                    voter_name = voter.get('name', f'DangVien_{voter_id}') if voter else f'DangVien_{voter_id}'
                    voter_choices = payload.get('choices')
                    if not voter_choices:
                        voter_choices = db.get('ballots', {}).get('voters', {}).get(str(voter_id), {}).get('ballot', None)
                        
                    docx_name = f'Phieu_Bau_Ca_Nhan_{slugify(voter_name)}.docx'
                    docx_path = os.path.join(BASE_DIR, docx_name)
                    create_ballot_docx(db, docx_path, voter_member=voter, voted_choices=voter_choices)
                    msg = f"Đã kết xuất Phiếu bầu của đồng chí {voter_name}!"
                else:
                    docx_name = 'Phieu_Bau_Kiem_Diem_Chi_Bo_2_2026.docx'
                    docx_path = os.path.join(BASE_DIR, docx_name)
                    create_ballot_docx(db, docx_path)
                    msg = "Đã kết xuất Phôi phiếu bầu Tập thể & Đảng viên chuẩn mẫu biểu Đảng!"
                    
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': msg,
                    'file': docx_name,
                    'download_url': f'/api/download_ballot_docx?file={docx_name}'
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/vote':
            try:
                from datetime import datetime
                payload = json.loads(body.decode('utf-8'))
                voter_id = int(payload.get('voter_id', 0))
                collective_vote = payload.get('collective_vote', 'HTTNV')
                member_votes = payload.get('member_votes', {})

                if not os.path.exists(DATA_FILE):
                    raise Exception("Không tìm thấy tệp dữ liệu app_data.json")
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    db = json.load(f)

                ballots = db.setdefault('ballots', {})
                if ballots.get('is_closed', False):
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Hòm phiếu đã khóa! Thời gian bỏ phiếu kín đã kết thúc.'}).encode('utf-8'))
                    return

                members = db.get('members', [])
                member = next((m for m in members if m.get('id') == voter_id), None)
                if not member:
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Không tìm thấy thông tin đảng viên!'}).encode('utf-8'))
                    return

                if member.get('status') == 'Dự bị':
                    self.send_response(403)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Theo Điều 3 Điều lệ Đảng Cộng sản Việt Nam, Đảng viên dự bị không có quyền biểu quyết/bỏ phiếu xếp loại!'}).encode('utf-8'))
                    return

                voters = ballots.setdefault('voters', {})
                voter_key = str(voter_id)
                if voter_key in voters or voter_id in voters:
                    v_info = voters.get(voter_key) or voters.get(voter_id)
                    v_time = v_info.get('voted_at', 'trước đó') if isinstance(v_info, dict) else 'trước đó'
                    self.send_response(409)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        'status': 'error',
                        'message': f'Đồng chí {member.get("name")} đã bỏ phiếu lúc {v_time}. Theo quy định, mỗi đảng viên chỉ được quyền bỏ phiếu 01 lần duy nhất!'
                    }).encode('utf-8'))
                    return

                xs_count = sum(1 for v in member_votes.values() if v == 'HTXSNV')
                if xs_count > 2:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Lá phiếu không hợp lệ: Số đồng chí chọn Xuất sắc vượt quá tỷ lệ khống chế 20% (tối đa 2 người)!'}).encode('utf-8'))
                    return

                now_str = datetime.now().strftime('%H:%M:%S %d/%m/%Y')
                voters[voter_key] = {
                    'voted_at': now_str,
                    'name': member.get('name'),
                    'role_party': member.get('role_party'),
                    'ballot': {
                        'collective_vote': collective_vote,
                        'member_votes': member_votes
                    }
                }

                coll_dict = ballots.setdefault('collective', {'HTXSNV': 0, 'HTTNV': 0, 'HTNV': 0, 'KHTNV': 0})
                if collective_vote in coll_dict:
                    coll_dict[collective_vote] = coll_dict.get(collective_vote, 0) + 1

                mem_ballots = ballots.setdefault('members', {})
                for mid_str, rank_val in member_votes.items():
                    m_dict = mem_ballots.setdefault(str(mid_str), {'HTXSNV': 0, 'HTTNV': 0, 'HTNV': 0, 'KHTNV': 0, 'final_rank': 'Hoàn thành tốt nhiệm vụ'})
                    if rank_val in m_dict:
                        m_dict[rank_val] = m_dict.get(rank_val, 0) + 1

                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(db, f, ensure_ascii=False, indent=2)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': f'Lá phiếu kín của đồng chí {member.get("name")} đã được gửi vào hòm phiếu an toàn!',
                    'voted_at': now_str,
                    'ballots': ballots
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/close_ballot':
            try:
                payload = json.loads(body.decode('utf-8'))
                user_id = int(payload.get('user_id', 0))

                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    db = json.load(f)

                member = next((m for m in db.get('members', []) if m.get('id') == user_id), None)
                if not member or member.get('role_party') not in ['Bí thư Chi bộ', 'Phó Bí thư Chi bộ']:
                    self.send_response(403)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Chỉ Bí thư / Phó Bí thư (Chủ trì kỳ họp) mới có quyền Khóa hòm phiếu và Kiểm phiếu!'}).encode('utf-8'))
                    return

                ballots = db.setdefault('ballots', {})
                ballots['is_closed'] = True

                mem_ballots = ballots.setdefault('members', {})
                # Sort members by HTXSNV desc
                scores = []
                for m in db.get('members', []):
                    mid = str(m.get('id'))
                    mb = mem_ballots.setdefault(mid, {'HTXSNV': 0, 'HTTNV': 0, 'HTNV': 0, 'KHTNV': 0, 'final_rank': 'Hoàn thành tốt nhiệm vụ'})
                    scores.append((mid, mb.get('HTXSNV', 0)))

                scores.sort(key=lambda x: x[1], reverse=True)
                # Max 2 HTXSNV
                for idx, (mid, xs_votes) in enumerate(scores):
                    if idx < 2 and xs_votes > 0:
                        mem_ballots[mid]['final_rank'] = 'Hoàn thành xuất sắc nhiệm vụ'
                    else:
                        mem_ballots[mid]['final_rank'] = 'Hoàn thành tốt nhiệm vụ'

                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(db, f, ensure_ascii=False, indent=2)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': 'Đã khóa hòm phiếu và hoàn thành kiểm phiếu tự động!',
                    'ballots': ballots
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/reset_ballot':
            try:
                payload = json.loads(body.decode('utf-8'))
                pin = str(payload.get('pin', '')).strip()
                username = str(payload.get('username', '')).strip()
                user_id = int(payload.get('user_id', 0))

                # Xác thực quyền Bí thư: nguyenthithao, user_id=6 hoặc mã PIN 686789/123456
                is_bi_thu = (username == 'nguyenthithao' or user_id == 6 or pin in ['686789', '123456'])
                if not is_bi_thu:
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'status': 'error', 'message': 'Chỉ đồng chí Bí thư Chi bộ (nguyenthithao) mới có quyền xóa toàn bộ phiếu bầu!'}).encode('utf-8'))
                    return

                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    db = json.load(f)

                db['ballots'] = {
                    'is_closed': False,
                    'voters': {},
                    'collective': {'HTXSNV': 0, 'HTTNV': 0, 'HTNV': 0, 'KHTNV': 0},
                    'members': {}
                }

                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(db, f, ensure_ascii=False, indent=2)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': 'Đã thiết lập lại hòm phiếu mới thành công!',
                    'ballots': db['ballots']
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        self.send_error(404, "API endpoint not found")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

def run_server():
    lan_ip = get_lan_ip()
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, AppRequestHandler)
    print("=" * 70)
    print(" ⭐ ỨNG DỤNG KIỂM ĐIỂM ĐẢNG VIÊN CHI BỘ 2 - VIETINBANK BẢO LỘC ⭐ ")
    print("=" * 70)
    print(f" Máy chủ nội bộ đang chạy trên:")
    print(f" 🖥️   Máy tính của bạn (Local): http://localhost:{PORT}")
    print(f" 📱  Điện thoại qua Wi-Fi (LAN): http://{lan_ip}:{PORT}")
    print(f" 🌐  Cổng chia sẻ & Quét Mã QR: http://{lan_ip}:{PORT}/share_portal.html")
    print("=" * 70)
    print(" MÃ QR CODE TRUY CẬP TỨC THÌ (Bật Camera hoặc Zalo quét trực tiếp):")
    try:
        import qrcode
        qr = qrcode.QRCode(box_size=1, border=1)
        qr.add_data(f"http://{lan_ip}:{PORT}")
        qr.make(fit=True)
        qr.print_ascii(invert=True)
    except Exception:
        pass
    print("=" * 70)
    print(" Đang tự động mở trình duyệt...")
    
    try:
        webbrowser.open(f'http://localhost:{PORT}')
    except Exception:
        pass
        
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng máy chủ.")
        httpd.server_close()

if __name__ == '__main__':
    run_server()
