# -*- coding: utf-8 -*-
"""
Module xuất Biên bản kiểm điểm cuối năm Chi bộ 2 - Đảng bộ VietinBank CN Bảo Lộc.
Chuẩn thể thức Nghị định 30/2020/NĐ-CP và Hướng dẫn của Đảng ủy.
"""

import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def create_bien_ban_docx(data, output_path):
    doc = docx.Document()
    
    # 1. THIẾT LẬP KHỔ GIẤY VÀ LỀ TRANG CHUẨN NGHỊ ĐỊNH 30/2020/NĐ-CP
    # Khổ A4 (210 x 297 mm), Lề: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(6)
    
    # Helper tạo đoạn văn bản chuẩn Nghị định 30 (canh đều 2 lề, lùi đầu dòng 1cm, cách đoạn 6pt)
    def add_nd30_paragraph(text="", bold=False, italic=False, is_heading=False, is_list_item=False):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)
        if is_heading:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.first_line_indent = Inches(0.3937) # 1.0 cm
        elif is_list_item:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.3937)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.3937)
            
        if text:
            r = p.add_run(text)
            r.bold = bold
            r.italic = italic
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)
        return p

    # 2. KHỐI TIÊU NGỮ VÀ TÊN CƠ QUAN (Bảng 1 hàng 2 cột không viền)
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False
    table_header.columns[0].width = Mm(75)
    table_header.columns[1].width = Mm(90)
    
    cell_l = table_header.cell(0, 0)
    cell_r = table_header.cell(0, 1)
    
    # Cột trái: Tên cơ quan ban hành văn bản & Số ký hiệu
    # ND30: Tên cơ quan chủ quản 12-13pt in hoa đứng, Tên cơ quan 12-13pt in hoa đậm
    p_l1 = cell_l.paragraphs[0]
    p_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l1.paragraph_format.space_after = Pt(1)
    p_l1.paragraph_format.line_spacing = 1.0
    r_l1 = p_l1.add_run("ĐẢNG BỘ NHCT VIỆT NAM\nĐẢNG BỘ CN BẢO LỘC")
    r_l1.font.name = 'Times New Roman'
    r_l1.font.size = Pt(12)
    r_l1.bold = True
    
    p_l2 = cell_l.add_paragraph()
    p_l2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l2.paragraph_format.space_after = Pt(2)
    p_l2.paragraph_format.line_spacing = 1.0
    r_l2 = p_l2.add_run("CHI BỘ 2")
    r_l2.font.name = 'Times New Roman'
    r_l2.font.size = Pt(13)
    r_l2.bold = True
    
    p_l3 = cell_l.add_paragraph()
    p_l3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l3.paragraph_format.space_after = Pt(0)
    # ND30: Số: .../<tên loại viết tắt>-<tên cơ quan viết tắt>, cỡ 13pt
    r_l3 = p_l3.add_run("Số:    /BB-CB2")
    r_l3.font.name = 'Times New Roman'
    r_l3.font.size = Pt(13)
    
    # Cột phải: Quốc hiệu, Tiêu ngữ & Địa danh ngày tháng
    # ND30: Quốc hiệu in hoa đậm 12-13pt, Tiêu ngữ in thường đậm 13-14pt, Địa danh ngày tháng in nghiêng 13-14pt
    p_r1 = cell_r.paragraphs[0]
    p_r1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r1.paragraph_format.space_after = Pt(1)
    p_r1.paragraph_format.line_spacing = 1.0
    r_r1 = p_r1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r_r1.font.name = 'Times New Roman'
    r_r1.bold = True
    r_r1.font.size = Pt(12)
    
    p_r2 = cell_r.add_paragraph()
    p_r2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r2.paragraph_format.space_after = Pt(4)
    p_r2.paragraph_format.line_spacing = 1.0
    r_r2 = p_r2.add_run("Độc lập - Tự do - Hạnh phúc")
    r_r2.font.name = 'Times New Roman'
    r_r2.bold = True
    r_r2.font.size = Pt(13)
    
    p_r3 = cell_r.add_paragraph()
    p_r3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r3.paragraph_format.space_after = Pt(0)
    date_str = data.get('meeting_date', 'Bảo Lộc, ngày 21 tháng 11 năm 2026')
    r_r3 = p_r3.add_run(date_str)
    r_r3.font.name = 'Times New Roman'
    r_r3.italic = True
    r_r3.font.size = Pt(13)

    # Khoảng cách trước Tiêu đề
    p_space_title = doc.add_paragraph()
    p_space_title.paragraph_format.space_before = Pt(8)
    p_space_title.paragraph_format.space_after = Pt(4)
    
    # 3. TÊN LOẠI VÀ TRÍCH YẾU VĂN BẢN (ND30: Tên loại 13-14pt in hoa đậm canh giữa; Trích yếu 13-14pt in thường đậm)
    p_title1 = doc.add_paragraph()
    p_title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title1.paragraph_format.space_after = Pt(2)
    r_t1 = p_title1.add_run("BIÊN BẢN")
    r_t1.font.name = 'Times New Roman'
    r_t1.bold = True
    r_t1.font.size = Pt(14)
    
    p_title2 = doc.add_paragraph()
    p_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title2.paragraph_format.space_after = Pt(12)
    r_t2 = p_title2.add_run("Hội nghị kiểm điểm, đánh giá, xếp loại chất lượng\ntập thể và đảng viên năm 2026")
    r_t2.font.name = 'Times New Roman'
    r_t2.bold = True
    r_t2.font.size = Pt(14)
    
    # 4. PHẦN I: THỜI GIAN, ĐỊA ĐIỂM, THÀNH PHẦN
    p_sec1 = add_nd30_paragraph("I. THỜI GIAN, ĐỊA ĐIỂM, THÀNH PHẦN", bold=True, is_heading=True)
    
    t_time = data.get('meeting_time', '08 giờ 00 phút, ngày 21 tháng 11 năm 2026')
    t_loc = data.get('meeting_location', 'Phòng họp Tầng 3 - Trụ sở VietinBank Chi nhánh Bảo Lộc')
    t_host = data.get('host_name', 'NGUYỄN THI THAO')
    t_sec = data.get('secretary_name', 'NGÔ THỊ KIM NGÂN')
    
    add_nd30_paragraph(f"- Thời gian: Bắt đầu vào hồi {t_time}.", is_list_item=True)
    add_nd30_paragraph(f"- Địa điểm: Tại {t_loc}.", is_list_item=True)
    add_nd30_paragraph(f"- Chủ trì hội nghị: Đồng chí {t_host} - Bí thư Chi bộ 2.", is_list_item=True)
    add_nd30_paragraph(f"- Thư ký ghi biên bản: Đồng chí {t_sec} - Phó Bí thư Chi bộ 2.", is_list_item=True)
    
    attendance = data.get('attendance', {})
    total_members = len(data.get('members', [])) if data.get('members') else 12
    present_count = attendance.get('present_count', total_members)
    absent_count = attendance.get('absent_count', 0)
    official_count = attendance.get('official_count', 11)
    probation_count = attendance.get('probation_count', 1)
    
    add_nd30_paragraph(f"- Thành phần tham dự: Tổng số đảng viên của Chi bộ được triệu tập là {total_members} đồng chí (trong đó: đảng viên chính thức là {official_count} đồng chí, đảng viên dự bị là {probation_count} đồng chí).", is_list_item=True)
    add_nd30_paragraph(f"+ Số lượng đảng viên có mặt tại hội nghị: {present_count} đồng chí.", is_list_item=True)
    if absent_count > 0:
        add_nd30_paragraph(f"+ Số lượng đảng viên vắng mặt: {absent_count} đồng chí (Lý do: {attendance.get('absent_reason', 'Đi công tác hoặc nghỉ phép theo chế độ quy định')}).", is_list_item=True)
    else:
        add_nd30_paragraph(f"+ Số lượng đảng viên vắng mặt: 0 đồng chí (Có mặt đầy đủ 100%).", is_list_item=True)
    
    # 5. PHẦN II: NỘI DUNG HỘI NGHỊ
    add_nd30_paragraph("II. NỘI DUNG HỘI NGHỊ", bold=True, is_heading=True)
    
    # 1. Kiểm điểm tập thể
    add_nd30_paragraph("1. Kiểm điểm, đánh giá, xếp loại đối với tập thể Chi bộ 2 năm 2026", bold=True, is_heading=True)
    add_nd30_paragraph(f"- Đồng chí {t_host} thay mặt Ban Chi ủy trình bày dự thảo Báo cáo kiểm điểm công tác lãnh đạo, chỉ đạo của tập thể Chi bộ 2 năm 2026 theo đúng khung hướng dẫn và quy định của Đảng ủy cấp trên.", is_list_item=True)
    c_score = data.get('collective_score', '97.0')
    c_rank = data.get('collective_ranking', 'Hoàn thành xuất sắc nhiệm vụ')
    add_nd30_paragraph(f"- Kết quả tự chấm điểm của tập thể Chi bộ theo Khung tiêu chí tổ chức Đảng: Đạt {c_score}/100 điểm. Tập thể Chi bộ tự đề xuất xếp loại chất lượng: {c_rank}.", is_list_item=True)
    
    p_cd = add_nd30_paragraph("- Ý kiến thảo luận, đóng góp của các đồng chí đảng viên tham dự hội nghị đối với tập thể Chi bộ:", italic=True, is_list_item=True)
    fb_text = data.get('collective_feedback', "Hội nghị đã tiến hành thảo luận sôi nổi, dân chủ và thẳng thắn. Toàn thể 100% đảng viên nhất trí cao với nội dung Báo cáo kiểm điểm của Chi bộ, đánh giá cao sự chỉ đạo quyết liệt của Ban Chi ủy trong việc hoàn thành toàn diện các nhiệm vụ chuyên môn và công tác xây dựng Đảng năm 2026.")
    add_nd30_paragraph(fb_text)
    
    # 2. Kiểm điểm từng cá nhân đảng viên
    add_nd30_paragraph("2. Kiểm điểm, tự đánh giá, xếp loại đối với từng đảng viên trong Chi bộ", bold=True, is_heading=True)
    add_nd30_paragraph("Tại hội nghị, lần lượt từng đồng chí đảng viên đã trình bày bản tự kiểm điểm cá nhân năm 2026; các đồng chí đảng viên được phân công phát biểu nhận xét chéo; các cá nhân tiếp thu ý kiến và đồng chí Chủ trì hội nghị kết luận cụ thể như sau:")
    
    members = data.get('members', [])
    reviews = data.get('reviews', {})
    
    for idx, m in enumerate(members, 1):
        m_id = str(m.get('id', idx))
        m_name = m.get('name', f'Đồng chí {idx}')
        m_role = m.get('role_party', 'Đảng viên')
        m_title = m.get('title_pro', 'Cán bộ')
        m_dept = m.get('dept', 'Chi nhánh Bảo Lộc')
        m_type = m.get('status', 'Chính thức')
        m_score = m.get('self_score', '95')
        m_self_rank = m.get('self_ranking', 'Hoàn thành tốt nhiệm vụ')
        m_receptive = m.get('receptive_note', 'Đồng chí nghiêm túc tiếp thu toàn bộ ý kiến góp ý của các đồng chí trong Chi bộ để không ngừng rèn luyện và hoàn thiện bản thân.')
        m_host_conclusion = m.get('host_conclusion', 'Đồng chí có phẩm chất chính trị vững vàng, gương mẫu chấp hành tốt nhiệm vụ được giao, hoàn thành tốt trọng trách.')
        
        p_mem = add_nd30_paragraph(f"2.{idx}. Kiểm điểm đồng chí {m_name}", bold=True, is_heading=True)
        add_nd30_paragraph(f"Chức vụ Đảng: {m_role}; Chức vụ chuyên môn: {m_title} ({m_dept}); Phân loại đảng viên: {m_type}.", is_list_item=True)
        add_nd30_paragraph(f"+ Kết quả tự chấm điểm: {m_score}/100 điểm. Cá nhân tự nhận mức xếp loại: {m_self_rank}.", is_list_item=True)
        add_nd30_paragraph(f"+ Tóm tắt nội dung tự kiểm điểm: {m.get('self_good', 'Bản lĩnh chính trị kiên định, chấp hành nghiêm quy chế ngân hàng và nguyên tắc sinh hoạt Đảng, luôn hoàn thành tốt các chỉ tiêu nhiệm vụ được giao.')}", is_list_item=True)
        
        assigned_reviews = reviews.get(m_id, [])
        if assigned_reviews:
            add_nd30_paragraph("+ Ý kiến nhận xét, đóng góp của các đồng chí được phân công:", italic=True, is_list_item=True)
            for rev in assigned_reviews:
                reviewer_name = rev.get('reviewer_name', 'Đồng chí nhận xét')
                good = rev.get('good', 'Nhất trí với bản tự kiểm điểm của đồng chí.')
                limit = rev.get('limit', 'Cần chủ động hơn nữa trong phối hợp liên phòng ban.')
                sol = rev.get('solution', 'Tiếp tục phát huy năng lực và sở trường chuyên môn.')
                add_nd30_paragraph(f"- Đồng chí {reviewer_name} nhận xét: Ưu điểm: {good}; Hạn chế cần khắc phục: {limit}; Biện pháp đề nghị: {sol}.")
        else:
            add_nd30_paragraph("+ Các đảng viên trong Chi bộ nhất trí cao với nội dung bản tự kiểm điểm của đồng chí; không có ý kiến bổ sung thêm.", is_list_item=True)
            
        add_nd30_paragraph(f"+ Ý kiến tiếp thu của cá nhân: {m_receptive}", is_list_item=True)
        add_nd30_paragraph(f"+ Kết luận của đồng chí Chủ trì hội nghị: {m_host_conclusion}", is_list_item=True)
    
    # 6. PHẦN III: KẾT QUẢ BỎ PHIẾU BÌNH BẦU, XẾP LOẠI CHẤT LƯỢNG
    add_nd30_paragraph("III. KẾT QUẢ BỎ PHIẾU BÌNH BẦU, XẾP LOẠI CHẤT LƯỢNG", bold=True, is_heading=True)
    add_nd30_paragraph("Hội nghị tiến hành bỏ phiếu kín xếp loại chất lượng đối với Tập thể Chi bộ và từng cá nhân đảng viên theo đúng nguyên tắc Điều lệ Đảng và quy định của cấp ủy cấp trên:")
    
    # Bảng tổng hợp kiểm phiếu
    voting_results = data.get('voting_results', {})
    table_vote = doc.add_table(rows=1, cols=6)
    table_vote.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_vote.autofit = False
    
    col_widths = [Mm(12), Mm(48), Mm(25), Mm(25), Mm(25), Mm(30)]
    for ci, w in enumerate(col_widths):
        table_vote.columns[ci].width = w
        
    v_headers = ["STT", "Đối tượng đánh giá", "Hoàn thành\nxuất sắc", "Hoàn thành\ntốt", "Hoàn\nthành", "Chi bộ đề xuất\nxếp loại"]
    hdr_cells = table_vote.rows[0].cells
    for ci, text in enumerate(v_headers):
        hdr_cells[ci].text = text
        set_cell_background(hdr_cells[ci], "002D62")
        set_cell_margins(hdr_cells[ci], 100, 100, 80, 80)
        p = hdr_cells[ci].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.bold = True
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(255, 255, 255)
            
    # Hàng tập thể
    row_t = table_vote.add_row()
    c_t = row_t.cells
    c_t[0].text = "1"
    c_t[1].text = "Tập thể Chi bộ 2"
    c_t[2].text = f"{voting_results.get('collective_votes', {}).get('HTXSNV', '11/11')} (100%)"
    c_t[3].text = f"{voting_results.get('collective_votes', {}).get('HTTNV', '0/11')} (0%)"
    c_t[4].text = "0/11 (0%)"
    c_t[5].text = "Xuất sắc"
    
    for ci in range(6):
        set_cell_margins(c_t[ci], 80, 80, 80, 80)
        p = c_t[ci].paragraphs[0]
        p.runs[0].font.name = 'Times New Roman'
        p.runs[0].font.size = Pt(10)
        if ci in [0, 2, 3, 4, 5]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if ci == 1:
            p.runs[0].font.bold = True
            
    # Các hàng đảng viên
    mem_votes = voting_results.get('member_votes', {})
    for idx, m in enumerate(members, 1):
        row = table_vote.add_row()
        cells = row.cells
        m_id = str(m.get('id', idx))
        mv = mem_votes.get(m_id, {})
        
        cells[0].text = str(idx + 1)
        cells[1].text = m.get('name', '')
        cells[2].text = str(mv.get('HTXSNV', '0'))
        cells[3].text = str(mv.get('HTTNV', '11'))
        cells[4].text = str(mv.get('HTNV', '0'))
        cells[5].text = mv.get('final_rank', 'Hoàn thành tốt')
        
        for ci in range(6):
            set_cell_margins(cells[ci], 80, 80, 80, 80)
            p = cells[ci].paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.name = 'Times New Roman'
                p.runs[0].font.size = Pt(9.5)
            if ci in [0, 2, 3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
    # 7. PHẦN IV: KẾT LUẬN CỦA HỘI NGHỊ
    add_nd30_paragraph("IV. KẾT LUẬN CỦA HỘI NGHỊ", bold=True, is_heading=True)
    add_nd30_paragraph("Hội nghị đã tiến hành biểu quyết thống nhất 100% nội dung biên bản và đề nghị Ban Thường vụ Đảng ủy Ngân hàng TMCP Công thương Việt Nam - Chi nhánh Bảo Lộc xem xét, phê duyệt kết quả đánh giá, xếp loại năm 2026 của tập thể và các cá nhân thuộc Chi bộ 2.")
    add_nd30_paragraph("Biên bản này được lập thành 02 bản có giá trị pháp lý như nhau, đã được thông qua trước toàn thể hội nghị vào hồi 11 giờ 30 phút cùng ngày và được toàn thể đảng viên dự họp nhất trí tuyệt đối biểu quyết thông qua./.")
    
    # 8. KHỐI CHỮ KÝ ĐÔI (ND30: Biên bản có 2 bên ký: Thư ký & Chủ trì cuộc họp)
    p_space_sign = doc.add_paragraph()
    p_space_sign.paragraph_format.space_before = Pt(12)
    p_space_sign.paragraph_format.space_after = Pt(4)
    
    table_sign = doc.add_table(rows=1, cols=2)
    table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sign.autofit = False
    table_sign.columns[0].width = Mm(80)
    table_sign.columns[1].width = Mm(85)
    
    c_s_left = table_sign.cell(0, 0)
    c_s_right = table_sign.cell(0, 1)
    
    # Bên trái: THƯ KÝ HỘI NGHỊ
    p_sl = c_s_left.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sl.paragraph_format.space_after = Pt(2)
    p_sl.paragraph_format.line_spacing = 1.0
    r_sl1 = p_sl.add_run("THƯ KÝ HỘI NGHỊ")
    r_sl1.font.name = 'Times New Roman'
    r_sl1.bold = True
    r_sl1.font.size = Pt(13)
    
    p_sl_sub = c_s_left.add_paragraph()
    p_sl_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sl_sub.paragraph_format.space_after = Pt(44) # Chừa khoảng trống ký tên
    p_sl_sub.paragraph_format.line_spacing = 1.0
    r_sl_sub = p_sl_sub.add_run("(Ký, ghi rõ họ tên)")
    r_sl_sub.font.name = 'Times New Roman'
    r_sl_sub.italic = True
    r_sl_sub.font.size = Pt(12)
    
    p_sl_name = c_s_left.add_paragraph()
    p_sl_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sl_name.paragraph_format.space_after = Pt(0)
    r_sl2 = p_sl_name.add_run(t_sec.upper())
    r_sl2.font.name = 'Times New Roman'
    r_sl2.bold = True
    r_sl2.font.size = Pt(13)
    
    # Bên phải: T/M CHI BỘ 2 - BÍ THƯ
    p_sr = c_s_right.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sr.paragraph_format.space_after = Pt(2)
    p_sr.paragraph_format.line_spacing = 1.0
    r_sr1 = p_sr.add_run("T/M CHI BỘ 2\nBÍ THƯ")
    r_sr1.font.name = 'Times New Roman'
    r_sr1.bold = True
    r_sr1.font.size = Pt(13)
    
    p_sr_sub = c_s_right.add_paragraph()
    p_sr_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sr_sub.paragraph_format.space_after = Pt(44) # Chừa khoảng trống ký tên
    p_sr_sub.paragraph_format.line_spacing = 1.0
    r_sr_sub = p_sr_sub.add_run("(Ký, ghi rõ họ tên)")
    r_sr_sub.font.name = 'Times New Roman'
    r_sr_sub.italic = True
    r_sr_sub.font.size = Pt(12)
    
    p_sr_name = c_s_right.add_paragraph()
    p_sr_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sr_name.paragraph_format.space_after = Pt(0)
    r_sr2 = p_sr_name.add_run(t_host.upper())
    r_sr2.font.name = 'Times New Roman'
    r_sr2.bold = True
    r_sr2.font.size = Pt(13)
    
    doc.save(output_path)
    return output_path

if __name__ == '__main__':
    test_data = {
        'meeting_date': 'Bảo Lộc, ngày 21 tháng 11 năm 2026',
        'host_name': 'NGUYỄN THI THAO',
        'secretary_name': 'NGÔ THỊ KIM NGÂN',
        'collective_score': '97.0',
        'collective_ranking': 'Hoàn thành xuất sắc nhiệm vụ',
        'members': [
            {'id': 1, 'name': 'NGUYỄN THI THAO', 'role_party': 'Bí thư Chi bộ', 'title_pro': 'Lãnh đạo quản lý', 'dept': 'Ban Chi ủy Chi bộ 2', 'status': 'Chính thức', 'self_score': 98, 'self_ranking': 'Hoàn thành xuất sắc nhiệm vụ'},
            {'id': 2, 'name': 'NGÔ THỊ KIM NGÂN', 'role_party': 'Phó Bí thư Chi bộ', 'title_pro': 'Lãnh đạo quản lý', 'dept': 'Ban Chi ủy Chi bộ 2', 'status': 'Chính thức', 'self_score': 97, 'self_ranking': 'Hoàn thành xuất sắc nhiệm vụ'},
            {'id': 3, 'name': 'VÕ VĂN MẠNH', 'role_party': 'Chi ủy viên', 'title_pro': 'Lãnh đạo quản lý', 'dept': 'Ban Chi ủy Chi bộ 2', 'status': 'Chính thức', 'self_score': 96, 'self_ranking': 'Hoàn thành tốt nhiệm vụ'}
        ]
    }
    out = os.path.join(os.path.dirname(__file__), 'test_bien_ban.docx')
    create_bien_ban_docx(test_data, out)
    print("Test docx generated successfully at:", out)
