# -*- coding: utf-8 -*-
"""
Module xuất Bản Tự Kiểm Điểm & Tự Chấm Điểm (Tập thể & Cá nhân)
Chi bộ - Đảng bộ VietinBank Chi nhánh Bảo Lộc (Năm 2026)
Chuẩn thể thức văn bản Đảng theo Hướng dẫn 05-HD/VPTW ngày 27/5/2026
và Mẫu 02A-HD KĐ.ĐG, 02B-HD KĐ.ĐG (Quy định 366-QĐ/TW & Hướng dẫn 43-HD/BTCTW).
"""

import os
import json
import docx
from docx.shared import Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_margins(cell, top=90, bottom=90, left=140, right=140):
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

def set_table_borders(table, color="000000", sz="4"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        borders.append(border)
    tblPr.append(borders)

def clear_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        borders.append(border)
    tblPr.append(borders)

def get_criteria_for_target(target_val, data):
    """
    Lấy danh sách tiêu chí & đồng bộ 100% điểm từng tiêu chí từ appData.
    Nếu chưa lưu điểm chi tiết, tự động phân bổ tỷ lệ khớp chính xác với tổng điểm đã lưu.
    """
    if target_val == 'collective':
        col_eval = data.get('collective_evaluation', {})
        saved_scores = col_eval.get('criteria_scores', {})
        target_total = float(col_eval.get('score', 97))
        groups = [
            {
                'group': 'I. CÁC TIÊU CHÍ CHUNG (30 ĐIỂM)',
                'items': [
                    {'key': 'c_1_1', 'stt': '1', 'name': 'Công tác xây dựng, chỉnh đốn Đảng; củng cố tổ chức đảng và nâng cao chất lượng đảng viên', 'max': 5},
                    {'key': 'c_1_2', 'stt': '2', 'name': 'Năng lực lãnh đạo của cấp ủy; hiệu quả lãnh đạo chính quyền, đoàn thể và đoàn kết nội bộ', 'max': 5},
                    {'key': 'c_1_3', 'stt': '3', 'name': 'Thanh tra, kiểm tra, giám sát, kỷ luật đảng; phòng chống tham nhũng, lãng phí, tiêu cực', 'max': 5},
                    {'key': 'c_1_4', 'stt': '4', 'name': 'Nguyên tắc tập trung dân chủ, tự phê bình và phê bình; chấp hành quy chế làm việc', 'max': 5},
                    {'key': 'c_1_5', 'stt': '5', 'name': 'Lãnh đạo thực hiện nhiệm vụ chính trị, tinh gọn bộ máy, cải cách hành chính, chuyển đổi số', 'max': 5},
                    {'key': 'c_1_6', 'stt': '6', 'name': 'Tinh thần tự soi tự sửa; kết quả khắc phục hạn chế các kỳ trước', 'max': 5}
                ]
            },
            {
                'group': 'II. KẾT QUẢ THỰC HIỆN NHIỆM VỤ (70 ĐIỂM - KHỐI DOANH NGHIỆP NHÀ NƯỚC)',
                'items': [
                    {'key': 'c_2_1', 'stt': '1', 'name': 'Việc cụ thể hóa nghị quyết, chương trình, kế hoạch công tác của cấp ủy cấp trên', 'max': 10},
                    {'key': 'c_2_2', 'stt': '2', 'name': 'Kết quả lãnh đạo, chỉ đạo trên các lĩnh vực được phân công phụ trách', 'max': 10},
                    {'key': 'c_2_3', 'stt': '3', 'name': 'Kết quả thực hiện chỉ tiêu kế hoạch kinh doanh VietinBank giao năm 2026 (Huy động, Tín dụng, Thu phí, Nợ xấu)', 'max': 40},
                    {'key': 'c_2_4', 'stt': '4', 'name': 'Đánh giá, xếp loại các tổ chức trực thuộc và cá nhân liên quan trực tiếp', 'max': 10}
                ]
            }
        ]
    else:
        mem_id = int(target_val)
        m = next((x for x in data.get('members', []) if x.get('id') == mem_id), None) or {}
        saved_scores = m.get('criteria_scores', {})
        target_total = float(m.get('self_score', 95))
        title_p = (m.get('title_pro') or '').lower()
        is_leader = (
            m.get('role_party') in ['Bí thư Chi bộ', 'Phó Bí thư Chi bộ', 'Chi ủy viên']
            or any(k in title_p for k in ['giám đốc', 'trưởng phòng', 'phó phòng', 'lãnh đạo', 'quản lý'])
        )
        
        if is_leader:
            groups = [
                {
                    'group': 'I. CÁC TIÊU CHÍ CHUNG (30 ĐIỂM - CẤP ỦY VIÊN / LÃNH ĐẠO)',
                    'items': [
                        {'key': 'l_1_1', 'stt': '1', 'name': 'Chính trị, tư tưởng: Bản lĩnh chính trị vững vàng, giữ nghiêm kỷ luật phát ngôn', 'max': 5},
                        {'key': 'l_1_2', 'stt': '2', 'name': 'Phẩm chất đạo đức, lối sống, ý thức tổ chức kỷ luật: Gương mẫu, không tham nhũng tiêu cực', 'max': 5},
                        {'key': 'l_1_3', 'stt': '3', 'name': 'Năng lực lãnh đạo, quản lý; đổi mới sáng tạo, dám nghĩ dám làm vì lợi ích chung', 'max': 10},
                        {'key': 'l_1_4', 'stt': '4', 'name': 'Uy tín, khả năng quy tụ đoàn kết nội bộ và phối hợp công tác', 'max': 5},
                        {'key': 'l_1_5', 'stt': '5', 'name': 'Tinh thần tự phê bình và phê bình, tự soi tự sửa, khắc phục khuyết điểm', 'max': 5}
                    ]
                },
                {
                    'group': 'II. KẾT QUẢ THỰC HIỆN NHIỆM VỤ (70 ĐIỂM)',
                    'items': [
                        {'key': 'l_2_1', 'stt': '1', 'name': 'Lãnh đạo, điều hành chi bộ và đơn vị hoàn thành các chỉ tiêu nhiệm vụ chính trị, kinh doanh', 'max': 35},
                        {'key': 'l_2_2', 'stt': '2', 'name': 'Tiến độ và chất lượng cụ thể hóa các văn bản chỉ đạo của Đảng ủy cấp trên', 'max': 15},
                        {'key': 'l_2_3', 'stt': '3', 'name': 'Kết quả đánh giá, xếp loại đảng viên và cán bộ thuộc quyền quản lý trực tiếp', 'max': 20}
                    ]
                }
            ]
        else:
            groups = [
                {
                    'group': 'I. CÁC TIÊU CHÍ CHUNG (30 ĐIỂM - ĐẢNG VIÊN)',
                    'items': [
                        {'key': 'm_1_1', 'stt': '1', 'name': 'Chính trị, phẩm chất đạo đức và ý thức tổ chức kỷ luật (Phục vụ khách hàng, không sai sót)', 'max': 10},
                        {'key': 'm_1_2', 'stt': '2', 'name': 'Năng lực chuyên môn nghiệp vụ, thái độ công tác, đổi mới sáng tạo, văn hóa VietinBank', 'max': 15},
                        {'key': 'm_1_3', 'stt': '3', 'name': 'Tinh thần tự phê bình và phê bình, tự soi tự sửa', 'max': 5}
                    ]
                },
                {
                    'group': 'II. KẾT QUẢ THỰC HIỆN NHIỆM VỤ CHUYÊN MÔN (70 ĐIỂM)',
                    'items': [
                        {'key': 'm_2_1', 'stt': '1', 'name': 'Mức độ hoàn thành định mức KPI công việc theo bản mô tả vị trí việc làm', 'max': 40},
                        {'key': 'm_2_2', 'stt': '2', 'name': 'Tính chính xác, kịp thời, không để phát sinh rủi ro tác nghiệp hoặc ảnh hưởng uy tín chi nhánh', 'max': 20},
                        {'key': 'm_2_3', 'stt': '3', 'name': 'Tham gia học tập nghị quyết, các phong trào thi đua và nhiệm vụ đột xuất', 'max': 10}
                    ]
                }
            ]

    all_items = [it for g in groups for it in g['items']]
    if saved_scores and any(it['key'] in saved_scores for it in all_items):
        for it in all_items:
            val = float(saved_scores.get(it['key'], it['max']))
            val = max(0.0, min(val, float(it['max'])))
            it['score'] = int(val) if val.is_integer() else round(val, 1)
    else:
        for it in all_items:
            it['score'] = float(it['max'])
        diff = round(100.0 - target_total, 1)
        for it in reversed(all_items):
            if diff <= 0:
                break
            deduct = min(it['score'], diff)
            it['score'] = round(it['score'] - deduct, 1)
            diff = round(diff - deduct, 1)
        for it in all_items:
            if float(it['score']).is_integer():
                it['score'] = int(it['score'])

    return groups

def parse_date_parts(date_str):
    import re
    m = re.search(r'(\d{1,2})\s*[/\-]\s*(\d{1,2})\s*[/\-]\s*(\d{4})', str(date_str))
    if m:
        return m.group(1).zfill(2), m.group(2).zfill(2), m.group(3)
    m2 = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', str(date_str))
    if m2:
        return m2.group(1).zfill(2), m2.group(2).zfill(2), m2.group(3)
    return '28', '11', '2026'

def render_evaluation_section_in_doc(doc, data, target_val):
    meeting = data.get('meeting', {})
    day, month, year = parse_date_parts(meeting.get('date', '28/11/2026'))
    branch_name = meeting.get('branch_name', 'CHI BỘ 2').upper()
    host_name = meeting.get('host_name', 'NGUYỄN THI THAO').upper()

    is_col = (target_val == 'collective')
    if is_col:
        col_eval = data.get('collective_evaluation', {})
        ranking_val = col_eval.get('ranking', 'Hoàn thành xuất sắc nhiệm vụ')
        good_val = col_eval.get('good', 'Chi bộ luôn giữ vững bản lĩnh chính trị vững vàng, chấp hành tuyệt đối các chủ trương đường lối của Đảng, chính sách pháp luật của Nhà nước và quy chế nội bộ của VietinBank. Lãnh đạo cán bộ, đảng viên đoàn kết, nỗ lực hoàn thành thắng lợi các chỉ tiêu kinh doanh được giao năm 2026.')
        limit_val = col_eval.get('limit', 'Công tác nắm bắt diễn biến tư tưởng cán bộ đôi lúc còn chưa chủ động; việc phối hợp giải quyết thủ tục nghiệp vụ giữa các bộ phận có thời điểm chưa thật sự đồng bộ.')
        sol_val = col_eval.get('sol', 'Tiếp tục đổi mới nội dung và phương thức sinh hoạt chi bộ, tăng cường công tác kiểm tra, giám sát thường xuyên; gắn việc thực hiện nhiệm vụ Đảng với nâng cao hiệu quả kinh doanh của Chi nhánh.')
        form_code = "Mẫu 01-HD KĐ.ĐG"
        title_text = f"BẢN TỰ KIỂM ĐIỂM VÀ TỰ CHẤM ĐIỂM ĐÁNH GIÁ, XẾP LOẠI\nCHẤT LƯỢNG TẬP THỂ {branch_name} NĂM {year}"
        sub_title = "(Khối Doanh nghiệp Nhà nước — Quy định 366-QĐ/TW & Hướng dẫn 43-HD/BTCTW)"
    else:
        mem_id = int(target_val)
        m = next((x for x in data.get('members', []) if x.get('id') == mem_id), None)
        if not m:
            return
        ranking_val = m.get('self_ranking', 'Hoàn thành tốt nhiệm vụ')
        good_val = m.get('self_good', f"Đồng chí {m.get('name')} luôn có lập trường tư tưởng kiên định, chấp hành nghiêm túc sự phân công của tổ chức Đảng và chuyên môn; nỗ lực hoàn thành tốt nhiệm vụ được giao, không để xảy ra sai sót tác nghiệp.")
        limit_val = m.get('self_limit', "Tinh thần tự phê bình và phê bình trong sinh hoạt có lúc còn e dè, cần chủ động phát biểu xây dựng chi bộ nhiều hơn.")
        sol_val = m.get('self_sol', "Chủ động học tập nâng cao trình độ nghiệp vụ ngân hàng và lý luận chính trị; tích cực tham gia các phong trào thi đua của Chi nhánh.")
        
        title_p = (m.get('title_pro') or '').lower()
        is_sample_2a = (
            m.get('role_party') in ['Bí thư Chi bộ', 'Phó Bí thư Chi bộ', 'Chi ủy viên']
            or any(k in title_p for k in ['giám đốc', 'trưởng phòng', 'phó phòng', 'lãnh đạo', 'quản lý'])
        )
        form_code = "Mẫu 02A-HD KĐ.ĐG" if is_sample_2a else "Mẫu 02B-HD KĐ.ĐG"
        title_text = f"BẢN KIỂM ĐIỂM VÀ TỰ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG CÁ NHÂN\nNĂM {year}"
        sub_title = (
            f"({form_code} — Cá nhân giữ chức vụ lãnh đạo, quản lý)"
            if is_sample_2a
            else f"({form_code} — Cá nhân không giữ chức vụ lãnh đạo, quản lý)"
        )

    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_code.paragraph_format.space_after = Pt(2)
    r_code = p_code.add_run(form_code)
    r_code.bold = True
    r_code.font.size = Pt(11)

    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False
    clear_table_borders(table_header)
    table_header.columns[0].width = Mm(82)
    table_header.columns[1].width = Mm(83)
    
    cell_l = table_header.cell(0, 0)
    cell_r = table_header.cell(0, 1)
    
    p_l = cell_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l.paragraph_format.line_spacing = 1.05
    r_l1 = p_l.add_run("ĐẢNG BỘ NHCT VIỆT NAM\nĐẢNG BỘ CN BẢO LỘC\n")
    r_l1.font.size = Pt(11)
    r_l2 = p_l.add_run(f"{branch_name}\n")
    r_l2.font.size = Pt(12)
    r_l2.bold = True
    r_l3 = p_l.add_run("*")
    r_l3.font.size = Pt(11)
    r_l3.bold = True
    
    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r.paragraph_format.line_spacing = 1.05
    r_r1 = p_r.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n\n")
    r_r1.font.size = Pt(12)
    r_r1.bold = True
    r_r2 = p_r.add_run(f"Bảo Lộc, ngày {day} tháng {month} năm {year}")
    r_r2.font.size = Pt(12)
    r_r2.italic = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = 1.15
    r_t = p_title.add_run(title_text)
    r_t.bold = True
    r_t.font.size = Pt(14)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    r_sub = p_sub.add_run(sub_title)
    r_sub.italic = True
    r_sub.font.size = Pt(11.5)

    if not is_col:
        p_info = doc.add_paragraph()
        p_info.paragraph_format.line_spacing = 1.2
        p_info.paragraph_format.space_after = Pt(4)
        
        p_info.add_run("Họ và tên: ").bold = True
        r_name = p_info.add_run(f"{m.get('name').upper()}          ")
        r_name.bold = True
        p_info.add_run("Ngày sinh: ").bold = True
        p_info.add_run(f"{m.get('dob', '01/01/1985')}\n")
        
        p_info.add_run("Chức vụ Đảng: ").bold = True
        p_info.add_run(f"{m.get('role_party', 'Đảng viên')}          ")
        p_info.add_run("Chức vụ chuyên môn: ").bold = True
        p_info.add_run(f"{m.get('title_pro', 'Cán bộ')}\n")
        
        p_info.add_run("Đơn vị công tác: ").bold = True
        p_info.add_run(f"{m.get('dept', 'VietinBank Chi nhánh Bảo Lộc')}\n")
        
        p_info.add_run("Sinh hoạt tại: ").bold = True
        p_info.add_run(f"{branch_name} — Đảng bộ VietinBank Chi nhánh Bảo Lộc ({m.get('status', 'Chính thức')})")

    criteria_groups = get_criteria_for_target(target_val, data)
    computed_total = round(sum(float(it['score']) for g in criteria_groups for it in g['items']), 1)
    if computed_total.is_integer():
        computed_total = int(computed_total)

    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(6)
    p_sec1.paragraph_format.space_after = Pt(4)
    r_s1 = p_sec1.add_run(f"I. BẢNG TỰ CHẤM ĐIỂM THEO KHUNG TIÊU CHÍ NĂM {year} (THANG ĐIỂM 100)")
    r_s1.bold = True
    r_s1.font.size = Pt(12.5)

    total_rows = 1 + sum(len(g['items']) + 1 for g in criteria_groups) + 1
    table_crit = doc.add_table(rows=total_rows, cols=4)
    table_crit.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_crit.autofit = False
    set_table_borders(table_crit, color="000000", sz="4")
    
    col_widths = [Mm(12), Mm(105), Mm(23), Mm(25)]
    for row in table_crit.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    hdr_cells = table_crit.rows[0].cells
    hdr_titles = ["STT", "NỘI DUNG TIÊU CHÍ ĐÁNH GIÁ", "ĐIỂM CHUẨN", "ĐIỂM TỰ CHẤM"]
    for i, title in enumerate(hdr_titles):
        set_cell_background(hdr_cells[i], "E2E8F0")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=80, right=80)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(10.5)

    curr_row = 1
    for g in criteria_groups:
        g_cells = table_crit.rows[curr_row].cells
        g_cell = g_cells[0]
        for c in g_cells[1:]:
            g_cell.merge(c)
        set_cell_background(g_cell, "F8FAFC")
        set_cell_margins(g_cell, top=80, bottom=80, left=100, right=100)
        p = g_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(g['group'])
        r.bold = True
        r.font.size = Pt(10.5)
        curr_row += 1

        for it in g['items']:
            r_cells = table_crit.rows[curr_row].cells
            set_cell_margins(r_cells[0], top=70, bottom=70, left=70, right=70)
            set_cell_margins(r_cells[1], top=70, bottom=70, left=90, right=70)
            set_cell_margins(r_cells[2], top=70, bottom=70, left=70, right=70)
            set_cell_margins(r_cells[3], top=70, bottom=70, left=70, right=70)

            p0 = r_cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p0.add_run(it['stt']).font.size = Pt(10.5)

            p1 = r_cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p1.add_run(it['name']).font.size = Pt(10.5)

            p2 = r_cells[2].paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_max = p2.add_run(f"{it['max']}đ")
            r_max.font.size = Pt(10.5)
            r_max.bold = True

            p3 = r_cells[3].paragraphs[0]
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_self = p3.add_run(f"{it['score']}đ")
            r_self.font.size = Pt(10.5)
            r_self.bold = True

            curr_row += 1

    tot_cells = table_crit.rows[curr_row].cells
    tot_cell = tot_cells[0]
    tot_cell.merge(tot_cells[1])
    set_cell_background(tot_cell, "F1F5F9")
    set_cell_margins(tot_cell, top=100, bottom=100, left=100, right=100)
    p_tot = tot_cell.paragraphs[0]
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot = p_tot.add_run("TỔNG CỘNG ĐIỂM TỰ CHẤM (THANG ĐIỂM 100):")
    r_tot.bold = True
    r_tot.font.size = Pt(11)

    set_cell_background(tot_cells[2], "F1F5F9")
    p_m = tot_cells[2].paragraphs[0]
    p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_m = p_m.add_run("100đ")
    r_m.bold = True
    r_m.font.size = Pt(11)

    set_cell_background(tot_cells[3], "F1F5F9")
    p_s = tot_cells[3].paragraphs[0]
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s = p_s.add_run(f"{computed_total}đ")
    r_s.bold = True
    r_s.font.size = Pt(11.5)

    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(8)
    p_sec2.paragraph_format.space_after = Pt(4)
    if is_col:
        r_s2 = p_sec2.add_run("II. NỘI DUNG TỰ KIỂM ĐIỂM CỦA TẬP THỂ CHI BỘ")
    else:
        r_s2 = p_sec2.add_run(f"II. NỘI DUNG TỰ KIỂM ĐIỂM ({form_code})")
    r_s2.bold = True
    r_s2.font.size = Pt(12.5)

    if not is_col:
        p_g = doc.add_paragraph()
        p_g.paragraph_format.line_spacing = 1.2
        p_g.paragraph_format.space_after = Pt(4)
        p_g.add_run("1. Ưu điểm, kết quả đạt được:\n").bold = True
        p_g.add_run("a) Về phẩm chất chính trị, đạo đức, lối sống; tinh thần đoàn kết nội bộ; ý thức tổ chức kỷ luật, trách nhiệm nêu gương, việc thực hiện những điều đảng viên không được làm; kỷ luật phát ngôn; liên hệ biểu hiện suy thoái, \"tự diễn biến\", \"tự chuyển hóa\" và học tập, làm theo tư tưởng, đạo đức, phong cách Hồ Chí Minh:\n").bold = True
        p_g.add_run(f"   {good_val}\n   Bản thân luôn giữ vững lập trường tư tưởng chính trị, không có biểu hiện suy thoái về tư tưởng chính trị, đạo đức, lối sống, \"tự diễn biến\", \"tự chuyển hóa\".\n")
        p_g.add_run("b) Về thực hiện nhiệm vụ, quyền hạn được giao (gắn với kết quả, tiến độ, chất lượng công việc):\n").bold = True
        p_g.add_run(f"   Hoàn thành tốt các chỉ tiêu KPI và nhiệm vụ chuyên môn tại {m.get('dept', 'Chi nhánh Bảo Lộc')}, đảm bảo chất lượng, tiến độ và tuân thủ quy trình nghiệp vụ VietinBank.\n")
        p_g.add_run("c) Việc thực hiện cam kết tu dưỡng, rèn luyện, phấn đấu hằng năm:\n").bold = True
        p_g.add_run("   Thực hiện nghiêm túc bản cam kết tu dưỡng, rèn luyện, phấn đấu cá nhân đã đăng ký đầu năm.\n")
        if is_sample_2a:
            p_g.add_run("d) Kết quả công tác lãnh đạo, chỉ đạo, quản lý, điều hành; mức độ hoàn thành nhiệm vụ của tập thể, đơn vị thuộc quyền quản lý; khả năng quy tụ, xây dựng đoàn kết nội bộ:\n").bold = True
            p_g.add_run("   Phát huy vai trò nêu gương của cán bộ quản lý/cấp ủy viên, chỉ đạo điều hành sâu sát, quy tụ đoàn kết nội bộ hoàn thành tốt nhiệm vụ chính trị và kinh doanh của đơn vị.\n")
        else:
            p_g.add_run("d) Ý thức tổ chức kỷ luật, chấp hành nội quy, quy chế cơ quan, chế độ sinh hoạt Đảng, đóng đảng phí và giữ mối liên hệ nơi cư trú (Quy định 213-QĐ/TW):\n").bold = True
            p_g.add_run("   Tham gia sinh hoạt chi bộ đầy đủ, đóng đảng phí đúng quy định, chấp hành nghiêm nội quy cơ quan và giữ mối liên hệ tốt với cấp ủy nơi cư trú.\n")

        p_l = doc.add_paragraph()
        p_l.paragraph_format.line_spacing = 1.2
        p_l.paragraph_format.space_after = Pt(4)
        p_l.add_run("2. Hạn chế, khuyết điểm và nguyên nhân:\n").bold = True
        p_l.add_run("a) Hạn chế, khuyết điểm:\n").bold = True
        p_l.add_run(f"   {limit_val}\n")
        p_l.add_run("b) Nguyên nhân của hạn chế, khuyết điểm (chủ quan, khách quan):\n").bold = True
        p_l.add_run("   Do áp lực khối lượng công việc kinh doanh và tác nghiệp ngân hàng lớn vào các giai đoạn cao điểm; bản thân có lúc chưa sắp xếp thời gian nghiên cứu chuyên sâu một cách khoa học.\n")

        p_k = doc.add_paragraph()
        p_k.paragraph_format.line_spacing = 1.2
        p_k.paragraph_format.space_after = Pt(4)
        p_k.add_run("3. Kết quả khắc phục những hạn chế, khuyết điểm đã được cấp có thẩm quyền kết luận hoặc được chỉ ra ở các kỳ kiểm điểm trước:\n").bold = True
        p_k.add_run("   Đã nghiêm túc tiếp thu và từng bước khắc phục hiệu quả các tồn tại được Chi bộ góp ý từ kỳ kiểm điểm trước.\n")

        p_s = doc.add_paragraph()
        p_s.paragraph_format.line_spacing = 1.2
        p_s.paragraph_format.space_after = Pt(4)
        if is_sample_2a:
            p_s.add_run("4. Giải trình những vấn đề được gợi ý kiểm điểm (nếu có): Không có.\n").bold = True
            p_s.add_run("5. Làm rõ trách nhiệm của cá nhân đối với những hạn chế, khuyết điểm của tập thể (nếu có): Nghiêm túc nhận trách nhiệm liên đới trong phạm vi lĩnh vực được phân công phụ trách.\n").bold = True
            p_s.add_run("6. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm:\n").bold = True
        else:
            p_s.add_run("4. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm:\n").bold = True
        p_s.add_run(f"   {sol_val}")
    else:
        p_g = doc.add_paragraph()
        p_g.paragraph_format.line_spacing = 1.2
        p_g.paragraph_format.space_after = Pt(3)
        p_g.add_run("1. Ưu điểm, kết quả đạt được:\n").bold = True
        p_g.add_run(f"   {good_val}")

        p_l = doc.add_paragraph()
        p_l.paragraph_format.line_spacing = 1.2
        p_l.paragraph_format.space_after = Pt(3)
        p_l.add_run("2. Hạn chế, khuyết điểm và nguyên nhân:\n").bold = True
        p_l.add_run(f"   {limit_val}")

        p_s = doc.add_paragraph()
        p_s.paragraph_format.line_spacing = 1.2
        p_s.paragraph_format.space_after = Pt(4)
        p_s.add_run("3. Phương hướng, biện pháp khắc phục:\n").bold = True
        p_s.add_run(f"   {sol_val}")

    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(6)
    p_sec3.paragraph_format.space_after = Pt(3)
    r_s3 = p_sec3.add_run("III. TỰ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG ĐỀ XUẤT")
    r_s3.bold = True
    r_s3.font.size = Pt(12.5)

    def chk(label):
        return "☒" if ranking_val == label else "☐"

    p_rank = doc.add_paragraph()
    p_rank.paragraph_format.line_spacing = 1.2
    p_rank.paragraph_format.space_after = Pt(6)
    p_rank.add_run("• Tổng số điểm tự chấm: ").bold = True
    p_rank.add_run(f"{computed_total} / 100 điểm\n").bold = True
    p_rank.add_run("• Mức tự xếp loại chất lượng đề xuất:\n").bold = True
    p_rank.add_run(
        f"   {chk('Hoàn thành xuất sắc nhiệm vụ')} Hoàn thành xuất sắc nhiệm vụ          "
        f"{chk('Hoàn thành tốt nhiệm vụ')} Hoàn thành tốt nhiệm vụ\n"
        f"   {chk('Hoàn thành nhiệm vụ')} Hoàn thành nhiệm vụ                         "
        f"{chk('Không hoàn thành nhiệm vụ')} Không hoàn thành nhiệm vụ"
    )

    table_sign_me = doc.add_table(rows=1, cols=2)
    table_sign_me.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sign_me.autofit = False
    clear_table_borders(table_sign_me)
    table_sign_me.columns[0].width = Mm(82)
    table_sign_me.columns[1].width = Mm(83)
    
    p_sr = table_sign_me.cell(0, 1).paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sr.paragraph_format.line_spacing = 1.05

    if is_col:
        p_sr.add_run(f"T/M {branch_name}\nBÍ THƯ\n").bold = True
        r_note = p_sr.add_run("(Ký, ghi rõ họ tên)\n\n\n")
        r_note.italic = True
        r_note.font.size = Pt(11)
        p_sr.add_run(host_name).bold = True
    else:
        p_sr.add_run("NGƯỜI TỰ KIỂM ĐIỂM\n").bold = True
        r_note = p_sr.add_run("(Ký, ghi rõ họ tên)\n\n\n")
        r_note.italic = True
        r_note.font.size = Pt(11)
        p_sr.add_run(m.get('name').upper()).bold = True

        p_b = doc.add_paragraph()
        p_b.paragraph_format.space_before = Pt(10)
        p_b.paragraph_format.space_after = Pt(4)
        r_b = p_b.add_run("B. ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG CỦA CHI BỘ")
        r_b.bold = True
        r_b.font.size = Pt(12.5)

        p_b_body = doc.add_paragraph()
        p_b_body.paragraph_format.line_spacing = 1.2
        p_b_body.add_run("1. Nhận xét, đánh giá của Chi bộ:\n").bold = True
        p_b_body.add_run(f"   {m.get('host_conclusion', 'Chi bộ thống nhất thông qua bản tự kiểm điểm của đồng chí.')}\n")
        p_b_body.add_run("2. Kết quả biểu quyết của Chi bộ và đề xuất mức xếp loại: ...........................................................................")

        table_sign_cb = doc.add_table(rows=1, cols=2)
        table_sign_cb.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_sign_cb.autofit = False
        clear_table_borders(table_sign_cb)
        table_sign_cb.columns[0].width = Mm(82)
        table_sign_cb.columns[1].width = Mm(83)

        p_cb = table_sign_cb.cell(0, 1).paragraphs[0]
        p_cb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cb.paragraph_format.line_spacing = 1.05
        p_cb.add_run(f"T/M {branch_name}\nBÍ THƯ\n").bold = True
        r_cb_note = p_cb.add_run("(Ký, ghi rõ họ tên)\n\n\n")
        r_cb_note.italic = True
        r_cb_note.font.size = Pt(11)
        p_cb.add_run(host_name).bold = True

def create_single_evaluation_docx(data, target_val, output_path):
    doc = docx.Document()
    for s in doc.sections:
        s.page_width = Mm(210)
        s.page_height = Mm(297)
        s.top_margin = Mm(20)
        s.bottom_margin = Mm(20)
        s.left_margin = Mm(30)
        s.right_margin = Mm(15)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(3)

    render_evaluation_section_in_doc(doc, data, target_val)
    doc.save(output_path)
    return output_path

def create_all_evaluations_docx(data, output_path):
    doc = docx.Document()
    for s in doc.sections:
        s.page_width = Mm(210)
        s.page_height = Mm(297)
        s.top_margin = Mm(20)
        s.bottom_margin = Mm(20)
        s.left_margin = Mm(30)
        s.right_margin = Mm(15)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(3)

    render_evaluation_section_in_doc(doc, data, 'collective')

    members = data.get('members', [])
    for m in members:
        doc.add_page_break()
        render_evaluation_section_in_doc(doc, data, str(m.get('id')))

    doc.save(output_path)
    return output_path
