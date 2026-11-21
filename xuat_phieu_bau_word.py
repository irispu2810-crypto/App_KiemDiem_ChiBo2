# -*- coding: utf-8 -*-
"""
Module xuất Phiếu bầu đánh giá, xếp loại chất lượng Tập thể & Đảng viên Chi bộ 2
Chuẩn mẫu biểu Đảng ủy và thể thức Nghị định 30/2020/NĐ-CP.
"""

import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
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

def create_ballot_docx(data, output_path, voter_member=None, voted_choices=None):
    """
    Sinh file Word Phiếu bầu đánh giá, xếp loại.
    Nếu voter_member và voted_choices được truyền vào, lá phiếu sẽ được tích sẵn kết quả của đảng viên đó.
    Nếu không, sẽ sinh phôi phiếu bầu trắng chuẩn mẫu biểu của Đảng để in ra phát tại cuộc họp.
    """
    doc = docx.Document()
    
    # Thiết lập khổ giấy A4, lề chuẩn ND30: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm
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
    normal_style.paragraph_format.space_after = Pt(4)
    normal_style.paragraph_format.space_before = Pt(0)
    
    # =========================================================================
    # PHIẾU SỐ 01: PHIẾU ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG TẬP THỂ CHI BỘ
    # =========================================================================
    tbl_hdr1 = doc.add_table(rows=1, cols=2)
    tbl_hdr1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr1.autofit = False
    tbl_hdr1.columns[0].width = Mm(75)
    tbl_hdr1.columns[1].width = Mm(90)
    
    # Cột trái: Tên cơ quan
    cl1 = tbl_hdr1.cell(0, 0)
    p_cl1 = cl1.paragraphs[0]
    p_cl1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_org1 = p_cl1.add_run("ĐẢNG BỘ NHCT VIỆT NAM\nĐẢNG BỘ CN BẢO LỘC\n")
    r_org1.font.size = Pt(11)
    r_org1.bold = True
    r_cb1 = p_cl1.add_run("CHI BỘ 2")
    r_cb1.font.size = Pt(12)
    r_cb1.bold = True
    
    # Cột phải: Đảng kỳ / Tiêu ngữ
    cr1 = tbl_hdr1.cell(0, 1)
    p_cr1 = cr1.paragraphs[0]
    p_cr1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_vn1 = p_cr1.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
    r_vn1.font.size = Pt(12)
    r_vn1.bold = True
    r_d1 = p_cr1.add_run(data.get('meeting_date', 'Bảo Lộc, ngày 21 tháng 11 năm 2026'))
    r_d1.font.size = Pt(11)
    r_d1.italic = True
    
    # Tiêu đề Phiếu 1
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_before = Pt(8)
    p_t1.paragraph_format.space_after = Pt(2)
    r_p1 = p_t1.add_run("PHIẾU ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG TỔ CHỨC ĐẢNG\n")
    r_p1.bold = True
    r_p1.font.size = Pt(13.5)
    r_p1_sub = p_t1.add_run("(Dùng cho đảng viên chính thức bỏ phiếu xếp loại Tập thể Chi bộ năm 2026)")
    r_p1_sub.font.size = Pt(11)
    r_p1_sub.italic = True

    if voter_member:
        p_voter1 = doc.add_paragraph()
        p_voter1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_voter1.paragraph_format.space_after = Pt(4)
        r_v1 = p_voter1.add_run(f"• Đảng viên bỏ phiếu: {voter_member.get('name')} - {voter_member.get('role_party', 'Đảng viên')}")
        r_v1.bold = True
        r_v1.font.size = Pt(11)
    
    # Bảng đánh giá Tập thể (1 hàng đối tượng)
    tbl_coll = doc.add_table(rows=2, cols=6)
    tbl_coll.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_coll.autofit = False
    
    widths_coll = [Mm(12), Mm(52), Mm(26), Mm(26), Mm(26), Mm(26)]
    for i, w in enumerate(widths_coll):
        tbl_coll.columns[i].width = w
        
    coll_hdrs = ["TT", "Đối tượng đánh giá", "Hoàn thành\nxuất sắc", "Hoàn thành\ntốt", "Hoàn\nthành", "Không\nhoàn thành"]
    for i, h in enumerate(coll_hdrs):
        c = tbl_coll.rows[0].cells[i]
        c.text = h
        set_cell_background(c, "F2F4F8")
        set_cell_margins(c, 80, 80, 60, 60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.size = Pt(10)
            
    # Hàng dữ liệu tập thể
    r_c_cells = tbl_coll.rows[1].cells
    r_c_cells[0].text = "1"
    r_c_cells[1].text = "Tập thể Chi bộ 2"
    
    # Lấy lựa chọn đã bầu nếu có
    coll_choice = None
    if voted_choices and 'collective_vote' in voted_choices:
        coll_choice = voted_choices['collective_vote']
        
    for i in range(2, 6):
        set_cell_margins(r_c_cells[i], 80, 80, 60, 60)
        p = r_c_cells[i].paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Kiểm tra tích ô
        val_map = {2: 'HTXSNV', 3: 'HTTNV', 4: 'HTNV', 5: 'KHTNV'}
        if coll_choice == val_map[i]:
            r = p.add_run("[ X ]")
            r.bold = True
            r.font.color.rgb = RGBColor(0, 45, 98) # Navy
        else:
            r = p.add_run("[   ]")
        r.font.size = Pt(11)
        r.font.name = 'Times New Roman'
        
    p0 = r_c_cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p0.runs:
        p0.runs[0].font.size = Pt(10)
        p0.runs[0].font.name = 'Times New Roman'
    p1 = r_c_cells[1].paragraphs[0]
    if p1.runs:
        p1.runs[0].font.size = Pt(10.5)
        p1.runs[0].font.bold = True
        p1.runs[0].font.name = 'Times New Roman'
    
    # Ghi chú hướng dẫn
    p_note1 = doc.add_paragraph()
    p_note1.paragraph_format.space_before = Pt(4)
    p_note1.paragraph_format.space_after = Pt(14)
    r_n1 = p_note1.add_run("Ghi chú: Đồng chí đánh dấu (X) vào 01 trong 04 ô tương ứng với mức xếp loại đề xuất cho Tập thể Chi bộ 2.\n" + "-"*85)
    r_n1.font.size = Pt(10)
    r_n1.italic = True
    
    # =========================================================================
    # PHIẾU SỐ 02: PHIẾU ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG ĐẢNG VIÊN
    # =========================================================================
    tbl_hdr2 = doc.add_table(rows=1, cols=2)
    tbl_hdr2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr2.autofit = False
    tbl_hdr2.columns[0].width = Mm(75)
    tbl_hdr2.columns[1].width = Mm(90)
    
    cl2 = tbl_hdr2.cell(0, 0)
    p_cl2 = cl2.paragraphs[0]
    p_cl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_org2 = p_cl2.add_run("ĐẢNG BỘ NHCT VIỆT NAM\nĐẢNG BỘ CN BẢO LỘC\n")
    r_org2.font.size = Pt(11)
    r_org2.bold = True
    r_cb2 = p_cl2.add_run("CHI BỘ 2")
    r_cb2.font.size = Pt(12)
    r_cb2.bold = True
    
    cr2 = tbl_hdr2.cell(0, 1)
    p_cr2 = cr2.paragraphs[0]
    p_cr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_vn2 = p_cr2.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
    r_vn2.font.size = Pt(12)
    r_vn2.bold = True
    r_d2 = p_cr2.add_run(data.get('meeting_date', 'Bảo Lộc, ngày 21 tháng 11 năm 2026'))
    r_d2.font.size = Pt(11)
    r_d2.italic = True
    
    # Tiêu đề Phiếu 2
    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(8)
    p_t2.paragraph_format.space_after = Pt(2)
    r_p2 = p_t2.add_run("PHIẾU ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG ĐẢNG VIÊN\n")
    r_p2.bold = True
    r_p2.font.size = Pt(13.5)
    r_p2_sub = p_t2.add_run("(Dùng cho đảng viên chính thức bỏ phiếu xếp loại đảng viên trong Chi bộ năm 2026)")
    r_p2_sub.font.size = Pt(11)
    r_p2_sub.italic = True

    if voter_member:
        p_voter2 = doc.add_paragraph()
        p_voter2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_voter2.paragraph_format.space_after = Pt(4)
        r_v2 = p_voter2.add_run(f"• Đảng viên bỏ phiếu: {voter_member.get('name')} - {voter_member.get('role_party', 'Đảng viên')}")
        r_v2.bold = True
        r_v2.font.size = Pt(11)

    # Bảng đánh giá 11 đảng viên chính thức
    members = [m for m in data.get('members', []) if m.get('status') == 'Chính thức']
    if not members:
        # Nếu chưa lọc, lấy tất cả trừ dự bị
        members = data.get('members', [])
        
    tbl_mem = doc.add_table(rows=len(members) + 1, cols=7)
    tbl_mem.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_mem.autofit = False
    
    widths_mem = [Mm(10), Mm(42), Mm(26), Mm(23), Mm(23), Mm(23), Mm(23)]
    for i, w in enumerate(widths_mem):
        tbl_mem.columns[i].width = w
        
    mem_hdrs = ["STT", "Họ và tên đảng viên", "Chức vụ", "Hoàn thành\nxuất sắc", "Hoàn thành\ntốt", "Hoàn\nthành", "Không\nhoàn thành"]
    for i, h in enumerate(mem_hdrs):
        c = tbl_mem.rows[0].cells[i]
        c.text = h
        set_cell_background(c, "F2F4F8")
        set_cell_margins(c, 80, 80, 50, 50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.size = Pt(9.5)
            
    # Lấy danh sách phiếu bầu của cá nhân nếu có
    mem_votes = voted_choices.get('member_votes', {}) if voted_choices else {}
    
    for idx, m in enumerate(members, 1):
        m_id = str(m.get('id', idx))
        row_cells = tbl_mem.rows[idx].cells
        row_cells[0].text = str(idx)
        row_cells[1].text = m.get('name', f'Đồng chí {idx}')
        row_cells[2].text = m.get('role_party', 'Đảng viên')
        
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
        
        row_cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
        row_cells[1].paragraphs[0].runs[0].font.bold = True
        
        row_cells[2].paragraphs[0].runs[0].font.size = Pt(9)
        
        # Tích chọn 4 cột đánh giá
        chosen_rank = mem_votes.get(m_id)
        val_map_mem = {3: 'HTXSNV', 4: 'HTTNV', 5: 'HTNV', 6: 'KHTNV'}
        
        for c_idx in range(3, 7):
            set_cell_margins(row_cells[c_idx], 60, 60, 40, 40)
            p = row_cells[c_idx].paragraphs[0]
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if chosen_rank == val_map_mem[c_idx]:
                r = p.add_run("[ X ]")
                r.bold = True
                r.font.color.rgb = RGBColor(0, 45, 98) # Navy
            else:
                r = p.add_run("[   ]")
            r.font.size = Pt(10)
            r.font.name = 'Times New Roman'
            
    # Ghi chú hướng dẫn bầu cử
    p_note2 = doc.add_paragraph()
    p_note2.paragraph_format.space_before = Pt(6)
    r_n2 = p_note2.add_run(
        "Ghi chú hướng dẫn bỏ phiếu:\n"
        "1. Đánh dấu (X) vào 01 trong 04 ô tương ứng với mức xếp loại đối với từng đồng chí đảng viên.\n"
        "2. Tỷ lệ xếp loại 'Hoàn thành xuất sắc nhiệm vụ' không vượt quá 20% số đảng viên hoàn thành tốt nhiệm vụ (Chi bộ 2 tối đa không quá 02 đồng chí).\n"
        "3. Đảng viên dự bị không ghi trong danh sách bỏ phiếu kín này."
    )
    r_n2.font.size = Pt(9.5)
    r_n2.italic = True
    
    doc.save(output_path)
    return output_path

if __name__ == '__main__':
    test_data = {
        'meeting_date': 'Bảo Lộc, ngày 21 tháng 11 năm 2026',
        'members': [
            {'id': 1, 'name': 'BÙI KHẮC DUY', 'role_party': 'Đảng viên', 'status': 'Chính thức'},
            {'id': 2, 'name': 'VŨ HOÀNG NGUYÊN', 'role_party': 'Đảng viên', 'status': 'Chính thức'},
            {'id': 6, 'name': 'NGUYỄN THI THAO', 'role_party': 'Bí thư Chi bộ', 'status': 'Chính thức'},
            {'id': 8, 'name': 'NGÔ THỊ KIM NGÂN', 'role_party': 'Phó Bí thư Chi bộ', 'status': 'Chính thức'},
        ]
    }
    out = os.path.join(os.path.dirname(__file__), 'test_phieu_bau.docx')
    create_ballot_docx(test_data, out)
    print("Exported test ballot to:", out)
