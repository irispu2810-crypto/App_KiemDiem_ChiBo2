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
                    {'key': 'c_1_1', 'stt': '1', 'name': 'Kết quả công tác xây dựng, chỉnh đốn Đảng; củng cố, xây dựng tổ chức đảng và hệ thống chính trị; nâng cao chất lượng đội ngũ cán bộ, đảng viên theo nghị quyết, kế hoạch, chỉ tiêu được giao.', 'max': 5},
                    {'key': 'c_1_2', 'stt': '2', 'name': 'Năng lực lãnh đạo của cấp ủy; hiệu quả lãnh đạo, chỉ đạo của cấp ủy đối với chính quyền, Mặt trận Tổ quốc và các tổ chức chính trị - xã hội; hiệu quả trong phối hợp công tác; xây dựng đoàn kết nội bộ.', 'max': 5},
                    {'key': 'c_1_3', 'stt': '3', 'name': 'Hiệu quả trong công tác thanh tra, kiểm tra, giám sát và kỷ luật đảng; thực hiện các quy định về kiểm soát quyền lực; không để xảy ra tham nhũng, lãng phí, tiêu cực, suy thoái về tư tưởng chính trị, đạo đức, lối sống, "tự diễn biến", "tự chuyển hoá" trong nội bộ.', 'max': 5},
                    {'key': 'c_1_4', 'stt': '4', 'name': 'Kết quả thực hiện các nguyên tắc tổ chức và hoạt động của Đảng, nhất là nguyên tắc tập trung dân chủ, tự phê bình và phê bình; việc chấp hành quy chế làm việc, nội quy, chế độ công tác; phân công nhiệm vụ, tổ chức thực hiện hiệu quả.', 'max': 5},
                    {'key': 'c_1_5', 'stt': '5', 'name': 'Kết quả lãnh đạo, chỉ đạo, kiểm tra, giám sát việc thực hiện nhiệm vụ chính trị; xây dựng bộ máy tinh gọn, tinh giản biên chế, hoạt động hiệu năng, hiệu lực, hiệu quả; thực hiện cải cách hành chính, chuyển đổi số, đẩy mạnh ứng dụng khoa học, công nghệ, phương pháp quản lý gắn với xây dựng văn hoá công vụ trong hoạt động quản lý, điều hành tổ chức, cơ quan, đơn vị.', 'max': 5},
                    {'key': 'c_1_6', 'stt': '6', 'name': 'Tinh thần tự phê bình và phê bình, tự soi, tự sửa của tập thể; kết quả khắc phục những hạn chế, yếu kém, khuyết điểm đã được chỉ ra trong các kỳ kiểm điểm trước hoặc qua thanh tra, kiểm tra, giám sát.', 'max': 5}
                ]
            },
            {
                'group': 'II. NHÓM TIÊU CHÍ VỀ KẾT QUẢ THỰC HIỆN NHIỆM VỤ — MỤC 9: KHỐI DOANH NGHIỆP NHÀ NƯỚC (70 ĐIỂM)',
                'items': [
                    {'is_header': True, 'stt': '9', 'name': 'Khối doanh nghiệp nhà nước (Ghi chú: Mỗi nội dung, nhiệm vụ trễ hạn, không đạt chỉ tiêu, kế hoạch giao hoặc bị cấp có thẩm quyền nhắc nhở, phê bình bị trừ 2 điểm.)', 'max': 70},
                    {'key': 'c_2_1', 'stt': '(1)', 'name': 'Việc cụ thể hoá nghị quyết, chương trình, kế hoạch công tác của cấp ủy hoặc cấp có thẩm quyền.', 'max': 10},
                    {'key': 'c_2_2', 'stt': '(2)', 'name': 'Kết quả thực hiện nhiệm vụ được giao (nếu có); kết quả sản xuất kinh doanh; về quản lý tài chính, tài sản; về quản trị doanh nghiệp, tổ chức bộ máy, nhân sự; xây dựng văn hoá doanh nghiệp; việc chăm lo đời sống người lao động; phòng, chống tham nhũng, lãng phí, tiêu cực... theo chức năng, nhiệm vụ được giao.', 'max': 50},
                    {'key': 'c_2_3', 'stt': '(3)', 'name': 'Kết quả đánh giá, xếp loại đối với các lĩnh vực, địa phương, cơ quan, đơn vị do tập thể phụ trách (nếu có); kết quả đánh giá chất lượng đối với các tổ chức trực thuộc và cá nhân liên quan trực tiếp.', 'max': 10}
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
                    'group': 'I. CÁC TIÊU CHÍ CHUNG — PHỤ LỤC 4: CÁ NHÂN GIỮ CHỨC VỤ LÃNH ĐẠO, QUẢN LÝ (30 ĐIỂM)',
                    'items': [
                        {'is_header': True, 'stt': '1', 'name': 'Về chính trị, tư tưởng', 'max': 5},
                        {'key': 'l_1_1a', 'stt': 'a)', 'name': 'Tuyệt đối trung thành với Đảng, với Tổ quốc và Nhân dân; kiên định lý tưởng cách mạng, chủ nghĩa Mác - Lênin, tư tưởng Hồ Chí Minh, mục tiêu độc lập dân tộc, chủ nghĩa xã hội và đường lối đổi mới của Đảng.', 'max': 1},
                        {'key': 'l_1_1b', 'stt': 'b)', 'name': 'Có lập trường, quan điểm, bản lĩnh chính trị vững vàng, không dao động trước mọi khó khăn, thách thức; kiên quyết bảo vệ nền tảng tư tưởng, Cương lĩnh, đường lối của Đảng, Hiến pháp và pháp luật của Nhà nước; nhận thức và hành động thống nhất về tư tưởng chính trị, tổ chức và đạo đức; đấu tranh phản bác các quan điểm sai trái, luận điệu xuyên tạc của các thế lực thù địch. Giữ nghiêm kỷ luật phát ngôn theo đúng nguyên tắc, quy định của Đảng.', 'max': 1},
                        {'key': 'l_1_1c', 'stt': 'c)', 'name': 'Có tinh thần yêu nước sâu sắc, tận tuỵ phục vụ Nhân dân, sâu sát cơ sở, luôn hành động vì lợi ích của Nhân dân, đặt lợi ích của Đảng, quốc gia, dân tộc, Nhân dân, tập thể lên trên lợi ích cá nhân; sẵn sàng hy sinh vì sự nghiệp của Đảng, vì độc lập, tự do của Tổ quốc, vì hạnh phúc của Nhân dân.', 'max': 1},
                        {'key': 'l_1_1d', 'stt': 'd)', 'name': 'Tuyệt đối chấp hành sự phân công của tổ chức, yên tâm công tác và hoàn thành tốt mọi nhiệm vụ được giao.', 'max': 1},
                        {'key': 'l_1_1e', 'stt': 'đ)', 'name': 'Tích cực nghiên cứu, học tập các chủ trương, nghị quyết của Đảng, chính sách, pháp luật của Nhà nước; thường xuyên bồi dưỡng, cập nhật kiến thức mới và vận dụng sáng tạo vào thực tiễn.', 'max': 0.5},
                        {'key': 'l_1_1f', 'stt': 'e)', 'name': 'Có năng lực tư duy và tầm nhìn đáp ứng với yêu cầu thay đổi của tình hình thực tiễn; phát huy tinh thần chủ động, đổi mới, sáng tạo; phấn đấu vì mục tiêu phát triển của tổ chức, cơ quan, đơn vị, đóng góp vào mục tiêu chung của đất nước.', 'max': 0.5},

                        {'is_header': True, 'stt': '2', 'name': 'Về phẩm chất đạo đức và ý thức tổ chức kỷ luật', 'max': 5},
                        {'key': 'l_1_2a', 'stt': 'a)', 'name': 'Có phẩm chất đạo đức, lối sống trong sáng, trung thực, khiêm tốn, chân thành, giản dị; cần, kiệm, liêm, chính, chí công vô tư; chấp hành nghiêm quy định về trách nhiệm nêu gương; không vi phạm Quy định về những điều đảng viên không được làm. Không vi phạm đạo đức, lối sống đến mức bị xử lý kỷ luật trong 5 năm gần nhất.', 'max': 1},
                        {'key': 'l_1_2b', 'stt': 'b)', 'name': 'Không tham ô, tham nhũng, tiêu cực, quan liêu, hách dịch, cửa quyền, vụ lợi; không để người thân, người quen lợi dụng chức vụ, quyền hạn của mình để trục lợi. Không có biểu hiện suy thoái về tư tưởng chính trị, đạo đức, lối sống, những biểu hiện "tự diễn biến", "tự chuyển hoá" trong nội bộ. Kiên quyết đấu tranh chống quan liêu, cửa quyền, tham nhũng, xa hoa, lãng phí, tiêu cực, chủ nghĩa cá nhân, lối sống cơ hội, thực dụng, bè phái, lợi ích nhóm, nói không đi đôi với làm.', 'max': 1},
                        {'key': 'l_1_2c', 'stt': 'c)', 'name': 'Có ý thức tự giác học tập, tu dưỡng, rèn luyện đạo đức bản thân; có tinh thần, trách nhiệm cao với công việc; cần cù, chịu khó, năng động, sáng tạo, dám nghĩ, dám nói, dám làm, dám chịu trách nhiệm, dám đương đầu với khó khăn, thử thách, dám hy sinh vì lợi ích chung.', 'max': 1},
                        {'key': 'l_1_2d', 'stt': 'd)', 'name': 'Thực hiện nghiêm các nguyên tắc tổ chức và hoạt động của Đảng, nhất là nguyên tắc tập trung dân chủ, tự phê bình và phê bình; thực hiện đúng, đầy đủ quyền hạn, trách nhiệm và chỉ đạo thực hiện các quy định, quy chế, nội quy của tổ chức, cơ quan, đơn vị nơi công tác; kiên quyết đấu tranh với những biểu hiện, việc làm trái với các quy định của Đảng, Nhà nước trong công tác cán bộ.', 'max': 1},
                        {'key': 'l_1_2e', 'stt': 'đ)', 'name': 'Thực hiện việc kê khai và công khai tài sản, thu nhập theo quy định.', 'max': 0.5},
                        {'key': 'l_1_2f', 'stt': 'e)', 'name': 'Báo cáo đầy đủ, trung thực, cung cấp thông tin chính xác, khách quan về những nội dung liên quan đến việc thực hiện chức trách, nhiệm vụ được giao và hoạt động của tổ chức, cơ quan, đơn vị với cấp trên khi được yêu cầu.', 'max': 0.5},

                        {'is_header': True, 'stt': '3', 'name': 'Năng lực lãnh đạo, quản lý và năng lực chuyên môn, nghiệp vụ theo yêu cầu của vị trí việc làm; khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao; thái độ công tác trong thực hiện nhiệm vụ; tinh thần đổi mới sáng tạo, dám nghĩ, dám làm, dám chịu trách nhiệm vì lợi ích chung', 'max': 10},
                        {'is_header': True, 'stt': 'a)', 'name': 'Năng lực lãnh đạo, quản lý', 'max': 2},
                        {'key': 'l_1_3a1', 'stt': '-', 'name': 'Có tư duy, khả năng hoạch định đường lối, chính sách và lãnh đạo, chỉ đạo tổ chức thực hiện đóng góp vào mục tiêu phát triển chung của địa phương, tổ chức, cơ quan, đơn vị; lãnh đạo, chỉ đạo xây dựng, củng cố, kiện toàn tổ chức đảng, đoàn thể trong sạch, vững mạnh.', 'max': 0.5},
                        {'key': 'l_1_3a2', 'stt': '-', 'name': 'Có tầm nhìn, phương pháp làm việc khoa học; có năng lực tổng hợp, phân tích và dự báo.', 'max': 0.5},
                        {'key': 'l_1_3a3', 'stt': '-', 'name': 'Có năng lực, kinh nghiệm thực tiễn để cụ thể hoá và tổ chức thực hiện có hiệu quả đường lối, chủ trương của Đảng, chính sách, pháp luật của Nhà nước ở lĩnh vực, địa bàn công tác được phân công phụ trách.', 'max': 0.5},
                        {'key': 'l_1_3a4', 'stt': '-', 'name': 'Có khả năng chỉ đạo, điều hành hoạt động của tổ chức, cơ quan, đơn vị, phân công công việc khoa học, giám sát chặt chẽ, giữ gìn kỷ cương, kỷ luật.', 'max': 0.5},
                        {'is_header': True, 'stt': 'b)', 'name': 'Năng lực chuyên môn, nghiệp vụ theo vị trí việc làm', 'max': 2},
                        {'key': 'l_1_3b1', 'stt': '-', 'name': 'Có kiến thức chuyên sâu về lĩnh vực công tác được phân công; am hiểu quy định pháp luật, quy trình nghiệp vụ có liên quan đến vị trí việc làm.', 'max': 1},
                        {'key': 'l_1_3b2', 'stt': '-', 'name': 'Có khả năng phát hiện các vấn đề mới, khó và những hạn chế, bất cập trong thực tiễn; chủ động đề xuất chủ trương, chính sách, nhiệm vụ, giải pháp phù hợp, khả thi, hiệu quả để tháo gỡ điểm nghẽn, thúc đẩy phát triển.', 'max': 0.5},
                        {'key': 'l_1_3b3', 'stt': '-', 'name': 'Có kỹ năng xử lý công việc độc lập, làm việc nhóm hiệu quả.', 'max': 0.5},
                        {'is_header': True, 'stt': 'c)', 'name': 'Khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao', 'max': 2},
                        {'key': 'l_1_3c1', 'stt': '-', 'name': 'Nhiệm vụ thường xuyên: Vận dụng hiệu quả kiến thức chuyên môn, nghiệp vụ để xử lý công việc chuyên môn theo kế hoạch định kỳ; duy trì ổn định chất lượng chuyên môn.', 'max': 1},
                        {'key': 'l_1_3c2', 'stt': '-', 'name': 'Nhiệm vụ đột xuất: Chủ động đề xuất giải pháp, thực hiện hiệu quả các công việc phát sinh; có khả năng phản ứng kịp thời, đáp ứng với yêu cầu, nhiệm vụ cụ thể.', 'max': 1},
                        {'is_header': True, 'stt': 'd)', 'name': 'Thái độ công tác trong thực hiện nhiệm vụ', 'max': 2},
                        {'key': 'l_1_3d1', 'stt': '-', 'name': 'Tinh thần trách nhiệm cao trong công việc; chủ động tiếp cận kiến thức mới để kịp thời điều chỉnh, đề xuất cải tiến quy trình hoặc giải pháp nâng cao hiệu quả công việc.', 'max': 1},
                        {'key': 'l_1_3d2', 'stt': '-', 'name': 'Có thái độ đúng mực, phong cách làm việc chuẩn mực, lề lối hành chính chuyên nghiệp trong quan hệ công tác.', 'max': 0.5},
                        {'key': 'l_1_3d3', 'stt': '-', 'name': 'Phối hợp có hiệu quả với tập thể, cá nhân có liên quan trong việc triển khai thực hiện các nhiệm vụ được phân công.', 'max': 0.5},
                        {'is_header': True, 'stt': 'đ)', 'name': 'Tinh thần đổi mới sáng tạo, dám nghĩ, dám làm, dám chịu trách nhiệm vì lợi ích chung', 'max': 2},
                        {'key': 'l_1_3e1', 'stt': '-', 'name': 'Có sáng kiến, mô hình cách làm mới hoặc đề xuất giải pháp hiệu quả, mang tính đột phá hoặc tạo chuyển biến, được áp dụng trong thực tiễn thuộc phạm vi, lĩnh vực phụ trách và được cấp có thẩm quyền ghi nhận, đánh giá cao.', 'max': 0.5},
                        {'key': 'l_1_3e2', 'stt': '-', 'name': 'Sẵn sàng tham gia thực hiện nhiệm vụ chính trị đặc biệt quan trọng, nhiệm vụ có tính chất đột xuất, phức tạp hoặc trong điều kiện khó khăn.', 'max': 0.5},
                        {'key': 'l_1_3e3', 'stt': '-', 'name': 'Có tinh thần chịu trách nhiệm trước kết quả công việc; chủ động nhận trách nhiệm khi có sai sót và có biện pháp khắc phục rõ ràng, cụ thể.', 'max': 0.5},
                        {'key': 'l_1_3e4', 'stt': '-', 'name': 'Quyết đoán, kịp thời đưa ra quyết định trong phạm vi thẩm quyền, không né tránh; có tinh thần tiên phong trong thực hiện những nhiệm vụ mới.', 'max': 0.5},

                        {'is_header': True, 'stt': '4', 'name': 'Về mức độ tín nhiệm, uy tín và khả năng quy tụ đoàn kết', 'max': 5},
                        {'key': 'l_1_4a', 'stt': 'a)', 'name': 'Có uy tín trong nội bộ, gương mẫu, gắn bó mật thiết với Nhân dân; được cán bộ, đảng viên, quần chúng nhân dân ở địa phương, tổ chức, cơ quan, đơn vị tin tưởng, tín nhiệm cao.', 'max': 2.5},
                        {'key': 'l_1_4b', 'stt': 'b)', 'name': 'Có khả năng quy tụ, đoàn kết nội bộ; xây dựng tập thể vững mạnh về chính trị, tư tưởng, tổ chức và đạo đức; phát huy được sức mạnh tập thể; xây dựng đội ngũ cán bộ, công chức đoàn kết, chuyên nghiệp, đáp ứng yêu cầu, nhiệm vụ.', 'max': 2.5},

                        {'is_header': True, 'stt': '5', 'name': 'Về tự phê bình và phê bình, tự soi, tự sửa, khắc phục hạn chế, khuyết điểm', 'max': 5},
                        {'key': 'l_1_5', 'stt': '-', 'name': 'Tinh thần tự phê bình, tự soi, tự sửa của cá nhân lãnh đạo, quản lý; sự chủ động nhận diện thiếu sót trong công tác lãnh đạo, chỉ đạo; kết quả khắc phục những hạn chế, khuyết điểm đã được chỉ ra của bản thân và trong phạm vi lãnh đạo, quản lý.', 'max': 5}
                    ]
                },
                {
                    'group': 'II. NHÓM TIÊU CHÍ VỀ KẾT QUẢ THỰC HIỆN CHỨC TRÁCH, NHIỆM VỤ ĐƯỢC GIAO — MỤC 8: KHỐI DOANH NGHIỆP NHÀ NƯỚC (70 ĐIỂM)',
                    'items': [
                        {'is_header': True, 'stt': '8', 'name': 'Khối doanh nghiệp nhà nước', 'max': 70},
                        {'is_header': True, 'stt': '8.1.', 'name': 'Kết quả thực hiện chức trách, nhiệm vụ được giao', 'max': 40},
                        {'key': 'l_2_1a', 'stt': 'a)', 'name': 'Các báo cáo phân tích, đánh giá hiệu quả đầu tư và kinh doanh được thực hiện chính xác, đầy đủ. Chiến lược phát triển doanh nghiệp rõ ràng, khả thi, được triển khai hiệu quả. Đội ngũ nhân sự được xây dựng theo hướng chuyên nghiệp, đáp ứng tốt yêu cầu hoạt động, nhận phản hồi tích cực từ nội bộ và đối tác.', 'max': 10},
                        {'key': 'l_2_1b', 'stt': 'b)', 'name': 'Dự báo tình hình kinh doanh mang tính thực tiễn, được áp dụng hiệu quả trong kế hoạch hoạt động. Các vấn đề phức tạp, tranh chấp, khiếu nại được giải quyết kịp thời, đúng quy định pháp luật, nhận được sự đồng thuận từ các bên liên quan. Các giải pháp đề xuất có tác dụng rõ rệt nâng cao hiệu quả hoạt động của doanh nghiệp.', 'max': 10},
                        {'key': 'l_2_1c', 'stt': 'c)', 'name': 'Các bộ phận trong doanh nghiệp phối hợp đồng bộ, chặt chẽ. Quan hệ với cơ quan nhà nước được duy trì minh bạch, tuân thủ quy định, nhận được phản hồi tích cực; các báo cáo, hồ sơ phối hợp được hoàn thành đầy đủ, đúng thời hạn, không có sai sót.', 'max': 10},
                        {'key': 'l_2_1d', 'stt': 'd)', 'name': 'Mức độ hoàn thành kế hoạch sản xuất - kinh doanh (doanh thu, lợi nhuận, cổ tức, bảo toàn và phát triển vốn nhà nước). Doanh nghiệp nộp ngân sách đúng, đủ và đúng hạn. Không có khoản bị truy thu hoặc bị xử phạt hành chính do gian lận; không để xảy ra tình trạng thua lỗ kéo dài, mất vốn nhà nước, mất khả năng thanh toán. Tỉ lệ giải quyết kiến nghị, khiếu nại nội bộ và từ người lao động bảo đảm theo quy định.', 'max': 10},

                        {'is_header': True, 'stt': '8.2.', 'name': 'Kết quả lãnh đạo, chỉ đạo, điều hành, tổ chức thực hiện nhiệm vụ trong phạm vi phụ trách', 'max': 20},
                        {'key': 'l_2_2a', 'stt': 'a)', 'name': 'Thiết lập hệ thống quản trị hiện đại: Quy trình vận hành rõ ràng, có thiết lập cơ chế kiểm soát nội bộ, góp phần giảm chi phí hoạt động cho doanh nghiệp.', 'max': 10},
                        {'key': 'l_2_2b', 'stt': 'b)', 'name': 'Chủ động tái cấu trúc, cổ phần hoá đúng lộ trình, không gây xáo trộn lớn; phát hiện rủi ro sớm và có biện pháp ứng phó, khắc phục phù hợp.', 'max': 10},

                        {'key': 'l_2_3', 'stt': '8.3.', 'name': 'Kết quả đánh giá, xếp loại chất lượng của tập thể cơ quan, đơn vị và của đội ngũ cán bộ, đảng viên thuộc quyền quản lý trực tiếp của cá nhân (nếu có). (Nếu trường hợp không có nội dung này thì cộng 10 điểm vào mục d (20 điểm), 8.1 (50 điểm)).', 'max': 10}
                    ]
                }
            ]
        else:
            groups = [
                {
                    'group': 'I. CÁC TIÊU CHÍ CHUNG — PHỤ LỤC 6: CÁ NHÂN KHÔNG GIỮ CHỨC VỤ LÃNH ĐẠO, QUẢN LÝ (30 ĐIỂM)',
                    'items': [
                        {'is_header': True, 'stt': '1', 'name': 'Về chính trị, phẩm chất đạo đức và ý thức tổ chức kỷ luật', 'max': 10},
                        {'key': 'm_1_1a', 'stt': 'a)', 'name': 'Có quan điểm, bản lĩnh chính trị vững vàng; kiên định lập trường; không dao động trước mọi khó khăn, thách thức.', 'max': 1},
                        {'key': 'm_1_1b', 'stt': 'b)', 'name': 'Thực hiện nghiêm các nguyên tắc tổ chức và hoạt động của Đảng, nhất là nguyên tắc tập trung dân chủ, tự phê bình và phê bình. Chấp hành đường lối, chủ trương của Đảng, chính sách, pháp luật của Nhà nước; thực hiện nghiêm về kỷ luật phát ngôn, bảo vệ bí mật nhà nước.', 'max': 1},
                        {'key': 'm_1_1c', 'stt': 'c)', 'name': 'Có ý thức nghiên cứu, học tập, vận dụng chủ nghĩa Mác - Lênin, tư tưởng Hồ Chí Minh, nghị quyết, chỉ thị, quyết định và các văn bản của Đảng.', 'max': 1},
                        {'key': 'm_1_1d', 'stt': 'd)', 'name': 'Giữ gìn phẩm chất đạo đức, lối sống trong sáng, trung thực, khiêm tốn, chân thành, giản dị; có ý thức tham gia công tác đấu tranh phòng, chống tham nhũng, lãng phí, tiêu cực; không có biểu hiện suy thoái về tư tưởng chính trị, đạo đức, lối sống, "tự diễn biến", "tự chuyển hoá"; không vi phạm Quy định về những điều đảng viên không được làm.', 'max': 1},
                        {'key': 'm_1_1e', 'stt': 'đ)', 'name': 'Không tham ô, tham nhũng, tiêu cực, quan liêu, hách dịch, cửa quyền, vụ lợi; không để người thân, người quen lợi dụng quyền hạn của mình để trục lợi.', 'max': 1},
                        {'key': 'm_1_1f', 'stt': 'e)', 'name': 'Có ý thức tổ chức kỷ luật, tinh thần trách nhiệm trong công tác; chấp hành sự phân công của tổ chức; thực hiện các quy định, quy chế, nội quy của tổ chức, cơ quan, đơn vị nơi công tác.', 'max': 1},
                        {'key': 'm_1_1g', 'stt': 'g)', 'name': 'Thực hiện việc kê khai và công khai tài sản, thu nhập theo quy định.', 'max': 1},
                        {'key': 'm_1_1h', 'stt': 'h)', 'name': 'Báo cáo đầy đủ, trung thực, cung cấp thông tin chính xác, khách quan về những nội dung liên quan đến việc thực hiện chức trách, nhiệm vụ được giao và hoạt động của tổ chức, cơ quan, đơn vị với cấp trên khi được yêu cầu.', 'max': 1},
                        {'key': 'm_1_1i', 'stt': 'i)', 'name': 'Giữ gìn đoàn kết nội bộ; có quan hệ tốt với đồng chí, đồng nghiệp; tích cực tham gia xây dựng tổ chức đảng, đoàn thể và các hoạt động xã hội, cộng đồng và phong trào tập thể do tổ chức, cơ quan, đơn vị tổ chức.', 'max': 1},
                        {'key': 'm_1_1k', 'stt': 'k)', 'name': 'Gần gũi, sâu sát với cơ sở; thực hiện tốt việc giữ mối liên hệ với cấp ủy và Nhân dân nơi cư trú.', 'max': 1},

                        {'is_header': True, 'stt': '2', 'name': 'Năng lực chuyên môn, nghiệp vụ theo yêu cầu của vị trí việc làm; khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao; thái độ công tác trong thực hiện nhiệm vụ; tinh thần đổi mới sáng tạo, dám nghĩ, dám làm, dám chịu trách nhiệm vì lợi ích chung', 'max': 15},
                        {'is_header': True, 'stt': 'a)', 'name': 'Năng lực chuyên môn, nghiệp vụ theo yêu cầu của vị trí việc làm', 'max': 5},
                        {'key': 'm_1_2a1', 'stt': '-', 'name': 'Có hiểu biết đầy đủ về lĩnh vực công tác được phân công; nắm vững quy định pháp luật, quy trình nghiệp vụ có liên quan đến vị trí việc làm.', 'max': 2},
                        {'key': 'm_1_2a2', 'stt': '-', 'name': 'Thường xuyên cập nhật kiến thức mới, có khả năng nghiên cứu, phân tích, tổng hợp và vận dụng sáng tạo vào công việc; đáp ứng yêu cầu đổi mới, cải cách hành chính.', 'max': 1},
                        {'key': 'm_1_2a3', 'stt': '-', 'name': 'Có kỹ năng xử lý công việc độc lập, làm việc nhóm hiệu quả.', 'max': 1},
                        {'key': 'm_1_2a4', 'stt': '-', 'name': 'Khả năng sử dụng công nghệ thông tin, áp dụng vào công việc chuyên môn.', 'max': 1},
                        {'is_header': True, 'stt': 'b)', 'name': 'Khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao', 'max': 5},
                        {'key': 'm_1_2b1', 'stt': '-', 'name': 'Nhiệm vụ thường xuyên: Có khả năng vận dụng thành thạo kiến thức chuyên môn, nghiệp vụ để xử lý công việc chuyên môn theo kế hoạch định kỳ; duy trì ổn định chất lượng chuyên môn.', 'max': 3},
                        {'key': 'm_1_2b2', 'stt': '-', 'name': 'Nhiệm vụ đột xuất: Có khả năng đề xuất giải pháp, thực hiện hiệu quả các công việc phát sinh; có khả năng phản ứng kịp thời, đáp ứng với yêu cầu, nhiệm vụ cụ thể.', 'max': 2},
                        {'is_header': True, 'stt': 'c)', 'name': 'Thái độ công tác trong thực hiện nhiệm vụ', 'max': 2.5},
                        {'key': 'm_1_2c1', 'stt': '-', 'name': 'Tinh thần trách nhiệm, tích cực trong công việc; kịp thời tiếp cận kiến thức mới để điều chỉnh, đề xuất cải tiến quy trình hoặc giải pháp nâng cao hiệu quả công việc.', 'max': 1},
                        {'key': 'm_1_2c2', 'stt': '-', 'name': 'Có thái độ đúng mực, phong cách làm việc chuẩn mực, lề lối hành chính chuyên nghiệp trong quan hệ công tác.', 'max': 1},
                        {'key': 'm_1_2c3', 'stt': '-', 'name': 'Phối hợp có hiệu quả với tập thể, cá nhân có liên quan trong việc triển khai thực hiện các nhiệm vụ được phân công.', 'max': 0.5},
                        {'is_header': True, 'stt': 'd)', 'name': 'Tinh thần đổi mới sáng tạo, dám nghĩ, dám làm, dám chịu trách nhiệm vì lợi ích chung', 'max': 2.5},
                        {'key': 'm_1_2d1', 'stt': '-', 'name': 'Có sản phẩm, giải pháp đột phá, sáng tạo đem lại giá trị, hiệu quả thiết thực, tác động tích cực đến kết quả thực hiện nhiệm vụ của tổ chức, cơ quan, đơn vị.', 'max': 1},
                        {'key': 'm_1_2d2', 'stt': '-', 'name': 'Sẵn sàng tham gia thực hiện nhiệm vụ chính trị, nhiệm vụ có tính chất đột xuất, phức tạp hoặc trong điều kiện khó khăn.', 'max': 1},
                        {'key': 'm_1_2d3', 'stt': '-', 'name': 'Có tinh thần chịu trách nhiệm trước kết quả công việc; sẵn sàng nhận trách nhiệm khi có sai sót và có biện pháp khắc phục rõ ràng, cụ thể.', 'max': 0.5},

                        {'is_header': True, 'stt': '3', 'name': 'Về tự phê bình và phê bình, khắc phục hạn chế, khuyết điểm', 'max': 5},
                        {'key': 'm_1_3', 'stt': '-', 'name': 'Tinh thần tự phê bình, tự soi, tự sửa của cá nhân; mức độ tự giác nhận diện hạn chế, khuyết điểm của bản thân và kết quả khắc phục những hạn chế, khuyết điểm đã được chỉ ra.', 'max': 5}
                    ]
                },
                {
                    'group': 'II. NHÓM TIÊU CHÍ VỀ KẾT QUẢ THỰC HIỆN NHIỆM VỤ — MỤC 8: KHỐI DOANH NGHIỆP NHÀ NƯỚC (70 ĐIỂM)',
                    'items': [
                        {'is_header': True, 'stt': '8', 'name': 'Khối doanh nghiệp nhà nước (Đánh giá mức độ hoàn thành các chỉ tiêu, nhiệm vụ được giao, kết quả được lượng hoá bằng sản phẩm, tiến độ, chất lượng hoặc chỉ số đo lường hiệu quả công việc theo chức năng, nhiệm vụ của tổ chức, cơ quan, đơn vị)', 'max': 70},
                        {'key': 'm_2_1', 'stt': '-', 'name': 'Mức độ hoàn thành nhiệm vụ theo định mức, vị trí công việc; không để phát sinh lỗi kỹ thuật, sai sót ảnh hưởng đến hiệu quả sản xuất, kinh doanh, sản phẩm, dịch vụ.', 'max': 70}
                    ]
                }
            ]

    all_items = [it for g in groups for it in g['items'] if not it.get('is_header')]
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
        form_code = "Mẫu 01"
        title_text = f"BÁO CÁO KIỂM ĐIỂM, TỰ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG\nCỦA TẬP THỂ {branch_name}\nNăm {year}"
        sub_title = ""
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
        form_code = "Mẫu 02A" if is_sample_2a else "Mẫu 02B"
        title_text = f"BẢN KIỂM ĐIỂM, TỰ ĐÁNH GIÁ, XẾP LOẠI CỦA CÁ NHÂN\nNăm {year}"
        sub_title = (
            "(Cá nhân giữ chức vụ lãnh đạo, quản lý)"
            if is_sample_2a
            else "(Cá nhân không giữ chức vụ lãnh đạo, quản lý)"
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

    if is_col:
        p_intro = doc.add_paragraph()
        p_intro.paragraph_format.space_after = Pt(4)
        p_intro.add_run(
            f"Căn cứ kết quả lãnh đạo, chỉ đạo thực hiện nhiệm vụ chính trị tại đơn vị; "
            f"tập thể {branch_name} kiểm điểm, tự đánh giá, xếp loại chất lượng với các nội dung chủ yếu sau:"
        )
    else:
        p_info = doc.add_paragraph()
        p_info.paragraph_format.line_spacing = 1.2
        p_info.paragraph_format.space_after = Pt(4)
        
        p_info.add_run("Họ và tên: ").bold = True
        r_name = p_info.add_run(f"{m.get('name').upper()}          ")
        r_name.bold = True
        p_info.add_run("Ngày sinh: ").bold = True
        p_info.add_run(f"{m.get('dob', '01/01/1985')}\n")
        if is_sample_2a:
            p_info.add_run("Chức vụ Đảng: ").bold = True
            p_info.add_run(f"{m.get('role_party', 'Cấp ủy viên')}\n")
            p_info.add_run("Chức vụ chính quyền: ").bold = True
            p_info.add_run(f"{m.get('title_pro', 'Lãnh đạo quản lý')}\n")
            p_info.add_run("Chức vụ đoàn thể: ").bold = True
            p_info.add_run(f"{m.get('union_role', 'Không')}\n")
        p_info.add_run("Đơn vị công tác: ").bold = True
        p_info.add_run(f"{m.get('dept', 'VietinBank Chi nhánh Bảo Lộc')}          ")
        p_info.add_run("Chi bộ: ").bold = True
        p_info.add_run(f"{branch_name}")

    criteria_groups = get_criteria_for_target(target_val, data)
    def grp_sum(g):
        s = round(sum(float(it['score']) for it in g['items'] if not it.get('is_header')), 1)
        return int(s) if s.is_integer() else s
    g1_score = grp_sum(criteria_groups[0]) if len(criteria_groups) > 0 else 0
    g2_score = grp_sum(criteria_groups[1]) if len(criteria_groups) > 1 else 0
    computed_total = round(g1_score + g2_score, 1)
    if float(computed_total).is_integer():
        computed_total = int(computed_total)

    p_sec_a = doc.add_paragraph()
    p_sec_a.paragraph_format.space_before = Pt(6)
    p_sec_a.paragraph_format.space_after = Pt(4)
    r_sa = p_sec_a.add_run("A. NỘI DUNG KIỂM ĐIỂM")
    r_sa.bold = True
    r_sa.font.size = Pt(12.5)

    if is_col:
        p_a1 = doc.add_paragraph()
        p_a1.paragraph_format.line_spacing = 1.2
        p_a1.add_run("I. Ưu điểm, kết quả đạt được\n").bold = True
        col_prompts = [
            "1. Việc chấp hành Cương lĩnh chính trị, Điều lệ Đảng, nghị quyết, chỉ thị, kết luận của Đảng, pháp luật của Nhà nước; nguyên tắc tổ chức và hoạt động của Đảng, nhất là nguyên tắc tập trung dân chủ; thực hiện quy chế làm việc.",
            "2. Kết quả thực hiện các mục tiêu, chỉ tiêu, nhiệm vụ được đề ra trong nghị quyết đại hội nhiệm kỳ, kế hoạch, chương trình công tác năm, được cấp có thẩm quyền giao, phê duyệt và các nhiệm vụ phát sinh, đột xuất được cấp có thẩm quyền giao.",
            "3. Công tác xây dựng, chỉnh đốn Đảng và hệ thống chính trị; xây dựng tổ chức bộ máy tinh gọn, hoạt động hiệu năng, hiệu lực, hiệu quả; tinh giản biên chế gắn với cải cách hành chính và chuyển đổi số; đấu tranh phòng, chống tham nhũng, lãng phí, quan liêu, tiêu cực và ngăn chặn, đẩy lùi sự suy thoái về tư tưởng chính trị, đạo đức, lối sống, \"tự diễn biến\", \"tự chuyển hoá\" trong nội bộ; công tác kiểm soát quyền lực, phòng, chống chạy chức, chạy quyền, lợi ích nhóm; xây dựng và thực hiện cơ chế công khai, minh bạch gắn với trách nhiệm giải trình trong công tác cán bộ; công tác kiểm tra, giám sát, kỷ luật đảng và công tác tiếp công dân, giải quyết đơn, thư, khiếu nại, tố cáo, kiến nghị, phản ánh của tổ chức, cá nhân.",
            "4. Công tác lãnh đạo và định hướng tư tưởng của cán bộ, đảng viên và dư luận tại đơn vị. Kết quả ban hành văn bản chỉ đạo, hướng dẫn hằng năm về công tác chính trị, tư tưởng; xây dựng, triển khai thực hiện chương trình (hoặc kế hoạch) công tác chính trị, tư tưởng, bảo đảm tiến độ, chất lượng. Việc nắm bắt, dự báo, định hướng, xử lý kịp thời các vấn đề về tư tưởng thuộc phạm vi quản lý.",
            "5. Thực hiện nghiêm trách nhiệm nêu gương theo quy định. Cấp ủy, người đứng đầu cấp ủy nói đi đôi với làm, làm đúng, hiệu quả. Cấp ủy viên chấp hành nghiêm các quy định về nêu gương, không có trường hợp vi phạm.",
            "6. Thực hiện nghiêm kỷ luật phát ngôn. Quản lý, kiểm soát hiệu quả công tác tuyên truyền, kỷ luật phát ngôn và cung cấp thông tin, sử dụng các nền tảng số.",
            "7. Lãnh đạo, chỉ đạo việc đẩy mạnh học tập, thực hành tư tưởng, đạo đức, phương pháp, phong cách Hồ Chí Minh.",
            "8. Trách nhiệm của tập thể lãnh đạo, quản lý trong triển khai thực hiện các chủ trương, đường lối, nghị quyết, văn bản chỉ đạo của Đảng, của cấp trên; thực hiện nhiệm vụ chính trị của địa phương, tổ chức, cơ quan, đơn vị."
        ]
        for pr in col_prompts:
            p_a1.add_run(pr + "\n")
        p_a1.add_run(f"   {good_val}")

        p_a2 = doc.add_paragraph()
        p_a2.paragraph_format.line_spacing = 1.2
        p_a2.add_run("II. Hạn chế, khuyết điểm và nguyên nhân\n").bold = True
        p_a2.add_run(f"   {limit_val}\n")
        p_a2.add_run("III. Kết quả khắc phục những hạn chế, khuyết điểm đã được cấp có thẩm quyền kết luận hoặc được chỉ ra ở các kỳ kiểm điểm trước\n").bold = True
        p_a2.add_run("   Tập thể Chi bộ 2 đã nghiêm túc kiểm điểm, đề ra giải pháp và tổ chức khắc phục kịp thời các hạn chế đã chỉ ra từ kỳ kiểm điểm trước; phân công rõ trách nhiệm cho cấp ủy và từng đảng viên phụ trách.\n")
        p_a2.add_run("IV. Giải trình những vấn đề được gợi ý kiểm điểm (nếu có): ").bold = True
        p_a2.add_run("Không có.\n")
        p_a2.add_run("V. Trách nhiệm của tập thể, cá nhân\n").bold = True
        p_a2.add_run("   Tập thể Chi ủy, Chi bộ 2 và người đứng đầu nghiêm túc nhận trách nhiệm chung trong công tác lãnh đạo, chỉ đạo thực hiện nhiệm vụ chính trị; không để xảy ra mất đoàn kết nội bộ, tham nhũng, lãng phí, tiêu cực.\n")
        p_a2.add_run("VI. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm\n").bold = True
        p_a2.add_run(f"   {sol_val}")
    else:
        p_a1 = doc.add_paragraph()
        p_a1.paragraph_format.line_spacing = 1.2
        p_a1.add_run("I. Ưu điểm, kết quả đạt được\n").bold = True
        if is_sample_2a:
            mem_prompts = [
                "1. Về phẩm chất chính trị, đạo đức, lối sống; tinh thần đoàn kết, thống nhất trong nội bộ; ý thức tổ chức kỷ luật, việc thực hiện những điều đảng viên không được làm; tác phong, lề lối làm việc.",
                "2. Việc thực hiện nhiệm vụ, quyền hạn và kết quả thực hiện các chỉ tiêu, nhiệm vụ được giao trong năm, gắn với hiệu quả, tiến độ, chất lượng công việc.",
                "3. Việc thực hiện cam kết tu dưỡng, rèn luyện, phấn đấu của cá nhân theo quy định.",
                "4. Kết quả công tác lãnh đạo, chỉ đạo, quản lý, điều hành; thực hiện chức trách, nhiệm vụ; mức độ hoàn thành nhiệm vụ của các cấp uỷ, tổ chức đảng trực thuộc, địa phương, tổ chức, cơ quan, đơn vị thuộc quyền quản lý; khả năng quy tụ, xây dựng đoàn kết nội bộ.",
                "5. Trách nhiệm trong công việc; tinh thần năng động, đổi mới sáng tạo, dám nghĩ, dám nói, dám làm, nói đi đôi với làm, dám chịu trách nhiệm, dám đương đầu với khó khăn, thử thách, dám hành động, dám hy sinh lợi ích cá nhân vì lợi ích chung; trách nhiệm trong xây dựng đội ngũ cán bộ, kiểm soát quyền lực, phòng, chống tiêu cực trong phạm vi chức trách, nhiệm vụ được giao; trách nhiệm trong lãnh đạo, chỉ đạo hoàn thiện thể chế góp phần phát triển khoa học, công nghệ, đổi mới sáng tạo và chuyển đổi số, về mở rộng hợp tác, hội nhập quốc tế, về đổi mới tư duy quản trị, tạo điều kiện phát triển kinh tế tư nhân; xử lý những vấn đề khó, phức tạp, nhạy cảm trong thực hiện nhiệm vụ.",
                "6. Trách nhiệm nêu gương của bản thân và gia đình; việc đấu tranh phòng, chống tham nhũng, tiêu cực, lãng phí và ngăn chặn, đẩy lùi những biểu hiện suy thoái về tư tưởng chính trị, đạo đức, lối sống, \"tự diễn biến\", \"tự chuyển hoá\"; sự tín nhiệm của cán bộ, đảng viên; việc nghiên cứu, học tập và triển khai nghị quyết.",
                "7. Công tác lãnh đạo và định hướng tư tưởng của cán bộ, đảng viên và dư luận tại đơn vị. Việc nắm bắt, dự báo, định hướng, xử lý kịp thời các vấn đề về tư tưởng thuộc phạm vi quản lý.",
                "8. Thực hiện nghiêm kỷ luật phát ngôn, quy định về cung cấp và sử dụng thông tin; nói và viết đúng quan điểm, đường lối của Đảng, chính sách, pháp luật của Nhà nước.",
                "9. Việc học tập, thực hành tư tưởng, đạo đức, phương pháp, phong cách Hồ Chí Minh; cần, kiệm, liêm, chính, chí công vô tư; việc xây dựng phong cách lãnh đạo đổi mới, gần dân, sát dân, trọng dân, thường xuyên nắm bắt tâm tư, nguyện vọng của Nhân dân; dám nghĩ, dám nói, dám làm, dám đương đầu với khó khăn, thử thách, dám đột phá, dám chịu trách nhiệm vì lợi ích chung."
            ]
        else:
            mem_prompts = [
                "1. Về phẩm chất chính trị, đạo đức, lối sống; tinh thần đoàn kết, thống nhất trong nội bộ; ý thức tổ chức kỷ luật, trách nhiệm nêu gương, việc thực hiện những điều đảng viên không được làm; tác phong, lề lối làm việc. Việc đấu tranh phòng, chống tham nhũng, tiêu cực, lãng phí và ngăn chặn, đẩy lùi những biểu hiện suy thoái về tư tưởng chính trị, đạo đức, lối sống, \"tự diễn biến\", \"tự chuyển hoá\"; việc nghiên cứu, học tập và triển khai nghị quyết.",
                "2. Việc thực hiện nhiệm vụ, quyền hạn và kết quả thực hiện các chỉ tiêu, nhiệm vụ được giao trong năm, gắn với hiệu quả, tiến độ, chất lượng công việc.",
                "3. Thực hiện nghiêm kỷ luật phát ngôn, quy định về cung cấp và sử dụng thông tin; nói và viết đúng quan điểm, đường lối của Đảng, chính sách, pháp luật của Nhà nước.",
                "4. Việc học tập và thực hành tư tưởng, đạo đức, phương pháp, phong cách Hồ Chí Minh; cần, kiệm, liêm, chính, chí công vô tư.",
                "5. Việc thực hiện cam kết tu dưỡng, rèn luyện, phấn đấu của cá nhân theo quy định."
            ]
        for pr in mem_prompts:
            p_a1.add_run(pr + "\n")
        p_a1.add_run(f"   {good_val}")

        p_a2 = doc.add_paragraph()
        p_a2.paragraph_format.line_spacing = 1.2
        p_a2.add_run("II. Hạn chế, khuyết điểm và nguyên nhân\n").bold = True
        p_a2.add_run("1. Hạn chế, khuyết điểm:\n").bold = True
        p_a2.add_run(f"   {limit_val}\n")
        p_a2.add_run("2. Nguyên nhân của hạn chế, khuyết điểm:\n").bold = True
        p_a2.add_run("   Do khối lượng công việc kinh doanh và tác nghiệp ngân hàng lớn vào các giai đoạn cao điểm; bản thân có lúc chưa sắp xếp thời gian tự nghiên cứu chuyên sâu thật sự khoa học.\n")
        p_a2.add_run("III. Kết quả khắc phục những hạn chế, khuyết điểm đã được cấp có thẩm quyền kết luận hoặc được chỉ ra ở các kỳ kiểm điểm trước\n").bold = True
        p_a2.add_run("   Đã nghiêm túc tiếp thu ý kiến góp ý của Chi bộ từ kỳ kiểm điểm trước và từng bước khắc phục có hiệu quả trong quá trình công tác.\n")
        p_a2.add_run("IV. Giải trình những vấn đề được gợi ý kiểm điểm (nếu có): ").bold = True
        p_a2.add_run("Không có.\n")
        p_a2.add_run("V. Làm rõ trách nhiệm của cá nhân đối với những hạn chế, khuyết điểm của tập thể (nếu có)\n").bold = True
        p_a2.add_run("   Nghiêm túc nhận trách nhiệm cá nhân trong phạm vi công việc được phân công phụ trách.\n")
        p_a2.add_run("VI. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm\n").bold = True
        p_a2.add_run(f"   {sol_val}")

    p_sec_b = doc.add_paragraph()
    p_sec_b.paragraph_format.space_before = Pt(8)
    p_sec_b.paragraph_format.space_after = Pt(4)
    r_sb = p_sec_b.add_run("B. NỘI DUNG TỰ ĐÁNH GIÁ, XẾP LOẠI")
    r_sb.bold = True
    r_sb.font.size = Pt(12.5)

    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.space_after = Pt(4)
    p_b1.add_run("I. Đánh giá theo các nhóm tiêu chí\n").bold = True
    if is_col:
        p_b1.add_run(f"Trên cơ sở cụ thể hoá khung tiêu chí đánh giá đối với tập thể được quy định tại Phụ lục 1, Quy định số 366-QĐ/TW, tập thể {branch_name} tự đánh giá, chấm điểm theo các tiêu chí được cụ thể hóa như sau:\n")
    elif is_sample_2a:
        p_b1.add_run("Trên cơ sở cụ thể hoá khung tiêu chí đánh giá đối với cá nhân giữ chức vụ lãnh đạo, quản lý được quy định tại Phụ lục 3, Quy định số 366-QĐ/TW, cá nhân tự đánh giá, chấm điểm theo các tiêu chí được cụ thể hóa như sau:\n")
    else:
        p_b1.add_run("Trên cơ sở cụ thể hoá khung tiêu chí đánh giá đối với cá nhân không giữ chức vụ lãnh đạo, quản lý được quy định tại Phụ lục 4, Quy định số 366-QĐ/TW, cá nhân tự đánh giá, chấm điểm theo các tiêu chí được cụ thể hóa như sau:\n")
    p_b1.add_run(f"   1. Nhóm tiêu chí chung: {g1_score}/30\n")
    p_b1.add_run(f"   2. Nhóm tiêu chí về kết quả thực hiện nhiệm vụ: {g2_score}/70\n")
    p_b1.add_run(f"   Tổng điểm: {computed_total}/100").bold = True

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

            is_hdr = bool(it.get('is_header'))
            if is_hdr:
                for c in r_cells:
                    set_cell_background(c, "F8FAFC")

            p0 = r_cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r0 = p0.add_run(it['stt'])
            r0.font.size = Pt(10.5)
            if is_hdr:
                r0.bold = True

            p1 = r_cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r1 = p1.add_run(it['name'])
            r1.font.size = Pt(10.5)
            if is_hdr:
                r1.bold = True

            p2 = r_cells[2].paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_max = p2.add_run(f"{it['max']}đ")
            r_max.font.size = Pt(10.5)
            r_max.bold = True

            p3 = r_cells[3].paragraphs[0]
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if not is_hdr:
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

    def chk(label):
        return "☒" if ranking_val == label else "☐"

    p_rank = doc.add_paragraph()
    p_rank.paragraph_format.line_spacing = 1.2
    p_rank.paragraph_format.space_before = Pt(6)
    p_rank.paragraph_format.space_after = Pt(6)
    p_rank.add_run("II. Đề xuất xếp loại mức chất lượng:\n").bold = True
    p_rank.add_run(
        f"   {chk('Hoàn thành xuất sắc nhiệm vụ')} Hoàn thành xuất sắc nhiệm vụ          "
        f"{chk('Hoàn thành tốt nhiệm vụ')} Hoàn thành tốt nhiệm vụ\n"
        f"   {chk('Hoàn thành nhiệm vụ')} Hoàn thành nhiệm vụ                         "
        f"{chk('Không hoàn thành nhiệm vụ')} Không hoàn thành nhiệm vụ\n"
    )
    r_appx = p_rank.add_run("(Gửi kèm Phụ lục về kết quả thực hiện nhiệm vụ trong đó thể hiện sản phẩm tiến độ, chất lượng cụ thể - theo hướng dẫn của các cấp uỷ, tổ chức đảng)")
    r_appx.italic = True
    r_appx.font.size = Pt(11)

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
        p_sr.add_run("T/M TẬP THỂ LÃNH ĐẠO, QUẢN LÝ\n").bold = True
        r_note = p_sr.add_run("(Ký, ghi rõ họ tên và đóng dấu)\n\n\n")
        r_note.italic = True
        r_note.font.size = Pt(11)
        p_sr.add_run(host_name).bold = True
    else:
        p_sr.add_run("NGƯỜI TỰ KIỂM ĐIỂM\n").bold = True
        r_note = p_sr.add_run("(Ký, ghi rõ họ tên)\n\n\n")
        r_note.italic = True
        r_note.font.size = Pt(11)
        p_sr.add_run(m.get('name').upper()).bold = True

        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(10)
        p_c.paragraph_format.space_after = Pt(4)
        r_c = p_c.add_run("C. KẾT QUẢ ĐÁNH GIÁ, XẾP LOẠI")
        r_c.bold = True
        r_c.font.size = Pt(12.5)

        p_c1 = doc.add_paragraph()
        p_c1.paragraph_format.line_spacing = 1.2
        p_c1.add_run("I. Đánh giá, xếp loại chất lượng cán bộ, công chức, viên chức\n").bold = True
        p_c1.add_run("- Nhận xét, đánh giá của người quản lý, sử dụng cán bộ, công chức, viên chức: ..........................................................................\n")
        p_c1.add_run("- Mức xếp loại chất lượng công chức, viên chức: ...........................................................")

        t_tt = doc.add_table(rows=1, cols=2)
        t_tt.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_tt.autofit = False
        clear_table_borders(t_tt)
        t_tt.columns[0].width = Mm(75)
        t_tt.columns[1].width = Mm(90)
        p_tt = t_tt.cell(0, 1).paragraphs[0]
        p_tt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tt.add_run("THỦ TRƯỞNG CƠ QUAN, ĐƠN VỊ\n").bold = True
        r_ttn = p_tt.add_run("(Xác lập thời điểm, ký, ghi rõ họ tên và đóng dấu)\n\n\n")
        r_ttn.italic = True
        r_ttn.font.size = Pt(11)

        p_c2 = doc.add_paragraph()
        p_c2.paragraph_format.line_spacing = 1.2
        p_c2.paragraph_format.space_before = Pt(6)
        p_c2.add_run("II. Đánh giá, xếp loại chất lượng đảng viên\n").bold = True
        p_c2.add_run(f"- Nhận xét, đánh giá của chi ủy: {m.get('host_conclusion', '....................................................................................')}\n")
        p_c2.add_run("- Chi bộ đề xuất xếp loại mức chất lượng: ...................................................................")

        table_sign_cb = doc.add_table(rows=1, cols=2)
        table_sign_cb.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_sign_cb.autofit = False
        clear_table_borders(table_sign_cb)
        table_sign_cb.columns[0].width = Mm(75)
        table_sign_cb.columns[1].width = Mm(90)
        p_cb = table_sign_cb.cell(0, 1).paragraphs[0]
        p_cb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cb.paragraph_format.line_spacing = 1.05
        p_cb.add_run("T/M CHI ỦY (CHI BỘ)\n").bold = True
        r_cb_note = p_cb.add_run("(Xác lập thời điểm, ký, ghi rõ họ tên)\n\n\n")
        r_cb_note.italic = True
        r_cb_note.font.size = Pt(11)
        p_cb.add_run(host_name).bold = True

        p_du = doc.add_paragraph()
        p_du.paragraph_format.space_before = Pt(6)
        p_du.add_run("- Ban Thường vụ Đảng ủy cơ sở hoặc Đảng ủy trực thuộc cơ sở xếp loại mức chất lượng: ...................................................")

        t_du = doc.add_table(rows=1, cols=2)
        t_du.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_du.autofit = False
        clear_table_borders(t_du)
        t_du.columns[0].width = Mm(75)
        t_du.columns[1].width = Mm(90)
        p_td = t_du.cell(0, 1).paragraphs[0]
        p_td.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_td.add_run("T/M ĐẢNG ỦY\n").bold = True
        r_tdn = p_td.add_run("(Xác lập thời điểm, ký, ghi rõ họ tên và đóng dấu)\n\n\n")
        r_tdn.italic = True
        r_tdn.font.size = Pt(11)

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
