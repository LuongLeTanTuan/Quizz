# -*- coding: utf-8 -*-
import json
import re
import fitz

# -------------------------------------------------------------
# 1. PARSE & AUDIT ADVANCED EXAM (TRAC NGHIEM NANG CAO.pdf)
# -------------------------------------------------------------
doc_nc = fitz.open('TRAC NGHIEM NANG CAO.pdf')
full_text_nc = '\n'.join([page.get_text('text') for page in doc_nc])

pattern = r'(?ms)^\s*(\d+)\.\s*([^:\n]+):\s*(.*?)\n\s*A\.\s*(.*?)\n\s*B\.\s*(.*?)\n\s*C\.\s*(.*?)\n\s*D\.\s*(.*?)\n\s*Answer:\s*([A-D])'
matches_nc = list(re.finditer(pattern, full_text_nc))

questions_nc = []
for idx, m in enumerate(matches_nc):
    num = int(m.group(1))
    mod = m.group(2).strip()
    q = ' '.join(m.group(3).split())
    a = ' '.join(m.group(4).split())
    b = ' '.join(m.group(5).split())
    c = ' '.join(m.group(6).split())
    d = ' '.join(m.group(7).split())
    ans = m.group(8).strip()
    
    note = ''
    # Verified answer checks & detailed learning notes
    if mod == 'MS Access 2013' and num == 1:
        ans = 'C'
        note = 'Chuẩn đáp án: Định dạng .accdb ra mắt từ Microsoft Access 2007 (.mdb dành cho bản 97-2003).'
    elif mod == 'MS Access 2013' and num == 14:
        ans = 'C'
        note = 'Chuẩn đáp án: Kiểu Number - Long Integer lưu được số 100.000 (Byte: 0..255, Integer: -32.768..32.767).'
    elif mod == 'MS Access 2013' and num == 10:
        note = 'Kiểu Text (Short Text) trong Access giới hạn tối đa 255 ký tự.'
    elif mod == 'MS Access 2013' and num == 11:
        note = 'Kiểu Memo (Long Text) cho phép lưu văn bản dài trên 255 ký tự.'
    elif mod == 'MS Access 2013' and num == 15:
        note = 'Kiểu Number - Double hỗ trợ lưu số thực có phần thập phân (như 2.34).'
    elif mod == 'MS Access 2013' and num == 16:
        note = 'Kiểu Yes/No lưu trữ hai giá trị logic TRUE / FALSE.'
    elif mod == 'MS Access 2013' and num == 18:
        note = 'Kiểu Date/Time lưu trữ ngày tháng năm.'
    elif mod == 'MS Access 2013' and num == 20:
        note = 'Field Size quy định kích thước dữ liệu hoặc kiểu con của trường Number/Text.'
    elif mod == 'MS Access 2013' and num == 22:
        note = 'Input Mask là mặt nạ quy định khuôn mẫu ràng buộc khi nhập dữ liệu.'
    elif mod == 'MS Access 2013' and num == 25:
        note = 'Validation Rule là biểu thức điều kiện ràng buộc tính hợp lệ của dữ liệu nhập.'
    elif mod == 'MS Access 2013' and num == 26:
        note = 'Validation Text là thông báo nhắc nhở khi người dùng nhập vi phạm Validation Rule.'
    elif mod == 'MS Access 2013' and num == 30:
        note = 'Query chứa câu lệnh SQL dùng trích lọc, cập nhật hoặc xóa dữ liệu từ Table.'
    elif mod == 'MS Access 2013' and num == 31:
        note = 'Table là đối tượng cơ sở dùng để lưu trữ dữ liệu người dùng nhập vào.'
    elif mod == 'MS Access 2013' and num == 32:
        note = 'Form là biểu mẫu giao diện dùng xem, sửa, xóa các bản ghi Record.'
    elif mod == 'MS Access 2013' and num == 33:
        note = 'Report là đối tượng dùng để tổng hợp và in ấn báo cáo.'
    elif mod == 'MS Access 2013' and num == 37:
        note = 'Parameter Query cho phép người dùng nhập điều kiện lọc qua hộp thoại lúc chạy.'
    elif mod == 'MS Access 2013' and num == 38:
        note = 'Total Query dùng để thống kê dữ liệu (SUM, COUNT, GROUP BY).'
    elif mod == 'MS Access 2013' and num == 39:
        note = 'Crosstab Query dùng để thống kê ma trận 2 chiều hàng x cột.'
    elif mod == 'MS Excel 2013' and num == 25:
        note = 'File > Info > Protect Workbook > Encrypt with Password để gắn mật khẩu bảo vệ file.'
    elif mod == 'MS Excel 2013' and num == 27:
        note = 'Khi đặt mật khẩu bảo vệ Workbook, hệ thống yêu cầu nhập xác nhận 2 lần.'
    elif mod == 'MS Excel 2013' and num == 28:
        note = 'Để mở Workbook đã gắn mật khẩu, hệ thống chỉ yêu cầu nhập 1 lần.'
    elif mod == 'MS Excel 2013' and num == 37:
        note = 'Các điều kiện lọc đồng thời (VÀ / AND) trong Advanced Filter được đặt trên CÙNG DÒNG.'
    elif mod == 'MS Excel 2013' and num == 38:
        note = 'Các điều kiện lọc hoặc (HOẶC / OR) trong Advanced Filter được đặt trên KHÁC DÒNG KHÁC CỘT.'
    elif mod == 'MS Excel 2013' and num == 91:
        note = 'Vào tab Data > nhóm Sort & Filter > chọn Advanced để lọc nâng cao.'
    elif mod == 'MS Excel 2013' and num == 94:
        note = 'MID(A1, 4, LEN(A1)-5) lấy chính xác phần tên hàng ở giữa mã.'
    elif mod == 'MS Excel 2013' and num == 95:
        note = 'LEFT(A1, FIND(",", A1)-1) trích xuất phần tên trước dấu phẩy.'
    elif mod == 'MS Excel 2013' and num == 118:
        note = 'SUM là hàm tính tổng thông thường, không phải hàm xử lý CSDL (các hàm CSDL bắt đầu bằng D: DSUM, DMIN...).'
    elif mod == 'MS Excel 2013' and num == 151:
        note = 'SUMIFS kết hợp ký tự đại diện * cho phép tính tổng theo nhiều điều kiện chuỗi.'
    elif mod == 'MS Winword 2013' and num == 1:
        note = 'AutoCorrect trong Word dùng thiết lập gõ tắt (Replace - With).'
    elif mod == 'MS Winword 2013' and num == 4:
        note = 'Tab References > Insert Citation dùng để chèn trích dẫn nguồn học thuật.'
    elif mod == 'MS Winword 2013' and num == 10:
        note = 'Ngắt đoạn sang đoạn mới (Section Break) bằng lệnh Page Layout > Break > Next Page.'
    elif mod == 'MS Winword 2013' and num == 11:
        note = 'Muốn đánh số trang bắt đầu từ 1 cho từng chương, ta phải ngắt đoạn (Section Break) ở cuối mỗi chương.'
    elif mod == 'MS Winword 2013' and num == 12:
        note = 'Chia cột văn bản: Page Layout > Columns > More Columns.'
    elif mod == 'MS Winword 2013' and num == 14:
        note = 'Lặp lại dòng tiêu đề bảng trên nhiều trang: Thẻ Layout > Repeat Header Rows.'
    elif mod == 'MS Winword 2013' and num == 21:
        note = 'Tạo tiêu đề chân trang cho các trang: Insert > Footer.'
    elif mod == 'MS Winword 2013' and num == 22:
        note = 'Footnote đặt chú thích ở cuối trang hiện hành (References > Insert Footnote).'
    elif mod == 'MS Winword 2013' and num == 23:
        note = 'Endnote đặt chú thích ở cuối toàn bộ văn bản (References > Insert Endnote).'
    elif mod == 'MS Winword 2013' and num == 30:
        note = 'Tạo mục lục tự động: References > Table of Contents > Insert Table of Contents.'
    elif mod == 'MS Winword 2013' and num == 31:
        note = 'Mailings > Select Recipients > Use Existing List: Chỉ đường dẫn đến file dữ liệu nguồn có sẵn.'
    elif mod == 'MS Winword 2013' and num == 32:
        note = 'Mailings > Select Recipients > Type New List: Tạo danh sách dữ liệu mới trực tiếp.'
    elif mod == 'MS Winword 2013' and num == 33:
        note = 'Mailings > Insert Merge Field: Chỉ định chèn trường dữ liệu vào tài liệu mẫu.'

    ans_idx = {'A': 0, 'B': 1, 'C': 2, 'D': 3}[ans]
    questions_nc.append({
        'id': idx + 1,
        'original_num': num,
        'module': mod,
        'question': q,
        'options': [a, b, c, d],
        'answer': ans,
        'ans_idx': ans_idx,
        'note': note
    })

with open('questions_nangcao_parsed.json', 'w', encoding='utf-8') as f:
    json.dump(questions_nc, f, ensure_ascii=False, indent=2)

# -------------------------------------------------------------
# 2. LOAD BASIC EXAM (TRAC NGHIEM CO BAN)
# -------------------------------------------------------------
with open('questions_parsed.json', 'r', encoding='utf-8') as f:
    questions_cb = json.load(f)

# -------------------------------------------------------------
# 3. MINDMAPS DATA (ADVANCED & BASIC)
# -------------------------------------------------------------
mindmaps_nc = [
    {
        "id": "module-nc-1",
        "title": "Module 7: Soạn Thảo Văn Bản MS Word 2013 Nâng Cao",
        "shortTitle": "M7: Word Nâng Cao",
        "category": "MS Winword 2013",
        "count": 130,
        "nodes": [
            {
                "title": "1. Trộn Thư Nâng Cao (Mailings)",
                "items": [
                    {"label": "Quy trình trộn thư", "detail": "Bắt đầu với Start Mail Merge (Letters/Envelopes) -> Chọn nguồn Select Recipients -> Chèn trường Insert Merge Field -> Xem trước Preview Results -> Hoàn tất Finish & Merge."},
                    {"label": "Select Recipients", "detail": "Use Existing List: Chọn file dữ liệu có sẵn (Excel/Access). Type New List: Tạo mới danh sách người nhận ngay trong Word."},
                    {"label": "Insert Merge Field", "detail": "Chèn các cột dữ liệu (Họ tên, Địa chỉ, Điểm...) vào các vị trí tương ứng trong tài liệu mẫu."},
                    {"label": "Các nhóm lệnh thuộc Mailings", "detail": "Create, Start Mail Merge, Write & Insert Fields, Preview Results, Finish. Chú ý: Lệnh 'Arrange' thuộc Page Layout, 'Footnotes/Index' thuộc References, KHÔNG thuộc Mailings."}
                ]
            },
            {
                "title": "2. Mục Lục & Tham Chiếu Học Thuật (References)",
                "items": [
                    {"label": "Mục lục tự động (Table of Contents)", "detail": "Thẻ References -> Table of Contents -> Insert Table of Contents sau khi đã gán Level/Heading cho các tiêu đề trong bài."},
                    {"label": "Footnote vs Endnote", "detail": "Footnote (Alt+Ctrl+F): Chú thích xuất hiện ở CUỐI MỖI TRANG. Endnote (Alt+Ctrl+D): Chú thích xuất hiện ở CUỐI TOÀN BỘ TÀI LIỆU."},
                    {"label": "Trích dẫn khoa học (Citations)", "detail": "Thẻ References -> Insert Citation: Quản lý nguồn trích dẫn theo chuẩn APA, MLA, Chicago và tự động xuất Bibliography."},
                    {"label": "Captions & Index", "detail": "Captions: Đánh nhãn hình ảnh/bảng biểu (Hình 1, Bảng 1). Index: Tạo bảng tra cứu thuật ngữ ở cuối sách."}
                ]
            },
            {
                "title": "3. Phân Đoạn & Bảng Biểu Nâng Cao",
                "items": [
                    {"label": "Section Break (Ngắt đoạn)", "detail": "Page Layout -> Breaks -> Next Page. Cực kỳ quan trọng để đánh số trang bắt đầu lại từ 1 cho từng chương, hoặc đổi trang ngang/dọc xen kẽ."},
                    {"label": "Lặp lại tiêu đề Bảng (Header Rows)", "detail": "Khi bảng biểu dài qua nhiều trang, chọn dòng đầu -> Thẻ Layout -> Repeat Header Rows để tự động lặp lại tiêu đề mỗi trang."},
                    {"label": "Thao tác trên ô Bảng", "detail": "Gộp ô: Merge Cells. Tách ô: Split Cells. Đổi hướng chữ: Text Direction. Căn chỉnh lề: Cell Alignment (9 hướng)."},
                    {"label": "Định vị hình ảnh (Wrap Text)", "detail": "In Line with Text (ảnh như 1 chữ cái), In Front of Text (ảnh nổi che chữ), Behind Text (ảnh chìm), Square/Tight/Through (chữ bao quanh ảnh)."}
                ]
            },
            {
                "title": "4. Gõ Tắt AutoCorrect & Bảo Mật Văn Bản",
                "items": [
                    {"label": "AutoCorrect (Gõ tắt)", "detail": "File -> Options -> Proofing -> AutoCorrect Options: Nhập từ viết tắt vào 'Replace', nhập cụm từ thay thế vào 'With' -> Add."},
                    {"label": "Bảo mật tài liệu", "detail": "File -> Info -> Protect Document -> Encrypt with Password. Người dùng bắt buộc phải nhập mật khẩu khi MỞ tệp văn bản."},
                    {"label": "Khổ giấy & Hướng in", "detail": "Page Setup: Orientation Portrait (In trang dọc), Landscape (In trang ngang). Giấy chuẩn thông dụng là A4."},
                    {"label": "Format Painter & Tìm kiếm", "detail": "Format Painter: Sao chép nhanh định dạng đã có. Ctrl + H: Tìm kiếm và Thay thế (Find What / Replace With)."}
                ]
            }
        ]
    },
    {
        "id": "module-nc-2",
        "title": "Module 8: Bảng Tính MS Excel 2013 Nâng Cao",
        "shortTitle": "M8: Excel Nâng Cao",
        "category": "MS Excel 2013",
        "count": 31,
        "nodes": [
            {
                "title": "1. Bảo Mật Workbook & Worksheet",
                "items": [
                    {"label": "Mật khẩu cho Workbook", "detail": "File -> Info -> Protect Workbook -> Encrypt with Password. Hệ thống yêu cầu nhập xác nhận 2 LẦN khi đặt, và 1 LẦN khi mở file."},
                    {"label": "Khóa bảo vệ Worksheet", "detail": "File -> Info -> Protect Current Sheet (hoặc chuột phải tên Sheet -> Protect Sheet). Nhập pass 2 lần để khóa, 1 lần để mở."},
                    {"label": "Xóa mật khẩu", "detail": "Vào lại mục File -> Info -> Protect Workbook -> Encrypt with Password, xóa trắng ô mật khẩu và bấm OK, lưu file."}
                ]
            },
            {
                "title": "2. Lọc Dữ Liệu Nâng Cao (Advanced Filter)",
                "items": [
                    {"label": "Đường dẫn thực hiện", "detail": "Vào thẻ Data -> trong nhóm Sort & Filter -> Chọn nút Advanced."},
                    {"label": "Quy tắc Miền điều kiện (Criteria)", "detail": "Các điều kiện viết trên CÙNG MỘT DÒNG = đồng thời thỏa mãn (VÀ / AND). Các điều kiện viết trên KHÁC DÒNG = thỏa mãn một trong các điều kiện (HOẶC / OR)."},
                    {"label": "Điều kiện VÀ trên cùng 1 cột", "detail": "Lặp lại tên cột đó 2 lần trên dòng tiêu đề miền điều kiện, rồi ghi 2 điều kiện trên cùng dòng (ví dụ: Điểm >= 5 và Điểm <= 8)."},
                    {"label": "Vị trí xuất kết quả lọc", "detail": "Cho phép lọc tại chỗ hoặc xuất sang vùng khác (Copy to another location). Đứng tại Sheet nào chạy lệnh thì kết quả trích xuất được phép xuất ra ở chính Sheet đó."}
                ]
            },
            {
                "title": "3. Hàm Xử Lý Chuỗi Nâng Cao & Ký Tự Đại Diện",
                "items": [
                    {"label": "Tách Họ - Tên linh hoạt", "detail": "Lấy tên trước dấu phẩy: =LEFT(A1, FIND(\",\", A1)-1). Tách chuỗi ở giữa: =MID(A1, 4, LEN(A1)-5)."},
                    {"label": "Ký tự đại diện (*)", "detail": "Dấu * đại diện cho chuỗi ký tự bất kỳ. Ví dụ: \"*L1*\" đại diện chuỗi có chứa L1 ở bất kỳ vị trí nào."},
                    {"label": "Đếm & Tính tổng nhiều điều kiện", "detail": "COUNTIF(range, \"*L1*\") hoặc COUNTIFS(range, \"*L1*\"). Tính tổng theo nhiều điều kiện: SUMIFS(sum_range, cr_range1, cond1, cr_range2, cond2)."}
                ]
            },
            {
                "title": "4. Hàm Cơ Sở Dữ Liệu (Database Functions)",
                "items": [
                    {"label": "Đặc điểm hàm CSDL", "detail": "Đều bắt đầu bằng chữ D: DSUM, DCOUNT, DCOUNTA, DMIN, DMAX, DAVERAGE. Cú pháp: =DSUM(database, field, criteria)."},
                    {"label": "Phân biệt với hàm thường", "detail": "SUM, COUNT, COUNTA là hàm thống kê tiêu chuẩn, không phải hàm xử lý CSDL vì không dùng cấu trúc bảng kèm miền điều kiện (Criteria)."}
                ]
            }
        ]
    },
    {
        "id": "module-nc-3",
        "title": "Module 9: Quản Trị CSDL MS Access 2013 Nâng Cao",
        "shortTitle": "M9: Access Nâng Cao",
        "category": "MS Access 2013",
        "count": 40,
        "nodes": [
            {
                "title": "1. Khái Niệm & 4 Đối Tượng Cốt Lõi",
                "items": [
                    {"label": "Table (Bảng dữ liệu)", "detail": "Đối tượng nền tảng để LƯU TRỮ DỮ LIỆU người dùng nhập vào. Cấu trúc gồm các dòng (Record - Bản ghi) và cột (Field - Trường)."},
                    {"label": "Query (Truy vấn)", "detail": "Chứa câu lệnh SQL dùng để TRÍCH LỌC, tính toán, kết nối, cập nhật hoặc xóa dữ liệu từ một hay nhiều Table."},
                    {"label": "Form (Biểu mẫu giao diện)", "detail": "Dùng để XEM, NHẬP, SỬA, XÓA các bản ghi Record thuận tiện cho người dùng cuối."},
                    {"label": "Report (Báo cáo)", "detail": "Dùng để TỔNG HỢP, thống kê định dạng chuyên nghiệp phục vụ IN ẤN hoặc xuất file."},
                    {"label": "Định dạng tệp .accdb", "detail": ".accdb là định dạng chuẩn từ Access 2007, 2010, 2013, 2016... Các phiên bản cũ (Access 97, 2003) sử dụng đuôi .mdb."}
                ]
            },
            {
                "title": "2. Thiết Kế Bảng & Các Kiểu Dữ Liệu",
                "items": [
                    {"label": "Table Design vs Table", "detail": "Table Design cho phép người dùng chỉ định chi tiết tên trường, kiểu dữ liệu và thuộc tính. Table (Datasheet) nhập liệu tự động."},
                    {"label": "Kiểu Text vs Memo", "detail": "Short Text (Text): Lưu trữ tối đa 255 ký tự. Long Text (Memo): Lưu trữ văn bản dài TRÊN 255 ký tự (lên tới 65.535 ký tự)."},
                    {"label": "Các kiểu số (Number)", "detail": "Byte (0..255). Integer (-32.768..32.767). Long Integer (-2.14 tỷ..2.14 tỷ, thích hợp lưu số 100.000). Single / Double (số thực có phần thập phân như 2.34)."},
                    {"label": "Kiểu Yes/No & Date/Time", "detail": "Yes/No: Lưu giá trị logic nhị phân TRUE / FALSE. Date/Time: Lưu ngày giờ (Short Date: 10/03/2017)."}
                ]
            },
            {
                "title": "3. Thuộc Tính Trường Dữ Liệu (Field Properties)",
                "items": [
                    {"label": "Field Size & Format", "detail": "Field Size: Xác định kích thước kiểu dữ liệu. Format: Định dạng hiển thị dữ liệu ra màn hình."},
                    {"label": "Input Mask & Caption", "detail": "Input Mask: Mặt nạ ràng buộc khuôn mẫu nhập liệu. Caption: Nhãn tiêu đề hiển thị thay cho tên trường gốc."},
                    {"label": "Default Value & Required", "detail": "Default Value: Giá trị ban đầu mặc định khi tạo mới bản ghi. Required: Yêu cầu bắt buộc phải nhập dữ liệu (không được rỗng)."},
                    {"label": "Validation Rule & Text", "detail": "Validation Rule: Điều kiện ràng buộc tính hợp lệ dữ liệu. Validation Text: Câu thông báo lỗi hiển thị khi vi phạm điều kiện."},
                    {"label": "Indexed & Allow Zero Length", "detail": "Indexed: Tạo chỉ mục để tìm kiếm nhanh và kiểm soát trùng lặp dữ liệu (No Duplicates). Allow Zero Length: Cho phép chuỗi rỗng."}
                ]
            },
            {
                "title": "4. Giao Diện & Phân Loại Query",
                "items": [
                    {"label": "Lưới thiết kế Query Design", "detail": "Field: Tên trường. Table: Bảng nguồn. Sort: Thứ tự sắp xếp. Show: Tích chọn để hiển thị trường. Criteria: Điều kiện lọc. Or: Điều kiện hoặc."},
                    {"label": "Select Query & Action Query", "detail": "Select: Trích lọc thông tin thông thường. Update: Cập nhật dữ liệu hàng loạt. Delete: Xóa dữ liệu. Append: Nối dữ liệu vào bảng khác."},
                    {"label": "Parameter Query", "detail": "Query có tham số, cho phép người dùng gõ điều kiện lọc linh hoạt qua hộp thoại mỗi khi chạy truy vấn."},
                    {"label": "Total Query & Crosstab Query", "detail": "Total Query: Thống kê tổng hợp (SUM, COUNT, GROUP BY). Crosstab Query: Thống kê dạng ma trận 2 chiều (hàng x cột, ô giao nhau là tổng)."}
                ]
            }
        ]
    }
]

with open('mindmaps_cb.json', 'r', encoding='utf-8') as f:
    mindmaps_cb = json.load(f)

# SVG Icon Library
ICONS = {
    "quiz": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle></svg>',
    "mindmap": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="2"></rect><rect x="14" y="3" width="7" height="7" rx="2"></rect><rect x="14" y="14" width="7" height="7" rx="2"></rect><rect x="3" y="14" width="7" height="7" rx="2"></rect><path d="M10 6.5h4M6.5 10v4M14 17.5H10M17.5 10v4"></path></svg>',
    "book": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>',
    "timer": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>',
    "streak": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"></path></svg>',
    "check": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>',
    "cross": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>',
    "search": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>',
    "chevronDown": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>',
    "settings": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>',
    "play": '<svg class="svg-ico" viewBox="0 0 24 24" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>',
    "refresh": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>',
    "arrowRight": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>',
    "bulb": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6"></path><path d="M10 22h4"></path><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5.76.76 1.23 1.52 1.41 2.5h6.18z"></path></svg>',
    "keyboard": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"></rect><line x1="6" y1="8" x2="6.01" y2="8"></line><line x1="10" y1="8" x2="10.01" y2="8"></line><line x1="14" y1="8" x2="14.01" y2="8"></line><line x1="18" y1="8" x2="18.01" y2="8"></line><line x1="6" y1="12" x2="6.01" y2="12"></line><line x1="18" y1="12" x2="18.01" y2="12"></line><line x1="8" y1="16" x2="16" y2="16"></line></svg>',
    "sun": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>',
    "moon": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>',
    "volume": '<svg class="svg-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="11 5 6 9 2 9 2 15 6 15 11 19 11 5"></polygon><path d="M19.07 4.93a10 10 0 0 1 0 14.14M15.54 8.46a5 5 0 0 1 0 7.07"></path></svg>',
    "star": '<svg class="svg-ico" viewBox="0 0 24 24" fill="currentColor"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>'
}

# Dump JSON data strings to embed into JavaScript
json_nc_str = json.dumps(questions_nc, ensure_ascii=False)
json_cb_str = json.dumps(questions_cb, ensure_ascii=False)
mindmaps_nc_str = json.dumps(mindmaps_nc, ensure_ascii=False)
mindmaps_cb_str = json.dumps(mindmaps_cb, ensure_ascii=False)

# -------------------------------------------------------------
# 4. BUILD SINGLE-FILE APPLICATION TEMPLATE
# -------------------------------------------------------------
html_content = f"""<!DOCTYPE html>
<html lang="vi" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <title>Ôn Thi Trắc Nghiệm CNTT Nâng Cao & Cơ Bản – Đề Cương Chuẩn & Sơ Đồ Tư Duy</title>
  <meta name="description" content="Hệ thống ôn luyện trắc nghiệm CNTT Nâng cao (201 câu) & Cơ bản (288 câu): Giao diện Quizizz chuẩn mực, chuẩn đáp án, giải thích chi tiết, Mindmaps trực quan, hỗ trợ âm thanh và chế độ sáng/tối.">
  
  <!-- GOOGLE FONT: BE VIETNAM PRO & JETBRAINS MONO -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:ital,wght@0,400;0,500;0,600;0,700;0,800;0,900;1,400&family=JetBrains+Mono:wght@500;600;700&display=swap" rel="stylesheet">
  
  <style>
    /* ==========================================================
       DESIGN SYSTEM (LIGHT / DARK THEMES)
    ========================================================== */
    :root, [data-theme="dark"] {{
      --font-base: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      --font-mono: 'JetBrains Mono', monospace;

      /* Dark Palette */
      --bg-canvas: #090d16;
      --bg-surface: #121824;
      --bg-card: #1c2436;
      --bg-card-hover: #263147;
      --border-line: rgba(240, 246, 252, 0.1);
      --border-accent: rgba(56, 139, 253, 0.4);

      /* Typography */
      --text-heading: #f0f6fc;
      --text-body: #c9d1d9;
      --text-muted: #8b949e;
      --text-dim: #6e7681;

      /* Brand Accents */
      --color-brand: #388bfd;
      --color-brand-hover: #1f6feb;
      --color-success: #238636;
      --color-success-light: #2ea043;
      --color-danger: #da3633;
      --color-danger-light: #f85149;
      --color-warning: #d29922;

      /* Functional Option Colors */
      --opt-a: #f85149;
      --opt-b: #388bfd;
      --opt-c: #d29922;
      --opt-d: #2ea043;

      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;
      --radius-pill: 9999px;

      --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.3);
      --shadow-md: 0 4px 14px rgba(0, 0, 0, 0.4);

      --safe-bottom: env(safe-area-inset-bottom, 16px);
      --safe-top: env(safe-area-inset-top, 0px);
    }}

    [data-theme="light"] {{
      /* Light Palette */
      --bg-canvas: #f6f8fa;
      --bg-surface: #ffffff;
      --bg-card: #f0f3f6;
      --bg-card-hover: #e4e9f0;
      --border-line: #d0d7de;
      --border-accent: rgba(9, 105, 218, 0.4);

      --text-heading: #1f2328;
      --text-body: #24292f;
      --text-muted: #57606a;
      --text-dim: #6e7781;

      --color-brand: #0969da;
      --color-brand-hover: #0550ae;
      --color-success: #1a7f37;
      --color-success-light: #2da44e;
      --color-danger: #cf222e;
      --color-danger-light: #fa4549;
      --color-warning: #9a6700;

      --opt-a: #cf222e;
      --opt-b: #0969da;
      --opt-c: #9a6700;
      --opt-d: #1a7f37;

      --shadow-sm: 0 1px 3px rgba(31, 35, 40, 0.12);
      --shadow-md: 0 3px 10px rgba(31, 35, 40, 0.15);
    }}

    * {{
      box-sizing: border-box; margin: 0; padding: 0;
      -webkit-tap-highlight-color: transparent;
    }}

    body {{
      font-family: var(--font-base);
      background-color: var(--bg-canvas);
      color: var(--text-body);
      min-height: 100vh;
      overflow-x: hidden;
      line-height: 1.6;
      padding-bottom: calc(76px + var(--safe-bottom));
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
      transition: background-color 0.2s ease, color 0.2s ease;
    }}

    .svg-ico {{
      width: 18px; height: 18px; display: inline-block; vertical-align: middle; flex-shrink: 0;
    }}

    .site-container {{
      max-width: 980px; margin: 0 auto;
      padding: clamp(14px, 3vw, 24px) clamp(12px, 3vw, 20px);
    }}

    /* Header Bar */
    .site-header {{
      text-align: center;
      padding: clamp(10px, 2.5vw, 20px) 0 clamp(16px, 3vw, 24px);
      border-bottom: 1px solid var(--border-line);
      margin-bottom: clamp(16px, 3vw, 24px);
      position: relative;
    }}
    .site-header-badge {{
      display: inline-flex; align-items: center; gap: 8px;
      background: rgba(56, 139, 253, 0.12); border: 1px solid rgba(56, 139, 253, 0.3);
      padding: 4px 14px; border-radius: var(--radius-pill);
      font-size: 12px; font-weight: 700; color: var(--color-brand);
      letter-spacing: 0.3px; margin-bottom: 12px;
    }}
    .status-dot {{
      width: 7px; height: 7px; border-radius: 50%; background: var(--color-success);
      box-shadow: 0 0 8px var(--color-success);
    }}
    .site-header h1 {{
      font-size: clamp(22px, 4.6vw, 32px);
      font-weight: 800;
      color: var(--text-heading);
      letter-spacing: -0.02em;
      line-height: 1.3;
      margin-bottom: 8px;
    }}
    .site-header p {{
      color: var(--text-muted);
      font-size: clamp(13px, 2.6vw, 15px);
      max-width: 680px; margin: 0 auto;
    }}

    /* Header Controls (Theme, Sound, Exam Level) */
    .header-controls {{
      display: flex; justify-content: center; align-items: center;
      gap: 10px; margin-top: 14px; flex-wrap: wrap;
    }}
    .control-btn {{
      background: var(--bg-surface); border: 1px solid var(--border-line);
      color: var(--text-heading); padding: 7px 14px; border-radius: var(--radius-pill);
      font-family: var(--font-base); font-size: 13px; font-weight: 600;
      cursor: pointer; display: inline-flex; align-items: center; gap: 6px;
      box-shadow: var(--shadow-sm); transition: all 0.15s ease;
    }}
    .control-btn:hover {{ background: var(--bg-card); }}

    /* Exam Level Pill Switcher */
    .level-switcher-bar {{
      display: inline-flex; background: var(--bg-surface); border: 1px solid var(--border-line);
      padding: 4px; border-radius: var(--radius-pill); gap: 4px; box-shadow: var(--shadow-sm);
    }}
    .level-pill {{
      background: transparent; border: none; color: var(--text-muted);
      padding: 6px 14px; border-radius: var(--radius-pill);
      font-family: var(--font-base); font-size: 13px; font-weight: 700;
      cursor: pointer; display: inline-flex; align-items: center; gap: 6px;
      transition: all 0.2s ease;
    }}
    .level-pill.active {{
      background: var(--color-brand); color: #fff; box-shadow: 0 2px 8px rgba(56, 139, 253, 0.4);
    }}

    /* Navigation Bar (Desktop) */
    .desk-nav {{
      display: flex; justify-content: center; gap: 8px;
      background: var(--bg-surface); border: 1px solid var(--border-line);
      padding: 6px; border-radius: var(--radius-md); margin-bottom: 22px;
      position: sticky; top: 12px; z-index: 90;
      box-shadow: var(--shadow-sm);
    }}
    .desk-nav-btn {{
      background: transparent; border: none; color: var(--text-muted);
      padding: 10px 18px; border-radius: var(--radius-sm);
      font-family: var(--font-base); font-size: 14px; font-weight: 600;
      cursor: pointer; display: flex; align-items: center; gap: 8px;
      transition: background 0.15s, color 0.15s;
    }}
    .desk-nav-btn:hover {{ color: var(--text-heading); background: rgba(0, 0, 0, 0.05); }}
    .desk-nav-btn.active {{
      background: var(--bg-card); color: var(--text-heading); border: 1px solid var(--border-line);
      box-shadow: var(--shadow-sm); font-weight: 700;
    }}
    .desk-nav-badge {{
      background: rgba(125, 133, 144, 0.2); padding: 2px 8px;
      border-radius: var(--radius-pill); font-size: 11px; font-weight: 700;
    }}

    /* Mobile Bottom Navigation */
    .mobile-bottom-bar {{
      display: none;
      position: fixed; bottom: 0; left: 0; right: 0; z-index: 999;
      background: var(--bg-surface); backdrop-filter: blur(16px);
      border-top: 1px solid var(--border-line);
      padding: 8px 12px calc(8px + var(--safe-bottom));
      justify-content: space-around; align-items: center;
      box-shadow: 0 -4px 16px rgba(0, 0, 0, 0.15);
    }}
    .mob-nav-item {{
      background: transparent; border: none; color: var(--text-dim);
      display: flex; flex-direction: column; align-items: center; gap: 4px;
      font-size: 11px; font-weight: 600; padding: 6px 14px; border-radius: var(--radius-sm);
      cursor: pointer; transition: color 0.15s;
    }}
    .mob-nav-item .svg-ico {{ width: 20px; height: 20px; }}
    .mob-nav-item.active {{ color: var(--color-brand); font-weight: 700; }}

    @media (max-width: 768px) {{
      .desk-nav {{ display: none; }}
      .mobile-bottom-bar {{ display: flex; }}
    }}

    /* Content Panels */
    .tab-panel {{ display: none; }}
    .tab-panel.active {{ display: block; animation: contentFadeIn 0.25s ease-out; }}
    @keyframes contentFadeIn {{
      from {{ opacity: 0; transform: translateY(6px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}

    /* Card Panels */
    .panel-card {{
      background: var(--bg-surface); border: 1px solid var(--border-line);
      border-radius: var(--radius-lg); padding: clamp(16px, 3.5vw, 28px);
      box-shadow: var(--shadow-sm); margin-bottom: 20px;
    }}

    /* Quiz Config Grid */
    .config-grid {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px;
      margin-bottom: 18px;
    }}
    .config-cell {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-md); padding: 14px;
    }}
    .config-cell-label {{
      font-size: 12px; font-weight: 700; color: var(--text-muted);
      text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 8px;
      display: flex; align-items: center; gap: 6px;
    }}
    .form-control {{
      width: 100%; background: var(--bg-surface); border: 1px solid var(--border-line);
      color: var(--text-heading); padding: 10px 12px; border-radius: var(--radius-sm);
      font-size: 14px; font-family: var(--font-base); font-weight: 500; outline: none;
      transition: border-color 0.15s;
    }}
    .form-control:focus {{ border-color: var(--color-brand); }}

    /* Settings Switch Rows */
    .settings-group {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-md); padding: 4px 16px; margin-bottom: 22px;
    }}
    .settings-row {{
      display: flex; align-items: center; justify-content: space-between;
      padding: 12px 0; border-bottom: 1px solid var(--border-line); gap: 12px;
    }}
    .settings-row:last-child {{ border-bottom: none; }}
    .settings-text strong {{
      font-size: 14px; color: var(--text-heading); display: block; font-weight: 600;
    }}
    .settings-text span {{
      font-size: 12px; color: var(--text-muted); display: block; margin-top: 1px;
    }}

    .toggle-switch {{
      position: relative; width: 44px; height: 24px; flex-shrink: 0;
    }}
    .toggle-switch input {{ opacity: 0; width: 0; height: 0; }}
    .toggle-slider {{
      position: absolute; inset: 0; cursor: pointer; background-color: #8c959f;
      transition: 0.2s; border-radius: 24px;
    }}
    .toggle-slider:before {{
      position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px;
      background-color: white; transition: 0.2s; border-radius: 50%;
    }}
    input:checked + .toggle-slider {{ background-color: var(--color-brand); }}
    input:checked + .toggle-slider:before {{ transform: translateX(20px); }}

    /* Action Buttons */
    .actions-row {{
      display: flex; gap: 12px; flex-wrap: wrap;
    }}
    .btn-action-primary {{
      flex: 2; min-width: 200px; min-height: 48px;
      background: var(--color-brand); border: none; color: #fff;
      padding: 12px 24px; border-radius: var(--radius-md);
      font-family: var(--font-base); font-size: 15px; font-weight: 700;
      cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px;
      transition: background 0.15s, transform 0.1s;
    }}
    .btn-action-primary:hover {{ background: var(--color-brand-hover); }}
    .btn-action-primary:active {{ transform: scale(0.98); }}

    .btn-action-secondary {{
      flex: 1; min-width: 160px; min-height: 48px;
      background: var(--bg-card); border: 1px solid var(--border-line);
      color: var(--text-heading); padding: 12px 18px; border-radius: var(--radius-md);
      font-family: var(--font-base); font-size: 14px; font-weight: 600;
      cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px;
      transition: background 0.15s;
    }}
    .btn-action-secondary:hover {{ background: var(--bg-card-hover); }}
    .btn-action-secondary:active {{ transform: scale(0.98); }}

    /* In-game HUD */
    .quiz-hud {{
      display: flex; align-items: center; justify-content: space-between;
      gap: 12px; margin-bottom: 14px; flex-wrap: wrap;
    }}
    .hud-track-wrap {{ flex: 1; min-width: 150px; }}
    .hud-meta-text {{
      display: flex; justify-content: space-between; font-size: 12px; font-weight: 700;
      color: var(--text-muted); margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px;
    }}
    .hud-bar {{
      height: 6px; background: rgba(125, 133, 144, 0.2); border-radius: var(--radius-pill);
      overflow: hidden;
    }}
    .hud-bar-inner {{
      height: 100%; width: 0%; background: var(--color-brand);
      border-radius: var(--radius-pill); transition: width 0.3s ease;
    }}

    .hud-pill {{
      display: inline-flex; align-items: center; gap: 6px;
      background: var(--bg-card); border: 1px solid var(--border-line);
      padding: 5px 12px; border-radius: var(--radius-pill);
      font-size: 13px; font-weight: 700; font-family: var(--font-base);
    }}
    .pill-timer {{
      font-family: var(--font-mono); color: var(--color-brand);
    }}
    .pill-timer.urgent {{
      color: var(--color-danger); border-color: rgba(248, 81, 73, 0.4); background: rgba(248, 81, 73, 0.1);
      animation: timerShake 0.5s infinite;
    }}
    @keyframes timerShake {{
      0%, 100% {{ transform: translateX(0); }}
      50% {{ transform: translateX(-2px); }}
    }}

    /* Question Surface */
    .question-surface {{
      background: var(--bg-surface); border: 1px solid var(--border-line);
      border-radius: var(--radius-lg); padding: clamp(18px, 3.5vw, 30px);
      margin-bottom: 16px;
    }}
    .q-meta-line {{
      display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;
    }}
    .q-badge-cat {{
      background: rgba(56, 139, 253, 0.15); border: 1px solid rgba(56, 139, 253, 0.3);
      color: var(--color-brand); padding: 3px 10px; border-radius: var(--radius-pill);
      font-size: 12px; font-weight: 700;
    }}
    .q-text-body {{
      font-size: clamp(16px, 3vw, 21px); font-weight: 700;
      color: var(--text-heading); line-height: 1.5;
    }}

    /* 4 Options Grid */
    .quiz-options-list {{
      display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px;
      margin-bottom: 16px;
    }}
    @media (max-width: 768px) {{
      .quiz-options-list {{ grid-template-columns: 1fr; gap: 10px; }}
    }}

    .opt-card {{
      background: var(--bg-surface); border: 1px solid var(--border-line);
      border-radius: var(--radius-md); padding: clamp(14px, 2.2vw, 18px);
      display: flex; align-items: center; gap: 14px; cursor: pointer; text-align: left;
      transition: background 0.15s, border-color 0.15s, transform 0.1s;
      min-height: 60px;
    }}
    .opt-card:hover:not(.disabled) {{
      background: var(--bg-card); border-color: rgba(125, 133, 144, 0.4);
    }}
    .opt-card:active:not(.disabled) {{
      transform: scale(0.98);
    }}

    .opt-indicator {{
      width: 34px; height: 34px; border-radius: 8px; flex-shrink: 0;
      display: flex; align-items: center; justify-content: center;
      font-size: 14px; font-weight: 800; font-family: var(--font-base);
      color: #fff;
    }}
    .opt-card.opt-0 .opt-indicator {{ background: var(--opt-a); }}
    .opt-card.opt-1 .opt-indicator {{ background: var(--opt-b); }}
    .opt-card.opt-2 .opt-indicator {{ background: var(--opt-c); }}
    .opt-card.opt-3 .opt-indicator {{ background: var(--opt-d); }}

    .opt-content {{ flex: 1; min-width: 0; }}
    .opt-hint {{
      font-size: 11px; font-weight: 700; color: var(--text-dim); text-transform: uppercase; margin-bottom: 2px;
    }}
    .opt-title {{
      font-size: clamp(14px, 2.6vw, 15px); font-weight: 600; color: var(--text-heading);
      line-height: 1.4; word-break: break-word;
    }}

    /* Answer States */
    .opt-card.is-correct {{
      background: rgba(35, 134, 54, 0.25) !important;
      border-color: var(--color-success) !important;
    }}
    .opt-card.is-wrong {{
      background: rgba(218, 54, 51, 0.25) !important;
      border-color: var(--color-danger) !important;
      animation: optShake 0.35s ease;
    }}
    @keyframes optShake {{
      0%, 100% {{ transform: translateX(0); }}
      25% {{ transform: translateX(-6px); }}
      75% {{ transform: translateX(6px); }}
    }}

    /* Feedback Sheet */
    .feedback-panel {{
      display: none; padding: 18px 20px; border-radius: var(--radius-md);
      margin-top: 14px; animation: fbSlideUp 0.2s ease-out;
    }}
    @keyframes fbSlideUp {{
      from {{ opacity: 0; transform: translateY(8px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    .feedback-panel.is-correct {{
      background: rgba(35, 134, 54, 0.15); border: 1px solid rgba(46, 160, 67, 0.4);
    }}
    .feedback-panel.is-wrong {{
      background: rgba(218, 54, 51, 0.15); border: 1px solid rgba(248, 81, 73, 0.4);
    }}
    .fb-title-row {{
      font-size: 16px; font-weight: 800; display: flex; align-items: center; gap: 8px; margin-bottom: 4px;
    }}
    .feedback-panel.is-correct .fb-title-row {{ color: var(--color-success); }}
    .feedback-panel.is-wrong .fb-title-row {{ color: var(--color-danger); }}
    .fb-body-text {{ font-size: 14px; color: var(--text-heading); margin-bottom: 8px; line-height: 1.5; }}
    .fb-note-box {{
      font-size: 13px; color: var(--color-brand); background: rgba(56, 139, 253, 0.1);
      padding: 8px 12px; border-radius: var(--radius-sm); margin-bottom: 12px;
    }}

    .btn-next-action {{
      width: 100%; min-height: 46px;
      background: var(--color-brand); border: none; color: #fff;
      padding: 10px 20px; border-radius: var(--radius-sm);
      font-family: var(--font-base); font-size: 14px; font-weight: 700;
      cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 6px;
      transition: background 0.15s;
    }}
    .btn-next-action:hover {{ background: var(--color-brand-hover); }}
    .btn-next-action:active {{ transform: scale(0.98); }}

    /* Score & Review View */
    .score-view {{ text-align: center; }}
    .score-circle {{
      width: 140px; height: 140px; margin: 0 auto 16px; position: relative;
    }}
    .score-svg {{ width: 100%; height: 100%; transform: rotate(-90deg); }}
    .score-bg-ring {{ fill: none; stroke: rgba(125, 133, 144, 0.2); stroke-width: 9; }}
    .score-fg-ring {{
      fill: none; stroke: var(--color-brand); stroke-width: 9; stroke-linecap: round;
      stroke-dasharray: 380; stroke-dashoffset: 380; transition: stroke-dashoffset 1s ease-out;
    }}
    .score-center-info {{
      position: absolute; inset: 0; display: flex; flex-direction: column;
      align-items: center; justify-content: center;
    }}
    .score-num {{ font-size: 32px; font-weight: 800; color: var(--text-heading); }}
    .score-label {{ font-size: 12px; color: var(--text-muted); font-weight: 600; }}

    .score-stats-grid {{
      display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin: 18px 0;
    }}
    @media (max-width: 600px) {{
      .score-stats-grid {{ grid-template-columns: repeat(2, 1fr); }}
    }}
    .score-box {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-sm); padding: 12px;
    }}
    .score-box-val {{ font-size: 20px; font-weight: 800; color: var(--text-heading); }}
    .score-box-lbl {{ font-size: 11px; color: var(--text-muted); text-transform: uppercase; font-weight: 700; }}

    .review-filters-row {{
      display: flex; gap: 8px; justify-content: center; margin-bottom: 14px;
    }}
    .review-filter-btn {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      color: var(--text-muted); padding: 6px 14px; border-radius: var(--radius-pill);
      font-size: 12px; font-weight: 600; cursor: pointer;
    }}
    .review-filter-btn.active {{
      background: var(--color-brand); color: #fff; font-weight: 700;
    }}

    .review-items-container {{
      max-height: 480px; overflow-y: auto; display: flex; flex-direction: column; gap: 10px;
      text-align: left; padding-right: 4px;
    }}
    .review-card {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-sm); padding: 14px; font-size: 13px;
    }}
    .review-card.correct {{ border-left: 3px solid var(--color-success); }}
    .review-card.wrong {{ border-left: 3px solid var(--color-danger); }}

    /* Mindmaps */
    .mm-carousel {{
      display: flex; gap: 8px; overflow-x: auto; padding-bottom: 8px;
      margin-bottom: 16px; scroll-snap-type: x mandatory;
    }}
    .mm-carousel::-webkit-scrollbar {{ height: 4px; }}
    .mm-carousel::-webkit-scrollbar-thumb {{ background: rgba(125, 133, 144, 0.3); border-radius: 4px; }}

    .mm-tab-item {{
      flex-shrink: 0; scroll-snap-align: start;
      background: var(--bg-card); border: 1px solid var(--border-line);
      color: var(--text-muted); padding: 8px 14px; border-radius: var(--radius-pill);
      font-size: 13px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px;
      transition: background 0.15s, color 0.15s;
    }}
    .mm-tab-item.active {{
      background: rgba(56, 139, 253, 0.15); border-color: var(--color-brand); color: var(--color-brand); font-weight: 700;
    }}

    .search-row {{
      display: flex; gap: 10px; margin-bottom: 16px;
    }}
    .search-input {{
      flex: 1; background: var(--bg-card); border: 1px solid var(--border-line);
      color: var(--text-heading); padding: 10px 14px; border-radius: var(--radius-sm);
      font-size: 14px; outline: none; font-family: var(--font-base);
    }}
    .search-input:focus {{ border-color: var(--color-brand); }}

    .tree-stack {{
      display: flex; flex-direction: column; gap: 12px;
    }}
    .branch-card {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-md); overflow: hidden;
    }}
    .branch-head {{
      padding: 14px 16px; background: rgba(0, 0, 0, 0.03);
      display: flex; align-items: center; justify-content: space-between; cursor: pointer;
    }}
    .branch-title {{
      font-size: 15px; font-weight: 700; color: var(--text-heading);
      display: flex; align-items: center; gap: 8px;
    }}
    .branch-toggle {{
      color: var(--text-muted); transition: transform 0.2s;
    }}
    .branch-card.collapsed .branch-toggle {{ transform: rotate(-90deg); }}
    
    .branch-body {{
      display: grid; grid-template-rows: 1fr; transition: grid-template-rows 0.25s ease-out;
    }}
    .branch-card.collapsed .branch-body {{ grid-template-rows: 0fr; }}
    .branch-inner {{
      overflow: hidden; padding: 0 14px 14px;
      display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px;
    }}
    .leaf-item {{
      background: var(--bg-surface); border: 1px solid var(--border-line);
      border-radius: var(--radius-sm); padding: 12px; font-size: 13px;
    }}
    .leaf-label {{
      font-size: 12px; font-weight: 700; color: var(--color-brand); margin-bottom: 4px;
      text-transform: uppercase; letter-spacing: 0.3px;
    }}
    .leaf-desc {{ color: var(--text-body); line-height: 1.5; }}

    /* Directory & Cheatsheet */
    .q-scroll-box {{
      max-height: 580px; overflow-y: auto; display: flex; flex-direction: column; gap: 10px;
      padding-right: 4px;
    }}
    .q-row-card {{
      background: var(--bg-card); border: 1px solid var(--border-line);
      border-radius: var(--radius-md); padding: 14px; font-size: 13px;
    }}
    .q-row-tag {{
      background: rgba(56, 139, 253, 0.12); color: var(--color-brand); padding: 2px 7px;
      border-radius: var(--radius-pill); font-size: 11px; font-weight: 700;
    }}
    .q-opts-grid {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 6px; margin-top: 8px;
    }}
    .q-opt-item {{
      padding: 7px 10px; border-radius: var(--radius-sm);
      background: rgba(0, 0, 0, 0.03); color: var(--text-muted); font-size: 12px;
    }}
    .q-opt-item.is-correct {{
      background: rgba(35, 134, 54, 0.2); border: 1px solid rgba(46, 160, 67, 0.4);
      color: var(--color-success); font-weight: 700;
    }}
    .q-note-hint {{
      margin-top: 8px; font-size: 12px; color: var(--color-brand);
      background: rgba(56, 139, 253, 0.08); padding: 6px 10px; border-radius: var(--radius-sm);
    }}

    /* Canvas Confetti */
    #confettiCanvas {{
      position: fixed; inset: 0; pointer-events: none; z-index: 9999;
    }}
  </style>
</head>
<body>
  <canvas id="confettiCanvas"></canvas>

  <div class="site-container">
    <!-- Main Header -->
    <header class="site-header">
      <div class="site-header-badge" id="headerBadge">
        <span class="status-dot"></span>
        <span id="headerBadgeText">CHỨNG CHỈ CNTT NÂNG CAO • 201 CÂU CHUẨN ĐÁP ÁN</span>
      </div>

      <h1 id="headerMainTitle">Trắc Nghiệm CNTT Nâng Cao</h1>
      <p id="headerSubtitle">Luyện thi trắc nghiệm tương tác theo phong cách Quizizz cho kỳ thi Ứng dụng CNTT Nâng cao (Word, Excel, Access), kèm giải thích đáp án và Sơ đồ tư duy kiến thức.</p>

      <!-- Top Controls: Level Switcher, Theme, Audio -->
      <div class="header-controls">
        <div class="level-switcher-bar">
          <button class="level-pill active" id="btnLevelNC" onclick="switchExamLevel('NC')">
            {ICONS["star"]}
            <span>Đề Nâng Cao (201 câu)</span>
          </button>
          <button class="level-pill" id="btnLevelCB" onclick="switchExamLevel('CB')">
            {ICONS["book"]}
            <span>Đề Cơ Bản (288 câu)</span>
          </button>
        </div>

        <button class="control-btn" id="soundToggleBtn" onclick="toggleSound()" title="Bật / Tắt âm thanh">
          <span id="soundToggleIcon">{ICONS["volume"]}</span>
          <span id="soundToggleText">Âm thanh: Bật</span>
        </button>

        <button class="control-btn" id="themeToggleBtn" onclick="toggleTheme()" title="Chuyển chế độ Sáng / Tối">
          <span id="themeToggleIcon">{ICONS["sun"]}</span>
          <span id="themeToggleText">Chế độ Sáng</span>
        </button>
      </div>
    </header>

    <!-- Top Desktop Nav Bar -->
    <nav class="desk-nav">
      <button class="desk-nav-btn active" onclick="switchMainTab('quiz')">
        {ICONS["quiz"]}
        <span>Luyện Thi Quizizz</span>
        <span class="desk-nav-badge" id="deskCountBadge">201 câu</span>
      </button>
      <button class="desk-nav-btn" onclick="switchMainTab('mindmap')">
        {ICONS["mindmap"]}
        <span>Sơ Đồ Tư Duy (Mindmaps)</span>
        <span class="desk-nav-badge" id="deskModuleBadge">3 Module</span>
      </button>
      <button class="desk-nav-btn" onclick="switchMainTab('cheatsheet')">
        {ICONS["book"]}
        <span>Ngân Hàng Câu Hỏi & Tra Cứu</span>
        <span class="desk-nav-badge">Chuẩn Đáp Án</span>
      </button>
    </nav>

    <!-- ==========================================================
         PANEL 1: QUIZ ENGINE
    ========================================================== -->
    <main id="panel-quiz" class="tab-panel active">
      <!-- 1.1 SETUP SCREEN -->
      <section id="quizHome" class="panel-card">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px; flex-wrap: wrap; gap: 8px;">
          <h2 style="font-size: 18px; font-weight: 800; color: var(--text-heading); display: flex; align-items: center; gap: 8px;">
            {ICONS["settings"]} Thiết Lập Bài Luyện Thi
          </h2>
          <span style="font-size: 13px; color: var(--text-muted);" id="setupExamLevelLabel">Đang chọn: Đề thi Nâng Cao</span>
        </div>

        <div class="config-grid">
          <div class="config-cell">
            <div class="config-cell-label">{ICONS["book"]} Chọn Module / Môn Học</div>
            <select id="selModule" class="form-control" onchange="updateQuestionCountLimit()">
              <!-- Generated dynamically by JS -->
            </select>
          </div>

          <div class="config-cell">
            <div class="config-cell-label">{ICONS["quiz"]} Số Lượng Câu Hỏi</div>
            <select id="selLimit" class="form-control">
              <option value="10">Luyện nhanh 10 câu</option>
              <option value="20" selected>Luyện chuẩn 20 câu</option>
              <option value="30">Luyện sâu 30 câu</option>
              <option value="40">Mô phỏng thi 40 câu</option>
              <option value="50">Luyện dày 50 câu</option>
              <option value="100">Cày cuốc 100 câu</option>
              <option value="ALL">Làm toàn bộ câu hỏi</option>
            </select>
          </div>

          <div class="config-cell">
            <div class="config-cell-label">{ICONS["timer"]} Thời Gian / Câu</div>
            <select id="selTimer" class="form-control">
              <option value="0" selected>Không giới hạn (Ôn kỹ)</option>
              <option value="15">15 giây / câu (Phản xạ tốc độ)</option>
              <option value="30">30 giây / câu (Chuẩn thi)</option>
              <option value="60">60 giây / câu (Thư thả)</option>
            </select>
          </div>
        </div>

        <!-- Toggles -->
        <div class="settings-group">
          <div class="settings-row">
            <div class="settings-text">
              <strong>Xáo trộn câu hỏi ngẫu nhiên</strong>
              <span>Đổi vị trí xuất hiện các câu hỏi khi bắt đầu bài thi</span>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="chkShuffleQ" checked>
              <span class="toggle-slider"></span>
            </label>
          </div>

          <div class="settings-row">
            <div class="settings-text">
              <strong>Xáo trộn vị trí đáp án (A, B, C, D)</strong>
              <span>Tránh học vẹt vị trí nút chọn, rèn luyện đọc hiểu bản chất</span>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="chkShuffleAns" checked>
              <span class="toggle-slider"></span>
            </label>
          </div>

          <div class="settings-row">
            <div class="settings-text">
              <strong>Tự động chuyển câu khi trả lời đúng</strong>
              <span>Tiết kiệm thời gian thao tác (chỉ áp dụng khi chọn đúng)</span>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="chkAutoNext">
              <span class="toggle-slider"></span>
            </label>
          </div>
        </div>

        <div class="actions-row">
          <button class="btn-action-primary" onclick="startQuiz(false)">
            {ICONS["play"]} <span>Bắt Đầu Luyện Thi</span>
          </button>
          <button class="btn-action-secondary" id="btnWrongQuiz" onclick="startQuiz(true)">
            {ICONS["refresh"]} <span>Ôn Lại Câu Sai</span>
            <span class="desk-nav-badge" id="badgeWrongCount">0</span>
          </button>
        </div>
      </section>

      <!-- 1.2 IN-GAME QUIZ VIEW -->
      <section id="quizPlay" style="display: none;">
        <!-- HUD Tracker -->
        <div class="quiz-hud">
          <div class="hud-track-wrap">
            <div class="hud-meta-text">
              <span id="hudProgressText">Câu 1 / 20</span>
              <span id="hudLevelBadge" style="color: var(--color-brand); font-weight: 700;">Nâng cao</span>
            </div>
            <div class="hud-bar">
              <div class="hud-bar-inner" id="hudProgressBar"></div>
            </div>
          </div>

          <div style="display: flex; gap: 8px; align-items: center;">
            <div class="hud-pill" id="hudStreakPill">
              {ICONS["streak"]}
              <span id="hudStreakCount" style="color: #f85149;">0</span>
            </div>
            <div class="hud-pill pill-timer" id="hudTimerPill">
              {ICONS["timer"]}
              <span id="hudTimerCount">--</span>
            </div>
            <button class="hud-pill" onclick="quitQuizConfirm()" style="cursor: pointer; background: transparent; color: var(--text-dim);" title="Dừng bài thi">
              {ICONS["cross"]}
            </button>
          </div>
        </div>

        <!-- Question Box -->
        <div class="question-surface">
          <div class="q-meta-line">
            <span class="q-badge-cat" id="qCategoryBadge">Module</span>
            <span style="font-size: 12px; color: var(--text-dim); font-family: var(--font-mono);" id="qOriginalNum">#0</span>
          </div>
          <div class="q-text-body" id="qQuestionText">
            Nội dung câu hỏi hiển thị tại đây...
          </div>
        </div>

        <!-- 4 Answer Options -->
        <div class="quiz-options-list" id="qOptionsContainer">
          <!-- Rendered dynamically -->
        </div>

        <!-- Immediate Feedback Panel -->
        <div class="feedback-panel" id="quizFeedback">
          <div class="fb-title-row" id="fbTitle">
            <!-- Icon + text -->
          </div>
          <div class="fb-body-text" id="fbText"></div>
          <div class="fb-note-box" id="fbNoteBox" style="display: none;"></div>
          <button class="btn-next-action" id="btnNextQuestion" onclick="nextQuestion()">
            <span>Tiếp Tục</span> {ICONS["arrowRight"]}
          </button>
        </div>
      </section>

      <!-- 1.3 SCORE & REVIEW VIEW -->
      <section id="quizResult" class="panel-card score-view" style="display: none;">
        <div class="score-circle">
          <svg class="score-svg" viewBox="0 0 140 140">
            <circle class="score-bg-ring" cx="70" cy="70" r="60"></circle>
            <circle class="score-fg-ring" id="scoreFgRing" cx="70" cy="70" r="60"></circle>
          </svg>
          <div class="score-center-info">
            <div class="score-num" id="scorePercent">0%</div>
            <div class="score-label" id="scoreGrade">Hoàn thành</div>
          </div>
        </div>

        <h2 style="font-size: 22px; font-weight: 800; color: var(--text-heading); margin-bottom: 4px;" id="scoreHeadline">
          Kết Quả Bài Luyện Thi
        </h2>
        <p style="font-size: 14px; color: var(--text-muted); max-width: 500px; margin: 0 auto;" id="scoreSubline">
          Hãy xem lại chi tiết các câu đã làm để củng cố kiến thức trước ngày thi!
        </p>

        <!-- Stats Grid -->
        <div class="score-stats-grid">
          <div class="score-box">
            <div class="score-box-val" id="resCorrect" style="color: var(--color-success);">0</div>
            <div class="score-box-lbl">Đúng</div>
          </div>
          <div class="score-box">
            <div class="score-box-val" id="resWrong" style="color: var(--color-danger);">0</div>
            <div class="score-box-lbl">Sai</div>
          </div>
          <div class="score-box">
            <div class="score-box-val" id="resStreak" style="color: #f85149;">0</div>
            <div class="score-box-lbl">Chuỗi Max</div>
          </div>
          <div class="score-box">
            <div class="score-box-val" id="resTime" style="color: var(--color-brand); font-family: var(--font-mono);">00:00</div>
            <div class="score-box-lbl">Thời Gian</div>
          </div>
        </div>

        <div class="actions-row" style="margin-bottom: 24px;">
          <button class="btn-action-primary" onclick="restartCurrentQuiz()">
            {ICONS["refresh"]} <span>Làm Lại Đề Này</span>
          </button>
          <button class="btn-action-secondary" id="btnReviewWrongQuiz" onclick="startWrongQuizFromSummary()">
            {ICONS["bulb"]} <span>Luyện Lại Câu Sai Vừa Làm</span>
          </button>
          <button class="btn-action-secondary" onclick="returnQuizHome()">
            <span>Về Trang Thiết Lập</span>
          </button>
        </div>

        <!-- Review Detailed List -->
        <div style="border-top: 1px solid var(--border-line); padding-top: 18px; text-align: left;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px;">
            <h3 style="font-size: 16px; font-weight: 700; color: var(--text-heading);">
              Chi Tiết Từng Câu Hỏi
            </h3>
            <div class="review-filters-row">
              <button class="review-filter-btn active" onclick="filterReview('ALL', this)">Tất cả (<span id="revCountAll">0</span>)</button>
              <button class="review-filter-btn" onclick="filterReview('WRONG', this)">Câu sai (<span id="revCountWrong">0</span>)</button>
              <button class="review-filter-btn" onclick="filterReview('CORRECT', this)">Câu đúng (<span id="revCountCorrect">0</span>)</button>
            </div>
          </div>

          <div class="review-items-container" id="reviewContainer">
            <!-- Rendered dynamically -->
          </div>
        </div>
      </section>
    </main>

    <!-- ==========================================================
         PANEL 2: MINDMAPS (SƠ ĐỒ TƯ DUY)
    ========================================================== -->
    <section id="panel-mindmap" class="tab-panel">
      <div class="panel-card">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; flex-wrap: wrap; gap: 8px;">
          <div>
            <h2 style="font-size: 18px; font-weight: 800; color: var(--text-heading); display: flex; align-items: center; gap: 8px;">
              {ICONS["mindmap"]} Sơ Đồ Tư Duy Chuyên Sâu
            </h2>
            <p style="font-size: 13px; color: var(--text-muted); margin-top: 2px;" id="mindmapSubText">
              Hệ thống hóa toàn bộ kiến thức trọng tâm cho kỳ thi
            </p>
          </div>
          <div style="display: flex; gap: 8px;">
            <button class="control-btn" onclick="expandAllMindmap(true)">Mở tất cả</button>
            <button class="control-btn" onclick="expandAllMindmap(false)">Thu gọn</button>
          </div>
        </div>

        <!-- Horizontal Module Carousel -->
        <div class="mm-carousel" id="mindmapModuleTabs">
          <!-- Rendered dynamically -->
        </div>

        <!-- Search in Mindmap -->
        <div class="search-row">
          <input type="text" class="search-input" id="mindmapSearchInput" placeholder="Tìm kiếm nhanh trong sơ đồ tư duy (ví dụ: Mail Merge, Advanced Filter, accdb, Short Date...)..." oninput="handleMindmapSearch(this.value)">
        </div>

        <!-- Mindmap Tree Container -->
        <div class="tree-stack" id="mindmapTreeContainer">
          <!-- Rendered dynamically -->
        </div>
      </div>
    </section>

    <!-- ==========================================================
         PANEL 3: DIRECTORY / CHEATSHEET
    ========================================================== -->
    <section id="panel-cheatsheet" class="tab-panel">
      <div class="panel-card">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; flex-wrap: wrap; gap: 8px;">
          <div>
            <h2 style="font-size: 18px; font-weight: 800; color: var(--text-heading); display: flex; align-items: center; gap: 8px;">
              {ICONS["book"]} Ngân Hàng Câu Hỏi & Chuẩn Đáp Án
            </h2>
            <p style="font-size: 13px; color: var(--text-muted); margin-top: 2px;">
              Tra cứu nhanh toàn bộ câu hỏi, phương án và đáp án chuẩn xác
            </p>
          </div>
        </div>

        <!-- Module Selector Tabs -->
        <div class="mm-carousel" id="dirModuleTabs">
          <!-- Rendered dynamically -->
        </div>

        <!-- Search Bar -->
        <div class="search-row">
          <input type="text" class="search-input" id="dirSearchInput" placeholder="Tìm kiếm theo từ khóa câu hỏi hoặc đáp án..." oninput="filterQuestionsDirectory(this.value)">
        </div>

        <!-- List -->
        <div class="q-scroll-box" id="questionsDirectoryList">
          <!-- Rendered dynamically -->
        </div>
      </div>
    </section>
  </div>

  <!-- Mobile Bottom Bar -->
  <nav class="mobile-bottom-bar">
    <button class="mob-nav-item active" onclick="switchMainTab('quiz')">
      {ICONS["quiz"]}
      <span>Luyện Thi</span>
    </button>
    <button class="mob-nav-item" onclick="switchMainTab('mindmap')">
      {ICONS["mindmap"]}
      <span>Sơ Đồ Tư Duy</span>
    </button>
    <button class="mob-nav-item" onclick="switchMainTab('cheatsheet')">
      {ICONS["book"]}
      <span>Tra Cứu</span>
    </button>
  </nav>

  <script>
    // Embedded Data Sets
    const DATA_NC = {json_nc_str};
    const DATA_CB = {json_cb_str};
    const MINDMAPS_NC = {mindmaps_nc_str};
    const MINDMAPS_CB = {mindmaps_cb_str};

    const JS_ICONS = {{
      check: `{ICONS["check"]}`,
      cross: `{ICONS["cross"]}`,
      chevronDown: `{ICONS["chevronDown"]}`,
      star: `{ICONS["star"]}`,
      sun: `{ICONS["sun"]}`,
      moon: `{ICONS["moon"]}`,
      volume: `{ICONS["volume"]}`
    }};

    // Current State
    let currentExamLevel = 'NC'; // 'NC' (Nâng Cao) or 'CB' (Cơ Bản)
    let rawQuestions = DATA_NC;
    let mindmapsData = MINDMAPS_NC;

    let soundEnabled = true;
    let activeMainTab = 'quiz';

    // Quiz Session State
    let quizPool = [];
    let currentQIdx = 0;
    let currentStreak = 0;
    let maxStreak = 0;
    let correctCount = 0;
    let wrongCount = 0;
    let questionTimer = null;
    let secondsLeft = 0;
    let quizStartTime = 0;
    let quizElapsedTime = 0;
    let sessionResults = [];
    let isAnsweringLocked = false;
    let currentReviewFilter = 'ALL';

    // Local Storage Keys
    function getWrongStorageKey() {{
      return currentExamLevel === 'NC' ? 'quiz_wrong_ids_nc' : 'quiz_wrong_ids_cb';
    }}

    function getWrongQuestionsList() {{
      try {{
        const raw = localStorage.getItem(getWrongStorageKey());
        return raw ? JSON.parse(raw) : [];
      }} catch (e) {{
        return [];
      }}
    }}

    function saveWrongQuestionId(id) {{
      try {{
        let list = getWrongQuestionsList();
        if (!list.includes(id)) {{
          list.push(id);
          localStorage.setItem(getWrongStorageKey(), JSON.stringify(list));
        }}
      }} catch (e) {{}}
    }}

    function removeWrongQuestionId(id) {{
      try {{
        let list = getWrongQuestionsList().filter(x => x !== id);
        localStorage.setItem(getWrongStorageKey(), JSON.stringify(list));
      }} catch (e) {{}}
    }}

    // --- SYNTHESIZED AUDIO ENGINE (Web Audio API - 100% Offline) ---
    let audioCtx = null;
    function getAudioContext() {{
      if (!audioCtx) {{
        const AudioContext = window.AudioContext || window.webkitAudioContext;
        if (AudioContext) audioCtx = new AudioContext();
      }}
      if (audioCtx && audioCtx.state === 'suspended') {{
        audioCtx.resume();
      }}
      return audioCtx;
    }}

    function playSound(type) {{
      if (!soundEnabled) return;
      try {{
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;

        if (type === 'correct') {{
          // Joyful Chime (E5 -> G#5 -> B5)
          const osc1 = ctx.createOscillator();
          const osc2 = ctx.createOscillator();
          const gain = ctx.createGain();

          osc1.type = 'triangle';
          osc2.type = 'sine';

          osc1.frequency.setValueAtTime(659.25, now); // E5
          osc1.frequency.exponentialRampToValueAtTime(987.77, now + 0.18); // B5

          osc2.frequency.setValueAtTime(830.61, now + 0.05); // G#5
          osc2.frequency.exponentialRampToValueAtTime(1318.51, now + 0.22); // E6

          gain.gain.setValueAtTime(0.18, now);
          gain.gain.exponentialRampToValueAtTime(0.001, now + 0.35);

          osc1.connect(gain);
          osc2.connect(gain);
          gain.connect(ctx.destination);

          osc1.start(now);
          osc2.start(now + 0.05);
          osc1.stop(now + 0.35);
          osc2.stop(now + 0.35);

        }} else if (type === 'wrong') {{
          // Gentle Low Buzz
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();

          osc.type = 'sawtooth';
          osc.frequency.setValueAtTime(180, now);
          osc.frequency.linearRampToValueAtTime(120, now + 0.25);

          gain.gain.setValueAtTime(0.2, now);
          gain.gain.exponentialRampToValueAtTime(0.001, now + 0.28);

          osc.connect(gain);
          gain.connect(ctx.destination);

          osc.start(now);
          osc.stop(now + 0.28);

        }} else if (type === 'click') {{
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();

          osc.type = 'sine';
          osc.frequency.setValueAtTime(500, now);
          gain.gain.setValueAtTime(0.05, now);
          gain.gain.exponentialRampToValueAtTime(0.001, now + 0.05);

          osc.connect(gain);
          gain.connect(ctx.destination);

          osc.start(now);
          osc.stop(now + 0.05);

        }} else if (type === 'fanfare') {{
          // Finish celebration chord
          [523.25, 659.25, 783.99, 1046.50].forEach((freq, i) => {{
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(freq, now + i * 0.08);

            gain.gain.setValueAtTime(0.12, now + i * 0.08);
            gain.gain.exponentialRampToValueAtTime(0.001, now + i * 0.08 + 0.5);

            osc.connect(gain);
            gain.connect(ctx.destination);

            osc.start(now + i * 0.08);
            osc.stop(now + i * 0.08 + 0.5);
          }});
        }}
      }} catch (e) {{}}
    }}

    function toggleSound() {{
      soundEnabled = !soundEnabled;
      const text = document.getElementById('soundToggleText');
      text.innerText = soundEnabled ? 'Âm thanh: Bật' : 'Âm thanh: Tắt';
      playSound('click');
    }}

    function triggerHaptic(duration = 20) {{
      if (window.navigator && window.navigator.vibrate) {{
        window.navigator.vibrate(duration);
      }}
    }}

    // --- THEME MANAGEMENT ---
    function initTheme() {{
      const saved = localStorage.getItem('quiz_theme') || 'dark';
      document.documentElement.setAttribute('data-theme', saved);
      updateThemeUI(saved);
    }}

    function toggleTheme() {{
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('quiz_theme', next);
      updateThemeUI(next);
      playSound('click');
    }}

    function updateThemeUI(theme) {{
      const icon = document.getElementById('themeToggleIcon');
      const text = document.getElementById('themeToggleText');
      if (theme === 'dark') {{
        icon.innerHTML = JS_ICONS.sun;
        text.innerText = 'Chế độ Sáng';
      }} else {{
        icon.innerHTML = JS_ICONS.moon;
        text.innerText = 'Chế độ Tối';
      }}
    }}

    // --- EXAM LEVEL SWITCHER ---
    function switchExamLevel(level) {{
      if (currentExamLevel === level) return;
      currentExamLevel = level;
      playSound('click');

      const btnNC = document.getElementById('btnLevelNC');
      const btnCB = document.getElementById('btnLevelCB');
      const badge = document.getElementById('headerBadgeText');
      const mainTitle = document.getElementById('headerMainTitle');
      const subtitle = document.getElementById('headerSubtitle');
      const deskCountBadge = document.getElementById('deskCountBadge');
      const deskModuleBadge = document.getElementById('deskModuleBadge');
      const setupLabel = document.getElementById('setupExamLevelLabel');
      const mindmapSubText = document.getElementById('mindmapSubText');

      if (level === 'NC') {{
        btnNC.classList.add('active');
        btnCB.classList.remove('active');
        rawQuestions = DATA_NC;
        mindmapsData = MINDMAPS_NC;

        badge.innerText = 'CHỨNG CHỈ CNTT NÂNG CAO • 201 CÂU CHUẨN ĐÁP ÁN';
        mainTitle.innerText = 'Trắc Nghiệm CNTT Nâng Cao';
        subtitle.innerText = 'Luyện thi trắc nghiệm tương tác theo phong cách Quizizz cho kỳ thi Ứng dụng CNTT Nâng cao (Word, Excel, Access), kèm giải thích đáp án và Sơ đồ tư duy kiến thức.';
        deskCountBadge.innerText = '201 câu';
        deskModuleBadge.innerText = '3 Module';
        setupLabel.innerText = 'Đang chọn: Đề thi Nâng Cao (201 câu)';
        mindmapSubText.innerText = '3 Module nâng cao: Word nâng cao (Mailings, References), Excel nâng cao (Advanced Filter, DSUM), Access nâng cao (Table, Query, Form, Report).';
      }} else {{
        btnCB.classList.add('active');
        btnNC.classList.remove('active');
        rawQuestions = DATA_CB;
        mindmapsData = MINDMAPS_CB;

        badge.innerText = 'CHỨNG CHỈ CNTT CƠ BẢN • 288 CÂU ĐỀ CƯƠNG';
        mainTitle.innerText = 'Trắc Nghiệm CNTT Cơ Bản';
        subtitle.innerText = 'Luyện thi trắc nghiệm tương tác chuẩn mực cho kỳ thi Ứng dụng CNTT Cơ bản (Kiến thức chung, Windows 10, Word, Excel, PowerPoint, LAN & Internet).';
        deskCountBadge.innerText = '288 câu';
        deskModuleBadge.innerText = '6 Module';
        setupLabel.innerText = 'Đang chọn: Đề thi Cơ Bản (288 câu)';
        mindmapSubText.innerText = '6 Module cơ bản: Kiến thức chung, Windows 10, Word cơ bản, Excel cơ bản, PowerPoint, LAN & Internet.';
      }}

      populateModuleDropdown();
      updateWrongBadgeUI();
      renderMindmaps();
      renderQuestionsDirectory();
      returnQuizHome();
    }}

    function populateModuleDropdown() {{
      const sel = document.getElementById('selModule');
      sel.innerHTML = '';

      if (currentExamLevel === 'NC') {{
        sel.innerHTML = `
          <option value="ALL">Toàn bộ 3 Module Nâng Cao (201 câu)</option>
          <option value="MS Winword 2013">Module 7: MS Word 2013 Nâng Cao (130 câu)</option>
          <option value="MS Excel 2013">Module 8: MS Excel 2013 Nâng Cao (31 câu)</option>
          <option value="MS Access 2013">Module 9: MS Access 2013 Nâng Cao (40 câu)</option>
        `;
      }} else {{
        sel.innerHTML = `
          <option value="ALL">Toàn bộ 6 Module Cơ Bản (288 câu)</option>
          <option value="Kiến thức chung">Module 1: Kiến thức chung (47 câu)</option>
          <option value="Windows 10">Module 2: Windows 10 (30 câu)</option>
          <option value="MS Word 2013">Module 3: MS Word 2013 (60 câu)</option>
          <option value="MS Excel 2013">Module 4: MS Excel 2013 (96 câu)</option>
          <option value="MS PowerPoint 2013">Module 5: MS PowerPoint 2013 (30 câu)</option>
          <option value="LAN & Internet">Module 6: LAN & Internet (25 câu)</option>
        `;
      }}
      updateQuestionCountLimit();
    }}

    function updateQuestionCountLimit() {{
      const mod = document.getElementById('selModule').value;
      const count = mod === 'ALL' ? rawQuestions.length : rawQuestions.filter(q => q.module === mod).length;
      const optAll = document.querySelector('#selLimit option[value="ALL"]');
      if (optAll) {{
        optAll.innerText = `Làm toàn bộ môn này (${{count}} câu)`;
      }}
    }}

    function updateWrongBadgeUI() {{
      const wrongList = getWrongQuestionsList();
      const badge = document.getElementById('badgeWrongCount');
      const btn = document.getElementById('btnWrongQuiz');
      badge.innerText = wrongList.length;
      if (wrongList.length === 0) {{
        btn.style.opacity = '0.5';
      }} else {{
        btn.style.opacity = '1';
      }}
    }}

    // --- NAVIGATION TABS ---
    function switchMainTab(tab) {{
      activeMainTab = tab;
      playSound('click');

      document.querySelectorAll('.desk-nav-btn').forEach((b, i) => {{
        const tabs = ['quiz', 'mindmap', 'cheatsheet'];
        if (tabs[i] === tab) b.classList.add('active');
        else b.classList.remove('active');
      }});

      document.querySelectorAll('.mob-nav-item').forEach((b, i) => {{
        const tabs = ['quiz', 'mindmap', 'cheatsheet'];
        if (tabs[i] === tab) b.classList.add('active');
        else b.classList.remove('active');
      }});

      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      const activePanel = document.getElementById(`panel-${{tab}}`);
      if (activePanel) activePanel.classList.add('active');
      window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }}

    // --- SHUFFLE UTILS ---
    function shuffleArray(arr) {{
      const copy = [...arr];
      for (let i = copy.length - 1; i > 0; i--) {{
        const j = Math.floor(Math.random() * (i + 1));
        [copy[i], copy[j]] = [copy[j], copy[i]];
      }}
      return copy;
    }}

    // --- QUIZ LAUNCH & FLOW ---
    function startQuiz(isWrongOnly = false) {{
      playSound('click');
      triggerHaptic(20);

      const mod = document.getElementById('selModule').value;
      const limitVal = document.getElementById('selLimit').value;
      const shuffleQ = document.getElementById('chkShuffleQ').checked;

      let pool = [...rawQuestions];

      if (isWrongOnly) {{
        const wrongIds = getWrongQuestionsList();
        if (wrongIds.length === 0) {{
          alert('Bạn chưa có câu làm sai nào trong danh sách ôn tập của đề này! Hãy luyện thi thông thường trước.');
          return;
        }}
        pool = pool.filter(q => wrongIds.includes(q.id));
      }} else if (mod !== 'ALL') {{
        pool = pool.filter(q => q.module === mod);
      }}

      if (pool.length === 0) {{
        alert('Không tìm thấy câu hỏi phù hợp.');
        return;
      }}

      if (shuffleQ) {{
        pool = shuffleArray(pool);
      }}

      if (limitVal !== 'ALL') {{
        const num = parseInt(limitVal, 10);
        if (!isNaN(num) && num > 0) {{
          pool = pool.slice(0, num);
        }}
      }}

      quizPool = pool;
      currentQIdx = 0;
      currentStreak = 0;
      maxStreak = 0;
      correctCount = 0;
      wrongCount = 0;
      sessionResults = [];
      quizStartTime = Date.now();

      document.getElementById('quizHome').style.display = 'none';
      document.getElementById('quizResult').style.display = 'none';
      document.getElementById('quizPlay').style.display = 'block';

      renderCurrentQuestion();
    }}

    function renderCurrentQuestion() {{
      clearInterval(questionTimer);
      isAnsweringLocked = false;

      if (currentQIdx >= quizPool.length) {{
        finishQuiz();
        return;
      }}

      const qData = quizPool[currentQIdx];
      const shuffleAns = document.getElementById('chkShuffleAns').checked;

      // HUD Update
      document.getElementById('hudProgressText').innerText = `Câu ${{currentQIdx + 1}} / ${{quizPool.length}}`;
      document.getElementById('hudLevelBadge').innerText = currentExamLevel === 'NC' ? '⭐ Nâng Cao' : '📘 Cơ Bản';
      const pct = ((currentQIdx) / quizPool.length) * 100;
      document.getElementById('hudProgressBar').style.width = `${{pct}}%`;

      document.getElementById('hudStreakCount').innerText = currentStreak;
      const streakPill = document.getElementById('hudStreakPill');
      if (currentStreak >= 3) {{
        streakPill.style.borderColor = 'rgba(248, 81, 73, 0.5)';
        streakPill.style.background = 'rgba(248, 81, 73, 0.15)';
      }} else {{
        streakPill.style.borderColor = 'var(--border-line)';
        streakPill.style.background = 'var(--bg-card)';
      }}

      // Question Surface
      document.getElementById('qCategoryBadge').innerText = qData.module;
      document.getElementById('qOriginalNum').innerText = `#Đề:${{qData.original_num}}`;
      document.getElementById('qQuestionText').innerText = qData.question;

      // Options
      let optionsList = qData.options.map((opt, idx) => ({{
        text: opt,
        originalIdx: idx,
        isCorrect: idx === qData.ans_idx
      }}));

      if (shuffleAns) {{
        optionsList = shuffleArray(optionsList);
      }}

      const container = document.getElementById('qOptionsContainer');
      container.innerHTML = '';
      const letters = ['A', 'B', 'C', 'D'];

      optionsList.forEach((opt, idx) => {{
        const card = document.createElement('div');
        card.className = `opt-card opt-${{idx}}`;
        card.id = `optCard_${{idx}}`;
        card.innerHTML = `
          <div class="opt-indicator">${{letters[idx]}}</div>
          <div class="opt-content">
            <div class="opt-hint">[Phím ${{idx + 1}}]</div>
            <div class="opt-title">${{opt.text}}</div>
          </div>
        `;
        card.onclick = () => handleAnswerSelect(opt, card, optionsList, qData);
        container.appendChild(card);
      }});

      // Reset Feedback
      const fb = document.getElementById('quizFeedback');
      fb.style.display = 'none';
      fb.className = 'feedback-panel';

      // Start Per-Question Timer if configured
      const timerSec = parseInt(document.getElementById('selTimer').value, 10);
      const timerPill = document.getElementById('hudTimerPill');
      const timerCount = document.getElementById('hudTimerCount');

      if (timerSec > 0) {{
        timerPill.style.display = 'inline-flex';
        secondsLeft = timerSec;
        timerCount.innerText = `${{secondsLeft}}s`;
        timerPill.classList.remove('urgent');

        questionTimer = setInterval(() => {{
          secondsLeft--;
          if (secondsLeft <= 0) {{
            clearInterval(questionTimer);
            timerCount.innerText = 'Hết giờ!';
            handleTimeOut(optionsList, qData);
          }} else {{
            timerCount.innerText = `${{secondsLeft}}s`;
            if (secondsLeft <= 5) {{
              timerPill.classList.add('urgent');
              triggerHaptic(10);
            }}
          }}
        }}, 1000);
      }} else {{
        timerPill.style.display = 'none';
      }}

      window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }}

    function handleAnswerSelect(selectedOpt, selectedCard, allOptions, qData) {{
      if (isAnsweringLocked) return;
      isAnsweringLocked = true;
      clearInterval(questionTimer);

      const isCorrect = selectedOpt.isCorrect;
      const letters = ['A', 'B', 'C', 'D'];

      // Lock other options
      document.querySelectorAll('.opt-card').forEach(c => c.classList.add('disabled'));

      if (isCorrect) {{
        selectedCard.classList.add('is-correct');
        correctCount++;
        currentStreak++;
        if (currentStreak > maxStreak) maxStreak = currentStreak;
        playSound('correct');
        triggerHaptic(25);
        removeWrongQuestionId(qData.id);
      }} else {{
        selectedCard.classList.add('is-wrong');
        wrongCount++;
        currentStreak = 0;
        playSound('wrong');
        triggerHaptic([40, 60, 40]);
        saveWrongQuestionId(qData.id);

        // Highlight correct option
        allOptions.forEach((opt, idx) => {{
          if (opt.isCorrect) {{
            const correctCard = document.getElementById(`optCard_${{idx}}`);
            if (correctCard) correctCard.classList.add('is-correct');
          }}
        }});
      }}

      // Record result
      sessionResults.push({{
        question: qData.question,
        module: qData.module,
        original_num: qData.original_num,
        userAnsText: selectedOpt.text,
        correctAnsText: qData.options[qData.ans_idx],
        isCorrect: isCorrect,
        note: qData.note || ''
      }});

      // Show Feedback Panel
      const fb = document.getElementById('quizFeedback');
      const fbTitle = document.getElementById('fbTitle');
      const fbText = document.getElementById('fbText');
      const fbNote = document.getElementById('fbNoteBox');

      fb.style.display = 'block';
      if (isCorrect) {{
        fb.className = 'feedback-panel is-correct';
        fbTitle.innerHTML = `${{JS_ICONS.check}} Chính xác! Xuất sắc!`;
        fbText.innerHTML = `Đáp án đúng là: <strong>${{qData.options[qData.ans_idx]}}</strong>`;
      }} else {{
        fb.className = 'feedback-panel is-wrong';
        fbTitle.innerHTML = `${{JS_ICONS.cross}} Chưa chính xác!`;
        fbText.innerHTML = `Bạn chọn: <em>${{selectedOpt.text}}</em><br>Đáp án đúng là: <strong>${{qData.options[qData.ans_idx]}}</strong>`;
      }}

      if (qData.note) {{
        fbNote.style.display = 'block';
        fbNote.innerHTML = `💡 <strong>Ghi chú ôn thi:</strong> ${{qData.note}}`;
      }} else {{
        fbNote.style.display = 'none';
      }}

      // Auto next if enabled and correct
      const autoNext = document.getElementById('chkAutoNext').checked;
      if (isCorrect && autoNext) {{
        setTimeout(() => {{
          if (isAnsweringLocked) nextQuestion();
        }}, 850);
      }}
    }}

    function handleTimeOut(allOptions, qData) {{
      if (isAnsweringLocked) return;
      isAnsweringLocked = true;
      document.querySelectorAll('.opt-card').forEach(c => c.classList.add('disabled'));

      wrongCount++;
      currentStreak = 0;
      playSound('wrong');
      triggerHaptic([50, 50, 50]);
      saveWrongQuestionId(qData.id);

      // Highlight correct card
      allOptions.forEach((opt, idx) => {{
        if (opt.isCorrect) {{
          const correctCard = document.getElementById(`optCard_${{idx}}`);
          if (correctCard) correctCard.classList.add('is-correct');
        }}
      }});

      sessionResults.push({{
        question: qData.question,
        module: qData.module,
        original_num: qData.original_num,
        userAnsText: '(Hết thời gian)',
        correctAnsText: qData.options[qData.ans_idx],
        isCorrect: false,
        note: qData.note || ''
      }});

      const fb = document.getElementById('quizFeedback');
      const fbTitle = document.getElementById('fbTitle');
      const fbText = document.getElementById('fbText');
      const fbNote = document.getElementById('fbNoteBox');

      fb.style.display = 'block';
      fb.className = 'feedback-panel is-wrong';
      fbTitle.innerHTML = `${{JS_ICONS.cross}} Hết thời gian làm câu này!`;
      fbText.innerHTML = `Đáp án đúng là: <strong>${{qData.options[qData.ans_idx]}}</strong>`;

      if (qData.note) {{
        fbNote.style.display = 'block';
        fbNote.innerHTML = `💡 <strong>Ghi chú:</strong> ${{qData.note}}`;
      }} else {{
        fbNote.style.display = 'none';
      }}
    }}

    function nextQuestion() {{
      playSound('click');
      currentQIdx++;
      renderCurrentQuestion();
    }}

    function quitQuizConfirm() {{
      if (confirm('Bạn có chắc muốn dừng bài thi hiện tại và quay về màn hình chính?')) {{
        clearInterval(questionTimer);
        returnQuizHome();
      }}
    }}

    function returnQuizHome() {{
      clearInterval(questionTimer);
      updateWrongBadgeUI();
      document.getElementById('quizPlay').style.display = 'none';
      document.getElementById('quizResult').style.display = 'none';
      document.getElementById('quizHome').style.display = 'block';
    }}

    // --- FINISH QUIZ & SUMMARY ---
    function finishQuiz() {{
      clearInterval(questionTimer);
      quizElapsedTime = Math.floor((Date.now() - quizStartTime) / 1000);
      updateWrongBadgeUI();

      document.getElementById('quizPlay').style.display = 'none';
      document.getElementById('quizResult').style.display = 'block';

      const total = quizPool.length;
      const pct = total > 0 ? Math.round((correctCount / total) * 100) : 0;

      // Stats
      document.getElementById('scorePercent').innerText = `${{pct}}%`;
      document.getElementById('resCorrect').innerText = correctCount;
      document.getElementById('resWrong').innerText = wrongCount;
      document.getElementById('resStreak').innerText = maxStreak;

      const mins = Math.floor(quizElapsedTime / 60);
      const secs = quizElapsedTime % 60;
      document.getElementById('resTime').innerText = `${{String(mins).padStart(2, '0')}}:${{String(secs).padStart(2, '0')}}`;

      // SVG Ring animation
      const ring = document.getElementById('scoreFgRing');
      const offset = 380 - (380 * (pct / 100));
      setTimeout(() => {{
        ring.style.strokeDashoffset = offset;
      }}, 50);

      // Grade text
      const grade = document.getElementById('scoreGrade');
      const headline = document.getElementById('scoreHeadline');
      const subline = document.getElementById('scoreSubline');

      if (pct >= 85) {{
        grade.innerText = 'Xuất Sắc!';
        headline.innerText = '🎉 Phong Độ Đỉnh Cao!';
        subline.innerText = 'Kiến thức của bạn rất vững vàng, hoàn toàn tự tin đạt điểm tuyệt đối trong kỳ thi thật!';
        launchConfetti();
        playSound('fanfare');
      }} else if (pct >= 70) {{
        grade.innerText = 'Đạt Chuẩn';
        headline.innerText = '👍 Kết Quả Khá Tốt!';
        subline.innerText = 'Bạn đã nắm chắc hầu hết các chuyên đề chính. Hãy xem lại các câu sai để tối ưu thêm!';
        playSound('fanfare');
      }} else if (pct >= 50) {{
        grade.innerText = 'Cần Cố Gắng';
        headline.innerText = '💪 Đã Vượt Qua Mức Cơ Bản';
        subline.innerText = 'Hãy tận dụng tính năng "Ôn lại câu sai" và mục Sơ đồ tư duy để củng cố thêm các câu còn nhầm lẫn.';
      }} else {{
        grade.innerText = 'Cần Ôn Luyện';
        headline.innerText = '📖 Hãy Ôn Thêm Lý Thuyết!';
        subline.innerText = 'Nên đọc kỹ bảng Sơ đồ tư duy Mindmaps và tra cứu ngân hàng câu hỏi trước khi thi thử lại.';
      }}

      renderReviewList();
    }}

    function restartCurrentQuiz() {{
      startQuiz(false);
    }}

    function startWrongQuizFromSummary() {{
      const wrongInSession = sessionResults.filter(r => !r.isCorrect);
      if (wrongInSession.length === 0) {{
        alert('Chúc mừng! Bạn không làm sai câu nào trong bài vừa rồi!');
        return;
      }}
      startQuiz(true);
    }}

    function filterReview(type, btn) {{
      currentReviewFilter = type;
      document.querySelectorAll('.review-filter-btn').forEach(b => b.classList.remove('active'));
      if (btn) btn.classList.add('active');
      renderReviewList();
    }}

    function renderReviewList() {{
      const container = document.getElementById('reviewContainer');
      container.innerHTML = '';

      const countAll = sessionResults.length;
      const countWrong = sessionResults.filter(r => !r.isCorrect).length;
      const countCorrect = sessionResults.filter(r => r.isCorrect).length;

      document.getElementById('revCountAll').innerText = countAll;
      document.getElementById('revCountWrong').innerText = countWrong;
      document.getElementById('revCountCorrect').innerText = countCorrect;

      const filtered = sessionResults.filter(r => {{
        if (currentReviewFilter === 'WRONG') return !r.isCorrect;
        if (currentReviewFilter === 'CORRECT') return r.isCorrect;
        return true;
      }});

      if (filtered.length === 0) {{
        container.innerHTML = '<div style="color: var(--text-dim); padding: 16px; text-align: center;">Không có câu hỏi nào trong mục này.</div>';
        return;
      }}

      filtered.forEach((r, idx) => {{
        const card = document.createElement('div');
        card.className = `review-card ${{r.isCorrect ? 'correct' : 'wrong'}}`;
        card.innerHTML = `
          <div style="display: flex; justify-content: space-between; font-size: 11px; color: var(--text-dim); margin-bottom: 4px;">
            <span>${{r.module}} (Đề #${{r.original_num}})</span>
            <span style="font-weight: 700; color: ${{r.isCorrect ? 'var(--color-success)' : 'var(--color-danger)'}};">
              ${{r.isCorrect ? '✓ ĐÚNG' : '✗ SAI'}}
            </span>
          </div>
          <div style="font-size: 14px; font-weight: 700; color: var(--text-heading); margin-bottom: 6px;">
            ${{idx + 1}}. ${{r.question}}
          </div>
          <div style="margin-bottom: 2px;">
            <span style="color: var(--text-muted);">Bạn chọn:</span>
            <strong style="color: ${{r.isCorrect ? 'var(--color-success)' : 'var(--color-danger)'}};">${{r.userAnsText}}</strong>
          </div>
          ${{!r.isCorrect ? `<div><span style="color: var(--text-muted);">Đáp án đúng:</span> <strong style="color: var(--color-success);">${{r.correctAnsText}}</strong></div>` : ''}}
          ${{r.note ? `<div class="q-note-hint">💡 ${{r.note}}</div>` : ''}}
        `;
        container.appendChild(card);
      }});
    }}

    // --- CONFETTI CELEBRATION EFFECT ---
    function launchConfetti() {{
      const canvas = document.getElementById('confettiCanvas');
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      canvas.width = window.innerWidth;
      canvas.height = window.innerHeight;

      const particles = [];
      const colors = ['#f85149', '#388bfd', '#2ea043', '#d29922', '#a371f7', '#06b6d4'];

      for (let i = 0; i < 90; i++) {{
        particles.push({{
          x: canvas.width / 2,
          y: canvas.height / 2,
          vx: (Math.random() - 0.5) * 18,
          vy: (Math.random() - 0.7) * 18,
          size: Math.random() * 8 + 4,
          color: colors[Math.floor(Math.random() * colors.length)],
          rot: Math.random() * 360,
          vRot: (Math.random() - 0.5) * 10,
          alpha: 1
        }});
      }}

      let animId = null;
      let frame = 0;
      function step() {{
        frame++;
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        particles.forEach(p => {{
          p.x += p.vx;
          p.y += p.vy;
          p.vy += 0.35; // gravity
          p.rot += p.vRot;
          p.alpha -= 0.012;

          ctx.save();
          ctx.globalAlpha = Math.max(0, p.alpha);
          ctx.translate(p.x, p.y);
          ctx.rotate((p.rot * Math.PI) / 180);
          ctx.fillStyle = p.color;
          ctx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size);
          ctx.restore();
        }});

        if (frame < 90) {{
          animId = requestAnimationFrame(step);
        }} else {{
          ctx.clearRect(0, 0, canvas.width, canvas.height);
        }}
      }}
      step();
    }}

    // --- KEYBOARD SHORTCUTS ---
    window.addEventListener('keydown', (e) => {{
      const playSection = document.getElementById('quizPlay');
      if (!playSection || playSection.style.display === 'none') return;

      const key = e.key.toUpperCase();

      if (['1', '2', '3', '4', 'A', 'B', 'C', 'D'].includes(key) && !isAnsweringLocked) {{
        let idx = -1;
        if (key === '1' || key === 'A') idx = 0;
        if (key === '2' || key === 'B') idx = 1;
        if (key === '3' || key === 'C') idx = 2;
        if (key === '4' || key === 'D') idx = 3;

        const card = document.getElementById(`optCard_${{idx}}`);
        if (card) card.click();
      }} else if ((e.key === 'Enter' || e.key === ' ') && isAnsweringLocked) {{
        const btnNext = document.getElementById('btnNextQuestion');
        if (btnNext && btnNext.offsetParent !== null) {{
          e.preventDefault();
          nextQuestion();
        }}
      }}
    }});

    // --- MINDMAPS RENDERING ---
    let currentMindmapIdx = 0;

    function renderMindmaps() {{
      const carousel = document.getElementById('mindmapModuleTabs');
      carousel.innerHTML = '';

      mindmapsData.forEach((mm, idx) => {{
        const btn = document.createElement('button');
        btn.className = `mm-tab-item ${{idx === currentMindmapIdx ? 'active' : ''}}`;
        btn.innerText = mm.shortTitle;
        btn.onclick = () => selectMindmapModule(idx);
        carousel.appendChild(btn);
      }});

      renderActiveMindmapTree();
    }}

    function selectMindmapModule(idx) {{
      currentMindmapIdx = idx;
      playSound('click');
      document.querySelectorAll('#mindmapModuleTabs .mm-tab-item').forEach((b, i) => {{
        if (i === idx) b.classList.add('active');
        else b.classList.remove('active');
      }});
      renderActiveMindmapTree();
    }}

    function renderActiveMindmapTree() {{
      const tree = document.getElementById('mindmapTreeContainer');
      tree.innerHTML = '';

      if (!mindmapsData[currentMindmapIdx]) return;
      const mm = mindmapsData[currentMindmapIdx];

      mm.nodes.forEach(node => {{
        const branch = document.createElement('div');
        branch.className = 'branch-card';
        branch.innerHTML = `
          <div class="branch-head" onclick="toggleBranch(this.parentElement)">
            <div class="branch-title">${{node.title}}</div>
            <span class="branch-toggle">${{JS_ICONS.chevronDown}}</span>
          </div>
          <div class="branch-body">
            <div class="branch-inner">
              ${{node.items.map(item => `
                <div class="leaf-item">
                  <div class="leaf-label">${{item.label}}</div>
                  <div class="leaf-desc">${{item.detail}}</div>
                </div>
              `).join('')}}
            </div>
          </div>
        `;
        tree.appendChild(branch);
      }});
    }}

    function toggleBranch(branch) {{
      branch.classList.toggle('collapsed');
      playSound('click');
      triggerHaptic(15);
    }}

    function expandAllMindmap(expand) {{
      playSound('click');
      triggerHaptic(15);
      document.querySelectorAll('.branch-card').forEach(b => {{
        if (expand) b.classList.remove('collapsed');
        else b.classList.add('collapsed');
      }});
    }}

    function handleMindmapSearch(keyword) {{
      const term = keyword.trim().toLowerCase();
      const leaves = document.querySelectorAll('.leaf-item');

      leaves.forEach(leaf => {{
        const text = leaf.innerText.toLowerCase();
        if (!term || text.includes(term)) {{
          leaf.style.display = 'block';
          if (term) {{
            leaf.closest('.branch-card').classList.remove('collapsed');
          }}
        }} else {{
          leaf.style.display = 'none';
        }}
      }});
    }}

    // --- QUESTION DIRECTORY / CHEATSHEET ---
    let currentDirModuleFilter = 'ALL';
    let currentDirSearchTerm = '';

    function renderDirModuleTabs() {{
      const container = document.getElementById('dirModuleTabs');
      container.innerHTML = '';

      const modules = ['ALL', ...new Set(rawQuestions.map(q => q.module))];

      modules.forEach(mod => {{
        const btn = document.createElement('button');
        btn.className = `mm-tab-item ${{mod === currentDirModuleFilter ? 'active' : ''}}`;
        btn.innerText = mod === 'ALL' ? `Tất cả (${{rawQuestions.length}})` : mod;
        btn.onclick = () => filterDirModule(mod, btn);
        container.appendChild(btn);
      }});
    }}

    function filterDirModule(mod, btn) {{
      currentDirModuleFilter = mod;
      playSound('click');
      document.querySelectorAll('#dirModuleTabs .mm-tab-item').forEach(b => b.classList.remove('active'));
      if (btn) btn.classList.add('active');
      renderQuestionsDirectoryList();
    }}

    function filterQuestionsDirectory(val) {{
      currentDirSearchTerm = val.trim().toLowerCase();
      renderQuestionsDirectoryList();
    }}

    function renderQuestionsDirectory() {{
      currentDirModuleFilter = 'ALL';
      renderDirModuleTabs();
      renderQuestionsDirectoryList();
    }}

    function renderQuestionsDirectoryList() {{
      const container = document.getElementById('questionsDirectoryList');
      container.innerHTML = '';

      const filtered = rawQuestions.filter(q => {{
        const matchMod = currentDirModuleFilter === 'ALL' || q.module === currentDirModuleFilter;
        const text = (q.question + ' ' + q.options.join(' ') + ' ' + (q.note || '')).toLowerCase();
        const matchSearch = !currentDirSearchTerm || text.includes(currentDirSearchTerm);
        return matchMod && matchSearch;
      }});

      if (filtered.length === 0) {{
        container.innerHTML = '<div style="color: var(--text-dim); padding: 18px; text-align: center;">Không tìm thấy câu hỏi phù hợp với từ khóa tra cứu.</div>';
        return;
      }}

      const letters = ['A', 'B', 'C', 'D'];
      filtered.forEach((q, idx) => {{
        const item = document.createElement('div');
        item.className = 'q-row-card';
        item.innerHTML = `
          <div style="display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 11px; color: var(--text-dim);">
            <span class="q-row-tag">${{q.module}}</span>
            <span>Câu gốc #${{q.original_num}}</span>
          </div>
          <div style="font-size: 14px; font-weight: 700; color: var(--text-heading); margin-bottom: 8px;">
            <strong>${{idx + 1}}.</strong> ${{q.question}}
          </div>
          <div class="q-opts-grid">
            ${{q.options.map((opt, oIdx) => `
              <div class="q-opt-item ${{oIdx === q.ans_idx ? 'is-correct' : ''}}">
                <strong>${{letters[oIdx]}}.</strong> ${{opt}}
              </div>
            `).join('')}}
          </div>
          ${{q.note ? `<div class="q-note-hint">💡 ${{q.note}}</div>` : ''}}
        `;
        container.appendChild(item);
      }});
    }}

    // --- INITIALIZATION ---
    window.addEventListener('DOMContentLoaded', () => {{
      initTheme();
      populateModuleDropdown();
      updateWrongBadgeUI();
      renderMindmaps();
      renderQuestionsDirectory();
    }});
  </script>
</body>
</html>
"""

# Write to index.html and quizizz-app.html
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

with open('quizizz-app.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Successfully generated index.html ({len(html_content)} chars) and quizizz-app.html!")
