# -*- coding: utf-8 -*-
"""
Tách từng biểu mẫu từ "Phu luc 68_2025.docx" thành file .docx riêng,
giữ NGUYÊN định dạng gốc, và chèn placeholder docxtemplater ({ten})
vào đúng các dòng "………" tương ứng với field thu thập trên web.

Chạy: python3 scripts/build_templates.py
Kết quả: templates/<id>.docx  (vd templates/pl1-1.docx)
"""
import copy
import re
import shutil
import os
import docx
from docx.oxml.ns import qn

SRC = "Thư Viện Pháp Luật/Phu luc 68_2025.docx"
OUT = "templates"

# ---------------------------------------------------------------------------
# 1. Xác định vị trí block (paragraph/table) bắt đầu mỗi mẫu
# ---------------------------------------------------------------------------
def block_text(el):
    if el.tag == qn('w:tbl'):
        return '[TBL]'
    return ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()


def scan_marks(path):
    d = docx.Document(path)
    els = [c for c in d.element.body.iterchildren()
           if c.tag in (qn('w:p'), qn('w:tbl'))]
    marks, plII = [], None
    for i, el in enumerate(els):
        t = block_text(el)
        if re.fullmatch(r'Mẫu số (\d+)', t):
            marks.append((int(t.split()[-1]), i))
        if t == 'PHỤ LỤC II':
            plII = i
    plI = [m for m in marks if m[1] < plII]
    pl2 = [m for m in marks if m[1] > plII]
    return els, plI, pl2, plII


# ---------------------------------------------------------------------------
# 2. Ánh xạ form web -> mẫu trong Phụ lục
#    ('I', n)  = Phụ lục I, Mẫu số n ; ('II', n) = Phụ lục II, Mẫu số n
# ---------------------------------------------------------------------------
FORM_SOURCE = {
    'pl1-1':  ('I', 1),
    'pl1-2':  ('I', 2),
    'pl1-3':  ('I', 3),
    'pl1-4':  ('I', 4),
    'pl1-5':  ('I', 5),
    'pl1-6':  ('I', 6),
    'pl1-7':  ('I', 7),
    'pl1-13': ('I', 13),
    'pl1-27': ('I', 27),
    'pl2-1':  ('II', 1),
}

# ---------------------------------------------------------------------------
# 3. Bản đồ chèn placeholder cho từng mẫu.
#    Mỗi entry: (regex áp lên text cả đoạn, chuỗi thay thế chứa {tag}).
#    Nếu một nhãn xuất hiện nhiều lần (vd địa chỉ cá nhân & trụ sở), liệt kê
#    nhiều entry cùng pattern -> tiêu thụ lần lượt theo thứ tự xuất hiện.
#    [BLANK] đại diện cho cụm "………"/"....." trong regex.
# ---------------------------------------------------------------------------
# Cụm "chỗ trống" trong mẫu: dấu chấm ASCII, ellipsis "…", "/", "_", tab, space
B = r'[.…/_\t ]*'          # 0+ ký tự trống
BD = r'[.…/_\t ]+'         # 1+ ký tự trống (bắt buộc có chỗ điền)

# Dòng ngày tháng ở đầu mẫu: "……, ngày ……tháng …… năm ……"
HDR_DATE = (r'^[.…]+,\s*ngày.*?năm[.… ]*$', r'{ngay_thang}')

# Các trường chung của Giấy đề nghị ĐKDN (DNTN/cty) – địa chỉ trụ sở dùng tiền tố ts_
def dn_company(name_tag='ten_cty', addr='', von_so=True, von_chu=True):
    p = addr + '_' if addr else ''
    rules = [
        HDR_DATE,
        (r'^(Kính gửi:.*?\))' + B + r'$', r'\1 {co_quan}'),
        (r'^(Tên (?:doanh nghiệp|công ty) viết bằng tiếng Việt.*?:)' + B + r'$', r'\1 {%s}' % name_tag),
        (r'^(Tên (?:doanh nghiệp|công ty) viết bằng tiếng nước ngoài.*?:)' + B + r'$', r'\1 {ten_dn_nn}'),
        (r'^(Tên (?:doanh nghiệp|công ty) viết tắt.*?:)' + B + r'$', r'\1 {ten_vt}'),
        (r'^(Số nhà/phòng.*?thôn:)' + B + r'$', r'\1 {%sso_nha}' % p),
        (r'^(Xã/Phường/Đặc khu:)' + B + r'$', r'\1 {%sxa}' % p),
        (r'^(Tỉnh/Thành phố trực thuộc trung ương:)' + B + r'$', r'\1 {%stinh}' % p),
    ]
    if von_so:
        rules.append((r'^(Vốn điều lệ \(bằng số.*?:)' + B + r'$', r'\1 {von_so}'))
    if von_chu:
        rules.append((r'^(Vốn điều lệ \(bằng chữ.*?:)' + B + r'$', r'\1 {von_chu}'))
    return rules


MAPS = {
    'pl1-1': [
        HDR_DATE,
        (r'^(Kính gửi:.*?\))' + B + r'$', r'\1 {co_quan}'),
        (r'^(Tôi là.*?in hoa\):)' + B + r'$', r'\1 {ho_ten}'),
        (r'^(Ngày, tháng, năm sinh:)' + B + r'$', r'\1 {ngay_sinh}'),
        (r'^(Giới tính:)' + B + r'$', r'\1 {gioi_tinh}'),
        (r'^(Số định danh cá nhân:)' + B + r'$', r'\1 {so_cccd}'),
        # địa chỉ cá nhân (xuất hiện trước) rồi địa chỉ trụ sở (lần 2)
        (r'^(Số nhà/phòng.*?thôn:)' + B + r'$', r'\1 {so_nha}'),
        (r'^(Số nhà/phòng.*?thôn:)' + B + r'$', r'\1 {ts_so_nha}'),
        (r'^(Xã/Phường/Đặc khu:)' + B + r'$', r'\1 {xa_phuong}'),
        (r'^(Xã/Phường/Đặc khu:)' + B + r'$', r'\1 {ts_xa}'),
        (r'^(Tỉnh/Thành phố trực thuộc trung ương:)' + B + r'$', r'\1 {tinh_tp}'),
        (r'^(Tỉnh/Thành phố trực thuộc trung ương:)' + B + r'$', r'\1 {ts_tinh}'),
        (r'^(Điện thoại \(nếu có\):)' + BD + r'(Thư điện tử \(nếu có\):)' + B + r'$',
         r'\1 {dien_thoai} \2 {email}'),
        (r'^(Tên doanh nghiệp viết bằng tiếng Việt.*?:)' + B + r'$', r'\1 {ten_dn_viet}'),
        (r'^(Tên doanh nghiệp viết bằng tiếng nước ngoài.*?:)' + B + r'$', r'\1 {ten_dn_nn}'),
        (r'^(Tên doanh nghiệp viết tắt.*?:)' + B + r'$', r'\1 {ten_dn_vt}'),
        (r'^(Điện thoại:)' + BD + r'(Số fax.*?:)' + B + r'$', r'\1 {ts_dien_thoai} \2'),
        (r'^(Thư điện tử \(nếu có\):)' + BD + r'(Website.*?:)' + B + r'$', r'\1 {ts_email} \2'),
        (r'^(Vốn đầu tư \(bằng số.*?:)' + B + r'$', r'\1 {von_so}'),
        (r'^(Vốn đầu tư \(bằng chữ.*?:)' + B + r'$', r'\1 {von_chu}'),
    ],
    'pl1-2': dn_company('ten_cty', addr='dc') + [
        (r'^(Tôi là.*?in hoa\):)' + B + r'$', r'\1 {ho_ten_chu}'),
        (r'^(Ngày, tháng, năm sinh:)' + B + r'$', r'\1 {ngay_sinh_chu}'),
        (r'^(Giới tính:)' + B + r'$', r'\1 {gt_chu}'),
        (r'^(Số định danh cá nhân:)' + B + r'$', r'\1 {cccd_chu}'),
    ],
    'pl1-3': dn_company('ten_cty', addr=''),
    'pl1-4': dn_company('ten_cty', addr='', von_chu=False),
    'pl1-5': dn_company('ten_cty', addr=''),
    'pl1-6': [],   # bảng danh sách thành viên – xử lý ở v2 (vòng lặp)
    'pl1-7': [],   # bảng danh sách cổ đông – xử lý ở v2
    'pl1-13': [
        HDR_DATE,
        (r'^(Kính gửi:.*?\))' + B + r'$', r'\1 {co_quan}'),
        (r'^(Tên doanh nghiệp.*?in hoa\):)' + B + r'$', r'\1 {ten_dn}'),
        (r'^(Mã số doanh nghiệp.*?:)' + B + r'$', r'\1 {ma_so_dn}'),
    ],
    'pl1-27': [
        HDR_DATE,
        (r'^(Kính gửi:.*?\))' + B + r'$', r'\1 {co_quan}'),
        (r'^(Tên doanh nghiệp.*?in hoa\):)' + B + r'$', r'\1 {ten_dn}'),
        (r'^(Mã số doanh nghiệp/Mã số thuế:)' + B + r'$', r'\1 {ma_so_dn}'),
        (r'^(Lý do.*?:)' + B + r'$', r'\1 {ly_do}'),
    ],
    'pl2-1': [
        HDR_DATE,
        (r'^(Kính gửi:.*?\))' + B + r'$', r'\1 {co_quan}'),
        (r'^(Tôi là.*?in hoa\):)' + B + r'$', r'\1 {chu_ten}'),
        (r'^(Sinh ngày:)' + B + r'$', r'\1 {chu_ns}'),
        (r'^(Giới tính:)' + B + r'$', r'\1 {chu_gt}'),
        (r'^(Số định danh cá nhân:?)' + B + r'$', r'\1 {chu_cccd}'),
        (r'^(Điện thoại \(nếu có\):)' + BD + r'(Thư điện tử \(nếu có\):)' + B + r'$',
         r'\1 {chu_tel} \2 {chu_email}'),
        (r'^(Tên hộ kinh doanh viết bằng tiếng Việt.*?:)' + B + r'$', r'\1 {hkd_ten}'),
        (r'^(Số nhà/phòng.*?thôn:)' + B + r'$', r'\1 {hkd_dc}'),
        (r'^(Xã/Phường/Đặc khu:)' + B + r'$', r'\1 {hkd_xa}'),
        (r'^(Tỉnh/Thành phố trực thuộc trung ương:)' + B + r'$', r'\1 {hkd_tinh}'),
        (r'^(Tổng số \(bằng số.*?:)' + B + r'$', r'\1 {von_so}'),
    ],
}


# ---------------------------------------------------------------------------
# Vòng lặp bảng (docxtemplater): {#arr} ở ô đầu, {/arr} ở ô cuối của hàng dữ liệu
#   cols: {chỉ_số_cột: 'placeholder'} ; row: chỉ số hàng dữ liệu cần biến thành lặp
# ---------------------------------------------------------------------------
NGANH_LOOP = {'kind': 'nganh', 'arr': 'nganh', 'row': 1,
              'cols': {0: '{stt}', 1: '{ten}', 2: '{ma}', 3: '{chinh}'}}

TABLE_LOOPS = {
    'pl1-1': [NGANH_LOOP], 'pl1-2': [NGANH_LOOP], 'pl1-3': [NGANH_LOOP],
    'pl1-4': [NGANH_LOOP], 'pl1-5': [NGANH_LOOP], 'pl2-1': [NGANH_LOOP],
    'pl1-6': [{'kind': 'tbl0', 'arr': 'thanhvien', 'row': 3,
               'cols': {0: '{stt}', 1: '{ten}', 4: '{cccd}', 8: '{von}', 9: '{tyle}'}}],
    'pl1-7': [{'kind': 'tbl0', 'arr': 'codong', 'row': 5,
               'cols': {0: '{stt}', 1: '{ten}', 4: '{cccd}', 8: '{cp}', 10: '{tyle}'}}],
}


def set_cell_text(cell, text):
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    p.add_run(text)


def find_nganh_table(doc):
    for tbl in doc.tables:
        cells = tbl.rows[0].cells
        if len(cells) == 4 and 'Tên ngành' in cells[1].text:
            return tbl
    return None


def add_table_loops(doc, fid):
    n = 0
    for cfg in TABLE_LOOPS.get(fid, []):
        tbl = find_nganh_table(doc) if cfg['kind'] == 'nganh' else doc.tables[0]
        if tbl is None or cfg['row'] >= len(tbl.rows):
            continue
        cells = tbl.rows[cfg['row']].cells
        for ci, ph in cfg['cols'].items():
            set_cell_text(cells[ci], ph)
        # bọc vòng lặp: mở ở ô đầu, đóng ở ô cuối (ô cuối thật, tránh trùng merge)
        first, last = cells[0], cells[-1]
        set_cell_text(first, '{#%s}%s' % (cfg['arr'], first.text))
        set_cell_text(last, '%s{/%s}' % (last.text, cfg['arr']))
        n += 1
    return n


def set_paragraph_text(p, new_text):
    """Ghi đè text của paragraph thành 1 run, giữ định dạng run đầu."""
    runs = p.runs
    if runs:
        runs[0].text = new_text
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        p.add_run(new_text)


def inject_placeholders(doc, mapping):
    """Áp các regex lên từng paragraph theo thứ tự, tiêu thụ entry trùng nhãn."""
    rules = [[pat, repl, 0] for pat, repl in mapping]  # [pat, repl, used]
    hits = {}
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        for rule in rules:
            pat, repl, used = rule
            if used:
                # với nhãn lặp: chỉ bỏ qua nếu entry này (cùng pat) đã dùng
                continue
            if re.match(pat, text):
                new = re.sub(pat, repl, text)
                set_paragraph_text(p, new)
                rule[2] = 1
                tag = re.search(r'\{(\w+)\}', repl)
                hits[tag.group(1) if tag else repl] = True
                break
    return hits


def extract_form(form_id, src_path, scope, num, plI, pl2, plII_idx, total_blocks):
    # tìm block bắt đầu & kết thúc
    marks = plI if scope == 'I' else pl2
    starts = dict(marks)
    start = starts[num]
    # end = mark kế tiếp trong cùng phụ lục, hoặc ranh giới Phụ lục II / hết tài liệu
    later = sorted(i for _, i in marks if i > start)
    if later:
        end = later[0]
    else:
        end = plII_idx if scope == 'I' else total_blocks
    # clone file gốc rồi cắt
    tmp = os.path.join(OUT, form_id + '.docx')
    shutil.copyfile(src_path, tmp)
    d = docx.Document(tmp)
    blocks = [c for c in d.element.body.iterchildren()
              if c.tag in (qn('w:p'), qn('w:tbl'))]
    for i, el in enumerate(blocks):
        if i < start or i >= end:
            el.getparent().remove(el)
    hits = inject_placeholders(d, MAPS.get(form_id, []))
    nloops = add_table_loops(d, form_id)
    d.save(tmp)
    return tmp, len(MAPS.get(form_id, [])), len(hits), nloops


def main():
    os.makedirs(OUT, exist_ok=True)
    els, plI, pl2, plII = scan_marks(SRC)
    total = len(els)
    print(f"Nguồn: {total} blocks | Phụ lục I: {len(plI)} mẫu | Phụ lục II: {len(pl2)} mẫu\n")
    for fid, (scope, num) in FORM_SOURCE.items():
        path, nrules, nhits, nloops = extract_form(fid, SRC, scope, num, plI, pl2, plII, total)
        size = os.path.getsize(path)
        print(f"  {fid:8s} <- Phụ lục {scope} Mẫu {num:<2d}  "
              f"-> {path} ({size:,}B)  field: {nhits}/{nrules}  bảng-lặp: {nloops}")
    print("\nXong.")


if __name__ == '__main__':
    main()
