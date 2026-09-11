# -*- coding: utf-8 -*-
"""
Công cụ thay thế dữ liệu giữa hai file Word (.doc/.docx)

- Đọc dữ liệu từ Tài liệu A: Tên / Số / Số tem kiểm định.
  Tài liệu A có thể chứa NHIỀU bộ dữ liệu (nhiều giấy chứng nhận /
  nhiều biên bản) — mỗi bộ ứng với một trang của Tài liệu B.
- Thay vào các trường tương ứng trong Tài liệu B theo từng trang:
  bộ thứ nhất -> trang 1, bộ thứ hai -> trang 2, ...
- Hỗ trợ nhiều dạng tài liệu: giấy chứng nhận (nội dung trong textbox),
  biên bản (đoạn văn + bảng), nhãn "Tên đối tượng/Tên phương tiện/..."
- Thay thế trực tiếp trong XML của file Word: chỉ sửa đúng đoạn text
  cần thay, mọi phần còn lại giữ nguyên từng byte -> định dạng, font,
  bảng biểu, header/footer... không thể bị ảnh hưởng.

Chạy giao diện :  python ThayTheDuLieu.py
Dòng lệnh      :  python ThayTheDuLieu.py <fileA> <fileB> <fileXuat>
"""

import base64
import io
import os
import re
import sys
import time
import zipfile
import tempfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# ---------------------------------------------------------------------------
# Các trường cần thay thế — dùng chung cho cả đọc (A) và thay (B).
# Khi gặp dạng tài liệu mới chỉ cần bổ sung nhãn vào pattern tương ứng.
# ---------------------------------------------------------------------------
# Giá trị = phần còn lại sau nhãn, dừng trước tab / 2 khoảng trắng liền /
# nhãn kế tiếp trên cùng dòng / hết đoạn; bỏ dấu chấm cuối.
# Giá trị được phép chứa dấu ":" (vd "Dòng điện sơ sấp: 35:√3 kV" của máy
# biến áp) nên phải nhận diện nhãn kế tiếp bằng tên nhãn + dấu ":".
_NEXT_LABEL = (r"(?:Dòng điện\s*(?:sơ|thứ)\s*[scx]?ấp|Dung lượng|"
               r"Cấp chính xác|Tần số(?:\s*làm việc)?|Mức cách điện|"
               r"Điện áp(?:\s*làm việc)?|Năm(?:\s*sản\s*xuất)?|Kiểu|"
               r"(?:Nơi|Cơ sở)\s*sản xuất|Số)\s*:")
_VAL = (r"(?P<val>\S(?:[^\t]*?[^\s.])?)\s*\.?\s*"
        r"(?=\t|\s\s|\s*$|\s+" + _NEXT_LABEL + r")")

def _adj_so_gcn(new, old):
    """Khách chỉ nhập số (vd 4111) -> giữ nguyên phần đuôi của mẫu
    ("4302 - 26 (YB)" -> "4111 - 26 (YB)")."""
    if new.isdigit():
        return new + re.sub(r"^\s*[0-9]+", "", old)
    return new


FIELDS = [
    {
        # Số hiệu giấy chứng nhận "Số (N°): 3681 - 26 (YB)" — KHÔNG lấy từ
        # tài liệu nguồn mà do chương trình tự đánh số tăng dần theo trang.
        "key": "so_gcn",
        "display": "Số (N°)",
        "pattern": re.compile(r"Số\s*\(N°\)\s*:\s*(?P<val>[0-9]+[^\n]*?)\s*$"),
        "auto_only": True,
        "adjust": _adj_so_gcn,
    },
    {
        "key": "ten",
        "display": "Tên",
        "pattern": re.compile(
            r"Tên (?:đối tượng|phương tiện|thiết bị)(?:\s*đo)?\s*:\s*(?P<val>.+?)\s*\.?\s*$"),
    },
    {
        "key": "so",
        "display": "Số",
        # "Số:" + mã máy >= 3 ký tự chữ-số đứng trọn vẹn, phải có ít nhất một
        # chữ số (nhận cả "20220300150F"); không nhận "Số: 01/TI-TH",
        # "Số: 2266/BB-TI(YB)" vì có ký tự bám ngay sau; "Số (N°):" không khớp
        "pattern": re.compile(
            r"(?<![\w])Số\s*:\s*(?P<val>(?=[0-9A-Za-z]*[0-9])[0-9A-Za-z]{3,})"
            r"\s*\.?\s*(?=\t|\s\s|$)"),
    },
    {
        "key": "nam",
        "display": "Năm sản xuất",
        "pattern": re.compile(
            r"(?<![\w])Năm(?:\s*sản\s*xuất)?\s*:\s*(?P<val>(?:19|20)[0-9]{2})"),
    },
    {
        "key": "kieu",
        "display": "Kiểu",
        "pattern": re.compile(r"(?<![\w])Kiểu\s*:\s*" + _VAL),
    },
    {
        "key": "noi_sx",
        "display": "Nơi sản xuất",
        # A ghi "Cơ sở sản xuất", B ghi "Nơi sản xuất" (không nhầm "Cơ sở sử dụng")
        "pattern": re.compile(r"(?:Cơ sở|Nơi) sản xuất\s*:\s*" + _VAL),
    },
    {
        "key": "so_cap",
        "display": "Dòng điện sơ cấp",
        # chấp nhận cả lỗi chính tả "sơ sấp" hay gặp trong biểu mẫu
        "pattern": re.compile(r"Dòng điện sơ [scx]ấp\s*:\s*" + _VAL),
    },
    {
        "key": "dung_luong",
        "display": "Dung lượng",
        "pattern": re.compile(r"Dung lượng\s*:\s*" + _VAL),
    },
    {
        "key": "thu_cap",
        "display": "Dòng điện thứ cấp",
        "pattern": re.compile(r"Dòng điện thứ cấp\s*:\s*" + _VAL),
    },
    {
        "key": "cap_cx",
        "display": "Cấp chính xác",
        "pattern": re.compile(r"Cấp chính xác\s*:\s*" + _VAL),
    },
    {
        "key": "tan_so",
        "display": "Tần số làm việc",
        "pattern": re.compile(r"Tần số làm việc\s*:\s*" + _VAL),
    },
    {
        "key": "cach_dien",
        "display": "Mức cách điện",
        "pattern": re.compile(r"Mức cách điện\s*:\s*" + _VAL),
    },
    {
        "key": "hang_so",
        "display": "Hằng số công tơ",
        # số đứng trước "xung/kW-h" trong dòng đặc trưng kỹ thuật đo lường
        "pattern": re.compile(
            r"(?P<val>[0-9]+(?:[.,][0-9]+)?)\s*xung\s*/\s*kW[.\-]?h"),
    },
    {
        "key": "pham_vi",
        "display": "Phạm vi đo",
        "pattern": re.compile(r"Phạm vi đo\s*:\s*" + _VAL),
    },
    {
        "key": "max_can",
        "display": "Max (mức cân lớn nhất)",
        "pattern": re.compile(r"(?<![\w])Max\s*=\s*" + _VAL),
    },
    {
        "key": "min_can",
        "display": "Min (mức cân nhỏ nhất)",
        "pattern": re.compile(r"(?<![\w])Min\s*=\s*" + _VAL),
    },
    {
        "key": "d_can",
        "display": "d (giá trị độ chia)",
        # "d = 50 g." hoặc "Giá trị độ chia: d = 0,001 g."
        "pattern": re.compile(r"(?<![\wđ])d\s*=\s*" + _VAL),
    },
    {
        "key": "e_can",
        "display": "e (độ chia kiểm)",
        "pattern": re.compile(r"(?<![\wđ])e\s*=\s*" + _VAL),
    },
    {
        "key": "che_do",
        "display": "Chế độ kiểm định",
        "pattern": re.compile(r"Chế độ kiểm định\s*:\s*(?P<val>[^¨þ£\t]+?)\s*\.?\s*$"),
    },
    {
        "key": "noi_sd",
        "display": "Nơi sử dụng",
        "pattern": re.compile(r"Nơi sử dụng\s*:\s*(?P<val>.+?)\s*\.?\s*$"),
    },
    {
        "key": "nguoi_sd",
        "display": "Người/Đơn vị sử dụng",
        "pattern": re.compile(r"Người\s*/\s*Đơn vị sử dụng\s*:\s*(?P<val>.+?)\s*\.?\s*$"),
    },
    {
        "key": "phuong_phap",
        "display": "Phương pháp thực hiện",
        # giá trị có thể chứa dấu ":" (vd "ĐLVN 39:2019")
        "pattern": re.compile(r"Phương pháp thực hiện\s*:\s*(?P<val>.+?)\s*\.?\s*$"),
    },
    {
        "key": "thoi_han",
        "display": "Thời hạn đến",
        "pattern": re.compile(r"Thời hạn đến\s*\(\*\)\s*(?P<val>.+?)\s*\.?\s*$"),
    },
    {
        "key": "ngay_cap",
        "display": "Ngày cấp (dòng ký)",
        # "Lào Cai, ngày 22 tháng 07 năm 2026" — thay bằng đúng ngày kiểm định
        "pattern": re.compile(
            r"[^,\n]{2,40},\s*(?P<val>ngày\s+\d{1,2}\s+tháng\s+\d{1,2}"
            r"\s+năm\s+\d{4})"),
    },
    {
        "key": "tem",
        "display": "Tem kiểm định",
        "pattern": re.compile(
            r"(?:Số\s*)?[Tt]em kiểm định\s*:\s*(?P<val>[0-9A-Za-z]+)"),
    },
]

FIELD_ORDER = {f["key"]: i for i, f in enumerate(FIELDS)}


# ---------------------------------------------------------------------------
# Danh sách word mẫu có sẵn (thư mục MauWord)
#   ready=True  : đã hỗ trợ, chọn là dùng được ngay
#   ready=False : chưa hỗ trợ, giao diện báo "sắp cập nhật"
# Bổ sung mẫu mới: thêm file vào thư mục MauWord rồi khai báo một dòng ở đây.
# ---------------------------------------------------------------------------
# Phiên bản chương trình — hiện trên thanh tiêu đề để biết đang dùng bản nào.
# Mỗi lần bàn giao bản mới nhớ tăng số này.
APP_VERSION = "21"
APP_DATE = "09/2026"
APP_TITLE = "Công cụ tạo giấy chứng nhận kiểm định  —  bản %s (%s)" % (
    APP_VERSION, APP_DATE)

TEMPLATE_DIR_NAME = "MauWord"
COMING_SOON = "   — sắp cập nhật"
# gợi ý mờ trong ô "Số bắt đầu (N°)"
SO_GCN_HINT = "4302 - 26 (YB)"
# dòng cuối danh sách mẫu: tự chọn file Word riêng
CHOOSE_OTHER = "📂  Chọn file Word khác..."

TEMPLATES = [
    {"file": "MBD trung thế.docx", "label": "Máy biến dòng đo lường trung thế",
     "ready": True},
    {"file": "MBD hạ thế.docx", "label": "Máy biến dòng đo lường hạ thế",
     "ready": True},
    {"file": "MB áp.docx", "label": "Máy biến áp đo lường", "ready": True},
    {"file": "3pha.docx", "label": "Công tơ điện xoay chiều 3 pha",
     "ready": True},
    {"file": "Cân đồng hồ lò xo.docx", "label": "Cân đồng hồ lò xo",
     "ready": True},
    {"file": "Cân kỹ thuật.docx", "label": "Cân kỹ thuật", "ready": True},
    {"file": "Cân bàn.docx", "label": "Cân bàn", "ready": True},
    {"file": "Cân đĩa.docx", "label": "Cân đĩa", "ready": True},
    {"file": "Taximet.docx", "label": "Taximet", "ready": True},
    {"file": "KX.docx", "label": "Giấy chứng nhận kiểm xạ (KX)", "ready": False},
    {"file": "Xquang.docx", "label": "Thiết bị X-quang", "ready": False},
    {"file": "xăng dầu.docx", "label": "Cột đo xăng dầu", "ready": False},
]


def app_dir():
    """Thư mục chứa chương trình (file exe hoặc file .py)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def template_dir():
    """Ưu tiên thư mục MauWord đặt cạnh chương trình — cho phép bổ sung/sửa
    mẫu mà không phải build lại; nếu không có thì dùng bản gói trong exe."""
    local = os.path.join(app_dir(), TEMPLATE_DIR_NAME)
    if os.path.isdir(local):
        return local
    return os.path.join(getattr(sys, "_MEIPASS", app_dir()), TEMPLATE_DIR_NAME)


def available_templates():
    """Các mẫu thực sự có file, kèm đường dẫn đầy đủ."""
    folder = template_dir()
    found = []
    for t in TEMPLATES:
        path = os.path.join(folder, t["file"])
        if os.path.isfile(path):
            found.append(dict(t, path=path))
    return found


def describe_document(path):
    """Đọc nội dung một file Word: trả về (danh sách (trường, giá trị),
    danh sách dòng chữ, số trang) — dùng cho nút 'Xem mẫu'."""
    doc_path = ensure_docx(path)
    with zipfile.ZipFile(doc_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    values = {}
    lines = []
    for para, _page in iter_paragraphs_with_page(root):
        text = para_text(para).strip()
        if not text or len(text) > 400:
            continue
        lines.append(text)
        for f in FIELDS:
            if f["key"] in values:
                continue
            m = f["pattern"].search(text)
            if m:
                values[f["key"]] = m.group("val").strip().rstrip(".").strip()
    n_pages, _starts, _t = _scan_part(root, [{}], False, True, None,
                                      collect_starts=True)
    fields = [(f["display"], values[f["key"]]) for f in FIELDS
              if f["key"] in values]
    return fields, lines, n_pages


def render_document_images(path, max_pages=3, dpi=105):
    """Dựng ảnh các trang đầu của file Word (Word xuất PDF -> ảnh PNG).
    Trả về danh sách bytes PNG; rỗng nếu máy không dựng được ảnh."""
    try:
        import pythoncom
        import win32com.client
        import fitz
    except ImportError:
        return []
    doc_path = os.path.abspath(ensure_docx(path))
    pdf_path = os.path.join(tempfile.gettempdir(), "_xem_mau.pdf")
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(doc_path, ReadOnly=True,
                                    AddToRecentFiles=False)
            # 17 = wdExportFormatPDF, 3 = wdExportFromTo (giới hạn số trang)
            d.ExportAsFixedFormat(pdf_path, 17, False, 0, 3, 1, max_pages)
            d.Close(False)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
    except Exception:
        return []
    finally:
        pythoncom.CoUninitialize()
    images = []
    try:
        with fitz.open(pdf_path) as pdf:
            for page in list(pdf)[:max_pages]:
                images.append(page.get_pixmap(dpi=dpi).tobytes("png"))
    except Exception:
        return []
    return images


def suggest_template(records, templates):
    """Đoán mẫu phù hợp với dữ liệu Excel: so tên phương tiện đo."""
    names = {(r.get("ten") or "").strip().lower() for r in records}
    names.discard("")
    if not names:
        return None
    for t in templates:
        if not t.get("ready"):
            continue
        try:
            fields, _lines, _n = describe_document(t["path"])
        except Exception:
            continue
        tpl_name = next((v for d, v in fields if d == "Tên"), "").strip().lower()
        if tpl_name and tpl_name in names:
            return t
    return None


# ---------------------------------------------------------------------------
# Chuyển .doc -> .docx (cần Microsoft Word trên máy)
# ---------------------------------------------------------------------------
def ensure_docx(path):
    """Trả về đường dẫn .docx; nếu là .doc thì chuyển đổi bằng MS Word."""
    if path.lower().endswith(".docx"):
        return path
    if not path.lower().endswith(".doc"):
        raise ValueError("Chỉ hỗ trợ file .doc hoặc .docx: %s" % path)
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError(
            "Bản chương trình này chưa hỗ trợ file .doc.\n"
            "Hãy mở file bằng Microsoft Word, chọn File > Save As,\n"
            "lưu sang định dạng .docx rồi chọn lại."
        )
    out = os.path.join(tempfile.gettempdir(),
                       os.path.splitext(os.path.basename(path))[0] + "_conv.docx")
    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except Exception:
            raise RuntimeError(
                "Không mở được Microsoft Word để chuyển file .doc sang .docx.\n"
                "Máy này cần cài Microsoft Word, hoặc bạn tự mở file,\n"
                "chọn File > Save As, lưu sang .docx rồi chọn lại."
            )
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                      AddToRecentFiles=False)
            doc.SaveAs2(out, FileFormat=16)  # wdFormatXMLDocument
            doc.Close(False)
        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(
                "Chuyển file .doc sang .docx thất bại:\n%s\n\n"
                "Bạn có thể tự mở file bằng Word, lưu sang .docx rồi chọn lại." % e
            )
        finally:
            try:
                word.Quit()
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()
    return out


# ---------------------------------------------------------------------------
# Tiện ích XML
# ---------------------------------------------------------------------------
def xml_part_names(z):
    """document.xml trước, sau đó header/footer."""
    names = [n for n in z.namelist()
             if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)]
    names.sort(key=lambda n: (0 if n == "word/document.xml" else 1, n))
    return names


def is_under(el, tag):
    a = el.getparent()
    while a is not None:
        if a.tag == tag:
            return True
        a = a.getparent()
    return False


def nearest_ancestor(el, tag):
    a = el.getparent()
    while a is not None:
        if a.tag == tag:
            return a
        a = a.getparent()
    return None


def own_texts(p):
    """Các node w:t thuộc CHÍNH đoạn văn p (không lấy của đoạn lồng trong
    textbox con) — tránh chuỗi gộp khổng lồ ở đoạn bao ngoài."""
    result = []
    for t in p.iter(W + "t"):
        a = t.getparent()
        while a is not None:
            if a.tag == W + "p":
                if a is p:
                    result.append(t)
                break
            a = a.getparent()
    return result


def para_text(p):
    return "".join(t.text or "" for t in own_texts(p))


def set_t_text(t, text):
    t.text = text
    if text != text.strip():
        t.set(XML_SPACE, "preserve")


def replace_span_in_texts(ts, start, end, new):
    """Thay đúng đoạn ký tự [start, end) của chuỗi ghép từ dãy w:t bằng new,
    giữ nguyên run/định dạng. Dùng vị trí thay vì tìm chuỗi để không đụng
    nhầm chỗ khác (vd giá trị "1" của Cấp chính xác nằm trong số "4310")."""
    texts = [t.text or "" for t in ts]
    pos = 0
    replaced = False
    for t, txt in zip(ts, texts):
        s, e = pos, pos + len(txt)
        pos = e
        if e <= start or s >= end:
            continue
        head = txt[: max(0, start - s)]
        tail = txt[max(0, min(len(txt), end - s)):]
        if not replaced:
            set_t_text(t, head + new + tail)
            replaced = True
        else:
            set_t_text(t, head + tail)
    return 1 if replaced else 0


def field_spans(pattern, full, value):
    """Vị trí các lần trường xuất hiện trong đoạn văn với đúng giá trị cần
    thay (dùng cả cho bản chính lẫn bản sao dự phòng mc:Fallback)."""
    spans = []
    for m in pattern.finditer(full):
        if m.group("val").strip() == value:
            spans.append((m.start("val"), m.end("val")))
    return spans


def replace_in_texts(ts, old, new):
    """Thay 'old' bằng 'new' trong dãy node w:t, giữ nguyên run/định dạng."""
    count = 0
    search_from = 0
    while True:
        texts = [t.text or "" for t in ts]
        full = "".join(texts)
        # tìm tiếp SAU vị trí vừa thay — tránh lặp vô hạn khi giá trị mới
        # chứa giá trị cũ (vd "Máy biến dòng" trong "Máy biến dòng đo lường")
        idx = full.find(old, search_from)
        if idx == -1:
            return count
        end = idx + len(old)
        pos = 0
        replaced = False
        for t, txt in zip(ts, texts):
            s, e = pos, pos + len(txt)
            pos = e
            if e <= idx or s >= end:
                continue
            head = txt[: max(0, idx - s)]
            tail = txt[max(0, min(len(txt), end - s)):]
            if not replaced:
                set_t_text(t, head + new + tail)
                replaced = True
            else:
                set_t_text(t, head + tail)
        search_from = idx + len(new)
        count += 1


def iter_paragraphs_with_page(root):
    """Duyệt mọi w:p theo thứ tự tài liệu (bỏ bản sao mc:Fallback),
    kèm số trang ước lượng của phần thân (tăng khi gặp ngắt trang)."""
    page = 0
    for el in root.iter():
        tag = el.tag
        if tag == W + "p":
            if not is_under(el, MC + "Fallback"):
                yield el, page
        elif tag == W + "br":
            if el.get(W + "type") == "page":
                page += 1
        elif tag == W + "lastRenderedPageBreak":
            page += 1
        elif tag == W + "sectPr":
            page += 1


def container_of(p, body_page):
    """Định danh 'khối chứa' đoạn văn: textbox nào, hoặc trang nào của thân."""
    tx = nearest_ancestor(p, W + "txbxContent")
    if tx is not None:
        return ("tx", id(tx))
    return ("body", body_page)


# ---------------------------------------------------------------------------
# Đọc các bộ dữ liệu từ Tài liệu A
# ---------------------------------------------------------------------------
def extract_records_from_a(docx_path):
    """Lấy TẤT CẢ các bộ dữ liệu (ten/so/tem) theo thứ tự xuất hiện.
    Gặp lại một trường đã có trong bộ hiện tại -> bắt đầu bộ mới."""
    records = []
    current = {}

    def push():
        if current and (not records or records[-1] != current):
            records.append(dict(current))

    with zipfile.ZipFile(docx_path) as z:
        for name in xml_part_names(z):
            root = etree.fromstring(z.read(name))
            for p, _page in iter_paragraphs_with_page(root):
                text = para_text(p).strip()
                if not text:
                    continue
                for f in FIELDS:
                    if f.get("auto_only"):
                        continue  # số giấy chứng nhận do chương trình tự đánh
                    m = f["pattern"].search(text)
                    if not m:
                        continue
                    val = m.group("val").strip().rstrip(".").strip()
                    if f["key"] in current:
                        push()
                        current.clear()
                    current[f["key"]] = val
    push()
    return records


def template_field_value(path, key):
    """Giá trị một trường đang in sẵn trong mẫu Word (chỉ đọc .docx cho nhanh,
    không mở Microsoft Word). Trả về '' nếu không tìm thấy."""
    if not path or not str(path).lower().endswith(".docx"):
        return ""
    try:
        with zipfile.ZipFile(path) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    except Exception:
        return ""
    pattern = next((f["pattern"] for f in FIELDS if f["key"] == key), None)
    if pattern is None:
        return ""
    for para, _page in iter_paragraphs_with_page(root):
        m = pattern.search(para_text(para))
        if m:
            return m.group("val").strip().rstrip(".").strip()
    return ""


def template_start_number(path):
    """Số (N°) đang có sẵn trong mẫu — dùng làm số bắt đầu mặc định khi
    khách không nhập gì."""
    return template_field_value(path, "so_gcn")


_ADDRESS_WORD = (r"(?:xã|phường|thị\s+trấn|thị\s+xã|huyện|quận|tỉnh|"
                 r"thành\s+phố|tp\.?|thôn|bản|ấp|tổ\s+dân\s+phố|tổ|"
                 r"khu\s+phố|khu|đường|phố|số\s+nhà|ngõ|chợ)")
_ADDR_CUT = re.compile(r",\s*" + _ADDRESS_WORD + r"\b", re.IGNORECASE)


def _org_name(value):
    """Tách tên đơn vị khỏi phần địa chỉ đi kèm.
    'Trạm y tế Trấn Yên, xã Trấn Yên, tỉnh Lào Cai' -> 'Trạm y tế Trấn Yên'."""
    text = (value or "").strip()
    m = _ADDR_CUT.search(text)
    return text[:m.start()].strip().rstrip(",").strip() if m else text


def apply_user_fields(records, template_path=""):
    """Dòng "Người/Đơn vị sử dụng" trên giấy chứng nhận thường CHỈ ghi tên đơn
    vị, còn địa chỉ nằm ở dòng "Nơi sử dụng". Biên bản Excel lại ghi gộp cả
    hai, làm dòng này dài quá khổ, xuống hàng và đẩy lệch phần cuối trang so
    với các trang khác. Nếu mẫu Word đang theo lối chỉ ghi tên thì cắt bớt
    địa chỉ cho khớp."""
    tpl_value = template_field_value(template_path, "nguoi_sd")
    if tpl_value and _ADDR_CUT.search(tpl_value):
        return []          # mẫu vốn ghi cả địa chỉ -> giữ nguyên
    count = 0
    for r in records:
        full = r.get("nguoi_sd")
        if not full:
            continue
        short = _org_name(full)
        if short and short != full:
            r["nguoi_sd"] = short
            count += 1
    if not count:
        return []
    return ["Đã bỏ phần địa chỉ khỏi \"Người/Đơn vị sử dụng\" của %d giấy cho "
            "khớp mẫu (địa chỉ vẫn giữ ở dòng \"Nơi sử dụng\")" % count]


def apply_validity(records, template_path=""):
    """Tính 'Thời hạn đến' cho các bộ chưa có (Excel không ghi 'Hiệu lực đến').

    Số năm tra theo phương pháp thực hiện trong CauHinhThoiHan.txt. Biểu mẫu
    nào không ghi phương pháp (vd biên bản cân bàn) thì lấy phương pháp in
    sẵn trong mẫu Word. Trả về danh sách cảnh báo."""
    need = [r for r in records if not r.get("thoi_han") and r.get("_ngay_kd")]
    if not need:
        return []
    rules = validity_rules()
    default_method = ""
    if any(not r.get("phuong_phap") for r in need):
        default_method = template_field_value(template_path, "phuong_phap")
    missing = {}
    for r in need:
        method = r.get("phuong_phap") or default_method
        months = rules.get(_norm_method(method))
        if months:
            han = _add_months(r["_ngay_kd"], months)
            if han:
                r["thoi_han"] = han
                continue
        missing[method or "(không rõ phương pháp)"] = \
            missing.get(method or "(không rõ phương pháp)", 0) + 1
    notes = []
    for method, count in sorted(missing.items()):
        notes.append("Chưa cấu hình thời hạn cho phương pháp \"%s\" (%d giấy) — "
                     "giữ nguyên thời hạn của mẫu; bấm \"Cấu hình thời hạn...\" "
                     "để bổ sung" % (method, count))
    return notes


def apply_auto_numbering(records, start_no):
    """Đánh số giấy chứng nhận tăng dần theo trang.

    Nhận cả dạng đầy đủ "4111 - 26 (YB)" (trang sau thành "4112 - 26 (YB)"...)
    lẫn dạng chỉ có số "4111" (khi đó giữ nguyên phần đuôi của mẫu).
    Trả về số bản ghi đã đánh số (0 nếu để trống)."""
    # KHÔNG so sánh với SO_GCN_HINT ở đây: có mẫu mang đúng số đó (3pha =
    # "4302 - 26 (YB)"), so sánh sẽ nuốt mất số khách nhập thật.
    # Việc lọc chữ gợi ý mờ do giao diện làm trước khi gọi vào đây.
    text = str(start_no or "").strip()
    if not text:
        return 0
    m = re.match(r"^\s*([0-9]+)\s*(.*)$", text)
    if not m:
        raise RuntimeError(
            "Số bắt đầu phải mở đầu bằng chữ số, ví dụ: %s" % SO_GCN_HINT)
    num = int(m.group(1))
    suffix = m.group(2).strip()
    for i, r in enumerate(records):
        r["so_gcn"] = str(num + i) + ((" " + suffix) if suffix else "")
    return len(records)


# ---------------------------------------------------------------------------
# Thay thế trong Tài liệu B — trực tiếp trên XML
# ---------------------------------------------------------------------------
WP_DOCPR = ("{http://schemas.openxmlformats.org/drawingml/2006/"
            "wordprocessingDrawing}docPr")


def _top_body_child(p, body):
    a = p
    while a is not None and a.getparent() is not body:
        a = a.getparent()
    return a


def _scan_part(root, records, apply_mode, is_main, rows, collect_starts=False):
    """Một lượt quét/thay cho một phần XML. Trả về (n_blocks, starts, total)."""
    body = root.find(W + "body") if is_main else None
    block = -1
    block_vals = {}
    starts = []
    total = 0

    for p, body_page in iter_paragraphs_with_page(root):
        ts = own_texts(p)
        full = "".join(t.text or "" for t in ts)
        if not full.strip():
            continue
        for f in FIELDS:
            m = f["pattern"].search(full)
            if not m:
                continue
            old = m.group("val").strip()
            if not old:
                continue

            if is_main:
                cont = container_of(p, body_page)
                prev = block_vals.get(f["key"])
                if prev is not None and (prev[0] != old or prev[1] != cont):
                    block += 1
                    block_vals = {}
                if block < 0:
                    block = 0
                block_vals.setdefault(f["key"], (old, cont))
                if collect_starts and len(starts) <= block:
                    starts.append(_top_body_child(p, body))
                bi = block
            else:
                bi = 0  # header/footer: luôn dùng bộ đầu tiên

            rec = records[min(bi, len(records) - 1)]
            new = rec.get(f["key"])
            if not new:
                continue
            adjust = f.get("adjust")
            if adjust:
                # một số trường cần dựa vào giá trị đang có trong mẫu
                # (giữ đuôi "- 26 (YB)", giữ năm của ngày cấp...)
                new = adjust(new, old)
                if not new:
                    continue

            spans = field_spans(f["pattern"], full, old)
            n = len(spans) or 1
            if apply_mode and old != new:
                for start, end in sorted(spans, reverse=True):
                    replace_span_in_texts(ts, start, end, new)
                # đồng bộ bản sao dự phòng (mc:Fallback): chỉ sửa đúng đoạn
                # văn cùng trường & cùng giá trị, không quét cả khối
                ac = nearest_ancestor(p, MC + "AlternateContent")
                if ac is not None:
                    for fb in ac.iter(MC + "Fallback"):
                        for fp in fb.iter(W + "p"):
                            fts = own_texts(fp)
                            ftext = "".join(t.text or "" for t in fts)
                            fspans = field_spans(f["pattern"], ftext, old)
                            for start, end in sorted(fspans, reverse=True):
                                replace_span_in_texts(fts, start, end, new)
            if old != new:
                total += n

            if rows is not None:
                label = f["display"] if is_main else f["display"] + " (header/footer)"
                row = rows.setdefault((f["key"], bi if is_main else -1),
                                      {"display": label, "old": old,
                                       "new": new, "count": 0})
                row["count"] += n

    return block + 1, starts, total


def _is_empty_paragraph(el):
    if el.tag != W + "p":
        return False
    for t in el.iter(W + "t"):
        if (t.text or "").strip():
            return False
    for tag in ("drawing", "pict", "br", "object"):
        for _ in el.iter(W + tag):
            return False
    return True


def _chunk_leads_with_break(chunk):
    """Khối này đã tự mở đầu bằng một ngắt trang chưa?"""
    for el in chunk:
        for x in el.iter():
            if x.tag in (W + "drawing", W + "pict", W + "object") or \
                    (x.tag == W + "t" and (x.text or "").strip()):
                return False
            if x.tag == W + "br" and x.get(W + "type") == "page":
                return True
    return False


def _strip_trailing_page_breaks(chunk):
    """Bỏ ngắt trang nằm SAU nội dung cuối của khối.

    Nhiều mẫu kết thúc trang bằng một đoạn rỗng chứa ngắt trang. Khi nhân
    bản, dấu kết đoạn ấy rơi xuống đầu trang kế, chiếm một dòng và làm lệch
    bố cục trang đó."""
    seq = [(el, x) for el in chunk for x in el.iter()]
    last_content = -1
    for i, (_el, x) in enumerate(seq):
        if x.tag in (W + "drawing", W + "pict", W + "object") or \
                (x.tag == W + "t" and (x.text or "").strip()):
            last_content = i
    removed = 0
    for _el, x in seq[last_content + 1:]:
        if x.tag == W + "br" and x.get(W + "type") == "page":
            parent = x.getparent()
            if parent is not None:
                parent.remove(x)
                removed += 1
    return removed


def _set_page_break_before(el):
    """Đánh dấu đoạn văn này BẮT ĐẦU một trang mới (w:pageBreakBefore).

    Không chèn ngắt trang vào cuối đoạn trước: khi đó dấu kết đoạn của đoạn
    ấy rơi xuống đầu trang sau, chiếm mất một dòng trống và đẩy lệch toàn bộ
    nội dung trang đó so với các trang khác."""
    if el is None or el.tag != W + "p":
        return False
    pPr = el.find(W + "pPr")
    if pPr is None:
        pPr = etree.Element(W + "pPr")
        el.insert(0, pPr)
    if pPr.find(W + "pageBreakBefore") is not None:
        return True
    node = etree.Element(W + "pageBreakBefore")
    # theo lược đồ OOXML, pageBreakBefore đứng sau pStyle/keepNext/keepLines
    after = None
    for tag in ("pStyle", "keepNext", "keepLines"):
        found = pPr.find(W + tag)
        if found is not None:
            after = found
    if after is not None:
        after.addnext(node)
    else:
        pPr.insert(0, node)
    return True


def _make_page_break_para():
    p = etree.Element(W + "p")
    r = etree.SubElement(p, W + "r")
    br = etree.SubElement(r, W + "br")
    br.set(W + "type", "page")
    return p


def _clone_last_page(root, starts, need):
    """Nhân bản khối trang cuối của thân tài liệu để có thêm >= need trang.

    Khối mẫu có thể chứa NHIỀU trang (nhiều textbox chung một đoạn thân) —
    số lần nhân bản tính theo số trang trong khối. Nếu sau nội dung cuối
    của khối không có ngắt trang thì chèn ngắt trang trước mỗi bản sao để
    các trang không đè lên nhau."""
    import copy
    body = root.find(W + "body")
    children = list(body)
    idx = {id(c): i for i, c in enumerate(children)}
    last_start = idx[id(starts[-1])]
    end = len(children)
    if children and children[-1].tag == W + "sectPr":
        end -= 1
    while end - 1 > last_start and _is_empty_paragraph(children[end - 1]):
        end -= 1
    template = children[last_start:end]
    if not template:
        return
    # số trang nằm trong khối mẫu (các block bắt đầu từ phần tử thuộc khối)
    k_pages = max(1, sum(1 for s in starts if idx.get(id(s), -1) >= last_start))
    n_clones = -(-need // k_pages)  # làm tròn lên
    # khối mẫu có tự ngăn trang với bản sao kế tiếp không? Được coi là có
    # nếu tồn tại ngắt trang TRƯỚC nội dung đầu (kiểu [br][textbox]) hoặc
    # SAU nội dung cuối (kiểu [textbox][br])
    seq = [x for el in template for x in el.iter()]
    content_idx = [i for i, x in enumerate(seq)
                   if x.tag in (W + "drawing", W + "pict", W + "object")
                   or (x.tag == W + "t" and (x.text or "").strip())]
    first_c = content_idx[0] if content_idx else 0
    last_c = content_idx[-1] if content_idx else -1
    def is_page_br(x):
        return x.tag == W + "br" and x.get(W + "type") == "page"
    has_separator = (any(is_page_br(x) for x in seq[:first_c])
                     or any(is_page_br(x) for x in seq[last_c + 1:]))
    # tăng id docPr trong bản sao để không trùng id hình vẽ/textbox
    max_id = 0
    for d in root.iter(WP_DOCPR):
        try:
            max_id = max(max_id, int(d.get("id", "0")))
        except ValueError:
            pass
    pos = end
    chunks = [list(template)]
    for _ in range(n_clones):
        current = []
        for el in template:
            cp = copy.deepcopy(el)
            for d in cp.iter(WP_DOCPR):
                max_id += 1
                d.set("id", str(max_id))
            body.insert(pos, cp)
            pos += 1
            current.append(cp)
        chunks.append(current)

    # Chuẩn hoá cách sang trang: bỏ ngắt trang thừa ở cuối mỗi khối, rồi cho
    # mỗi khối sau TỰ bắt đầu trang mới. Nhờ vậy mọi trang bắt đầu ở cùng một
    # vị trí, không trang nào bị dôi ra một dòng trống ở đầu.
    for i, chunk in enumerate(chunks):
        _strip_trailing_page_breaks(chunk)
        if i > 0 and not _chunk_leads_with_break(chunk):
            _set_page_break_before(chunk[0])


def _split_multi_page_paragraphs(root):
    """Tách các đoạn thân chứa NHIỀU textbox-trang thành mỗi trang một đoạn.

    Một số tài liệu để 2+ giấy chứng nhận (textbox) neo chung một đoạn với
    vị trí trùng nhau -> khi in bị chồng trang. Mỗi run chứa textbox có từ
    2 trường dữ liệu trở lên được coi là một trang riêng; giữa các trang
    chèn ngắt trang. Trả về số đoạn đã tách."""
    import copy
    n_split = 0
    body = root.find(W + "body")
    if body is None:
        return 0
    for p in list(body):
        if p.tag != W + "p":
            continue
        kids = list(p)
        idxs = []
        for i, r in enumerate(kids):
            if r.tag != W + "r":
                continue
            keys = set()
            for pp in r.iter(W + "p"):
                t = para_text(pp)
                for f in FIELDS:
                    if f["key"] not in keys and f["pattern"].search(t):
                        keys.add(f["key"])
            if len(keys) >= 2:
                idxs.append(i)
        if len(idxs) <= 1:
            continue
        pPr = p.find(W + "pPr")
        pos = list(body).index(p)
        new_ps = []
        prev = idxs[0]
        for cut in idxs[1:]:
            np = etree.Element(W + "p")
            if pPr is not None:
                np.append(copy.deepcopy(pPr))
            for el in kids[prev + 1: cut + 1]:
                p.remove(el)
                np.append(el)
            new_ps.append(np)
            prev = cut
        # phần đuôi (vd run chứa ngắt trang cuối) chuyển vào đoạn cuối
        for el in kids[prev + 1:]:
            p.remove(el)
            new_ps[-1].append(el)
        # giữa hai đoạn liên tiếp phải có ngắt trang: thêm vào cuối đoạn
        # trước TRỪ KHI đoạn sau đã có ngắt trang trước nội dung đầu
        def _leads_with_break(par):
            for x in par.iter():
                if x.tag == W + "br" and x.get(W + "type") == "page":
                    return True
                if x.tag in (W + "drawing", W + "pict", W + "object") \
                        or (x.tag == W + "t" and (x.text or "").strip()):
                    return False
            return False

        chain = [p] + new_ps
        for _par, nxt in zip(chain[:-1], chain[1:]):
            if not _leads_with_break(nxt):
                _set_page_break_before(nxt)
        for k, np in enumerate(new_ps):
            body.insert(pos + 1 + k, np)
        n_split += 1
    return n_split


def _trim_extra_pages(root, starts, keep):
    """Xóa các trang thừa của mẫu: bỏ mọi phần tử thân kể từ đầu trang
    thứ keep+1 (giữ lại sectPr cuối). Trả về True nếu có xóa."""
    body = root.find(W + "body")
    children = list(body)
    idx = {id(c): i for i, c in enumerate(children)}
    cut_el = starts[keep]
    start_idx = idx[id(cut_el)]
    # trang trước và trang cần cắt nằm chung một phần tử -> giữ phần tử đó
    if keep > 0 and cut_el is starts[keep - 1]:
        start_idx += 1
    end = len(children)
    if children and children[-1].tag == W + "sectPr":
        end -= 1
    removed = False
    for el in children[start_idx:end]:
        body.remove(el)
        removed = True
    return removed


def process_document(b_path, records, out_path=None, exact=False):
    """Dò (out_path=None) hoặc thay thế rồi lưu file mới.

    - Chia B thành khối/trang theo cấu trúc (textbox, ngắt trang, hoặc gặp
      lại một trường với giá trị khác). Khối thứ k dùng bộ dữ liệu thứ k.
    - Nếu A có nhiều bộ hơn số trang của B: tự nhân bản trang cuối của B
      để đủ mỗi bộ một trang.
    - exact=True (chế độ tạo từ mẫu): B là template — cắt bớt các trang
      thừa để số trang đúng bằng số bộ dữ liệu.
    Trả về (rows, total, notes)."""
    if not records:
        raise RuntimeError("Không có bộ dữ liệu nào để thay thế.")
    apply_mode = out_path is not None
    replaced_parts = {}
    rows = {}
    notes = []
    total = 0
    n_blocks = 0
    orig_blocks = 0
    trimmed = False

    with zipfile.ZipFile(b_path) as z:
        for name in xml_part_names(z):
            raw = z.read(name)
            root = etree.fromstring(raw)
            is_main = (name == "word/document.xml")
            cloned = False

            if is_main:
                # tách các textbox-trang bị neo chung một đoạn (chống chồng trang)
                n_split = _split_multi_page_paragraphs(root)
                if n_split:
                    cloned = True
                    notes.append("Đã tách %d đoạn chứa nhiều trang chồng nhau "
                                 "thành từng trang riêng." % n_split)
                # lượt dò để biết B có bao nhiêu trang, trang cuối bắt đầu ở đâu
                orig_blocks, starts, _t = _scan_part(root, records, False,
                                                     True, None, collect_starts=True)
                need = len(records) - orig_blocks
                if orig_blocks > 0 and need > 0 and starts:
                    _clone_last_page(root, starts, need)
                    cloned = True
                    # khối mẫu nhiều trang có thể nhân dôi — cắt phần thừa
                    n2, starts2, _t2 = _scan_part(root, records, False, True,
                                                  None, collect_starts=True)
                    if n2 > len(records) and starts2:
                        _trim_extra_pages(root, starts2, len(records))
                elif exact and need < 0 and starts:
                    trimmed = _trim_extra_pages(root, starts, len(records))
                    cloned = cloned or trimmed
                if exact or cloned:
                    # dọn các đoạn rỗng cuối thân để không dư trang trắng
                    body = root.find(W + "body")
                    kids = list(body)
                    end_i = len(kids) - (1 if kids and kids[-1].tag == W + "sectPr" else 0)
                    while end_i > 0 and _is_empty_paragraph(kids[end_i - 1]):
                        body.remove(kids[end_i - 1])
                        end_i -= 1
                        cloned = True
                    # bỏ ngắt trang nằm sau nội dung cuối cùng (gây trang trắng)
                    elems = list(body.iter())
                    last_text = -1
                    for i, el in enumerate(elems):
                        if el.tag == W + "t" and (el.text or "").strip():
                            last_text = i
                    for i, el in enumerate(elems):
                        if i > last_text and el.tag == W + "br" \
                                and el.get(W + "type") == "page":
                            el.getparent().remove(el)
                            cloned = True
                n_blocks, _s, t = _scan_part(root, records, apply_mode, True, rows)
                total += t
            else:
                _b, _s, t = _scan_part(root, records, apply_mode, False, rows)
                total += t

            if apply_mode and (t > 0 or cloned):
                replaced_parts[name] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True)

        if apply_mode:
            with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in z.infolist():
                    data = replaced_parts.get(item.filename, z.read(item.filename))
                    zout.writestr(item, data)

    if 0 < orig_blocks < len(records):
        notes.append("B có %d trang, A có %d bộ dữ liệu — %s nhân bản thêm %d trang "
                     "theo mẫu trang cuối của B."
                     % (orig_blocks, len(records),
                        "đã" if apply_mode else "sẽ", len(records) - orig_blocks))
    elif trimmed:
        notes.append("Mẫu có %d trang, chỉ cần %d — %s cắt bớt %d trang thừa."
                     % (orig_blocks, len(records),
                        "đã" if apply_mode else "sẽ", orig_blocks - len(records)))
    elif n_blocks > len(records):
        notes.append("B có %d trang/khối nhưng A chỉ có %d bộ dữ liệu — "
                     "các trang cuối dùng lại bộ cuối cùng." % (n_blocks, len(records)))

    out_rows = []
    for (key, gi) in sorted(rows, key=lambda t: (FIELD_ORDER.get(t[0], 99), t[1])):
        row = rows[(key, gi)]
        label = row["display"]
        if n_blocks > 1 and gi >= 0:
            label = "%s (trang %d)" % (label, gi + 1)
            if gi >= orig_blocks:
                label += " *mới"
        out_rows.append((label, row["old"], row["new"], row["count"]))
    return out_rows, total, notes


# ---------------------------------------------------------------------------
# Đọc dữ liệu từ file Excel (sheet BIEN_BAN_IN) để tạo giấy chứng nhận
# ---------------------------------------------------------------------------
EXCEL_SHEET = "BIEN_BAN_IN"
EXCEL_SKIP_PREFIX = "bien ban giao"

# nhãn viết thành câu trong Excel -> khóa trường Word
EXCEL_LABELS = [
    ("ten", re.compile(r"(?:Loại công tơ|Tên (?:đối tượng|phương tiện|thiết bị)"
                       r"(?:\s*đo)?)\s*:?\s*(?P<val>.*)$")),
    ("kieu", re.compile(r"^\s*-?\s*Kiểu\s*:?\s*(?P<val>.*)$")),
    ("noi_sx", re.compile(r"(?:Nơi|Nước|Cơ sở) sản xuất\s*:?\s*(?P<val>.*)$")),
    ("nam", re.compile(r"Năm sản xuất\s*:?\s*(?P<val>.*)$")),
    # số máy ghi thẳng trong câu (vd "Cơ sở sản xuất: Việt Nam.   Số: 10797");
    # chỉ nhận dãy số đứng trọn để không nhầm số biên bản "Số: 55/BB-KL-CB"
    ("so", re.compile(r"(?<![\w])Số\s*:\s*(?P<val>[0-9]{3,})\s*\.?\s*$")),
    ("cap_cx", re.compile(r"Cấp chính xác\s*:?\s*(?P<val>.*)$")),
    ("noi_sd", re.compile(
        r"(?:(?:Nơi|Đơn vị|Cơ sở) sử dụng|Đơn vị chủ quản)\s*:?\s*(?P<val>.*)$")),
    ("dia_diem", re.compile(
        r"(?:Địa điểm thực hiện|Nơi đặt)\s*:?\s*(?P<val>.*)$")),
    ("phuong_phap", re.compile(r"Phương pháp thực hiện\s*:?\s*(?P<val>.*)$")),
    ("che_do", re.compile(r"Chế độ kiểm định\s*:?\s*(?P<val>.*)$")),
    # lấy đúng mã tem, không nuốt phần "Tem NP: ..." ghi cùng dòng
    ("tem", re.compile(r"[Tt]em kiểm định\s*:?\s*(?P<val>\S*)")),
    # biên bản taximet: đơn vị quản lý + biển số xe
    ("donvi", re.compile(r"Tên đơn vị quản lý\s*:?\s*(?P<val>.*)$")),
    ("bien_so", re.compile(r"Biển (?:đăng ký|số)\s*xe\s*:?\s*(?P<val>.*)$")),
    ("hieu_luc", re.compile(r"Hiệu lực đến\s*:?\s*(?P<val>.*)$")),
    # ngày kiểm định — có nơi ghi "Ngày thực hiện: : 05/02/2026." (2 dấu hai chấm)
    ("ngay_kd", re.compile(
        r"Ngày (?:thực hiện|kiểm(?:\s*định)?|kđ)\s*[:\s]*(?P<val>.*)$")),
]

# bảng "vị trí kiểm" (nhiều công tơ trên một biên bản)
EXCEL_ROW_LABELS = [
    ("vi_tri", re.compile(r"Vị trí kiểm")),
    ("so", re.compile(r"Số công tơ")),
    ("nam", re.compile(r"Năm sản xuất")),
    ("hang_so", re.compile(r"Hằng số công tơ")),
]

# khối "Thông số kỹ thuật cân": ô chỉ chứa nhãn, giá trị nằm ô kế bên,
# đơn vị (kg/g/mg...) nằm ô kế tiếp nữa
EXCEL_SPEC_LABELS = [
    ("so", re.compile(r"^\s*Số\s*:?\s*$")),
    ("nam", re.compile(r"^\s*Năm\s*sx\s*:?\s*$")),
    ("noi_sx", re.compile(r"^\s*Nước\s*sx\s*:?\s*$")),
    ("max_can", re.compile(r"^\s*Max\s*=?\s*$")),
    ("min_can", re.compile(r"^\s*Min\s*=?\s*$")),
    ("d_can", re.compile(r"^\s*d\s*=?\s*$")),
    ("e_can", re.compile(r"^\s*e\s*=?\s*$")),
]

_UNIT_RX = re.compile(r"^(?:kg|g|mg|t|tấn|ml|l|m)$", re.IGNORECASE)


def _fmt_month_first(number_format):
    """Định dạng ngày của ô có đặt THÁNG trước NGÀY không (kiểu Mỹ m/d/yyyy)?"""
    fmt = re.sub(r"\[[^\]]*\]|\"[^\"]*\"|\\.", "", str(number_format or "")).lower()
    if "d" not in fmt and "y" not in fmt:
        return False        # định dạng giờ ("h:mm") — "m" ở đây là phút
    for ch in fmt:
        if ch in "ymd":
            # chỉ khi THÁNG đứng đầu mới là kiểu Mỹ; "yyyy-mm-dd" bắt đầu
            # bằng năm nên vẫn đọc theo giá trị thật của ô
            return ch == "m"
    return False


def _cell_text(v, number_format=""):
    """Đổi giá trị ô Excel về chuỗi gọn (số nguyên bỏ .0, ngày -> dd/mm/yyyy).

    Với ô ngày định dạng kiểu Mỹ (m/d/yyyy), chuỗi Excel hiển thị là
    "5/8/2026" nhưng người dùng đọc theo kiểu Việt là NGÀY 5 THÁNG 8 —
    nên lấy đúng chuỗi đang hiện rồi hiểu theo thứ tự ngày/tháng.
    """
    import datetime
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        if _fmt_month_first(number_format):
            return "%02d/%02d/%04d" % (v.month, v.day, v.year)
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _cell_str(cell):
    """Như _cell_text nhưng lấy thêm định dạng của ô để đọc ngày cho đúng."""
    if cell is None:
        return ""
    return _cell_text(cell.value, getattr(cell, "number_format", ""))


# ---------------------------------------------------------------------------
# Cấu hình thời hạn hiệu lực theo phương pháp thực hiện
#   Mỗi dòng: <Phương pháp>|<số năm>     ví dụ:  ĐLVN 30: 2019|2
#   Lưu ở file cạnh chương trình để khách tự bổ sung được.
# ---------------------------------------------------------------------------
VALIDITY_FILE = "CauHinhThoiHan.txt"
VALIDITY_HEADER = (
    "# Thời hạn hiệu lực của giấy chứng nhận theo phương pháp thực hiện.\n"
    "# Mỗi dòng một quy định, dạng:  <Phương pháp thực hiện>|<số năm>\n"
    "# Ví dụ:  ĐLVN 30: 2019|2   nghĩa là phương pháp ĐLVN 30: 2019 có hạn 2 năm.\n"
    "# Muốn tính theo THÁNG thì thêm chữ t:  ĐLVN 01: 2019|18t  = 18 tháng.\n"
    "# Dòng bắt đầu bằng dấu # là ghi chú, chương trình bỏ qua.\n"
)
DEFAULT_VALIDITY = ("ĐLVN 30: 2019|2\n"
                    "ĐLVN 16: 2021|2\n"
                    "ĐLVN 39: 2019|3\n"
                    "ĐLVN 14: 2009|1\n"
                    "ĐLVN 15: 2009|1\n"
                    "ĐLVN 01: 2019|18t\n")

_VALIDITY_CACHE = {"rules": None}


def validity_file_path():
    return os.path.join(app_dir(), VALIDITY_FILE)


def _norm_method(text):
    """Bỏ mọi khoảng trắng + chữ thường để so 'ĐLVN 30:2019' == 'ĐLVN 30: 2019'."""
    return re.sub(r"\s+", "", str(text or "")).lower().rstrip(".")


def load_validity_text():
    """Nội dung file cấu hình (tạo sẵn mặc định nếu chưa có)."""
    path = validity_file_path()
    if os.path.isfile(path):
        try:
            with io.open(path, encoding="utf-8") as f:
                return f.read()
        except OSError:
            pass
    return VALIDITY_HEADER + DEFAULT_VALIDITY


def save_validity_text(text):
    with io.open(validity_file_path(), "w", encoding="utf-8") as f:
        f.write(text)
    _VALIDITY_CACHE["rules"] = None      # đọc lại ở lần dùng sau


def validity_rules(reload=False):
    """{phương pháp đã chuẩn hoá: số năm}"""
    if reload or _VALIDITY_CACHE["rules"] is None:
        rules = {}
        for line in load_validity_text().splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "|" in line[:1]:
                continue
            method, sep, value = line.rpartition("|")
            if not sep:
                continue
            # "2" = 2 năm; "18t" / "18 tháng" = 18 tháng
            m = re.match(r"^\s*(\d+)\s*(t|th|tháng|thang)?\s*$", value.strip(),
                         re.IGNORECASE)
            if not m:
                continue
            months = int(m.group(1)) * (1 if m.group(2) else 12)
            rules[_norm_method(method)] = months
        _VALIDITY_CACHE["rules"] = rules
    return _VALIDITY_CACHE["rules"]


def _add_months(date_text, months):
    """'04/9/2026' + 18 tháng -> 'ngày 04 tháng 03 năm 2028'."""
    m = re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})\s*$",
                 str(date_text or ""))
    if not m:
        return ""
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    total = (year * 12 + month - 1) + int(months)
    return "ngày %02d tháng %02d năm %d" % (day, total % 12 + 1, total // 12)


def _add_years(date_text, years):
    """'05/02/2026' + 2 năm -> 'ngày 05 tháng 02 năm 2028'."""
    return _add_months(date_text, int(years) * 12)


def _to_thoi_han(hieu_luc):
    """'29/07/2029' -> 'ngày 29 tháng 07 năm 2029' (khớp định dạng template)."""
    m = re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})\s*$", hieu_luc)
    if m:
        return "ngày %02d tháng %02d năm %s" % (int(m.group(1)), int(m.group(2)),
                                                m.group(3))
    return hieu_luc


def _checkbox_choice(text):
    """'¨ Ban đầu  þ Định kỳ  ¨ Sau sửa chữa' -> 'Định kỳ' (ô được đánh dấu)."""
    m = re.search(r"[þ☒]\s*(?P<val>[^¨þ£☐☒\t]+)", text)
    return m.group("val").strip() if m else ""


def _find_bien_ban_sheet(wb):
    """Tìm sheet chứa biên bản kiểm định. Tên sheet mỗi nơi một khác
    (BIEN_BAN_IN / Đạt / Sheet3...) nên nhận diện theo nội dung."""
    if EXCEL_SHEET in wb.sheetnames:
        return wb[EXCEL_SHEET]
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 15)):
            for cell in row:
                if isinstance(cell.value, str) and "BIÊN BẢN KIỂM ĐỊNH" in cell.value:
                    return ws
    return None


def _load_workbook(path):
    """Mở file Excel cho openpyxl.

    File .xls đời cũ được đọc bằng thư viện xlrd rồi dựng lại thành workbook
    trong bộ nhớ — KHÔNG cần Microsoft Excel, tránh hẳn lỗi COM 'Call was
    rejected by callee' khi máy đang mở sẵn Excel."""
    import openpyxl
    if not str(path).lower().endswith(".xls"):
        return openpyxl.load_workbook(path, data_only=True)

    import xlrd
    try:
        book = xlrd.open_workbook(path, formatting_info=True)
        has_format = True
    except Exception:
        book = xlrd.open_workbook(path)
        has_format = False

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in book.sheet_names():
        sh = book.sheet_by_name(name)
        title = re.sub(r"[\\/*?:\[\]]", "-", name)[:31] or "Sheet"
        ws = wb.create_sheet(title=title)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                value = cell.value
                if value is None or value == "":
                    continue
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    except Exception:
                        continue
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                new = ws.cell(r + 1, c + 1, value)
                if has_format:   # giữ định dạng để đọc ngày/tháng cho đúng
                    try:
                        xf = book.xf_list[cell.xf_index]
                        new.number_format = book.format_map[xf.format_key].format_str
                    except Exception:
                        pass
    return wb


BIEN_BAN_MARK = "BIÊN BẢN KIỂM ĐỊNH"


def _find_block_rows(ws):
    """Các khối biên bản ghi liền nhau trong một sheet.
    Trả về [(dòng đầu, dòng cuối), ...]; rỗng nếu không phải dạng này."""
    starts = []
    max_row = ws.max_row or 0
    for row in ws.iter_rows(min_row=1, max_row=max_row):
        for cell in row:
            if isinstance(cell.value, str) and BIEN_BAN_MARK in cell.value:
                starts.append(cell.row)
                break
    if len(starts) < 2:
        return []
    return [(s, (starts[i + 1] - 1) if i + 1 < len(starts) else max_row)
            for i, s in enumerate(starts)]


def _cut_at_next_label(value):
    """Một ô Excel có thể nhồi hai nhãn cách nhau bằng nhiều dấu cách
    ("Kiểu: BATM.        Năm sản xuất: 2024") — cắt lấy phần của nhãn đầu."""
    return re.split(r"\s{2,}(?=[^:\n]{1,30}:)", value)[0]


def _looks_like_label(text):
    """Ô kiểu 'Tên gì đó: ...' (là nhãn) hay chỉ là dòng chữ nối tiếp?"""
    return bool(re.match(r"^\s*[-–\s]*[^\n:]{1,45}:", text))


def _scan_label_block(ws, r_from, r_to):
    """Quét các nhãn trong một vùng dòng của sheet.
    Trả về (common, row_of, spec)."""
    common = {}
    row_of = {}
    spec = {}
    label_row = {}

    def right_values(row_idx, col, count=9):
        return [_cell_str(ws.cell(row_idx, c2))
                for c2 in range(col + 1, col + 1 + count)]

    for row in ws.iter_rows(min_row=r_from, max_row=r_to):
        for cell in row:
            if cell.value is None:
                continue
            text = _cell_str(cell)
            if not text:
                continue
            for key, rx in EXCEL_LABELS:
                if key in common:
                    continue
                m = rx.search(text)
                if m is None:
                    continue
                val = m.group("val").strip()
                if not val:  # giá trị nằm ở ô kế bên phải
                    val = next((v for v in right_values(cell.row, cell.column) if v), "")
                    # ô kế bên là một nhãn khác -> coi như không có giá trị
                    if val.rstrip().endswith((":", "=")):
                        val = ""
                val = _cut_at_next_label(val).strip().rstrip(".").strip()
                if key == "che_do" and ("þ" in val or "¨" in val):
                    val = _checkbox_choice(val)
                if val:
                    common[key] = val
                    label_row[key] = cell.row
            for key, rx in EXCEL_SPEC_LABELS:
                if key in spec:
                    continue
                if rx.match(text):
                    vals = [v for v in right_values(cell.row, cell.column, 3) if v]
                    if vals:
                        num = vals[0]
                        unit = vals[1] if len(vals) > 1 and _UNIT_RX.match(vals[1]) else ""
                        spec[key] = (num + " " + unit).strip()
            for key, rx in EXCEL_ROW_LABELS:
                if key not in row_of and isinstance(cell.value, str) \
                        and rx.search(cell.value):
                    row_of[key] = (cell.row, cell.column)

    # tên đơn vị quản lý thường xuống dòng phần địa chỉ (dòng kế, không có nhãn)
    if "donvi" in common and label_row.get("donvi", 0) < r_to:
        for cell in ws[label_row["donvi"] + 1]:
            more = _cell_str(cell).strip()
            if more and not _looks_like_label(more):
                common["donvi"] = common["donvi"].rstrip(".") + ", " + \
                    more.strip().rstrip(".")
            if more:
                break
    return common, row_of, spec


def _finish_common(common):
    """Ghép các trường phụ thành giá trị dùng cho giấy chứng nhận."""
    donvi = common.pop("donvi", "")
    bien_so = common.pop("bien_so", "")
    if donvi:
        common.setdefault("nguoi_sd", donvi)
        common.setdefault("noi_sd", donvi)
    if bien_so:
        # giấy ghi "Nơi sử dụng: Xe ô tô biển đăng ký 21A-061.28"
        common["noi_sd"] = "Xe ô tô biển đăng ký " + bien_so
    ngay_kd = common.pop("ngay_kd", "")
    if ngay_kd:
        common["ngay_cap"] = _to_thoi_han(ngay_kd)
        common["_ngay_kd"] = ngay_kd
    if "hieu_luc" in common:
        common["thoi_han"] = _to_thoi_han(common.pop("hieu_luc"))
    return common


def parse_excel_file(path):
    """Đọc biên bản trong Excel -> (danh sách bộ dữ liệu, cảnh báo).

    Ba dạng biên bản:
      - Nhiều khối "BIÊN BẢN KIỂM ĐỊNH" ghi liền trong một sheet (taximet):
        mỗi khối là một trang Word.
      - Có bảng "Vị trí kiểm": mỗi CỘT có tem là một trang Word.
      - Không có hai dạng trên (cân, thiết bị lẻ): cả file là MỘT trang Word.
    """
    wb = _load_workbook(path)
    name = os.path.basename(path)
    try:
        ws = _find_bien_ban_sheet(wb)
        if ws is None:
            return [], ["%s: không tìm thấy sheet biên bản kiểm định" % name]
        # biên bản ghi liền nhiều thiết bị trong một sheet (vd taximet):
        # mỗi khối "BIÊN BẢN KIỂM ĐỊNH" là một trang Word
        blocks = _find_block_rows(ws)
        if len(blocks) >= 2:
            units = []
            warns = []
            for i, (r_from, r_to) in enumerate(blocks, 1):
                common, _row_of, _spec = _scan_label_block(ws, r_from, r_to)
                _finish_common(common)
                tem = common.get("tem", "")
                if not tem or not re.search(r"[0-9]", tem):
                    warns.append("%s: biên bản thứ %d không có số tem kiểm định "
                                 "— bỏ qua" % (name, i))
                    continue
                common["_file"] = name
                units.append(common)
            return units, warns

        common, row_of, spec = _scan_label_block(ws, 1,
                                                 min(ws.max_row or 0, 300))
        warns = []
        _finish_common(common)
        dia_diem = common.pop("dia_diem", "")
        if "noi_sd" in common:
            common.setdefault("nguoi_sd", common["noi_sd"])
            # "Nơi sử dụng" trong giấy chứng nhận là đơn vị + địa chỉ; chỉ ghép
            # thêm địa điểm khi tên đơn vị chưa bao gồm địa chỉ đó
            if dia_diem and dia_diem.lower() not in common["noi_sd"].lower():
                common["noi_sd"] = common["noi_sd"] + ", " + dia_diem

        # --- dạng biên bản lẻ: cả file là một trang ---
        if "so" not in row_of or "hang_so" not in row_of:
            unit = dict(common)
            for key, val in spec.items():
                unit.setdefault(key, val)
            if "min_can" in spec and "max_can" in spec:
                lo, hi = spec["min_can"].split(), spec["max_can"].split()
                if len(lo) == 2 and len(hi) == 2 and lo[1] == hi[1]:
                    unit.setdefault("pham_vi", "%s ÷ %s %s" % (lo[0], hi[0], hi[1]))
            if not unit.get("tem") or not re.search(r"[0-9]", unit.get("tem", "")):
                warns.append("%s: không có số tem kiểm định — bỏ qua" % name)
                return [], warns
            if not unit.get("so"):
                warns.append("%s: không tìm thấy số máy" % name)
            unit["_file"] = name
            return [unit], warns

        # --- dạng bảng nhiều cột ---
        units = []
        r_so, c_label = row_of["so"]
        r_nam = row_of.get("nam", (None,))[0]
        r_vitri = row_of.get("vi_tri", (None,))[0]
        r_tem = row_of["hang_so"][0] + 1  # hàng tem nằm ngay dưới Hằng số công tơ
        max_col = min(ws.max_column or 0, 40)
        for c in range(c_label + 1, max_col + 1):
            so = _cell_str(ws.cell(r_so, c))
            tem = _cell_str(ws.cell(r_tem, c))
            if not so and not tem:
                continue
            vi_tri = _cell_str(ws.cell(r_vitri, c)) if r_vitri else ""
            # tem hợp lệ phải chứa chữ số (ô ghi "Không đạt"... coi như không có tem)
            if not tem or not re.search(r"[0-9]", tem):
                warns.append("%s: vị trí %s có số %s nhưng KHÔNG có tem%s — bỏ qua"
                             % (name, vi_tri or "?", so,
                                " (ô ghi %r)" % tem if tem else ""))
                continue
            unit = dict(common)
            unit["so"] = so
            unit["tem"] = tem
            if r_nam:
                nam = _cell_str(ws.cell(r_nam, c))
                if nam:
                    unit["nam"] = nam
            hs = _cell_str(ws.cell(row_of["hang_so"][0], c))
            if hs:
                unit["hang_so"] = hs
            unit["_file"] = name
            units.append(unit)
        return units, warns
    finally:
        wb.close()


def collect_excel_records(folder, progress=None):
    """Duyệt đệ quy thư mục, đọc mọi file Excel (bỏ 'Bien ban giao...').
    progress(done, total, text) được gọi sau mỗi file để báo tiến độ.
    Trả về (records, notes)."""
    records = []
    notes = []
    files = []
    for base, _dirs, names in os.walk(folder):
        for n in sorted(names):
            low = n.lower()
            if low.startswith("~$"):
                continue
            if low.startswith(EXCEL_SKIP_PREFIX):
                continue
            if low.endswith((".xlsx", ".xlsm", ".xls")):
                files.append(os.path.join(base, n))
    files.sort()

    for i, f in enumerate(files, 1):
        if progress:
            progress(i, len(files),
                     "Đang đọc file %d/%d: %s" % (i, len(files), os.path.basename(f)))
        try:
            units, warns = parse_excel_file(f)
        except Exception as e:
            notes.append("%s: lỗi đọc file (%s)" % (os.path.basename(f), e))
            continue
        notes.extend(warns)
        records.extend(units)
    return records, notes


# ---------------------------------------------------------------------------
# Giao diện
# ---------------------------------------------------------------------------
def run_gui():
    import queue
    import threading
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("920x640")
    root.minsize(780, 540)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True)

    FILETYPES = [("File Word", "*.docx *.doc"), ("Tất cả", "*.*")]

    def start_task(worker, on_done, buttons, bar, status_var):
        """Chạy worker(progress) ở luồng nền để giao diện không bị treo.
        progress(done, total, text): cập nhật thanh tiến độ + dòng trạng thái."""
        q = queue.Queue()
        for b in buttons:
            b.config(state="disabled")
        bar.config(mode="indeterminate")
        bar.start(60)

        def progress(done=None, total=None, text=None):
            q.put(("prog", done, total, text))

        def thread_fn():
            try:
                q.put(("done", worker(progress)))
            except Exception as e:
                q.put(("error", e))

        threading.Thread(target=thread_fn, daemon=True).start()

        def finish():
            for b in buttons:
                b.config(state="normal")
            bar.stop()
            bar.config(mode="determinate", value=0)

        def poll():
            try:
                while True:
                    item = q.get_nowait()
                    if item[0] == "prog":
                        _k, done, total, text = item
                        if text:
                            status_var.set(text)
                        if done is not None and total:
                            bar.stop()
                            bar.config(mode="determinate", maximum=total, value=done)
                        elif str(bar.cget("mode")) != "indeterminate":
                            bar.config(mode="indeterminate")
                            bar.start(60)
                    elif item[0] == "done":
                        finish()
                        on_done(item[1])
                        return
                    else:
                        finish()
                        status_var.set("Lỗi: %s" % item[1])
                        messagebox.showerror("Lỗi", str(item[1]))
                        return
            except queue.Empty:
                pass
            root.after(100, poll)

        poll()

    def make_row(frame, r, label, var, cmd=None, btn_text="Chọn...", hint=""):
        ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", pady=4)
        ent = ttk.Entry(frame, textvariable=var)
        ent.grid(row=r, column=1, sticky="ew", padx=8, pady=4)
        if cmd:
            ttk.Button(frame, text=btn_text, command=cmd).grid(row=r, column=2, pady=4)
        if hint:
            add_hint(ent, var, hint)
        return ent

    def add_hint(entry, var, hint):
        """Chữ gợi ý mờ hiện khi ô còn trống, tự mất khi người dùng gõ.
        Gắn sẵn entry.set_hint(...) để đổi gợi ý theo mẫu đang chọn."""
        state = {"hint": hint}
        normal = entry.cget("foreground") or ""

        def show_hint():
            if not var.get():
                var.set(state["hint"])
                entry.config(foreground="#9a9a9a")

        def on_focus_in(_e=None):
            if var.get() == state["hint"]:
                var.set("")
            entry.config(foreground=normal)

        def set_hint(new_hint):
            if not new_hint:
                return
            untouched = var.get() in ("", state["hint"])
            state["hint"] = new_hint
            if untouched:
                var.set("")
                show_hint()

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", lambda _e=None: show_hint())
        entry.set_hint = set_hint
        entry.hint_state = state
        show_hint()

    def ask_out_path(var_dir, var_name, messagebox):
        name = var_name.get().strip()
        if not name.lower().endswith(".docx"):
            name += ".docx"
        try:
            os.makedirs(var_dir.get(), exist_ok=True)
        except OSError as e:
            messagebox.showerror("Lỗi", "Không tạo được thư mục lưu kết quả:\n%s" % e)
            return None
        out_path = os.path.join(var_dir.get(), name)
        if os.path.exists(out_path):
            if not messagebox.askyesno("Xác nhận",
                                       "File đã tồn tại:\n%s\nGhi đè?" % out_path):
                return None
        return out_path

    # danh sách mẫu dùng chung cho cả hai tab
    tpl_list = available_templates()
    tpl_display = [t["label"] + ("" if t["ready"] else COMING_SOON)
                   for t in tpl_list]

    def show_doc_preview(path, title):
        """Cửa sổ xem mẫu: ảnh trang giấy + bảng các trường sẽ thay."""
        try:
            fields, lines, n_pages = describe_document(path)
        except Exception as e:
            messagebox.showerror("Lỗi", "Không đọc được mẫu:\n%s" % e)
            return
        win = tk.Toplevel(root)
        win.title("Xem mẫu — %s" % title)
        win.geometry("900x680")
        ttk.Label(win, text="%s   (%d trang)" % (title, n_pages),
                  font=("", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
        nb2 = ttk.Notebook(win)
        nb2.pack(fill="both", expand=True, padx=10)

        # --- tab ảnh ---
        img_tab = ttk.Frame(nb2)
        nb2.add(img_tab, text="  Hình ảnh  ")
        canvas = tk.Canvas(img_tab, background="#8a8a8a", highlightthickness=0)
        vsb = ttk.Scrollbar(img_tab, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(img_tab, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        img_tab.rowconfigure(0, weight=1)
        img_tab.columnconfigure(0, weight=1)
        note = canvas.create_text(20, 20, anchor="nw", fill="white",
                                  text="Đang dựng ảnh trang giấy, xin chờ...")
        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(-int(e.delta / 120), "units"))

        # --- tab chi tiết ---
        det = ttk.Frame(nb2, padding=8)
        nb2.add(det, text="  Chi tiết trường  ")
        ttk.Label(det, text="Các trường chương trình sẽ thay trong mẫu này:",
                  foreground="#555").pack(anchor="w", pady=(0, 4))
        tv = ttk.Treeview(det, columns=("f", "v"), show="headings", height=10)
        tv.heading("f", text="Trường")
        tv.heading("v", text="Giá trị đang có trong mẫu")
        tv.column("f", width=210, anchor="w")
        tv.column("v", width=600, anchor="w")
        for d, v in fields:
            tv.insert("", "end", values=(d, v))
        tv.pack(fill="x")
        ttk.Label(det, text="Toàn văn nội dung mẫu:",
                  foreground="#555").pack(anchor="w", pady=(10, 2))
        box = ttk.Frame(det)
        box.pack(fill="both", expand=True)
        txt = tk.Text(box, wrap="word", height=10)
        sc = ttk.Scrollbar(box, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sc.set)
        txt.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")
        txt.insert("1.0", "\n".join(lines))
        txt.config(state="disabled")
        ttk.Button(win, text="Đóng", command=win.destroy).pack(pady=8)

        # dựng ảnh ở luồng nền để cửa sổ không bị treo
        holder = {}

        def build():
            try:
                holder["imgs"] = render_document_images(path)
            except Exception:
                holder["imgs"] = []

        th = threading.Thread(target=build, daemon=True)
        th.start()

        def poll_images():
            if th.is_alive():
                win.after(150, poll_images)
                return
            pngs = holder.get("imgs") or []
            canvas.delete(note)
            if not pngs:
                canvas.create_text(
                    20, 20, anchor="nw", fill="white",
                    text="Không dựng được ảnh trên máy này.\n"
                         "(cần có Microsoft Word)\n"
                         "Bạn xem nội dung ở thẻ \"Chi tiết trường\".")
                return
            win._preview_images = []      # giữ tham chiếu kẻo ảnh bị xoá
            y = 10
            for png in pngs:
                img = tk.PhotoImage(data=base64.b64encode(png))
                win._preview_images.append(img)
                canvas.create_image(10, y, anchor="nw", image=img)
                y += img.height() + 12
            canvas.configure(scrollregion=(0, 0, 900, y))

        win.after(150, poll_images)

    def show_validity_dialog():
        """Cửa sổ cấu hình thời hạn hiệu lực theo phương pháp thực hiện."""
        win = tk.Toplevel(root)
        win.title("Cấu hình thời hạn hiệu lực")
        win.geometry("640x480")
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Mỗi dòng một quy định:  "
                            "<Phương pháp thực hiện>|<số năm>",
                  font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, foreground="#555", justify="left",
                  text="Ví dụ:  ĐLVN 30: 2019|2   →  phương pháp ĐLVN 30: 2019 "
                       "có thời hạn 2 năm.\n"
                       "Thời hạn = ngày kiểm định trong Excel cộng thêm số năm này.\n"
                       "Dòng bắt đầu bằng dấu # là ghi chú.").pack(anchor="w",
                                                                   pady=(4, 8))
        box = ttk.Frame(frm)
        box.pack(fill="both", expand=True)
        txt = tk.Text(box, wrap="none", height=14)
        sc = ttk.Scrollbar(box, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sc.set)
        txt.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")
        txt.insert("1.0", load_validity_text())
        info = tk.StringVar(value="File lưu tại: %s" % validity_file_path())
        ttk.Label(frm, textvariable=info, foreground="#555",
                  wraplength=600, justify="left").pack(anchor="w", pady=(8, 0))

        def do_save():
            content = txt.get("1.0", "end").rstrip() + "\n"
            try:
                save_validity_text(content)
            except OSError as e:
                messagebox.showerror("Lỗi", "Không lưu được cấu hình:\n%s" % e)
                return
            n = len(validity_rules(reload=True))
            info.set("Đã lưu %d quy định vào: %s" % (n, validity_file_path()))
            messagebox.showinfo("Đã lưu",
                                "Đã lưu %d quy định thời hạn." % n)

        btns_v = ttk.Frame(frm)
        btns_v.pack(pady=10)
        ttk.Button(btns_v, text="Lưu", command=do_save).pack(side="left", padx=6)
        ttk.Button(btns_v, text="Đóng", command=win.destroy).pack(side="left", padx=6)

    def make_template_row(frame, r, var_target, after_pick=None,
                          extra_button=None):
        """Hàng 'Mẫu Word': chọn mẫu có sẵn (hoặc file Word riêng ở dòng cuối
        danh sách) là điền luôn đường dẫn vào var_target. Mẫu chưa hỗ trợ thì
        báo 'sắp cập nhật' và không chọn."""
        ttk.Label(frame, text="Mẫu Word:").grid(row=r, column=0, sticky="w", pady=4)
        cb = ttk.Combobox(frame, state="readonly",
                          values=tpl_display + [CHOOSE_OTHER])
        cb.grid(row=r, column=1, sticky="ew", padx=8, pady=4)

        def do_view():
            path = var_target.get().strip()
            if not path or not os.path.isfile(path):
                messagebox.showwarning("Chưa có mẫu",
                                       "Hãy chọn một mẫu Word trước.")
                return
            show_doc_preview(path, os.path.splitext(os.path.basename(path))[0])

        btn_box = ttk.Frame(frame)
        btn_box.grid(row=r, column=2, pady=4, sticky="w")
        ttk.Button(btn_box, text="Xem mẫu", command=do_view).pack(side="left")
        if extra_button:
            ttk.Button(btn_box, text=extra_button[0],
                       command=extra_button[1]).pack(side="left", padx=(6, 0))

        def on_select(_evt=None):
            i = cb.current()
            if i < 0:
                return
            if i >= len(tpl_list):          # dòng "Chọn file Word khác..."
                p = filedialog.askopenfilename(title="Chọn file Word mẫu",
                                               filetypes=FILETYPES)
                if p:
                    var_target.set(p)
                    cb.set(os.path.basename(p))
                    if after_pick:
                        after_pick({"label": os.path.splitext(
                            os.path.basename(p))[0], "path": p})
                else:
                    cb.set("")
                return
            t = tpl_list[i]
            if not t["ready"]:
                messagebox.showinfo(
                    "Sắp cập nhật",
                    "Mẫu \"%s\" đang được cập nhật, chưa dùng được.\n\n"
                    "Bạn hãy chọn mẫu khác, hoặc chọn dòng \"%s\" ở cuối danh "
                    "sách để dùng file Word của mình." % (t["label"], CHOOSE_OTHER))
                cb.set("")
                return
            var_target.set(t["path"])
            if after_pick:
                after_pick(t)

        cb.bind("<<ComboboxSelected>>", on_select)
        return cb

    # ================= TAB 1: Word -> Word =================
    tab1 = ttk.Frame(nb, padding=12)
    nb.add(tab1, text="  Word → Word  ")
    tab1.columnconfigure(1, weight=1)

    var_a = tk.StringVar()
    var_b = tk.StringVar()
    var_dir = tk.StringVar()
    var_name = tk.StringVar()
    var_start = tk.StringVar()

    def pick_a():
        p = filedialog.askopenfilename(title="Chọn tài liệu nguồn (A)", filetypes=FILETYPES)
        if p:
            var_a.set(p)

    def pick_dir():
        p = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if p:
            var_dir.set(p)

    def after_pick_b(t):
        """Chọn mẫu xong thì điền sẵn tên file xuất, thư mục lưu và cập nhật
        gợi ý số bắt đầu theo đúng số đang có trong mẫu."""
        if not var_name.get().strip():
            var_name.set(t["label"] + ".docx")
        if not var_dir.get().strip() and var_a.get().strip():
            var_dir.set(os.path.join(os.path.dirname(var_a.get()), "KetQua"))
        ent_start.set_hint(template_start_number(t["path"]))

    make_row(tab1, 0, "Tài liệu nguồn (A):", var_a, pick_a)
    make_template_row(tab1, 1, var_b, after_pick_b)
    make_row(tab1, 2, "Thư mục lưu kết quả:", var_dir, pick_dir)
    make_row(tab1, 3, "Tên file xuất:", var_name)
    ent_start = make_row(tab1, 4, "Số bắt đầu (N°):", var_start,
                         hint=SO_GCN_HINT)

    ttk.Label(tab1, text="Xem trước các thay đổi:").grid(row=5, column=0, columnspan=3,
                                                         sticky="w", pady=(12, 2))
    cols = ("field", "old", "new", "count")
    tree = ttk.Treeview(tab1, columns=cols, show="headings", height=9)
    tree.heading("field", text="Trường")
    tree.heading("old", text="Giá trị cũ (B)")
    tree.heading("new", text="Giá trị mới (A)")
    tree.heading("count", text="Số vị trí")
    tree.column("field", width=180, anchor="w")
    tree.column("old", width=250, anchor="w")
    tree.column("new", width=250, anchor="w")
    tree.column("count", width=70, anchor="center")
    sb1 = ttk.Scrollbar(tab1, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb1.set)
    tree.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=4)
    sb1.grid(row=6, column=3, sticky="ns", pady=4)
    tab1.rowconfigure(6, weight=1)

    status = tk.StringVar(
        value="Chọn tài liệu nguồn (A), chọn mẫu có sẵn (hoặc file B), rồi bấm Xem trước.")
    ttk.Label(tab1, textvariable=status, foreground="#555", wraplength=840,
              justify="left").grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 2))

    def validate_inputs(need_out=False):
        if not var_a.get().strip() or not os.path.isfile(var_a.get()):
            messagebox.showwarning("Thiếu thông tin", "Hãy chọn Tài liệu nguồn (A) hợp lệ.")
            return False
        if not var_b.get().strip() or not os.path.isfile(var_b.get()):
            messagebox.showwarning("Thiếu thông tin", "Hãy chọn Mẫu Word hợp lệ.")
            return False
        if need_out:
            if not var_dir.get().strip():
                messagebox.showwarning("Thiếu thông tin", "Hãy chọn thư mục lưu kết quả.")
                return False
            if not var_name.get().strip():
                messagebox.showwarning("Thiếu thông tin", "Hãy nhập tên file xuất.")
                return False
        return True

    def load_inputs(a_input, b_input):
        a_path = ensure_docx(a_input)
        b_path = ensure_docx(b_input)
        records = extract_records_from_a(a_path)
        if not records:
            raise RuntimeError("Không tìm thấy trường dữ liệu nào trong Tài liệu A "
                               "(Tên / Số / Số tem kiểm định).")
        return b_path, records

    def show(rows):
        tree.delete(*tree.get_children())
        for label, old, new, count in rows:
            tree.insert("", "end", values=(label, old, new, count))

    bar1 = ttk.Progressbar(tab1, mode="determinate")
    bar1.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(4, 0))

    def tab1_worker(out_path):
        a_input = var_a.get()
        b_input = var_b.get()
        start_no = var_start.get().strip()
        if start_no == ent_start.hint_state["hint"]:
            start_no = ""   # khách chưa gõ gì, mới chỉ là chữ gợi ý mờ

        def worker(progress):
            progress(text="Đang đọc tài liệu (chuyển đổi .doc nếu cần)...")
            b_path, records = load_inputs(a_input, b_input)
            # không nhập gì -> đánh số tiếp theo số đang có trong mẫu
            apply_auto_numbering(records, start_no or template_start_number(b_path))
            progress(text="Đang %s %d bộ dữ liệu..."
                          % ("thay thế" if out_path else "dò", len(records)))
            rows, total, notes = process_document(b_path, records, out_path=out_path)
            return records, rows, total, notes
        return worker

    def do_preview():
        if not validate_inputs():
            return

        def on_done(result):
            records, rows, total, notes = result
            show(rows)
            msg = "Tài liệu A có %d bộ dữ liệu. Sẽ thay %d vị trí." % (len(records), total)
            if notes:
                msg += "\nLưu ý: " + " ".join(notes)
            status.set(msg)
            if not rows:
                messagebox.showinfo("Xem trước",
                                    "Không tìm thấy trường nào trong B khớp với dữ liệu từ A.")

        start_task(tab1_worker(None), on_done, tab1_btns, bar1, status)

    def do_replace():
        if not validate_inputs(need_out=True):
            return
        out_path = ask_out_path(var_dir, var_name, messagebox)
        if not out_path:
            return

        def on_done(result):
            records, rows, total, notes = result
            show(rows)
            msg = "Đã thay %d vị trí. Đã lưu: %s" % (total, out_path)
            if notes:
                msg += "\nLưu ý: " + " ".join(notes)
            status.set(msg)
            messagebox.showinfo("Hoàn tất",
                                "Đã tạo file mới:\n%s\n\nSố vị trí đã thay: %d"
                                % (out_path, total))

        start_task(tab1_worker(out_path), on_done, tab1_btns, bar1, status)

    btns = ttk.Frame(tab1)
    btns.grid(row=9, column=0, columnspan=3, pady=8)
    tab1_btns = [ttk.Button(btns, text="Xem trước", command=do_preview),
                 ttk.Button(btns, text="Bắt đầu thay thế", command=do_replace)]
    for b in tab1_btns:
        b.pack(side="left", padx=6)

    # ================= TAB 2: Excel -> Word =================
    tab2 = ttk.Frame(nb, padding=12)
    nb.add(tab2, text="  Excel → Word  ")
    tab2.columnconfigure(1, weight=1)

    var_folder = tk.StringVar()
    var_tpl = tk.StringVar()
    var_dir2 = tk.StringVar()
    var_name2 = tk.StringVar()
    var_start2 = tk.StringVar()

    # mặc định dùng mẫu công tơ 3 pha (dạng hay gặp nhất của luồng Excel)
    _default_tpl = next((t["path"] for t in tpl_list
                         if t["ready"] and t["file"].startswith("3pha")), "")
    if not _default_tpl:
        _fallback = os.path.join(app_dir(), "template.docx")
        _default_tpl = _fallback if os.path.isfile(_fallback) else ""
    if _default_tpl:
        var_tpl.set(_default_tpl)

    def pick_folder():
        p = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel")
        if p:
            var_folder.set(p)
            if not var_dir2.get().strip():
                var_dir2.set(os.path.join(p, "KetQua"))
            if not var_name2.get().strip():
                var_name2.set("Giay chung nhan %s.docx" % os.path.basename(p))

    def pick_dir2():
        p = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if p:
            var_dir2.set(p)

    def after_pick_tpl(t):
        ent_start2.set_hint(template_start_number(t["path"]))

    make_row(tab2, 0, "Thư mục Excel:", var_folder, pick_folder)
    tpl_combo2 = make_template_row(
        tab2, 1, var_tpl, after_pick_tpl,
        extra_button=("Cấu hình thời hạn...", show_validity_dialog))
    make_row(tab2, 2, "Thư mục lưu kết quả:", var_dir2, pick_dir2)
    make_row(tab2, 3, "Tên file xuất:", var_name2)
    ent_start2 = make_row(tab2, 4, "Số bắt đầu (N°):", var_start2,
                          hint=SO_GCN_HINT)
    ent_start2.set_hint(template_start_number(var_tpl.get()))

    # hiển thị sẵn tên mẫu mặc định trên ô chọn
    for _i, _t in enumerate(tpl_list):
        if _t["path"] == var_tpl.get():
            tpl_combo2.set(tpl_display[_i])
            break

    ttk.Label(tab2, text="Xem trước các trang sẽ tạo:").grid(
        row=5, column=0, columnspan=3, sticky="w", pady=(12, 2))
    cols2 = ("page", "file", "so", "nam", "tem")
    tree2 = ttk.Treeview(tab2, columns=cols2, show="headings", height=9)
    tree2.heading("page", text="Trang")
    tree2.heading("file", text="File Excel")
    tree2.heading("so", text="Số công tơ")
    tree2.heading("nam", text="Năm SX")
    tree2.heading("tem", text="Tem kiểm định")
    tree2.column("page", width=55, anchor="center")
    tree2.column("file", width=330, anchor="w")
    tree2.column("so", width=130, anchor="w")
    tree2.column("nam", width=70, anchor="center")
    tree2.column("tem", width=110, anchor="w")
    sb2 = ttk.Scrollbar(tab2, orient="vertical", command=tree2.yview)
    tree2.configure(yscrollcommand=sb2.set)
    tree2.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=4)
    sb2.grid(row=6, column=3, sticky="ns", pady=4)
    tab2.rowconfigure(6, weight=1)

    status2 = tk.StringVar(value="Chọn thư mục Excel và mẫu Word rồi bấm Xem trước.")
    ttk.Label(tab2, textvariable=status2, foreground="#555", wraplength=840,
              justify="left").grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 2))

    def validate2(need_out=False):
        if not var_folder.get().strip() or not os.path.isdir(var_folder.get()):
            messagebox.showwarning("Thiếu thông tin", "Hãy chọn thư mục chứa file Excel.")
            return False
        if not var_tpl.get().strip() or not os.path.isfile(var_tpl.get()):
            messagebox.showwarning("Thiếu thông tin", "Hãy chọn file Word mẫu (template).")
            return False
        if need_out:
            if not var_dir2.get().strip():
                messagebox.showwarning("Thiếu thông tin", "Hãy chọn thư mục lưu kết quả.")
                return False
            if not var_name2.get().strip():
                messagebox.showwarning("Thiếu thông tin", "Hãy nhập tên file xuất.")
                return False
        return True

    bar2 = ttk.Progressbar(tab2, mode="determinate")
    bar2.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(4, 0))

    def show2(records):
        tree2.delete(*tree2.get_children())
        for i, r in enumerate(records, 1):
            tree2.insert("", "end", values=(i, r.get("_file", ""), r.get("so", ""),
                                            r.get("nam", ""), r.get("tem", "")))

    def tab2_worker(out_path):
        folder = var_folder.get()
        tpl_input = var_tpl.get()
        start_no = var_start2.get().strip()
        if start_no == ent_start2.hint_state["hint"]:
            start_no = ""   # khách chưa gõ gì, mới chỉ là chữ gợi ý mờ

        def worker(progress):
            progress(text="Đang quét thư mục Excel...")
            records, notes = collect_excel_records(folder, progress)
            if not records:
                raise RuntimeError("Không đọc được bộ dữ liệu nào từ thư mục Excel.\n"
                                   + "\n".join(notes[:10]))
            # không nhập gì -> đánh số tiếp theo số đang có trong mẫu
            apply_auto_numbering(records,
                                 start_no or template_start_number(tpl_input))
            notes.extend(apply_user_fields(records, tpl_input))
            notes.extend(apply_validity(records, tpl_input))
            pnotes = []
            total = 0
            suggest = None
            if out_path:
                progress(text="Đang chuẩn bị file Word mẫu...")
                tpl = ensure_docx(tpl_input)
                progress(text="Đang tạo file Word %d trang (bước này có thể mất "
                              "vài phút, xin chờ)..." % len(records))
                _rows, total, pnotes = process_document(tpl, records,
                                                        out_path=out_path, exact=True)
            else:
                progress(text="Đang dò mẫu Word phù hợp với dữ liệu...")
                try:
                    suggest = suggest_template(records, tpl_list)
                except Exception:
                    suggest = None
            return records, notes, total, pnotes, suggest
        return worker

    def do_preview2():
        if not validate2():
            return

        def on_done(result):
            records, notes, _total, _pnotes, suggest = result
            show2(records)
            n_files = len(set(r.get("_file") for r in records))
            msg = "Đọc được %d bộ dữ liệu từ %d file Excel — sẽ tạo %d trang." \
                  % (len(records), n_files, len(records))
            if notes:
                msg += "\nLưu ý: " + " | ".join(notes[:6])
                if len(notes) > 6:
                    msg += " | (+%d cảnh báo nữa)" % (len(notes) - 6)
            status2.set(msg)
            # gợi ý mẫu khớp với dữ liệu vừa đọc
            if suggest and suggest["path"] != var_tpl.get():
                if messagebox.askyesno(
                        "Gợi ý mẫu",
                        "Dữ liệu Excel là \"%s\", khớp với mẫu:\n\n    %s\n\n"
                        "Dùng mẫu này không?"
                        % (records[0].get("ten", ""), suggest["label"])):
                    var_tpl.set(suggest["path"])
                    for i, t in enumerate(tpl_list):
                        if t["path"] == suggest["path"]:
                            tpl_combo2.set(tpl_display[i])
                            break
                    status2.set(msg + "\nĐã chuyển sang mẫu: " + suggest["label"])

        start_task(tab2_worker(None), on_done, tab2_btns, bar2, status2)

    def do_generate():
        if not validate2(need_out=True):
            return
        out_path = ask_out_path(var_dir2, var_name2, messagebox)
        if not out_path:
            return

        def on_done(result):
            records, notes, total, pnotes, _suggest = result
            show2(records)
            msg = "Đã tạo %d trang (%d vị trí thay). Đã lưu: %s" \
                  % (len(records), total, out_path)
            allnotes = notes + pnotes
            if allnotes:
                msg += "\nLưu ý: " + " | ".join(allnotes[:6])
            status2.set(msg)
            messagebox.showinfo("Hoàn tất",
                                "Đã tạo file:\n%s\n\nSố trang: %d"
                                % (out_path, len(records)))

        start_task(tab2_worker(out_path), on_done, tab2_btns, bar2, status2)

    btns2 = ttk.Frame(tab2)
    btns2.grid(row=9, column=0, columnspan=3, pady=8)
    tab2_btns = [ttk.Button(btns2, text="Xem trước", command=do_preview2),
                 ttk.Button(btns2, text="Bắt đầu tạo", command=do_generate)]
    for b in tab2_btns:
        b.pack(side="left", padx=6)

    root.mainloop()


# ---------------------------------------------------------------------------
# Chế độ dòng lệnh (kiểm thử):  python ThayTheDuLieu.py A.docx B.docx out.docx
# ---------------------------------------------------------------------------
def run_cli(a, b, out, start_no=""):
    # console Windows có thể dùng cp1252 không in được tiếng Việt;
    # exe dạng --windowed thậm chí không có stdout
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    a_path = ensure_docx(a)
    b_path = ensure_docx(b)
    records = extract_records_from_a(a_path)
    apply_auto_numbering(records, start_no or template_start_number(b_path))
    print("Tai lieu A co %d bo du lieu:" % len(records))
    for i, r in enumerate(records[:10], 1):
        print("  Bo %d: %s" % (i, r))
    if len(records) > 10:
        print("  ... va %d bo nua" % (len(records) - 10))
    rows, total, notes = process_document(b_path, records, out_path=out)
    for label, old, new, count in rows:
        print("  %-32s | %-28s -> %-28s (x%d)" % (label, old, new, count))
    for n in notes:
        print("  LUU Y:", n)
    print("Da thay %d vi tri. Luu: %s" % (total, out))


def run_cli_excel(folder, template, out, start_no=""):
    # bản exe chạy chế độ cửa sổ không có màn hình chữ -> ghi nhật ký ra
    # file <tên file xuất>.log để còn xem được khi cần đối chiếu
    if sys.stdout is None:
        try:
            sys.stdout = io.open(out + ".log", "w", encoding="utf-8")
        except OSError:
            sys.stdout = open(os.devnull, "w", encoding="utf-8")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    records, notes = collect_excel_records(folder)
    apply_auto_numbering(records, start_no or template_start_number(template))
    notes.extend(apply_user_fields(records, template))
    notes.extend(apply_validity(records, template))
    print("Doc duoc %d bo du lieu" % len(records))
    for i, r in enumerate(records[:8], 1):
        print("  Trang %d [%s]: so=%s nam=%s tem=%s" %
              (i, r.get("_file"), r.get("so"), r.get("nam"), r.get("tem")))
    if len(records) > 8:
        print("  ...")
    for n in notes:
        print("  LUU Y:", n)
    tpl = ensure_docx(template)
    _rows, total, pnotes = process_document(tpl, records, out_path=out, exact=True)
    for n in pnotes:
        print("  LUU Y:", n)
    print("Da tao %d trang (%d vi tri thay). Luu: %s" % (len(records), total, out))


def run_cli_list_templates():
    """Liệt kê các word mẫu chương trình đang nhận (để kiểm tra nhanh)."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    found = available_templates()
    print(APP_TITLE)
    print("Thu muc mau:", template_dir())
    print("Tim thay %d/%d mau:" % (len(found), len(TEMPLATES)))
    for t in found:
        print("  [%s] %s  (%s)"
              % ("v" if t["ready"] else " ", t["label"],
                 "da ho tro" if t["ready"] else "sap cap nhat"))
    missing = [t["file"] for t in TEMPLATES
               if t["file"] not in [f["file"] for f in found]]
    for f in missing:
        print("  [!] THIEU FILE:", f)


if __name__ == "__main__":
    # tham số tuỳ chọn --so=400: đánh số giấy chứng nhận từ 400 tăng dần
    _start = ""
    _args = []
    for _a in sys.argv[1:]:
        if _a.startswith("--so="):
            _start = _a[5:]
        else:
            _args.append(_a)

    if len(_args) == 1 and _args[0] == "--mau":
        run_cli_list_templates()
    elif len(_args) == 4 and _args[0] == "--excel":
        run_cli_excel(_args[1], _args[2], _args[3], _start)
    elif len(_args) == 3:
        run_cli(_args[0], _args[1], _args[2], _start)
    else:
        run_gui()
