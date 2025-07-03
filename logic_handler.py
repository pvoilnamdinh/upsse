import openpyxl
import re
import io
import pandas as pd
import csv
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle

def clean_string(s):
    """Hàm hỗ trợ làm sạch chuỗi, loại bỏ khoảng trắng thừa và dấu nháy đơn ở đầu."""
    if s is None:
        return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"):
        cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def to_float(value):
    """Hàm hỗ trợ chuyển đổi một giá trị (có thể là text) sang dạng số."""
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0

def format_tax_code(raw_vat_value):
    """Chuẩn hóa giá trị VAT sang định dạng chuỗi 2 chữ số."""
    if raw_vat_value is None:
        return ""
    try:
        s_value = str(raw_vat_value).replace('%', '').strip()
        f_value = float(s_value)
        if 0 < f_value < 1:
            f_value *= 100
        return f"{round(f_value):02d}"
    except (ValueError, TypeError):
        return ""

def load_static_data(data_file_path, mahh_file_path, dskh_file_path):
    """Hàm này đọc tất cả các file cấu hình và trả về một dictionary."""
    static_data = {}
    try:
        # Đọc Data.xlsx
        wb = load_workbook(data_file_path, data_only=True)
        ws = wb.active
        chxd_list, tk_mk_map, khhd_map, chxd_to_khuvuc_map, vu_viec_map = [], {}, {}, {}, {}
        vu_viec_headers = [clean_string(cell.value) for cell in ws[2][4:9]]
        for row_values in ws.iter_rows(min_row=3, max_col=12, values_only=True):
            chxd_name = clean_string(row_values[3])
            if chxd_name:
                ma_kho, khhd, khu_vuc = clean_string(row_values[9]), clean_string(row_values[10]), clean_string(row_values[11])
                if chxd_name not in tk_mk_map: chxd_list.append(chxd_name)
                if ma_kho: tk_mk_map[chxd_name] = ma_kho
                if khhd: khhd_map[chxd_name] = khhd
                if khu_vuc: chxd_to_khuvuc_map[chxd_name] = khu_vuc
                vu_viec_map[chxd_name] = {}
                vu_viec_data_row = row_values[4:9]
                for i, header in enumerate(vu_viec_headers):
                    if header:
                        key = "Dầu mỡ nhờn" if i == len(vu_viec_headers) - 1 else header
                        vu_viec_map[chxd_name][key] = clean_string(vu_viec_data_row[i])
        if not chxd_list: return None, "Không tìm thấy Tên CHXD nào trong cột D của file Data.xlsx."
        static_data.update({"DS_CHXD": chxd_list, "tk_mk": tk_mk_map, "khhd_map": khhd_map, "chxd_to_khuvuc_map": chxd_to_khuvuc_map, "vu_viec_map": vu_viec_map})
        def get_lookup_map(min_r, max_r, min_c=1, max_c=2):
            return {clean_string(row[0]): row[1] for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c, values_only=True) if row[0] and row[1] is not None}
        phi_bvmt_map_raw = get_lookup_map(10, 13)
        static_data.update({"phi_bvmt_map": {k: to_float(v) for k, v in phi_bvmt_map_raw.items()}, "tk_no_map": get_lookup_map(29, 31), "tk_doanh_thu_map": get_lookup_map(33, 35), "tk_thue_co_map": get_lookup_map(38, 40), "tk_gia_von_value": ws['B36'].value, "tk_no_bvmt_map": get_lookup_map(44, 46), "tk_dt_thue_bvmt_map": get_lookup_map(48, 50), "tk_gia_von_bvmt_value": ws['B51'].value, "tk_thue_co_bvmt_map": get_lookup_map(53, 55)})

        # Đọc MaHH.xlsx
        wb_mahh = load_workbook(mahh_file_path, data_only=True)
        static_data["ma_hang_map"] = {clean_string(row[0]): clean_string(row[2]) for row in wb_mahh.active.iter_rows(min_row=2, max_col=3, values_only=True) if row[0] and row[2]}

        # Đọc DSKH.xlsx
        wb_dskh = load_workbook(dskh_file_path, data_only=True)
        static_data["mst_to_makh_map"] = {clean_string(row[2]): clean_string(row[3]) for row in wb_dskh.active.iter_rows(min_row=2, max_col=4, values_only=True) if row[2]}
        
        return static_data, None
    except FileNotFoundError as e: return None, f"Lỗi: Không tìm thấy file cấu hình. Chi tiết: {e.filename}"
    except Exception as e: return None, f"Lỗi khi đọc file cấu hình: {e}"

def _load_uploaded_workbook(file_content_bytes):
    """Đọc file người dùng tải lên trực tiếp như một file Excel."""
    try:
        return load_workbook(io.BytesIO(file_content_bytes), data_only=True)
    except Exception as e:
        raise ValueError(f"Lỗi khi đọc file Bảng kê hóa đơn. Hãy chắc chắn file tải lên là file Excel. Lỗi: {e}")

def _analyze_date_ambiguity(worksheet):
    """Phân tích ngày tháng trong BKHD."""
    unique_dates = set()
    for row in worksheet.iter_rows(min_row=11, values_only=True):
        if to_float(row[8] if len(row) > 8 else None) > 0:
            date_val = row[20] if len(row) > 20 else None
            converted_date_obj = None
            if isinstance(date_val, datetime): converted_date_obj = date_val
            elif isinstance(date_val, (int, float)):
                try: converted_date_obj = pd.to_datetime(date_val, unit='D', origin='1899-12-30').to_pydatetime()
                except (ValueError, TypeError): pass
            if converted_date_obj: unique_dates.add(converted_date_obj.date())
    if len(unique_dates) > 1: raise ValueError("Công cụ chỉ chạy được khi bạn kết xuất hóa đơn trong 1 ngày duy nhất.")
    if not unique_dates: raise ValueError("Không tìm thấy dữ liệu hóa đơn hợp lệ nào trong file Bảng kê.")
    the_date = unique_dates.pop()
    if the_date.day > 12: return False, datetime(the_date.year, the_date.month, the_date.day), None
    else:
        try:
            date_as_is, swapped_date = datetime(the_date.year, the_date.month, the_date.day), datetime(the_date.year, the_date.day, the_date.month)
            return (True, date_as_is, swapped_date) if date_as_is != swapped_date else (False, date_as_is, None)
        except ValueError: return False, datetime(the_date.year, the_date.month, the_date.day), None

def _validate_address_length(worksheet):
    """Kiểm tra độ dài địa chỉ trên cột E của BKHD."""
    for i, row in enumerate(worksheet.iter_rows(min_row=11, values_only=True), 11):
        if to_float(row[8] if len(row) > 8 else None) > 0 and len(str(row[4] or '')) > 128:
            raise ValueError(f"Dòng địa chỉ tại ô E{i} quá dài, hãy chỉnh sửa rồi chạy lại công cụ.")

# *** SỬA LỖI: THÊM LẠI HÀM BỊ MẤT ***
def _create_upsse_workbook():
    """Tạo một file Excel mới với cấu trúc tiêu đề cho UpSSE."""
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    wb = Workbook()
    ws = wb.active
    for _ in range(4): ws.append([''] * len(headers))
    ws.append(headers)
    return wb

def _create_bvmt_row(original_row, phi_bvmt, static_data, khu_vuc):
    """Tạo dòng thuế BVMT dựa trên dòng hóa đơn gốc."""
    bvmt_row = list(original_row)
    tk_no_bvmt, tk_dt_thue_bvmt, tk_gia_von_bvmt, tk_thue_co_bvmt = static_data.get('tk_no_bvmt_map', {}).get(khu_vuc), static_data.get('tk_dt_thue_bvmt_map', {}).get(khu_vuc), static_data.get('tk_gia_von_bvmt_value'), static_data.get('tk_thue_co_bvmt_map', {}).get(khu_vuc)
    so_luong, ma_thue = to_float(original_row[12]), original_row[17]
    thue_suat = to_float(ma_thue) / 100.0 if ma_thue else 0.0
    bvmt_row[6], bvmt_row[7], bvmt_row[13], bvmt_row[14], bvmt_row[18], bvmt_row[19], bvmt_row[20], bvmt_row[21], bvmt_row[36] = "TMT", "Thuế bảo vệ môi trường", phi_bvmt, round(phi_bvmt * so_luong), tk_no_bvmt, tk_dt_thue_bvmt, tk_gia_von_bvmt, tk_thue_co_bvmt, round(phi_bvmt * so_luong * thue_suat)
    for i in [5, 23, 31, 32, 33]: bvmt_row[i] = ''
    return bvmt_row

def _generate_upsse_data(source_rows, static_data, selected_chxd, final_date, is_new_price_period=False):
    """Hàm lõi để xử lý một danh sách các dòng hóa đơn và tạo ra dữ liệu cho file UpSSE."""
    # Lấy các bản đồ tra cứu
    ma_kho, ma_hang_map, phi_bvmt_map, vu_viec_map, mst_to_makh_map, xang_dau_group = static_data['tk_mk'].get(selected_chxd), static_data['ma_hang_map'], static_data['phi_bvmt_map'], static_data['vu_viec_map'], static_data['mst_to_makh_map'], ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]
    khu_vuc = static_data['chxd_to_khuvuc_map'].get(selected_chxd)
    tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co = static_data['tk_no_map'].get(khu_vuc), static_data['tk_doanh_thu_map'].get(khu_vuc), static_data['tk_gia_von_value'], static_data['tk_thue_co_map'].get(khu_vuc)

    original_invoice_rows, bvmt_rows, summary_data, first_invoice_prefix_source = [], [], {}, ""

    for bkhd_row in source_rows:
        if to_float(bkhd_row[8] if len(bkhd_row) > 8 else None) <= 0: continue
        ten_kh, ten_mat_hang = clean_string(bkhd_row[3]), clean_string(bkhd_row[6])
        is_anonymous, is_petrol = (ten_kh == "Người mua không lấy hóa đơn"), (ten_mat_hang in xang_dau_group)
        
        new_upsse_row = [''] * 37
        new_upsse_row[9], new_upsse_row[1], new_upsse_row[31], new_upsse_row[2] = ma_kho, ten_kh, ten_kh, final_date
        ky_hieu_shd, so_hd_goc = str(bkhd_row[18] or '').strip(), str(bkhd_row[19] or '').strip()
        so_hoa_don_moi = f"HN{so_hd_goc[-6:]}" if selected_chxd == "Nguyễn Huệ" else f"{ky_hieu_shd[-2:]}{so_hd_goc[-6:]}"
        new_upsse_row[3], new_upsse_row[4], new_upsse_row[5] = so_hoa_don_moi, clean_string(bkhd_row[17]) + clean_string(bkhd_row[18]), "Xuất bán hàng theo hóa đơn số " + so_hoa_don_moi
        new_upsse_row[7], new_upsse_row[6], new_upsse_row[8] = ten_mat_hang, ma_hang_map.get(ten_mat_hang, ''), clean_string(bkhd_row[10])
        so_luong = to_float(bkhd_row[8])
        new_upsse_row[12] = round(so_luong, 3)
        phi_bvmt = phi_bvmt_map.get(ten_mat_hang, 0.0) if is_petrol else 0.0
        new_upsse_row[13] = to_float(bkhd_row[9]) - phi_bvmt
        ma_thue, thue_suat = format_tax_code(bkhd_row[14]), to_float(format_tax_code(bkhd_row[14])) / 100.0 if format_tax_code(bkhd_row[14]) else 0.0
        new_upsse_row[17] = ma_thue
        new_upsse_row[36] = round(to_float(bkhd_row[15]) - round(phi_bvmt * so_luong * thue_suat))
        new_upsse_row[14] = round(to_float(bkhd_row[16]) - to_float(bkhd_row[15]) - round(phi_bvmt * so_luong)) if is_petrol else round(to_float(bkhd_row[13]))
        new_upsse_row[18], new_upsse_row[19], new_upsse_row[20], new_upsse_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
        chxd_vu_viec_map = vu_viec_map.get(selected_chxd, {})
        new_upsse_row[23] = chxd_vu_viec_map.get(ten_mat_hang, chxd_vu_viec_map.get("Dầu mỡ nhờn", ''))
        new_upsse_row[32], new_upsse_row[33] = clean_string(bkhd_row[4]), clean_string(bkhd_row[5])
        ma_kh_fast, mst_khach_hang = clean_string(bkhd_row[2]), clean_string(bkhd_row[5])
        ma_khach_final = ma_kho
        if ma_kh_fast and len(ma_kh_fast) < 12: ma_khach_final = ma_kh_fast
        elif mst_khach_hang:
            ma_kh_dskh = mst_to_makh_map.get(mst_khach_hang)
            if ma_kh_dskh: ma_khach_final = ma_kh_dskh
        new_upsse_row[0] = ma_khach_final

        if not is_anonymous or not is_petrol:
            original_invoice_rows.append(new_upsse_row)
            if is_petrol: bvmt_rows.append(_create_bvmt_row(new_upsse_row, phi_bvmt, static_data, khu_vuc))
        else:
            if not first_invoice_prefix_source: first_invoice_prefix_source = str(bkhd_row[18] or '').strip()
            if ten_mat_hang not in summary_data: summary_data[ten_mat_hang] = {'total_so_luong': 0, 'total_tien_hang': 0, 'total_tien_thue': 0, 'first_invoice_data': {'ky_hieu_mau_so': clean_string(bkhd_row[17]), 'ky_hieu_ky_hieu': clean_string(bkhd_row[18]), 'don_gia': to_float(bkhd_row[9]), 'vat_raw': bkhd_row[14]}}
            summary_data[ten_mat_hang]['total_so_luong'] += new_upsse_row[12]
            summary_data[ten_mat_hang]['total_tien_hang'] += new_upsse_row[14]
            summary_data[ten_mat_hang]['total_tien_thue'] += new_upsse_row[36]

    # Tạo dòng tổng hợp
    suffix_map = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
    suffix_map_new = {"Xăng E5 RON 92-II": "5", "Xăng RON 95-III": "6", "Dầu DO 0,05S-II": "7", "Dầu DO 0,001S-V": "8"}
    current_suffix_map = suffix_map_new if is_new_price_period else suffix_map
    prefix = first_invoice_prefix_source[-2:] if len(first_invoice_prefix_source) >= 2 else first_invoice_prefix_source

    for product_name, data in summary_data.items():
        summary_row = [''] * 37
        first_data = data['first_invoice_data']
        date_part, suffix = f"{final_date.day:02d}.{final_date.month:02d}", current_suffix_map.get(product_name, "")
        summary_invoice_number = f"{prefix}BK.{date_part}.{suffix}"
        phi_bvmt_unit = phi_bvmt_map.get(product_name, 0.0)
        summary_row[0], summary_row[1], summary_row[31], summary_row[2], summary_row[3], summary_row[4], summary_row[5], summary_row[7], summary_row[6], summary_row[8], summary_row[9], summary_row[12], summary_row[13], summary_row[14], summary_row[17], summary_row[18], summary_row[19], summary_row[20], summary_row[21], summary_row[23], summary_row[36] = ma_kho, f"Khách hàng mua {product_name} không lấy hóa đơn", f"Khách hàng mua {product_name} không lấy hóa đơn", final_date, summary_invoice_number, first_data['ky_hieu_mau_so'] + first_data['ky_hieu_ky_hieu'], "Xuất bán hàng theo hóa đơn số " + summary_invoice_number, product_name, ma_hang_map.get(product_name, ''), "Lít", ma_kho, round(data['total_so_luong'], 3), first_data['don_gia'] - phi_bvmt_unit, data['total_tien_hang'], format_tax_code(first_data['vat_raw']), tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co, vu_viec_map.get(selected_chxd, {}).get(product_name, ''), data['total_tien_thue']
        original_invoice_rows.append(summary_row)
        bvmt_rows.append(_create_bvmt_row(summary_row, phi_bvmt_unit, static_data, khu_vuc))

    return original_invoice_rows, bvmt_rows

def process_uploaded_file(uploaded_file_content, static_data, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None):
    """Hàm chính để xử lý file bảng kê, có hỗ trợ xác nhận ngày tháng."""
    bkhd_wb = _load_uploaded_workbook(uploaded_file_content)
    bkhd_ws = bkhd_wb.active
    
    final_date = None
    if confirmed_date_str:
        final_date = datetime.strptime(confirmed_date_str, '%Y-%m-%d')
    else:
        is_ambiguous, date1, date2 = _analyze_date_ambiguity(bkhd_ws)
        if is_ambiguous: return {'choice_needed': True, 'options': [{'text': date1.strftime('%d/%m/%Y'), 'value': date1.strftime('%Y-%m-%d')}, {'text': date2.strftime('%d/%m/%Y'), 'value': date2.strftime('%Y-%m-%d')}]}
        final_date = date1

    khhd_map = static_data.get('khhd_map', {})
    khhd_from_data = khhd_map.get(selected_chxd)
    if not khhd_from_data: raise ValueError(f"Lỗi cấu hình: Không tìm thấy 'Ký hiệu hóa đơn' cho '{selected_chxd}'.")
    khhd_suffix_expected = khhd_from_data[-6:]
    validation_value_raw = bkhd_ws['S11'].value
    if khhd_suffix_expected not in clean_string(validation_value_raw): raise ValueError(f"Bảng kê hóa đơn không khớp. Ký hiệu trên file: '{clean_string(validation_value_raw)}', mong đợi chứa: '{khhd_suffix_expected}'.")
    _validate_address_length(bkhd_ws)

    all_source_rows = list(bkhd_ws.iter_rows(min_row=11, values_only=True))

    def create_final_buffer(rows_to_process, is_new_period):
        original_rows, bvmt_rows_list = _generate_upsse_data(rows_to_process, static_data, selected_chxd, final_date, is_new_period)
        if not original_rows and not bvmt_rows_list: return None
        wb = _create_upsse_workbook()
        ws = wb.active
        for row_data in (original_rows + bvmt_rows_list): ws.append(row_data)
        for i in range(6, ws.max_row + 1): ws[f'C{i}'].number_format = 'dd/mm/yyyy'
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    if price_periods == '1':
        return create_final_buffer(all_source_rows, False)
    else:
        split_index = -1
        for i, row in enumerate(all_source_rows):
            if clean_string(row[19]) == new_price_invoice_number:
                split_index = i
                break
        if split_index == -1: raise ValueError(f"Không tìm thấy số hóa đơn '{new_price_invoice_number}' để chia giai đoạn giá.")
        
        old_price_rows, new_price_rows = all_source_rows[:split_index], all_source_rows[split_index:]
        return {'old': create_final_buffer(old_price_rows, False), 'new': create_final_buffer(new_price_rows, True)}
