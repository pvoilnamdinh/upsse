from flask import Flask, request, render_template, flash, redirect, url_for, send_file
import io
import base64
import zipfile

# Import các hàm cần thiết từ logic_handler
from logic_handler import load_static_data, process_uploaded_file

# --- Cài đặt Flask App cơ bản ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_super_secret_key_12345'
DATA_FILE_PATH = "Data.xlsx" 
MAHH_FILE_PATH = "MaHH.xlsx"
DSKH_FILE_PATH = "DSKH.xlsx"

# --- Route chính để hiển thị trang upload ---
@app.route('/', methods=['GET'])
def index():
    """Hiển thị trang upload chính."""
    static_data, error_message = load_static_data(DATA_FILE_PATH, MAHH_FILE_PATH, DSKH_FILE_PATH)
    chxd_list = static_data.get("DS_CHXD", []) if static_data else []
    if error_message:
        flash(error_message, "danger")

    return render_template('index.html', chxd_list=chxd_list, form_data={})

# --- Route để xử lý file ---
@app.route('/process', methods=['POST'])
def process():
    """Xử lý file, hỗ trợ 1 hoặc 2 giai đoạn giá."""
    try:
        static_data, error = load_static_data(DATA_FILE_PATH, MAHH_FILE_PATH, DSKH_FILE_PATH)
        if error:
            raise ValueError(error)
        chxd_list = static_data.get("DS_CHXD", [])

        # --- Lấy dữ liệu từ form ---
        form_data = {
            "selected_chxd": request.form.get('chxd'),
            "price_periods": request.form.get('price_periods'),
            "invoice_number": request.form.get('invoice_number', '').strip(),
            "confirmed_date": request.form.get('confirmed_date'),
            "encoded_file": request.form.get('file_content_b64')
        }

        # Kiểm tra các trường bắt buộc
        if not form_data["selected_chxd"]:
            flash('Vui lòng chọn CHXD.', 'warning')
            return redirect(url_for('index'))
        if form_data["price_periods"] == '2' and not form_data["invoice_number"]:
            flash('Vui lòng nhập "Số hóa đơn đầu tiên của giá mới" khi chọn 2 giai đoạn giá.', 'warning')
            return redirect(url_for('index'))

        # Xác định nội dung file (từ file upload hoặc từ trường ẩn)
        file_content = None
        if form_data["encoded_file"]:
            file_content = base64.b64decode(form_data["encoded_file"])
        elif 'file' in request.files and request.files['file'].filename != '':
            file_content = request.files['file'].read()
        else:
            flash('Vui lòng tải lên file Bảng kê hóa đơn.', 'warning')
            return redirect(url_for('index'))

        # Gọi hàm xử lý chính từ logic_handler
        result = process_uploaded_file(
            uploaded_file_content=file_content, 
            static_data=static_data, 
            selected_chxd=form_data["selected_chxd"],
            price_periods=form_data["price_periods"],
            new_price_invoice_number=form_data["invoice_number"],
            confirmed_date_str=form_data["confirmed_date"]
        )
        
        # --- Xử lý kết quả trả về ---
        if isinstance(result, dict) and result.get('choice_needed'):
            # Trường hợp cần người dùng xác nhận ngày
            form_data["encoded_file"] = base64.b64encode(file_content).decode('utf-8')
            return render_template('index.html', chxd_list=chxd_list, date_ambiguous=True, date_options=result['options'], form_data=form_data)
        
        elif isinstance(result, dict) and 'old' in result:
            # Trường hợp 2 giai đoạn giá, tạo file ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if result.get('old'):
                    zipf.writestr('UpSSE_gia_cu.xlsx', result['old'].getvalue())
                if result.get('new'):
                    zipf.writestr('UpSSE_gia_moi.xlsx', result['new'].getvalue())
            
            zip_buffer.seek(0)
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='UpSSE_2_giai_doan.zip',
                mimetype='application/zip'
            )

        elif isinstance(result, io.BytesIO):
            # Trường hợp 1 giai đoạn giá, trả về file Excel
            return send_file(
                result,
                as_attachment=True,
                download_name='UpSSE.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            raise ValueError("Hàm xử lý không trả về kết quả hợp lệ.")

    except ValueError as ve:
        flash(str(ve), 'danger')
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn: {e}", 'danger')

    return redirect(url_for('index'))

# --- Chạy App ---
if __name__ == '__main__':
    app.run(debug=True)
