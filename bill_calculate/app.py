"""
app.py — Flask Web App: Upload PDF → Đối chiếu master_data → Download kết quả
"""
import os
import uuid
import shutil
import time
from datetime import datetime

from flask import Flask, request, render_template_string, send_file, redirect, url_for

from calculator import process_all

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
MASTER_PATH = os.path.join(BASE_DIR, "master_data.xlsx")

app = Flask(__name__)
app.secret_key = "bill_calculator_secret_key_2024"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# HTML TEMPLATES
# ============================================================

UPLOAD_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bill Calculator — Tách SKU Tự Động</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #f0f4f8;
            color: #1e293b;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
            padding: 40px;
            max-width: 600px;
            width: 90%;
        }
        h1 {
            font-size: 24px;
            color: #2563eb;
            margin-bottom: 4px;
        }
        .subtitle {
            color: #64748b;
            font-size: 14px;
            margin-bottom: 28px;
        }
        .upload-zone {
            border: 2px dashed #cbd5e1;
            border-radius: 12px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: border-color 0.2s, background 0.2s;
            margin-bottom: 20px;
        }
        .upload-zone:hover, .upload-zone.dragover {
            border-color: #2563eb;
            background: #f8fafc;
        }
        .upload-zone svg {
            display: block;
            margin: 0 auto 12px;
        }
        .upload-zone p {
            color: #64748b;
            font-size: 15px;
        }
        .upload-zone .browse {
            color: #2563eb;
            font-weight: 600;
            text-decoration: underline;
        }
        input[type="file"] { display: none; }
        .file-list {
            list-style: none;
            margin-bottom: 20px;
        }
        .file-list li {
            background: #f1f5f9;
            padding: 8px 14px;
            border-radius: 8px;
            margin-bottom: 6px;
            font-size: 13px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .file-list .remove {
            margin-left: auto;
            color: #ef4444;
            cursor: pointer;
            font-weight: bold;
            font-size: 16px;
        }
        .btn {
            display: inline-block;
            padding: 12px 28px;
            font-size: 15px;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            transition: background 0.2s;
            text-decoration: none;
        }
        .btn-primary {
            background: #2563eb;
            color: #fff;
            width: 100%;
        }
        .btn-primary:hover { background: #1d4ed8; }
        .btn-primary:disabled { background: #93c5fd; cursor: not-allowed; }
        .flash-msg {
            padding: 12px 16px;
            border-radius: 8px;
            margin-bottom: 16px;
            font-size: 14px;
        }
        .flash-error { background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
        .flash-success { background: #f0fdf4; color: #059669; border: 1px solid #bbf7d0; }
        .info-box {
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 8px;
            padding: 14px 18px;
            margin-bottom: 20px;
            font-size: 13px;
            color: #3b82f6;
            line-height: 1.6;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📦 Bill Calculator</h1>
        <p class="subtitle">Upload PDF / Excel danh sách sản phẩm — tự động tách SKU ghép & tính số lượng thực</p>

        {% if msg %}
        <div class="flash-msg {{ msg_type }}">{{ msg }}</div>
        {% endif %}

        <div class="info-box">
            <strong>Hỗ trợ:</strong> File PDF Picking List từ Haravan<br>
            <strong>Kết quả:</strong> File Excel đối chiếu master_data + PDF + TXT<br>
            <strong>Yêu cầu:</strong> File <code>master_data.xlsx</code> phải có trong thư mục gốc
        </div>

        <form method="POST" action="/upload" enctype="multipart/form-data" id="upload-form">
            <div class="upload-zone" id="drop-zone">
                <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="1.5">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17 8 12 3 7 8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                </svg>
                <p>Kéo thả file PDF / Excel vào đây, hoặc <span class="browse">chọn file</span></p>
            </div>
            <input type="file" name="pdf_files" id="file-input" accept=".pdf" multiple>
            <ul class="file-list" id="file-list"></ul>
            <button type="submit" class="btn btn-primary" id="submit-btn" disabled>
                ⚡ Xử lý ngay
            </button>
        </form>
    </div>

    <script>
        const input = document.getElementById('file-input');
        const dropZone = document.getElementById('drop-zone');
        const fileList = document.getElementById('file-list');
        const submitBtn = document.getElementById('submit-btn');

        dropZone.addEventListener('click', () => input.click());
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            input.files = e.dataTransfer.files;
            updateList();
        });
        input.addEventListener('change', updateList);

        function updateList() {
            fileList.innerHTML = '';
            if (input.files.length > 0) {
                for (const f of input.files) {
                    const isPdf = f.name.toLowerCase().endsWith('.pdf');
                    const icon = isPdf ? '📕' : '📊';
                    const li = document.createElement('li');
                    li.innerHTML = `${icon} ${f.name} <span class="remove" data-name="${f.name}">&times;</span>`;
                    fileList.appendChild(li);
                }
                submitBtn.disabled = false;
                submitBtn.textContent = `⚡ Xử lý ${input.files.length} file PDF`;
            } else {
                submitBtn.disabled = true;
                submitBtn.textContent = '⚡ Xử lý ngay';
            }
        }
    </script>
</body>
</html>"""

RESULT_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kết quả — Bill Calculator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #f0f4f8;
            color: #1e293b;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
            padding: 40px;
            max-width: 700px;
            width: 90%;
        }
        h1 { font-size: 24px; color: #059669; margin-bottom: 4px; }
        .subtitle { color: #64748b; font-size: 14px; margin-bottom: 28px; }
        .stats {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 12px;
            margin-bottom: 28px;
        }
        .stat-box {
            background: #f8fafc;
            border-radius: 10px;
            padding: 16px;
            text-align: center;
        }
        .stat-value { font-size: 22px; font-weight: 700; color: #2563eb; }
        .stat-label { font-size: 12px; color: #64748b; margin-top: 4px; }
        h2 { font-size: 16px; margin-bottom: 12px; color: #334155; }
        .download-list { list-style: none; margin-bottom: 28px; }
        .download-list li { margin-bottom: 8px; }
        .download-link {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 12px 16px;
            background: #f1f5f9;
            border-radius: 10px;
            text-decoration: none;
            color: #1e293b;
            font-weight: 500;
            transition: background 0.2s;
        }
        .download-link:hover { background: #e2e8f0; }
        .download-link .icon { font-size: 20px; }
        .download-link .ext {
            margin-left: auto;
            font-size: 11px;
            background: #cbd5e1;
            padding: 3px 8px;
            border-radius: 4px;
            color: #475569;
        }
        .btn {
            display: inline-block;
            padding: 12px 28px;
            font-size: 15px;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            background: #2563eb;
            color: #fff;
            text-decoration: none;
            transition: background 0.2s;
        }
        .btn:hover { background: #1d4ed8; }
    </style>
</head>
<body>
    <div class="container">
        <h1>✅ Xử lý thành công!</h1>
        <p class="subtitle">{{ info.sl_goc }} — đối chiếu với master_data.xlsx</p>

        <div class="stats">
            <div class="stat-box">
                <div class="stat-value">{{ info.rows }}</div>
                <div class="stat-label">Dòng kết quả</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">{{ info.tong_cu }}</div>
                <div class="stat-label">Tổng Qty</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">{{ info.tong_moi }}</div>
                <div class="stat-label">Tổng Qty Sold</div>
            </div>
        </div>

        <h2>📥 Tải file kết quả ({{ info.pdf_count }} PDF)</h2>
        <ul class="download-list">
            {% for f in files %}
            <li>
                <a href="/download/{{ sid }}/{{ f.name }}" class="download-link">
                    <span class="icon">{{ f.icon }}</span>
                    {{ f.label }}
                    <span class="ext">.{{ f.ext }}</span>
                </a>
            </li>
            {% endfor %}
        </ul>

        <a href="/" class="btn">⬅ Xử lý thêm file</a>
    </div>
</body>
</html>"""


# ============================================================
# AUTO CLEANUP
# ============================================================

def cleanup_old_dirs():
    """Xóa thư mục uploads/outputs cũ hơn 1 giờ."""
    now = time.time()
    for d in [UPLOAD_DIR, OUTPUT_DIR]:
        for name in os.listdir(d):
            path = os.path.join(d, name)
            if os.path.isdir(path) and (now - os.path.getmtime(path)) > 3600:
                try:
                    shutil.rmtree(path)
                except Exception:
                    pass


# ============================================================
# ROUTES
# ============================================================

@app.route("/")
def index():
    msg = request.args.get("msg", "")
    msg_type = request.args.get("type", "")
    return render_template_string(UPLOAD_PAGE, msg=msg, msg_type=msg_type)


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("pdf_files")
    pdf_files = [f for f in files if f.filename and f.filename.lower().endswith(".pdf")]

    if not pdf_files:
        return redirect(url_for("index", msg="Vui lòng chọn ít nhất 1 file PDF!", type="flash-error"))

    # Kiểm tra master_data tồn tại
    if not os.path.exists(MASTER_PATH):
        return redirect(url_for("index",
                        msg="Không tìm thấy master_data.xlsx trong thư mục gốc!",
                        type="flash-error"))

    # Cleanup cũ
    cleanup_old_dirs()

    # Tạo thư mục session
    sid = uuid.uuid4().hex[:12]
    upload_dir = os.path.join(UPLOAD_DIR, sid)
    output_dir = os.path.join(OUTPUT_DIR, sid)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # Lưu file upload
    saved_paths = []
    for f in pdf_files:
        safe_name = f.filename.replace("\\", "/").split("/")[-1]
        path = os.path.join(upload_dir, safe_name)
        f.save(path)
        saved_paths.append(path)

    try:
        all_results = process_all(saved_paths, output_dir, MASTER_PATH)
    except ValueError as e:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        return redirect(url_for("index", msg=str(e), type="flash-error"))
    except Exception as e:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        return redirect(url_for("index", msg=f"Lỗi xử lý: {e}", type="flash-error"))

    # Dọn file upload (không cần nữa)
    shutil.rmtree(upload_dir, ignore_errors=True)

    # Tổng hợp stats từ tất cả PDF
    total_rows = sum(r["rows"] for r in all_results)
    total_qty = sum(r["tong_qty"] for r in all_results)
    total_sold = sum(r["tong_sold"] for r in all_results)

    info = {
        "rows": total_rows,
        "tong_cu": total_qty,
        "tong_moi": total_sold,
        "sl_goc": f"Đã xử lý {len(pdf_files)} file PDF",
        "pdf_count": len(all_results),
    }

    # Gom tất cả file download từ các PDF
    download_files = []
    for r in all_results:
        base = r["base_name"]
        for fkey, label, icon, ext in [
            ("excel", f"Excel kết quả [{base}]", "📊", "xlsx"),
            ("pdf", f"PDF danh sách [{base}]", "📕", "pdf"),
        ]:
            fpath = r["files"].get(fkey)
            if fpath:
                download_files.append({
                    "name": os.path.basename(fpath),
                    "label": label,
                    "icon": icon,
                    "ext": ext,
                })

    return render_template_string(RESULT_PAGE, info=info, files=download_files, sid=sid)


@app.route("/download/<sid>/<filename>")
def download(sid: str, filename: str):
    # Security: only allow safe filenames
    safe_name = filename.replace("\\", "/").split("/")[-1]
    path = os.path.join(OUTPUT_DIR, sid, safe_name)
    if not os.path.exists(path):
        return "File không tồn tại hoặc đã hết hạn.", 404
    return send_file(path, as_attachment=True, download_name=safe_name)


# ============================================================
# REST API — Cho script & extension gọi tự động
# ============================================================

@app.route("/api/status")
def api_status():
    """Health check cho script và extension."""
    master_ok = os.path.exists(MASTER_PATH)
    return {
        "status": "ok" if master_ok else "no_master_data",
        "master_data": master_ok,
        "version": "1.0.0",
    }


@app.route("/api/process", methods=["POST"])
def api_process():
    """
    Nhận file PDF qua multipart/form-data, xử lý, trả về JSON kết quả.

    Request:
        POST /api/process
        Content-Type: multipart/form-data
        Body: file=@danh_sach.pdf

    Response:
        {
            "ok": true,
            "base_name": "tiktok_20260622_0830",
            "rows": 45,
            "tong_qty": 120,
            "tong_sold": 100,
            "files": {
                "excel": "kết quả tiktok_20260622_0830.xlsx",
                "pdf": "Danh_sach_tiktok_20260622_0830.pdf"
            },
            "downloads": {
                "excel": "/download/{sid}/{filename}",
                "pdf": "/download/{sid}/{filename}"
            }
        }
    """
    from flask import jsonify

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Thiếu field 'file' chứa PDF."}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "File phải là PDF."}), 400

    if not os.path.exists(MASTER_PATH):
        return jsonify({
            "ok": False,
            "error": "Không tìm thấy master_data.xlsx. Vui lòng kiểm tra thư mục bill_calculate."
        }), 500

    # Cleanup cũ
    cleanup_old_dirs()

    # Tạo session
    sid = uuid.uuid4().hex[:12]
    upload_dir = os.path.join(UPLOAD_DIR, sid)
    output_dir = os.path.join(OUTPUT_DIR, sid)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # Lưu file upload
    safe_name = f.filename.replace("\\", "/").split("/")[-1]
    saved_path = os.path.join(upload_dir, safe_name)
    f.save(saved_path)

    try:
        results = process_all([saved_path], output_dir, MASTER_PATH)
    except ValueError as e:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        return jsonify({"ok": False, "error": str(e)}), 422
    except Exception as e:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        return jsonify({"ok": False, "error": f"Lỗi xử lý: {e}"}), 500

    # Dọn upload
    shutil.rmtree(upload_dir, ignore_errors=True)

    r = results[0]
    base_name = r["base_name"]
    excel_filename = os.path.basename(r["files"]["excel"])
    pdf_filename = os.path.basename(r["files"]["pdf"])

    return jsonify({
        "ok": True,
        "base_name": base_name,
        "rows": r["rows"],
        "tong_qty": r["tong_qty"],
        "tong_sold": r["tong_sold"],
        "tong_promo": r["tong_promo"],
        "files": {
            "excel": excel_filename,
            "pdf": pdf_filename,
        },
        "downloads": {
            "excel": f"/download/{sid}/{excel_filename}",
            "pdf": f"/download/{sid}/{pdf_filename}",
        },
    })


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print("=" * 50)
    print("  📦 Bill Calculator — Tách SKU Tự Động")
    print("  Mở trình duyệt: http://localhost:5000")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000)
