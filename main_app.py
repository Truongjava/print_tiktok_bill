"""
main_app.py — Desktop App TikTok Seller Automation + Bill Calculate
====================================================================
Chạy: python main_app.py
Đóng gói .exe: pyinstaller --onefile --windowed --name "TTS_Bill" main_app.py
"""
import os, sys, json, time, shutil, threading, queue
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from playwright.sync_api import sync_playwright

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'bill_calculate'))
try:
    from calculator import process_all
except ImportError:
    process_all = None  # Calculator không khả dụng (vd: khi chạy file .exe)

# ── Config ──
import os as _os
BASE_DIR = Path(__file__).parent
BILL_DIR = BASE_DIR / 'bill_calculate'
UPLOAD_DIR = BILL_DIR / 'uploads'

if getattr(sys, 'frozen', False):
    # Chạy file .exe — dùng Playwright browser đã cài sẵn trên máy
    _os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH',
        _os.path.join(_os.path.expanduser('~'), 'AppData', 'Local', 'ms-playwright'))
    BASE_DIR = Path(sys.executable).parent
    BILL_DIR = BASE_DIR / 'bill_calculate'
    UPLOAD_DIR = BILL_DIR / 'uploads'
    DEFAULT_COOKIE = Path(sys._MEIPASS) / 'seller-vn.tiktok.com_23-06-2026.json'
    MASTER_DEFAULT = Path(sys._MEIPASS) / 'mã combosss.xlsx'
    RETAIL_DEFAULT = Path(sys._MEIPASS) / 'sản phẩm bán lẻ.xlsx'
else:
    # Chạy source code
    _os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH',
        _os.path.join(_os.path.expanduser('~'), 'AppData', 'Local', 'ms-playwright'))
    DEFAULT_COOKIE = BASE_DIR / 'seller-vn.tiktok.com_23-06-2026.json'
    MASTER_DEFAULT = BASE_DIR / 'mã combosss.xlsx'
    RETAIL_DEFAULT = BASE_DIR / 'sản phẩm bán lẻ.xlsx'

TARGET_URL = 'https://seller-vn.tiktok.com'
ORDERS_URL = 'https://seller-vn.tiktok.com/order?order_status%5B%5D=1&selected_sort=11&tab=to_ship&page_size=50'
BATCH_SIZE = 50
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ============================================================
# AUTOMATION
# ============================================================
def run_automation(cookie_path, doc_type, output_dir, max_orders, log_cb, state_cb, stop_event=None,
                   existing_playwright=None, existing_browser=None):
    with open(cookie_path, 'r', encoding='utf-8') as f:
        cd = json.load(f)
    cookies_list = cd.get('cookies', cd if isinstance(cd, list) else [])
    pdf_files = []
    total_printed = 0
    target = max_orders if max_orders > 0 else 10**9
    batch_num = 0
    if stop_event is None:
        stop_event = threading.Event()  # fallback

    # ── Tái sử dụng browser nếu có, nếu không thì tạo mới ──
    if existing_browser and existing_playwright:
        playwright = existing_playwright
        browser = existing_browser
        # Dùng context hiện có (đã đăng nhập sẵn)
        context = browser.contexts[0] if browser.contexts else browser.new_context(
            viewport={'width': 1366, 'height': 768},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
            accept_downloads=True)
        # Tìm page có sẵn để F5 thay vì mở tab mới
        existing_pages = [p for p in context.pages if not p.is_closed()]
        if existing_pages:
            page = existing_pages[-1]  # dùng page cuối cùng (đang ở TikTok)
            log_cb('♻ Dùng lại browser — F5 trang đơn hàng...', 'info')
        else:
            page = context.new_page()
            log_cb('♻ Dùng lại browser — mở tab mới...', 'info')
        page.goto(ORDERS_URL, wait_until='networkidle', timeout=60000)
        page.wait_for_timeout(4000)
    else:
        playwright = sync_playwright().start()
        browser = playwright.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled'])
        context = browser.new_context(viewport={'width': 1366, 'height': 768},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
            accept_downloads=True)
        pw_cookies = []
        for c in cookies_list:
            if not c.get('name') or not c.get('value'): continue
            pw = {'name': c['name'], 'value': c['value'], 'domain': c.get('domain', '.tiktok.com'), 'path': c.get('path', '/')}
            if c.get('expirationDate') and not c.get('session'): pw['expires'] = int(c['expirationDate'])
            if 'httpOnly' in c: pw['httpOnly'] = c['httpOnly']
            if 'secure' in c: pw['secure'] = c['secure']
            if c.get('sameSite'):
                pw['sameSite'] = {'strict':'Strict','lax':'Lax','no_restriction':'None','unspecified':'Lax'}.get(c['sameSite'],'Lax')
            pw_cookies.append(pw)
        context.add_cookies(pw_cookies)
        page = context.new_page()
        page.goto(TARGET_URL, wait_until='domcontentloaded', timeout=30000)
        page.wait_for_timeout(2000)

    try:
        while total_printed < target:
            if stop_event.is_set():
                log_cb('⏹ Đã dừng theo yêu cầu.', 'warn'); break
            batch_num += 1
            batch_target = min(BATCH_SIZE, target - total_printed)
            log_cb(f'▸ Batch {batch_num}: chọn {batch_target} đơn (đã in {total_printed}/{target})', 'batch')

            # Vào trang đơn hàng
            state_cb('navigating', f'Batch {batch_num}: Đang tải danh sách đơn...')
            page.goto(ORDERS_URL, wait_until='networkidle', timeout=60000)
            page.wait_for_timeout(4000)

            # Đếm & chọn
            total_avail = page.evaluate("() => document.querySelectorAll('td.col-checkbox label.p-checkbox').length")
            if total_avail == 0:
                log_cb('  ✓ Hết đơn khả dụng — hoàn thành.', 'ok'); break
            to_select = min(batch_target, total_avail)

            cbs = page.query_selector_all('td.col-checkbox label.p-checkbox')
            checked = 0
            for cb in cbs[:to_select]:
                try: cb.click(); checked += 1; page.wait_for_timeout(120)
                except: pass
            page.wait_for_timeout(800)

            if total_avail < batch_target:
                log_cb(f'  ⚠ Chỉ có {total_avail} đơn, chọn hết {checked}', 'warn')
            else:
                log_cb(f'  ✓ Đã chọn {checked}/{batch_target} đơn', 'ok')
            if checked == 0: log_cb('  ✓ Hết đơn — hoàn thành.', 'ok'); break

            # ── Lưu ID các đơn hàng đã chọn vào Excel (phòng trường hợp tải lỗi) ──
            try:
                order_ids = page.evaluate('''() => {
                    const ids = [];
                    const rows = document.querySelectorAll('tr');
                    rows.forEach(row => {
                        const cb = row.querySelector('td.col-checkbox input[type="checkbox"]:checked');
                        if (cb) {
                            // Tìm ô chứa Order ID (thường là chuỗi số 15+ ký tự)
                            const cells = row.querySelectorAll('td');
                            cells.forEach(cell => {
                                const text = cell.textContent.trim();
                                if (/^\\d{15,}$/.test(text)) {
                                    ids.push(text);
                                }
                            });
                        }
                    });
                    return ids;
                }''')
                if order_ids:
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, Alignment, Border, Side
                    order_file = str(Path(output_dir) / f'ID_don_hang_da_in_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
                    wb = Workbook(); ws = wb.active; ws.title = "Order IDs"
                    ws.cell(row=1, column=1, value="STT").font = Font(bold=True)
                    ws.cell(row=1, column=2, value="Order ID").font = Font(bold=True)
                    ws.cell(row=1, column=3, value="Batch").font = Font(bold=True)
                    for i, oid in enumerate(order_ids, 1):
                        ws.cell(row=i+1, column=1, value=i)
                        ws.cell(row=i+1, column=2, value=oid)
                        ws.cell(row=i+1, column=3, value=batch_num)
                    ws.column_dimensions['A'].width = 8
                    ws.column_dimensions['B'].width = 22
                    ws.column_dimensions['C'].width = 10
                    wb.save(order_file)
                    log_cb(f'  📋 Đã lưu {len(order_ids)} ID đơn hàng vào {Path(order_file).name}', 'ok')
            except Exception as e:
                log_cb(f'  ⚠ Không lưu được ID đơn hàng: {e}', 'warn')

            # Click nút "Sắp xếp vận chuyển và in"
            state_cb('printing', f'Batch {batch_num}: Đang in...')
            ship_btn = None
            for btn_text in ['Sắp xếp vận chuyển và in', 'Arrange shipment and print', 'Sắp xếp vận chuyển']:
                try:
                    btn = page.locator(f'button:has-text("{btn_text}")').first
                    if btn.count() > 0 and btn.is_visible(timeout=3000):
                        ship_btn = btn; break
                except: pass
            if not ship_btn:
                # Fallback: tìm button chứa text "vận chuyển" và "in"
                all_btns = page.locator('button').all()
                for b in all_btns:
                    try:
                        txt = b.inner_text().strip().lower()
                        if 'vận chuyển' in txt and 'in' in txt:
                            ship_btn = b; break
                    except: pass
            if not ship_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Sắp xếp vận chuyển và in"!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_ship_btn_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break
            ship_btn.click()
            log_cb('  ✓ Đã bấm "Sắp xếp vận chuyển và in"', 'ok')
            page.wait_for_timeout(3000)

            # ── Bước 1: Bấm "Tiếp theo" trong popup "Tính năng mới In nhãn" ──
            state_cb('printing', f'Batch {batch_num}: Đợi popup "Tiếp theo"...')
            tieptheo_btn = None
            for _ in range(20):
                # Tìm nút "Tiếp theo" (có thể có dấu →)
                for selector in ['button:has-text("Tiếp theo")', 'button:has-text("Next")']:
                    try:
                        btns = page.locator(selector)
                        cnt = btns.count()
                        for j in range(cnt):
                            b = btns.nth(j)
                            if b.is_visible(timeout=500):
                                tieptheo_btn = b; break
                    except: pass
                    if tieptheo_btn: break
                if tieptheo_btn: break
                page.wait_for_timeout(1000)
            if not tieptheo_btn:
                log_cb('  ⚠ Không tìm thấy nút "Tiếp theo" — thử tiếp tục...', 'warn')
            else:
                tieptheo_btn.click(timeout=5000)
                log_cb('  ✓ Đã bấm "Tiếp theo"', 'ok')
                page.wait_for_timeout(3000)

            # ── Bước 2: Tick chọn Danh sách đóng gói + Danh sách lấy hàng ──
            state_cb('printing', f'Batch {batch_num}: Chọn loại chứng từ...')
            # Đợi popup hiện ra với các checkbox
            page.wait_for_timeout(2000)
            for doc_label in ['Danh sách đóng gói', 'Danh sách lấy hàng']:
                try:
                    lbl = page.locator('label').filter(has_text=doc_label).first
                    if lbl.count() > 0:
                        inp = lbl.locator('input')
                        if inp.count() > 0 and not inp.is_checked():
                            lbl.click(); page.wait_for_timeout(300)
                            log_cb(f'  ✓ Đã tick: {doc_label}', 'ok')
                except: pass
            page.wait_for_timeout(1000)

            # ── Bấm nút "In nhãn ngay sau khi vận chuyển" ──
            in_btn = None
            for btn_text in ['In nhãn ngay sau khi vận chuyển', 'In nhãn ngay', 'Print label immediately']:
                try:
                    btn = page.locator(f'button:has-text("{btn_text}")').first
                    if btn.count() > 0 and btn.is_visible(timeout=2000):
                        in_btn = btn; break
                except: pass
            if not in_btn:
                # Fallback: tìm button có text "In nhãn"
                for b in page.locator('button').all():
                    try:
                        txt = b.inner_text().strip().lower()
                        if 'in nhãn' in txt or 'in nhan' in txt:
                            in_btn = b; break
                    except: pass
            if not in_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "In nhãn ngay" — thử tiếp tục...', 'warn')
            else:
                in_btn.click(timeout=5000)
                log_cb('  ✓ Đã bấm "In nhãn ngay"', 'ok')
                page.wait_for_timeout(3000)

            # ── Bước 3: Popup "Tải xuống tất cả" ──
            state_cb('downloading', f'Batch {batch_num}: Đợi popup "Tải xuống tất cả"...')

            taixuong_btn = None
            for _ in range(30):
                btns = page.locator('button:has-text("Tải xuống tất cả"), button:has-text("Download all"), button:has-text("Tải xuống")')
                cnt = btns.count()
                for j in range(cnt):
                    b = btns.nth(j)
                    try:
                        if b.is_visible(timeout=500):
                            taixuong_btn = b; break
                    except: pass
                if taixuong_btn: break
                page.wait_for_timeout(1000)

            if not taixuong_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Tải xuống tất cả tập tin" — kiểm tra popup!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_taixuong_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break

            log_cb('  📥 Đang bấm nút tải xuống...', 'info')
            # Dùng event listener để bắt TẤT CẢ download (nhiều file 1 lúc)
            downloaded_files = []
            def on_download(dl):
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                suggested = dl.suggested_filename or f'PDF_goc_TTS_batch{batch_num}_{len(downloaded_files)}_{ts}.pdf'
                bp = str(Path(output_dir) / suggested)
                dl.save_as(bp)
                downloaded_files.append(bp)
                log_cb(f'  💾 Đã tải: {Path(bp).name}', 'ok')
            page.on('download', on_download)
            taixuong_btn.click(timeout=5000)
            log_cb('  ✓ Đã bấm "Tải xuống tất cả"', 'ok')
            # Đợi đủ lâu để tất cả download hoàn tất
            page.wait_for_timeout(8000)
            page.remove_listener('download', on_download)
            pdf_files.extend(downloaded_files)
            log_cb(f'  📥 Tổng cộng {len(downloaded_files)} file đã tải', 'info')
            if not downloaded_files:
                log_cb('  ✗ Không bắt được download nào!', 'err')
                total_printed += checked; break

            total_printed += checked
            log_cb(f'  📊 Tiến độ: {total_printed}/{target} đơn, {len(pdf_files)} file PDF', 'info')

            if total_printed < target and total_avail > 0:
                page.wait_for_timeout(2000)

        # KHÔNG đóng browser — giữ lại cho lần chạy sau
        return pdf_files, playwright, browser
    except Exception as e:
        log_cb(f'  ✗ Lỗi: {e}', 'err')
        # KHÔNG đóng browser — giữ lại để xem lỗi
        return pdf_files, playwright, browser

# ============================================================
# CALCULATOR
# ============================================================
def run_calculator(pdf_paths, output_dir, master_path, retail_path, log_cb):
    if process_all is None:
        log_cb('✗ Calculator không khả dụng (thiếu module calculator)', 'err'); return []
    if not Path(master_path).exists(): log_cb(f'✗ Không tìm thấy master_data: {master_path}', 'err'); return []
    out_dir = str(output_dir)
    for p in pdf_paths:
        shutil.copy2(p, str(UPLOAD_DIR / Path(p).name))
    try:
        results = process_all(pdf_paths, out_dir, master_path, retail_path)
        for r in results:
            log_cb(f'  ✓ {r["rows"]} dòng | Qty={r["tong_qty"]} | Sold={r["tong_sold"]} | Promo={r["tong_promo"]}', 'ok')
            for key, fb in r['files'].items():
                src, dst = Path(fb), Path(out_dir) / Path(fb).name
                if src != dst and src.exists(): shutil.copy2(str(src), str(dst)); r['files'][key] = str(dst)
        return results
    except Exception as e: log_cb(f'  ✗ Lỗi: {e}', 'err'); return []

# ============================================================
# SCHEDULER
# ============================================================
class Scheduler:
    def __init__(self, callback, log_cb, state_cb):
        self.callback = callback; self.log = log_cb; self.set_state = state_cb
        self.running = False; self.mode = 'once'; self.interval_hours = 1
        self.daily_times = []; self.next_run = None; self.last_run = None
        self._stop = threading.Event(); self._thread = None

    def configure(self, mode, interval_hours=1, daily_times=None):
        self.mode = mode; self.interval_hours = interval_hours; self.daily_times = daily_times or []

    def start(self):
        if self._thread and self._thread.is_alive(): return
        self.running = True; self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True); self._thread.start()

    def stop(self):
        self.running = False; self._stop.set(); self.set_state('idle', '⏹ Đã dừng')

    def _loop(self):
        while self.running and not self._stop.is_set():
            now = datetime.now()
            if self.mode == 'once': self._run_job(); self.running = False; break
            elif self.mode == 'interval':
                self.next_run = (self.last_run or now) + timedelta(hours=self.interval_hours)
                if now >= self.next_run: self._run_job()
                else:
                    w = (self.next_run - now).total_seconds()
                    self._countdown(w)
                    for _ in range(int(min(w, 60))):
                        if self._stop.is_set(): return
                        time.sleep(1)
            elif self.mode == 'daily':
                runs = []
                for t in self.daily_times:
                    try:
                        h, m = t.strip().split(':'); rt = now.replace(hour=int(h), minute=int(m), second=0, microsecond=0)
                        if rt <= now: rt += timedelta(days=1)
                        runs.append(rt)
                    except: pass
                if runs:
                    self.next_run = min(runs); w = (self.next_run - now).total_seconds()
                    if w <= 1: self._run_job()
                    else:
                        self._countdown(w)
                        for _ in range(int(min(w, 60))):
                            if self._stop.is_set(): return; time.sleep(1)
                else: time.sleep(60)
            time.sleep(1)

    def _run_job(self):
        self.set_state('running', '🔄 Đang chạy...'); start = datetime.now()
        try: self.callback()
        except Exception as e: self.log(f'✗ Job lỗi: {e}', 'err')
        self.last_run = datetime.now()
        self.log(f'🏁 Hoàn thành ({int((datetime.now()-start).total_seconds())}s)', 'ok')

    def _countdown(self, seconds):
        h, m = int(seconds//3600), int((seconds%3600)//60)
        self.set_state('waiting', f'⏳ Chạy tiếp sau {h}h{m:02d}')

# ============================================================
# AUTOMATION WORKER — dedicated thread cho Playwright
# ============================================================
class AutomationWorker:
    """Chạy tất cả automation trên 1 thread duy nhất để tránh lỗi thread-safety của Playwright.
    Browser được tạo 1 lần và tái sử dụng cho mọi job về sau."""

    def __init__(self):
        self._queue = queue.Queue()
        self._playwright = None
        self._browser = None
        self._thread = None
        self._running = False
        self._job_stop = threading.Event()

    def start(self):
        """Khởi động worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def run_job(self, job_fn):
        """Gửi 1 job vào queue. job_fn(pw, browser, stop_event) -> result được gọi trên worker thread."""
        self._queue.put((job_fn,))

    def stop_job(self):
        """Ra hiệu dừng job hiện tại (không đóng browser)."""
        self._job_stop.set()

    def shutdown(self):
        """Dừng worker thread và đóng browser."""
        self._running = False
        self._queue.put(None)  # sentinel để thoát _loop

    def _loop(self):
        """Vòng lặp chính của worker thread."""
        while self._running:
            item = self._queue.get()
            if item is None:
                break
            (job_fn,) = item
            self._job_stop.clear()
            try:
                result = job_fn(self._playwright, self._browser, self._job_stop)
                # Lưu browser/playwright nếu job vừa tạo mới
                if isinstance(result, dict):
                    if 'playwright' in result:
                        self._playwright = result['playwright']
                    if 'browser' in result:
                        self._browser = result['browser']
            except Exception:
                pass  # Lỗi đã được job_fn tự xử lý và log

        # ── Cleanup khi app tắt ──
        if self._browser:
            try:
                self._browser.close()
            except Exception:
                pass
        if self._playwright:
            try:
                self._playwright.stop()
            except Exception:
                pass

# ============================================================
# GUI
# ============================================================
class App:
    def __init__(self, root):
        self.root = root; self.root.title('📦 TTS Bill — In & Tính Bill TikTok Seller')
        self.root.geometry('680x820'); self.root.minsize(580, 700)
        self.root.configure(bg='#f0f2f5')

        # Hiển thị tên file gọn gàng, không hiện đường dẫn _MEI* dài dòng
        self._cookie_real = str(DEFAULT_COOKIE) if DEFAULT_COOKIE.exists() else ''
        self._master_real = str(MASTER_DEFAULT) if MASTER_DEFAULT.exists() else ''
        # Retail: ưu tiên file cạnh exe (nếu có), fallback file nhúng
        local_retail = BASE_DIR / 'sản phẩm bán lẻ.xlsx'
        if local_retail.exists():
            self._retail_real = str(local_retail)
        elif RETAIL_DEFAULT.exists():
            self._retail_real = str(RETAIL_DEFAULT)
        else:
            self._retail_real = ''
        if getattr(sys, 'frozen', False) and DEFAULT_COOKIE.exists():
            self.cookie_path = tk.StringVar(value=f'[{DEFAULT_COOKIE.name} — đã tích hợp sẵn]')
        else:
            self.cookie_path = tk.StringVar(value=self._cookie_real)
        if getattr(sys, 'frozen', False) and MASTER_DEFAULT.exists():
            self.master_path = tk.StringVar(value=f'[{MASTER_DEFAULT.name} — đã tích hợp sẵn]')
        else:
            self.master_path = tk.StringVar(value=self._master_real)
        self.retail_path = tk.StringVar(value=self._retail_real if self._retail_real else '')
        if self._retail_real:
            p = Path(self._retail_real)
            if getattr(sys, 'frozen', False):
                self.retail_path.set(f'[{p.name} — đã tích hợp sẵn]')
        self.doc_type = tk.StringVar(value='a4'); self.max_orders = tk.IntVar(value=0)
        self.auto_print = tk.BooleanVar(value=False)  # tự động in PDF sau khi tải
        self.printer_name = tk.StringVar(value='HP LaserJet Pro 4001 4002 4003 4004 PCL-6 (V4)')
        self.pdf_print_settings = tk.StringVar(value='paper=A4')  # cài đặt in cho PDF (SumatraPDF)
        self.output_dir = tk.StringVar(value=str(BASE_DIR / 'outputs'))
        self.schedule_mode = tk.StringVar(value='once')
        self.interval_hours = tk.IntVar(value=1)
        self.daily_times = tk.StringVar(value='08:00, 14:00, 20:00')
        self.running = False; self.scheduler_running = False; self.result_files = []
        self.stop_event = threading.Event()  # cờ dừng cho automation loop
        self._warned_sumatra = False  # tránh log cảnh báo SumatraPDF nhiều lần

        # Worker thread chuyên biệt cho Playwright (tránh lỗi thread-safety)
        self._worker = AutomationWorker()
        self._worker.start()

        self.scheduler = Scheduler(
            callback=self._execute_job_sync,
            log_cb=lambda m, t='': self.root.after(0, self.log, m, t),
            state_cb=lambda s, m: self.root.after(0, self._set_state, s, m))

        self._build_ui(); self._update_cookie_status(); self._update_master_status(); self._update_retail_status()
        os.makedirs(self.output_dir.get(), exist_ok=True)

        # Bắt sự kiện đóng cửa sổ → cleanup browser
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ═══════════════════════════════════════════════════
    # UI BUILDERS
    # ═══════════════════════════════════════════════════
    def _s(self, title, builder):
        f = tk.LabelFrame(self.root, text=title, bg='#f0f2f5', fg='#374151',
                          font=('Segoe UI', 10, 'bold'), padx=12, pady=8, bd=0, highlightthickness=1, highlightbackground='#e5e7eb')
        f.pack(fill='x', padx=14, pady=(0,5)); builder(f)

    def _row(self, parent, label, var, cmd, status_lbl_key):
        r = tk.Frame(parent, bg='#f0f2f5'); r.pack(fill='x')
        tk.Label(r, text=label, font=('Segoe UI',9), bg='#f0f2f5', fg='#6b7280', width=10, anchor='w').pack(side='left')
        tk.Entry(r, textvariable=var, font=('Segoe UI',9), relief='solid', bd=1).pack(side='left', fill='x', expand=True)
        tk.Button(r, text='📂', command=cmd, font=('Segoe UI',9), relief='flat', bg='#e5e7eb',
                  cursor='hand2', width=3).pack(side='left', padx=(4,0))
        lbl = tk.Label(parent, text='', font=('Segoe UI',8), bg='#f0f2f5', fg='#059669')
        lbl.pack(anchor='w', pady=(2,0))
        setattr(self, status_lbl_key, lbl)

    def _build_ui(self):
        # Header
        h = tk.Frame(self.root, bg='#2563eb', height=56)
        h.pack(fill='x'); h.pack_propagate(False)
        tk.Label(h, text='📦 TTS Bill', font=('Segoe UI',16,'bold'), bg='#2563eb', fg='white').pack(side='left', padx=16, pady=12)
        tk.Label(h, text='In & Tính Bill TikTok Seller', font=('Segoe UI',9), bg='#2563eb', fg='#bfdbfe').pack(side='left', pady=16)

        # Config sections
        self._s('Cấu hình', self._build_config)
        self._s('⚙ Tùy chọn in', self._build_options)
        self._s('🕐 Hẹn giờ', self._build_schedule)

        # Buttons
        btn = tk.Frame(self.root, bg='#f0f2f5'); btn.pack(fill='x', padx=14, pady=(8,0))
        self.start_btn = tk.Button(btn, text='▶ CHẠY NGAY', command=self._start_once,
            font=('Segoe UI',11,'bold'), bg='#2563eb', fg='white', relief='flat', cursor='hand2', padx=16, pady=8,
            activebackground='#1d4ed8', activeforeground='white')
        self.start_btn.pack(side='left', fill='x', expand=True, padx=(0,3))
        self.sched_btn = tk.Button(btn, text='⏰ CHẠY LỊCH', command=self._start_schedule,
            font=('Segoe UI',11,'bold'), bg='#059669', fg='white', relief='flat', cursor='hand2', padx=16, pady=8,
            activebackground='#047857', activeforeground='white')
        self.sched_btn.pack(side='left', fill='x', expand=True, padx=(3,3))
        self.stop_btn = tk.Button(btn, text='⏹ DỪNG', command=self._stop_all,
            font=('Segoe UI',11,'bold'), bg='#e5e7eb', fg='#6b7280', relief='flat', cursor='hand2', padx=16, pady=8,
            state='disabled', activebackground='#d1d5db')
        self.stop_btn.pack(side='left', fill='x', expand=True, padx=(3,0))

        # Status bar
        sf = tk.Frame(self.root, bg='#f0f2f5'); sf.pack(fill='x', padx=14, pady=(6,0))
        self.status_bar = tk.Label(sf, text='🟢 Sẵn sàng', font=('Segoe UI',9,'bold'),
                                   bg='#f0f2f5', fg='#059669', anchor='w')
        self.status_bar.pack(fill='x')

        # LOG — phần chính
        self._build_log()

        # Results
        self._build_results()

    def _build_config(self, parent):
        self._row(parent, 'Cookie', self.cookie_path, self._select_cookie, 'cookie_status')
        self._row(parent, 'Master (Combo)', self.master_path, self._select_master, 'master_status')
        self._row(parent, 'Bán lẻ', self.retail_path, self._select_retail, 'retail_status')
        self._row(parent, 'Thư mục lưu', self.output_dir, self._select_output_dir, 'output_status')
        self.output_status.config(text='✅ Sẵn sàng', fg='#059669')

    def _build_options(self, parent):
        r = tk.Frame(parent, bg='#f0f2f5'); r.pack(fill='x')
        tk.Label(r, text='Số đơn in:', font=('Segoe UI',9), bg='#f0f2f5', fg='#6b7280').pack(side='left')
        tk.Spinbox(r, from_=0, to=9999, width=6, textvariable=self.max_orders, font=('Segoe UI',9)).pack(side='left', padx=(8,0))
        tk.Label(r, text='(0 = tất cả, >50 = tự chia batch)', font=('Segoe UI',8), bg='#f0f2f5', fg='#9ca3af').pack(side='left', padx=6)
        tk.Label(parent, text='📑 Luôn in: Danh sách lấy hàng (A4) + Đơn vận chuyển (A6)',
                 font=('Segoe UI',8), bg='#f0f2f5', fg='#6b7280').pack(anchor='w', pady=(4,0))
        # In tự động
        r3 = tk.Frame(parent, bg='#f0f2f5'); r3.pack(fill='x', pady=(4,0))
        tk.Checkbutton(r3, text='🖨️ In tự động ra máy in', variable=self.auto_print,
                       bg='#f0f2f5', font=('Segoe UI',9), activebackground='#f0f2f5').pack(side='left')
        tk.Entry(r3, textvariable=self.printer_name, font=('Segoe UI',8), width=30).pack(side='left', padx=(6,0))
        r4 = tk.Frame(parent, bg='#f0f2f5'); r4.pack(fill='x')
        tk.Label(r4, text='   Cài đặt in PDF:', font=('Segoe UI',8), bg='#f0f2f5', fg='#6b7280').pack(side='left')
        tk.Entry(r4, textvariable=self.pdf_print_settings, font=('Segoe UI',8), width=35).pack(side='left', padx=(4,0))
        tk.Label(r4, text='(vd: paper=A4,pagespersheet=2)', font=('Segoe UI',7), bg='#f0f2f5', fg='#9ca3af').pack(side='left', padx=4)

    def _build_schedule(self, parent):
        modes = [('Chạy 1 lần','once'), ('Lặp mỗi N giờ','interval'), ('Giờ cố định trong ngày','daily')]
        for i, (lbl, mode) in enumerate(modes):
            tk.Radiobutton(parent, text=lbl, variable=self.schedule_mode, value=mode,
                           bg='#f0f2f5', font=('Segoe UI',9), command=self._on_sched_change,
                           activebackground='#f0f2f5').grid(row=i, column=0, sticky='w', pady=1)

        self.intv_f = tk.Frame(parent, bg='#f0f2f5')
        tk.Label(self.intv_f, text='Mỗi', bg='#f0f2f5', font=('Segoe UI',9)).pack(side='left')
        tk.Spinbox(self.intv_f, from_=1, to=24, width=3, textvariable=self.interval_hours, font=('Segoe UI',9)).pack(side='left', padx=4)
        tk.Label(self.intv_f, text='giờ', bg='#f0f2f5', font=('Segoe UI',9)).pack(side='left')

        self.day_f = tk.Frame(parent, bg='#f0f2f5')
        tk.Label(self.day_f, text='Lúc:', bg='#f0f2f5', font=('Segoe UI',9)).pack(side='left')
        tk.Entry(self.day_f, textvariable=self.daily_times, width=22, font=('Segoe UI',9)).pack(side='left', padx=4)
        tk.Label(self.day_f, text='VD: 08:00, 14:00', bg='#f0f2f5', font=('Segoe UI',8), fg='#9ca3af').pack(side='left')
        self._on_sched_change()

    def _build_log(self):
        lf = tk.LabelFrame(self.root, text='📋 Log', bg='#f0f2f5', fg='#374151',
                           font=('Segoe UI',10,'bold'), padx=4, pady=4, bd=0, highlightthickness=1, highlightbackground='#e5e7eb')
        lf.pack(fill='both', expand=True, padx=14, pady=(6,0))

        # Toolbar nhỏ: clear log
        tb = tk.Frame(lf, bg='#f0f2f5'); tb.pack(fill='x', pady=(0,2))
        tk.Button(tb, text='Xóa log', command=self._clear_log, font=('Segoe UI',8),
                  relief='flat', bg='#e5e7eb', cursor='hand2', padx=8).pack(side='right')

        self.log_text = tk.Text(lf, font=('Consolas',9),
                                bg='#1a1d23', fg='#d4d4d4', relief='flat',
                                padx=12, pady=10, state='disabled', wrap='word',
                                selectbackground='#264f78', selectforeground='white')
        self.log_text.pack(fill='both', expand=True)

        # Tags màu
        for tag, cfg in {
            'ts':     '#6b7280',
            'ok':     '#4ade80',
            'err':    '#f87171',
            'warn':   '#fbbf24',
            'info':   '#60a5fa',
            'batch':  '#c084fc',
            'result': '#f472b6',
            'dim':    '#6b7280',
        }.items():
            self.log_text.tag_config(tag, foreground=cfg)
        self.log_text.tag_config('bold_ok', foreground='#4ade80', font=('Consolas',9,'bold'))

    def _build_results(self):
        rf = tk.LabelFrame(self.root, text='📁 Kết quả — double-click để mở file', bg='#f0f2f5', fg='#374151',
                           font=('Segoe UI',10,'bold'), padx=6, pady=6, bd=0, highlightthickness=1, highlightbackground='#e5e7eb')
        rf.pack(fill='x', padx=14, pady=(6,10))

        self.result_list = tk.Listbox(rf, height=4, font=('Segoe UI',9), relief='flat',
                                      selectbackground='#2563eb', selectforeground='white',
                                      activestyle='none')
        self.result_list.pack(fill='x'); self.result_list.bind('<Double-Button-1>', self._open_result)

        br = tk.Frame(rf, bg='#f0f2f5'); br.pack(fill='x', pady=(4,0))
        tk.Button(br, text='📂 Mở thư mục', command=self._open_output_dir,
                  font=('Segoe UI',9), relief='flat', bg='#e5e7eb', cursor='hand2').pack(side='right')
        tk.Button(br, text='🗑 Xóa hết', command=self._clear_results,
                  font=('Segoe UI',9), relief='flat', bg='#e5e7eb', cursor='hand2').pack(side='right', padx=4)

    # ═══════════════════════════════════════════════════
    # ACTIONS
    # ═══════════════════════════════════════════════════
    def _on_sched_change(self):
        self.intv_f.grid_forget(); self.day_f.grid_forget()
        if self.schedule_mode.get() == 'interval': self.intv_f.grid(row=1, column=0, sticky='w', padx=(20,0), pady=(4,0))
        elif self.schedule_mode.get() == 'daily': self.day_f.grid(row=2, column=0, sticky='w', padx=(20,0), pady=(4,0))

    def _select_cookie(self):
        p = filedialog.askopenfilename(filetypes=[('JSON','*.json')])
        if p: self._cookie_real = p; self.cookie_path.set(p); self._update_cookie_status()
    def _select_master(self):
        p = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
        if p: self._master_real = p; self.master_path.set(p); self._update_master_status()
    def _select_retail(self):
        p = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
        if p: self._retail_real = p; self.retail_path.set(p); self._update_retail_status()
    def _select_output_dir(self):
        p = filedialog.askdirectory()
        if p: self.output_dir.set(p)

    def _update_cookie_status(self):
        p = Path(self._cookie_real) if self._cookie_real else Path('')
        if self._cookie_real and p.exists():
            try:
                d = json.loads(p.read_text(encoding='utf-8'))
                cs = d.get('cookies', d); n = len(cs) if isinstance(cs, list) else 0
                self.cookie_status.config(text=f'✅ {n} cookies', fg='#059669')
            except: self.cookie_status.config(text='⚠ File lỗi', fg='#d97706')
        else: self.cookie_status.config(text='❌ Chưa chọn', fg='#dc2626')

    def _update_master_status(self):
        p = Path(self._master_real) if self._master_real else Path('')
        if self._master_real and p.exists() and p.suffix.lower() == '.xlsx':
            try:
                from openpyxl import load_workbook
                wb = load_workbook(str(p), data_only=True); n = wb.active.max_row - 1; wb.close()
                self.master_status.config(text=f'✅ {n} SKU', fg='#059669')
            except: self.master_status.config(text='⚠ File lỗi', fg='#d97706')
        else: self.master_status.config(text='❌ Chưa chọn', fg='#dc2626')

    def _update_retail_status(self):
        p = Path(self._retail_real) if self._retail_real else Path('')
        if self._retail_real and p.exists():
            try:
                from openpyxl import load_workbook
                wb = load_workbook(str(p), data_only=True); n = wb.active.max_row - 1; wb.close()
                self.retail_status.config(text=f'✅ {n} SKU', fg='#059669')
            except: self.retail_status.config(text='⚠ File lỗi', fg='#d97706')
        else: self.retail_status.config(text='❌ Chưa chọn', fg='#dc2626')

    def log(self, msg, tag=''):
        self.log_text.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert('end', ts + '  ', 'ts')
        self.log_text.insert('end', msg + '\n', tag)
        self.log_text.see('end'); self.log_text.config(state='disabled')

    def _set_state(self, step, msg):
        c = {'idle':'#059669','running':'#d97706','waiting':'#2563eb','done':'#059669','error':'#dc2626'}
        self.status_bar.config(text=msg, fg=c.get(step,'#333'))

    def _set_buttons(self, s):
        if s in ('running','scheduled'):
            self.start_btn.config(state='disabled', bg='#93c5fd', text='▶ ĐANG CHẠY')
            self.sched_btn.config(state='disabled', bg='#6ee7b7')
            self.stop_btn.config(state='normal', bg='#fca5a5', fg='#7f1d1d')
        else:
            self.start_btn.config(state='normal', bg='#2563eb', text='▶ CHẠY NGAY')
            self.sched_btn.config(state='normal', bg='#059669', text='⏰ CHẠY LỊCH')
            self.stop_btn.config(state='disabled', bg='#e5e7eb', fg='#6b7280', text='⏹ DỪNG')

    def _start_once(self):
        if self.running: return
        if not Path(self._cookie_real).exists(): messagebox.showerror('Lỗi','Chọn file cookie JSON hợp lệ.'); return
        self.running = True; self._set_buttons('running'); self._clear_results(); self._clear_log()
        self.log('▶ Bắt đầu...', 'bold_ok')
        self._worker.run_job(self._execute_job)

    def _start_schedule(self):
        if not Path(self._cookie_real).exists(): messagebox.showerror('Lỗi','Chọn file cookie JSON hợp lệ.'); return
        mode = self.schedule_mode.get()
        if mode == 'once': self._start_once(); return
        if mode == 'daily':
            times = [t.strip() for t in self.daily_times.get().split(',') if t.strip()]
            if not times: messagebox.showerror('Lỗi','Nhập ít nhất 1 giờ (VD: 08:00).'); return
        self.scheduler.configure(mode=mode, interval_hours=self.interval_hours.get(),
            daily_times=[t.strip() for t in self.daily_times.get().split(',') if t.strip()])
        self._clear_log(); self._clear_results()
        self.log(f'⏰ Hẹn giờ: {mode}', 'info')
        if mode == 'interval': self.log(f'   Chạy mỗi {self.interval_hours.get()} giờ', 'info')
        elif mode == 'daily': self.log(f'   Chạy lúc {self.daily_times.get()}', 'info')
        self.scheduler_running = True; self._set_buttons('scheduled'); self.scheduler.start()

    def _stop_all(self):
        self.scheduler.stop(); self.scheduler_running = False; self.running = False
        self.stop_event.set()  # báo hiệu dừng cho automation loop
        self._worker.stop_job()  # dừng job đang chạy (giữ browser)
        self._set_buttons('idle'); self._set_state('idle','⏹ Đã dừng'); self.log('⏹ Đã dừng', 'warn')

    def _execute_job(self, playwright=None, browser=None, job_stop=None):
        """Thực thi 1 job automation + calculator. Được gọi từ AutomationWorker thread."""
        try:
            cookie = self._cookie_real; out_dir = self.output_dir.get()
            master = self._master_real; retail = self._retail_real
            doc = self.doc_type.get(); n = self.max_orders.get()
            os.makedirs(out_dir, exist_ok=True)

            # Step 1: Download
            self.stop_event.clear()
            # Dùng job_stop của worker nếu có, fallback về stop_event của app
            stop_signal = job_stop if job_stop is not None else self.stop_event
            self.log(f'📥 Tải PDF ({n if n>0 else "tất cả"} đơn)...', 'info')
            pdf_paths, pw, br = run_automation(cookie, doc, out_dir, n,
                lambda m, t='': self.root.after(0, self.log, m, t),
                lambda s, m: self.root.after(0, self._set_state, s, m),
                stop_signal,
                existing_playwright=playwright, existing_browser=browser)

            for p in pdf_paths:
                if Path(p).exists(): self._add_result(f'📄 {Path(p).name}')
            self.log(f'📥 Đã tải {len(pdf_paths)} file PDF', 'ok' if pdf_paths else 'warn')

            # Step 2: Calculator
            results = []
            if pdf_paths:
                self.log('📊 Đang tính bill...', 'info')
                results = run_calculator(pdf_paths, out_dir, master, retail,
                    lambda m, t='': self.root.after(0, self.log, m, t))
                for r in results:
                    for key, lbl in [('excel','📊'),('pdf','📕'),('pdf_grouped','📋')]:
                        fp = r['files'].get(key)
                        if fp and Path(fp).exists(): self._add_result(f'{lbl} {Path(fp).name}')

            # Step 3: In tự động ra máy in (nếu bật)
            if self.auto_print.get():
                printer = self.printer_name.get().strip()
                self.log(f'🖨️ Đang in ra "{printer}"...', 'info')
                # Thu thập các file cần in từ lần chạy này
                files_to_print = []
                # File PDF Shipping Label (từ TikTok)
                for p in pdf_paths:
                    name = Path(p).name.lower()
                    if Path(p).exists() and ('shipping' in name or 'vận chuyển' in name):
                        files_to_print.append(str(p))
                # File Excel báo cáo gộp SKU (từ calculator)
                for r in results:
                    excel_path = r['files'].get('excel')
                    if excel_path and Path(excel_path).exists():
                        files_to_print.append(excel_path)
                # In từng file
                pdf_settings = self.pdf_print_settings.get().strip()
                for fp in files_to_print:
                    try:
                        is_pdf = fp.lower().endswith('.pdf')
                        # PDF: dùng SumatraPDF với cài đặt tùy chỉnh
                        # Excel: dùng PowerShell với máy in mặc định
                        self._print_file(fp, printer,
                                       pdf_settings=pdf_settings if is_pdf else '')
                        self.log(f'  ✓ Đã gửi in: {Path(fp).name}', 'ok')
                    except Exception as e:
                        self.log(f'  ✗ Lỗi in {Path(fp).name}: {e}', 'err')
                if not files_to_print:
                    self.log('  ⚠ Không có file nào để in', 'warn')

            self.log('🏁 HOÀN THÀNH!', 'bold_ok')
            self._set_state('done', f'✅ Hoàn thành lúc {datetime.now().strftime("%H:%M:%S")}')
            # Trả về browser/playwright để worker lưu lại cho lần sau
            return {'playwright': pw, 'browser': br, 'pdf_paths': pdf_paths}
        except Exception as e:
            self.log(f'✗ Lỗi: {e}', 'err'); self._set_state('error', f'✗ {e}')
            return {}
        finally:
            self.running = False
            if self.schedule_mode.get() == 'once' or not self.scheduler_running:
                self.root.after(0, lambda: self._set_buttons('idle'))

    def _execute_job_sync(self):
        """Wrapper cho scheduler — gửi job vào worker và đợi hoàn thành."""
        done = threading.Event()
        def job_wrapper(pw, br, stop):
            try:
                return self._execute_job(pw, br, stop)
            finally:
                self.root.after(0, done.set)
        self._worker.run_job(job_wrapper)
        done.wait()  # block cho đến khi job hoàn thành

    def _add_result(self, text):
        self.root.after(0, lambda: self.result_list.insert('end', text)); self.result_files.append(text)
    def _clear_results(self): self.result_list.delete(0,'end'); self.result_files.clear()
    def _clear_log(self):
        self.log_text.config(state='normal'); self.log_text.delete('1.0','end'); self.log_text.config(state='disabled')

    def _open_result(self, event):
        sel = self.result_list.curselection()
        if sel:
            item = self.result_list.get(sel[0])
            for pfx in ['📄 ','📊 ','📕 ','📋 ']:
                if item.startswith(pfx): item = item[len(pfx):]; break
            fp = str(Path(self.output_dir.get()) / item)
            if Path(fp).exists(): os.startfile(fp)

    def _open_output_dir(self):
        d = self.output_dir.get()
        if Path(d).exists(): os.startfile(d)

    def _print_file(self, file_path, printer_name, pdf_settings='paper=A4'):
        """In file (PDF hoặc Excel) ra máy in chỉ định (Windows).
        pdf_settings: chuỗi cài đặt cho SumatraPDF (vd: 'paper=A4,pagespersheet=2')"""
        import subprocess
        fp = str(file_path)
        is_pdf = fp.lower().endswith('.pdf')

        # Cách 1: SumatraPDF — in PDF thẳng đến máy in chỉ định (silent)
        if is_pdf:
            sumatra_paths = [
                r'C:\Users\thanh\AppData\Local\SumatraPDF\SumatraPDF.exe',
                r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
                r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe',
            ]
            for sp in sumatra_paths:
                if Path(sp).exists():
                    cmd = [sp, '-print-to', printer_name, fp]
                    if pdf_settings:
                        cmd += ['-print-settings', pdf_settings]
                    subprocess.run(cmd, check=False, timeout=60)
                    return
            # Không tìm thấy SumatraPDF → cảnh báo 1 lần
            if not getattr(self, '_warned_sumatra', False):
                self.root.after(0, self.log,
                    '⚠ Chưa cài SumatraPDF — in qua PowerShell (máy in mặc định). '
                    'Tải tại: https://www.sumatrapdfreader.org', 'warn')
                self._warned_sumatra = True

        # Cách 2: PowerShell Start-Process — hoạt động từ mọi thread
        ps_cmd = f'Start-Process -FilePath "{fp}" -Verb Print'
        subprocess.run(['powershell', '-Command', ps_cmd],
                      capture_output=True, timeout=30)

    def _on_close(self):
        """Dọn dẹp khi tắt app: dừng scheduler, đóng browser, hủy cửa sổ."""
        self.scheduler.stop()
        self.scheduler_running = False
        self.running = False
        self.stop_event.set()
        self._worker.shutdown()  # đóng browser + dừng worker thread
        self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
