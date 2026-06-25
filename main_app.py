"""
main_app.py — Desktop App TikTok Seller Automation + Bill Calculate
====================================================================
Chạy: python main_app.py
Đóng gói .exe: pyinstaller --onefile --windowed --name "TTS_Bill" main_app.py
"""
import os, sys, json, time, shutil, threading
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
    # Chạy file .exe
    _os.environ['PLAYWRIGHT_BROWSERS_PATH'] = _os.path.join(sys._MEIPASS, 'ms-playwright')
    # Outputs luôn lưu cạnh exe
    BASE_DIR = Path(sys.executable).parent
    BILL_DIR = BASE_DIR / 'bill_calculate'
    UPLOAD_DIR = BILL_DIR / 'uploads'
    # Mặc định: dùng file đã nhúng trong exe (không giải nén ra ngoài)
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
ORDERS_URL = 'https://seller-vn.tiktok.com/order?order_status%5B%5D=2&selected_sort=11&tab=to_ship'
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

    # ── Tái sử dụng browser cũ nếu có, nếu không thì tạo mới ──
    if existing_browser and existing_playwright:
        playwright = existing_playwright
        browser = existing_browser
        # Dùng context hiện có (đã đăng nhập + có cookie sẵn)
        context = browser.contexts[0] if browser.contexts else browser.new_context(
            viewport={'width': 1366, 'height': 768},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
            accept_downloads=True)
        page = context.new_page()  # tạo tab mới, không đụng tab cũ
        log_cb('♻ Dùng lại browser đang mở — vào thẳng trang đơn hàng...', 'info')
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

            # In chứng từ → chọn A4 → In
            state_cb('printing', f'Batch {batch_num}: Đang in...')
            page.locator('button:has-text("In chứng từ")').first.wait_for(timeout=10000)
            page.locator('button:has-text("In chứng từ")').first.click()
            page.wait_for_timeout(3000)

            page.wait_for_selector('text=Nhãn vận chuyển', timeout=10000)
            page.wait_for_timeout(1000)
            # Giữ nguyên A6 vận chuyển (đã checked) + check thêm A4
            # → 2 loại: A4 + A6 vận chuyển
            for doc_label in ['Danh sách lấy hàng (A4)']:
                try:
                    lbl = page.locator('label').filter(has_text=doc_label).first
                    if lbl.count() > 0:
                        inp = lbl.locator('input')
                        if inp.count() > 0 and not inp.is_checked():
                            lbl.click(); page.wait_for_timeout(300)
                except: pass
            page.wait_for_timeout(2000)
            log_cb('  ✓ Đã chọn: A4 + A6 vận chuyển', 'ok')

            # Click nút In trong modal
            in_btn = None
            for _ in range(20):
                for b in page.locator('.p-modal button, [class*="modal"] button').all():
                    if b.inner_text().strip() == 'In' and b.is_enabled(): in_btn = b; break
                if in_btn: break; page.wait_for_timeout(500)
            if not in_btn:
                in_btn = page.locator('.p-modal button:has-text("In")').first
                in_btn.evaluate('el => { el.disabled = false; el.classList.remove("p-btn-disabled"); }')
            in_btn.click()
            log_cb('  ✓ Đã bấm In', 'ok')
            page.wait_for_timeout(3000)

            # ── Bước 1: Bấm "Tiếp tục in" trong popup đầu tiên ──
            state_cb('printing', f'Batch {batch_num}: Đợi popup "Tiếp tục in"...')
            tieptuc_btn = None
            for _ in range(30):
                btns = page.locator('button:has-text("Tiếp tục in"), button:has-text("Continue"), button:has-text("Xác nhận")')
                cnt = btns.count()
                for j in range(cnt):
                    b = btns.nth(j)
                    try:
                        if b.is_visible(timeout=500):
                            tieptuc_btn = b; break
                    except: pass
                if tieptuc_btn: break
                page.wait_for_timeout(1000)
            if not tieptuc_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Tiếp tục in" — kiểm tra popup!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_tieptucin_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break
            tieptuc_btn.click(timeout=5000)
            log_cb('  ✓ Đã bấm "Tiếp tục in"', 'ok')
            page.wait_for_timeout(3000)

            # ── Bước 2: Bấm "Tải xuống tất cả tập tin" trong popup thứ hai ──
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

        # KHÔNG đóng browser — giữ lại cho đến khi tắt chương trình
        return pdf_files, playwright, browser
    except Exception as e:
        log_cb(f'  ✗ Lỗi: {e}', 'err')
        # KHÔNG đóng browser — giữ lại để user xem lỗi
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
        self.output_dir = tk.StringVar(value=str(BASE_DIR / 'outputs'))
        self.schedule_mode = tk.StringVar(value='once')
        self.interval_hours = tk.IntVar(value=1)
        self.daily_times = tk.StringVar(value='08:00, 14:00, 20:00')
        self.running = False; self.scheduler_running = False; self.result_files = []
        self.stop_event = threading.Event()  # cờ dừng cho automation loop
        self._playwright = None   # giữ playwright instance
        self._browser = None      # giữ browser instance — KHÔNG đóng cho đến khi tắt app

        # Bắt sự kiện đóng cửa sổ để dọn dẹp browser
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.scheduler = Scheduler(
            callback=self._execute_job,
            log_cb=lambda m, t='': self.root.after(0, self.log, m, t),
            state_cb=lambda s, m: self.root.after(0, self._set_state, s, m))

        self._build_ui(); self._update_cookie_status(); self._update_master_status(); self._update_retail_status()
        os.makedirs(self.output_dir.get(), exist_ok=True)

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
        tk.Label(parent, text='📑 Luôn in cả 3 loại: Danh sách lấy hàng (A4) + Đóng gói (A6) + Vận chuyển (A6)',
                 font=('Segoe UI',8), bg='#f0f2f5', fg='#6b7280').pack(anchor='w', pady=(4,0))

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
        threading.Thread(target=self._execute_job, daemon=True).start()

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
        self._set_buttons('idle'); self._set_state('idle','⏹ Đã dừng'); self.log('⏹ Đã dừng', 'warn')

    def _execute_job(self):
        try:
            cookie = self._cookie_real; out_dir = self.output_dir.get()
            master = self._master_real; retail = self._retail_real
            doc = self.doc_type.get(); n = self.max_orders.get()
            os.makedirs(out_dir, exist_ok=True)

            # Step 1: Download
            self.stop_event.clear()
            self.log(f'📥 Tải PDF ({n if n>0 else "tất cả"} đơn)...', 'info')
            result = run_automation(cookie, doc, out_dir, n,
                lambda m, t='': self.root.after(0, self.log, m, t),
                lambda s, m: self.root.after(0, self._set_state, s, m),
                self.stop_event,
                existing_playwright=self._playwright,
                existing_browser=self._browser)
            pdf_paths, new_playwright, new_browser = result

            # Lưu lại để lần sau tái sử dụng (KHÔNG đóng browser)
            self._browser = new_browser
            self._playwright = new_playwright

            for p in pdf_paths:
                if Path(p).exists(): self._add_result(f'📄 {Path(p).name}')
            self.log(f'📥 Đã tải {len(pdf_paths)} file PDF', 'ok' if pdf_paths else 'warn')

            # Step 2: Calculator
            if pdf_paths:
                self.log('📊 Đang tính bill...', 'info')
                results = run_calculator(pdf_paths, out_dir, master, retail,
                    lambda m, t='': self.root.after(0, self.log, m, t))
                for r in results:
                    for key, lbl in [('excel','📊'),('pdf','📕'),('pdf_grouped','📋')]:
                        fp = r['files'].get(key)
                        if fp and Path(fp).exists(): self._add_result(f'{lbl} {Path(fp).name}')

            self.log('🏁 HOÀN THÀNH!', 'bold_ok')
            self._set_state('done', f'✅ Hoàn thành lúc {datetime.now().strftime("%H:%M:%S")}')
        except Exception as e:
            self.log(f'✗ Lỗi: {e}', 'err'); self._set_state('error', f'✗ {e}')
        finally:
            self.running = False
            if self.schedule_mode.get() == 'once' or not self.scheduler_running:
                self.root.after(0, lambda: self._set_buttons('idle'))

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

    def _on_close(self):
        """Đóng browser + playwright khi tắt chương trình."""
        self.scheduler.stop()
        self.scheduler_running = False
        self.running = False
        self.stop_event.set()
        if self._browser:
            try: self._browser.close()
            except: pass
        if self._playwright:
            try: self._playwright.stop()
            except: pass
        self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
