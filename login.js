const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const { spawn, execSync } = require('child_process');

const COOKIE_FILE = path.join(__dirname, 'seller-vn.tiktok.com_22-06-2026.json');
const TARGET_URL = 'https://seller-vn.tiktok.com';
const ORDERS_URL = 'https://seller-vn.tiktok.com/order?order_status%5B%5D=2&selected_sort=11&tab=to_ship';

// Bill Calculate paths
const CALC_DIR = path.join(__dirname, 'bill_calculate');
const CALC_UPLOAD_DIR = path.join(CALC_DIR, 'uploads');
const CALC_OUTPUT_DIR = path.join(CALC_DIR, 'outputs');
const CALC_SCRIPT = path.join(CALC_DIR, 'calculator.py');
const MASTER_PATH = path.join(CALC_DIR, 'master_data.xlsx');

// Output naming: timestamp giây để phân biệt giữa các lần chạy
const NOW = new Date();
const TS = `${NOW.getFullYear()}${String(NOW.getMonth()+1).padStart(2,'0')}${String(NOW.getDate()).padStart(2,'0')}_${String(NOW.getHours()).padStart(2,'0')}${String(NOW.getMinutes()).padStart(2,'0')}${String(NOW.getSeconds()).padStart(2,'0')}`;
const PDF_RAW = path.join(__dirname, `PDF_goc_TTS_${TS}.pdf`);   // PDF gốc từ TikTok
const PDF_NAME_RAW = `PDF_goc_TTS_${TS}.pdf`;

// Số đơn cần in: 0 = tất cả, >0 = số lượng cụ thể
const MAX_ORDERS = parseInt(process.env.MAX_ORDERS || '0', 10);

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================================
// PIPELINE: Gọi calculator.py để xử lý PDF đã tải
// ============================================================
async function processWithCalculator(pdfPath, originalFileName) {
  console.log('\n' + '='.repeat(50));
  console.log('📊 PIPELINE: Chạy bill_calculate xử lý PDF...');
  console.log('='.repeat(50));

  // Kiểm tra môi trường
  if (!fs.existsSync(CALC_SCRIPT)) {
    console.log('⚠️ Không tìm thấy calculator.py tại:', CALC_SCRIPT);
    console.log('   Bỏ qua bước xử lý. Bạn có thể chạy thủ công.');
    return null;
  }

  if (!fs.existsSync(MASTER_PATH)) {
    console.log('⚠️ Không tìm thấy master_data.xlsx tại:', MASTER_PATH);
    console.log('   Bỏ qua bước xử lý. Bạn có thể chạy thủ công.');
    return null;
  }

  // Tạo tên file đầu vào (timestamp)
  const now = new Date();
  const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
  const inputName = `tiktok_${ts}.pdf`;
  const inputPath = path.join(CALC_UPLOAD_DIR, inputName);

  // Copy PDF vào uploads/
  fs.mkdirSync(CALC_UPLOAD_DIR, { recursive: true });
  fs.copyFileSync(pdfPath, inputPath);
  console.log(`📄 Đã copy PDF vào: ${inputPath}`);

  // Thử gọi Flask API trước
  const flaskResult = await processViaFlaskApi(pdfPath);
  if (flaskResult) {
    console.log('='.repeat(50));
    return flaskResult;
  }

  // Fallback: Gọi Python trực tiếp
  // Kiểm tra Python
  let pythonCmd = 'python';
  try {
    execSync('python --version', { stdio: 'ignore' });
  } catch {
    try {
      execSync('python3 --version', { stdio: 'ignore' });
      pythonCmd = 'python3';
    } catch {
      console.log('⚠️ Không tìm thấy Python. Bỏ qua bước xử lý.');
      console.log('   Bạn có thể chạy Flask API: cd bill_calculate && python app.py');
      return null;
    }
  }

  // Chạy calculator.py
  console.log(`🐍 Đang chạy: ${pythonCmd} calculator.py "${inputName}"...`);
  fs.mkdirSync(CALC_OUTPUT_DIR, { recursive: true });

  return new Promise((resolve) => {
    const proc = spawn(pythonCmd, [CALC_SCRIPT, inputPath], {
      cwd: CALC_DIR,
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => {
      const text = data.toString();
      stdout += text;
      process.stdout.write(text);
    });

    proc.stderr.on('data', (data) => {
      const text = data.toString();
      stderr += text;
      process.stderr.write(text);
    });

    proc.on('close', (code) => {
      console.log(`\n📊 Python exit code: ${code}`);

      const baseName = path.parse(inputName).name;
      const expectedExcel = path.join(CALC_OUTPUT_DIR, `kết quả ${baseName}.xlsx`);
      const expectedPdf = path.join(CALC_OUTPUT_DIR, `Danh_sach_${baseName}.pdf`);

      const results = {
        success: code === 0,
        inputPath,
        excelPath: null,
        pdfPath: null,
        stdout,
        stderr,
      };

      if (code === 0) {
        if (fs.existsSync(expectedExcel)) {
          results.excelPath = expectedExcel;
          console.log(`✅ Excel kết quả: ${expectedExcel}`);
        }
        if (fs.existsSync(expectedPdf)) {
          results.pdfPath = expectedPdf;
          console.log(`✅ PDF báo cáo:   ${expectedPdf}`);
        }
      } else {
        console.log('❌ Python trả về lỗi. Kiểm tra log phía trên.');
      }

      console.log('='.repeat(50));
      resolve(results);
    });
  });
}

// ============================================================
// PIPELINE ALT: Gọi Flask API thay vì Python trực tiếp
// ============================================================
async function processViaFlaskApi(pdfPath) {
  const FLASK_URL = 'http://localhost:5000/api/process';

  try {
    // Kiểm tra Flask có đang chạy không
    const statusRes = await fetch('http://localhost:5000/api/status', { signal: AbortSignal.timeout(2000) });
    if (!statusRes.ok) return null;

    console.log('🌐 Flask API đang chạy, gửi file qua HTTP...');

    // Dùng native FormData + Blob (có sẵn trong Node 20)
    const pdfBuffer = fs.readFileSync(pdfPath);
    const blob = new Blob([pdfBuffer], { type: 'application/pdf' });
    const form = new FormData();
    form.append('file', blob, PDF_NAME_RAW);

    const response = await fetch(FLASK_URL, {
      method: 'POST',
      body: form,
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      console.log(`⚠️ Flask API lỗi: ${err.error || response.statusText}`);
      return null;
    }

    const data = await response.json();
    console.log('✅ Flask API xử lý thành công!');
    console.log(`   Base: ${data.base_name}`);
    console.log(`   Rows: ${data.rows} | Qty: ${data.tong_qty} | Sold: ${data.tong_sold}`);

    // Tải file kết quả về
    const results = {
      success: true,
      inputPath: pdfPath,
      excelPath: null,
      pdfPath: null,
    };

    if (data.downloads?.excel) {
      const excelUrl = `http://localhost:5000${data.downloads.excel}`;
      const excelRes = await fetch(excelUrl);
      if (excelRes.ok) {
        const excelPath = path.join(__dirname, `kết_quả_${data.base_name}.xlsx`);
        const buf = await excelRes.arrayBuffer();
        fs.writeFileSync(excelPath, Buffer.from(buf));
        results.excelPath = excelPath;
        console.log(`💾 Đã tải Excel về: ${excelPath}`);
      }
    }

    if (data.downloads?.pdf) {
      const pdfUrl = `http://localhost:5000${data.downloads.pdf}`;
      const pdfRes = await fetch(pdfUrl);
      if (pdfRes.ok) {
        const pdfOutPath = path.join(__dirname, `Danh_sach_${data.base_name}.pdf`);
        const buf = await pdfRes.arrayBuffer();
        fs.writeFileSync(pdfOutPath, Buffer.from(buf));
        results.pdfPath = pdfOutPath;
        console.log(`💾 Đã tải PDF báo cáo về: ${pdfOutPath}`);
      }
    }

    return results;
  } catch (e) {
    // Flask không chạy hoặc lỗi kết nối
    console.log('🌐 Flask API không khả dụng, dùng Python trực tiếp...');
    return null;
  }
}

async function main() {
  // ========== 1. Đọc cookies ==========
  console.log('📂 Đang đọc file cookie...');
  const raw = fs.readFileSync(COOKIE_FILE, 'utf-8');
  const cookieData = JSON.parse(raw);
  if (!cookieData.cookies || !Array.isArray(cookieData.cookies)) {
    console.error('❌ File cookie không đúng định dạng.');
    process.exit(1);
  }
  console.log(`✅ ${cookieData.cookies.length} cookies.`);

  // ========== 2. Convert cookies ==========
  const cookies = cookieData.cookies
    .filter(c => c.name && c.value)
    .map(c => {
      const pw = {
        name: c.name, value: c.value,
        domain: c.domain || '.tiktok.com', path: c.path || '/',
      };
      if (c.expirationDate && !c.session) pw.expires = Math.floor(c.expirationDate);
      if (c.httpOnly !== undefined) pw.httpOnly = c.httpOnly;
      if (c.secure !== undefined) pw.secure = c.secure;
      if (c.sameSite) {
        const m = { strict: 'Strict', lax: 'Lax', no_restriction: 'None', unspecified: 'Lax' };
        pw.sameSite = m[c.sameSite] || 'Lax';
      }
      return pw;
    });
  console.log(`✅ Converted ${cookies.length} cookies.`);

  // ========== 3. Launch browser ==========
  console.log('🚀 Đang mở Chromium...');
  const browser = await chromium.launch({
    headless: false,
    args: ['--disable-blink-features=AutomationControlled', '--no-sandbox'],
  });

  let context = null;
  try {
    context = await browser.newContext({
      viewport: { width: 1366, height: 768 },
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
      acceptDownloads: true,
    });

    await context.addCookies(cookies);
    console.log('✅ Cookies injected.');

  const page = await context.newPage();

  // ========== 4. Điều hướng ==========
  console.log(`🌐 Đang vào ${TARGET_URL}...`);
  await page.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await sleep(2000);

  console.log(`📦 Đang vào trang đơn hàng cần giao...`);
  await page.goto(ORDERS_URL, { waitUntil: 'networkidle', timeout: 60000 });
  await sleep(4000);

  await page.screenshot({ path: path.join(__dirname, '01_orders_page.png'), fullPage: false });
  console.log('📸 01_orders_page.png');

  // ========== 5. Click chọn đơn hàng ==========
  console.log(`🖱️ Đang chọn ${MAX_ORDERS > 0 ? MAX_ORDERS + ' đơn' : 'tất cả'}...`);

  if (MAX_ORDERS <= 0) {
    // Chọn tất cả — click header checkbox
    const selectAllClicked = await page.evaluate(() => {
      const thCheckboxes = document.querySelectorAll('th.col-checkbox, th[class*="col-checkbox"]');
      for (const th of thCheckboxes) {
        const rect = th.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          const label = th.querySelector('label.p-checkbox');
          if (label) { label.click(); return { ok: true, method: 'label.p-checkbox in th' }; }
          th.click(); return { ok: true, method: 'th.col-checkbox' };
        }
      }
      const firstTh = document.querySelector('thead th');
      if (firstTh) { firstTh.click(); return { ok: true, method: 'first thead th' }; }
      return { ok: false };
    });
    console.log(`  Kết quả: ${JSON.stringify(selectAllClicked)}`);
    await sleep(3000);
  } else {
    // Chọn từng đơn theo số lượng
    await page.evaluate((max) => {
      const checkboxes = document.querySelectorAll('td.col-checkbox label.p-checkbox, td[class*="col-checkbox"] label.p-checkbox');
      let count = 0;
      for (const cb of checkboxes) {
        if (count >= max) break;
        cb.click();
        count++;
      }
    }, MAX_ORDERS);
    await sleep(1500);
  }

  // Kiểm tra số lượng đã chọn
  const checkedCount = await page.evaluate(() => {
    return document.querySelectorAll('input[type="checkbox"]:checked').length;
  });
  console.log(`📊 Đã chọn: ${checkedCount} đơn hàng`);

  await page.screenshot({ path: path.join(__dirname, '02_after_select_all.png'), fullPage: false });
  console.log('📸 02_after_select_all.png');

  if (checkedCount === 0) {
    // Kiểm tra xem có đơn hàng nào hiển thị không (có thể checkbox bị disabled hoặc không có đơn)
    const orderRows = await page.evaluate(() => {
      // Đếm số dòng trong bảng đơn hàng (không tính header)
      const rows = document.querySelectorAll('tbody tr, table tbody tr, [class*="table"] tbody tr');
      // Cũng kiểm tra text "Không có đơn hàng" hoặc "No orders"
      const bodyText = document.body.innerText || '';
      const noOrders = bodyText.includes('Không có đơn') || bodyText.includes('No order') || bodyText.includes('không tìm thấy');
      return { rowCount: rows.length, noOrdersText: noOrders, bodySnippet: bodyText.substring(0, 300) };
    });

    if (orderRows.noOrdersText || orderRows.rowCount === 0) {
      console.log('ℹ️ KHÔNG CÓ ĐƠN HÀNG NÀO CẦN GIAO HÔM NAY.');
      console.log('   Đây không phải lỗi — đơn giản là không có đơn ở trạng thái "to_ship".');
    } else if (orderRows.rowCount > 0) {
      console.log(`⚠️ Có ${orderRows.rowCount} dòng trong bảng nhưng không chọn được checkbox.`);
      console.log('   Có thể checkbox đã bị disabled hoặc cần scroll để thấy.');
    } else {
      console.log('⚠️ Không tìm thấy bảng đơn hàng. Có thể trang chưa load hết hoặc cần đăng nhập lại.');
    }

    console.log('👋 Done - Không có đơn hàng để xử lý.');
    return;
  }

  // ========== 6. Đợi popup hiện ra và click "In chứng từ" ==========
  console.log('🔍 Đang tìm nút "In chứng từ"...');

  // Popup thường là 1 thanh bottom bar hoặc 1 modal hiện ra sau khi chọn đơn
  // Tìm button có text "In chứng từ"
  let printDocsClicked = false;

  // Thử tìm bằng text
  const printTexts = ['In chứng từ', 'In chứng từ', 'Print documents', 'Print Documents'];
  for (const txt of printTexts) {
    try {
      const btn = page.locator(`button:has-text("${txt}")`).first();
      if (await btn.count() > 0) {
        const visible = await btn.isVisible({ timeout: 2000 }).catch(() => false);
        if (visible) {
          console.log(`  🎯 Tìm thấy nút "${txt}"`);
          await btn.click({ timeout: 5000 });
          printDocsClicked = true;
          console.log(`  ✅ Đã click "In chứng từ"!`);
          break;
        }
      }
    } catch (e) {}
  }

  // Thử tìm bằng text trong span/div
  if (!printDocsClicked) {
    for (const txt of printTexts) {
      try {
        const el = page.locator(`text="${txt}"`).first();
        if (await el.count() > 0) {
          const visible = await el.isVisible({ timeout: 2000 }).catch(() => false);
          if (visible) {
            console.log(`  🎯 Tìm thấy text "${txt}"`);
            // Click vào button cha
            const parentBtn = el.locator('..');
            await parentBtn.click({ timeout: 5000 });
            printDocsClicked = true;
            console.log(`  ✅ Đã click!`);
            break;
          }
        }
      } catch (e) {}
    }
  }

  // Thử dump các button visible để debug
  if (!printDocsClicked) {
    console.log('  ⚠️ Đang dump tất cả button trên trang...');
    const buttons = await page.evaluate(() => {
      const btns = document.querySelectorAll('button');
      const info = [];
      btns.forEach((b, i) => {
        const rect = b.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          info.push({
            i, top: Math.round(rect.top), left: Math.round(rect.left),
            text: (b.innerText || '').substring(0, 100),
            className: (b.className || '').substring(0, 60),
          });
        }
      });
      return info;
    });
    console.log(`  Tìm thấy ${buttons.length} button visible:`);
    buttons.forEach(b => {
      console.log(`    [${b.i}] top=${b.top} left=${b.left} "${b.text}"`);
    });

    // Thử click button có chứa text "in"
    for (const b of buttons) {
      if (b.text.toLowerCase().includes('in chứng từ') || b.text.toLowerCase().includes('in ')) {
        console.log(`  🎯 Thử click button [${b.i}] "${b.text}"`);
        try {
          await page.locator('button').nth(b.i).click({ timeout: 3000 });
          printDocsClicked = true;
          console.log(`  ✅ Đã click!`);
          break;
        } catch (e) {}
      }
    }
  }

  await sleep(3000);
  await page.screenshot({ path: path.join(__dirname, '03_after_click_print_docs.png'), fullPage: false });
  console.log('📸 03_after_click_print_docs.png');

  // ========== 7. Đợi popup thứ 2 load nội dung ==========
  console.log('⏳ Đợi popup chọn tài liệu in load nội dung...');

  // Đợi modal xuất hiện với nội dung đầy đủ
  await sleep(3000);

  // Chờ cho nội dung popup load (thử đợi text xuất hiện)
  try {
    await page.waitForSelector('text=Nhãn', { timeout: 10000 });
    console.log('  ✅ Tìm thấy text "Nhãn" trong popup');
  } catch (e) {
    console.log('  ⚠️ Không tìm thấy text "Nhãn", thử đợi thêm...');
    await sleep(5000);
  }

  // Dump TOÀN BỘ nội dung popup (có retry nếu popup chưa render kịp)
  console.log('🔍 DUMP TOÀN BỘ NỘI DUNG POPUP:');
  let fullPopupDump = [];
  for (let retry = 0; retry < 5; retry++) {
    fullPopupDump = await page.evaluate(() => {
    // Helper: safe class name (handles SVG elements where className is SVGAnimatedString)
    const getClass = (el) => {
      try {
        const cls = el.className;
        if (typeof cls === 'string') return cls;
        if (cls && cls.baseVal) return cls.baseVal;
        return el.getAttribute('class') || '';
      } catch(e) { return ''; }
    };

    // Tìm modal đang hiển thị
    const modals = document.querySelectorAll('.p-modal, [role="dialog"], .modal, [class*="modal-wrapper"]');
    const results = [];
    modals.forEach((modal, i) => {
      const rect = modal.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0 && modal.innerText.trim().length > 0) {
        // Lấy toàn bộ text
        const allText = modal.innerText.substring(0, 2000);
        // Lấy tất cả label
        const labels = [];
        modal.querySelectorAll('label').forEach(l => {
          const lr = l.getBoundingClientRect();
          labels.push({
            text: (l.innerText || '').trim().substring(0, 100),
            top: Math.round(lr.top),
            left: Math.round(lr.left),
            w: lr.width, h: lr.height,
            className: getClass(l).substring(0, 80),
          });
        });
        // Lấy tất cả input
        const inputs = [];
        modal.querySelectorAll('input').forEach(inp => {
          const ir = inp.getBoundingClientRect();
          inputs.push({
            type: inp.type,
            checked: inp.checked,
            top: Math.round(ir.top),
            left: Math.round(ir.left),
            w: ir.width, h: ir.height,
            name: inp.name || '',
            value: inp.value || '',
          });
        });
        // Lấy tất cả button
        const buttons = [];
        modal.querySelectorAll('button').forEach(b => {
          const br = b.getBoundingClientRect();
          buttons.push({
            text: (b.innerText || '').trim().substring(0, 100),
            top: Math.round(br.top),
            left: Math.round(br.left),
            w: br.width, h: br.height,
            className: getClass(b).substring(0, 80),
          });
        });
        // Lấy tất cả checkbox/radio containers
        const checkables = [];
        modal.querySelectorAll('.p-checkbox, .p-radiobutton, [class*="checkbox"], [class*="radio"]').forEach(el => {
          const er = el.getBoundingClientRect();
          if (er.width > 0 && er.height > 0) {
            checkables.push({
              tag: el.tagName,
              text: (el.innerText || '').trim().substring(0, 100),
              top: Math.round(er.top),
              left: Math.round(er.left),
              className: getClass(el).substring(0, 80),
            });
          }
        });

        results.push({
          index: i,
          className: getClass(modal).substring(0, 100),
          textPreview: allText.substring(0, 1000),
          labels: labels,
          inputs: inputs,
          buttons: buttons,
          checkables: checkables,
        });
      }
    });
    return results;
    });
    if (fullPopupDump.length > 0) break;
    console.log(`  ⏳ Popup chưa render (retry ${retry+1}/5)...`);
    await sleep(2000);
  }

  console.log(`  Popups có nội dung: ${fullPopupDump.length}`);
  for (const popup of fullPopupDump) {
    console.log(`\n  === POPUP [${popup.index}] ${popup.className} ===`);
    console.log(`  Text preview: "${popup.textPreview.substring(0, 500)}"`);
    console.log(`  Labels (${popup.labels.length}):`);
    popup.labels.forEach(l => console.log(`    - "${l.text}" top=${l.top} left=${l.left} ${l.w}x${l.h} class="${l.className}"`));
    console.log(`  Inputs (${popup.inputs.length}):`);
    popup.inputs.forEach(inp => console.log(`    - type=${inp.type} checked=${inp.checked} top=${inp.top} name="${inp.name}"`));
    console.log(`  Buttons (${popup.buttons.length}):`);
    popup.buttons.forEach(b => console.log(`    - "${b.text}" top=${b.top} left=${b.left} ${b.w}x${b.h} class="${b.className}"`));
    console.log(`  Checkables (${popup.checkables.length}):`);
    popup.checkables.forEach(c => console.log(`    - [${c.tag}] "${c.text}" top=${c.top} class="${c.className}"`));
  }

  await page.screenshot({ path: path.join(__dirname, '04_print_options_popup.png'), fullPage: false });
  console.log('📸 04_print_options_popup.png');

  // ========== 8. Chọn cả 3 loại: A6 vận chuyển (giữ) + A6 đóng gói + A4 ==========
  console.log('🖱️ Đang chọn cả 3 loại tài liệu...');

  // Chỉ xử lý popup có buttons (popup chính, bỏ qua wrapper)
  const mainPopup = fullPopupDump.find(p => p.buttons && p.buttons.length > 0) || fullPopupDump[fullPopupDump.length - 1];
  if (!mainPopup) {
    console.log('⚠️ Popup không có nội dung, thử click trực tiếp...');
    // Thử tìm label trong page thay vì popup
  } else {
    console.log(`  Chọn popup chính (có ${mainPopup.buttons?.length || 0} buttons)`);
  }

  const labels = mainPopup.labels;
  const inputs = mainPopup.inputs;

  let actions = [];

  // Check thêm A4 (A6 vận chuyển đã checked sẵn, bỏ qua A6 đóng gói)
  const toCheck = ['Danh sách lấy hàng'];
  for (const labelText of toCheck) {
    try {
      const loc = page.locator('label.p-checkbox').filter({ hasText: labelText });
      const cnt = await loc.count();
      if (cnt > 0) {
        const first = loc.first();
        const isChecked = await first.evaluate(el => el.classList.contains('p-checkbox-checked'));
        if (!isChecked) {
          await first.click({ timeout: 3000, force: true });
          actions.push(`checked_${labelText}`);
          console.log(`  ✅ Đã chọn: ${labelText}`);
          await sleep(500);
        } else {
          console.log(`  ⏭ Đã checked sẵn: ${labelText}`);
        }
      } else {
        console.log(`  ⚠️ Không tìm thấy label: ${labelText}`);
      }
    } catch (e) {
      console.log(`  ❌ Lỗi chọn ${labelText}: ${e.message}`);
    }
  }

  // Verify trạng thái sau khi click
  const verifyState = await page.evaluate(() => {
    const inputs = document.querySelectorAll('.p-modal input[type="checkbox"]');
    const states = [];
    inputs.forEach((inp, i) => {
      const label = inp.closest('label');
      states.push({
        i,
        checked: inp.checked,
        text: label ? label.innerText.trim().substring(0, 60) : '',
      });
    });
    return states;
  });
  console.log(`  📊 Trạng thái sau khi chọn: ${JSON.stringify(verifyState)}`);

  console.log(`  📊 Actions đã thực hiện: ${JSON.stringify(actions)}`);

  await sleep(2000);
  await page.screenshot({ path: path.join(__dirname, '05_after_select_A4.png'), fullPage: false });
  console.log('📸 05_after_select_A4.png');

  // ========== 9. Click nút "In" và bắt tab mới ==========
  console.log('🖱️ Đang click nút "In"...');

  // Bắt sự kiện page mới (popup/tab) TRƯỚC KHI click
  const newPagePromise = context.waitForEvent('page', { timeout: 30000 }).catch(() => null);
  // Cũng bắt download
  const downloadPromise = page.waitForEvent('download', { timeout: 30000 }).catch(() => null);

  // Click nút "In" - dùng selector chính xác từ dump
  let printClicked = false;
  try {
    // Dùng class chính xác: p-btn p-btn-primary p-btn-size-large
    const inBtn = page.locator('button.p-btn-primary:has-text("In")').first();
    if (await inBtn.count() > 0 && await inBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
      await inBtn.click({ timeout: 5000 });
      printClicked = true;
      console.log('  ✅ Đã click nút "In" (p-btn-primary)!');
    }
  } catch (e) {
    console.log(`  ❌ Lỗi click In: ${e.message}`);
  }

  // Fallback: click bằng text chính xác
  if (!printClicked) {
    try {
      // Nút In trong popup (text chính xác là "In")
      const exactInBtn = page.locator('button').filter({ hasText: /^In$/ }).first();
      if (await exactInBtn.count() > 0 && await exactInBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
        await exactInBtn.click({ timeout: 5000 });
        printClicked = true;
        console.log('  ✅ Đã click nút "In" (exact text)!');
      }
    } catch (e) {}
  }

  // Fallback: click bằng toạ độ
  if (!printClicked) {
    const inBtnCoord = mainPopup.buttons.find(b => b.text === 'In');
    if (inBtnCoord) {
      console.log(`  🎯 Click nút In tại (${inBtnCoord.left + 20}, ${inBtnCoord.top + 20})`);
      await page.mouse.click(inBtnCoord.left + 20, inBtnCoord.top + 20);
      printClicked = true;
      console.log('  ✅ Đã click bằng tọa độ!');
    }
  }

  await sleep(2000);

  // ========== Xử lý popup xác nhận "Tiếp tục in" (xuất hiện khi chọn nhiều loại) ==========
  console.log('🔍 Kiểm tra popup xác nhận...');
  let continueBtn = null;
  for (let i = 0; i < 15; i++) {
    const btns = page.locator('button:has-text("Tiếp tục in"), button:has-text("Continue"), button:has-text("Xác nhận")');
    const cnt = await btns.count();
    if (cnt > 0) {
      for (let j = 0; j < cnt; j++) {
        const b = btns.nth(j);
        if (await b.isVisible({ timeout: 500 }).catch(() => false)) {
          continueBtn = b;
          break;
        }
      }
    }
    if (continueBtn) break;
    await sleep(1000);
  }

  if (continueBtn) {
    console.log('  📥 Popup 1: click "Tiếp tục in"...');
    await continueBtn.click({ timeout: 5000 });
    console.log('  ✅ Đã click "Tiếp tục in"');
    await sleep(3000);

    // Đợi popup 2: nút "Tải xuống tất cả"
    console.log('  🔍 Đợi popup 2: "Tải xuống tất cả"...');
    let downloadAllBtn = null;
    for (let i = 0; i < 20; i++) {
      const btns = page.locator('button:has-text("Tải xuống tất cả"), button:has-text("Download all"), button:has-text("Tải xuống")');
      const cnt = await btns.count();
      for (let j = 0; j < cnt; j++) {
        const b = btns.nth(j);
        if (await b.isVisible({ timeout: 500 }).catch(() => false)) {
          downloadAllBtn = b;
          break;
        }
      }
      if (downloadAllBtn) break;
      await sleep(1000);
    }

    if (downloadAllBtn) {
      console.log('  📥 Popup 2: click "Tải xuống tất cả"...');
      // Bắt download event
      const dlPromise = page.waitForEvent('download', { timeout: 60000 }).catch(() => null);
      await downloadAllBtn.click({ timeout: 5000 });
      console.log('  ✅ Đã click "Tải xuống tất cả"');
      await sleep(3000);

      const dl = await dlPromise;
      if (dl) {
        const suggested = dl.suggestedFilename() || `download_${TS}.zip`;
        const savePath = path.join(__dirname, suggested);
        await dl.saveAs(savePath);
        console.log(`💾 ĐÃ TẢI VỀ: ${savePath}`);
        fs.copyFileSync(savePath, PDF_RAW);
      } else {
        // Fallback: check tab mới
        console.log('  ⚠️ Không bắt được download, check tab...');
        await sleep(3000);
        for (const pg of context.pages()) {
          if (pg.url().includes('/easesafe/')) {
            const resp = await page.request.get(pg.url(), { timeout: 30000 });
            if (resp.ok()) {
              fs.writeFileSync(PDF_RAW, await resp.body());
              console.log(`💾 ĐÃ TẢI VỀ: ${PDF_NAME_RAW} (${resp.body().length} bytes)`);
            }
          }
        }
      }
    } else {
      console.log('  ⚠️ Không thấy nút "Tải xuống tất cả"');
    }
  }

  await page.screenshot({ path: path.join(__dirname, '06_after_click_print.png'), fullPage: false });
  console.log('📸 06_after_click_print.png');

  // ========== 10. Đợi tab mới hoặc download ==========
  console.log('⏳ Đang đợi tab mới mở ra hoặc file download...');

  // Đợi 1 trong 2 sự kiện
  const result = await Promise.race([
    newPagePromise.then(p => ({ type: 'page', page: p })),
    downloadPromise.then(d => ({ type: 'download', download: d })),
    sleep(15000).then(() => ({ type: 'timeout' })),
  ]);

  if (result.type === 'page' && result.page) {
    const newPage = result.page;
    console.log('✅ Tab mới đã mở!');
    await sleep(5000);

    // Đợi trang load
    await newPage.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await sleep(3000);

    const newUrl = newPage.url();
    console.log(`🔗 URL tab mới: ${newUrl}`);

    await newPage.screenshot({ path: path.join(__dirname, '07_new_tab.png'), fullPage: false });
    console.log('📸 07_new_tab.png');

    // ========== TÌM VÀ CLICK NÚT TẢI XUỐNG TRONG TAB MỚI ==========
    console.log('🔍 Đang tìm nút tải xuống trong tab mới...');

    // Bắt sự kiện download TRƯỚC KHI click
    const dlPromise = newPage.waitForEvent('download', { timeout: 30000 }).catch(() => null);

    // Dump tất cả button/link trong tab mới
    const tabElements = await newPage.evaluate(() => {
      const items = [];

      // Tìm buttons
      document.querySelectorAll('button').forEach((b, i) => {
        const rect = b.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          items.push({
            type: 'button',
            i,
            text: (b.innerText || '').trim().substring(0, 100),
            top: Math.round(rect.top),
            left: Math.round(rect.left),
            w: Math.round(rect.width),
            h: Math.round(rect.height),
            className: (b.className && typeof b.className === 'string') ? b.className.substring(0, 80) : '',
          });
        }
      });

      // Tìm links
      document.querySelectorAll('a[href]').forEach((a, i) => {
        const rect = a.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          items.push({
            type: 'link',
            i,
            text: (a.innerText || '').trim().substring(0, 100),
            href: (a.href || '').substring(0, 200),
            top: Math.round(rect.top),
            left: Math.round(rect.left),
          });
        }
      });

      // Tìm elements có icon download
      document.querySelectorAll('[class*="download"], [class*="Download"], svg').forEach((el, i) => {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0 && rect.width < 100) {
          const parent = el.closest('button, a');
          if (parent) {
            items.push({
              type: 'download-icon-parent',
              i,
              parentTag: parent.tagName,
              parentText: (parent.innerText || '').trim().substring(0, 80),
              top: Math.round(rect.top),
              className: (el.className && typeof el.className === 'string') ? el.className.substring(0, 60) : '',
            });
          }
        }
      });

      return items;
    });

    console.log(`  Elements trong tab mới:`);
    tabElements.forEach(el => {
      if (el.type === 'button') {
        console.log(`    [BTN ${el.i}] "${el.text}" top=${el.top} left=${el.left} size=${el.w}x${el.h} class="${el.className}"`);
      } else if (el.type === 'link') {
        console.log(`    [LINK ${el.i}] "${el.text}" href="${el.href}" top=${el.top}`);
      } else if (el.type === 'download-icon-parent') {
        console.log(`    [ICON ${el.i}] <${el.parentTag}> "${el.parentText}" class="${el.className}" top=${el.top}`);
      }
    });

    // Tìm và click nút tải xuống
    let downloadClicked = false;

    // Các text pattern của nút tải xuống
    const downloadTexts = ['Tải xuống', 'Download', 'Tải về', 'Tải', 'Save', 'Export', 'Xuất'];

    for (const txt of downloadTexts) {
      try {
        const btn = newPage.locator(`button:has-text("${txt}")`).first();
        if (await btn.count() > 0 && await btn.isVisible({ timeout: 1000 }).catch(() => false)) {
          console.log(`  🎯 Click nút "${txt}"...`);
          await btn.click({ timeout: 5000 });
          downloadClicked = true;
          console.log(`  ✅ Đã click!`);
          break;
        }
      } catch (e) {}
    }

    // Thử click link download
    if (!downloadClicked) {
      for (const el of tabElements) {
        if (el.type === 'link' &&
            (el.text.toLowerCase().includes('tải') ||
             el.text.toLowerCase().includes('download') ||
             el.href.includes('download'))) {
          try {
            console.log(`  🎯 Click link: "${el.text}"`);
            await newPage.locator(`a:has-text("${el.text}")`).first().click({ timeout: 3000 });
            downloadClicked = true;
            console.log(`  ✅ Đã click!`);
            break;
          } catch (e) {}
        }
      }
    }

    // Thử click vào icon download parent
    if (!downloadClicked) {
      const downloadIcons = tabElements.filter(el => el.type === 'download-icon-parent');
      if (downloadIcons.length > 0) {
        const icon = downloadIcons[0];
        console.log(`  🎯 Click icon download parent: <${icon.parentTag}> "${icon.parentText}"`);
        try {
          await newPage.mouse.click(icon.left + 10, icon.top + 10);
          downloadClicked = true;
          console.log(`  ✅ Đã click icon!`);
        } catch (e) {}
      }
    }

    // Thử dùng evaluate để click tất cả button có download icon
    if (!downloadClicked) {
      const jsClicked = await newPage.evaluate(() => {
        const btns = document.querySelectorAll('button');
        for (const b of btns) {
          const text = (b.innerText || '').toLowerCase();
          const cls = ((b.className || '') + '').toLowerCase();
          const html = (b.innerHTML || '').toLowerCase();
          if (text.includes('tải') || text.includes('download') ||
              cls.includes('download') || html.includes('download') ||
              html.includes('tải')) {
            b.click();
            return { clicked: true, text: b.innerText.trim().substring(0, 50) };
          }
        }
        // Thử click link
        const links = document.querySelectorAll('a');
        for (const a of links) {
          const text = (a.innerText || '').toLowerCase();
          if (text.includes('tải') || text.includes('download')) {
            a.click();
            return { clicked: true, text: a.innerText.trim().substring(0, 50) };
          }
        }
        return { clicked: false };
      });
      if (jsClicked.clicked) {
        downloadClicked = true;
        console.log(`  ✅ JS click: "${jsClicked.text}"`);
      }
    }

    // ========== ĐỢI FILE DOWNLOAD ==========
    if (downloadClicked) {
      console.log('⏳ Đang đợi file tải về...');
      const download = await dlPromise;
      if (download) {
        const fileName = download.suggestedFilename() || `danh_sach_lay_hang_${Date.now()}`;
        const savePath = path.join(__dirname, fileName);
        await download.saveAs(savePath);
        console.log(`💾 ĐÃ TẢI VỀ: ${savePath}`);
      } else {
        console.log('⚠️ Timeout đợi download sau 30s.');
        // Fallback: in trang ra PDF
        console.log('  📄 Fallback: in trang thành PDF...');
        await newPage.pdf({
          path: PDF_RAW,
          format: 'A4',
          printBackground: true,
        });
        console.log(`  💾 Đã lưu ${PDF_NAME_RAW} (từ page.pdf)`);
      }
    } else {
      // Tab hiển thị PDF trực tiếp (Chrome PDF viewer) - không có DOM elements
      // Fetch trực tiếp signed URL để lấy raw PDF
      console.log('📥 Tab hiển thị PDF trực tiếp, fetch URL để lấy file gốc...');
      let downloaded = false;

      // Cách 1: Fetch bằng context của newPage
      try {
        const response = await newPage.request.get(newUrl, { timeout: 30000 });
        console.log(`  Status: ${response.status()}, Content-Type: ${response.headers()['content-type']}`);
        const buffer = await response.body();
        console.log(`  Body: ${buffer.length} bytes, header: ${buffer.slice(0, 20).toString('utf-8')}`);
        if (buffer.length > 1000) {
          fs.writeFileSync(PDF_RAW, buffer);
          console.log(`💾 ĐÃ TẢI VỀ: ${PDF_NAME_RAW} (${buffer.length} bytes)`);
          downloaded = true;
        }
      } catch (e) {
        console.log(`  ❌ Lỗi fetch newPage: ${e.message}`);
      }

      // Cách 2: Fetch bằng page chính
      if (!downloaded) {
        try {
          const response = await page.request.get(newUrl, { timeout: 30000 });
          console.log(`  Status: ${response.status()}, Content-Type: ${response.headers()['content-type']}`);
          const buffer = await response.body();
          console.log(`  Body: ${buffer.length} bytes`);
          if (buffer.length > 1000) {
            fs.writeFileSync(PDF_RAW, buffer);
            console.log(`💾 ĐÃ TẢI VỀ: ${PDF_NAME_RAW} (${buffer.length} bytes)`);
            downloaded = true;
          }
        } catch (e) {
          console.log(`  ❌ Lỗi fetch page: ${e.message}`);
        }
      }

      // Cách 3: Fallback cuối cùng - page.pdf()
      if (!downloaded) {
        console.log('  📄 Fallback: in trang ra PDF...');
        await newPage.pdf({
          path: PDF_RAW,
          format: 'A4',
          printBackground: true,
        });
        console.log('  💾 Đã lưu danh_sach_lay_hang.pdf (page.pdf)');
      }
    }

  } else if (result.type === 'download' && result.download) {
    const download = result.download;
    const fileName = download.suggestedFilename() || `download_${Date.now()}.pdf`;
    const savePath = path.join(__dirname, fileName);
    await download.saveAs(savePath);
    console.log(`💾 Đã tải file (download): ${savePath}`);

  } else {
    console.log('⚠️ Timeout - không có tab mới hoặc download sau 15s.');
    console.log(`  URL hiện tại: ${page.url()}`);

    // Kiểm tra xem popup còn không
    const popupStillThere = await page.evaluate(() => {
      const modals = document.querySelectorAll('.p-modal');
      let visibleCount = 0;
      modals.forEach(m => {
        if (m.getBoundingClientRect().width > 0) visibleCount++;
      });
      return visibleCount;
    });
    console.log(`  Popup còn hiển thị: ${popupStillThere}`);

    // Chụp màn hình cuối cùng
    await page.screenshot({ path: path.join(__dirname, '07_no_new_tab.png'), fullPage: true });
    console.log('📸 07_no_new_tab.png');
  }

  await page.screenshot({ path: path.join(__dirname, '08_final.png'), fullPage: false });
  console.log('📸 08_final.png');

  // ============================================================
  // PIPELINE: Nếu có file PDF đã tải → xử lý qua bill_calculate
  // ============================================================
  const downloadedPdf = PDF_RAW;
  if (fs.existsSync(downloadedPdf) && fs.statSync(downloadedPdf).size > 1000) {
    const result = await processWithCalculator(downloadedPdf, PDF_NAME_RAW);

    console.log('\n' + '='.repeat(50));
    console.log('🏁 PIPELINE HOÀN THÀNH!');
    console.log('='.repeat(50));
    console.log(`📥 File gốc từ TikTok: ${downloadedPdf}`);
    if (result && result.excelPath) {
      console.log(`📊 Excel kết quả:     ${result.excelPath}`);
    }
    if (result && result.pdfPath) {
      console.log(`📕 PDF báo cáo:       ${result.pdfPath}`);
    }
    console.log('='.repeat(50));
  } else {
    console.log('\n⚠️ Không có file PDF hợp lệ để xử lý.');
  }

  } catch (err) {
    // Bắt mọi lỗi trong quá trình automation
    console.log('\n' + '='.repeat(50));
    console.log('❌ LỖI TRONG QUÁ TRÌNH TỰ ĐỘNG:');
    console.log('='.repeat(50));
    console.log(`   ${err.message}`);
    console.log('='.repeat(50));

    // Chụp màn hình lỗi nếu page còn tồn tại
    try {
      const pages = context?.pages?.();
      if (pages && pages.length > 0) {
        await pages[0].screenshot({
          path: path.join(__dirname, 'error_screenshot.png'),
          fullPage: true,
        });
        console.log('📸 Đã chụp error_screenshot.png');
      }
    } catch {}
  } finally {
    // LUÔN đóng browser - dù thành công hay thất bại
    console.log('\n🔚 Đóng trình duyệt...');
    if (browser) {
      try { await browser.close(); } catch {}
    }
    console.log('👋 Done!');
  }
}

main().catch((err) => {
  console.error('❌ Lỗi:', err.message);
  console.error(err.stack);
  process.exit(1);
});
