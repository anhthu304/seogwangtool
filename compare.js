/*
 * Đối chiếu BOOK (file xuất ra) với Phiếu xuất kho (PDF) và sheet MẪU 27 của INVOICE.
 * Thuần logic, không đụng tới giao diện -> chạy được cả trong trình duyệt lẫn Node (để test).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PxkCompare = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  // ───────────── tiện ích ─────────────
  const stripVi = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D');
  // Khoá so sánh tên: giữ dấu tiếng Việt, bỏ mọi ký tự không phải chữ/số (dấu cách, gạch nối, dấu phẩy...)
  const keyName = s => String(s == null ? '' : s).normalize('NFC').toUpperCase().replace(/[^\p{L}\p{N}]/gu, '');
  const keyCode = s => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const isNumTok = s => /^[\d.,]+$/.test(s) && /\d/.test(s);
  // Số kiểu Việt: 8.700.048 | 294,00 | 29.592
  const parseVnNum = s => {
    const t = String(s).trim().replace(/\./g, '').replace(/,/g, '.');
    const n = Number(t);
    return isNaN(n) ? null : n;
  };
  const toNum = v => {
    if (typeof v === 'number') return v;
    if (v == null || v === '') return null;
    const n = Number(String(v).replace(/[^\d.-]/g, ''));
    return isNaN(n) ? null : n;
  };
  const stripLeadZeros = s => String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, '');

  const PHI_GC_RE = /PHÍ\s*GIA\s*CÔNG/giu;
  const hasPhiGC = name => { PHI_GC_RE.lastIndex = 0; return PHI_GC_RE.test(String(name).normalize('NFC').toUpperCase()); };
  const removePhiGC = name => String(name).normalize('NFC').toUpperCase().replace(PHI_GC_RE, ' ');
  // Khi so tên hàng, bỏ qua: chữ "phí gia công", cụm "BỘ PHẬN CỦA", và chữ "VN" đứng riêng (không nằm trong từ khác)
  const NOISE_VN_RE = /(^|[^\p{L}\p{N}])VN(?![\p{L}\p{N}])/gu;
  const nameKeyN = name => keyName(removePhiGC(name).replace(/BỘ\s*PHẬN\s*CỦA/gu, ' ').replace(NOISE_VN_RE, '$1 '));

  function fmtDate(d, m, y) { return String(d).padStart(2, '0') + '/' + String(m).padStart(2, '0') + '/' + y; }
  // Chuẩn hoá nhiều kiểu ngày về dd/mm/yyyy
  function normDate(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') { // số serial Excel
      const dt = new Date(Math.round((v - 25569) * 86400000));
      return fmtDate(dt.getUTCDate(), dt.getUTCMonth() + 1, dt.getUTCFullYear());
    }
    if (v instanceof Date) return fmtDate(v.getDate(), v.getMonth() + 1, v.getFullYear());
    const s = String(v).trim();
    let m = s.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
    if (m) return fmtDate(m[1], m[2], m[3]);
    m = s.match(/(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/);
    if (m) return fmtDate(m[3], m[2], m[1]);
    return s;
  }

  // ───────────── đọc PDF phiếu xuất kho ─────────────
  // pages: mảng theo trang, mỗi trang là mảng {s, x, y, w}  (s = chữ, x/y = toạ độ, w = chiều rộng)
  function groupLines(items, tol) {
    const sorted = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    const lines = [];
    for (const it of sorted) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - it.y) <= tol) last.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    }
    lines.forEach(l => l.items.sort((a, b) => a.x - b.x));
    return lines;
  }
  const joinItems = items => items.reduce((acc, it) => acc === '' ? it.s : (/-$/.test(acc) ? acc + it.s : acc + ' ' + it.s), '');

  function parsePxk(pages, fileName) {
    const res = { file: fileName || '', soPhieu: '', ngay: '', tongGia: null, rows: [], warnings: [] };

    // Văn bản toàn phiếu để bắt số phiếu / ngày
    const fullText = pages.map(items => groupLines(items.filter(i => i.s.trim()), 3).map(l => joinItems(l.items)).join('\n')).join('\n');
    let m = fullText.match(/Số\s*\(No\.?\)\s*:?\s*(\d+)/);
    if (m) res.soPhieu = m[1];
    else { const f = String(fileName || '').match(/(\d{5,})\D*$/); if (f) res.soPhieu = f[1]; else res.warnings.push('Không tìm thấy số phiếu'); }
    // Ngày phiếu: lần xuất hiện đầu tiên (nằm ngay dưới tiêu đề PHIẾU XUẤT KHO), khác với ngày của lệnh điều động nằm thấp hơn
    m = fullText.match(/Ngày\s*\(Date\)\s*(\d{1,2})\s*tháng\s*\(month\)\s*(\d{1,2})\s*năm\s*\(year\)\s*(\d{4})/);
    if (m) res.ngay = fmtDate(m[1], m[2], m[3]); else res.warnings.push('Không tìm thấy ngày phiếu');

    let sumAmount = 0;
    pages.forEach((raw, pi) => {
      const items = raw.filter(i => i.s.trim() && !/^\d+\/\d+$/.test(i.s.trim()));
      const hStt = items.find(i => i.s.trim() === 'STT');
      if (!hStt) return; // trang không có bảng hàng
      const sttCx = hStt.x + hStt.w / 2;
      const endCells = items.filter(i => /(output\)|input\))$/.test(i.s.trim()));
      const headerBottom = endCells.length ? Math.min(...endCells.map(i => i.y)) : hStt.y - 40;
      const tot = items.find(i => /^Tổng cộng/.test(i.s.trim()));
      if (tot) {
        const sameLine = items.filter(i => Math.abs(i.y - tot.y) <= 3 && isNumTok(i.s.trim()));
        if (sameLine.length) { sameLine.sort((a, b) => a.x - b.x); res.tongGia = parseVnNum(sameLine[sameLine.length - 1].s.trim()); }
      }
      const yLow = tot ? tot.y + 2 : -Infinity;
      const data = items.filter(i => i.y < headerBottom - 1 && i.y > yLow);

      const stts = data.filter(i => /^\d+$/.test(i.s.trim()) && Math.abs(i.x + i.w / 2 - sttCx) <= 12).sort((a, b) => b.y - a.y);
      if (!stts.length) return;
      const rowsRaw = stts.map(s => ({ stt: s.s.trim(), y: s.y, items: [] }));
      for (const it of data) {
        if (stts.includes(it)) continue;
        let best = rowsRaw[0], bd = Infinity;
        for (const r of rowsRaw) { const d = Math.abs(r.y - it.y); if (d < bd) { bd = d; best = r; } }
        best.items.push(it);
      }
      for (const r of rowsRaw) {
        const its = r.items.slice().sort((a, b) => a.x - b.x);
        const nums = [];
        while (its.length && isNumTok(its[its.length - 1].s.trim()) && nums.length < 4) nums.unshift(its.pop());
        if (nums.length < 3) { res.warnings.push(`Trang ${pi + 1}, dòng STT ${r.stt}: không đọc đủ số lượng/đơn giá/thành tiền`); continue; }
        const amount = parseVnNum(nums[nums.length - 1].s.trim());
        const price = parseVnNum(nums[nums.length - 2].s.trim());
        const qty = parseVnNum(nums[0].s.trim());
        const unit = its.length ? its.pop().s.trim() : '';
        const code = its.length ? its.pop().s.trim() : '';
        const name = groupLines(its, 3).map(l => joinItems(l.items)).reduce((a, l) => a === '' ? l : (/-$/.test(a) ? a + l : a + ' ' + l), '');
        res.rows.push({ stt: r.stt, code, name, unit, qty, price, amount });
        sumAmount += amount || 0;
      }
    });
    res.sumAmount = sumAmount;
    res.tongSL = res.rows.reduce((s, r) => s + (r.qty || 0), 0);
    if (res.tongGia == null) res.warnings.push('Không tìm thấy dòng Tổng cộng');
    if (!res.rows.length) res.warnings.push('Không đọc được dòng hàng nào (PDF dạng ảnh scan?)');
    return res;
  }

  // Một file PDF có thể gộp NHIỀU phiếu (mỗi phiếu có tiêu đề riêng, trang đầu có "Số (No.): ..."). Tách ra từng phiếu.
  function parsePxkFile(pages, fileName) {
    const starts = [];
    pages.forEach((items, i) => {
      const t = groupLines(items.filter(x => x.s.trim()), 3).map(l => joinItems(l.items)).join('\n');
      if (/Số\s*\(No\.?\)\s*:?\s*\d+/.test(t)) starts.push(i);
    });
    if (starts.length <= 1) return [parsePxk(pages, fileName)];
    starts[0] = 0; // trang nào nằm trước phiếu đầu tiên thì tính vào phiếu đầu
    return starts.map((s, k) => {
      const p = parsePxk(pages.slice(s, k + 1 < starts.length ? starts[k + 1] : pages.length), fileName);
      p.file = fileName + ' (phiếu ' + p.soPhieu + ')';
      return p;
    });
  }

  // ───────────── đọc sheet MẪU 27 ─────────────
  // rows: mảng 2 chiều (sheet_to_json header:1)
  function parseMau27(rows) {
    const out = { phieu: [], total: null, error: '' };
    let hr = -1, col = {};
    for (let r = 0; r < rows.length && hr < 0; r++) {
      const cells = (rows[r] || []).map(c => keyName(stripVi(c)));
      const iSo = cells.findIndex(c => c.includes('SOPHIEU'));
      if (iSo >= 0) {
        hr = r;
        col = {
          so: iSo,
          ngay: cells.findIndex(c => c.includes('NGAYPHIEU')),
          sl: cells.findIndex(c => c.includes('SOLUONG')),
          gia: cells.findIndex(c => c.includes('TRIGIA'))
        };
      }
    }
    if (hr < 0) { out.error = 'Không tìm thấy dòng tiêu đề (Số phiếu…) trong sheet Mẫu 27'; return out; }
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const first = keyName(stripVi(row[0] != null && row[0] !== '' ? row[0] : row[1]));
      const firstAny = (row || []).map(c => keyName(stripVi(c))).find(c => c) || '';
      if (/^TONG/.test(first) || (/^TONG/.test(firstAny) && !row[col.so])) {
        out.total = { sl: col.sl >= 0 ? toNum(row[col.sl]) : null, gia: col.gia >= 0 ? toNum(row[col.gia]) : null };
        break;
      }
      const so = row[col.so];
      if (so == null || String(so).trim() === '') continue;
      out.phieu.push({ so: String(so).trim(), ngay: normDate(col.ngay >= 0 ? row[col.ngay] : ''), sl: col.sl >= 0 ? toNum(row[col.sl]) : null, gia: col.gia >= 0 ? toNum(row[col.gia]) : null });
    }
    return out;
  }

  // ───────────── đối chiếu ─────────────
  const near = (a, b, tol) => a != null && b != null && Math.abs(a - b) <= tol;

  // ───────────── đọc sheet INVOICE (INV) ─────────────
  // rows: mảng 2 chiều (sheet_to_json header:1)
  // Giá: ưu tiên nhóm cột "Processing fee"; sheet không có nhóm này (hoặc dòng đó để trống) thì lấy nhóm "DAP VALUE".
  function parseInvoice(rows) {
    const out = { lines: [], error: '', usedDap: 0, hasProcessing: false };
    let hr = -1;
    for (let r = 0; r < rows.length && hr < 0; r++) {
      const cells = (rows[r] || []).map(c => keyName(stripVi(c)));
      if (cells.some(c => c.includes('ITEMNO')) && cells.some(c => c.includes('DESCRIPTION'))) hr = r;
    }
    if (hr < 0) { out.error = 'Không tìm thấy dòng tiêu đề (Item No., Description…) trong sheet INVOICE'; return out; }
    const head = (rows[hr] || []).map(c => keyName(stripVi(c)));
    const col = { ma: head.findIndex(c => c.includes('ITEMNO')), ten: head.findIndex(c => c.includes('DESCRIPTION')), sl: head.findIndex(c => c.includes('QTY') || c.includes('QUANTITY')) };
    // Tiêu đề U/price, Amount có thể nằm cùng dòng "Item No." hoặc 1-2 dòng bên dưới (tuỳ mẫu invoice) -> quét cả 3 dòng
    const priceCols = [], amountCols = [];
    for (let r = hr; r <= hr + 2 && r < rows.length; r++) (rows[r] || []).forEach((c, i) => {
      const k = keyName(stripVi(c));
      if (k.includes('PRICE')) { if (!priceCols.includes(i)) priceCols.push(i); }
      else if (k.includes('AMOUNT')) { if (!amountCols.includes(i)) amountCols.push(i); }
    });
    priceCols.sort((a, b) => a - b); amountCols.sort((a, b) => a - b);
    // nhãn nhóm (Processing fee / DAP VALUE) nằm cùng dòng Item No. hoặc phía trên
    const groups = [];
    for (let r = Math.max(0, hr - 3); r <= hr + 1; r++) (rows[r] || []).forEach((c, i) => {
      const k = keyName(stripVi(c));
      if (k.includes('PROCESSING')) groups.push({ type: 'proc', col: i });
      else if (k.startsWith('DAP')) groups.push({ type: 'dap', col: i });
    });
    groups.sort((a, b) => a.col - b.col);
    const groupOf = i => { let g = null; groups.forEach(x => { if (x.col <= i) g = x.type; }); return g; };
    const pick = (list, type) => list.find(i => groupOf(i) === type);
    let proc = { price: pick(priceCols, 'proc'), amount: pick(amountCols, 'proc') };
    let dap = { price: pick(priceCols, 'dap'), amount: pick(amountCols, 'dap') };
    if (!groups.length) { dap = { price: priceCols[priceCols.length - 1], amount: amountCols[amountCols.length - 1] }; proc = {}; } // không có nhãn nhóm -> 1 bộ giá duy nhất
    out.hasProcessing = proc.price != null && proc.amount != null;
    if (col.ma < 0 || col.sl < 0 || (!out.hasProcessing && (dap.price == null || dap.amount == null))) { out.error = 'Không xác định được các cột Item No. / Q\'ty / U/price / Amount trong sheet INVOICE'; return out; }
    // cột HS code: tiêu đề "Mã hscode" nằm trong vài dòng quanh dòng tiêu đề
    col.hs = -1;
    for (let r = Math.max(0, hr - 1); r <= hr + 3 && col.hs < 0; r++) col.hs = (rows[r] || []).findIndex(c => keyName(stripVi(c)).includes('HSCODE'));
    out.cols = { proc, dap, ma: col.ma, ten: col.ten, sl: col.sl, hs: col.hs };

    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const code = row[col.ma] == null ? '' : String(row[col.ma]).trim();
      const qty = toNum(row[col.sl]);
      if (!code || code === '0' || /TOTAL/i.test(code) || !qty || qty <= 0) continue;
      const has = v => toNum(v) != null && toNum(v) !== 0;
      let price, amount, src;
      if (out.hasProcessing && (has(row[proc.price]) || has(row[proc.amount]))) { price = toNum(row[proc.price]); amount = toNum(row[proc.amount]); src = 'Processing fee'; }
      else if (dap.price != null && dap.amount != null) { price = toNum(row[dap.price]); amount = toNum(row[dap.amount]); src = 'DAP VALUE'; out.usedDap++; }
      else continue;
      // HS code: ưu tiên cột có tiêu đề "Mã hscode"; không được thì thử cột K rồi cột M. HS hợp lệ = đúng 8 chữ số.
      let hs = '';
      for (const ci of [col.hs, 10, 12]) {
        if (ci == null || ci < 0 || row[ci] == null) continue;
        const d = String(row[ci]).replace(/\D/g, '');
        if (/^\d{8}$/.test(d)) { hs = d; break; }
      }
      out.lines.push({ row: r + 1, ma: code, ten: row[col.ten] == null ? '' : String(row[col.ten]), qty, price, amount, src, hs });
    }
    if (!out.lines.length) out.error = 'Không đọc được dòng hàng nào trong sheet INVOICE';
    return out;
  }

  // ───────────── đọc file BCCT (cùng định dạng với trang NPL đã lưu) ─────────────
  // arr: mảng 2 chiều (sheet_to_json header:1) của Sheet1. Kết quả: { map: { <mã>: {code, hs, name, all:[{hs,name}]} }, count }
  const BCCT_VER = 2; // tăng khi đổi cách đọc BCCT -> bản đã lưu cũ hơn sẽ bị yêu cầu chọn lại
  const normHS = s => String(s == null ? '' : s).trim().replace(/\.0+$/, '').replace(/[\s.]/g, '');
  // Ô "Tên hàng" của BCCT có nhiều dạng:
  //   "MÃ - TÊN. Hàng mới 100%#&VN"   (đuôi #&VN = xuất xứ)
  //   "MÃ#&TÊN#&VN"  |  "MÃ#&TÊN"  |  "TÊN"
  // -> bỏ đuôi xuất xứ, bỏ mã ở đầu, bỏ "Hàng mới 100%", chỉ giữ TÊN.
  function cleanBcctName(raw, code) {
    const parts = String(raw == null ? '' : raw).split('#&').map(s => s.trim());
    if (parts.length > 1 && /^[A-Za-z]{2,3}$/.test(parts[parts.length - 1])) parts.pop(); // xuất xứ (VN…)
    if (parts.length > 1) parts.shift();                                                    // mã ở đầu
    let name = parts.join(' ').trim();
    const c = String(code || '').trim();
    if (c && name.toUpperCase().startsWith(c.toUpperCase() + ' - ')) name = name.slice(c.length + 3);
    else name = name.replace(/^(?=[A-Z0-9\-\/.]*\d)[A-Z0-9][A-Z0-9\-\/.]{3,}\s+-\s+/i, '');   // "DA63-09551B - " (mã khác mã cột)
    return name.replace(/[.\s]*(HÀNG\s*)?MỚI\s*100\s*%\s*$/iu, '').trim();
  }
  const upNoSpace = s => String(s == null ? '' : s).toUpperCase().replace(/\s+/g, '');
  function parseBcct(arr) {
    let h = -1, cc = -1, hc = -1, nc = -1;
    for (let i = 0; i < Math.min(15, arr.length); i++) {
      const row = arr[i] || [];
      const c = row.findIndex(x => upNoSpace(x) === 'MÃNPL/SP');
      if (c >= 0) { h = i; cc = c; hc = row.findIndex(x => upNoSpace(x) === 'MÃHS'); nc = row.findIndex(x => upNoSpace(x) === 'TÊNHÀNG'); break; }
    }
    if (h < 0 || hc < 0 || nc < 0) throw new Error('BCCT thiếu cột "Mã NPL/SP", "Mã HS" hoặc "Tên hàng"');
    const map = {};
    let count = 0;
    for (let i = h + 1; i < arr.length; i++) {
      const row = arr[i];
      if (!row) continue;
      const code = String(row[cc] == null ? '' : row[cc]).trim();
      const key = keyCode(code);
      if (!key) continue;
      const raw = String(row[nc] == null ? '' : row[nc]).trim();
      const name = cleanBcctName(raw, code);
      const hs = normHS(row[hc]);
      if (map[key]) { if (!map[key].all.some(x => x.hs === hs && keyName(x.name) === keyName(name))) map[key].all.push({ hs, name }); continue; }
      map[key] = { code, hs, name, all: [{ hs, name }] };
      count++;
    }
    return { map, count, ver: BCCT_VER };
  }

  // ───────────── đối chiếu BCCT (HS code + tên hàng) ─────────────
  // Tên "trùng khớp" = tên này nằm trong tên kia (sau khi bỏ chữ "phí gia công", dấu cách, dấu câu).
  const nameMatch = (a, b) => {
    const ka = nameKeyN(a), kb = nameKeyN(b);
    return !!ka && !!kb && (ka.includes(kb) || kb.includes(ka));
  };
  // Tên chung chung kiểu "BỘ PHẬN CỦA ..." thì bỏ qua, không cảnh báo
  const isGenericName = s => /^BỘPHẬNCỦA/.test(keyName(s));

  /*
   * bcct: { map: { <mã>: { code, hs, name, all?: [{hs,name}] } }, fileName, savedAt }  (bản BCCT NPL đã lưu)
   * Với mỗi mã hàng: MỌI dòng BCCT của mã đó phải có HS code = HS trên INVOICE, và tên hàng khớp tên trên INVOICE + trên phiếu PDF.
   */
  function compareBcct(inv, pxks, bcct) {
    const res = { rows: [], errors: 0, noAll: false };
    const byCode = new Map();
    const add = (code, f) => { const k = keyCode(code); if (!byCode.has(k)) byCode.set(k, { code, hs: new Set(), invNames: [], pdfNames: [] }); f(byCode.get(k)); };
    inv.lines.forEach(l => add(l.ma, e => { if (l.hs) e.hs.add(l.hs); if (l.ten && !e.invNames.includes(l.ten)) e.invNames.push(l.ten); }));
    pxks.forEach(p => p.rows.forEach(r => add(r.code, e => { if (r.name && !e.pdfNames.includes(r.name)) e.pdfNames.push(r.name); })));
    byCode.forEach(e => {
      const b = bcct.map[keyCode(e.code)];
      const row = { code: e.code, invHs: Array.from(e.hs), invNames: e.invNames, pdfNames: e.pdfNames, bcctRows: [], issues: [], notes: [] };
      if (!b) { row.issues.push('Mã hàng không có trong BCCT'); res.errors++; res.rows.push(row); return; }
      if (!b.all) res.noAll = true;
      row.bcctRows = (b.all || [{ hs: b.hs, name: b.name }]).map(x => ({ hs: x.hs, name: x.name, hsOk: true, nameOk: true, generic: false }));
      row.bcctRows.forEach((x, i) => {
        const tag = row.bcctRows.length > 1 ? `Dòng BCCT ${i + 1}: ` : '';
        if (row.invHs.length && !row.invHs.includes(String(x.hs).replace(/\D/g, ''))) { x.hsOk = false; row.issues.push(`${tag}HS code BCCT ${x.hs || '(trống)'} ≠ INVOICE ${row.invHs.join(' / ')}`); }
        if (!row.invHs.length) row.notes.push('INVOICE không có HS code của mã này');
        const others = e.invNames.map(n => ({ n, src: 'INVOICE' })).concat(e.pdfNames.map(n => ({ n, src: 'phiếu PDF' })));
        const miss = others.filter(o => !nameMatch(x.name, o.n));
        if (miss.length) {
          if (isGenericName(x.name)) { x.generic = true; row.notes.push(`${tag}tên BCCT dạng "BỘ PHẬN CỦA…" nên bỏ qua`); }
          else { x.nameOk = false; row.issues.push(`${tag}tên BCCT không khớp tên ${Array.from(new Set(miss.map(o => o.src))).join(' & ')}`); }
        }
      });
      row.issues = Array.from(new Set(row.issues));
      res.errors += row.issues.length ? 1 : 0;
      res.rows.push(row);
    });
    return res;
  }

  /*
   * inv  : parseInvoice(...)   (sheet INVOICE)
   * pxks : [parsePxk(...)]     (các phiếu PDF)
   * mau27: parseMau27(...) | null (null nếu không có sheet)
   * bcct : bản BCCT NPL đã lưu | null (null = chưa có, bỏ qua phần BCCT)
   */
  function compare(inv, pxks, mau27, bcct) {
    const book = inv.lines.map(l => ({ ma: l.ma, ten: l.ten, luong: l.qty, donGia: l.price, triGia: l.amount, src: l.src, row: l.row }));
    const out = { phieu: [], lines: [], totals: [], errors: 0, warnings: [] };
    const bad = () => { out.errors++; };

    // 1) từng phiếu: số phiếu / ngày / tổng SL / tổng trị giá so với Mẫu 27 + quy tắc "phí gia công"
    const usedMau = new Set();
    pxks.forEach(p => {
      const ph = { file: p.file, soPhieu: p.soPhieu, ngay: p.ngay, tongSL: p.tongSL, tongGia: p.tongGia, checks: [], warnings: p.warnings.slice() };
      // tổng các dòng phải bằng "Tổng cộng" in trên phiếu
      if (p.tongGia != null) {
        const ok = near(p.sumAmount, p.tongGia, 1);
        ph.checks.push({ label: 'Cộng các dòng = Tổng cộng trên phiếu', a: p.sumAmount, b: p.tongGia, ok });
        if (!ok) bad();
      }
      if (mau27 && !mau27.error) {
        const mr = mau27.phieu.find(x => stripLeadZeros(x.so) === stripLeadZeros(p.soPhieu));
        if (!mr) { ph.checks.push({ label: 'Số phiếu có trong Mẫu 27', a: p.soPhieu, b: '(không có)', ok: false }); bad(); }
        else {
          usedMau.add(mr);
          const c = [
            { label: 'Số phiếu', a: p.soPhieu, b: mr.so, ok: stripLeadZeros(p.soPhieu) === stripLeadZeros(mr.so) },
            { label: 'Ngày phiếu', a: p.ngay, b: mr.ngay, ok: p.ngay === mr.ngay },
            { label: 'Tổng số lượng', a: p.tongSL, b: mr.sl, ok: near(p.tongSL, mr.sl, 0.0001) },
            { label: 'Tổng trị giá', a: p.tongGia != null ? p.tongGia : p.sumAmount, b: mr.gia, ok: near(p.tongGia != null ? p.tongGia : p.sumAmount, mr.gia, 1) }
          ];
          c.forEach(x => { ph.checks.push(x); if (!x.ok) bad(); });
        }
      }
      // quy tắc "phí gia công": đã có dòng nào ghi thì mọi dòng của phiếu đều phải ghi
      const withPhi = p.rows.filter(r => hasPhiGC(r.name));
      if (withPhi.length) {
        const missing = p.rows.filter(r => !hasPhiGC(r.name));
        ph.phiGC = { total: p.rows.length, withCount: withPhi.length, missing: missing.map(r => `${r.stt}. ${r.code}`), ok: missing.length === 0 };
        if (missing.length) bad();
      } else ph.phiGC = { total: p.rows.length, withCount: 0, missing: [], ok: true, none: true };
      out.phieu.push(ph);
    });
    if (mau27 && !mau27.error) {
      mau27.phieu.filter(x => !usedMau.has(x)).forEach(x => {
        out.phieu.push({ file: '(không có file PDF)', soPhieu: x.so, ngay: x.ngay, tongSL: x.sl, tongGia: x.gia, checks: [{ label: 'Phiếu có trong Mẫu 27 nhưng chưa có file PDF', a: '(không có PDF)', b: x.so, ok: false }], phiGC: { none: true, ok: true, total: 0, withCount: 0, missing: [] }, warnings: [] });
        bad();
      });
    }

    // 2) từng dòng hàng: sheet INVOICE <-> PDF  (ghép theo mã hàng + số lượng, mỗi dòng chỉ ghép 1 lần)
    // Một mã hàng trên INVOICE thường là 1 dòng tổng, còn trên phiếu xuất kho bị chia ra nhiều dòng / nhiều phiếu.
    // Nên gom theo (mã hàng + đơn giá) ở cả hai bên rồi so TỔNG số lượng, TỔNG trị giá.
    const cluster = (list, priceOf) => {
      const cs = [];
      list.forEach(x => { const c = cs.find(k => near(k.price, priceOf(x), 0.5)); if (c) c.items.push(x); else cs.push({ price: priceOf(x), items: [x] }); });
      return cs;
    };
    const sum = (arr, f) => arr.reduce((s, x) => s + (Number(f(x)) || 0), 0);
    const codes = [];
    book.forEach(d => { const k = keyCode(d.ma); if (!codes.includes(k)) codes.push(k); });
    pxks.forEach(p => p.rows.forEach(r => { const k = keyCode(r.code); if (!codes.includes(k)) codes.push(k); }));
    const pdfAll = [];
    pxks.forEach(p => p.rows.forEach(r => pdfAll.push(Object.assign({ soPhieu: p.soPhieu }, r))));

    codes.forEach(k => {
      const invL = book.filter(d => keyCode(d.ma) === k);
      const pdfL = pdfAll.filter(r => keyCode(r.code) === k);
      const pdfCl = cluster(pdfL, r => r.price).map(c => Object.assign(c, { used: false }));
      const rowsTxt = arr => arr.map(d => d.row).join(', ');
      invL.length && cluster(invL, d => d.donGia).forEach(ic => {
        const items = ic.items;
        const line = { idx: rowsTxt(items), ma: items[0].ma, issues: [], invCount: items.length };
        const luong = sum(items, d => d.luong), triGia = sum(items, d => d.triGia);
        line.book = { luong, donGia: ic.price, triGia, src: items[0].src };
        // cụm PDF cùng đơn giá; không có thì so với toàn bộ các dòng PDF còn lại của mã này (để báo lệch giá)
        let pc = pdfCl.find(c => !c.used && near(c.price, ic.price, 0.5));
        let prs;
        if (pc) { pc.used = true; prs = pc.items; }
        else { const rest = pdfCl.filter(c => !c.used); rest.forEach(c => { c.used = true; }); prs = rest.reduce((a, c) => a.concat(c.items), []); }
        if (!prs.length) { line.pdf = null; line.issues.push('Không có trên phiếu xuất kho nào'); bad(); out.lines.push(line); return; }
        const pQty = sum(prs, r => r.qty), pAmt = sum(prs, r => r.amount);
        const pPrice = prs.every(r => near(r.price, prs[0].price, 0.5)) ? prs[0].price : null; // null = các dòng PDF có đơn giá khác nhau
        line.pdf = { qty: pQty, price: pPrice, amount: pAmt };
        line.pdfCount = prs.length;
        line.pdfPhieu = Array.from(new Set(prs.map(r => r.soPhieu)));
        const invKeys = items.map(d => nameKeyN(d.ten));
        if (!prs.every(r => invKeys.includes(nameKeyN(r.name)))) line.issues.push('Tên hàng không khớp');
        if (!near(luong, pQty, 0.0001)) line.issues.push('Số lượng không khớp');
        if (pPrice == null || !near(ic.price, pPrice, 0.5)) line.issues.push('Đơn giá không khớp');
        if (!near(triGia, pAmt, 1)) line.issues.push('Trị giá không khớp');
        if (line.issues.length) bad();
        out.lines.push(line);
      });
      pdfCl.filter(c => !c.used).forEach(c => {
        c.items.forEach(r => out.lines.push({ idx: null, ma: r.code, issues: ['Có trên phiếu ' + r.soPhieu + ' nhưng không có trong sheet INVOICE'], book: null, pdf: r }));
        bad();
      });
    });

    // 3) dòng "Tổng" của Mẫu 27 <-> cộng tất cả các phiếu PDF
    const sumPdfSL = pxks.reduce((s, p) => s + p.tongSL, 0);
    const sumPdfGia = pxks.reduce((s, p) => s + (p.tongGia != null ? p.tongGia : p.sumAmount), 0);
    const T = (label, a, b, tol) => { const ok = near(a, b, tol); out.totals.push({ label, a, b, ok }); if (!ok) bad(); };
    if (mau27 && !mau27.error && mau27.total) {
      if (mau27.total.sl != null) T('Tổng số lượng: dòng Tổng Mẫu 27 ↔ cộng các phiếu PDF', mau27.total.sl, sumPdfSL, 0.0001);
      if (mau27.total.gia != null) T('Tổng trị giá: dòng Tổng Mẫu 27 ↔ cộng các phiếu PDF', mau27.total.gia, sumPdfGia, 1);
    }
    if (bcct && bcct.map) {
      out.bcct = compareBcct(inv, pxks, bcct);
      out.errors += out.bcct.errors;
      if (out.bcct.noAll) out.warnings.push('BCCT đã lưu ở trang NPL là bản cũ, chỉ có dòng đầu tiên của mỗi mã. Vào trang NPL bấm "Cập nhật BCCT" chọn lại file để kiểm tra được mọi dòng BCCT.');
    }
    if (inv.usedDap && inv.hasProcessing) out.warnings.push(`${inv.usedDap} dòng trong sheet INVOICE không có giá Processing fee nên đã lấy giá DAP VALUE để so.`);
    if (!inv.hasProcessing) out.warnings.push('Sheet INVOICE không có cột giá Processing fee — dùng giá DAP VALUE để so với phiếu.');
    if (!mau27) out.warnings.push('Không tìm thấy sheet MẪU 27 trong file INVOICE — bỏ qua phần đối chiếu Mẫu 27');
    else if (mau27.error) out.warnings.push(mau27.error);
    return out;
  }

  return { parsePxk, parsePxkFile, parseMau27, parseInvoice, parseBcct, cleanBcctName, BCCT_VER, compare, compareBcct, nameMatch, groupLines, normDate, keyName };
});
