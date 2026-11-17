/* Logic khung bên phải của add-in Excel. Dùng chung lõi với tool Gop_file_Excel.html:
   worker_gop.js là bản ghép SheetJS + tho_doc.js + loi_xu_ly.js + worker_chinh.js do dong_goi_addin.js tạo,
   nói chuyện theo giao thức ở mục 4.4 YEU_CAU_KY_THUAT.md. Phần riêng của add-in là ghi kết quả vào workbook:
   - Microsoft 365 (ExcelApi 1.13): chèn nguyên tệp .xlsx do worker tạo thành các sheet mới.
   - Excel 2016, 2019, 2021 (ExcelApi 1.2 đến 1.12): worker gửi kèm các bảng, add-in ghi từng khối ô vào sheet mới.
   - Mở kết quả thành workbook mới khi có ExcelApi 1.8. */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var GF = window.GopFile, coLoi = !!(GF && GF.CAU_HINH_MAU), tep = [], ketQua = null;
  var trongExcel = false, coChen = false, coGhiO = false, coTaoMoi = false, daBietCheDo = false;
  var DUOI = /\.(xlsx|xlsm|xls|csv)$/i, KHOA_LUU = 'gop-file-cau-hinh';
  var TEN_SHEET = ['Du_lieu_gop', 'Doi_soat', 'Nhat_ky_loi', 'Tong_hop'];
  var GIOI_HAN_GHI_O = 200000;   // ghi từng ô thì trên số dòng này dữ liệu gộp chỉ tải về, chỉ ghi các sheet báo cáo
  var KHOI_GHI = 5000;           // mỗi lần đồng bộ với Excel ghi 5.000 dòng

  function ghiDuoc() { return trongExcel && (coChen || coGhiO); }

  // ---------- Kết nối Excel ----------

  function capNhatCheDo() {
    daBietCheDo = true;
    var cd = $('che-do');
    if (trongExcel) {
      cd.className = 'che-do';
      cd.textContent = coChen || coGhiO ? 'Đã kết nối Excel. Kết quả sẽ thêm thành các sheet mới ở cuối workbook.'
        : coTaoMoi ? 'Bản Excel này chưa ghi sheet được, kết quả sẽ mở thành workbook mới.'
        : 'Bản Excel này chưa ghi kết quả được, hãy tải tệp kết quả về.';
    } else {
      cd.className = 'che-do canh';
      cd.textContent = 'Đang mở ngoài Excel. Vẫn gộp và đối soát được, kết quả chỉ tải về.';
    }
    $('dong-ghi').classList.toggle('an', !ghiDuoc());
    $('ghi-lai').classList.toggle('an', !ghiDuoc());
    $('mo-moi').classList.toggle('an', !(trongExcel && coTaoMoi));
  }
  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(function (info) {
      trongExcel = !!info && info.host === Office.HostType.Excel;
      if (trongExcel) {
        var rq = Office.context.requirements;
        coChen = rq.isSetSupported('ExcelApi', '1.13');
        coGhiO = rq.isSetSupported('ExcelApi', '1.2');
        coTaoMoi = rq.isSetSupported('ExcelApi', '1.8');
      }
      capNhatCheDo();
    });
    setTimeout(function () { if (!daBietCheDo) capNhatCheDo(); }, 4000);
  } else {
    capNhatCheDo();
  }

  // ---------- Cấu hình, giống tool HTML ----------

  function chMacDinh() { return coLoi ? JSON.stringify(GF.CAU_HINH_MAU, null, 1) : ''; }
  var luu = null; try { luu = localStorage.getItem(KHOA_LUU); } catch (e) {}
  $('cau-hinh').value = (coLoi && luu) || chMacDinh();
  function docCH() {
    if (!coLoi) { $('ten-mau').innerHTML = '<span class="canh">Thiếu lõi xử lý loi_xu_ly.js.</span> Đóng gói lại bằng lệnh node 04_Addin_Excel/dong_goi_addin.js.'; return null; }
    try {
      var c = JSON.parse($('cau-hinh').value);
      var cotTien = (c.cot || []).filter(function (x) { return x.cotTien; })[0];
      $('ten-mau').textContent = 'Mẫu đang dùng là ' + c.ten + ', ' + c.cot.length + ' cột chuẩn, cột tiền để đối soát là ' + (cotTien ? cotTien.ten : 'chưa chọn') + '.';
      return c;
    } catch (e) { $('ten-mau').textContent = 'Cấu hình chưa đúng định dạng JSON.'; return null; }
  }
  docCH();
  $('cau-hinh').addEventListener('input', function () { if (docCH()) { try { localStorage.setItem(KHOA_LUU, $('cau-hinh').value); } catch (e) {} } });
  $('mac-dinh').onclick = function () { $('cau-hinh').value = chMacDinh(); docCH(); try { localStorage.removeItem(KHOA_LUU); } catch (e) {} };
  $('luu-ch').onclick = function () { taiVe(new Blob([$('cau-hinh').value], { type: 'application/json' }), 'cau_hinh_gop_file.json'); };
  $('nap-ch').onchange = function (e) {
    var f = e.target.files[0]; if (!f) return;
    f.text().then(function (t) { $('cau-hinh').value = t; if (docCH()) { try { localStorage.setItem(KHOA_LUU, t); } catch (er) {} } });
    e.target.value = '';
  };

  // ---------- Tiện ích ----------

  function mb(b) { return (b / 1048576).toFixed(1).replace('.', ',') + ' MB'; }
  function soVN(n) { return n == null || n === '' ? '' : (typeof n === 'number' ? n.toLocaleString('vi-VN') : String(n)); }
  function giay(ms) { return (ms / 1000).toFixed(1).replace('.', ',') + ' giây'; }
  function an(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
  function thongDiep(e) { return String(e && (e.message || e.code) || e || 'không rõ nguyên nhân'); }
  function taiVe(blob, ten) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = ten;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function blobBase64(blob) {
    return new Promise(function (ok, hong) {
      var r = new FileReader();
      r.onload = function () { var s = String(r.result), i = s.indexOf(','); ok(i >= 0 ? s.slice(i + 1) : s); };
      r.onerror = function () { hong(r.error); };
      r.readAsDataURL(blob);
    });
  }

  // ---------- Chọn tệp ----------

  function themTep(list) {
    for (var i = 0; i < list.length; i++) {
      var f = list[i];
      if (!DUOI.test(f.name) || f.name.indexOf('~$') === 0) continue;
      if (!tep.some(function (x) { return x.name === f.name && x.size === f.size; })) tep.push(f);
    }
    var tong = tep.reduce(function (a, f) { return a + f.size; }, 0);
    $('ds-tep').textContent = tep.length ? tep.length + ' tệp, ' + mb(tong) + '. ' + tep.slice(0, 4).map(function (f) { return f.name; }).join(', ') + (tep.length > 4 ? ' và ' + (tep.length - 4) + ' tệp khác.' : '.') : 'Chưa có tệp nào.';
    $('chay').disabled = !tep.length;
    $('trang-thai').textContent = tep.length ? 'Sẵn sàng gộp ' + tep.length + ' tệp.' : 'Chọn tệp ở bước 1 để bắt đầu.';
  }
  $('chon-tep').onchange = function (e) { themTep(e.target.files); e.target.value = ''; };
  $('chon-thu-muc').onchange = function (e) { themTep(e.target.files); e.target.value = ''; };
  $('xoa-tep').onclick = function () { tep = []; themTep([]); };
  var tha = $('tha');
  tha.addEventListener('dragover', function (e) { e.preventDefault(); tha.classList.add('vao'); });
  tha.addEventListener('dragleave', function () { tha.classList.remove('vao'); });
  tha.addEventListener('drop', function (e) { e.preventDefault(); tha.classList.remove('vao'); themTep(e.dataTransfer.files); });

  // ---------- Ghi vào Excel ----------

  // Hậu tố để sheet mới không trùng tên sheet có sẵn: lần đầu để trống, sau đó _2, _3...
  function timHauTo() {
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets;
      ws.load('items/name');
      return ctx.sync().then(function () {
        var co = ws.items.map(function (w) { return w.name.toLowerCase(); });
        for (var i = 1; i < 100; i++) {
          var h = i === 1 ? '' : '_' + i;
          if (!TEN_SHEET.some(function (t) { return co.indexOf((t + h).toLowerCase()) >= 0; })) return h;
        }
        return '_' + (Date.now() % 100000);
      });
    });
  }

  // Chèn mọi sheet của một tệp .xlsx vào cuối workbook rồi mở sheet Doi_soat
  function chenSheet(blob, hauTo) {
    return blobBase64(blob).then(function (b64) {
      return Excel.run(function (ctx) {
        ctx.workbook.insertWorksheetsFromBase64(b64, { positionType: Excel.WorksheetPositionType.end });
        return ctx.sync().then(function () {
          var ds = ctx.workbook.worksheets.getItemOrNullObject('Doi_soat' + hauTo);
          return ctx.sync().then(function () { if (!ds.isNullObject) { ds.activate(); return ctx.sync(); } });
        });
      });
    });
  }

  function tenCacSheet(m, coDuLieu) {
    var h = m.hauTo || '', ds = [];
    if (coDuLieu) ds.push('Du_lieu_gop' + h);
    ds.push('Doi_soat' + h);
    if (m.loiMau) ds.push('Nhat_ky_loi' + h);
    if (m.tongHop) ds.push('Tong_hop' + h);
    return ds.join(', ');
  }

  // Cách 1, Microsoft 365. Tệp dữ liệu .xlsx đã gồm cả các sheet báo cáo.
  // Dữ liệu là CSV, hoặc quá lớn để Excel nhận một lần, thì chỉ chèn báo cáo
  function ghiBangChen(m) {
    var coXlsx = m.dinhDang === 'xlsx' && m.duLieu && m.duLieu.blob, loiDuLieu = null;
    var buoc = coXlsx ? chenSheet(m.duLieu.blob, m.hauTo).then(function () { return true; }, function (e) { loiDuLieu = e; return false; }) : Promise.resolve(false);
    return buoc.then(function (xong) {
      if (xong) return 'Đã thêm các sheet ' + tenCacSheet(m, true) + ' vào cuối workbook.';
      return chenSheet(m.baoCao.blob, m.hauTo).then(function () {
        return 'Đã thêm các sheet ' + tenCacSheet(m, false) + ' vào cuối workbook. ' +
          (loiDuLieu ? 'Dữ liệu gộp quá lớn để chèn thẳng (' + thongDiep(loiDuLieu) + '), bấm Tải dữ liệu gộp để lấy tệp.' : 'Dữ liệu gộp ở dạng CSV, bấm Tải dữ liệu gộp để lấy tệp.');
      });
    });
  }

  // Cách 2, Excel 2016, 2019, 2021: ghi từng khối ô vào sheet mới, chỉ dùng hàm của ExcelApi 1.1 và 1.2
  function chuCot(n) { var t = ''; n++; while (n > 0) { var m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); } return t; }

  // Chữ giống số hoặc ngày, hay bắt đầu bằng = + - @, thì thêm dấu nháy để Excel giữ nguyên là chữ
  var GIONG_SO = /^[\s\d.,:\/\-+()%]*\d[\s\d.,:\/\-+()%]*$/;
  function oGhi(v) {
    if (v == null) return '';
    if (typeof v === 'number') return isFinite(v) ? v : '';
    if (typeof v === 'boolean') return v;
    var s = String(v);
    return /^[=+\-@]/.test(s) || GIONG_SO.test(s) ? "'" + s : s;
  }

  // Ghi một bảng thành sheet mới, mỗi lần 5.000 dòng. dinhDang[k] là định dạng số của cột k, áp cho các dòng dưới tiêu đề
  function ghiMotBang(ten, aoa, dinhDang, baoTienDo) {
    var soCot = aoa.reduce(function (mx, r) { return Math.max(mx, r.length); }, 1), soDong = aoa.length, cuoi = chuCot(soCot - 1);
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.add(ten), batDau = 0;
      function mot() {
        var khoi = aoa.slice(batDau, batDau + KHOI_GHI).map(function (r) { var o = new Array(soCot); for (var k = 0; k < soCot; k++) o[k] = oGhi(r[k]); return o; });
        var d1 = batDau + 1, d2 = batDau + khoi.length, f1 = Math.max(d1, 2);
        ws.getRange('A' + d1 + ':' + cuoi + d2).values = khoi;
        if (d2 >= f1) (dinhDang || []).forEach(function (f, k) {
          if (!f) return;
          var mang = []; for (var i = f1; i <= d2; i++) mang.push([f]);
          ws.getRange(chuCot(k) + f1 + ':' + chuCot(k) + d2).numberFormat = mang;
        });
        batDau = d2;
        return ctx.sync().then(function () {
          if (baoTienDo) baoTienDo(batDau, soDong);
          if (batDau < soDong) return mot();
          ws.getRange('A1:' + cuoi + '1').format.font.bold = true;
          ws.getRange('A1:' + cuoi + Math.min(soDong, 200)).format.autofitColumns();
          return ctx.sync();
        });
      }
      return mot();
    });
  }

  function dinhDangTheoTen(hdr, tenCot) { return hdr.map(function (h) { return tenCot.indexOf(h) >= 0 ? '#,##0' : null; }); }

  function ghiBangO(m, tb) {
    var b = m.bang, viec = [];
    if (!b) return Promise.reject(new Error('thiếu bảng kết quả, hãy bấm Gộp và đối soát lại'));
    // Tính hậu tố lúc ghi vì người dùng có thể đã thêm sheet sau khi gộp
    return timHauTo().then(function (h) {
      m.hauTo = h;
      if (b.duLieu) viec.push(['Du_lieu_gop' + h, b.duLieu, b.dinhDang.concat([null, null, null])]);
      viec.push(['Doi_soat' + h, b.doiSoat, dinhDangTheoTen(b.doiSoat[0], ['Tổng các dòng', 'Tổng ghi trong tệp', 'Chênh lệch'])]);
      if (b.nhatKy) viec.push(['Nhat_ky_loi' + h, b.nhatKy, []]);
      if (b.tongHop) viec.push(['Tong_hop' + h, b.tongHop, b.tongHop[0].map(function (x, k) { return k ? '#,##0' : null; })]);
      return viec.reduce(function (p, v) {
        return p.then(function () {
          return ghiMotBang(v[0], v[1], v[2], function (d, n) { tb.textContent = 'Đang ghi sheet ' + v[0] + ', ' + soVN(Math.max(0, d - 1)) + '/' + soVN(n - 1) + ' dòng…'; });
        });
      }, Promise.resolve()).then(function () {
        return Excel.run(function (ctx) { ctx.workbook.worksheets.getItem('Doi_soat' + h).activate(); return ctx.sync(); });
      });
    }).then(function () {
      return 'Đã thêm các sheet ' + viec.map(function (v) { return v[0]; }).join(', ') + ' vào cuối workbook.' +
        (b.duLieu ? '' : ' Dữ liệu gộp có ' + soVN(m.thongKe.dongGop) + ' dòng, quá lớn để ghi vào sheet, bấm Tải dữ liệu gộp để lấy tệp.');
    });
  }

  function ghiVaoExcel() {
    var m = ketQua; if (!m || !ghiDuoc()) return Promise.resolve();
    var tb = $('ghi-trang-thai'), nut = $('ghi-lai');
    tb.className = 'ghi'; tb.textContent = 'Đang ghi kết quả vào workbook…'; tb.classList.remove('an');
    nut.disabled = true;
    return (coChen ? ghiBangChen(m) : ghiBangO(m, tb)).then(function (chu) {
      tb.textContent = chu + ' Nhấn Ctrl+S để lưu workbook.';
      nut.textContent = 'Đã ghi vào workbook';
    }, function (e) {
      tb.className = 'ghi loi';
      tb.textContent = 'Chưa ghi được vào workbook, ' + thongDiep(e) + '.' + (coTaoMoi ? ' Có thể bấm Mở thành workbook mới hoặc tải tệp về.' : ' Có thể tải tệp về.');
      nut.disabled = false;
    });
  }
  $('ghi-lai').onclick = function () { ghiVaoExcel(); };

  $('mo-moi').onclick = function () {
    if (!ketQua) return;
    var blob = ketQua.dinhDang === 'xlsx' ? ketQua.duLieu.blob : ketQua.baoCao.blob, tb = $('ghi-trang-thai');
    blobBase64(blob).then(function (b64) { return Excel.createWorkbook(b64); }).then(function () {
      tb.className = 'ghi'; tb.textContent = 'Đã mở ' + (ketQua.dinhDang === 'xlsx' ? 'dữ liệu gộp và báo cáo' : 'báo cáo đối soát') + ' thành workbook mới.';
    }, function (e) {
      tb.className = 'ghi loi'; tb.textContent = 'Chưa mở được workbook mới, ' + thongDiep(e) + '. Hãy tải tệp về.';
    }).then(function () { tb.classList.remove('an'); });
  };

  // ---------- Hiện kết quả ----------

  function veBang(id, hdr, rows, khiTrong) {
    if (!rows || !rows.length) { $(id).innerHTML = '<p>' + an(khiTrong) + '</p>'; return; }
    var h = '<table><thead><tr>' + hdr.map(function (x) { return '<th>' + an(x) + '</th>'; }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      h += '<tr>' + r.map(function (v) {
        var t = an(soVN(v));
        if (t === 'ĐẠT') t = '<span class="dat">ĐẠT</span>';
        else if (/^LỆCH/.test(t)) t = '<span class="lech-chu">LỆCH</span>';
        return '<td' + (typeof v === 'number' ? ' class="num"' : '') + '>' + t + '</td>';
      }).join('') + '</tr>';
    });
    $(id).innerHTML = h + '</tbody></table>';
  }

  // Lấy một số cột của bảng hai chiều theo tên cột ở dòng đầu
  function layCot(aoa, ten) {
    if (!aoa || !aoa.length) return [];
    var vt = ten.map(function (t) { return aoa[0].indexOf(t); });
    return aoa.slice(1).map(function (r) { return vt.map(function (i) { return i >= 0 ? r[i] : null; }); });
  }

  function hienKetQua(m) {
    var tk = m.thongKe || {}, o = function (v) { return v == null ? 'chưa có' : soVN(v); };
    $('kpi').innerHTML = [[o(tk.soSheet), 'sheet đã gộp'], [o(tk.dongGop), 'dòng gộp'], [o(tk.tongTienGop), 'tổng thành tiền'],
      [o(tk.dongLoi), 'dòng lỗi'], [o(tk.dongTrung), 'dòng trùng'], [o(tk.soLech), 'sheet lệch dòng tổng', tk.soLech > 0]]
      .map(function (x) { return '<div' + (x[2] ? ' class="lech"' : '') + '><b>' + an(x[0]) + '</b><span>' + x[1] + '</span></div>'; }).join('');
    veBang('bang-ds', ['Tệp', 'Sheet', 'Dòng gộp', 'Chênh lệch', 'Kết luận'],
      layCot(m.doiSoat, ['Tệp', 'Sheet', 'Dòng gộp', 'Chênh lệch', 'Kết luận']).slice(0, 300), 'Chưa có bảng đối soát.');
    veBang('bang-loi', ['Tệp', 'Dòng gốc', 'Cột', 'Lỗi'],
      layCot(m.loiMau ? [['Tệp', 'Sheet', 'Dòng gốc', 'Cột', 'Giá trị', 'Lỗi']].concat(m.loiMau) : null, ['Tệp', 'Dòng gốc', 'Cột', 'Lỗi']).slice(0, 20),
      'Không có dòng lỗi hay dòng trùng.');
    var coDL = m.duLieu && m.duLieu.blob, coBC = m.baoCao && m.baoCao.blob;
    $('tai-du-lieu').disabled = !coDL; $('tai-bao-cao').disabled = !coBC;
    $('tai-du-lieu').textContent = coDL ? 'Tải dữ liệu gộp (.' + (m.dinhDang || 'xlsx') + ', ' + mb(m.duLieu.blob.size) + ')' : 'Chưa có tệp dữ liệu gộp';
    $('ghi-lai').disabled = false; $('ghi-lai').textContent = 'Ghi vào workbook này';
    $('ghi-trang-thai').classList.add('an');
    $('ket-qua').classList.remove('an');
  }
  $('tai-du-lieu').onclick = function () { if (ketQua && ketQua.duLieu) taiVe(ketQua.duLieu.blob, ketQua.duLieu.ten); };
  $('tai-bao-cao').onclick = function () { if (ketQua && ketQua.baoCao) taiVe(ketQua.baoCao.blob, ketQua.baoCao.ten); };

  // ---------- Chạy ----------

  function chay(ch, hauTo) {
    $('trang-thai').textContent = 'Đang khởi động luồng xử lý.';
    var w;
    try { w = new Worker('worker_gop.js'); }
    catch (e) { $('chay').disabled = false; $('trang-thai').textContent = 'Không mở được luồng xử lý, ' + thongDiep(e); return; }
    var batDau = performance.now();
    function ketThuc() { w.terminate(); $('chay').disabled = false; }
    w.onmessage = function (e) {
      var m = e.data || {};
      if (m.loai === 'tien-do') {
        $('thanh').style.width = Math.round(Math.min(1, m.i / Math.max(1, m.n)) * 100) + '%';
        $('trang-thai').textContent = (m.i < m.n ? 'Đang đọc tệp ' + (m.i + 1) + '/' + m.n + ', ' + m.ten : m.ten) + '. Đã gộp ' + soVN(m.dong || 0) + ' dòng, ' + giay(performance.now() - batDau) + '.';
      } else if (m.loai === 'xong') {
        ketThuc();
        m.hauTo = hauTo; ketQua = m;
        var tg = m.thoiGian;
        $('thanh').style.width = '100%';
        $('trang-thai').textContent = tg ? 'Xong sau ' + giay(tg.tong) + '. Đọc ' + giay(tg.doc) + ', xử lý ' + giay(tg.xuLy) + ', ghi tệp ' + giay(tg.ghi) + '.' : 'Xong sau ' + giay(performance.now() - batDau) + '.';
        hienKetQua(m);
        if (ghiDuoc() && $('ghi-excel').checked) ghiVaoExcel();
      } else if (m.loai === 'loi') {
        ketThuc(); $('trang-thai').textContent = 'Có lỗi, ' + m.thongDiep;
      }
    };
    w.onerror = function (e) { ketThuc(); $('trang-thai').textContent = 'Luồng xử lý lỗi, ' + (e.message || 'không rõ nguyên nhân'); };
    // Ghi từng ô (Excel 2016 đến 2021) cần worker gửi kèm các bảng, chèn tệp (Microsoft 365) thì không
    var guiBang = trongExcel && !coChen && coGhiO;
    w.postMessage({ lenh: 'chay', tep: tep, cauHinh: ch, tuyChon: { xuat: $('xuat').value, nguongCSV: 300000, hauTo: hauTo, guiBang: guiBang, gioiHanBang: GIOI_HAN_GHI_O } });
  }

  $('chay').onclick = function () {
    var ch = docCH();
    if (!ch) { $('trang-thai').textContent = coLoi ? 'Sửa cấu hình cho đúng JSON rồi chạy lại.' : 'Thiếu lõi xử lý, đóng gói lại add-in.'; return; }
    $('chay').disabled = true; $('ket-qua').classList.add('an'); $('thanh').style.width = '0%';
    // Chèn tệp thì tên sheet nằm sẵn trong tệp worker tạo, nên tìm hậu tố trước khi gộp; ghi từng ô thì tìm lúc ghi
    if (trongExcel && coChen) {
      $('trang-thai').textContent = 'Đang xem tên các sheet có sẵn.';
      timHauTo().then(function (h) { chay(ch, h); }, function () { chay(ch, '_' + (Date.now() % 100000)); });
    } else {
      chay(ch, '');
    }
  };
})();
