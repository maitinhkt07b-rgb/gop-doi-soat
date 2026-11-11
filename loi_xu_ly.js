/* Lõi xử lý tool gộp file Excel (khối loi-xu-ly). Claude viết tệp này theo YEU_CAU_KY_THUAT.md.
   Bộ khởi động chỉ để sẵn cấu hình mẫu và tên các hàm cần có, hàm nào chưa xây thì báo "Chưa xây".
   Không chạm giao diện, không gọi mạng. Trên trang, khung chỉ đọc GopFile.CAU_HINH_MAU,
   mọi việc còn lại chạy trong Web Worker. Hợp đồng các hàm ở mục 4.5. */
(function (root) {
  'use strict';

  // Cấu hình mẫu ở mục 5. Giữ đủ 12 cột theo đúng thứ tự, bộ kiểm thử đọc cấu hình này.
  var CAU_HINH_MAU = {
    ten: 'Bảng kê chi phí vật tư công trường',
    cot: [
      { ten: 'Ngày', kieu: 'ngay', batBuoc: true, dongNghia: ['Ngày CT', 'Ngày chứng từ', 'Ngày phiếu'] },
      { ten: 'Số phiếu', kieu: 'ma', batBuoc: true, dongNghia: ['Số CT', 'Số chứng từ'] },
      { ten: 'Công trường', kieu: 'ma', batBuoc: true, dongNghia: ['Tên công trường', 'CT'], layTuTenTep: '^(CT\\d+)' },
      { ten: 'Hạng mục', kieu: 'ma', dongNghia: ['Mã HM', 'HM'] },
      { ten: 'Mã vật tư', kieu: 'ma', batBuoc: true, dongNghia: ['Mã VT'] },
      { ten: 'Tên vật tư', kieu: 'chu', dongNghia: ['Tên VT', 'Diễn giải'] },
      { ten: 'Đơn vị tính', kieu: 'chu', dongNghia: ['ĐVT', 'Đơn vị'] },
      { ten: 'Số lượng', kieu: 'so', batBuoc: true, dongNghia: ['SL', 'Khối lượng'] },
      { ten: 'Đơn giá', kieu: 'so', dinhDang: '#,##0', dongNghia: ['Giá', 'Đơn giá (đ)'] },
      { ten: 'Thành tiền', kieu: 'so', batBuoc: true, cotTien: true, dongNghia: ['Số tiền', 'Thành tiền (VNĐ)', 'Giá trị'] },
      { ten: 'Nhà cung cấp', kieu: 'chu', dongNghia: ['NCC', 'Đơn vị bán'] },
      { ten: 'Ghi chú', kieu: 'chu', dongNghia: [] }
    ],
    khoaTrung: ['Số phiếu', 'Mã vật tư', 'Ngày', 'Thành tiền'],
    dongTong: ['Tổng cộng', 'Cộng', 'Tổng', 'Total', 'Cộng trang'],
    tongHop: { hang: 'Công trường', cot: 'Hạng mục' }
  };

  var COT_NGUON = ['Tệp nguồn', 'Sheet', 'Dòng gốc'];
  var GIOI_HAN_DONG = 1048575; // Excel cho 1.048.576 dòng một sheet, trừ dòng tiêu đề

  var COT_DOI_SOAT = ['Tệp', 'Sheet', 'Dòng tiêu đề', 'Dòng đọc', 'Dòng gộp', 'Dòng lỗi', 'Dòng trùng', 'Tổng các dòng',
    'Tổng ghi trong tệp', 'Chênh lệch', 'Kết luận', 'Cột lạ', 'Cột thiếu', 'Lấy từ tên tệp'];
  var SO_COT_TOI_THIEU = 3;          // khớp từ 3 cột chuẩn trở lên mới coi là bảng
  var SO_DONG_TIM_TIEU_DE = 25;      // tìm dòng tiêu đề trong 25 dòng đầu
  var SO_O_TIM_DONG_TONG = 8;        // chữ dòng tổng nằm trong 8 ô đầu dòng
  var NGAY_GOC_EXCEL = 25569;        // số seri Excel của ngày 01/01/1970
  var LY_DO_KHONG_TIEU_DE = 'Không tìm thấy dòng tiêu đề khớp bộ cột chuẩn';
  var LOI_THIEU = 'Thiếu giá trị bắt buộc', LOI_SO = 'Không đọc được số', LOI_NGAY = 'Ngày không hợp lệ';
  var COT_CA_DONG = '(cả dòng)';     // cột Cột của nhật ký khi dòng trùng
  var SO_DONG_KHOI_CSV = 20000;      // mỗi khối chuỗi CSV khoảng 20.000 dòng
  var GIOI_HAN_GOM = 50000;          // cột có quá 50.000 giá trị khác nhau (như số phiếu) thì thôi gom chuỗi
  var GIO = '(?:[ T]+\\d{1,2}:\\d{2}(?::\\d{2}(?:\\.\\d+)?)?(?:\\s*[AaPp][Mm])?)?';
  var NGAY_NTN = new RegExp('^(\\d{1,2})[/.-](\\d{1,2})[/.-](\\d{4}|\\d{2})' + GIO + '$');  // dd/mm/yyyy, d/m/yy, dd-mm-yyyy, dd.mm.yyyy
  var NGAY_NTN_NGUOC = new RegExp('^(\\d{4})[/.-](\\d{1,2})[/.-](\\d{1,2})' + GIO + '$');  // yyyy-mm-dd

  // ---------- Chuẩn hóa ô và tên cột ----------

  function laTrong(v) { return v == null || (typeof v === 'string' && !v.trim()); }

  function dongTrong(row) {
    if (!row) return true;
    for (var i = 0; i < row.length; i++) if (!laTrong(row[i])) return false;
    return true;
  }

  // So tên cột sau khi bỏ dấu tiếng Việt (đ thành d), đổi chữ thường, đổi ký tự không phải chữ hoặc số thành khoảng trắng
  function chuanHoaTen(s) {
    if (laTrong(s)) return '';
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'd').toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function lamTron2(n) { var r = Math.round(n * 100) / 100; return r === 0 ? 0 : r; }

  // Số gõ dạng chữ theo cách viết của người Việt: "1.250.000" và "1.250" có dấu chấm tách nghìn, "12,5" có dấu phẩy thập phân,
  // từ hai nhóm phẩy như "1,250,000" là kiểu tiếng Anh. Bỏ khoảng trắng, bỏ đuôi đ hoặc VNĐ. Số trong ngoặc hoặc có dấu trừ là số âm
  function docSo(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (laTrong(v)) return NaN;
    var s = String(v).replace(/\s+/g, '').replace(/(vnđ|vnd|đ)$/i, ''), am = false, n;
    if (/^\(.*\)$/.test(s)) { am = true; s = s.slice(1, -1); }
    if (s.charAt(0) === '-') { am = !am; s = s.slice(1); }
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    else if (/^\d{1,3}(,\d{3}){2,}(\.\d+)?$/.test(s)) n = parseFloat(s.replace(/,/g, ''));
    else if (/^\d+(,\d+)?$/.test(s)) n = parseFloat(s.replace(',', '.'));
    else if (/^\d+(\.\d+)?$/.test(s)) n = parseFloat(s);
    else return NaN;
    return am ? -n : n;
  }

  // Số seri Excel của một ngày, NaN khi ngày không có thật
  function seriNgay(nam, thang, ngay) {
    if (thang < 1 || thang > 12 || ngay < 1 || ngay > 31) return NaN;
    var t = Date.UTC(nam, thang - 1, ngay), d = new Date(t);
    if (d.getUTCFullYear() !== nam || d.getUTCMonth() !== thang - 1 || d.getUTCDate() !== ngay) return NaN;
    return t / 86400000 + NGAY_GOC_EXCEL;
  }

  // Ngày đọc được khi ô là số seri ngày của Excel, hoặc chữ dạng dd/mm/yyyy, d/m/yyyy, dd-mm-yyyy, dd.mm.yyyy, yyyy-mm-dd
  // có thể kèm giờ. Năm 2 chữ số tính là 20yy. Ngày không có thật như 31/09/2026 là lỗi
  function docNgay(v) {
    if (typeof v === 'number') return isFinite(v) && v >= 1 && v < 2958466 ? Math.floor(v) : NaN;
    if (laTrong(v)) return NaN;
    var s = String(v).trim(), m = NGAY_NTN_NGUOC.exec(s);
    if (m) return seriNgay(+m[1], +m[2], +m[3]);
    m = NGAY_NTN.exec(s);
    if (!m) return NaN;
    return seriNgay(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  }

  // Ngày seri Excel thành chữ dd/mm/yyyy để ghi vào nhật ký và tệp CSV, nhớ sẵn vì số ngày khác nhau ít
  var nhoNgay = new Map();
  function chuNgay(seri) {
    var s = nhoNgay.get(seri);
    if (s !== undefined) return s;
    var d = new Date((seri - NGAY_GOC_EXCEL) * 86400000), p = function (n) { return (n < 10 ? '0' : '') + n; };
    s = p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    if (nhoNgay.size < 100000) nhoNgay.set(seri, s);
    return s;
  }

  // Gom chuỗi lặp lại như tên vật tư, nhà cung cấp về một bản để đỡ tốn bộ nhớ khi dữ liệu lớn
  function gom(st, k, s) {
    var m = st.chuoi[k];
    if (!m || typeof s !== 'string') return s;
    var c = m.get(s);
    if (c !== undefined) return c;
    if (m.size >= GIOI_HAN_GOM) { st.chuoi[k] = null; return s; }
    m.set(s, s);
    return s;
  }

  // Chuẩn hóa một ô theo kiểu cột: null khi trống, { loi } khi không đọc được, còn lại là giá trị ghi ra
  function chuanHoaO(v, cot) {
    if (laTrong(v)) return null;
    var n;
    switch (cot.kieu) {
      case 'so': n = docSo(v); return isNaN(n) ? { loi: LOI_SO } : n;
      case 'ngay': n = docNgay(v); return isNaN(n) ? { loi: LOI_NGAY } : n;
      case 'ma': return String(v).trim().toUpperCase();
      default: return typeof v === 'string' ? v.trim() : v;
    }
  }

  // ---------- Nhận dạng bảng ----------

  // Vị trí của một cột trong cấu hình theo tên, -1 khi không có
  function viTriCot(cot, ten) {
    var t = chuanHoaTen(ten);
    for (var i = 0; t && i < cot.length; i++) if (chuanHoaTen(cot[i].ten) === t) return i;
    return -1;
  }

  // Khớp một dòng với bộ cột chuẩn. viTri[k] là cột trong sheet của cột chuẩn k, -1 khi sheet thiếu cột đó
  // Tên cột so với cả tên chuẩn lẫn dongNghia, cột trong sheet có thể đổi thứ tự
  function khopCot(hdr, st) {
    var viTri = st.cb.cot.map(function () { return -1; }), soKhop = 0, cotLa = [];
    for (var j = 0; hdr && j < hdr.length; j++) {
      var t = chuanHoaTen(hdr[j]);
      if (!t) continue;
      var k = st.banDoTen.has(t) ? st.banDoTen.get(t) : -1;
      if (k >= 0 && viTri[k] < 0) { viTri[k] = j; soKhop++; }
      else cotLa.push(String(hdr[j]).trim());
    }
    return { viTri: viTri, soKhop: soKhop, cotLa: cotLa };
  }

  // Dòng tiêu đề là dòng khớp nhiều cột chuẩn nhất trong 25 dòng đầu, null khi không dòng nào đủ cột chuẩn
  function timDongTieuDe(aoa, st) {
    var tot = null;
    for (var r = 0; r < aoa.length && r < SO_DONG_TIM_TIEU_DE; r++) {
      if (dongTrong(aoa[r])) continue;
      var kc = khopCot(aoa[r], st);
      if (!tot || kc.soKhop > tot.soKhop) { kc.r = r; tot = kc; }
    }
    return tot && tot.soKhop >= SO_COT_TOI_THIEU ? tot : null;
  }

  // Dòng tổng: một trong 8 ô đầu có chữ của dongTong (khớp cả khi chữ đó đứng đầu như "Tổng cộng tháng 9"),
  // đồng thời các cột khóa trừ cột ngày và cột tiền đều trống
  function laDongTong(row, kc, st) {
    // Kiểm cột khóa trước vì nhanh, dòng dữ liệu thường có số phiếu nên dừng ngay ở đây
    for (var i = 0; i < st.khoaDongTong.length; i++) {
      var c = kc.viTri[st.khoaDongTong[i]];
      if (c >= 0 && !laTrong(row[c])) return false;
    }
    for (var j = 0; j < row.length && j < SO_O_TIM_DONG_TONG; j++) {
      if (typeof row[j] !== 'string') continue;
      var t = chuanHoaTen(row[j]);
      for (var w = 0; w < st.dongTong.length; w++) {
        var x = st.dongTong[w];
        if (x && (t === x || t.indexOf(x + ' ') === 0)) return true;
      }
    }
    return false;
  }

  // ---------- Hợp đồng GopFile ----------

  // Tạo trạng thái cho một lần chạy từ cấu hình đang dùng (cauHinh gửi kèm hoặc CAU_HINH_MAU).
  function taoTrangThai(cauHinh) {
    var cb = cauHinh || CAU_HINH_MAU, cot = cb.cot || [], iTien = -1, banDoTen = new Map();
    for (var k = 0; k < cot.length; k++) if (cot[k].cotTien) { iTien = k; break; }
    // Tên chuẩn được ưu tiên trước tên đồng nghĩa khi hai cột trùng tên sau chuẩn hóa
    cot.forEach(function (c, i) { var t = chuanHoaTen(c.ten); if (t && !banDoTen.has(t)) banDoTen.set(t, i); });
    cot.forEach(function (c, i) {
      (c.dongNghia || []).forEach(function (s) { var t = chuanHoaTen(s); if (t && !banDoTen.has(t)) banDoTen.set(t, i); });
    });
    var iKhoa = (cb.khoaTrung || []).map(function (t) { return viTriCot(cot, t); }).filter(function (i) { return i >= 0; });
    return {
      cb: cb, rows: [], doiSoat: [], boQua: [], loi: [], banDoTen: banDoTen, iTien: iTien, iKhoa: iKhoa,
      daGop: new Map(),  // khóa trùng → vị trí dòng đã gộp trong rows
      chuoi: cot.map(function (c) { return c.kieu === 'ma' || c.kieu === 'chu' ? new Map() : null; }),
      dongTong: (cb.dongTong || []).map(chuanHoaTen),
      // Cột khóa phải trống ở dòng tổng, không tính cột ngày và cột tiền
      khoaDongTong: iKhoa.filter(function (i) { return i !== iTien && cot[i].kieu !== 'ngay'; }),
      tuTenTep: cot.map(function (c) { try { return c.layTuTenTep ? new RegExp(c.layTuTenTep, 'i') : null; } catch (e) { return null; } })
    };
  }

  // Xử lý một sheet. aoa là mảng hai chiều từ ThoDoc, dongBatDau là số dòng Excel nằm trước aoa[0].
  // Thêm dòng hợp lệ vào st.rows, thêm một mục đối soát vào st.doiSoat hoặc ghi sheet bỏ qua vào st.boQua.
  function xuLySheet(aoa, tenTep, tenSheet, st, dongBatDau) {
    aoa = aoa || []; dongBatDau = dongBatDau || 0;
    var kc = timDongTieuDe(aoa, st);
    if (!kc) { st.boQua.push({ tep: tenTep, sheet: tenSheet, lyDo: LY_DO_KHONG_TIEU_DE }); return; }
    var cot = st.cb.cot, nCot = cot.length, viTri = kc.viTri, iTien = st.iTien;
    // Cột thiếu mà cấu hình có layTuTenTep thì lấy giá trị từ tên tệp, ví dụ CT03_T09.xlsx cho CT03
    var tuTen = viTri.map(function (c, k) {
      var m = c < 0 && st.tuTenTep[k] ? st.tuTenTep[k].exec(tenTep) : null;
      return m ? (m[1] != null ? m[1] : m[0]) : null;
    });
    var ds = {
      tep: tenTep, sheet: tenSheet, dongTieuDe: kc.r + 1 + dongBatDau, dongDoc: 0, dongGop: 0, dongLoi: 0, dongTrung: 0,
      tongCacDong: 0, tongGop: 0, tongGhi: null, cotLa: kc.cotLa,
      cotThieu: cot.filter(function (c, k) { return viTri[k] < 0; }).map(function (c) { return c.ten; }),
      layTuTen: cot.map(function (c, k) { return tuTen[k] == null ? null : c.ten + ' = ' + tuTen[k]; }).filter(Boolean)
    };
    for (var r = kc.r + 1; r < aoa.length; r++) {
      var row = aoa[r];
      if (dongTrong(row)) continue;
      if (laDongTong(row, kc, st)) {
        var tg = iTien >= 0 && viTri[iTien] >= 0 ? docSo(row[viTri[iTien]]) : NaN;
        if (!isNaN(tg)) ds.tongGhi = (ds.tongGhi || 0) + tg;
        continue;
      }
      ds.dongDoc++;
      var ra = new Array(nCot + 3), loi = null;
      for (var k = 0; k < nCot; k++) {
        var tho = viTri[k] >= 0 ? row[viTri[k]] : tuTen[k], v = chuanHoaO(tho, cot[k]);
        if (v !== null && typeof v === 'object') { if (!loi) loi = { cot: cot[k].ten, giaTri: tho, loi: v.loi }; v = null; }
        else if (v === null && cot[k].batBuoc && !loi) loi = { cot: cot[k].ten, giaTri: tho, loi: LOI_THIEU };
        ra[k] = gom(st, k, v);
      }
      // Tổng các dòng tính cả dòng lỗi miễn là đọc được tiền
      if (iTien >= 0 && typeof ra[iTien] === 'number') ds.tongCacDong += ra[iTien];
      var dongGoc = r + 1 + dongBatDau;
      // Mỗi dòng lỗi ghi một dòng nhật ký với lỗi đầu tiên theo thứ tự cột, dòng lỗi không tham gia so trùng
      if (loi) {
        ds.dongLoi++;
        st.loi.push([tenTep, tenSheet, dongGoc, loi.cot, laTrong(loi.giaTri) ? null : loi.giaTri, loi.loi]);
        continue;
      }
      // Khóa trùng ghép từ giá trị đã chuẩn hóa của các cột khoaTrung, so cả với dòng đã gộp ở tệp khác
      if (st.iKhoa.length) {
        var khoa = st.iKhoa.map(function (i) { return ra[i] == null ? '' : String(ra[i]); }).join('\u0001');
        var truoc = st.daGop.get(khoa);
        if (truoc !== undefined) {
          var goc = st.rows[truoc];
          ds.dongTrung++;
          st.loi.push([tenTep, tenSheet, dongGoc, COT_CA_DONG,
            st.iKhoa.map(function (i) { return ra[i] == null ? '' : cot[i].kieu === 'ngay' ? chuNgay(ra[i]) : String(ra[i]); }).join(' | '),
            'Trùng với ' + goc[nCot] + ' dòng ' + goc[nCot + 2]]);
          continue;
        }
        st.daGop.set(khoa, st.rows.length);
      }
      ra[nCot] = tenTep; ra[nCot + 1] = tenSheet; ra[nCot + 2] = dongGoc;
      st.rows.push(ra);
      ds.dongGop++;
      if (iTien >= 0) ds.tongGop += ra[iTien];
    }
    st.doiSoat.push(ds);
  }

  function chenhLech(ds) { return ds.tongGhi == null ? null : lamTron2(ds.tongGhi - ds.tongCacDong); }

  function ketLuan(ds) {
    var cl = chenhLech(ds);
    return cl == null ? 'Không có dòng tổng' : Math.abs(cl) < 0.5 ? 'ĐẠT' : 'LỆCH, kiểm lại tệp nguồn';
  }

  // Bảng đối soát dạng mảng hai chiều, dòng đầu là 14 tên cột ở mục 6.2.
  function bangDoiSoat(st) {
    var aoa = [COT_DOI_SOAT.slice()], coTien = st.iTien >= 0;
    st.doiSoat.forEach(function (d) {
      aoa.push([d.tep, d.sheet, d.dongTieuDe, d.dongDoc, d.dongGop, d.dongLoi, d.dongTrung,
        coTien ? lamTron2(d.tongCacDong) : null, d.tongGhi == null ? null : lamTron2(d.tongGhi), chenhLech(d), ketLuan(d),
        d.cotLa.join(', ') || null, d.cotThieu.join(', ') || null, d.layTuTen.join(', ') || null]);
    });
    st.boQua.forEach(function (b) {
      aoa.push([b.tep, b.sheet || null, null, null, null, null, null, null, null, null, 'Bỏ qua, ' + b.lyDo, null, null, null]);
    });
    return aoa;
  }

  // Số liệu tóm tắt cho giao diện và bộ kiểm thử, gồm soSheet, soBoQua, dongGop, dongLoi, dongTrung, tongTienGop, soLech.
  function thongKe(st) {
    var tk = { soSheet: st.doiSoat.length, soBoQua: st.boQua.length, dongGop: st.rows.length, dongLoi: 0, dongTrung: 0, tongTienGop: 0, soLech: 0 };
    st.doiSoat.forEach(function (d) {
      tk.dongLoi += d.dongLoi; tk.dongTrung += d.dongTrung; tk.tongTienGop += d.tongGop;
      var cl = chenhLech(d);
      if (cl != null && Math.abs(cl) >= 0.5) tk.soLech++;
    });
    tk.tongTienGop = lamTron2(tk.tongTienGop);
    return tk;
  }

  // 15 tên cột của tệp dữ liệu gộp, gồm 12 cột của cấu hình đang dùng rồi đến COT_NGUON.
  function tieuDeRa(st) { return st.cb.cot.map(function (c) { return c.ten; }).concat(COT_NGUON); }

  // Bảng tổng hợp thành tiền các dòng gộp: hàng và cột theo tongHop của cấu hình, cả hai sắp tăng dần,
  // thêm cột Cộng và dòng Cộng ở cuối. null khi cấu hình thiếu cột hàng, cột hoặc cột tiền
  function bangTongHop(st) {
    var cot = st.cb.cot, th = st.cb.tongHop || {}, iH = viTriCot(cot, th.hang), iC = viTriCot(cot, th.cot), iT = st.iTien;
    if (iH < 0 || iC < 0 || iT < 0) return null;
    var nhan = function (v) { return v == null || v === '' ? '(trống)' : String(v); };
    var o = new Map(), dsHang = new Set(), dsCot = new Set();
    st.rows.forEach(function (r) {
      var h = nhan(r[iH]), c = nhan(r[iC]), k = h + '\u0001' + c;
      o.set(k, (o.get(k) || 0) + (r[iT] || 0));
      dsHang.add(h); dsCot.add(c);
    });
    var tang = function (a, b) { return a < b ? -1 : a > b ? 1 : 0; };
    var hang = Array.from(dsHang).sort(tang), cotTH = Array.from(dsCot).sort(tang);
    var aoa = [[cot[iH].ten + ' \\ ' + cot[iC].ten].concat(cotTH, ['Cộng'])];
    var congCot = cotTH.map(function () { return 0; }), cong = 0;
    hang.forEach(function (h) {
      var dong = [h], s = 0;
      cotTH.forEach(function (c, j) { var v = o.get(h + '\u0001' + c) || 0; dong.push(lamTron2(v)); s += v; congCot[j] += v; });
      dong.push(lamTron2(s)); cong += s;
      aoa.push(dong);
    });
    aoa.push(['Cộng'].concat(congCot.map(lamTron2), [lamTron2(cong)]));
    return aoa;
  }

  // Ô CSV có dấu phẩy, ngoặc kép hoặc xuống dòng thì bọc trong ngoặc kép
  function oCSV(v) {
    if (v == null) return '';
    var s = String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  // Tệp CSV UTF-8 có BOM, ngăn bằng dấu phẩy, 15 cột như mục 6.1, ngày dạng dd/mm/yyyy.
  // Trả về mảng chuỗi để ghép Blob, mỗi chuỗi một khối khoảng 20.000 dòng thay vì một chuỗi khổng lồ
  function taoCSV(st) {
    var cot = st.cb.cot, laNgay = cot.map(function (c) { return c.kieu === 'ngay'; });
    var ra = [], khoi = ['﻿' + tieuDeRa(st).map(oCSV).join(',')];
    for (var i = 0; i < st.rows.length; i++) {
      var r = st.rows[i], o = new Array(r.length);
      for (var k = 0; k < r.length; k++) o[k] = laNgay[k] && typeof r[k] === 'number' ? chuNgay(r[k]) : oCSV(r[k]);
      khoi.push(o.join(','));
      if (khoi.length >= SO_DONG_KHOI_CSV) { ra.push(khoi.join('\r\n') + '\r\n'); khoi = []; }
    }
    if (khoi.length) ra.push(khoi.join('\r\n') + '\r\n');
    return ra;
  }

  root.GopFile = {
    CAU_HINH_MAU: CAU_HINH_MAU, COT_NGUON: COT_NGUON, GIOI_HAN_DONG: GIOI_HAN_DONG,
    taoTrangThai: taoTrangThai, xuLySheet: xuLySheet, bangDoiSoat: bangDoiSoat, thongKe: thongKe, tieuDeRa: tieuDeRa,
    bangTongHop: bangTongHop, taoCSV: taoCSV
  };
})(typeof self !== 'undefined' ? self : globalThis);
