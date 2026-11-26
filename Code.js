// =========================================================================
// CẤU HÌNH ID CÁC FILE GOOGLE SHEETS
// =========================================================================
const DB_CONFIG = {
  MAIN_ID: '1L4waVVXo2c9WTqylub1f_ie7lbullQMGP1UhLUOadGU',   // VanDongVien, GiaiDau, LichSuThiDau, LichSuDiemTrinh
  TOUR_ID: '1Bu6gAleRUwKfsa9zMsf4oVe5zALnn8X-RNAK8QhA8zc',   // DoiThiDau, TranThiDau
  REF_ID:  '1MkCPS63R1Znb2dz-kV9mdzUreh8ovoNqNKxWNacz0Oc'    // TrongTai, TrongTaiGiaiDau
};

/**
 * Hàm xử lý HTTP GET Request từ Front-end (GitHub Pages)
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getTournaments';
  const maGiaiDau = (e && e.parameter && e.parameter.maGiaiDau) ? e.parameter.maGiaiDau : '';

  let result = {};
  try {
    if (action === 'getTournaments') {
      result = {
        status: 'success',
        data: getVisibleTournaments()
      };
    } else if (action === 'getTournamentData') {
      if (!maGiaiDau) throw new Error('Thiếu tham số maGiaiDau');
      result = {
        status: 'success',
        data: getTournamentDetails(maGiaiDau)
      };
    } else if (action === 'tts') {
      const text = (e && e.parameter && e.parameter.text) ? String(e.parameter.text).trim() : '';
      if (!text) throw new Error('Thiếu nội dung text');
      result = getVietnameseTtsAudio(text);
    } else {
      result = { status: 'error', message: 'Action không hợp lệ' };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Tạo file âm thanh MP3 giọng đọc tiếng Việt chuẩn từ văn bản
 */
function getVietnameseTtsAudio(text) {
  const cleanText = String(text || '').trim().slice(0, 200);
  const url = 'https://translate.google.com/translate_tts?ie=UTF-8&tl=vi&client=tw-ob&q=' + encodeURIComponent(cleanText);
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0'
    }
  });
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Không tạo được âm thanh tiếng Việt (HTTP ' + code + ')');
  }
  const blob = response.getBlob();
  return {
    status: 'success',
    audioBase64: Utilities.base64Encode(blob.getBytes()),
    mimeType: 'audio/mpeg'
  };
}

/**
 * Chuyển đổi dữ liệu Sheet thành mảng Object dựa vào dòng tiêu đề đầu tiên
 */
function getSheetDataAsObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => h.trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rowObj = {};
    let isEmpty = true;
    for (let j = 0; j < headers.length; j++) {
      const val = data[i][j] !== undefined ? String(data[i][j]).trim() : '';
      rowObj[headers[j]] = val;
      if (val !== '') isEmpty = false;
    }
    if (!isEmpty) rows.push(rowObj);
  }
  return rows;
}

/**
 * Chuyển đổi link Google Drive sang định dạng thumbnail nhúng trực tiếp ổn định
 */
function formatDriveImage(url) {
  if (!url) return '';
  const match = url.match(/[-\w]{25,}/);
  if (match && url.includes('drive.google.com')) {
    return 'https://drive.google.com/thumbnail?id=' + match[0] + '&sz=w800';
  }
  return url;
}

/**
 * Lấy danh sách các giải đấu có HienThi === "Hiển thị"
 */
function getVisibleTournaments() {
  const ssMain = SpreadsheetApp.openById(DB_CONFIG.MAIN_ID);
  const tournaments = getSheetDataAsObjects(ssMain, 'GiaiDau');

  return tournaments
    .filter(t => t.HienThi && t.HienThi.toLowerCase() === 'hiển thị')
    .map(t => ({
      MaGiaiDau: t.MaGiaiDau,
      TenGiaiDau: t.TenGiaiDau,
      NgayThiDau: t.NgayThiDau,
      DonViToChuc: t.DonViToChuc,
      MoTa: t.MoTa,
      TrangThai: t.TrangThai,
      Poster: formatDriveImage(t.Poster),
      URL: t.URL
    }));
}

/**
 * Lấy toàn bộ dữ liệu chi tiết của 1 giải đấu được chọn
 */
function getTournamentDetails(maGiaiDau) {
  const ssMain = SpreadsheetApp.openById(DB_CONFIG.MAIN_ID);
  const ssTour = SpreadsheetApp.openById(DB_CONFIG.TOUR_ID);
  const ssRef  = SpreadsheetApp.openById(DB_CONFIG.REF_ID);

  // 1. Lấy thông tin giải đấu
  const allTournaments = getSheetDataAsObjects(ssMain, 'GiaiDau');
  const tournamentInfo = allTournaments.find(t => t.MaGiaiDau === maGiaiDau) || {};

  // 2. Lấy danh sách Đội thi đấu của giải này
  const allTeams = getSheetDataAsObjects(ssTour, 'DoiThiDau');
  const teams = allTeams.filter(d => d.MaGiaiDau === maGiaiDau);

  const neededVdvIds = new Set();
  teams.forEach(t => {
    if (t.MaSoVDV1) neededVdvIds.add(t.MaSoVDV1);
    if (t.MaSoVDV2) neededVdvIds.add(t.MaSoVDV2);
  });

  // 3. Lấy thông tin Vận động viên liên quan
  const allPlayers = getSheetDataAsObjects(ssMain, 'VanDongVien');
  const playersMap = {};
  allPlayers.forEach(p => {
    if (neededVdvIds.has(p.MaSoVDV)) {
      playersMap[p.MaSoVDV] = {
        MaSoVDV: p.MaSoVDV,
        HoTen: p.HoTen,
        BietDanh: p.BietDanh,
        GioiTinh: p.GioiTinh,
        HinhAnhDaiDien: formatDriveImage(p.HinhAnhDaiDien)
      };
    }
  });

  // 4. Lấy danh sách Trọng tài
  const allRefs = getSheetDataAsObjects(ssRef, 'TrongTai');
  const refsMap = {};
  allRefs.forEach(r => {
    refsMap[r.MaTaiKhoan] = r.HoTen;
  });

  // 5. Lấy danh sách Trận thi đấu của giải này
  const allMatches = getSheetDataAsObjects(ssTour, 'TranThiDau');
  const matches = allMatches
    .filter(m => m.MaGiaiDau === maGiaiDau)
    .map(m => ({
      ...m,
      TenTrongTai: refsMap[m.TrongTai] || m.TrongTai || ''
    }));

  // 6. Gắn thông tin VĐV vào từng Đội
  const enrichedTeams = teams.map(t => ({
    ...t,
    VDV1: playersMap[t.MaSoVDV1] || { HoTen: t.MaSoVDV1 || 'Chưa cập nhật', HinhAnhDaiDien: '' },
    VDV2: playersMap[t.MaSoVDV2] || { HoTen: t.MaSoVDV2 || '', HinhAnhDaiDien: '' }
  }));

  return {
    tournament: tournamentInfo,
    teams: enrichedTeams,
    matches: matches
  };
}