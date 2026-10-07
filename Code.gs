/**
 * 부산 상가 철거·원상복구 리드 접수용 Google Apps Script
 *
 * 역할:
 * GitHub Pages의 index.html에서 POST로 받은 고객 정보를
 * Google Sheets에 저장하고, 사진은 Google Drive에 저장합니다.
 */

const CONFIG = {
  SPREADSHEET_ID: "114GDKFNyJ5mPie6KjrI2LA8uY6a6Re9VpSRS0WjRvB4",
  SHEET_NAME: "Leads",
  DRIVE_FOLDER_NAME: "철거견적_현장사진",
  TIMEZONE: "Asia/Seoul",

  PROVIDER_NAME: "부산철거 논스톱철거 -영도구-",
  PROVIDER_PHONE: "010-2984-1115"
};

/**
 * 최초 1회 실행:
 * Apps Script 편집기에서 setup()을 선택하고 실행하세요.
 */
function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  const headers = [
    "접수일시",
    "접수번호",
    "상태",
    "지역",
    "현장주소",
    "업종",
    "규모",
    "철거범위",
    "예정시기",
    "추가설명",
    "고객명",
    "연락처",
    "동의",
    "사진URL",
    "배정업체",
    "업체연락처"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  getOrCreatePhotoFolder_();

  return "설정 완료";
}

/**
 * Apps Script 웹앱 직접 접속 시 확인용 화면
 */
function doGet() {
  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html lang="ko"><meta charset="utf-8">' +
      '<body style="font-family:sans-serif;padding:30px">' +
      '<h2>부산 철거 견적 접수 서버 정상</h2>' +
      '<p>이 주소는 고객 페이지가 아니라 접수 서버입니다.</p>' +
      '</body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * GitHub Pages → Apps Script POST
 */
function doPost(e) {
  let result;

  try {
    if (!e || !e.parameter) {
      throw new Error("전송된 데이터가 없습니다.");
    }

    const p = e.parameter;

    if (p.consent !== "on" && p.consent !== "true") {
      throw new Error("개인정보 수집·이용 및 제3자 제공 동의가 필요합니다.");
    }

    const required = [
      ["name", "성함"],
      ["phone", "휴대전화 번호"],
      ["region", "지역"],
      ["address", "현장 주소"],
      ["businessType", "업종"],
      ["size", "매장 규모"],
      ["schedule", "철거 예정 시기"],
      ["demolitionText", "철거 범위"]
    ];

    for (const item of required) {
      if (!String(p[item[0]] || "").trim()) {
        throw new Error(item[1] + "을(를) 입력해주세요.");
      }
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      writeHeaders_(sheet);
    }

    const leadId = makeLeadId_(sheet);
    const now = new Date();

    let photoUrls = [];

    const photoCount = Math.min(
      Math.max(parseInt(p.photoCount || "0", 10) || 0, 0),
      3
    );

    if (photoCount > 0) {
      const photos = [];

      for (let i = 0; i < photoCount; i++) {
        const data = p["photoData" + i];
        if (!data) continue;

        photos.push({
          name: p["photoName" + i] || ("photo_" + i + ".jpg"),
          mimeType: p["photoMime" + i] || "image/jpeg",
          data: data
        });
      }

      if (photos.length) {
        photoUrls = savePhotos_(leadId, photos);
      }
    }

    sheet.appendRow([
      Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss"),
      leadId,
      "신규",
      clean_(p.region),
      clean_(p.address),
      clean_(p.businessType),
      clean_(p.size),
      clean_(p.demolitionText),
      clean_(p.schedule),
      clean_(p.notes),
      clean_(p.name),
      clean_(p.phone),
      "동의",
      photoUrls.join("\n"),
      CONFIG.PROVIDER_NAME,
      CONFIG.PROVIDER_PHONE
    ]);

    result = {
      source: "busan-demolition",
      ok: true,
      leadId: leadId
    };

  } catch (err) {
    result = {
      source: "busan-demolition",
      ok: false,
      message: String(err && err.message ? err.message : err)
    };
  }

  return makeResultPage_(result);
}

/**
 * 헤더 생성
 */
function writeHeaders_(sheet) {
  const headers = [
    "접수일시",
    "접수번호",
    "상태",
    "지역",
    "현장주소",
    "업종",
    "규모",
    "철거범위",
    "예정시기",
    "추가설명",
    "고객명",
    "연락처",
    "동의",
    "사진URL",
    "배정업체",
    "업체연락처"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

/**
 * BUS-0001, BUS-0002... 생성
 */
function makeLeadId_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return "BUS-0001";
  }

  const values = sheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat();

  let max = 0;

  values.forEach(v => {
    const m = String(v || "").match(/^BUS-(\d+)$/);
    if (m) {
      max = Math.max(max, parseInt(m[1], 10));
    }
  });

  return "BUS-" + String(max + 1).padStart(4, "0");
}

/**
 * 사진 저장 폴더
 */
function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

/**
 * 압축된 Base64 사진을 Drive에 저장
 */
function savePhotos_(leadId, photos) {
  const folder = getOrCreatePhotoFolder_();
  const urls = [];

  photos.forEach((photo, index) => {
    try {
      const bytes = Utilities.base64Decode(photo.data);
      const blob = Utilities.newBlob(
        bytes,
        photo.mimeType || "image/jpeg",
        sanitizeFileName_(leadId + "_" + index + "_" + photo.name)
      );

      const file = folder.createFile(blob);

      // 링크를 아는 사람에게 공개하지 않고, 본인 Drive 권한 기준으로 유지합니다.
      urls.push(file.getUrl());

    } catch (err) {
      // 사진 하나가 실패해도 고객 접수 자체는 계속 진행합니다.
      urls.push("사진 저장 실패: " + String(err.message || err));
    }
  });

  return urls;
}

/**
 * Apps Script → GitHub iframe으로 결과 전달
 */
function makeResultPage_(result) {
  const payload = JSON.stringify(result)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026");

  const html =
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'window.parent.postMessage(' + payload + ', "*");' +
    '</script>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 셀에 들어가는 문자열 정리
 */
function clean_(value) {
  return String(value == null ? "" : value).trim();
}

/**
 * 파일명 정리
 */
function sanitizeFileName_(name) {
  return String(name || "photo.jpg")
    .replace(/[\\\/:*?"<>|]/g, "_")
    .slice(0, 180);
}
