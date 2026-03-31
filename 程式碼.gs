var SPREADSHEET_ID = '1EIxN0W67Ph6BuqlHhbxW1NCoJ3TvgWvL_W4GofbTSyg';
var FOLDER_ID = '1CEhx5uJkUysIAHhhaVWPBdmh8TSaHzMH';

// ── 快取：同一次執行內重用同一個 Spreadsheet 物件 ──────────────
var _ss = null;
function getSpreadsheet() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('資遣流程管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(name) {
  var sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('找不到工作表：' + name);
  return sheet;
}

function buildHeaderMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = headers[i].toString().trim();
    if (key) map[key] = i + 1;
  }
  return map;
}

function rowToObj(headers, row) {
  var obj = {};
  for (var j = 0; j < headers.length; j++) {
    if (row[j] instanceof Date) {
      obj[headers[j]] = Utilities.formatDate(row[j], 'GMT+8', 'yyyy/MM/dd HH:mm');
    } else {
      obj[headers[j]] = (row[j] !== null && row[j] !== undefined) ? row[j].toString() : '';
    }
  }
  return obj;
}

// ── 讀取所有案件（統一回傳 { success, data } 格式）──────────────
function getAllActiveCases() {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    var headers = data[0].map(function(h) { return h.toString().trim(); });

    if (headers.indexOf('目前狀態') === -1) {
      return { success: false, error: '找不到目前狀態欄位，目前欄位：' + headers.join('、') };
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] || row[0].toString().trim() === '') continue;
      result.push(rowToObj(headers, row));
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: '後台執行失敗：' + e.toString() };
  }
}

function getHRDashboardData() {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var statusIdx = headers.indexOf('目前狀態');

    if (statusIdx === -1) {
      return { success: false, error: '找不到目前狀態欄位，目前欄位：' + headers.join('、') };
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] || row[0].toString().trim() === '') continue;
      var status = row[statusIdx] ? row[statusIdx].toString().replace(/\s+/g, '') : '';
      if (status === '待人資回覆') {
        result.push(rowToObj(headers, row));
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: '後台執行失敗：' + e.toString() };
  }
}

function getCaseData(caseId) {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var searchId = caseId.toString().trim();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === searchId) {
        return { success: true, data: rowToObj(headers, data[i]) };
      }
    }
    return { success: true, data: null }; // 查無此案件
  } catch (e) {
    return { success: false, error: '查詢失敗：' + e.toString() };
  }
}

// ── 處理步驟（新增 / 更新案件）───────────────────────────────────
function processStep(formData, fileData) {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('案件總表');
    var logSheet    = ss.getSheetByName('操作紀錄');
    var now = new Date();

    var caseId  = formData.caseId ? formData.caseId.toString().trim() : '';
    var fileUrl = formData.existingFileUrl || '';
    // ★ 加入「待人資再審」的判斷，確保補件寫入「主管補充附件」不覆蓋原檔
    var isManagerSupplemental = (formData.nextStatus === '待老闆同意' || formData.nextStatus === '待人資再審');

    // ★ 修改點：移除 setSharing，避免企業版權限阻擋，並整合多檔案處理
    if (fileData) {
      var folder = DriveApp.getFolderById(FOLDER_ID);
      
      if (Array.isArray(fileData)) {
        var urls = [];
        for (var f = 0; f < fileData.length; f++) {
          var blob = Utilities.newBlob(Utilities.base64Decode(fileData[f].data), fileData[f].type, fileData[f].name);
          var file = folder.createFile(blob);
          // 拿掉 file.setSharing，因為資料夾已經是公開的了，檔案會自動繼承
          urls.push(file.getUrl());
        }
        fileUrl = urls.join('\n');
      } else if (fileData.data) {
        var blobSingle = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.type, fileData.name);
        var fileSingle = folder.createFile(blobSingle);
        // 拿掉 file.setSharing，因為資料夾已經是公開的了，檔案會自動繼承
        fileUrl = fileSingle.getUrl();
      }
    }

    var hMap = buildHeaderMap(masterSheet);
    var requiredCols = ['案件編號', '目前狀態', '最後更新時間'];
    for (var r = 0; r < requiredCols.length; r++) {
      if (!hMap[requiredCols[r]]) {
        throw new Error('試算表缺少必要欄位：' + requiredCols[r]);
      }
    }

    var allData = masterSheet.getDataRange().getValues();
    var rowIndex = -1;
    if (caseId) {
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][0].toString().trim() === caseId) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex === -1) {
      // ── 新案件 ────────────────────────────────────────────────
      var dateStr = Utilities.formatDate(now, 'GMT+8', 'yyyyMMdd');
      var seq = ('000' + masterSheet.getLastRow()).slice(-3);
      caseId = 'TERM-' + dateStr + '-' + seq;

      var totalCols = masterSheet.getLastColumn();
      var newRow = [];
      for (var n = 0; n < totalCols; n++) newRow.push('');

      if (hMap['案件編號'])     newRow[hMap['案件編號'] - 1]     = caseId;
      if (hMap['目前狀態'])     newRow[hMap['目前狀態'] - 1]     = formData.nextStatus || '待人資回覆';
      if (hMap['所屬公司'])     newRow[hMap['所屬公司'] - 1]     = formData.company || '';
      if (hMap['單位'])         newRow[hMap['單位'] - 1]         = formData.unit || '';
      if (hMap['申請主管'])     newRow[hMap['申請主管'] - 1]     = formData.applicant || '';
      if (hMap['被資遣員工'])   newRow[hMap['被資遣員工'] - 1]   = formData.employeeName || '';
      if (hMap['資遣原因'])     newRow[hMap['資遣原因'] - 1]     = formData.reason || '';
      if (hMap['附件連結'])     newRow[hMap['附件連結'] - 1]     = fileUrl;
      if (hMap['最後更新時間']) newRow[hMap['最後更新時間'] - 1] = now;

      masterSheet.appendRow(newRow);

    } else {
      // ── 更新現有案件 ──────────────────────────────────────────
      var updates = {};
      function stage(col, val) {
        if (hMap[col] && val !== undefined && val !== null && val.toString().trim() !== '') {
          updates[hMap[col]] = val;
        }
      }

      stage('目前狀態',     formData.nextStatus);
      stage('風險等級',     formData.risk);
      stage('建議方案',     formData.hrPlan);
      stage('尚缺資料',     formData.missingDocs);
      stage('主管確認簽署', formData.managerSign);
      stage('老闆審核',     formData.bossSign);
      stage('執行結果',     formData.executionResult);
      updates[hMap['最後更新時間']] = now;

      if (fileUrl) {
        // 如果是主管確認階段 (isManagerSupplemental 為 true)
        if (isManagerSupplemental) {
          // 只寫入「主管補充附件」欄位。如果試算表沒有這個欄位，請在 G 欄(附件連結) 右邊手動新增一欄名為「主管補充附件」
          if (hMap['主管補充附件']) {
            updates[hMap['主管補充附件']] = fileUrl;
          }
        } 
        // 只有非主管補充階段，才寫入「附件連結」(初始附件)
        else if (hMap['附件連結']) {
          updates[hMap['附件連結']] = fileUrl;
        }
      }

      for (var col in updates) {
        masterSheet.getRange(rowIndex, parseInt(col)).setValue(updates[col]);
      }
    }

    var operator = formData.applicant || formData.managerSign || formData.bossSign || '系統';
    logSheet.appendRow([
      now,
      caseId,
      operator,
      '變更至：' + (formData.nextStatus || ''),
      formData.missingDocs || formData.executionResult || ''
    ]);

    return { success: true, caseId: caseId };

  } catch (e) {
    throw e.toString();
  }
}
// ── 撤回送件 ───────────────────────────────────
function recallCase(caseId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('案件總表');
    var logSheet = ss.getSheetByName('操作紀錄');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    
    var idIdx = headers.indexOf('案件編號');
    var statusIdx = headers.indexOf('目前狀態');
    var signIdx = headers.indexOf('主管確認簽署');
    var updateIdx = headers.indexOf('最後更新時間');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx].toString().trim() === caseId) {
        var currentStatus = data[i][statusIdx];
        if (currentStatus === '待老闆同意' || currentStatus === '待人資再審') {
          // 狀態退回，並清除主管的電子簽名
          sheet.getRange(i + 1, statusIdx + 1).setValue('待主管確認');
          if (signIdx !== -1) sheet.getRange(i + 1, signIdx + 1).setValue('');
          if (updateIdx !== -1) sheet.getRange(i + 1, updateIdx + 1).setValue(new Date());

          logSheet.appendRow([new Date(), caseId, '主管', '撤回送件', '重新編輯']);
          return { success: true };
        } else {
          return { success: false, error: '案件目前狀態無法撤回' };
        }
      }
    }
    return { success: false, error: '找不到案件' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}