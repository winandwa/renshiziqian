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

// ── 讀取所有案件 ──────────────
function getAllActiveCases() {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === '') continue;
      result.push(rowToObj(headers, data[i]));
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: '後台執行失敗：' + e.toString() };
  }
}

// ── 人資後台：撈取「待回覆」與「待再審」的案件 ──────────────
function getHRDashboardData() {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var statusIdx = headers.indexOf('目前狀態');
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === '') continue;
      var status = data[i][statusIdx] ? data[i][statusIdx].toString().replace(/\s+/g, '') : '';
      // ★ 同步 Index.html 的邏輯，讓再審案件也能在後台出現
      if (status === '待人資回覆' || status === '待人資再審') {
        result.push(rowToObj(headers, data[i]));
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: '後台執行失敗：' + e.toString() };
  }
}

function getCaseData(searchKey) {
  try {
    var sheet = getSheet('案件總表');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var key = searchKey.toString().trim().toLowerCase();
    
    var nameIdx = headers.indexOf('被資遣員工');

    for (var i = 1; i < data.length; i++) {
      // 比對 A 欄 (案件編號) 或 F 欄 (被資遣員工姓名)
      var idMatch = data[i][0].toString().trim().toLowerCase() === key;
      var nameMatch = (nameIdx !== -1) && data[i][nameIdx].toString().trim().toLowerCase() === key;
      
      if (idMatch || nameMatch) {
        return { success: true, data: rowToObj(headers, data[i]) };
      }
    }
    return { success: true, data: null };
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
    var isManagerSupplemental = (formData.nextStatus === '待老闆同意' || formData.nextStatus === '待人資再審');

    // ★ 處理線上填寫的 PIP 資料並轉為 PDF
    if (formData.pipData) {
      var folder = DriveApp.getFolderById(FOLDER_ID);
      var pipFileUrl = generatePipPdf(folder, formData.pipData, formData.employeeName || '未命名員工');
      fileUrl = fileUrl ? fileUrl + '\n' + pipFileUrl : pipFileUrl;
    }

    // 處理檔案上傳
    if (fileData) {
      var folder = DriveApp.getFolderById(FOLDER_ID);
      if (Array.isArray(fileData)) {
        var urls = [];
        for (var f = 0; f < fileData.length; f++) {
          var blob = Utilities.newBlob(Utilities.base64Decode(fileData[f].data), fileData[f].type, fileData[f].name);
          var file = folder.createFile(blob);
          urls.push(file.getUrl());
        }
        fileUrl = fileUrl ? fileUrl + '\n' + urls.join('\n') : urls.join('\n');
      } else if (fileData.data) {
        var blobSingle = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.type, fileData.name);
        var fileSingle = folder.createFile(blobSingle);
        fileUrl = fileUrl ? fileUrl + '\n' + fileSingle.getUrl() : fileSingle.getUrl();
      }
    }

    var hMap = buildHeaderMap(masterSheet);
    var allData = masterSheet.getDataRange().getValues();
    var rowIndex = -1;
    if (caseId) {
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][0].toString().trim() === caseId) { rowIndex = i + 1; break; }
      }
    }

    if (rowIndex === -1) {
      // ── 新案件 ──
      var dateStr = Utilities.formatDate(now, 'GMT+8', 'yyyyMMdd');
      var seq = ('000' + masterSheet.getLastRow()).slice(-3);
      caseId = 'TERM-' + dateStr + '-' + seq;
      var newRow = new Array(masterSheet.getLastColumn()).fill('');
      if (hMap['案件編號'])     newRow[hMap['案件編號'] - 1]     = caseId;
      if (hMap['目前狀態'])     newRow[hMap['目前狀態'] - 1]     = formData.nextStatus || '待人資回覆';
      if (hMap['所屬公司'])     newRow[hMap['所屬公司'] - 1]     = formData.company || '';
      if (hMap['單位'])         newRow[hMap['單位'] - 1]         = formData.unit || '';
      if (hMap['申請主管'])     newRow[hMap['申請主管'] - 1]     = formData.applicant || '';
      if (hMap['被資遣員工'])   newRow[hMap['被資遣員工'] - 1]   = formData.employeeName || '';
      if (hMap['資遣原因'])     newRow[hMap['資遣原因'] - 1]     = formData.reason || '';
      if (hMap['附件連結'])     newRow[hMap['附件連結'] - 1]     = fileUrl;
      if (hMap['最後更新時間']) newRow[hMap['最後更新時間'] - 1] = now;
      if (hMap['預計資遣日'])   newRow[hMap['預計資遣日'] - 1]   = formData.estDate || '';
      if (hMap['主管申請說明']) newRow[hMap['主管申請說明'] - 1] = formData.managerDesc || '';
      // ★ 紀錄申請時間
      if (hMap['主管申請時間']) newRow[hMap['主管申請時間'] - 1] = now;
      masterSheet.appendRow(newRow);
    } else {
      // ── 更新案件 ──
      var updates = {};
      function stage(col, val) { if (hMap[col] && val) updates[hMap[col]] = val; }
      stage('目前狀態',     formData.nextStatus);
      stage('風險等級',     formData.risk);
      stage('建議方案',     formData.hrPlan);
      stage('尚缺資料',     formData.missingDocs);
      stage('主管確認簽署', formData.managerSign);
      stage('老闆審核',     formData.bossSign);
      stage('執行結果',     formData.executionResult);
      stage('預計資遣日',   formData.estDate);
      stage('主管申請說明', formData.managerDesc);
      
      // ★ 紀錄各階段 HR 與主管的操作時間點
      if (formData.nextStatus === '待人資回覆' && !allData[rowIndex-1][hMap['主管申請時間']-1]) stage('主管申請時間', now);
      if (formData.nextStatus === '待主管確認') stage('HR評估時間', now);
      if (formData.nextStatus === '待老闆同意') stage('主管簽署時間', now);
      if (formData.nextStatus === '待人資再審') stage('主管簽署時間', now); 
      
      // 偵測是否為再審完成：若原狀態含有「再審」且下一步是「送老闆」，紀錄再審時間
      var currentStatusInSheet = rowIndex !== -1 ? masterSheet.getRange(rowIndex, hMap['目前狀態']).getValue().toString() : "";
      if (currentStatusInSheet.indexOf('再審') !== -1 && formData.nextStatus === '待老闆同意') {
        stage('HR再審時間', now);
      }
      
      if (formData.bossSign) stage('老闆核准時間', now);
      if (formData.nextStatus === '已結案') stage('執行結案時間', now);

      updates[hMap['最後更新時間']] = now;

      if (fileUrl) {
        if (isManagerSupplemental && hMap['主管補充附件']) {
          updates[hMap['主管補充附件']] = fileUrl;
        } else if (hMap['附件連結']) {
          updates[hMap['附件連結']] = fileUrl;
        }
      }
      for (var col in updates) masterSheet.getRange(rowIndex, parseInt(col)).setValue(updates[col]);
    }

    logSheet.appendRow([now, caseId, formData.applicant || '系統', '變更至：' + formData.nextStatus, formData.missingDocs || '']);
    return { success: true, caseId: caseId };
  } catch (e) { throw e.toString(); }
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

    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx].toString().trim() === caseId) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('待主管確認');
        if (signIdx !== -1) sheet.getRange(i + 1, signIdx + 1).setValue('');
        logSheet.appendRow([new Date(), caseId, '主管', '撤回送件', '重新編輯']);
        return { success: true };
      }
    }
    return { success: false, error: '找不到案件' };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── 自動產生 PIP PDF (獨立函式，不可放在 recallCase 內) ───────────────────────
function generatePipPdf(folder, pip, empName) {
  var doc = DocumentApp.create('PIP_績效改善計畫_' + empName + '_' + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd"));
  var body = doc.getBody();
  
  body.appendParagraph("績效改善計畫 (PIP)").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("被資遣員工：" + empName);
  body.appendParagraph("輔導期間：" + (pip.start || '未填') + " 至 " + (pip.end || '未填'));
  body.appendHorizontalRule();
  body.appendParagraph("【目前工作表現之具體問題】").setBold(true);
  body.appendParagraph(pip.issue || "無");
  body.appendParagraph("【預期改善目標與行動建議】").setBold(true);
  body.appendParagraph(pip.target || "無");
  body.appendParagraph("【公司提供之資源與協助】").setBold(true);
  body.appendParagraph(pip.support || "無");
  body.appendParagraph("\n\n本表單由系統於 " + Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd") + " 自動產生，內容僅供輔導參考。");

  doc.saveAndClose();
  var pdfBlob = doc.getAs('application/pdf');
  var file = folder.createFile(pdfBlob);
  
  // 如果先前有存取遭拒的問題，下面這行建議拿掉或確保資料夾權限已開
  // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return file.getUrl();
}

function authTrigger() { 
  DriveApp.getFolderById(FOLDER_ID);
  // 先建立文件
  var tempDoc = DocumentApp.create('權限測試文件');
  // 透過 DriveApp 取得該 ID 並丟進垃圾桶
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
}