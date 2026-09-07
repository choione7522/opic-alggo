// ── OPIc AL go go Apps Script ──
// 구글 시트 > 확장 프로그램 > Apps Script > 코드 전체 교체 후 저장
// ★ 저장 후 반드시 "배포 > 새 배포"로 새 버전 배포!

var SHEET_NAME = '전체';

function progressToStar(p) {
  var s = String(p || '').trim().toUpperCase();
  if (s === 'PASS') return 3;
  if (s === 'MORE') return 2;
  if (s === 'NG')   return 1;
  return 0;
}
function starToProgress(star) {
  if (star === 3) return 'PASS';
  if (star === 2) return 'MORE';
  if (star === 1) return 'NG';
  return '';
}
function makeResponse(e, data) {
  var json = JSON.stringify(data);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
function isIndexRow(val) {
  if (val === null || val === undefined || val === '') return false;
  var n = parseFloat(String(val).trim());
  return !isNaN(n) && isFinite(n) && n >= 0 && n === Math.floor(n);
}

// Rich Text → HTML (색상 있는 텍스트만 span으로 감쌈)
function richToHtml(richText) {
  if (!richText) return '';
  try {
    var runs = richText.getRuns();
    if (!runs || runs.length === 0) return escHtml(richText.getText());
    var html = '';
    for (var i = 0; i < runs.length; i++) {
      var text = runs[i].getText();
      if (!text) continue;
      var color = null;
      try { color = runs[i].getTextStyle().getForegroundColor(); } catch(e) {}
      if (color && color !== '#000000' && color !== '#000') {
        html += '<span style="color:' + color + '">' + escHtml(text) + '</span>';
      } else {
        html += escHtml(text);
      }
    }
    return html;
  } catch(e) {
    try { return escHtml(richText.getText()); } catch(e2) { return ''; }
  }
}
function escHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function doGet(e) {
  try {
    // action=setStar: 별점 업데이트 (JSONP GET)
    if (e && e.parameter && e.parameter.action === 'setStar') {
      var qid  = parseInt(e.parameter.id);
      var star = parseInt(e.parameter.star);
      var ss2  = SpreadsheetApp.getActiveSpreadsheet();
      var sh2  = ss2.getSheetByName(SHEET_NAME);
      var rows = sh2.getLastRow();
      var vals = sh2.getRange(1, 1, rows, 4).getValues();
      for (var k = 1; k < vals.length; k++) {
        if (isIndexRow(vals[k][0]) && parseInt(String(vals[k][0]).trim()) === qid) {
          sh2.getRange(k + 1, 4).setValue(starToProgress(star));
          break;
        }
      }
      return makeResponse(e, {ok: true});
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      var names = ss.getSheets().map(function(s){ return s.getName(); });
      return makeResponse(e, {ok: false, error: '시트 없음: ' + SHEET_NAME, availableSheets: names});
    }

    var lastRow = sheet.getLastRow();

    // 1단계: 값만 읽기 (빠름)
    var values = sheet.getRange(1, 1, lastRow, 5).getValues();

    // 2단계: script 셀만 Rich Text로 읽기
    // script 셀 행 번호 목록 수집 (문제 타이틀행 + 2)
    var scriptRows = [];  // 1-indexed 행 번호
    for (var i = 1; i < values.length; i++) {
      if (isIndexRow(values[i][0])) {
        var subjVal = String(values[i][1] || '').trim();
        if (subjVal && subjVal !== 'Subject') {
          var scriptRowIdx = i + 2;  // 0-indexed
          if (scriptRowIdx < values.length) {
            scriptRows.push(scriptRowIdx + 1);  // 1-indexed for getRange
          }
        }
      }
    }

    // script 셀(C열=3번째)만 Rich Text 배치 읽기
    // 연속된 개별 셀보다 getRange로 각각 읽는 게 안정적
    var scriptHtmlMap = {};  // key: 1-indexed row → html
    for (var r = 0; r < scriptRows.length; r++) {
      var rowNum = scriptRows[r];
      try {
        var rt = sheet.getRange(rowNum, 3).getRichTextValue();
        scriptHtmlMap[rowNum] = richToHtml(rt);
      } catch(e2) {
        scriptHtmlMap[rowNum] = escHtml(String(values[rowNum-1][2] || ''));
      }
    }

    // 문제 데이터 조립
    var questions = [];
    var subjMap = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (!isIndexRow(row[0])) continue;
      var subj = String(row[1] || '').trim();
      if (!subj || subj === 'Subject') continue;

      var questionText = (i+1 < values.length) ? String(values[i+1][2] || '') : '';
      var scriptRowNum = i + 3;  // 1-indexed
      var scriptHtml = scriptHtmlMap[scriptRowNum] || escHtml(String((i+2 < values.length) ? values[i+2][2] : ''));

      subjMap[subj] = true;
      questions.push({
        id:       parseInt(String(row[0]).trim()),
        subject:  subj,
        title:    String(row[2] || '').trim(),
        question: questionText,
        script:   scriptHtml,
        star:     progressToStar(String(row[3] || '').trim())
      });
    }

    // 대표유형 시트
    var tmplSheet = ss.getSheetByName('대표 유형');
    var templates = [];
    if (tmplSheet) {
      var tmplData = tmplSheet.getDataRange().getValues();
      var curName = '', curScript = [];
      for (var j = 0; j < tmplData.length; j++) {
        var nm = String(tmplData[j][0] || '').trim();
        var sc = String(tmplData[j][1] || '');
        if (nm) {
          if (curName) templates.push({name: curName, script: curScript.join('\n')});
          curName = nm;
          curScript = sc ? [sc] : [];
        } else if (sc) {
          curScript.push(sc);
        }
      }
      if (curName) templates.push({name: curName, script: curScript.join('\n')});
    }

    return makeResponse(e, {
      ok: true,
      count: questions.length,
      subjects: Object.keys(subjMap).length,
      questions: questions,
      templates: templates,
      richText: true
    });

  } catch(err) {
    return makeResponse(e, {ok: false, error: err.message});
  }
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var qid  = parseInt(params.id);
    var star = parseInt(params.star);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data  = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
    for (var i = 1; i < data.length; i++) {
      if (isIndexRow(data[i][0]) && parseInt(String(data[i][0]).trim()) === qid) {
        sheet.getRange(i + 1, 4).setValue(starToProgress(star));
        break;
      }
    }
    return makeResponse(e, {ok: true});
  } catch(err) {
    return makeResponse(e, {ok: false, error: err.message});
  }
}
