/**
 * 공간잇기 — 성찰로그 수합 + 대시보드 Apps Script
 * ════════════════════════════════════════════════════
 * 1) 모듈1(삼수선)·모듈2(정사영) 성찰로그를 각 시트에 분리 기록
 * 2) '대시보드' 탭에 제출 현황·사고 유형 분포·자유 서술을 자동 요약
 *
 * ── 설치 ──
 * 1. 성찰로그를 모을 Google 스프레드시트 → 확장 프로그램 → Apps Script
 * 2. 이 코드 전체를 붙여넣고 저장
 * 3. 배포 → 새 배포 → 유형: 웹 앱
 *      실행 사용자: 나 / 액세스: 모든 사용자
 * 4. 웹 앱 URL을 samsuseon.html · jeongsayeong.html 의 SHEET_URL 에 입력
 *
 * ── 대시보드 보기 ──
 * 스프레드시트 상단 메뉴에 '📊 대시보드' 가 생깁니다 (새로고침하면 나타남)
 *   → '대시보드 새로고침' 클릭하면 최신 데이터로 갱신됩니다
 */

// ───────────────────────────────────────────────
// 설정: 사고 유형 4가지 (두 모듈 공통 구조)
// ───────────────────────────────────────────────
var THINK_TYPES = ['🔢', '🔁', '👁', '📋'];
var TYPE_LABEL = {
  '🔢': '논리·공식형',
  '🔁': '탐색·시행착오형',
  '👁': '직관·관찰형',
  '📋': '절차·분석형'
};

// ───────────────────────────────────────────────
// 1) 성찰로그 수신 → 모듈별 시트에 기록
// ───────────────────────────────────────────────
function doPost(e) {
  try {
    var p = e.parameter;
    var sheetName = (p.module && p.module.indexOf('M2') === 0) ? '모듈2' : '모듈1';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['제출시각', '모듈', '학번', '이름', '사고유형', '자유서술']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    sheet.appendRow([
      new Date(),
      p.module || '',
      p.sid    || '',
      p.sname  || '',
      p.method || '',
      p.extra  || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('공간잇기 성찰로그 수합 스크립트 정상 작동 중.');
}

// ───────────────────────────────────────────────
// 2) 메뉴 등록 (스프레드시트 열 때)
// ───────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 대시보드')
    .addItem('대시보드 새로고침', 'buildDashboard')
    .addToUi();
}

// ───────────────────────────────────────────────
// 3) 대시보드 생성 / 갱신
// ───────────────────────────────────────────────
function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('대시보드');
  if (dash) {
    dash.getCharts().forEach(function(c){ dash.removeChart(c); });
    dash.clear();
  } else {
    dash = ss.insertSheet('대시보드', 0); // 맨 앞 탭
  }

  var m1 = readModule(ss, '모듈1');
  var m2 = readModule(ss, '모듈2');

  var row = 1;

  // 제목
  dash.getRange(row, 1).setValue('🛰️ 공간잇기 — 학습 현황 대시보드')
      .setFontSize(16).setFontWeight('bold');
  row += 1;
  dash.getRange(row, 1).setValue('갱신: ' + Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm'))
      .setFontColor('#888888');
  row += 2;

  // 제출 현황
  dash.getRange(row, 1).setValue('📥 제출 현황').setFontWeight('bold').setFontSize(12);
  row += 1;
  dash.getRange(row, 1, 1, 3).setValues([['모듈', '제출 인원', '제출 완료 학생(학번)']]).setFontWeight('bold');
  row += 1;
  dash.getRange(row, 1, 1, 3).setValues([['모듈1 · 삼수선', m1.count, m1.sids.join(', ')]]);
  row += 1;
  dash.getRange(row, 1, 1, 3).setValues([['모듈2 · 정사영', m2.count, m2.sids.join(', ')]]);
  row += 2;

  // 사고 유형 분포 (표)
  dash.getRange(row, 1).setValue('🧠 사고 유형 분포').setFontWeight('bold').setFontSize(12);
  row += 1;
  var typeHeaderRow = row;
  dash.getRange(row, 1, 1, 3).setValues([['사고 유형', '모듈1', '모듈2']]).setFontWeight('bold');
  row += 1;
  THINK_TYPES.forEach(function(t){
    dash.getRange(row, 1, 1, 3).setValues([[
      t + ' ' + TYPE_LABEL[t],
      m1.typeCount[t] || 0,
      m2.typeCount[t] || 0
    ]]);
    row += 1;
  });
  var typeEndRow = row - 1;
  row += 1;

  // 자유 서술 모아보기
  dash.getRange(row, 1).setValue('✏️ 자유 서술 모아보기').setFontWeight('bold').setFontSize(12);
  row += 1;
  dash.getRange(row, 1, 1, 4).setValues([['모듈', '학번', '이름', '서술']]).setFontWeight('bold');
  row += 1;
  var notes = m1.notes.concat(m2.notes);
  if (notes.length === 0) {
    dash.getRange(row, 1).setValue('(아직 자유 서술 응답이 없습니다)').setFontColor('#888888');
    row += 1;
  } else {
    notes.forEach(function(n){
      dash.getRange(row, 1, 1, 4).setValues([[n.mod, n.sid, n.name, n.text]]);
      row += 1;
    });
  }

  // 차트: 사고 유형 분포 (묶은 세로 막대)
  var typeRange = dash.getRange(typeHeaderRow, 1, typeEndRow - typeHeaderRow + 1, 3);
  var chart = dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(typeRange)
    .setPosition(typeHeaderRow, 5, 0, 0)
    .setOption('title', '사고 유형 분포 (모듈별)')
    .setOption('legend', { position: 'top' })
    .setOption('width', 460)
    .setOption('height', 280)
    .build();
  dash.insertChart(chart);

  // 열 너비
  dash.setColumnWidth(1, 160);
  dash.setColumnWidth(2, 90);
  dash.setColumnWidth(3, 220);
  dash.setColumnWidth(4, 360);

  SpreadsheetApp.getUi().alert('✅ 대시보드를 갱신했습니다.');
}

// ───────────────────────────────────────────────
// 보조: 한 모듈 시트를 읽어 집계
// ───────────────────────────────────────────────
function readModule(ss, sheetName) {
  var out = { count: 0, sids: [], typeCount: {}, notes: [] };
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return out;

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return out;

  // 컬럼: 0제출시각 1모듈 2학번 3이름 4사고유형 5자유서술
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[2] && !r[3]) continue;
    out.count += 1;
    if (r[2]) out.sids.push(String(r[2]));

    var method = String(r[4] || '');
    THINK_TYPES.forEach(function(t){
      if (method.indexOf(t) === 0) {
        out.typeCount[t] = (out.typeCount[t] || 0) + 1;
      }
    });

    var note = String(r[5] || '').trim();
    if (note) {
      out.notes.push({
        mod: (sheetName === '모듈2' ? '정사영' : '삼수선'),
        sid: r[2], name: r[3], text: note
      });
    }
  }
  return out;
}
