/**
 * 공간잇기 — 학습 활동 로그 · 성찰로그 수합 + 대시보드 Apps Script  (v2)
 * ════════════════════════════════════════════════════════════════════
 * 받는 데이터 (kosa.js가 보냄)
 *  ① kind=activity : 단계별 활동 기록 → 단계마다 별도 시트
 *       삼수선M1 · 삼수선M2 · 삼수선M3 · 정사영_개념탐구 ·
 *       정사영_개념이해1 · 정사영_개념이해2 · 정사영_개념이해3 · 정사영_미션탐구
 *       열: 제출시각·학번·이름·상태·활동시간(초)·활동시간·오답횟수·오답상세·힌트·이탈횟수·이탈시간(초)·방문·추가
 *  ② kind=reflect  : 성찰로그 → '모듈1' / '모듈2' 시트 (기존과 동일)
 *  ③ kind=login    : 접속 기록 → '접속기록' 시트
 *
 * 교사용 웹 대시보드 (그래프)
 *   웹 앱 URL 뒤에  ?key=교사비밀번호  를 붙여 열면 브라우저에서 그래프로 봅니다.
 *   예) https://script.google.com/macros/s/XXXX/exec?key=0000
 *   (아래 TEACHER_KEY 값을 꼭 바꾸세요)
 *
 * ── 설치 ──
 * 1. 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드 전체를 붙여넣고 저장
 * 2. 배포 → 배포 관리 → (기존 배포) 편집 → 버전: 새 버전 → 배포
 *      ※ '새 배포'를 만들면 URL이 바뀌므로, 기존 배포를 '편집'해야 kosa.js의 URL을 그대로 쓸 수 있습니다.
 *      실행 사용자: 나 / 액세스: 모든 사용자
 * 3. (처음이라면) 웹 앱 URL을 kosa.js 의 KOSA.SHEET_URL 에 입력
 */

var TEACHER_KEY = '0000';   // ◀ 교사용 웹 대시보드 비밀번호 (초기값 0000 — 변경 권장)

var STAGE_SHEETS = ['삼수선M1','삼수선M2','삼수선M3','삼수선_문제해결','삼수선_추가문제',
  '정사영_개념탐구','정사영_개념이해1','정사영_개념이해2','정사영_개념이해3','정사영_미션탐구','정사영_추가문제'];
var STAGE_LABEL = {
  '삼수선M1':'삼수선 M1','삼수선M2':'삼수선 M2','삼수선M3':'삼수선 M3','삼수선_문제해결':'삼수선 문제해결','삼수선_추가문제':'삼수선 추가문제',
  '정사영_개념탐구':'개념탐구','정사영_개념이해1':'개념이해①','정사영_개념이해2':'개념이해②',
  '정사영_개념이해3':'개념이해③','정사영_미션탐구':'미션탐구','정사영_추가문제':'정사영 추가문제'
};
var ACT_HEADER = ['제출시각','학번','이름','상태','활동시간(초)','활동시간','오답횟수','오답상세','힌트','이탈횟수','이탈시간(초)','방문','추가'];

var THINK_TYPES = ['🔢', '🔁', '👁', '📋'];
var TYPE_LABEL = { '🔢':'논리·공식형', '🔁':'탐색·시행착오형', '👁':'직관·관찰형', '📋':'절차·분석형' };

// ───────────────────────────────────────────────
// 수신
// ───────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    var p = e.parameter || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kind = p.kind || 'reflect';

    if (kind === 'activity') {
      var name = STAGE_SHEETS.indexOf(p.sheet) >= 0 ? p.sheet : ('기타_' + (p.stage || 'unknown'));
      var sh = getOrCreate(ss, name, ACT_HEADER);
      var sec = Number(p.activeSec || 0);
      sh.appendRow([new Date(), p.sid||'', p.sname||'', p.status||'', sec, fmtSec(sec),
        Number(p.wrongTotal||0), p.wrongDetail||'', Number(p.hints||0),
        Number(p.leaves||0), Number(p.leaveSec||0), Number(p.visits||0), p.extra||'']);
      return ok({sheet:name});
    }
    if (kind === 'heartbeat') {
      var rt = getOrCreate(ss, '실시간', ['마지막수신','학번','이름','단계','활동시간(초)','오답','힌트','이탈','이탈시간(초)']);
      var vals = rt.getLastRow() > 1 ? rt.getRange(2,2,rt.getLastRow()-1,1).getValues() : [];
      var rowIdx = -1; for (var i=0;i<vals.length;i++){ if(String(vals[i][0])===String(p.sid)){ rowIdx=i+2; break; } }
      var rowVals = [new Date(), p.sid||'', p.sname||'', p.label||p.stage||'', Number(p.activeSec||0), Number(p.wrongTotal||0), Number(p.hints||0), Number(p.leaves||0), Number(p.leaveSec||0)];
      if (rowIdx>0) rt.getRange(rowIdx,1,1,rowVals.length).setValues([rowVals]); else rt.appendRow(rowVals);
      return ok({sheet:'실시간'});
    }
    if (kind === 'login') {
      var lg = getOrCreate(ss, '접속기록', ['접속시각','학번','이름','화면','기기']);
      lg.appendRow([new Date(), p.sid||'', p.sname||'', p.screen||'', p.ua||'']);
      return ok({sheet:'접속기록'});
    }
    // 성찰로그 (kind=reflect 또는 구버전)
    var sheetName = (p.module && p.module.indexOf('M2') === 0) ? '모듈2' : '모듈1';
    var rs = getOrCreate(ss, sheetName, ['제출시각','모듈','학번','이름','사고유형','자유서술']);
    rs.appendRow([new Date(), p.module||'', p.sid||'', p.sname||'', p.method||'', p.extra||'']);
    return ok({sheet:sheetName});
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result:'error', message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally { try{ lock.releaseLock(); }catch(e){} }
}
function ok(o){ o.result='success'; return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function getOrCreate(ss, name, header){
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.getRange(1,1,1,header.length).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
  }
  return sh;
}
function fmtSec(s){ s=Math.round(s); return Math.floor(s/60)+'분 '+(s%60)+'초'; }

// ───────────────────────────────────────────────
// 교사용 웹 대시보드 (GET)
// ───────────────────────────────────────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.key !== TEACHER_KEY) {
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font-family:sans-serif;padding:40px;text-align:center;color:#334">' +
      '<h2>🛰️ 공간잇기 성찰로그 수합 스크립트 — 정상 작동 중</h2>' +
      '<p>교사용 대시보드는 URL 끝에 <code>?key=비밀번호</code> 를 붙여 여세요.</p></div>');
  }
  var data = collectAll();
  var t = HtmlService.createTemplate(DASH_HTML);
  t.json = JSON.stringify(data);
  return t.evaluate().setTitle('공간잇기 — 학습 현황 대시보드')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 모든 시트를 읽어 JSON으로 집계 (웹 대시보드·시트 대시보드 공용)
function collectAll(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {updated: Utilities.formatDate(new Date(),'GMT+9','yyyy-MM-dd HH:mm'),
             stages:[], students:{}, reflect:{m1:readModule(ss,'모듈1'), m2:readModule(ss,'모듈2')}, logins:[]};
  var students = out.students;
  function stu(sid, name){ if(!students[sid]) students[sid]={sid:sid,name:name||'',stages:{}}; if(name&&!students[sid].name) students[sid].name=name; return students[sid]; }

  var lg = ss.getSheetByName('접속기록');
  if (lg && lg.getLastRow()>1){
    lg.getRange(2,1,lg.getLastRow()-1,3).getValues().forEach(function(r){
      if(!r[1]) return; stu(String(r[1]), String(r[2]));
      out.logins.push({at:fmtDate(r[0]), sid:String(r[1]), name:String(r[2])});
    });
  }
  // 실시간 상태 (하트비트) + 주의 학생 판정
  out.live = []; var now = Date.now();
  var rt = ss.getSheetByName('실시간');
  if (rt && rt.getLastRow()>1){
    rt.getRange(2,1,rt.getLastRow()-1,9).getValues().forEach(function(r){
      if(!r[1]) return; var ago = Math.round((now - new Date(r[0]).getTime())/1000);
      var flags = [];
      if (Number(r[5])>=3) flags.push('오답 '+r[5]+'회');
      if (Number(r[4])>=480) flags.push('머묾 '+Math.round(r[4]/60)+'분');
      if (Number(r[7])>=2) flags.push('이탈 '+r[7]+'회');
      out.live.push({sid:String(r[1]), name:String(r[2]), stage:String(r[3]), sec:Number(r[4]), wrong:Number(r[5]), hints:Number(r[6]), leaves:Number(r[7]), leaveSec:Number(r[8]), ago:ago, online:ago<180, flags:flags});
      stu(String(r[1]), String(r[2]));
    });
    out.live.sort(function(a,b){ return (b.flags.length-a.flags.length) || (a.ago-b.ago); });
  }
  STAGE_SHEETS.forEach(function(name){
    var st = {sheet:name, label:STAGE_LABEL[name], rows:[], done:0, abandoned:0,
              avgSec:0, avgWrong:0, avgHints:0, avgLeaves:0, wrongDist:{}, stepWrong:{}};
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow()>1){
      var vals = sh.getRange(2,1,sh.getLastRow()-1,ACT_HEADER.length).getValues();
      // 학생별 최신 행 (완료 우선)
      var best = {};
      vals.forEach(function(r){
        var sid = String(r[1]||''); if(!sid) return;
        var row = {at:fmtDate(r[0]), sid:sid, name:String(r[2]||''), status:String(r[3]||''), sec:Number(r[4]||0),
                   wrong:Number(r[6]||0), wrongDetail:String(r[7]||''), hints:Number(r[8]||0),
                   leaves:Number(r[9]||0), leaveSec:Number(r[10]||0), visits:Number(r[11]||0)};
        var b = best[sid];
        if(!b || (row.status==='완료' && b.status!=='완료') || (row.status===b.status)) best[sid]=row;
      });
      var doneRows=[];
      Object.keys(best).forEach(function(sid){
        var row=best[sid]; st.rows.push(row);
        stu(sid,row.name).stages[name]=row;
        if(row.status==='완료'){ st.done++; doneRows.push(row); } else st.abandoned++;
        var w = Math.min(row.wrong,5); var k = w>=5?'5+':String(w);
        st.wrongDist[k]=(st.wrongDist[k]||0)+1;
        // 문항(단계)별 오답: wrongDetail JSON {"step3":2,...} → 오답 총횟수 · 틀린 학생 수
        try { var wd = JSON.parse(row.wrongDetail||'{}'); Object.keys(wd).forEach(function(q){
          if(!st.stepWrong[q]) st.stepWrong[q]={total:0, students:0, sids:[]};
          st.stepWrong[q].total += Number(wd[q]||0); st.stepWrong[q].students += 1; st.stepWrong[q].sids.push(sid); }); } catch(e) {}
      });
      if(doneRows.length){
        st.avgSec   = avg(doneRows.map(function(r){return r.sec;}));
        st.avgWrong = avg(doneRows.map(function(r){return r.wrong;}));
        st.avgHints = avg(doneRows.map(function(r){return r.hints;}));
        st.avgLeaves= avg(doneRows.map(function(r){return r.leaves;}));
      }
    }
    out.stages.push(st);
  });
  // 누적 기준 '도움이 필요한 학생': 오답 5회↑ / 중단 단계 있음 / 이탈 3회↑ (오답 총합 내림차순)
  out.attention = Object.keys(students).map(function(sid){
    var st = students[sid], wrong=0, leaves=0, abandoned=[], worst=null, last='';
    Object.keys(st.stages).forEach(function(sh){ var r=st.stages[sh]; wrong+=Number(r.wrong||0); leaves+=Number(r.leaves||0);
      if(r.status!=='완료') abandoned.push(STAGE_LABEL[sh]||sh);
      if(!worst || Number(r.wrong||0)>worst.wrong) worst={label:STAGE_LABEL[sh]||sh, wrong:Number(r.wrong||0)};
      if(String(r.at||'')>last) last=String(r.at||''); });
    var flags=[]; if(wrong>=5) flags.push('오답 '+wrong+'회'); if(abandoned.length) flags.push('중단: '+abandoned.join('·')); if(leaves>=3) flags.push('이탈 '+leaves+'회');
    return {sid:sid, name:st.name, wrong:wrong, leaves:leaves, abandoned:abandoned, worst:worst, flags:flags, last:last};
  }).filter(function(a){ return a.flags.length; }).sort(function(a,b){ return (b.wrong-a.wrong)||(b.abandoned.length-a.abandoned.length); });
  return out;
}
function getData(){ return JSON.stringify(collectAll()); }   // 웹 대시보드가 주기적으로 호출
function avg(a){ if(!a.length) return 0; return Math.round(a.reduce(function(x,y){return x+y;},0)/a.length*10)/10; }
function fmtDate(d){ try{ return Utilities.formatDate(new Date(d),'GMT+9','MM-dd HH:mm'); }catch(e){ return String(d); } }

// ───────────────────────────────────────────────
// 시트 안 대시보드 (메뉴)
// ───────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 공간잇기 교사용')
    .addItem('교사용 시트(공간잇기_교사용) 새로고침', 'buildDashboard')
    .addItem('교사용 웹 대시보드 주소 보기', 'showDashUrl')
    .addToUi();
}
function showDashUrl(){
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('교사용 웹 대시보드(그래프):\n\n' + url + '?key=' + TEACHER_KEY +
    '\n\n(배포된 웹 앱 URL 뒤에 ?key=비밀번호 를 붙인 주소입니다)');
}

function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('공간잇기_교사용');
  if (dash) { dash.getCharts().forEach(function(c){ dash.removeChart(c); }); dash.clear(); }
  else dash = ss.insertSheet('공간잇기_교사용', 0);

  var data = collectAll();
  var m1 = data.reflect.m1, m2 = data.reflect.m2;
  var row = 1;
  dash.getRange(row,1).setValue('🛰️ 공간잇기_교사용 — 학습 현황').setFontSize(16).setFontWeight('bold'); row++;
  dash.getRange(row,1).setValue('갱신: ' + data.updated).setFontColor('#888888'); row+=2;

  // 🆘 도움이 필요한 학생 (오답 5회↑ · 중단 · 이탈 3회↑)
  dash.getRange(row,1).setValue('🆘 도움이 필요한 학생 — 오답 총합 순 (오답 5회↑ · 중단 단계 있음 · 이탈 3회↑)').setFontWeight('bold').setFontSize(12).setFontColor('#c0392b'); row++;
  dash.getRange(row,1,1,7).setValues([['학번','이름','총 오답','가장 막힌 단계','중단한 단계','이탈','마지막 활동']]).setFontWeight('bold').setBackground('#fdecea'); row++;
  if (!data.attention.length) { dash.getRange(row,1).setValue('(해당 학생 없음)').setFontColor('#888888'); row++; }
  data.attention.forEach(function(a){
    dash.getRange(row,1,1,7).setValues([[a.sid, a.name, a.wrong, a.worst?(a.worst.label+' ('+a.worst.wrong+'회)'):'', a.abandoned.join(', ')||'-', a.leaves, a.last]]);
    if (a.wrong>=8 || a.abandoned.length>=2) dash.getRange(row,1,1,7).setBackground('#ffd6d6');
    row++;
  });
  row++;

  // 단계별 활동 요약
  dash.getRange(row,1).setValue('📈 단계별 활동 요약 (완료 학생 기준 평균)').setFontWeight('bold').setFontSize(12); row++;
  var actHead = row;
  dash.getRange(row,1,1,7).setValues([['단계','완료','중단','평균 활동시간(초)','평균 오답','평균 힌트','평균 이탈']]).setFontWeight('bold'); row++;
  data.stages.forEach(function(s){
    dash.getRange(row,1,1,7).setValues([[s.label, s.done, s.abandoned, s.avgSec, s.avgWrong, s.avgHints, s.avgLeaves]]); row++;
  });
  var actEnd = row-1; row++;

  // 학생별 진행표
  dash.getRange(row,1).setValue('🧑‍🎓 학생별 진행 (완료=✅ 시간/오답, 중단=⏸, 미참여=·)').setFontWeight('bold').setFontSize(12); row++;
  var head = ['학번','이름'].concat(data.stages.map(function(s){return s.label;}));
  dash.getRange(row,1,1,head.length).setValues([head]).setFontWeight('bold'); row++;
  var sids = Object.keys(data.students).sort();
  sids.forEach(function(sid){
    var st = data.students[sid];
    var line = [sid, st.name];
    data.stages.forEach(function(s){
      var r = st.stages[s.sheet];
      line.push(!r ? '·' : (r.status==='완료' ? '✅ '+fmtSec(r.sec)+' / 오답'+r.wrong+(r.leaves?' / 이탈'+r.leaves:'') : '⏸ '+fmtSec(r.sec)));
    });
    dash.getRange(row,1,1,line.length).setValues([line]); row++;
  });
  row++;

  // 성찰로그 제출 현황
  dash.getRange(row,1).setValue('📥 성찰로그 제출 현황').setFontWeight('bold').setFontSize(12); row++;
  dash.getRange(row,1,1,3).setValues([['모듈','제출 인원','제출 완료 학생(학번)']]).setFontWeight('bold'); row++;
  dash.getRange(row,1,1,3).setValues([['모듈1 · 삼수선', m1.count, m1.sids.join(', ')]]); row++;
  dash.getRange(row,1,1,3).setValues([['모듈2 · 정사영', m2.count, m2.sids.join(', ')]]); row+=2;

  // 사고 유형 분포
  dash.getRange(row,1).setValue('🧠 사고 유형 분포').setFontWeight('bold').setFontSize(12); row++;
  var typeHeaderRow = row;
  dash.getRange(row,1,1,3).setValues([['사고 유형','모듈1','모듈2']]).setFontWeight('bold'); row++;
  THINK_TYPES.forEach(function(t){
    dash.getRange(row,1,1,3).setValues([[t+' '+TYPE_LABEL[t], m1.typeCount[t]||0, m2.typeCount[t]||0]]); row++;
  });
  var typeEndRow = row-1; row++;

  // 자유 서술
  dash.getRange(row,1).setValue('✏️ 자유 서술 모아보기').setFontWeight('bold').setFontSize(12); row++;
  dash.getRange(row,1,1,4).setValues([['모듈','학번','이름','서술']]).setFontWeight('bold'); row++;
  var notes = m1.notes.concat(m2.notes);
  if(!notes.length){ dash.getRange(row,1).setValue('(아직 자유 서술 응답이 없습니다)').setFontColor('#888888'); row++; }
  else notes.forEach(function(n){ dash.getRange(row,1,1,4).setValues([[n.mod,n.sid,n.name,n.text]]); row++; });

  // 차트 1: 단계별 완료/중단
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(dash.getRange(actHead,1,actEnd-actHead+1,3)).setPosition(actHead,9,0,0)
    .setOption('title','단계별 완료·중단 인원').setOption('legend',{position:'top'})
    .setOption('width',480).setOption('height',260).build());
  // 차트 2: 단계별 평균 활동시간
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(dash.getRange(actHead,1,actEnd-actHead+1,1)).addRange(dash.getRange(actHead,4,actEnd-actHead+1,1))
    .setPosition(actHead+14,9,0,0)
    .setOption('title','단계별 평균 활동시간(초)').setOption('legend',{position:'none'})
    .setOption('width',480).setOption('height',260).build());
  // 차트 3: 사고 유형
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(dash.getRange(typeHeaderRow,1,typeEndRow-typeHeaderRow+1,3)).setPosition(typeHeaderRow,9,0,0)
    .setOption('title','사고 유형 분포 (모듈별)').setOption('legend',{position:'top'})
    .setOption('width',480).setOption('height',280).build());

  dash.setColumnWidth(1,150); dash.setColumnWidth(2,90); dash.setColumnWidth(3,140);
  for (var c=4;c<=10;c++) dash.setColumnWidth(c,150);
  SpreadsheetApp.getUi().alert('✅ 공간잇기_교사용 시트를 갱신했습니다.');
}

function readModule(ss, sheetName) {
  var out = { count: 0, sids: [], typeCount: {}, notes: [] };
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return out;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return out;
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[2] && !r[3]) continue;
    out.count += 1;
    if (r[2]) out.sids.push(String(r[2]));
    var method = String(r[4] || '');
    THINK_TYPES.forEach(function(t){ if (method.indexOf(t) === 0) out.typeCount[t] = (out.typeCount[t]||0)+1; });
    var note = String(r[5] || '').trim();
    if (note) out.notes.push({ mod:(sheetName==='모듈2'?'정사영':'삼수선'), sid:r[2], name:r[3], text:note });
  }
  return out;
}

// ───────────────────────────────────────────────
// 교사용 웹 대시보드 HTML (Google Charts)
// ───────────────────────────────────────────────
var DASH_HTML = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
'<script src="https://www.gstatic.com/charts/loader.js"></script>' +
'<style>' +
'body{font-family:"Noto Sans KR",sans-serif;background:#0b1424;color:#dbe7f7;margin:0;padding:16px 20px}' +
'h1{font-size:20px;margin:0}.sub{color:#7f9cc4;font-size:12px}' +
'.top{display:flex;align-items:center;gap:14px;margin-bottom:14px;flex-wrap:wrap}.top .sp{flex:1}' +
'.kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}' +
'.k{background:#111e33;border:1px solid #223a60;border-radius:10px;padding:12px 14px}.k .l{font-size:11px;color:#7f9cc4;letter-spacing:.1em}.k .v{font-size:26px;font-weight:900;color:#fff}.k.warn .v{color:#ff8a65}.k.ok .v{color:#5ff0a0}' +
'.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:14px}' +
'.card{background:#111e33;border:1px solid #223a60;border-radius:10px;padding:14px}.card h2{font-size:14px;margin:0 0 8px;color:#9ec5ff}' +
'.chart{width:100%;height:260px}' +
'table{border-collapse:collapse;width:100%;font-size:12px}th,td{border-bottom:1px solid #223a60;padding:5px 6px;text-align:left;white-space:nowrap}th{color:#9ec5ff;position:sticky;top:0;background:#111e33}' +
'.wrap{max-height:460px;overflow:auto}.no{color:#4a5a75}.ok{color:#5ff0a0}.ab{color:#ffb070}.warn{color:#ff8a65}' +
'.att{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px}' +
'.st{background:#1a0e10;border:1px solid #6a2a2a;border-left:4px solid #ff5a3a;border-radius:8px;padding:9px 11px;font-size:12.5px}.st b{font-size:14px;color:#fff}.st .tg{display:inline-block;margin:4px 4px 0 0;padding:2px 7px;border-radius:999px;background:#3a1010;color:#ffb0a0;font-size:11px;font-weight:700}.st .m{color:#9ec5ff;font-size:11px;margin-top:3px}' +
'.st.off{opacity:.55;border-left-color:#4a5a75}' +
'.cell{display:inline-block;min-width:34px;text-align:center;padding:3px 6px;border-radius:5px;font-weight:700}.c-done{background:#0f3a24;color:#8affc0}.c-ab{background:#3a2410;color:#ffcf9a}.c-live{background:#102a4a;color:#8ec3ff;box-shadow:0 0 0 1px #5599ff}.c-no{background:#141c2c;color:#4a5a75}' +
'.note{font-size:12.5px;line-height:1.6;padding:6px 8px;border-left:3px solid #3a5a90;margin:6px 0;background:#0e1a2d;white-space:normal}' +
'.legend{font-size:11px;color:#7f9cc4;margin-top:6px}' +
'</style></head><body>' +
'<div class="top"><div><h1>🛰️ 공간잇기_교사용 — 실시간 학습 현황</h1><div class="sub" id="upd"></div></div><div class="sp"></div>' +
'<label style="font-size:12px;color:#9ec5ff"><input type="checkbox" id="auto" checked> 45초마다 자동 갱신</label><button onclick="refresh()" style="background:#1c2f4d;border:1px solid #5599ff;color:#cfe6ff;border-radius:6px;padding:6px 12px;cursor:pointer">🔄 지금 갱신</button></div>' +
'<div class="kpi"><div class="k"><div class="l">접속 학생</div><div class="v" id="kAll">-</div></div><div class="k ok"><div class="l">지금 활동 중 (3분 내 수신)</div><div class="v" id="kOn">-</div></div><div class="k warn"><div class="l">⚠ 주의 학생</div><div class="v" id="kWarn">-</div></div><div class="k"><div class="l">모듈2 미션탐구 완료</div><div class="v" id="kDone">-</div></div></div>' +
'<div class="card" style="margin-bottom:14px"><h2>⚠ 지금 살펴볼 학생 <span class="legend">— 오답 3회↑ · 한 단계 8분↑ · 창 이탈 2회↑ (최근 수신 순)</span></h2><div class="att" id="att"></div></div>' +
'<div class="card" style="margin-bottom:14px"><h2>🆘 도움이 필요한 학생 <span class="legend">— 누적 기준: 오답 5회↑ · 중단 단계 있음 · 이탈 3회↑ (오답 순)</span></h2><div class="wrap" style="max-height:260px"><table id="attTb"></table></div></div>' +
'<div class="card" style="margin-bottom:14px"><h2>🧑‍🎓 진행 보드 <span class="legend">— 완료 ✅(활동시간′/오답) · 중단 ⏸ · 진행 중 🔵 · 미참여 ·</span></h2><div class="wrap"><table id="tb"></table></div></div>' +
'<div class="grid">' +
'<div class="card"><h2>단계별 완료 · 중단 인원</h2><div id="c1" class="chart"></div></div>' +
'<div class="card"><h2>단계별 평균 활동시간(초) · 평균 오답</h2><div id="c2" class="chart"></div></div>' +
'<div class="card"><h2>단계별 오답 횟수 분포 (학생 수)</h2><div id="c3" class="chart"></div></div>' +
'<div class="card"><h2>🧩 문제해결 — 단계별 막힌 지점</h2><div id="c6" class="chart"></div></div>' +
'<div class="card"><h2>성찰로그 사고 유형 분포</h2><div id="c5" class="chart"></div></div>' +
'<div class="card"><h2>성찰로그 자유 서술</h2><div id="notes" style="max-height:260px;overflow:auto"></div></div>' +
'</div>' +
'<script>' +
'var D=<?!= json ?>;' +
'google.charts.load("current",{packages:["corechart","bar"]});google.charts.setOnLoadCallback(function(){draw();});' +
'var OPT={backgroundColor:"transparent",legend:{position:"top",textStyle:{color:"#c9d8ee"}},hAxis:{textStyle:{color:"#c9d8ee"}},vAxis:{textStyle:{color:"#c9d8ee"},gridlines:{color:"#223a60"},minValue:0},chartArea:{width:"85%",height:"65%"}};' +
'function o(x){var r=JSON.parse(JSON.stringify(OPT));for(var k in x)r[k]=x[k];return r;}' +
'function fmt(s){return Math.floor(s/60)+"′"+(s%60)+"″";}' +
'function draw(){' +
' document.getElementById("upd").textContent="갱신: "+D.updated;' +
' var live=D.live||[], on=live.filter(function(l){return l.online;}), warn=live.filter(function(l){return l.flags.length;});' +
' document.getElementById("kAll").textContent=Object.keys(D.students).length+"명"; document.getElementById("kOn").textContent=on.length+"명"; document.getElementById("kWarn").textContent=warn.length+"명";' +
' var fin=D.stages.filter(function(s){return s.sheet==="정사영_미션탐구";})[0]; document.getElementById("kDone").textContent=(fin?fin.done:0)+"명";' +
' document.getElementById("att").innerHTML = warn.length ? warn.map(function(l){return "<div class=st"+(l.online?"":" off")+"><b>"+l.sid+" "+l.name+"</b> <span class=m>"+l.stage+" · "+fmt(l.sec)+(l.online?" · 활동 중":" · "+Math.round(l.ago/60)+"분 전 수신")+"</span><br>"+l.flags.map(function(f){return "<span class=tg>"+f+"</span>";}).join("")+"</div>";}).join("") : "<div class=no>지금은 주의가 필요한 학생이 없습니다.</div>";' +
' var at=D.attention||[]; document.getElementById("attTb").innerHTML = at.length ? "<tr><th>학번</th><th>이름</th><th>총 오답</th><th>가장 막힌 단계</th><th>중단한 단계</th><th>이탈</th><th>마지막 활동</th></tr>"+at.map(function(a){return "<tr"+((a.wrong>=8||a.abandoned.length>=2)?" style=background:#2a1010":"")+"><td><b>"+a.sid+"</b></td><td>"+a.name+"</td><td class=warn><b>"+a.wrong+"</b></td><td>"+(a.worst?a.worst.label+" ("+a.worst.wrong+"회)":"")+"</td><td class=ab>"+(a.abandoned.join(", ")||"-")+"</td><td>"+a.leaves+"</td><td class=no>"+a.last+"</td></tr>";}).join("") : "<tr><td class=no>해당 학생 없음</td></tr>";' +
' var liveMap={}; live.forEach(function(l){ liveMap[l.sid]=l; });' +
' var h="<tr><th>학번</th><th>이름</th><th>현재</th>"+D.stages.map(function(s){return "<th>"+s.label+"</th>";}).join("")+"</tr>";' +
' Object.keys(D.students).sort().forEach(function(sid){var st=D.students[sid], lv=liveMap[sid];' +
'  h+="<tr><td>"+sid+"</td><td>"+st.name+"</td><td>"+(lv?(lv.online?"<span class=ok>● </span>":"<span class=no>○ </span>")+lv.stage+" "+fmt(lv.sec):"<span class=no>-</span>")+"</td>"+D.stages.map(function(s){var r=st.stages[s.sheet];' +
'   if(lv&&lv.online&&lv.stage===s.label&&!(r&&r.status==="완료")) return "<td><span class=\\"cell c-live\\">🔵 "+fmt(lv.sec)+"</span></td>";' +
'   if(!r) return "<td><span class=\\"cell c-no\\">·</span></td>";' +
'   if(r.status==="완료") return "<td><span class=\\"cell c-done\\">✅ "+fmt(r.sec)+" / "+r.wrong+"</span></td>";' +
'   return "<td><span class=\\"cell c-ab\\">⏸ "+fmt(r.sec)+" / "+r.wrong+"</span></td>";}).join("")+"</tr>";});' +
' document.getElementById("tb").innerHTML=h;' +
' var d1=[["단계","완료","중단"]];D.stages.forEach(function(s){d1.push([s.label,s.done,s.abandoned]);});new google.visualization.ColumnChart(document.getElementById("c1")).draw(google.visualization.arrayToDataTable(d1),o({colors:["#4fd18b","#ffb070"],isStacked:true}));' +
' var d2=[["단계","평균 활동시간(초)","평균 오답"]];D.stages.forEach(function(s){d2.push([s.label,s.avgSec,s.avgWrong]);});new google.visualization.ColumnChart(document.getElementById("c2")).draw(google.visualization.arrayToDataTable(d2),o({colors:["#5aa9ff","#ff6b8a"],series:{1:{targetAxisIndex:1}},vAxes:{1:{textStyle:{color:"#c9d8ee"}}}}));' +
' var keys=["0","1","2","3","4","5+"];var d3=[["단계"].concat(keys.map(function(k){return "오답 "+k+"회";}))];D.stages.forEach(function(s){d3.push([s.label].concat(keys.map(function(k){return s.wrongDist[k]||0;})));});new google.visualization.ColumnChart(document.getElementById("c3")).draw(google.visualization.arrayToDataTable(d3),o({isStacked:true,colors:["#4fd18b","#9ad5ff","#ffd166","#ffb070","#ff8a65","#ff5c7a"]}));' +
' var PB=D.stages.filter(function(s){return s.sheet==="삼수선_문제해결";})[0];var PBN={step1:"1 C의 위치",step2:"2 BD′",step3:"3① ∠AD′B",step4:"3② AD′",step5:"4① D′H⊥AB",step6:"4② 삼수선",step7:"5① D′H",step8:"5② DH",step9:"6 정답"};' +
' if(PB){var d6=[["단계","오답 총횟수","틀린 학생 수"]];Object.keys(PBN).forEach(function(k){var v=PB.stepWrong[k]||{total:0,students:0};d6.push([PBN[k],v.total,v.students]);});new google.visualization.ColumnChart(document.getElementById("c6")).draw(google.visualization.arrayToDataTable(d6),o({colors:["#ff6b8a","#ffd166"]}));} else document.getElementById("c6").textContent="(아직 기록 없음)";' +
' var T={"🔢":"논리·공식형","🔁":"탐색·시행착오형","👁":"직관·관찰형","📋":"절차·분석형"};var d5=[["유형","모듈1","모듈2"]];Object.keys(T).forEach(function(k){d5.push([k+" "+T[k],D.reflect.m1.typeCount[k]||0,D.reflect.m2.typeCount[k]||0]);});new google.visualization.ColumnChart(document.getElementById("c5")).draw(google.visualization.arrayToDataTable(d5),o({colors:["#ff9355","#ffd34a"]}));' +
' var ns=D.reflect.m1.notes.concat(D.reflect.m2.notes);document.getElementById("notes").innerHTML=ns.length?ns.map(function(n){return "<div class=note><b>["+n.mod+"] "+n.sid+" "+n.name+"</b> — "+n.text+"</div>";}).join(""):"<div class=no>(아직 응답 없음)</div>";' +
'}' +
'function refresh(){ google.script.run.withSuccessHandler(function(j){ D=JSON.parse(j); draw(); }).getData(); }' +
'setInterval(function(){ if(document.getElementById("auto").checked) refresh(); },45000);' +
'</script></body></html>';
