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

// ── 공간잇기_교사용 시트 '핵심 요약' 판정 기준 (숫자만 바꾸면 기준이 바뀝니다) ──────
var FLAG_INACTIVE_DAYS = 3;    // 🚨 이탈자 (a): 마지막 활동으로부터 이 값(일) 이상 미접속
var FLAG_STALL_DAYS    = 2;    // 🚨 이탈자 (b): 중단 단계가 있고, 이 값(일) 이상 진전 없이 정체
var FLAG_WRONG_TOTAL   = 5;    // 🐢 학습부진자(오답형): 누적 오답이 이 값 이상
var FLAG_SLOW_RATIO    = 1.8;  // 🐢 학습부진자(정체형): 그 단계 평균 소요시간의 이 배수 이상
var FLAG_SLOW_MIN_SEC  = 900;  // 🐢 학습부진자(정체형): 그리고 절대 시간도 이 값(초) 이상 (기본 15분)
var EXCEL_MIN_DONE_RATIO = 0.6;// 🌟 학습우수자: 전체 단계 중 이 비율 이상 완료해야 후보로 간주
var EXCEL_MAX_AVG_WRONG  = 1;  // 🌟 학습우수자: 완료 단계당 평균 오답이 이 값 이하
var EXCEL_FAST_RATIO     = 0.85;// 🌟 학습우수자: 평균 소요시간이 '전체 학생 평균'의 이 비율 이하(더 빠름)

var COVER_NAME = '🔒 표지';       // 표지(대문) 시트 이름 — 항상 맨 왼쪽에 위치
var LOCK_LEVEL_KEY = 'kosa_lock_level'; // 잠금 단계 저장 키 (0=표지만 · 1=교사용보드까지 · 2=원본 학생데이터까지)

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
    if (getLockLevel() < 2) { try{ sh.hideSheet(); }catch(e){} }   // 원본 학생데이터 단계가 잠겨있으면 새로 생기는 시트도 바로 숨김
  }
  return sh;
}
function fmtSec(s){ s=Math.round(s); return Math.floor(s/60)+'분 '+(s%60)+'초'; }
// 표 전체(헤더+본문)에 격자 테두리를 둘러줌 — 표끼리 시각적으로 확실히 구분되도록
function gridBorder(sheet, r1, c1, r2, c2){
  if (r2<r1 || c2<c1) return;
  sheet.getRange(r1,c1,r2-r1+1,c2-c1+1).setBorder(true,true,true,true,true,true,'#c7d2e0',SpreadsheetApp.BorderStyle.SOLID);
}

// ───────────────────────────────────────────────
// 교사용 웹 대시보드 (GET)
// ───────────────────────────────────────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  var selfUrl = ScriptApp.getService().getUrl();
  if (p.key !== TEACHER_KEY) {
    var wrong = (typeof p.key !== 'undefined' && p.key !== '');
    var html = LOGIN_HTML.split('{{ACTION}}').join(selfUrl)
      .split('{{ERR}}').join(wrong ? '<div class="err">❌ 비밀번호가 올바르지 않습니다. 다시 입력해주세요.</div>' : '');
    return HtmlService.createHtmlOutput(html)
      .setTitle('공간잇기 — 교사용 로그인')
      .addMetaTag('viewport','width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  function stu(sid, name, ts){
    if(!students[sid]) students[sid]={sid:sid,name:name||'',stages:{},lastMs:0};
    if(name&&!students[sid].name) students[sid].name=name;
    if(ts && ts>students[sid].lastMs) students[sid].lastMs=ts;
    return students[sid];
  }

  var lg = ss.getSheetByName('접속기록');
  if (lg && lg.getLastRow()>1){
    lg.getRange(2,1,lg.getLastRow()-1,3).getValues().forEach(function(r){
      if(!r[1]) return; var ts=new Date(r[0]).getTime(); stu(String(r[1]), String(r[2]), ts);
      out.logins.push({at:fmtDate(r[0]), sid:String(r[1]), name:String(r[2])});
    });
  }
  // 실시간 상태 (하트비트) + 주의 학생 판정
  out.live = []; var now = Date.now();
  var rt = ss.getSheetByName('실시간');
  if (rt && rt.getLastRow()>1){
    rt.getRange(2,1,rt.getLastRow()-1,9).getValues().forEach(function(r){
      if(!r[1]) return; var ts=new Date(r[0]).getTime(); var ago = Math.round((now - ts)/1000);
      var flags = [];
      if (Number(r[5])>=3) flags.push('오답 '+r[5]+'회');
      if (Number(r[4])>=480) flags.push('머묾 '+Math.round(r[4]/60)+'분');
      if (Number(r[7])>=2) flags.push('이탈 '+r[7]+'회');
      out.live.push({sid:String(r[1]), name:String(r[2]), stage:String(r[3]), sec:Number(r[4]), wrong:Number(r[5]), hints:Number(r[6]), leaves:Number(r[7]), leaveSec:Number(r[8]), ago:ago, online:ago<180, flags:flags});
      stu(String(r[1]), String(r[2]), ts);
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
        var row = {at:fmtDate(r[0]), atMs:new Date(r[0]).getTime(), sid:sid, name:String(r[2]||''), status:String(r[3]||''), sec:Number(r[4]||0),
                   wrong:Number(r[6]||0), wrongDetail:String(r[7]||''), hints:Number(r[8]||0),
                   leaves:Number(r[9]||0), leaveSec:Number(r[10]||0), visits:Number(r[11]||0)};
        var b = best[sid];
        if(!b || (row.status==='완료' && b.status!=='완료') || (row.status===b.status)) best[sid]=row;
      });
      var doneRows=[];
      Object.keys(best).forEach(function(sid){
        var row=best[sid]; st.rows.push(row);
        stu(sid,row.name,row.atMs).stages[name]=row;
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

  // ── 핵심 요약 4종: 🚨이탈자 · 🐢학습부진자(오답형/정체형) · 🌟학습우수자 ──────────────
  var totalStages = STAGE_SHEETS.length;
  Object.keys(students).forEach(function(sid){
    var st = students[sid], doneCnt=0, abandonedCnt=0, wrongSum=0, secSum=0;
    Object.keys(st.stages).forEach(function(k){
      var r = st.stages[k];
      if(r.status==='완료'){ doneCnt++; wrongSum+=Number(r.wrong||0); secSum+=Number(r.sec||0); }
      else { abandonedCnt++; wrongSum+=Number(r.wrong||0); }
    });
    st.doneCnt=doneCnt; st.abandonedCnt=abandonedCnt; st.wrongSum=wrongSum;
    st.avgWrongDone = doneCnt? Math.round(wrongSum/doneCnt*10)/10 : 0;
    st.avgSecDone   = doneCnt? Math.round(secSum/doneCnt) : 0;
  });
  var secList = Object.keys(students).map(function(sid){return students[sid].avgSecDone;}).filter(function(v){return v>0;});
  var globalAvgSec = secList.length ? secList.reduce(function(a,b){return a+b;},0)/secList.length : 0;

  out.dropouts=[]; out.strugglingWrong=[]; out.strugglingStuck=[]; out.excellent=[];
  Object.keys(students).forEach(function(sid){
    var st = students[sid];
    var daysSince = st.lastMs ? (now-st.lastMs)/86400000 : 999;

    // 🚨 이탈자: (a) 장기 미접속  또는  (b) 중단 상태로 정체
    var reasons=[];
    if(daysSince>=FLAG_INACTIVE_DAYS) reasons.push('미접속 '+Math.floor(daysSince)+'일');
    if(st.abandonedCnt>0 && daysSince>=FLAG_STALL_DAYS) reasons.push('중단 '+st.abandonedCnt+'단계 · 정체 '+Math.floor(daysSince)+'일');
    if(reasons.length) out.dropouts.push({sid:sid, name:st.name, reasons:reasons,
      lastSeen: st.lastMs? fmtDate(st.lastMs):'(기록 없음)', doneCnt:st.doneCnt, abandonedCnt:st.abandonedCnt});

    // 🐢 학습부진자 — 오답형
    if(st.wrongSum>=FLAG_WRONG_TOTAL) out.strugglingWrong.push({sid:sid, name:st.name, wrongSum:st.wrongSum, doneCnt:st.doneCnt});

    // 🐢 학습부진자 — 정체형 (한 단계에 유독 오래 머묾)
    var slow=null;
    Object.keys(st.stages).forEach(function(k){
      var r=st.stages[k]; var sInfo=null;
      for(var i=0;i<out.stages.length;i++){ if(out.stages[i].sheet===k){ sInfo=out.stages[i]; break; } }
      if(!sInfo || !sInfo.avgSec) return;
      if(r.sec>=sInfo.avgSec*FLAG_SLOW_RATIO && r.sec>=FLAG_SLOW_MIN_SEC){
        if(!slow || r.sec>slow.sec) slow={label:STAGE_LABEL[k]||k, sec:r.sec, status:r.status};
      }
    });
    if(slow) out.strugglingStuck.push({sid:sid, name:st.name, stage:slow.label, sec:slow.sec, status:slow.status});

    // 🌟 학습우수자 — 완료율 + 낮은 오답 + 평균보다 빠른 속도, 종합 반영
    var doneRatio = totalStages? st.doneCnt/totalStages : 0;
    if(doneRatio>=EXCEL_MIN_DONE_RATIO && st.avgWrongDone<=EXCEL_MAX_AVG_WRONG &&
       globalAvgSec>0 && st.avgSecDone>0 && st.avgSecDone<=globalAvgSec*EXCEL_FAST_RATIO){
      out.excellent.push({sid:sid, name:st.name, doneCnt:st.doneCnt, avgWrong:st.avgWrongDone, avgSec:st.avgSecDone});
    }
  });
  out.dropouts.sort(function(a,b){ return b.abandonedCnt-a.abandonedCnt || a.sid.localeCompare(b.sid); });
  out.strugglingWrong.sort(function(a,b){ return b.wrongSum-a.wrongSum; });
  out.strugglingStuck.sort(function(a,b){ return b.sec-a.sec; });
  out.excellent.sort(function(a,b){ return b.doneCnt-a.doneCnt || a.avgSec-b.avgSec; });

  return out;
}
function getData(){ return JSON.stringify(collectAll()); }   // 웹 대시보드가 주기적으로 호출
function avg(a){ if(!a.length) return 0; return Math.round(a.reduce(function(x,y){return x+y;},0)/a.length*10)/10; }
function fmtDate(d){ try{ return Utilities.formatDate(new Date(d),'GMT+9','MM-dd HH:mm'); }catch(e){ return String(d); } }

// ───────────────────────────────────────────────
// 시트 안 대시보드 (메뉴)
// ───────────────────────────────────────────────
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().createMenu('📊 공간잇기 교사용')
    .addItem('🔓 교사용 보드 열기', 'unlockBoard')
    .addItem('📂 학생 데이터(원본) 추가로 열기', 'unlockStudentData')
    .addItem('🔒 잠그기', 'lockTeacherSheets')
    .addSeparator()
    .addItem('교사용 시트(공간잇기_교사용) 새로고침', 'buildDashboard')
    .addItem('교사용 웹 대시보드 주소 보기', 'showDashUrl')
    .addToUi();
  try{
    var cov = buildCoverSheet();   // 열 때마다 항상 최신 디자인으로 다시 그림 (이미 있어도 다시 그림)
    ss.setActiveSheet(cov);
  }catch(e){}
}
function showDashUrl(){
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('교사용 웹 대시보드(그래프):\n\n' + url + '?key=' + TEACHER_KEY +
    '\n\n(배포된 웹 앱 URL 뒤에 ?key=비밀번호 를 붙인 주소입니다)');
}

// ───────────────────────────────────────────────
// 표지(대문) + 잠금
// ───────────────────────────────────────────────
function getLockLevel(){ var v = PropertiesService.getDocumentProperties().getProperty(LOCK_LEVEL_KEY); return v ? Number(v) : 0; }
function setLockLevel(n){ PropertiesService.getDocumentProperties().setProperty(LOCK_LEVEL_KEY, String(n)); }
function rawDataSheetNames(){ return STAGE_SHEETS.concat(['모듈1','모듈2','접속기록','실시간']); }

function buildCoverSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cov = ss.getSheetByName(COVER_NAME);
  if (!cov) cov = ss.insertSheet(COVER_NAME, 0);
  else { cov.clear(); cov.clearFormats(); cov.showSheet(); if (ss.getSheets()[0].getSheetId()!==cov.getSheetId()){ ss.setActiveSheet(cov); ss.moveActiveSheet(1); } }
  cov.setHiddenGridlines(true);
  cov.setColumnWidth(1,220); cov.setColumnWidth(2,260); cov.setColumnWidth(3,560);

  // 배너
  cov.setRowHeight(1,48);
  cov.getRange(1,1,1,3).merge().setValue('🛰️  공간잇기 — 교사용 관제실')
    .setBackground('#16233f').setFontColor('#ffffff').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  cov.setRowHeight(2,28);
  cov.getRange(2,1,1,3).merge().setValue('학생들의 학습 활동을 확인하는 교사 전용 공간입니다')
    .setBackground('#26375c').setFontColor('#cfe0ff').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  var row = 4;

  // 카드: 여는 방법
  var openHeadRow = row;
  cov.getRange(row,1,1,3).merge().setValue('🔓  여는 방법').setFontWeight('bold').setFontSize(13).setFontColor('#0b3d91'); row++;
  cov.getRange(row,1,1,3).merge().setValue('① 위쪽 메뉴 [📊 공간잇기 교사용] → [🔓 학생 데이터 열기] 클릭   ② 비밀번호 입력 → 최신 데이터까지 자동으로 갱신되어 열립니다.').setWrap(true); row++;
  cov.getRange(row,1,1,3).merge().setValue('다 보신 뒤에는 [📊 공간잇기 교사용] → [🔒 잠그기]를 눌러 다시 가려두세요 (수업 중 화면 공유 시 학생 정보 노출 방지).').setWrap(true); row++;
  cov.getRange(openHeadRow,1,row-openHeadRow,3).setBackground('#eef4ff').setBorder(true,true,true,true,false,false,'#b9cdf0',SpreadsheetApp.BorderStyle.SOLID);
  row++;

  // 카드: 보안 주의
  var warnHeadRow = row;
  cov.getRange(row,1,1,3).merge().setValue('※ 이 잠금은 "실수로 보이는 것"을 막는 용도예요. 진짜 보안은 이 스프레드시트의 [공유] 권한입니다 — 학생·외부인에게 편집/조회 권한이 없는지 꼭 확인하세요.')
    .setFontColor('#8a4b12').setFontStyle('italic').setWrap(true); row++;
  cov.getRange(warnHeadRow,1,row-warnHeadRow,3).setBackground('#fff6e8').setBorder(true,true,true,true,false,false,'#f0c98a',SpreadsheetApp.BorderStyle.SOLID);
  row+=2;

  // 목차
  cov.getRange(row,1,1,3).merge().setValue('📋  시트 목차').setFontWeight('bold').setFontSize(14).setFontColor('#16233f'); row++;
  var tocHeadRow = row;
  cov.getRange(row,1,1,3).setValues([['구분','시트','설명']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#4a6da8'); row++;
  var toc = [
    ['⭐ 가장 먼저 볼 시트','공간잇기_교사용','이탈자·학습부진자·학습우수자 요약 + 단계별 통계. 평소엔 이 시트 하나만 봐도 충분합니다.'],
    ['📥 원본 데이터 (직접 볼 필요 없음)','삼수선M1 · M2 · M3','삼수선의 정리 미션 1~3 — 단계별 완료/중단·오답·힌트·이탈 기록'],
    ['📥 원본 데이터','삼수선_문제해결 · 삼수선_추가문제','모듈1 문제풀이 활동 기록'],
    ['📥 원본 데이터','정사영_개념탐구 · 개념이해1~3 · 미션탐구','모듈2 단계별 활동 기록'],
    ['📥 원본 데이터','정사영_추가문제','모듈2 문제풀이 활동 기록'],
    ['📥 원본 데이터','모듈1 · 모듈2','성찰로그(자기보고) 제출 내역'],
    ['📥 원본 데이터','접속기록','학생 로그인 시각 기록'],
    ['⏱ 실시간 전용','실시간','학생당 최신 상태 1줄 — 웹 대시보드가 실시간 현황을 보여줄 때만 사용']
  ];
  toc.forEach(function(t,i){
    var r = cov.getRange(row,1,1,3); r.setValues([t]).setWrap(true);
    if (i%2===1) r.setBackground('#f4f7fc');
    row++;
  });
  var tocEndRow = row-1;
  cov.getRange(tocHeadRow,1,tocEndRow-tocHeadRow+1,3).setBorder(true,true,true,true,true,true,'#c7d2e0',SpreadsheetApp.BorderStyle.SOLID);
  row+=2;

  // 흐름
  var flowHeadRow = row;
  cov.getRange(row,1,1,3).merge().setValue('🔄  데이터가 만들어지는 흐름').setFontWeight('bold').setFontSize(13).setFontColor('#1e6b3a'); row++;
  cov.getRange(row,1,1,3).merge().setValue('학생 활동 → 원본 로그 시트에 자동 기록 → [새로고침] 클릭 → 공간잇기_교사용에 요약 → 이 표지에서 확인').setWrap(true); row++;
  cov.getRange(flowHeadRow,1,row-flowHeadRow,3).setBackground('#eefaf1').setBorder(true,true,true,true,false,false,'#bfe3cb',SpreadsheetApp.BorderStyle.SOLID);

  try{ cov.autoResizeRows(3, row-3); }catch(e){}
  return cov;
}

function lockTeacherSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cov = ss.getSheetByName(COVER_NAME) || buildCoverSheet();
  ss.getSheets().forEach(function(s){ if (s.getName()!==COVER_NAME){ try{ s.hideSheet(); }catch(e){} } });
  setLockLevel(0);
  ss.setActiveSheet(cov);
  SpreadsheetApp.getUi().alert('🔒 모두 잠갔습니다. 표지만 보입니다.');
}

// 1단: 교사용 보드(요약)만 열기 — 원본 학생데이터는 계속 숨김 상태 유지
function unlockBoard(){
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('🔓 교사용 보드 열기', '비밀번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  if (res.getResponseText() !== TEACHER_KEY) { ui.alert('❌ 비밀번호가 틀렸습니다.'); return; }
  setLockLevel(Math.max(getLockLevel(), 1));
  buildDashboard();   // 최신 데이터로 자동 갱신 + 표시 (완료 시 자체 알림 표시)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('공간잇기_교사용');
  if (dash) ss.setActiveSheet(dash);
}

// 2단: 원본 학생데이터(단계별 로그 · 성찰로그 · 접속기록 · 실시간)까지 추가로 열기
function unlockStudentData(){
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('📂 학생 데이터(원본) 열기', '비밀번호를 다시 입력하세요 (원본 데이터는 더 민감한 정보라 한 번 더 확인합니다).', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  if (res.getResponseText() !== TEACHER_KEY) { ui.alert('❌ 비밀번호가 틀렸습니다.'); return; }
  setLockLevel(2);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildDashboard();   // 보드도 함께 열려있도록 보장
  rawDataSheetNames().forEach(function(n){ var s = ss.getSheetByName(n); if (s) { try{ s.showSheet(); }catch(e){} } });
  SpreadsheetApp.getUi().alert('📂 원본 학생 데이터 시트까지 모두 열었습니다.');
}

function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(COVER_NAME)) buildCoverSheet();   // 표지가 없으면 먼저 만들어 순서를 보장
  var old = ss.getSheetByName('공간잇기_교사용');
  if (old) ss.deleteSheet(old);   // 접기(그룹) 상태가 새로고침마다 누적되는 걸 막기 위해 통째로 새로 만듦
  var dash = ss.insertSheet('공간잇기_교사용', 1);   // 표지(0번) 바로 다음 자리

  var data = collectAll();
  var m1 = data.reflect.m1, m2 = data.reflect.m2;
  var row = 1;
  dash.getRange(row,1).setValue('🛰️ 공간잇기_교사용 — 학습 현황').setFontSize(16).setFontWeight('bold'); row++;
  dash.getRange(row,1).setValue('갱신: ' + data.updated +
    '   (기준: 미접속 '+FLAG_INACTIVE_DAYS+'일↑ · 오답누적 '+FLAG_WRONG_TOTAL+'회↑ · 완료율 '+Math.round(EXCEL_MIN_DONE_RATIO*100)+'%↑ 우수)').setFontColor('#888888'); row+=2;

  // 🚨 이탈자
  dash.getRange(row,1).setValue('🚨 이탈자 — 미접속 '+FLAG_INACTIVE_DAYS+'일↑ 또는 중단 상태로 '+FLAG_STALL_DAYS+'일↑ 정체').setFontWeight('bold').setFontSize(12).setFontColor('#c0392b'); row++;
  var dropHead=row;
  dash.getRange(row,1,1,5).setValues([['학번','이름','사유','완료/중단 단계수','마지막 활동']]).setFontWeight('bold').setBackground('#fdecea'); row++;
  if (!data.dropouts.length) { dash.getRange(row,1).setValue('(해당 학생 없음)').setFontColor('#888888'); row++; }
  data.dropouts.forEach(function(a){
    dash.getRange(row,1,1,5).setValues([[a.sid, a.name, a.reasons.join(' / '), a.doneCnt+'완료 · '+a.abandonedCnt+'중단', a.lastSeen]]).setBackground('#ffe8e8');
    row++;
  });
  gridBorder(dash, dropHead, 1, row-1, 5); row++;

  // 🐢 학습부진자 — 오답형
  dash.getRange(row,1).setValue('🐢 학습부진자 — 오답형 (누적 오답 '+FLAG_WRONG_TOTAL+'회↑)').setFontWeight('bold').setFontSize(12).setFontColor('#b8860b'); row++;
  var wrongHead=row;
  dash.getRange(row,1,1,3).setValues([['학번','이름','누적 오답']]).setFontWeight('bold').setBackground('#fff6dd'); row++;
  if (!data.strugglingWrong.length) { dash.getRange(row,1).setValue('(해당 학생 없음)').setFontColor('#888888'); row++; }
  data.strugglingWrong.forEach(function(a){ dash.getRange(row,1,1,3).setValues([[a.sid, a.name, a.wrongSum]]); row++; });
  gridBorder(dash, wrongHead, 1, row-1, 3); row++;

  // 🐢 학습부진자 — 정체형
  dash.getRange(row,1).setValue('🐢 학습부진자 — 정체형 (한 단계에 '+Math.round(FLAG_SLOW_MIN_SEC/60)+'분↑ · 평균의 '+FLAG_SLOW_RATIO+'배↑ 머묾)').setFontWeight('bold').setFontSize(12).setFontColor('#b8860b'); row++;
  var stuckHead=row;
  dash.getRange(row,1,1,4).setValues([['학번','이름','단계','소요시간']]).setFontWeight('bold').setBackground('#fff6dd'); row++;
  if (!data.strugglingStuck.length) { dash.getRange(row,1).setValue('(해당 학생 없음)').setFontColor('#888888'); row++; }
  data.strugglingStuck.forEach(function(a){ dash.getRange(row,1,1,4).setValues([[a.sid, a.name, a.stage, fmtSec(a.sec)]]); row++; });
  gridBorder(dash, stuckHead, 1, row-1, 4); row++;

  // 🌟 학습우수자
  dash.getRange(row,1).setValue('🌟 학습우수자 — 완료율 '+Math.round(EXCEL_MIN_DONE_RATIO*100)+'%↑ · 평균오답 '+EXCEL_MAX_AVG_WRONG+'회 이하 · 평균보다 빠름').setFontWeight('bold').setFontSize(12).setFontColor('#1e7d32'); row++;
  var excelHead=row;
  dash.getRange(row,1,1,4).setValues([['학번','이름','완료 단계수','평균오답 · 평균시간']]).setFontWeight('bold').setBackground('#e8f8ee'); row++;
  if (!data.excellent.length) { dash.getRange(row,1).setValue('(해당 학생 없음)').setFontColor('#888888'); row++; }
  data.excellent.forEach(function(a){ dash.getRange(row,1,1,4).setValues([[a.sid, a.name, a.doneCnt, a.avgWrong+'회 · '+fmtSec(a.avgSec)]]); row++; });
  gridBorder(dash, excelHead, 1, row-1, 4); row+=2;

  // ✅ 모듈1(삼수선) · 모듈2(정사영) 완료 학생
  var doneM1 = Object.keys(data.students).map(function(sid){return data.students[sid];})
    .filter(function(st){ var r=st.stages['삼수선M3']; return r && r.status==='완료'; })
    .sort(function(a,b){ return (a.name||'').localeCompare(b.name||''); });
  var doneM2 = Object.keys(data.students).map(function(sid){return data.students[sid];})
    .filter(function(st){ var r=st.stages['정사영_미션탐구']; return r && r.status==='완료'; })
    .sort(function(a,b){ return (a.name||'').localeCompare(b.name||''); });
  dash.getRange(row,1).setValue('✅ 모듈1(삼수선) 완료 학생').setFontWeight('bold').setFontSize(12).setFontColor('#1e6b3a'); row++;
  var done1Head=row;
  dash.getRange(row,1,1,3).setValues([['학번','이름','완료 일시']]).setFontWeight('bold').setBackground('#eef7ef'); row++;
  if (!doneM1.length) { dash.getRange(row,1).setValue('(아직 없음)').setFontColor('#888888'); row++; }
  doneM1.forEach(function(st){ var r=st.stages['삼수선M3']; dash.getRange(row,1,1,3).setValues([[st.sid, st.name, r.at||'-']]); row++; });
  gridBorder(dash, done1Head, 1, row-1, 3); row++;

  dash.getRange(row,1).setValue('✅ 모듈2(정사영) 완료 학생').setFontWeight('bold').setFontSize(12).setFontColor('#1e6b3a'); row++;
  var done2Head=row;
  dash.getRange(row,1,1,3).setValues([['학번','이름','완료 일시']]).setFontWeight('bold').setBackground('#eef7ef'); row++;
  if (!doneM2.length) { dash.getRange(row,1).setValue('(아직 없음)').setFontColor('#888888'); row++; }
  doneM2.forEach(function(st){ var r=st.stages['정사영_미션탐구']; dash.getRange(row,1,1,3).setValues([[st.sid, st.name, r.at||'-']]); row++; });
  gridBorder(dash, done2Head, 1, row-1, 3); row+=2;

  dash.getRange(row,1).setValue('🔽 상세 데이터 — 왼쪽 여백의 [+]를 눌러 펼치기').setFontWeight('bold').setFontColor('#607080'); row++;
  var detailStart = row;

  // 단계별 활동 요약
  dash.getRange(row,1).setValue('📈 단계별 활동 요약 (완료 학생 기준 평균)').setFontWeight('bold').setFontSize(12); row++;
  var actHead = row;
  dash.getRange(row,1,1,7).setValues([['단계','완료','중단','평균 활동시간(초)','평균 오답','평균 힌트','평균 이탈']]).setFontWeight('bold').setBackground('#eef2f8'); row++;
  data.stages.forEach(function(s){
    dash.getRange(row,1,1,7).setValues([[s.label, s.done, s.abandoned, s.avgSec, s.avgWrong, s.avgHints, s.avgLeaves]]); row++;
  });
  var actEnd = row-1; gridBorder(dash, actHead, 1, actEnd, 7); row++;

  // 학생별 진행표
  dash.getRange(row,1).setValue('🧑‍🎓 학생별 진행 (완료=✅ 시간/오답, 중단=⏸, 미참여=·)').setFontWeight('bold').setFontSize(12); row++;
  var progHead = row;
  var head = ['학번','이름'].concat(data.stages.map(function(s){return s.label;}));
  dash.getRange(row,1,1,head.length).setValues([head]).setFontWeight('bold').setBackground('#eef2f8'); row++;
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
  gridBorder(dash, progHead, 1, row-1, head.length); row++;

  // 성찰로그 제출 현황
  dash.getRange(row,1).setValue('📥 성찰로그 제출 현황').setFontWeight('bold').setFontSize(12); row++;
  var reflHead = row;
  dash.getRange(row,1,1,3).setValues([['모듈','제출 인원','제출 완료 학생(학번)']]).setFontWeight('bold').setBackground('#eef2f8'); row++;
  dash.getRange(row,1,1,3).setValues([['모듈1 · 삼수선', m1.count, m1.sids.join(', ')]]); row++;
  dash.getRange(row,1,1,3).setValues([['모듈2 · 정사영', m2.count, m2.sids.join(', ')]]); row++;
  gridBorder(dash, reflHead, 1, row-1, 3); row++;

  // 사고 유형 분포
  dash.getRange(row,1).setValue('🧠 사고 유형 분포').setFontWeight('bold').setFontSize(12); row++;
  var typeHeaderRow = row;
  dash.getRange(row,1,1,3).setValues([['사고 유형','모듈1','모듈2']]).setFontWeight('bold').setBackground('#eef2f8'); row++;
  THINK_TYPES.forEach(function(t){
    dash.getRange(row,1,1,3).setValues([[t+' '+TYPE_LABEL[t], m1.typeCount[t]||0, m2.typeCount[t]||0]]); row++;
  });
  var typeEndRow = row-1; gridBorder(dash, typeHeaderRow, 1, typeEndRow, 3); row++;

  // 자유 서술
  dash.getRange(row,1).setValue('✏️ 자유 서술 모아보기').setFontWeight('bold').setFontSize(12); row++;
  var noteHead = row;
  dash.getRange(row,1,1,4).setValues([['모듈','학번','이름','서술']]).setFontWeight('bold').setBackground('#eef2f8'); row++;
  var notes = m1.notes.concat(m2.notes);
  if(!notes.length){ dash.getRange(row,1).setValue('(아직 자유 서술 응답이 없습니다)').setFontColor('#888888'); row++; }
  else notes.forEach(function(n){ dash.getRange(row,1,1,4).setValues([[n.mod,n.sid,n.name,n.text]]); row++; });
  gridBorder(dash, noteHead, 1, row-1, 4);
  var detailEnd = row-1;

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

  // 상세 데이터 구간을 접어서 기본적으로 숨김 — 핵심 요약만 바로 보이게
  if (detailEnd >= detailStart) {
    try {
      dash.getRange(detailStart,1,detailEnd-detailStart+1,1).shiftRowGroupDepth(1);
      dash.getRowGroup(detailStart, 1).collapse();
    } catch (e) { /* 그룹 생성 실패해도 데이터 자체는 정상 */ }
  }

  // 잠긴 상태에서 (열지 않고) 새로고침만 눌렀다면, 갱신된 시트도 그대로 숨김 유지
  if (getLockLevel() < 1) { try{ dash.hideSheet(); }catch(e){} }

  SpreadsheetApp.getUi().alert('✅ 공간잇기_교사용 시트를 갱신했습니다. (상세 데이터는 접혀있어요 — 왼쪽 [+]로 펼치기)');
}

function readModule(ss, sheetName) {
  var out = { count: 0, sids: [], typeCount: {}, notes: [], types: [] };
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
    THINK_TYPES.forEach(function(t){ if (method.indexOf(t) === 0){ out.typeCount[t] = (out.typeCount[t]||0)+1; out.types.push({sid:String(r[2]||''), t:t}); } });
    var note = String(r[5] || '').trim();
    if (note) out.notes.push({ mod:(sheetName==='모듈2'?'정사영':'삼수선'), sid:r[2], name:r[3], text:note });
  }
  return out;
}

// ───────────────────────────────────────────────
// 교사용 웹 대시보드 HTML (Google Charts)
// ───────────────────────────────────────────────
var LOGIN_HTML = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
'<style>' +
'body{font-family:"Noto Sans KR",sans-serif;background:#0b1424;color:#dbe7f7;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box}' +
'.box{background:#111e33;border:1px solid #223a60;border-radius:14px;padding:32px 28px;max-width:340px;width:100%;box-shadow:0 8px 30px rgba(0,0,0,.4);text-align:center}' +
'.box h1{font-size:20px;margin:0 0 6px}' +
'.box .sub{color:#7f9cc4;font-size:12.5px;margin-bottom:22px;line-height:1.5}' +
'.box input{width:100%;box-sizing:border-box;background:#0b1424;border:1px solid #2e4d78;border-radius:8px;color:#dbe7f7;padding:11px 12px;font-size:15px;margin-bottom:12px;text-align:center;letter-spacing:.15em}' +
'.box button{width:100%;background:#5599ff;border:none;border-radius:8px;color:#06111f;font-weight:900;padding:11px 12px;font-size:15px;cursor:pointer}' +
'.err{color:#ff8a75;font-size:12.5px;margin-bottom:12px}' +
'</style></head><body>' +
'<form class="box" method="GET" action="{{ACTION}}">' +
'<h1>🛰️ 공간잇기 교사용</h1>' +
'<div class="sub">비밀번호를 입력하면<br>실시간 학습 현황으로 바로 이동합니다</div>' +
'{{ERR}}' +
'<input type="password" name="key" placeholder="비밀번호" autofocus>' +
'<button type="submit">입장 →</button>' +
'</form></body></html>';

var DASH_HTML = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
'<script src="https://www.gstatic.com/charts/loader.js"></script>' +
'<style>' +
':root{--bg:#0b1424;--fg:#dbe7f7;--sub:#7f9cc4;--line:#1c2f4d;--cardbg:#111e33;--cardbd:#223a60;--hover:#152540;--shadow:rgba(0,0,0,.25);--thbg:#111e33}' +
'body.light{--bg:#f3f6fb;--fg:#1c2636;--sub:#5c6c88;--line:#dde5f1;--cardbg:#ffffff;--cardbd:#dde5f1;--hover:#eef3fb;--shadow:rgba(30,60,110,.10);--thbg:#eef2f8}' +
'body{font-family:"Noto Sans KR",sans-serif;background:var(--bg);color:var(--fg);margin:0;padding:16px 20px;transition:background .2s,color .2s}' +
'h1{font-size:20px;margin:0}.sub{color:var(--sub);font-size:12px}' +
'.top{display:flex;align-items:center;gap:14px;margin-bottom:10px;padding-bottom:12px;border-bottom:1px solid var(--line);flex-wrap:wrap}.top .sp{flex:1}' +
'.themeBtn{background:var(--cardbg);border:1px solid var(--cardbd);color:var(--fg);border-radius:8px;padding:7px 12px;cursor:pointer;font-size:13px}' +
'.miniStat{font-size:12.5px;color:var(--sub);margin-bottom:12px}' +
'.filterbar{display:flex;gap:14px;flex-wrap:wrap;align-items:center;background:var(--cardbg);border:1px solid var(--cardbd);border-radius:10px;padding:9px 12px;margin-bottom:14px;font-size:12.5px;color:var(--sub)}' +
'.filterbar select{background:var(--bg);color:var(--fg);border:1px solid var(--cardbd);border-radius:6px;padding:4px 8px;font-size:12.5px;font-family:inherit;margin-left:5px}' +
'.filterbar #fReset{margin-left:auto;background:transparent;border:1px solid var(--cardbd);color:var(--sub);border-radius:6px;padding:5px 10px;font-size:12px;cursor:pointer;font-family:inherit}' +
'.navrow{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}' +
'.navbtn{display:block;background:var(--cardbg);border:1px solid var(--cardbd);border-radius:10px;padding:12px 14px;box-shadow:0 2px 8px var(--shadow);text-decoration:none;color:inherit;transition:transform .12s}' +
'.navbtn:hover{transform:translateY(-2px)}' +
'.navbtn .nl{font-size:11px;color:var(--sub);letter-spacing:.06em}.navbtn .nv{font-size:24px;font-weight:900;margin-top:2px}' +
'.navbtn.crit .nv{color:#ff6b6b}.navbtn.warnc .nv{color:#ffb84d}.navbtn.warn .nv{color:#ff8a65}.navbtn.good .nv{color:#2fb673}' +
'.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:14px}' +
'.card{background:var(--cardbg);border:1px solid var(--cardbd);border-radius:10px;padding:14px;box-shadow:0 2px 8px var(--shadow)}.card h2{font-size:14px;margin:0 0 8px;color:#5599ff}' +
'.card:target{animation:flash 1.4s ease}' +
'@keyframes flash{0%{box-shadow:0 0 0 3px #5599ff}100%{box-shadow:0 2px 8px var(--shadow)}}' +
'.subh{font-size:12px;font-weight:700;color:var(--sub);margin:2px 0 6px}' +
'.chart{width:100%;height:260px}' +
'table{border-collapse:collapse;width:100%;font-size:12px}th,td{border-bottom:1px solid var(--cardbd);padding:5px 6px;text-align:left;white-space:nowrap}th{color:#5599ff;position:sticky;top:0;background:var(--thbg)}tbody tr:hover{background:var(--hover)}' +
'.wrap{max-height:460px;overflow:auto}.no{color:#8a97ab}.ok{color:#2fb673}.ab{color:#d98a3d}.warn{color:#e06a53}' +
'.att{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px}' +
'.st{background:#1a0e10;border:1px solid #6a2a2a;border-left:4px solid #ff5a3a;border-radius:8px;padding:9px 11px;font-size:12.5px;color:#f2d9d2}.st b{font-size:14px;color:#fff}.st .tg{display:inline-block;margin:4px 4px 0 0;padding:2px 7px;border-radius:999px;background:#3a1010;color:#ffb0a0;font-size:11px;font-weight:700}.st .m{color:#e0aa9a;font-size:11px;margin-top:3px}' +
'body.light .st{background:#fff3ef;border-color:#f0c3b3;color:#7a3520}body.light .st b{color:#5a2010}body.light .st .m{color:#a05030}body.light .st .tg{background:#ffe0d0;color:#a03818}' +
'.st.off{opacity:.55;border-left-color:#8a97ab}' +
'.cell{display:inline-block;min-width:34px;text-align:center;padding:3px 6px;border-radius:5px;font-weight:700}.c-done{background:#0f3a24;color:#8affc0}.c-ab{background:#3a2410;color:#ffcf9a}.c-live{background:#102a4a;color:#8ec3ff;box-shadow:0 0 0 1px #5599ff}.c-no{background:#141c2c;color:#5c6c88}' +
'body.light .c-done{background:#d9f5e6;color:#127a41}body.light .c-ab{background:#ffe9d2;color:#a15e12}body.light .c-live{background:#dcebff;color:#1a5fb4}body.light .c-no{background:#eef1f6;color:#93a0b5}' +
'.note{font-size:12.5px;line-height:1.6;padding:6px 8px;border-left:3px solid #3a5a90;margin:6px 0;background:var(--hover);white-space:normal}' +
'.legend{font-size:11px;color:var(--sub);margin-top:6px}' +
'.grid4{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:10px;margin-bottom:14px}' +
'.card.crit{border-left:4px solid #ff5c5c}.card.crit h2{color:#e05555}' +
'.card.warnc{border-left:4px solid #ffcc55}.card.warnc h2{color:#c98a1e}' +
'.card.good{border-left:4px solid #2fb673}.card.good h2{color:#1e8c58}' +
'.detailBtnWrap{margin:6px 0 14px}' +
'#detailBtn{background:var(--cardbg);border:1px solid var(--cardbd);color:var(--fg);border-radius:8px;padding:9px 14px;cursor:pointer;font-size:13px}' +
'</style></head><body>' +
'<div class="top"><div><h1>🛰️ 공간잇기_교사용 — 실시간 학습 현황</h1><div class="sub" id="upd"></div></div><div class="sp"></div>' +
'<button class="themeBtn" id="themeBtn" onclick="toggleTheme()">☀️ 밝게</button>' +
'<label style="font-size:12px;color:var(--sub)"><input type="checkbox" id="auto" checked> 45초마다 자동 갱신</label><button onclick="refresh()" style="background:var(--cardbg);border:1px solid #5599ff;color:#5599ff;border-radius:6px;padding:6px 12px;cursor:pointer">🔄 지금 갱신</button></div>' +
'<div class="miniStat" id="miniStat">접속 학생 - · 지금 활동 중 -</div>' +
'<div class="filterbar">'+
'<label>🏫 반 <select id="fClass"><option value="">전체</option></select></label>'+
'<label>📅 날짜(완료 기준) <select id="fDate"><option value="">전체</option></select></label>'+
'<label>📘 모듈 <select id="fMod"><option value="">전체</option><option value="1">모듈1(삼수선)</option><option value="2">모듈2(정사영)</option></select></label>'+
'<button id="fReset" type="button">필터 초기화</button>'+
'</div>' +
'<div class="navrow">' +
'<a href="#sec-warn" class="navbtn warn"><div class="nl">⚠ 주의 학생</div><div class="nv" id="nWarn">-</div></a>' +
'<a href="#sec-drop" class="navbtn crit"><div class="nl">🚨 이탈자</div><div class="nv" id="nDrop">-</div></a>' +
'<a href="#sec-strug" class="navbtn warnc"><div class="nl">🐢 학습부진자</div><div class="nv" id="nStrug">-</div></a>' +
'<a href="#sec-excel" class="navbtn good"><div class="nl">🌟 학습우수자</div><div class="nv" id="nExcel">-</div></a>' +
'<a href="#sec-done1" class="navbtn"><div class="nl">✅ 모듈1(삼수선) 완료</div><div class="nv" id="nDone1">-</div></a>' +
'<a href="#sec-done" class="navbtn"><div class="nl">✅ 모듈2(정사영) 완료</div><div class="nv" id="nDone">-</div></a>' +
'</div>' +
'<div class="card" style="margin-bottom:14px" id="sec-warn"><h2>⚠ 지금 살펴볼 학생 <span class="legend">— 오답 3회↑ · 한 단계 8분↑ · 창 이탈 2회↑ (최근 수신 순)</span></h2><div class="att" id="att"></div></div>' +
'<div class="grid4">' +
'<div class="card crit" id="sec-drop"><h2>🚨 이탈자</h2><div class="legend" style="margin:-4px 0 8px">미접속 오래됨 또는 중단 상태로 정체</div><div class="wrap" style="max-height:220px"><table id="dropTb"></table></div></div>' +
'<div class="card warnc" id="sec-strug"><h2>🐢 학습부진자</h2>' +
'<div class="subh">오답형</div><div class="wrap" style="max-height:130px"><table id="wrongTb"></table></div>' +
'<div class="subh" style="margin-top:10px">정체형 (한 단계에 오래 머묾)</div><div class="wrap" style="max-height:130px"><table id="stuckTb"></table></div>' +
'</div>' +
'<div class="card good" id="sec-excel"><h2>🌟 학습우수자</h2><div class="legend" style="margin:-4px 0 8px">완료율 높고 · 오답 적고 · 평균보다 빠름</div><div class="wrap" style="max-height:220px"><table id="excelTb"></table></div></div>' +
'<div class="card" id="sec-done1"><h2>✅ 모듈1(삼수선) 완료 학생</h2><div class="wrap" style="max-height:220px"><table id="doneTb1"></table></div></div>' +
'<div class="card" id="sec-done"><h2>✅ 모듈2(정사영) 완료 학생</h2><div class="wrap" style="max-height:220px"><table id="doneTb"></table></div></div>' +
'</div>' +
'<div class="detailBtnWrap"><button id="detailBtn" onclick="toggleDetail()">📂 상세 데이터 보기 (진행 보드 · 그래프)</button></div>' +
'<div id="detailWrap" style="display:none">' +
'<div class="card" style="margin-bottom:14px"><h2>🧑‍🎓 진행 보드 <span class="legend">— 완료 ✅(활동시간′/오답) · 중단 ⏸ · 진행 중 🔵 · 미참여 ·</span></h2><div class="wrap"><table id="tb"></table></div></div>' +
'<div class="grid">' +
'<div class="card"><h2>단계별 완료 · 중단 인원</h2><div id="c1" class="chart"></div></div>' +
'<div class="card"><h2>단계별 평균 활동시간(초) · 평균 오답</h2><div id="c2" class="chart"></div></div>' +
'<div class="card"><h2>단계별 오답 횟수 분포 (학생 수)</h2><div id="c3" class="chart"></div></div>' +
'<div class="card"><h2>🧩 문제해결 — 단계별 막힌 지점</h2><div id="c6" class="chart"></div></div>' +
'<div class="card"><h2>성찰로그 사고 유형 분포</h2><div id="c5" class="chart"></div></div>' +
'<div class="card"><h2>성찰로그 자유 서술</h2><div id="notes" style="max-height:260px;overflow:auto"></div></div>' +
'</div>' +
'</div>' +
'<script>' +
'var D=<?!= json ?>;' +
'google.charts.load("current",{packages:["corechart","bar"]});' +
'function isLight(){ return document.body.classList.contains("light"); }' +
'function chartOpt(extra){ var tc=isLight()?"#33415c":"#c9d8ee", gc=isLight()?"#dde5f1":"#223a60";' +
' var r={backgroundColor:"transparent",legend:{position:"top",textStyle:{color:tc}},hAxis:{textStyle:{color:tc}},vAxis:{textStyle:{color:tc},gridlines:{color:gc},minValue:0},chartArea:{width:"85%",height:"65%"}};' +
' for(var k in extra) r[k]=extra[k]; return r; }' +
'function fmt(s){return Math.floor(s/60)+"′"+(s%60)+"″";}' +
'var detailShown=false;' +
'function toggleDetail(){ detailShown=!detailShown; document.getElementById("detailWrap").style.display=detailShown?"block":"none";' +
' document.getElementById("detailBtn").textContent = detailShown ? "📂 상세 데이터 접기" : "📂 상세 데이터 보기 (진행 보드 · 그래프)";' +
' if(detailShown){ if(window.google&&google.visualization) drawDetail(); else google.charts.setOnLoadCallback(drawDetail); } }' +
'function toggleTheme(){ var light=!isLight(); document.body.classList.toggle("light",light);' +
' document.getElementById("themeBtn").textContent = light ? "🌙 어둡게" : "☀️ 밝게";' +
' try{ localStorage.setItem("kosa_theme", light?"light":"dark"); }catch(e){}' +
' if(detailShown) drawDetail(); }' +
'(function(){ var t="dark"; try{ t=localStorage.getItem("kosa_theme")||"dark"; }catch(e){}' +
' if(t==="light"){ document.body.classList.add("light"); document.getElementById("themeBtn").textContent="🌙 어둡게"; } })();' +
'function classOf(sid){ sid=String(sid); return sid.length>=3 ? sid.slice(0,3) : sid; }' +
'function populateFilters(){' +
' var cs=document.getElementById("fClass"); if(cs.options.length<=1){ var classes={}; Object.keys(D.students).forEach(function(sid){ classes[classOf(sid)]=1; });' +
'   Object.keys(classes).sort().forEach(function(c){ var o=document.createElement("option"); o.value=c; o.textContent=c.length===3?(c[0]+"학년 "+c.slice(1)+"반"):c; cs.appendChild(o); }); }' +
' var ds=document.getElementById("fDate"); if(ds.options.length<=1){ var dates={};' +
'   Object.keys(D.students).forEach(function(sid){ var st=D.students[sid]; Object.keys(st.stages).forEach(function(k){ var r=st.stages[k]; if(r&&r.at) dates[r.at.slice(0,5)]=1; }); });' +
'   Object.keys(dates).sort().forEach(function(d){ var o=document.createElement("option"); o.value=d; o.textContent=d+"일"; ds.appendChild(o); }); } }' +
'function fClassV(){ return document.getElementById("fClass").value; }' +
'function fDateV(){ return document.getElementById("fDate").value; }' +
'function fModV(){ return document.getElementById("fMod").value; }' +
'function passClass(sid){ var v=fClassV(); return !v || classOf(sid)===v; }' +
'function passDate(atStr){ var v=fDateV(); return !v || (atStr && atStr.slice(0,5)===v); }' +
'function drawCore(){' +
' populateFilters();' +
' document.getElementById("upd").textContent="갱신: "+D.updated;' +
' var allSids=Object.keys(D.students).filter(passClass);' +
' var live=(D.live||[]).filter(function(l){return passClass(l.sid);}), on=live.filter(function(l){return l.online;}), warn=live.filter(function(l){return l.flags.length;});' +
' document.getElementById("miniStat").textContent="접속 학생 "+allSids.length+"명 · 지금 활동 중 "+on.length+"명"+(fClassV()||fDateV()?" · (필터 적용 중)":"");' +
' document.getElementById("nWarn").textContent=warn.length+"명";' +
' document.getElementById("att").innerHTML = warn.length ? warn.map(function(l){return "<div class=st"+(l.online?"":" off")+"><b>"+l.sid+" "+l.name+"</b> <span class=m>"+l.stage+" · "+fmt(l.sec)+(l.online?" · 활동 중":" · "+Math.round(l.ago/60)+"분 전 수신")+"</span><br>"+l.flags.map(function(f){return "<span class=tg>"+f+"</span>";}).join("")+"</div>";}).join("") : "<div class=no>지금은 주의가 필요한 학생이 없습니다.</div>";' +
' var drop=(D.dropouts||[]).filter(function(a){return passClass(a.sid);}), sw=(D.strugglingWrong||[]).filter(function(a){return passClass(a.sid);}), ss=(D.strugglingStuck||[]).filter(function(a){return passClass(a.sid);}), ex=(D.excellent||[]).filter(function(a){return passClass(a.sid);});' +
' document.getElementById("nDrop").textContent=drop.length+"명"; document.getElementById("nStrug").textContent=(sw.length+ss.length)+"건"; document.getElementById("nExcel").textContent=ex.length+"명";' +
' document.getElementById("dropTb").innerHTML = drop.length ? "<tr><th>학번</th><th>이름</th><th>사유</th></tr>"+drop.map(function(a){return "<tr><td><b>"+a.sid+"</b></td><td>"+a.name+"</td><td class=warn>"+a.reasons.join(" / ")+"</td></tr>";}).join("") : "<tr><td class=no>해당 학생 없음</td></tr>";' +
' document.getElementById("wrongTb").innerHTML = sw.length ? "<tr><th>학번</th><th>이름</th><th>누적 오답</th></tr>"+sw.map(function(a){return "<tr><td><b>"+a.sid+"</b></td><td>"+a.name+"</td><td class=ab><b>"+a.wrongSum+"</b></td></tr>";}).join("") : "<tr><td class=no>해당 학생 없음</td></tr>";' +
' document.getElementById("stuckTb").innerHTML = ss.length ? "<tr><th>학번</th><th>이름</th><th>단계</th><th>소요시간</th></tr>"+ss.map(function(a){return "<tr><td><b>"+a.sid+"</b></td><td>"+a.name+"</td><td>"+a.stage+"</td><td class=ab>"+fmt(a.sec)+"</td></tr>";}).join("") : "<tr><td class=no>해당 학생 없음</td></tr>";' +
' document.getElementById("excelTb").innerHTML = ex.length ? "<tr><th>학번</th><th>이름</th><th>완료</th><th>평균오답·시간</th></tr>"+ex.map(function(a){return "<tr><td><b>"+a.sid+"</b></td><td>"+a.name+"</td><td class=ok>"+a.doneCnt+"</td><td>"+a.avgWrong+"회 · "+fmt(a.avgSec)+"</td></tr>";}).join("") : "<tr><td class=no>해당 학생 없음</td></tr>";' +
' var sec1=document.getElementById("sec-done1"), sec2=document.getElementById("sec-done");' +
' if(sec1) sec1.style.display = (fModV()==="2") ? "none" : "";' +
' if(sec2) sec2.style.display = (fModV()==="1") ? "none" : "";' +
' var doneList1=allSids.map(function(sid){return D.students[sid];}).filter(function(st){var r=st.stages["삼수선M3"]; return r&&r.status==="완료"&&passDate(r.at);}).sort(function(a,b){return (a.name||"").localeCompare(b.name||"");});' +
' document.getElementById("nDone1").textContent=doneList1.length+"명";' +
' document.getElementById("doneTb1").innerHTML = doneList1.length ? "<tr><th>학번</th><th>이름</th><th>완료 일시</th></tr>"+doneList1.map(function(st){var r=st.stages["삼수선M3"];return "<tr><td><b>"+st.sid+"</b></td><td>"+st.name+"</td><td class=no>"+(r.at||"-")+"</td></tr>";}).join("") : "<tr><td class=no>해당 없음</td></tr>";' +
' var doneList=allSids.map(function(sid){return D.students[sid];}).filter(function(st){var r=st.stages["정사영_미션탐구"]; return r&&r.status==="완료"&&passDate(r.at);}).sort(function(a,b){return (a.name||"").localeCompare(b.name||"");});' +
' document.getElementById("nDone").textContent=doneList.length+"명";' +
' document.getElementById("doneTb").innerHTML = doneList.length ? "<tr><th>학번</th><th>이름</th><th>완료 일시</th></tr>"+doneList.map(function(st){var r=st.stages["정사영_미션탐구"];return "<tr><td><b>"+st.sid+"</b></td><td>"+st.name+"</td><td class=no>"+(r.at||"-")+"</td></tr>";}).join("") : "<tr><td class=no>해당 없음</td></tr>";' +
'}' +
'function avgN(arr){ if(!arr.length) return 0; return Math.round((arr.reduce(function(a,b){return a+b;},0)/arr.length)*10)/10; }' +
'function computeStages(sids){' +
' return D.stages.map(function(base){' +
'   var st={sheet:base.sheet,label:base.label,done:0,abandoned:0,avgSec:0,avgWrong:0,avgHints:0,avgLeaves:0,wrongDist:{},stepWrong:{}}, doneRows=[];' +
'   sids.forEach(function(sid){ var r=D.students[sid]&&D.students[sid].stages[base.sheet]; if(!r) return;' +
'     if(r.status==="완료"){ st.done++; doneRows.push(r); } else st.abandoned++;' +
'     var w=Math.min(r.wrong,5), k=w>=5?"5+":String(w); st.wrongDist[k]=(st.wrongDist[k]||0)+1;' +
'     try{ var wd=JSON.parse(r.wrongDetail||"{}"); Object.keys(wd).forEach(function(q){ if(!st.stepWrong[q]) st.stepWrong[q]={total:0,students:0}; st.stepWrong[q].total+=Number(wd[q]||0); st.stepWrong[q].students+=1; }); }catch(e){}' +
'   });' +
'   if(doneRows.length){ st.avgSec=Math.round(doneRows.reduce(function(a,r){return a+r.sec;},0)/doneRows.length); st.avgWrong=avgN(doneRows.map(function(r){return r.wrong;})); st.avgHints=avgN(doneRows.map(function(r){return r.hints;})); st.avgLeaves=avgN(doneRows.map(function(r){return r.leaves;})); }' +
'   return st;' +
' }); }' +
'function drawDetail(){' +
' var allSids=Object.keys(D.students).filter(passClass);' +
' var stages=computeStages(allSids);' +
' var live=D.live||[]; var liveMap={}; live.forEach(function(l){ liveMap[l.sid]=l; });' +
' var h="<tr><th>학번</th><th>이름</th><th>현재</th>"+D.stages.map(function(s){return "<th>"+s.label+"</th>";}).join("")+"</tr>";' +
' Object.keys(D.students).filter(passClass).sort().forEach(function(sid){var st=D.students[sid], lv=liveMap[sid];' +
'  h+="<tr><td>"+sid+"</td><td>"+st.name+"</td><td>"+(lv?(lv.online?"<span class=ok>● </span>":"<span class=no>○ </span>")+lv.stage+" "+fmt(lv.sec):"<span class=no>-</span>")+"</td>"+D.stages.map(function(s){var r=st.stages[s.sheet];' +
'   if(lv&&lv.online&&lv.stage===s.label&&!(r&&r.status==="완료")) return "<td><span class=\\"cell c-live\\">🔵 "+fmt(lv.sec)+"</span></td>";' +
'   if(!r) return "<td><span class=\\"cell c-no\\">·</span></td>";' +
'   if(r.status==="완료") return "<td><span class=\\"cell c-done\\">✅ "+fmt(r.sec)+" / "+r.wrong+"</span></td>";' +
'   return "<td><span class=\\"cell c-ab\\">⏸ "+fmt(r.sec)+" / "+r.wrong+"</span></td>";}).join("")+"</tr>";});' +
' document.getElementById("tb").innerHTML=h;' +
' var d1=[["단계","완료","중단"]];stages.forEach(function(s){d1.push([s.label,s.done,s.abandoned]);});new google.visualization.ColumnChart(document.getElementById("c1")).draw(google.visualization.arrayToDataTable(d1),chartOpt({colors:["#4fd18b","#ffb070"],isStacked:true}));' +
' var d2=[["단계","평균 활동시간(초)","평균 오답"]];stages.forEach(function(s){d2.push([s.label,s.avgSec,s.avgWrong]);});new google.visualization.ColumnChart(document.getElementById("c2")).draw(google.visualization.arrayToDataTable(d2),chartOpt({colors:["#5aa9ff","#ff6b8a"],series:{1:{targetAxisIndex:1}}}));' +
' var keys=["0","1","2","3","4","5+"];var d3=[["단계"].concat(keys.map(function(k){return "오답 "+k+"회";}))];stages.forEach(function(s){d3.push([s.label].concat(keys.map(function(k){return s.wrongDist[k]||0;})));});new google.visualization.ColumnChart(document.getElementById("c3")).draw(google.visualization.arrayToDataTable(d3),chartOpt({isStacked:true,colors:["#4fd18b","#9ad5ff","#ffd166","#ffb070","#ff8a65","#ff5c7a"]}));' +
' var PB=stages.filter(function(s){return s.sheet==="삼수선_문제해결";})[0];var PBN={step1:"1 C의 위치",step2:"2 BD′",step3:"3① ∠AD′B",step4:"3② AD′",step5:"4① D′H⊥AB",step6:"4② 삼수선",step7:"5① D′H",step8:"5② DH",step9:"6 정답"};' +
' if(PB){var d6=[["단계","오답 총횟수","틀린 학생 수"]];Object.keys(PBN).forEach(function(k){var v=PB.stepWrong[k]||{total:0,students:0};d6.push([PBN[k],v.total,v.students]);});new google.visualization.ColumnChart(document.getElementById("c6")).draw(google.visualization.arrayToDataTable(d6),chartOpt({colors:["#ff6b8a","#ffd166"]}));} else document.getElementById("c6").textContent="(아직 기록 없음)";' +
' var T={"🔢":"논리·공식형","🔁":"탐색·시행착오형","👁":"직관·관찰형","📋":"절차·분석형"};' +
' var t1=(D.reflect.m1.types||[]).filter(function(x){return passClass(x.sid);}), t2=(D.reflect.m2.types||[]).filter(function(x){return passClass(x.sid);});' +
' function cnt(list,k){ return list.filter(function(x){return x.t===k;}).length; }' +
' var d5=[["유형","모듈1","모듈2"]];Object.keys(T).forEach(function(k){d5.push([k+" "+T[k],cnt(t1,k),cnt(t2,k)]);});new google.visualization.ColumnChart(document.getElementById("c5")).draw(google.visualization.arrayToDataTable(d5),chartOpt({colors:["#ff9355","#ffd34a"]}));' +
' var ns=D.reflect.m1.notes.concat(D.reflect.m2.notes).filter(function(n){return passClass(n.sid);});document.getElementById("notes").innerHTML=ns.length?ns.map(function(n){return "<div class=note><b>["+n.mod+"] "+n.sid+" "+n.name+"</b> — "+n.text+"</div>";}).join(""):"<div class=no>(아직 응답 없음)</div>";' +
'}' +
'function refresh(){ google.script.run.withSuccessHandler(function(j){ D=JSON.parse(j); drawCore(); if(detailShown) drawDetail(); }).getData(); }' +
'["fClass","fDate","fMod"].forEach(function(id){ document.getElementById(id).addEventListener("change", function(){ drawCore(); if(detailShown) drawDetail(); }); });' +
'document.getElementById("fReset").addEventListener("click", function(){ document.getElementById("fClass").value=""; document.getElementById("fDate").value=""; document.getElementById("fMod").value=""; drawCore(); if(detailShown) drawDetail(); });' +
'drawCore();' +
'setInterval(function(){ if(document.getElementById("auto").checked) refresh(); },45000);' +
'</script></body></html>';