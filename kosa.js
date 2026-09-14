/* ══════════════════════════════════════════════════════════════════
   kosa.js — 공간잇기 공통 라이브러리
   ─────────────────────────────────────────────────────────────────
   ① 화면 배율  : 태블릿(약 1180×760) 기준 레이아웃을 PC 큰 화면에서
                  글자·메뉴가 함께 커지도록 <body>를 통째로 확대
   ② 로그인     : index.html에서 입력한 학번·이름을 localStorage에 보관,
                  모듈은 KOSA.requireLogin()으로 접속 차단
   ③ 활동 추적  : 단계별 소요시간(실제 활동시간) · 오답 횟수 · 힌트 ·
                  창이탈 횟수/시간을 sessionStorage에 누적 → 완료 시 시트 전송
   ④ 창이탈     : 탭 전환·앱 전환·창 최소화 감지, 복귀 시 안내 토스트
   ⑤ 성찰로그   : 로그인 정보로 자동 채워 전송
   ─────────────────────────────────────────────────────────────────
   모든 모듈 HTML의 <head> 맨 앞에서 <script src="kosa.js"></script>
   ══════════════════════════════════════════════════════════════════ */
(function(){
  var KOSA = window.KOSA = window.KOSA || {};

  // ▼▼▼ 선생님 설정: Apps Script 웹 앱 URL (reflectlog_apps_script.gs 배포 후) ▼▼▼
  KOSA.SHEET_URL = 'https://script.google.com/macros/s/AKfycby1AZ5m-tBNb3Wj3v59oeIiXD1K4fu8G9m_xqDbJkJVYyMTCotcWN9xcqRmDet17UVqhw/exec';

  // 단계 키 → 구글시트 탭 이름 (Apps Script도 이 이름으로 시트를 만듭니다)
  KOSA.STAGES = {
    S_M1:'삼수선M1',        S_M2:'삼수선M2',        S_M3:'삼수선M3',        S_P1:'삼수선_문제해결', S_HW:'삼수선_추가문제',
    J_F1:'정사영_개념탐구',
    J_A0:'정사영_개념이해1', J_A1:'정사영_개념이해2', J_A2:'정사영_개념이해3',
    J_F3:'정사영_미션탐구', J_HW:'정사영_추가문제'
  };
  KOSA.STAGE_LABEL = {
    S_M1:'삼수선 미션1', S_M2:'삼수선 미션2', S_M3:'삼수선 미션3', S_P1:'삼수선 문제해결', S_HW:'삼수선 추가문제',
    J_F1:'정사영 개념탐구', J_A0:'개념이해 기본개념', J_A1:'개념이해 심화①', J_A2:'개념이해 심화②',
    J_F3:'정사영 미션탐구', J_HW:'정사영 추가문제'
  };

  /* ────────────────────────────────────────────────
     ① 화면 배율 — 태블릿 폭 이하는 그대로(배율 1), 넓은 PC 화면만 확대
     ──────────────────────────────────────────────── */
  var BASE_W = 1180, BASE_H = 760, MAX_Z = 1.8;
  KOSA.ZOOM = 1;
  var synResize = false;
  function applyScale(){
    var b = document.body; if(!b) return;
    if(window.self !== window.top) return;   // iframe 안: 부모 문서가 이미 확대하므로 중복 확대 금지
    var w = window.innerWidth, h = window.innerHeight;
    var z = Math.min(w/BASE_W, h/BASE_H);
    var touch = false; try{ touch = window.matchMedia && window.matchMedia('(pointer:coarse)').matches; }catch(e){}
    if(touch){ if(z < 1.08) z = 1; }          // 태블릿: 원본 그대로 (작은 태블릿을 축소하지 않음)
    else { z = Math.max(0.7, z); if(Math.abs(z-1) < 0.02) z = 1; }   // PC: 창 크기에 맞춰 태블릿 비율 그대로 확대/축소
    z = Math.min(z, MAX_Z);
    z = Math.round(z*1000)/1000;
    if(z === 1){
      b.style.transform=''; b.style.transformOrigin=''; b.style.width=''; b.style.height='';
    }else{
      b.style.transformOrigin='0 0';
      b.style.transform='scale('+z+')';
      b.style.width=(w/z)+'px';
      b.style.height=(h/z)+'px';
      document.documentElement.style.overflow='hidden';
    }
    KOSA.ZOOM = z;
    // 3D 렌더러: 확대된 만큼 해상도를 올려 흐릿함 방지
    var r = window.rdr || window.renderer;
    if(r && r.setPixelRatio){ try{ r.setPixelRatio(Math.min((window.devicePixelRatio||1)*z, 2.5)); }catch(e){} }
  }
  KOSA.applyScale = applyScale;
  function scaleAndNotify(){
    applyScale();
    // 모듈의 자체 resize 핸들러(캔버스 크기 재계산)를 한 번 더 깨움
    if(window.requestAnimationFrame){
      requestAnimationFrame(function(){ synResize=true; try{ window.dispatchEvent(new Event('resize')); }catch(e){} synResize=false; });
    }
  }
  window.addEventListener('resize', function(){ if(synResize) return; applyScale(); });
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', scaleAndNotify);
  else scaleAndNotify();
  window.addEventListener('load', function(){ setTimeout(scaleAndNotify, 60); });

  // (되돌림) 본문 position:fixed·touchmove 차단은 페이지 이동 시 화면이 튀는 원인이 되어 제거함.
  //  문서 스크롤만 잠그고(overflow:hidden), 밀림은 아래 감시로 되돌린다.
  (function(){
    function lock(){ var d=document.documentElement, b=document.body; if(!b) return;
      d.style.overflow='hidden'; d.style.overscrollBehavior='none'; b.style.overflow='hidden'; b.style.overscrollBehavior='none'; }
    if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', lock); else lock();
  })();
  // iOS Safari 화면 밀림 방지: 입력창(키보드) 사용 뒤에만 페이지 스크롤을 0으로 되돌림
  // (스크롤·툴바 변화마다 강제로 되돌리던 방식은 페이지 이동 시 튐을 유발해 제거)
  (function(){
    function inputFocused(){ var a=document.activeElement; return a && (a.tagName==='INPUT'||a.tagName==='TEXTAREA'||a.tagName==='SELECT'||a.isContentEditable); }
    function reset(){ if(!inputFocused() && (window.scrollY||window.pageYOffset||document.documentElement.scrollTop||document.body.scrollTop)){ window.scrollTo(0,0); document.documentElement.scrollTop=0; document.body.scrollTop=0; } }
    document.addEventListener('focusout', function(){ setTimeout(reset,150); setTimeout(reset,450); });
    window.addEventListener('orientationchange', function(){ setTimeout(reset,300); });
  })();

  // Apple Pencil 탭 보정: 펜 탭에 click 이 안 오는 경우 pointerup 으로 보완 (중복 방지)
  document.addEventListener('click', function(e){ var b=e.target.closest&&e.target.closest('button,.mt,.kchip,.pbtn,.adv-btn,.goal-opt'); if(b) b.__lastClick=Date.now(); }, true);
  document.addEventListener('pointerup', function(e){
    if(e.pointerType!=='pen') return;
    var b=e.target.closest&&e.target.closest('button,.mt,.kchip,.pbtn,.adv-btn,.goal-opt'); if(!b||b.disabled) return;
    setTimeout(function(){ if(!b.__lastClick || Date.now()-b.__lastClick>250){ b.__lastClick=Date.now(); b.click(); } }, 120);
  }, true);
  // 안쪽 패널만 스크롤 (scrollIntoView 가 문서 전체를 밀지 않도록)
  KOSA.scrollTo = function(el, block){
    if(!el) return; var p=el.parentElement;
    while(p && p!==document.body){ var cs=getComputedStyle(p); if(/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight>p.clientHeight+1) break; p=p.parentElement; }
    if(!p || p===document.body) return;
    var top = el.getBoundingClientRect().top - p.getBoundingClientRect().top + p.scrollTop - (block==='center' ? (p.clientHeight-el.offsetHeight)/2 : 8);
    if(p.scrollTo) p.scrollTo({top:Math.max(0,top), behavior:'smooth'}); else p.scrollTop=Math.max(0,top);
  };

  // 세로 모드 안내 (태블릿을 세로로 들었을 때) — 닫으면 이 세션 동안 다시 표시 안 함
  function rotateHint(){
    if(window.self!==window.top) return;
    var portrait = window.innerHeight > window.innerWidth && window.innerWidth < 1000;
    var el = document.getElementById('kosa-rotate');
    if(!portrait){ if(el) el.style.display='none'; return; }
    try{ if(sessionStorage.getItem('kosa_rotate_seen')==='1') return; }catch(e){}
    if(!el){
      el=document.createElement('div'); el.id='kosa-rotate';
      el.style.cssText='position:fixed;left:50%;bottom:14px;transform:translateX(-50%);z-index:99998;background:#0b1830;border:1.5px solid #5599ff;border-radius:10px;padding:10px 14px;color:#dfeeff;font:700 13px/1.5 "Noto Sans KR",sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.5);text-align:center;max-width:92%';
      el.innerHTML='🔄 태블릿을 <b style="color:#8ec3ff">가로</b>로 돌리면 3D 화면과 활동 카드를 한눈에 볼 수 있어요.<br><button style="margin-top:6px;background:#1a2a4a;border:1px solid #5599ff;color:#cfe6ff;border-radius:6px;padding:4px 12px;font:inherit;font-size:12px;cursor:pointer">알겠어요</button>';
      el.querySelector('button').onclick=function(){ el.style.display='none'; try{ sessionStorage.setItem('kosa_rotate_seen','1'); }catch(e){} };
      (document.body||document.documentElement).appendChild(el);
    }
    el.style.display='block';
    clearTimeout(el._t); el._t=setTimeout(function(){ el.style.display='none'; try{ sessionStorage.setItem('kosa_rotate_seen','1'); }catch(e){} }, 6000);   // 6초 뒤 자동으로 사라짐 (조작 방해 방지)
  }
  window.addEventListener('resize', rotateHint);
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', rotateHint); else rotateHint();

  /* ────────────────────────────────────────────────
     ② 로그인
     ──────────────────────────────────────────────── */
  KOSA.getUser = function(){
    try{ var u = JSON.parse(localStorage.getItem('kosa_user')); if(u && u.sid && u.name) return u; }catch(e){}
    return null;
  };
  KOSA.setUser = function(sid, name){
    var u = {sid:String(sid).trim(), name:String(name).trim(), at:Date.now()};
    try{ localStorage.setItem('kosa_user', JSON.stringify(u)); }catch(e){}
    return u;
  };
  KOSA.logout = function(){
    try{ localStorage.removeItem('kosa_user'); }catch(e){}
    try{ Object.keys(sessionStorage).forEach(function(k){ if(k.indexOf('kosa_')===0) sessionStorage.removeItem(k); }); }catch(e){}
  };
  // 현재 학번의 모듈 진행 기록만 삭제 (로그인은 유지)
  KOSA.resetProgress = function(sid){
    var u = KOSA.getUser(); sid = sid || (u && u.sid); if(!sid) return 0;
    var n = 0;
    try{ Object.keys(localStorage).forEach(function(k){
      if(k === 'kosa_pp_prog_'+sid || k === 'kosa_pp_lab_'+sid || k === 'kosa_ss_prog_'+sid || k === 'kosa_jsy_stage_'+sid || k === 'kosa_jsy_hw_'+sid || k === 'kosa_jsy_refl_'+sid || k === 'kosa_adv_unlock_'+sid || k === 'kosa_adv_done_'+sid){ localStorage.removeItem(k); n++; }
    }); }catch(e){}
    try{ Object.keys(sessionStorage).forEach(function(k){ if(k.indexOf('kosa_')===0 && k!=='kosa_story_seen') sessionStorage.removeItem(k); }); }catch(e){}
    return n;
  };
  // 이 브라우저의 공간잇기 기록 전부 삭제 (모든 학번 + 로그인)
  KOSA.resetAll = function(){
    try{ Object.keys(localStorage).forEach(function(k){ if(k.indexOf('kosa_')===0) localStorage.removeItem(k); }); }catch(e){}
    try{ Object.keys(sessionStorage).forEach(function(k){ if(k.indexOf('kosa_')===0) sessionStorage.removeItem(k); }); }catch(e){}
  };
  // 로그인 안 했으면 대문으로 돌려보냄 (iframe 안에서는 부모가 이미 확인했으므로 통과)
  KOSA.requireLogin = function(){
    if(window.self !== window.top) return true;
    if(KOSA.getUser()) return true;
    try{ sessionStorage.setItem('kosa_after_login', location.pathname.split('/').pop() + location.search); }catch(e){}
    location.replace('index.html?login=1');
    return false;
  };

  /* ────────────────────────────────────────────────
     전송 유틸 (Apps Script는 form-urlencoded / no-cors 로 받음)
     ──────────────────────────────────────────────── */
  function toParams(obj){
    var p = new URLSearchParams();
    Object.keys(obj).forEach(function(k){ if(obj[k]!==undefined && obj[k]!==null) p.append(k, String(obj[k])); });
    return p.toString();
  }
  // 반환: Promise (성공/실패 모두 resolve — 수업 흐름을 막지 않음)
  KOSA.send = function(obj, useBeacon){
    var url = KOSA.SHEET_URL;
    var body = toParams(obj);
    if(!url){ try{ console.warn('[KOSA] SHEET_URL 미설정 — 콘솔에만 기록', obj); }catch(e){} return Promise.resolve(true); }
    if(useBeacon && navigator.sendBeacon){
      try{ navigator.sendBeacon(url, new Blob([body], {type:'application/x-www-form-urlencoded'})); return Promise.resolve(true); }catch(e){}
    }
    if(!window.fetch) return Promise.resolve(false);
    var req = fetch(url, {method:'POST', mode:'no-cors', keepalive:true,
      headers:{'Content-Type':'application/x-www-form-urlencoded'}, body:body})
      .then(function(){ return true; }, function(){ return false; });
    // 응답이 8초 넘게 없으면(학교망 지연 등) 화면을 붙잡지 않도록 '전송됨'으로 간주하고 넘어감 (no-cors 응답은 원래 확인 불가)
    var timeout = new Promise(function(res){ setTimeout(function(){ res(true); }, 8000); });
    return Promise.race([req, timeout]);
  };

  /* ────────────────────────────────────────────────
     ③ 활동 추적 (단계 단위)
        KOSA.begin('S_M1')  … 단계 시작(또는 재진입)
        KOSA.wrong('step3') … 오답 1회 (현재 단계)
        KOSA.hint()         … 힌트 1회
        KOSA.pause()        … 잠시 벗어남(다른 탭·iframe 진입)
        KOSA.complete('S_M1', {추가정보}) … 완료 → 시트 전송
     ──────────────────────────────────────────────── */
  var cur = null;   // {key, data, segStart}
  function fresh(key){
    return {key:key, activeMs:0, wrong:{}, wrongTotal:0, hints:0, leaves:0, leaveMs:0,
            visits:0, completed:false, firstAt:Date.now(), partialAt:0};
  }
  function load(key){
    try{ var d = JSON.parse(sessionStorage.getItem('kosa_st_'+key)); if(d && d.key===key) return d; }catch(e){}
    return fresh(key);
  }
  function save(d){ try{ sessionStorage.setItem('kosa_st_'+d.key, JSON.stringify(d)); }catch(e){} }
  function tick(){
    if(cur && cur.segStart){ cur.data.activeMs += Date.now()-cur.segStart; cur.segStart = Date.now(); save(cur.data); }
  }
  function pageVisible(){ return !document.hidden && !leaving; }

  KOSA.begin = function(key){
    if(!KOSA.STAGES[key]) return;
    if(cur && cur.key===key) return;
    KOSA.pause();
    var d = load(key);
    if(d.completed){ cur = null; return; }      // 복습 모드: 다시 재지 않음
    d.visits++;
    cur = {key:key, data:d, segStart: pageVisible() ? Date.now() : null};
    save(d);
  };
  KOSA.pause = function(){
    if(!cur) return;
    tick(); cur.segStart = null; save(cur.data); cur = null;
  };
  KOSA.current = function(){ return cur ? cur.key : null; };
  KOSA.wrong = function(qid){
    if(!cur) return;
    var d = cur.data; qid = qid || 'q';
    d.wrong[qid] = (d.wrong[qid]||0)+1; d.wrongTotal++; save(d);
  };
  KOSA.hint = function(){ if(!cur) return; cur.data.hints++; save(cur.data); };
  KOSA.stat = function(key){ return (cur && cur.key===key) ? cur.data : load(key); };

  function rowOf(d, status, extra){
    var u = KOSA.getUser() || {sid:'', name:''};
    return {
      kind:'activity', stage:d.key, sheet:KOSA.STAGES[d.key], label:KOSA.STAGE_LABEL[d.key],
      sid:u.sid, sname:u.name, status:status,
      activeSec:Math.round(d.activeMs/1000),
      wrongTotal:d.wrongTotal, wrongDetail:JSON.stringify(d.wrong),
      hints:d.hints, leaves:d.leaves, leaveSec:Math.round(d.leaveMs/1000), visits:d.visits,
      extra: extra ? JSON.stringify(extra) : ''
    };
  }
  KOSA.complete = function(key, extra){
    if(!KOSA.STAGES[key]) return Promise.resolve(false);
    var d = (cur && cur.key===key) ? cur.data : load(key);
    if(d.completed) return Promise.resolve(false);
    if(cur && cur.key===key){ tick(); cur = null; }
    d.completed = true; d.completedAt = Date.now(); save(d);
    return KOSA.send(rowOf(d, '완료', extra));
  };
  // 완료 전에 페이지를 떠나면 '중단' 행을 남김 (15초 이상 활동했을 때만, 중복 방지)
  function flushPartial(){
    if(!cur) return;
    tick();
    var d = cur.data;
    if(d.completed) return;
    if(d.activeMs - (d.partialAt||0) < 15000) return;
    d.partialAt = d.activeMs; save(d);
    KOSA.send(rowOf(d, '중단'), true);
  }
  // 실시간 상태 전송(하트비트): 단계가 활성이고 화면이 보일 때 60초마다 현재 상태를 '실시간' 시트에 갱신
  function heartbeat(){
    if(!cur || document.hidden || leaving) return;
    tick(); var d = cur.data, u = KOSA.getUser() || {sid:'', name:''};
    KOSA.send({kind:'heartbeat', sid:u.sid, sname:u.name, stage:d.key, label:KOSA.STAGE_LABEL[d.key]||d.key,
      activeSec:Math.round(d.activeMs/1000), wrongTotal:d.wrongTotal, hints:d.hints, leaves:d.leaves, leaveSec:Math.round(d.leaveMs/1000)}, true);
  }
  setInterval(heartbeat, 60000);
  window.addEventListener('pagehide', flushPartial);
  window.addEventListener('beforeunload', flushPartial);

  /* ────────────────────────────────────────────────
     ④ 창이탈 감지
     ──────────────────────────────────────────────── */
  var leaving = false, leaveStart = 0, alertOpen = false;
  var inIframe = (window.self !== window.top);
  // alert()/confirm() 중에는 창이 흐려져도 이탈로 치지 않음
  ['alert','confirm','prompt'].forEach(function(fn){
    var orig = window[fn];
    if(typeof orig !== 'function') return;
    window[fn] = function(){ alertOpen = true; try{ return orig.apply(window, arguments); } finally { setTimeout(function(){ alertOpen=false; }, 50); } };
  });
  function onLeave(){
    if(leaving || alertOpen) return;
    leaving = true; leaveStart = Date.now();
    if(cur){ tick(); cur.segStart = null; cur.data.leaves++; save(cur.data); }
    try{ sessionStorage.setItem('kosa_leave_total', String((+sessionStorage.getItem('kosa_leave_total')||0)+1)); }catch(e){}
  }
  function onReturn(){
    if(!leaving) return;
    leaving = false;
    var ms = Date.now()-leaveStart;
    if(cur){ cur.data.leaveMs += ms; save(cur.data); if(!document.hidden) cur.segStart = Date.now(); }
    if(ms > 1500) showLeaveToast(ms);
  }
  document.addEventListener('visibilitychange', function(){ if(document.hidden) onLeave(); else onReturn(); });
  if(!inIframe){
    // 다른 창으로 이동(포커스 상실). iframe으로 포커스가 옮겨간 경우는 제외
    window.addEventListener('blur', function(){
      setTimeout(function(){
        var ae = document.activeElement;
        if(ae && ae.tagName === 'IFRAME') return;
        if(document.hasFocus && document.hasFocus()) return;
        onLeave();
      }, 30);
    });
    window.addEventListener('focus', onReturn);
  }
  function showLeaveToast(ms){
    var d = cur ? cur.data : null;
    var n = d ? d.leaves : (+sessionStorage.getItem('kosa_leave_total')||1);
    var el = document.getElementById('kosa-leave-toast');
    if(!el){
      el = document.createElement('div'); el.id='kosa-leave-toast';
      el.style.cssText='position:fixed;left:50%;top:calc(env(safe-area-inset-top,0px) + 14px);transform:translateX(-50%) translateY(-20px);'+
        'z-index:99999;background:#2a0f05;border:1.5px solid #ff7744;border-radius:10px;padding:10px 16px;'+
        'color:#ffcc99;font:bold 13px/1.5 "Noto Sans KR",sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.5);'+
        'max-width:420px;text-align:center;opacity:0;transition:opacity .3s,transform .3s;pointer-events:none;word-break:keep-all';
      (document.body||document.documentElement).appendChild(el);
    }
    var sec = Math.round(ms/1000);
    el.innerHTML = '⚠️ 화면 이탈 감지 — ' + (sec>=60 ? Math.floor(sec/60)+'분 '+(sec%60)+'초' : sec+'초') +
      ' 동안 자리를 비웠어요 (누적 '+n+'회)<br><span style="font-weight:400;font-size:11.5px;color:#ffaa88">이탈 기록은 선생님께 함께 전달됩니다. 탐구에 집중해 주세요!</span>';
    el.style.opacity='1'; el.style.transform='translateX(-50%) translateY(0)';
    clearTimeout(el._t);
    el._t = setTimeout(function(){ el.style.opacity='0'; el.style.transform='translateX(-50%) translateY(-20px)'; }, 4500);
  }

  /* ────────────────────────────────────────────────
     ⑤ 성찰로그 전송 (로그인 정보 자동 사용)
     ──────────────────────────────────────────────── */
  KOSA.sendReflection = function(module, method, extra){
    var u = KOSA.getUser() || {sid:'', name:''};
    // 응답을 읽을 수 없는(no-cors) 전송이므로 기다릴 이유가 없음 → sendBeacon으로 즉시 완료 (Apps Script 콜드스타트 3~8초 대기 제거)
    return KOSA.send({kind:'reflect', module:module, sid:u.sid, sname:u.name, method:method||'', extra:extra||''}, true);
  };

  // 접속 기록 (index.html 로그인 시)
  KOSA.sendLogin = function(){
    var u = KOSA.getUser(); if(!u) return Promise.resolve(false);
    return KOSA.send({kind:'login', sid:u.sid, sname:u.name, ua:navigator.userAgent.slice(0,120),
                      screen:window.innerWidth+'x'+window.innerHeight});
  };

  // 기지 전력(%) — 세 모듈의 저장된 진행으로 계산 (허브·정사영 미션이 같은 값을 씀)
  KOSA.powerPct = function(){
    // 스토리: 전력은 태양광 패널을 바로잡아야(정사영 미션탐구 완료) 복구된다.
    // 그 전까지는 시작 23%에서 남은 시간(72h 카운트다운)에 비례해 줄어든다 (최소 5%).
    var u = KOSA.getUser() || {sid:'anon'}; function J(k){ try{ return JSON.parse(localStorage.getItem(k)); }catch(e){ return null; } }
    var jy = J('kosa_jsy_stage_'+u.sid); if (jy && jy[2]) return 100;
    var left = KOSA.collapseLeftSec ? KOSA.collapseLeftSec() : 72*3600;
    return Math.max(5, Math.round(23 * left / (72*3600)));
  };


  // 완료한 차시 수 (1차시 직선⊥평면 = 설계 승인, 2차시 삼수선 = 미션3, 3차시 정사영 = 미션탐구)
  KOSA.lessonsDone = function(){
    var u = KOSA.getUser() || {sid:'anon'}, sid = u.sid, n = 0;
    function J(k){ try{ return JSON.parse(localStorage.getItem(k)); }catch(e){ return null; } }
    var pp = J('kosa_pp_prog_'+sid); if(pp && pp.approved) n++;
    var ss = J('kosa_ss_prog_'+sid); if(ss && ss.cleared && ss.cleared[2]) n++;
    var jy = J('kosa_jsy_stage_'+sid); if(jy && jy[2]) n++;
    return n;
  };
  // 붕괴까지 남은 초: 72h 에서 완료 차시마다 24h 씩 줄고, 이번 접속 동안 실시간으로 조금씩 더 줄어듦 (최대 24h)
  KOSA.collapseLeftSec = function(){
    var t0 = 0; try{ t0 = +sessionStorage.getItem('kosa_t0') || 0; if(!t0){ t0 = Date.now(); sessionStorage.setItem('kosa_t0', t0); } }catch(e){ t0 = Date.now(); }
    var elapsed = Math.min(Date.now()-t0, 24*3600e3);
    return Math.max(0, Math.round(((3 - KOSA.lessonsDone())*24*3600e3 - elapsed)/1000));
  };
  // 시트 연결 테스트 (교사용): '접속기록' 시트에 (테스트) 행이 생기면 연결 정상
  KOSA.testSheet = function(){
    var u = KOSA.getUser() || {sid:'0000', name:'테스트'};
    return KOSA.send({kind:'login', sid:u.sid, sname:'(연결테스트) '+u.name, ua:'test', screen:window.innerWidth+'x'+window.innerHeight});
  };
  // 화면 우상단 등에 붙일 '로그인 표시' 텍스트
  KOSA.userLabel = function(){ var u = KOSA.getUser(); return u ? (u.sid+' '+u.name) : ''; };
})();
