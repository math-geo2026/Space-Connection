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
  KOSA.SHEET_URL = 'https://script.google.com/macros/s/AKfycbyZbPrunr9E_eV1w_ihP2rYL5uOpkbpKn_AvgKTCr2wDm0XnhH5DaxxiDgXBloTR9wnxg/exec';
  KOSA.DASH_URL = KOSA.SHEET_URL; // 교사용 관제소(실시간 웹 대시보드) — 학생 데이터 기록과 같은 배포 주소로 통일

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

  // (예전에 여기 있던 "페이지 전체를 훑어서 스크롤 영역마다 overscroll-behavior를 자동으로 붙여주는"
  // 안전망 코드는 제거했습니다 — 요소가 추가/삭제될 때마다 문서 전체를 getComputedStyle로 다시
  // 훑는 방식이라 무거웠고, 지금은 모든 스크롤 패널에 CSS로 직접 overscroll-behavior:contain을
  // 넣어뒀기 때문에 더는 필요하지 않습니다.)

  // Apple Pencil 탭 보정: 펜 탭에 click 이 안 오는 경우 pointerup 으로 보완 (중복 방지)
  document.addEventListener('click', function(e){ var b=e.target.closest&&e.target.closest('button,.mt,.kchip,.pbtn,.adv-btn,.goal-opt'); if(b) b.__lastClick=Date.now(); }, true);
  document.addEventListener('pointerup', function(e){
    if(e.pointerType!=='pen') return;
    var b=e.target.closest&&e.target.closest('button,.mt,.kchip,.pbtn,.adv-btn,.goal-opt'); if(!b||b.disabled) return;
    setTimeout(function(){ if(!b.__lastClick || Date.now()-b.__lastClick>250){ b.__lastClick=Date.now(); b.click(); } }, 120);
  }, true);
  // 안쪽 패널만 스크롤 (scrollIntoView 가 문서 전체를 밀지 않도록)
  /* ────────────────────────────────────────────────
     ⑥ 다음 행동 안내 (반짝이는 테두리 + 손가락 포인터)
     사용법: KOSA.guide(요소또는선택자, { text:'여기를 눌러 시작하세요', key:'m1-dial' })
       - text 생략 가능(테두리만 반짝임)
       - key 를 주면 localStorage에 "이미 본 힌트"로 기록해 재접속 시 다시 안 뜸
       - 요소를 클릭/터치하거나 KOSA.clearGuide(el) 호출 시 자동 종료
     ──────────────────────────────────────────────── */
  if(!document.getElementById('kosa-hint-style')){
    var hs=document.createElement('style'); hs.id='kosa-hint-style';
    hs.textContent =
      '@keyframes kosaGlow{0%,100%{box-shadow:0 0 0 2px #ffee00,0 0 6px 2px rgba(255,238,0,.5)}50%{box-shadow:0 0 0 3px #fff35c,0 0 18px 6px rgba(255,238,0,.85)}}'+
      '.kosa-hint-glow{animation:kosaGlow 1.1s ease-in-out infinite;border-radius:8px;position:relative;z-index:2}'+
      '@keyframes kosaHandBob{0%,100%{transform:translateY(0)}50%{transform:translateY(6px)}}'+
      '@keyframes kosaHandBobUp{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}'+
      '.kosa-hint-hand{position:absolute;font-size:26px;line-height:1;pointer-events:none;z-index:9999;'+
      'right:2px;top:-28px;animation:kosaHandBob .85s ease-in-out infinite;filter:drop-shadow(0 2px 3px rgba(0,0,0,.5))}'+
      '.kosa-hint-hand.up{animation-name:kosaHandBobUp}'+
      '.kosa-hint-bubble{position:absolute;left:50%;bottom:100%;transform:translateX(-50%);margin-bottom:32px;'+
      'background:#1a1300;border:1.5px solid #ffee00;color:#ffee88;font:700 12px/1.4 "Noto Sans KR",sans-serif;'+
      'padding:6px 10px;border-radius:7px;white-space:nowrap;pointer-events:none;z-index:9999;'+
      'box-shadow:0 4px 14px rgba(0,0,0,.5)}'+
      '.kosa-hint-bubble:after{content:"";position:absolute;top:100%;left:50%;transform:translateX(-50%);'+
      'border:6px solid transparent;border-top-color:#ffee00}';
    document.head.appendChild(hs);
  }
  KOSA.clearGuide = function(el){
    el = (typeof el==='string') ? document.querySelector(el) : el;
    if(!el || !el.__kosaHint) return;
    el.classList.remove('kosa-hint-glow');
    if(el.__kosaHint.hand) el.__kosaHint.hand.remove();
    if(el.__kosaHint.bubble) el.__kosaHint.bubble.remove();
    if(el.__kosaHint.off) el.__kosaHint.off();
    if(el.__kosaHint.posRestore!=null) el.style.position = el.__kosaHint.posRestore;
    el.__kosaHint = null;
  };
  KOSA.guide = function(el, opts){
    opts = opts || {};
    el = (typeof el==='string') ? document.querySelector(el) : el;
    if(!el) return;
    if(opts.key){
      try{ if(localStorage.getItem('kosa_hint_'+opts.key)==='1') return; }catch(e){}
    }
    KOSA.clearGuide(el);
    var posRestore = null;
    var cs = getComputedStyle(el);
    if(cs.position==='static'){ posRestore=''; el.style.position='relative'; }
    el.classList.add('kosa-hint-glow');
    var hand=null, bubble=null;
    if(opts.hand !== false){
      hand=document.createElement('span'); hand.className='kosa-hint-hand'; el.appendChild(hand);
      // 화면에서 이 요소의 위치를 보고 손가락 방향·위치를 자동으로 정함 (버튼이 잘리거나 스크롤이 새로 생기지 않게)
      var rect = el.getBoundingClientRect();
      var vw = window.innerWidth, vh = window.innerHeight;
      var spaceAbove = rect.top, spaceBelow = vh - rect.bottom;
      var vertical = (spaceAbove >= 34 && spaceAbove >= spaceBelow) ? 'above' : (spaceBelow >= 34 ? 'below' : 'above');
      var spaceRight = vw - rect.right, spaceLeft = rect.left;
      var horiz = (spaceRight < 24 && spaceLeft > spaceRight) ? 'left' : 'right';
      hand.textContent = vertical==='above' ? '👇' : '👆';
      if(vertical==='below') hand.classList.add('up');
      hand.style.top = vertical==='above' ? '-28px' : '';
      hand.style.bottom = vertical==='below' ? '-28px' : '';
      hand.style.right = horiz==='right' ? '2px' : '';
      hand.style.left = horiz==='left' ? '2px' : '';
    }
    if(opts.text){
      bubble=document.createElement('span'); bubble.className='kosa-hint-bubble'; bubble.textContent=opts.text; el.appendChild(bubble);
    }
    function dismiss(){
      if(opts.key){ try{ localStorage.setItem('kosa_hint_'+opts.key,'1'); }catch(e){} }
      KOSA.clearGuide(el);
    }
    el.addEventListener('pointerdown', dismiss, {once:true});
    el.addEventListener('input', dismiss, {once:true});
    // 안내가 가리키는 요소가 아니라 근처의 다른 입력(드롭다운 선택 등)으로 답을 완성하는 경우도 있어서,
    // 문서 전체의 변경(change) 이벤트에도 반응해 안내가 화면에 계속 남아있지 않게 함
    document.addEventListener('change', dismiss, {once:true, capture:true});
    var timer = setTimeout(dismiss, opts.timeout || 8000);   // 무엇으로도 안 닫히는 경우를 대비한 최종 안전장치
    el.__kosaHint = { hand:hand, bubble:bubble, posRestore:posRestore, off:function(){ clearTimeout(timer); document.removeEventListener('change', dismiss, {capture:true}); } };
  };

  /* ────────────────────────────────────────────────
     ⑦ 메모장 (Apple Pencil / 손가락 필기, 접기·펼치기 지원)
     사용법: KOSA.mountScratchpad('#어딘가', {
       height:200, label:'✏️ 메모장',
       collapsed:true,   // 기본 접힘 여부 (기본값 true)
       prepend:true,     // 컨테이너 맨 앞에 넣을지 (기본 false = 맨 뒤)
       overlay:true       // true면 펼쳤을 때 뒤 내용 위에 떠서 겹쳐 보임(레이아웃 안 밀림)
     })
       - 펜/손가락 모두 지원 (Pointer Events, 필압 반영)
       - 접었다 펴도 그린 내용은 유지됨 (캔버스 자체는 항상 DOM에 남아있음)
       - 같은 요소에 두 번 부르면 무시(중복 방지)
     ──────────────────────────────────────────────── */
  if(!document.getElementById('kosa-pad-style')){
    var ps=document.createElement('style'); ps.id='kosa-pad-style';
    ps.textContent =
      '.kosa-pad{margin-top:10px;border:1.5px dashed #3a5a8a;border-radius:9px;background:#0a1424;overflow:visible;position:relative}'+
      '.kosa-pad.sticky{position:sticky;top:0;z-index:70;box-shadow:0 6px 14px rgba(0,0,0,.35)}'+
      '.kosa-pad-tab{display:block;width:100%;text-align:left;border:none;background:#0e1c33;color:#8fb4e0;font:700 12px "Noto Sans KR",sans-serif;padding:7px 10px;border-radius:9px;cursor:pointer;font-family:inherit}'+
      '.kosa-pad-tab:hover{background:#132449}'+
      '.kosa-pad-body{border-top:1px solid #22375c;border-radius:0 0 9px 9px;overflow:hidden;background:#0a1424}'+
      '.kosa-pad.overlay .kosa-pad-body{position:absolute;top:100%;left:0;right:0;z-index:80;box-shadow:0 10px 26px rgba(0,0,0,.5);border:1.5px solid #3a5a8a;border-top:1px solid #22375c;border-radius:0 0 9px 9px}'+
      '.kosa-pad-bar{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:6px 8px;background:#0e1c33;border-bottom:1px solid #22375c}'+
      '.kosa-pad-bar .lbl{font:700 12px "Noto Sans KR",sans-serif;color:#8fb4e0;margin-right:auto}'+
      '.kosa-pad-bar button{border:1.5px solid #2e4d78;background:#122140;color:#cfe0ff;border-radius:6px;padding:4px 8px;font-size:11px;cursor:pointer;font-family:inherit}'+
      '.kosa-pad-bar button.on{border-color:#ffee00;background:#2a2400;color:#ffee88}'+
      '.kosa-pad-bar .sw{width:20px;height:20px;border-radius:50%;border:2px solid #3a5a8a;padding:0;cursor:pointer}'+
      '.kosa-pad-bar .sw.on{border-color:#fff}'+
      '.kosa-pad-canvas-wrap{position:relative;touch-action:none}'+
      '.kosa-pad-canvas-wrap canvas{display:block;width:100%;background:#fff}';
    document.head.appendChild(ps);
  }
  KOSA.mountScratchpad = function(container, opts){
    opts = opts || {};
    container = (typeof container==='string') ? document.querySelector(container) : container;
    if(!container || container.__kosaPad) return;
    container.__kosaPad = true;
    var H = opts.height || 200;
    var overlay = !!opts.overlay;
    var open = (opts.collapsed === false);

    var pad = document.createElement('div'); pad.className='kosa-pad'+(overlay?' overlay':'');
    if(overlay){ var cs=getComputedStyle(container); if(cs.position==='static') container.style.position='relative'; }
    var tab = document.createElement('button'); tab.type='button'; tab.className='kosa-pad-tab';
    var body = document.createElement('div'); body.className='kosa-pad-body'; body.style.display = open ? 'block' : 'none';
    var bar = document.createElement('div'); bar.className='kosa-pad-bar';
    var colors = ['#111111','#e02020','#1a66ff','#0a9a4a'];
    var curColor = colors[0], curSize = 2.4, erasing = false;
    bar.innerHTML =
      '<span class="lbl">'+(opts.label||'✏️ 메모장')+'</span>'+
      colors.map(function(c,i){ return '<button type="button" class="sw'+(i===0?' on':'')+'" data-c="'+c+'" style="background:'+c+'"></button>'; }).join('')+
      '<button type="button" data-sz="1.4">가늘게</button><button type="button" data-sz="2.4" class="on">보통</button><button type="button" data-sz="5">굵게</button>'+
      '<button type="button" data-er="1">🧹 지우개</button><button type="button" data-clr="1">전체 지우기</button>';
    var cwrap = document.createElement('div'); cwrap.className='kosa-pad-canvas-wrap';
    var canvas = document.createElement('canvas'); cwrap.appendChild(canvas);
    body.appendChild(bar); body.appendChild(cwrap);
    pad.appendChild(tab); pad.appendChild(body);
    if(opts.prepend) container.insertBefore(pad, container.firstChild); else container.appendChild(pad);

    function setTabText(){ tab.textContent = (opts.label||'✏️ 메모장') + (open ? ' 접기 ▾' : ' 펼치기 ▸'); }
    setTabText();

    var ctx = canvas.getContext('2d');
    function resize(){
      var w = cwrap.clientWidth || 300, dpr = Math.min(window.devicePixelRatio||1, 2);
      canvas.width = Math.round(w*dpr); canvas.height = Math.round(H*dpr);
      canvas.style.height = H+'px';
      ctx.scale(dpr,dpr); ctx.lineCap='round'; ctx.lineJoin='round';
    }
    function resizePreserve(){
      var img = null; try{ img = canvas.toDataURL(); }catch(e){}
      var oldDpr = Math.min(window.devicePixelRatio||1, 2);
      resize();
      if(img){ var im=new Image(); im.onload=function(){ ctx.drawImage(im,0,0,canvas.width/((window.devicePixelRatio||1)),canvas.height/((window.devicePixelRatio||1))); }; im.src=img; }
    }
    if(open) resize();

    tab.addEventListener('click', function(){
      open = !open;
      body.style.display = open ? 'block' : 'none';
      setTabText();
      if(open) resizePreserve();   // 펼칠 때 캔버스 크기를 다시 잡되, 그려둔 내용은 유지
    });

    var ro; try{ ro = new ResizeObserver(function(){ if(open) resizePreserve(); }); ro.observe(cwrap); }catch(e){}

    bar.addEventListener('click', function(e){
      var b = e.target.closest('button'); if(!b) return;
      if(b.dataset.c){ curColor=b.dataset.c; erasing=false; bar.querySelectorAll('.sw').forEach(function(x){x.classList.toggle('on',x===b);}); bar.querySelector('[data-er]').classList.remove('on'); }
      else if(b.dataset.sz){ curSize=+b.dataset.sz; bar.querySelectorAll('[data-sz]').forEach(function(x){x.classList.toggle('on',x===b);}); }
      else if(b.dataset.er){ erasing=!erasing; b.classList.toggle('on',erasing); }
      else if(b.dataset.clr){ ctx.clearRect(0,0,canvas.width,canvas.height); }
    });

    var drawing=false, lastX=0, lastY=0;
    function pos(e){ var r=canvas.getBoundingClientRect(); return { x:e.clientX-r.left, y:e.clientY-r.top }; }
    canvas.addEventListener('pointerdown', function(e){
      drawing=true; canvas.setPointerCapture(e.pointerId);
      var p=pos(e); lastX=p.x; lastY=p.y;
      ctx.beginPath(); ctx.arc(p.x,p.y, (erasing?curSize*4:curSize)/2, 0, Math.PI*2);
      ctx.fillStyle = erasing?'#fff':curColor; ctx.fill();
    });
    canvas.addEventListener('pointermove', function(e){
      if(!drawing) return;
      var p=pos(e);
      var pr = (e.pointerType==='pen' && e.pressure>0) ? e.pressure : 1;
      ctx.strokeStyle = erasing?'#fff':curColor;
      ctx.lineWidth = (erasing?curSize*4:curSize) * (0.6+0.8*pr);
      ctx.beginPath(); ctx.moveTo(lastX,lastY); ctx.lineTo(p.x,p.y); ctx.stroke();
      lastX=p.x; lastY=p.y;
    });
    ['pointerup','pointercancel','pointerleave'].forEach(function(ev){
      canvas.addEventListener(ev, function(){ drawing=false; });
    });
  };

  /* ────────────────────────────────────────────────
     ⑦-2 화면 고정 메모장 묶음 (스크롤·본문 길이와 무관하게 항상 같은 자리)
     사용법: var dock = KOSA.mountDockedScratchpads('#probBox', 3, {height:150, label:'✏️ 메모장', corner:'br'});
       - outerBox는 반드시 position:fixed/absolute로 화면을 덮는, 그 자체는 스크롤되지 않는 바깥 박스여야 함
         (#probBox, #hwBox 처럼) — 그 박스의 한쪽 구석에 항상 고정으로 붙고, 안쪽 스크롤과 무관하게 안 움직임.
       - count가 2 이상이면 문제별로 서로 다른 메모장을 만들어두고, dock.show(i)로 지금 보이는 것만 바꿔치기
         (문제 1·2·3 필기 내용이 서로 안 섞임). count가 1이면 화면 전체에 하나만.
       - 반환값의 dock.show(i)를 탭 전환 함수(예: hwTab, hw2Tab) 안에서 호출해 연동한다.
     ──────────────────────────────────────────────── */
  if(!document.getElementById('kosa-dock-style')){
    var dks=document.createElement('style'); dks.id='kosa-dock-style';
    dks.textContent =
      '.kosa-pad-dock{position:absolute;right:14px;bottom:14px;z-index:200;width:min(340px,88vw)}'+
      '.kosa-pad-dock.tl{right:auto;bottom:auto;left:14px;top:92px}'+
      '.kosa-pad-dock.tr{bottom:auto;top:92px}'+
      '.kosa-pad-dock.bl{right:auto;left:14px}'+
      '.kosa-pad-dock .kosa-pad{margin-top:0;box-shadow:0 8px 22px rgba(0,0,0,.5)}';
    document.head.appendChild(dks);
  }
  KOSA.mountDockedScratchpads = function(outerBox, count, opts){
    opts = opts || {};
    outerBox = (typeof outerBox==='string') ? document.querySelector(outerBox) : outerBox;
    if(!outerBox || outerBox.__kosaDock) return outerBox && outerBox.__kosaDock;
    var docks = [];
    for(var i=0;i<count;i++){
      var d=document.createElement('div');
      d.className='kosa-pad-dock'+(opts.corner?(' '+opts.corner):'');
      d.style.display = i===0 ? '' : 'none';
      outerBox.appendChild(d);
      KOSA.mountScratchpad(d, { height:opts.height||150, label:opts.label||'✏️ 메모장', collapsed:true, overlay:false });
      docks.push(d);
    }
    var api = { show:function(i){ docks.forEach(function(d,idx){ d.style.display = idx===i ? '' : 'none'; }); } };
    outerBox.__kosaDock = api;
    return api;
  };
  // ⑧ 공용 안내 모달 (브라우저 기본 alert() 대체용) — KOSA.modal('메시지', {tone, title})
  if(!document.getElementById('kosa-modal-style')){
    var ms=document.createElement('style'); ms.id='kosa-modal-style';
    ms.textContent =
      '.kosa-modal-ov{position:fixed;inset:0;background:rgba(4,8,16,.72);z-index:99997;display:flex;align-items:center;justify-content:center;padding:20px;'+
      'opacity:0;transition:opacity .18s ease}'+
      '.kosa-modal-ov.on{opacity:1}'+
      '.kosa-modal-card{background:#0e1c33;border:1.5px solid #3a5a8a;border-radius:14px;max-width:380px;width:100%;padding:22px 20px 18px;'+
      'box-shadow:0 14px 40px rgba(0,0,0,.5);text-align:center;transform:translateY(8px);transition:transform .18s ease}'+
      '.kosa-modal-ov.on .kosa-modal-card{transform:translateY(0)}'+
      '.kosa-modal-card.warn{border-color:#aa6622}.kosa-modal-card.success{border-color:#2f7d4f}'+
      '.kosa-modal-icon{font-size:30px;margin-bottom:8px}'+
      '.kosa-modal-title{font:900 15px "Noto Sans KR",sans-serif;color:#eaf2ff;margin-bottom:6px}'+
      '.kosa-modal-msg{font:400 13.5px/1.6 "Noto Sans KR",sans-serif;color:#cfe0ff;margin-bottom:16px;white-space:pre-line}'+
      '.kosa-modal-ok{border:none;border-radius:8px;padding:9px 26px;font:700 13px "Noto Sans KR",sans-serif;cursor:pointer;background:#5599ff;color:#06111f}'+
      '.kosa-modal-card.warn .kosa-modal-ok{background:#ffcc55}'+
      '.kosa-modal-card.success .kosa-modal-ok{background:#44ffaa}';
    document.head.appendChild(ms);
  }
  KOSA.modal = function(msg, opts){
    opts = opts || {};
    var tone = opts.tone || 'info';
    var icon = opts.icon || (tone==='warn'?'⚠️':tone==='success'?'🎉':'ℹ️');
    var ov = document.createElement('div'); ov.className='kosa-modal-ov';
    var card = document.createElement('div'); card.className='kosa-modal-card '+tone;
    card.innerHTML =
      '<div class="kosa-modal-icon">'+icon+'</div>'+
      (opts.title ? '<div class="kosa-modal-title">'+opts.title+'</div>' : '')+
      '<div class="kosa-modal-msg"></div>'+
      '<button type="button" class="kosa-modal-ok">확인</button>';
    card.querySelector('.kosa-modal-msg').textContent = msg;   // 사용자 문구를 그대로 텍스트로만 삽입(HTML 이스케이프)
    ov.appendChild(card); document.body.appendChild(ov);
    requestAnimationFrame(function(){ ov.classList.add('on'); });
    function close(){ ov.classList.remove('on'); setTimeout(function(){ ov.remove(); }, 200); }
    card.querySelector('.kosa-modal-ok').addEventListener('click', close);
    ov.addEventListener('click', function(e){ if(e.target===ov) close(); });
    return close;
  };

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
  // 심사용 모드: URL에 ?judge=1 이 한 번이라도 붙으면 이 브라우저 세션 내내(다른 페이지로 이동해도) 유지됨.
  // 각 모듈의 '잠금' 체크 코드에서 이 값을 확인해, 순서를 건너뛰고 모든 단계를 열어볼 수 있게 한다.
  KOSA.isJudge = function(){
    try{
      if(/[?&]judge=1(&|$)/.test(location.search)){ sessionStorage.setItem('kosa_judge','1'); }
      if(sessionStorage.getItem('kosa_judge')==='1') return true;
    }catch(e){}
    // 세션 신호가 없어도(브라우저를 새로 열었거나 hub.html을 직접 열어 들어온 경우 등),
    // 로그인된 학번이 심사용 계정(99999)이면 항상 잠금을 우회한다 — 로그인 정보(localStorage)는
    // 세션이 끝나도 남아있으니, 이 계정으로 로그인돼 있는 한 항상 동일하게 동작해야 자연스럽다.
    try{
      var u = KOSA.getUser && KOSA.getUser();
      if(u && String(u.sid)==='99999') return true;
    }catch(e){}
    return false;
  };
  KOSA.requireLogin = function(){
    if(window.self !== window.top) return true;
    if(KOSA.getUser()) return true;
    if(KOSA.isJudge()){ KOSA.setUser('99999','심사용'); return true; }   // 심사 모드면 로그인 화면 없이 바로 통과
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
  var readyAt = Date.now();
  var LEAVE_MIN_MS = 1500;   // 이 시간 이상 벗어나야 '진짜 이탈'로 집계 (안내 토스트 기준과 통일)
  var LEAVE_GRACE_MS = 1000; // 페이지가 막 열린 직후 이 시간 동안은 전환 잡음으로 보고 무시
  // alert()/confirm() 중에는 창이 흐려져도 이탈로 치지 않음
  ['alert','confirm','prompt'].forEach(function(fn){
    var orig = window[fn];
    if(typeof orig !== 'function') return;
    window[fn] = function(){ alertOpen = true; try{ return orig.apply(window, arguments); } finally { setTimeout(function(){ alertOpen=false; }, 50); } };
  });
  function onLeave(){
    if(leaving || alertOpen) return;
    if(Date.now() - readyAt < LEAVE_GRACE_MS) return;   // 페이지 전환 직후 오탐 방지
    leaving = true; leaveStart = Date.now();
    if(cur){ tick(); cur.segStart = null; }
  }
  function onReturn(){
    if(!leaving) return;
    leaving = false;
    var ms = Date.now()-leaveStart;
    if(ms < LEAVE_MIN_MS) { if(cur && !document.hidden) cur.segStart = Date.now(); return; }   // 너무 짧으면 '진짜 이탈'로 안 셈
    if(cur){
      cur.data.leaves++; cur.data.leaveMs += ms; save(cur.data);
      if(!document.hidden) cur.segStart = Date.now();
    }
    try{ sessionStorage.setItem('kosa_leave_total', String((+sessionStorage.getItem('kosa_leave_total')||0)+1)); }catch(e){}
    showLeaveToast(ms);
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
