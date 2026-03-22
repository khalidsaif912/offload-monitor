/**
 * eid-overlay.js
 * ملف مستقل - أضفه فقط في مجلد docs/ دون المساس بأي ملف آخر
 * 
 * كيفية الاستخدام:
 * أضف هذا السطر قبل </body> في أي صفحة HTML تريد:
 * <script src="/eid-overlay.js"></script>
 * 
 * أو إذا كانت الصفحة في مجلد فرعي (مثل /now/):
 * <script src="../eid-overlay.js"></script>
 */

(function () {
  'use strict';

  // ════════════════════════════════════════════
  //  ⚙️ الإعدادات — عدّلها حسب رغبتك
  // ════════════════════════════════════════════
  var CONFIG = {
    duration: 5000,          // مدة عرض الـ Overlay بالمللي ثانية (5 ثوان)
    fireworkCount: 80,       // عدد جزيئات الألعاب النارية لكل انفجار
    burstInterval: 700,      // الفاصل الزمني بين الانفجارات (مللي ثانية)
    messages: [              // العبارات المعروضة — تتناوب بالترتيب
      'عيدكم مبارك 🌙',
      'كل عام وأنتم بخير ✨',
      'تقبّل الله منا ومنكم 🤲',
      'عيد سعيد 🎉',
    ],
    msgInterval: 1200,       // الفاصل بين تغيير العبارات (مللي ثانية)
    // التحكم في تفعيل الـ Overlay:
    // 'always'   → يظهر في كل زيارة
    // 'once'     → يظهر مرة واحدة فقط (يحفظ في localStorage)
    // 'session'  → يظهر مرة لكل جلسة (يحفظ في sessionStorage)
    showMode: 'once',
    storageKey: 'eid_overlay_seen_2025',
  };

  // ════════════════════════════════════════════
  //  🔍 تحقق هل يجب عرض الـ Overlay
  // ════════════════════════════════════════════
  function shouldShow() {
    if (CONFIG.showMode === 'always') return true;
    var store = CONFIG.showMode === 'session' ? sessionStorage : localStorage;
    return !store.getItem(CONFIG.storageKey);
  }

  function markSeen() {
    if (CONFIG.showMode === 'always') return;
    var store = CONFIG.showMode === 'session' ? sessionStorage : localStorage;
    store.setItem(CONFIG.storageKey, '1');
  }

  if (!shouldShow()) return;

  // ════════════════════════════════════════════
  //  💅 CSS
  // ════════════════════════════════════════════
  var style = document.createElement('style');
  style.textContent = [
    '#eid-overlay{',
      'position:fixed;inset:0;z-index:999999;',
      'display:flex;flex-direction:column;align-items:center;justify-content:center;',
      'background:radial-gradient(ellipse at 50% 40%,#0a0a2e 0%,#000010 100%);',
      'animation:eidFadeIn .4s ease forwards;',
      'cursor:pointer;',
      'overflow:hidden;',
    '}',
    '@keyframes eidFadeIn{from{opacity:0}to{opacity:1}}',
    '@keyframes eidFadeOut{from{opacity:1}to{opacity:0;pointer-events:none}}',

    /* نجوم الخلفية */
    '#eid-stars{',
      'position:absolute;inset:0;pointer-events:none;',
    '}',

    /* الهلال */
    '#eid-moon{',
      'font-size:clamp(60px,14vw,110px);',
      'line-height:1;margin-bottom:8px;',
      'animation:moonFloat 3s ease-in-out infinite;',
      'filter:drop-shadow(0 0 30px rgba(255,200,50,.6));',
      'position:relative;z-index:2;',
    '}',
    '@keyframes moonFloat{',
      '0%,100%{transform:translateY(0) rotate(-5deg)}',
      '50%{transform:translateY(-12px) rotate(5deg)}',
    '}',

    /* نص العبارات */
    '#eid-msg{',
      'font-family:"Amiri","Scheherazade New","Traditional Arabic",serif;',
      'font-size:clamp(26px,7vw,54px);',
      'font-weight:700;',
      'color:#fff;',
      'text-align:center;',
      'padding:0 20px;',
      'text-shadow:0 0 40px rgba(255,200,50,.8), 0 2px 8px rgba(0,0,0,.6);',
      'letter-spacing:2px;',
      'position:relative;z-index:2;',
      'min-height:1.3em;',
      'transition:opacity .4s ease;',
    '}',

    /* سطر تلميح الإغلاق */
    '#eid-hint{',
      'margin-top:28px;',
      'font-family:system-ui,sans-serif;',
      'font-size:13px;color:rgba(255,255,255,.45);',
      'position:relative;z-index:2;',
      'letter-spacing:.5px;',
    '}',

    /* شريط العد التنازلي */
    '#eid-bar-wrap{',
      'position:absolute;bottom:0;left:0;right:0;height:4px;',
      'background:rgba(255,255,255,.1);',
    '}',
    '#eid-bar{',
      'height:100%;width:100%;',
      'background:linear-gradient(to right,#f59e0b,#fbbf24,#fde68a);',
      'transform-origin:left;',
      'transition:transform linear;',
    '}',

    /* جزيئات الألعاب النارية */
    '.eid-spark{',
      'position:absolute;',
      'width:5px;height:5px;border-radius:50%;',
      'pointer-events:none;',
      'will-change:transform,opacity;',
    '}',
  ].join('');
  document.head.appendChild(style);

  // ════════════════════════════════════════════
  //  🏗️ بناء الـ DOM
  // ════════════════════════════════════════════
  var overlay = document.createElement('div');
  overlay.id = 'eid-overlay';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', 'تهنئة العيد');

  // نجوم
  var starsEl = document.createElement('canvas');
  starsEl.id = 'eid-stars';
  overlay.appendChild(starsEl);

  // هلال
  var moon = document.createElement('div');
  moon.id = 'eid-moon';
  moon.textContent = '🌙';
  overlay.appendChild(moon);

  // عبارة
  var msgEl = document.createElement('div');
  msgEl.id = 'eid-msg';
  msgEl.textContent = CONFIG.messages[0];
  overlay.appendChild(msgEl);

  // تلميح
  var hint = document.createElement('div');
  hint.id = 'eid-hint';
  hint.textContent = 'اضغط في أي مكان للمتابعة';
  overlay.appendChild(hint);

  // شريط التقدم
  var barWrap = document.createElement('div');
  barWrap.id = 'eid-bar-wrap';
  var bar = document.createElement('div');
  bar.id = 'eid-bar';
  barWrap.appendChild(bar);
  overlay.appendChild(barWrap);

  document.body.appendChild(overlay);

  // ════════════════════════════════════════════
  //  ⭐ رسم النجوم على Canvas
  // ════════════════════════════════════════════
  function drawStars() {
    var W = starsEl.width  = window.innerWidth;
    var H = starsEl.height = window.innerHeight;
    var ctx = starsEl.getContext('2d');
    for (var i = 0; i < 160; i++) {
      var x = Math.random() * W;
      var y = Math.random() * H;
      var r = Math.random() * 1.6 + 0.2;
      var alpha = Math.random() * 0.7 + 0.2;
      ctx.beginPath();
      ctx.arc(x, y, r, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(255,255,220,' + alpha + ')';
      ctx.fill();
    }
  }
  drawStars();
  window.addEventListener('resize', drawStars);

  // ════════════════════════════════════════════
  //  🎆 الألعاب النارية
  // ════════════════════════════════════════════
  var COLORS = [
    '#ff6b6b','#ffa500','#ffd700','#00ffcc',
    '#00bfff','#ff69b4','#adff2f','#ff4500',
    '#da70d6','#ffffff','#ffe066','#7fffd4',
  ];

  function rand(min, max) { return min + Math.random() * (max - min); }

  function burst(cx, cy) {
    var count = CONFIG.fireworkCount;
    for (var i = 0; i < count; i++) {
      (function (i) {
        var spark = document.createElement('div');
        spark.className = 'eid-spark';
        var color = COLORS[Math.floor(Math.random() * COLORS.length)];
        spark.style.cssText = [
          'left:' + cx + 'px',
          'top:' + cy + 'px',
          'background:' + color,
          'box-shadow:0 0 6px 2px ' + color,
        ].join(';');
        overlay.appendChild(spark);

        var angle = (i / count) * Math.PI * 2 + rand(-0.2, 0.2);
        var speed = rand(60, 220);
        var dx = Math.cos(angle) * speed;
        var dy = Math.sin(angle) * speed;
        var dur = rand(600, 1100);
        var start = performance.now();

        function animate(now) {
          var elapsed = now - start;
          var progress = Math.min(elapsed / dur, 1);
          var ease = 1 - Math.pow(1 - progress, 3);
          var gravity = progress * progress * 80;
          var x = cx + dx * ease;
          var y = cy + dy * ease + gravity;
          var opacity = 1 - progress;
          var scale = 1 - progress * 0.7;
          spark.style.transform = 'translate(' + (x - cx) + 'px,' + (y - cy) + 'px) scale(' + scale + ')';
          spark.style.opacity = opacity;
          if (progress < 1) {
            requestAnimationFrame(animate);
          } else {
            if (spark.parentNode) spark.parentNode.removeChild(spark);
          }
        }
        requestAnimationFrame(animate);
      })(i);
    }
  }

  function randomBurst() {
    var x = rand(window.innerWidth * 0.1, window.innerWidth * 0.9);
    var y = rand(window.innerHeight * 0.1, window.innerHeight * 0.65);
    burst(x, y);
  }

  // ════════════════════════════════════════════
  //  🔄 تناوب العبارات
  // ════════════════════════════════════════════
  var msgIdx = 0;
  var msgTimer = setInterval(function () {
    msgIdx = (msgIdx + 1) % CONFIG.messages.length;
    msgEl.style.opacity = '0';
    setTimeout(function () {
      msgEl.textContent = CONFIG.messages[msgIdx];
      msgEl.style.opacity = '1';
    }, 400);
  }, CONFIG.msgInterval);

  // ════════════════════════════════════════════
  //  ⏱️ شريط التقدم + الإغلاق التلقائي
  // ════════════════════════════════════════════
  var startTime = performance.now();

  function updateBar(now) {
    var elapsed = now - startTime;
    var ratio = Math.max(0, 1 - elapsed / CONFIG.duration);
    bar.style.transform = 'scaleX(' + ratio + ')';
    if (ratio > 0) requestAnimationFrame(updateBar);
  }
  requestAnimationFrame(updateBar);

  // انفجارات تلقائية
  randomBurst();
  var burstTimer = setInterval(randomBurst, CONFIG.burstInterval);

  // الإغلاق
  function closeOverlay() {
    clearInterval(msgTimer);
    clearInterval(burstTimer);
    window.removeEventListener('resize', drawStars);
    overlay.style.animation = 'eidFadeOut .5s ease forwards';
    markSeen();
    setTimeout(function () {
      if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
    }, 520);
  }

  var autoClose = setTimeout(closeOverlay, CONFIG.duration);

  overlay.addEventListener('click', function () {
    clearTimeout(autoClose);
    closeOverlay();
  });

  overlay.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'Escape') {
      clearTimeout(autoClose);
      closeOverlay();
    }
  });

})();
