(() => {
  'use strict';
  const canvas = document.querySelector('#canvas');
  const ctx = canvas.getContext('2d');
  const overlay = document.querySelector('#overlay');
  const action = document.querySelector('#action');
  const levelLabel = document.querySelector('#level');
  const effectsLabel = document.querySelector('#effects');
  const notice = document.querySelector('#notice');
  const touch = document.querySelector('#touch');
  const W = 960, H = 640, R = 7, PADDLE_Y = 602;
  const names = { expand: '바 확장', multi: '공 추가', pierce: '벽돌 관통' };
  const symbols = { expand: '확', multi: '공', pierce: '관' };
  const animals = [{ face: '🐰', type: 'expand' }, { face: '🐱', type: 'multi' }, { face: '🐻', type: 'pierce' }];
  let level = 1, state = 'ready', bricks, balls, items, paddle, effects, broken, bounces, lost, secrets, keys = {}, particles = [], celebrationTime = 0, last = 0, noticeTimer = 0;

  function speed() { return 310 * (1 + (level - 1) * 0.005); }
  function resetStage() {
    bricks = [];
    const maxHits = Math.min(5, Math.ceil(level / 20));
    for (let row = 0; row < 7; row++) for (let col = 0; col < 9; col++) {
      bricks.push({ x: 14 + col * 104, y: 38 + row * 34, w: 100, h: 30, animal: animals[Math.floor(Math.random() * animals.length)], type: null, hits: 1 + Math.floor(Math.random() * maxHits), alive: true });
    }
    const positions = Array.from({ length: 63 }, (_, i) => i);
    for (let i = positions.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [positions[i], positions[j]] = [positions[j], positions[i]]; }
    animals.forEach((animal, i) => { const brick = bricks[positions[i]]; brick.animal = animal; brick.type = animal.type; });
    paddle = { x: 405, w: 150, h: 15 };
    balls = [{ x: 480, y: PADDLE_Y - R, vx: speed() * 0.55, vy: -speed() * Math.sqrt(1 - 0.55 ** 2), contacts: new Set() }];
    items = []; effects = { expand: 0, pierce: 0 }; broken = 0; bounces = 0; lost = false; secrets = {}; keys = {}; particles = [];
    levelLabel.innerHTML = `${String(level).padStart(2, '0')} <small>/ 100</small>`;
    updateEffects();
  }
  function showOverlay(kicker, message, hint, button, celebration = false) {
    document.querySelector('#kicker').textContent = kicker;
    document.querySelector('#message').textContent = message;
    document.querySelector('#hint').textContent = hint;
    action.innerHTML = `${button} <span>→</span>`;
    overlay.classList.toggle('celebration', celebration);
    overlay.hidden = false;
  }
  function startAction() {
    if (state === 'playing') return;
    if (state === 'lost' || state === 'won') { level = 1; resetStage(); }
    else if (state === 'between') { level++; resetStage(); }
    state = 'playing'; overlay.hidden = true; last = 0;
  }
  function alertUser(message) {
    notice.textContent = message; notice.classList.add('show'); clearTimeout(noticeTimer);
    noticeTimer = setTimeout(() => notice.classList.remove('show'), 2600);
  }
  function drop(type, x, y) { items.push({ type, x: Math.max(30, Math.min(W - 30, x)), y }); }
  function secret(key, type, x, y) {
    if (!secrets[key]) { secrets[key] = true; drop(type, x, y); alertUser('숨겨진 조건 달성! 아이템이 나타났어요'); }
  }
  function applyItem(type) {
    if (type === 'expand') { effects.expand = 10; paddle.w = 225; paddle.x = Math.min(paddle.x, W - paddle.w); }
    if (type === 'pierce') effects.pierce = 10;
    if (type === 'multi') {
      for (const ball of [...balls]) {
        if (balls.length >= 6) break;
        const angle = Math.atan2(ball.vy, ball.vx) + 0.55;
        balls.push({ x: ball.x, y: ball.y, vx: Math.cos(angle) * speed(), vy: Math.sin(angle) * speed(), contacts: new Set(ball.contacts) });
      }
    }
    alertUser(`${names[type]} 효과를 얻었어요!`);
  }
  function updateEffects() {
    const labels = [];
    if (effects.expand > 0) labels.push(`바 확장 ${Math.ceil(effects.expand)}초`);
    if (effects.pierce > 0) labels.push(`벽돌 관통 ${Math.ceil(effects.pierce)}초`);
    effectsLabel.textContent = labels.join(' · ') || '효과 없음';
  }
  function endStage() {
    if (level === 100) {
      state = 'won'; celebrationTime = 10;
      particles = Array.from({ length: 120 }, () => ({ x: Math.random() * W, y: Math.random() * H, vx: (Math.random() - .5) * 100, vy: 40 + Math.random() * 130, size: 4 + Math.random() * 8, color: ['#e23f39','#f4b541','#4a9eaa','#8c70bb','#f187a0'][Math.floor(Math.random()*5)] }));
      showOverlay('모든 단계를 완료했어요', '100단계 정복! 정말 대단해요!', '끝까지 해낸 당신을 축하합니다!', '다시 시작', true);
    } else { state = 'between'; showOverlay(`${level}단계 완료`, '멋져요! 다음 단계로 가볼까요?', '스페이스키 또는 아래 버튼을 눌러 계속해요', '다음 단계'); }
  }
  function update(dt) {
    if (state === 'won' && celebrationTime > 0) {
      celebrationTime -= dt;
      for (const p of particles) { p.x += p.vx * dt; p.y += p.vy * dt; if (p.y > H) p.y = -10; }
      return;
    }
    if (state !== 'playing') return;
    if (keys.ArrowLeft) paddle.x -= 520 * dt;
    if (keys.ArrowRight) paddle.x += 520 * dt;
    paddle.x = Math.max(0, Math.min(W - paddle.w, paddle.x));
    for (const type of ['expand', 'pierce']) if (effects[type] > 0) {
      effects[type] = Math.max(0, effects[type] - dt);
      if (type === 'expand' && effects.expand === 0) { paddle.x += (paddle.w - 150) / 2; paddle.w = 150; paddle.x = Math.min(W - paddle.w, paddle.x); }
    }
    updateEffects();
    for (const ball of [...balls]) {
      const oldX = ball.x, oldY = ball.y;
      ball.x += ball.vx * dt; ball.y += ball.vy * dt;
      for (const brick of ball.contacts) {
        if (!brick.alive || ball.x + R < brick.x || ball.x - R > brick.x + brick.w || ball.y + R < brick.y || ball.y - R > brick.y + brick.h) ball.contacts.delete(brick);
      }
      if (ball.x < R) { ball.x = R; ball.vx = Math.abs(ball.vx); }
      if (ball.x > W - R) { ball.x = W - R; ball.vx = -Math.abs(ball.vx); }
      if (ball.y < R) { ball.y = R; ball.vy = Math.abs(ball.vy); }
      if (ball.vy > 0 && oldY + R <= PADDLE_Y && ball.y + R >= PADDLE_Y && ball.x + R >= paddle.x && ball.x - R <= paddle.x + paddle.w) {
        ball.y = PADDLE_Y - R; const relative = Math.max(-1, Math.min(1, (ball.x - (paddle.x + paddle.w / 2)) / (paddle.w / 2)));
        const angle = relative * 1.05; ball.vx = speed() * Math.sin(angle); ball.vy = -speed() * Math.cos(angle);
        bounces++; if (bounces >= 15) secret('bounce', 'pierce', ball.x, ball.y);
      }
      for (const brick of bricks) {
        if (!brick.alive || ball.contacts.has(brick) || ball.x + R < brick.x || ball.x - R > brick.x + brick.w || ball.y + R < brick.y || ball.y - R > brick.y + brick.h) continue;
        ball.contacts.add(brick);
        brick.hits--;
        if (brick.hits === 0) {
          brick.alive = false; broken++;
          if (brick.type) drop(brick.type, brick.x + brick.w / 2, brick.y + brick.h);
          if (!lost && broken >= 20) secret('twenty', 'expand', ball.x, ball.y);
          if (!lost && broken >= 40) secret('forty', 'multi', ball.x, ball.y);
        }
        if (!effects.pierce) {
          if (oldY + R <= brick.y || oldY - R >= brick.y + brick.h) ball.vy *= -1;
          else if (oldX + R <= brick.x || oldX - R >= brick.x + brick.w) ball.vx *= -1;
          else ball.vy *= -1;
          break;
        }
      }
      if (broken === 63) { endStage(); return; }
      if (ball.y - R > H) { lost = true; balls.splice(balls.indexOf(ball), 1); }
    }
    if (!balls.length) { state = 'lost'; showOverlay('게임 종료', '아쉽지만, 다시 도전해요!', `${level}단계까지 도착했어요`, '다시 시작'); return; }
    for (const item of [...items]) {
      item.y += 115 * dt;
      if (item.y + 12 >= PADDLE_Y && item.y - 12 <= PADDLE_Y + paddle.h && item.x + 25 >= paddle.x && item.x - 25 <= paddle.x + paddle.w) { applyItem(item.type); items.splice(items.indexOf(item), 1); }
      else if (item.y > H + 20) items.splice(items.indexOf(item), 1);
    }
  }
  function draw() {
    ctx.clearRect(0, 0, W, H);
    for (const brick of bricks) if (brick.alive) {
      ctx.fillStyle = '#e23f39'; ctx.fillRect(brick.x, brick.y, brick.w, brick.h);
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.font = '20px sans-serif';
      ctx.fillText(brick.animal.face, brick.x + 34, brick.y + 15 + Math.sin(last / 280 + brick.x / 80) * 2);
      ctx.fillStyle = '#fff'; ctx.font = 'bold 15px sans-serif';
      ctx.fillText(`${brick.type ? '★ ' : ''}${brick.hits}`, brick.x + 72, brick.y + 15);
    }
    ctx.fillStyle = '#191919'; ctx.fillRect(paddle.x, PADDLE_Y, paddle.w, paddle.h);
    for (const ball of balls) { ctx.beginPath(); ctx.fillStyle = effects.pierce ? '#e99b25' : '#191919'; ctx.arc(ball.x, ball.y, R, 0, Math.PI * 2); ctx.fill(); }
    for (const item of items) { ctx.fillStyle = '#272727'; ctx.fillRect(item.x - 27, item.y - 12, 54, 24); ctx.fillStyle = '#fff'; ctx.font = 'bold 14px sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(symbols[item.type], item.x, item.y); }
    if (state === 'won' && celebrationTime > 0) {
      for (const p of particles) { ctx.fillStyle = p.color; ctx.fillRect(p.x, p.y, p.size, p.size * 1.6); ctx.beginPath(); ctx.arc((p.x * 1.7) % W, (p.y * .6) % H, p.size / 2, 0, Math.PI * 2); ctx.fill(); }
    }
  }
  function frame(time) {
    const dt = last ? Math.min((time - last) / 1000, .035) : 0; last = time;
    update(dt); draw(); requestAnimationFrame(frame);
  }
  document.addEventListener('keydown', event => {
    if (['ArrowLeft', 'ArrowRight', 'Space'].includes(event.code)) event.preventDefault();
    if (event.code === 'Space' && !event.repeat) startAction();
    if (event.code === 'ArrowLeft' || event.code === 'ArrowRight') keys[event.code] = true;
  });
  document.addEventListener('keyup', event => { keys[event.code] = false; });
  window.addEventListener('blur', () => { keys = {}; });
  action.addEventListener('click', startAction);
  let pointer = null, previousX = 0;
  touch.addEventListener('pointerdown', event => { if (state !== 'playing') return; pointer = event.pointerId; previousX = event.clientX; touch.setPointerCapture(pointer); event.preventDefault(); });
  touch.addEventListener('pointermove', event => { if (pointer !== event.pointerId || state !== 'playing') return; const scale = W / canvas.getBoundingClientRect().width; paddle.x = Math.max(0, Math.min(W - paddle.w, paddle.x + (event.clientX - previousX) * scale)); previousX = event.clientX; event.preventDefault(); });
  for (const name of ['pointerup', 'pointercancel', 'lostpointercapture']) touch.addEventListener(name, () => { pointer = null; });
  resetStage(); showOverlay('준비되셨나요?', '벽돌을 모두 깨보세요!', '스페이스키 또는 아래 버튼을 눌러 시작해요', '게임 시작'); requestAnimationFrame(frame);
})();
