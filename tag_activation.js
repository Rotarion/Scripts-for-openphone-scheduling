copy(String((() => {
  function visible(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
  }

  function textOf(el) {
    return ((el && (el.innerText || el.textContent)) || '').replace(/\s+/g, ' ').trim();
  }

  function centerDist(el) {
    const r = el.getBoundingClientRect();
    const cx = r.left + (r.width / 2);
    const cy = r.top + (r.height / 2);
    return Math.hypot(cx - (window.innerWidth / 2), cy - (window.innerHeight / 2));
  }

  function fireMouse(el, type, x, y) {
    const evtType =
      type === 'pointerdown' || type === 'pointerup' || type === 'pointermove'
        ? PointerEvent
        : MouseEvent;

    el.dispatchEvent(new evtType(type, {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX: x,
      clientY: y,
      pointerId: 1,
      pointerType: 'mouse',
      isPrimary: true,
      buttons: type === 'pointerdown' || type === 'mousedown' ? 1 : 0,
      button: 0
    }));
  }

  function activate(el, x, y) {
    if (!el) return false;

    try { el.scrollIntoView({ block: 'center', inline: 'center' }); } catch (e) {}
    try { el.focus(); } catch (e) {}

    if (x == null || y == null) {
      const r = el.getBoundingClientRect();
      x = Math.round(r.left + (r.width / 2));
      y = Math.round(r.top + (r.height / 2));
    }

    try { fireMouse(el, 'pointerdown', x, y); } catch (e) {}
    try { fireMouse(el, 'mousedown', x, y); } catch (e) {}
    try { fireMouse(el, 'pointerup', x, y); } catch (e) {}
    try { fireMouse(el, 'mouseup', x, y); } catch (e) {}
    try { fireMouse(el, 'click', x, y); } catch (e) {}
    try { el.click && el.click(); } catch (e) {}
    try { el.focus(); } catch (e) {}

    return true;
  }

  function nearestClickable(el) {
    let cur = el;
    while (cur && cur !== document.body) {
      if (
        typeof cur.matches === 'function' &&
        cur.matches('button,[role="button"],[role="combobox"],input,textarea,[contenteditable="true"],[tabindex]')
      ) {
        return cur;
      }
      if (typeof cur.onclick === 'function') return cur;
      cur = cur.parentElement;
    }
    return null;
  }

  function clickableAncestor(el) {
    let cur = el;
    while (cur && cur !== document.body) {
      if (
        typeof cur.matches === 'function' &&
        cur.matches('button,[role="button"],[role="combobox"],input,textarea,[contenteditable="true"],[tabindex]')
      ) {
        return cur;
      }
      if (typeof cur.onclick === 'function') return cur;
      cur = cur.parentElement;
    }
    return null;
  }

  function bestClickableFromPoint(x, y) {
    const stack = document.elementsFromPoint(x, y) || [];
    for (const el of stack) {
      if (!visible(el)) continue;
      const clickable = nearestClickable(el);
      if (clickable && visible(clickable)) {
        return clickable;
      }
    }
    return stack[0] || null;
  }

  function findLocalClickable(anchor) {
    const container =
      anchor.closest('[role="row"], .row, .field, .form-group, .input-group, li, section, article') ||
      anchor.parentElement ||
      anchor.closest('div');

    if (!container) return null;

    const candidates = Array.from(
      container.querySelectorAll(
        'button,[role="button"],[role="combobox"],input,textarea,[contenteditable="true"],[tabindex],div,span'
      )
    ).filter(el => visible(el));

    const ar = anchor.getBoundingClientRect();

    let best = null;
    let bestScore = Infinity;

    for (const el of candidates) {
      if (el === anchor) continue;

      const clickable = nearestClickable(el);
      if (!clickable || !visible(clickable)) continue;
      if (clickable === anchor) continue;

      const r = clickable.getBoundingClientRect();
      const dx = Math.max(0, r.left - ar.right, ar.left - r.right);
      const dy = Math.abs((r.top + r.bottom) / 2 - (ar.top + ar.bottom) / 2);
      const score = dx * 10 + dy;

      if (score < bestScore) {
        bestScore = score;
        best = clickable;
      }
    }

    return best;
  }

  function plusNearAnchor(anchor) {
    const ar = anchor.getBoundingClientRect();
    const plusButtons = Array.from(document.querySelectorAll('button'))
      .filter(el => visible(el) && textOf(el) === '+');

    let best = null;
    let bestScore = Infinity;

    for (const btn of plusButtons) {
      const r = btn.getBoundingClientRect();
      const btnCy = (r.top + r.bottom) / 2;
      const anchorCy = (ar.top + ar.bottom) / 2;
      const isRightSide = r.left >= ar.left - 10;
      const verticalClose = Math.abs(btnCy - anchorCy) < 50;
      const horizontalClose = Math.abs(r.left - ar.right) < 180;

      if (!isRightSide || !verticalClose || !horizontalClose) continue;

      const score = Math.abs(r.left - ar.right) + Math.abs(btnCy - anchorCy);
      if (score < bestScore) {
        bestScore = score;
        best = btn;
      }
    }

    return best;
  }

  const anchors = Array.from(document.querySelectorAll('div,span,button'))
    .filter(el => {
      if (!visible(el)) return false;
      const t = textOf(el);
      return /^set tags\.{3}$/i.test(t) || /^tags$/i.test(t);
    });

  anchors.sort((a, b) => centerDist(a) - centerDist(b));

  for (const anchor of anchors) {
    const direct = clickableAncestor(anchor);
    if (direct && visible(direct)) {
      activate(direct, null, null);
      return 'ANCESTOR_TARGET';
    }

    const local = findLocalClickable(anchor);
    if (local && visible(local)) {
      activate(local, null, null);
      return 'LOCAL_TARGET';
    }

    const r = anchor.getBoundingClientRect();
    const testPoints = [
      [Math.round(r.right + 6),  Math.round(r.top + r.height / 2)],
      [Math.round(r.right + 18), Math.round(r.top + r.height / 2)],
      [Math.round(r.right - 8),  Math.round(r.top + r.height / 2)],
      [Math.round(r.left + r.width / 2), Math.round(r.top + r.height / 2)],
      [Math.round(r.left + 8),   Math.round(r.top + r.height / 2)]
    ];

    for (const [x, y] of testPoints) {
      const target = bestClickableFromPoint(x, y);
      if (target && visible(target)) {
        activate(target, x, y);
        return 'HITTEST_TARGET';
      }
    }

    const structural =
      nearestClickable(anchor.previousElementSibling) ||
      nearestClickable(anchor.nextElementSibling) ||
      nearestClickable(anchor.parentElement);

    if (structural && visible(structural)) {
      activate(structural, null, null);
      return 'STRUCTURE_TARGET';
    }

    const localPlus = plusNearAnchor(anchor);
    if (localPlus && visible(localPlus)) {
      activate(localPlus, null, null);
      return 'PLUS_FALLBACK';
    }
  }

  return 'NO_TARGET';
})()))