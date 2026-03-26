(() => {
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
    el.dispatchEvent(new MouseEvent(type, {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX: x,
      clientY: y
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
    try { fireMouse(el, 'mouseup', x, y); } catch (e) {}
    try { fireMouse(el, 'click', x, y); } catch (e) {}
    try { el.click && el.click(); } catch (e) {}
    try { el.focus(); } catch (e) {}
    return true;
  }

  function bestClickableFromPoint(x, y) {
    const stack = document.elementsFromPoint(x, y) || [];
    for (const el of stack) {
      if (!visible(el)) continue;
      if (typeof el.matches === 'function' &&
          el.matches('button,[tabindex],[role="button"],[role="combobox"],input,textarea,[contenteditable="true"]')) {
        return el;
      }
      if (typeof el.onclick === 'function') return el;
    }
    return stack[0] || null;
  }

  const anchors = Array.from(document.querySelectorAll('div,span,button'))
    .filter(el => visible(el) && (/^set tags\.{3}$/i.test(textOf(el)) || /^tags$/i.test(textOf(el))));

  anchors.sort((a, b) => centerDist(a) - centerDist(b));

  for (const anchor of anchors) {
    const r = anchor.getBoundingClientRect();
    const testPoints = [
      [Math.round(r.left + r.width - 8), Math.round(r.top + r.height / 2)],
      [Math.round(r.left + r.width / 2), Math.round(r.top + r.height / 2)],
      [Math.round(r.left + 8), Math.round(r.top + r.height / 2)]
    ];

    let target = null;

    for (const [x, y] of testPoints) {
      target = bestClickableFromPoint(x, y);
      if (target && target !== anchor) {
        activate(target, x, y);
        return 'HITTEST_TARGET';
      }
    }

    target =
      anchor.closest('button,[tabindex],[role="button"],[role="combobox"]') ||
      anchor.previousElementSibling ||
      anchor.nextElementSibling;

    if (target && visible(target)) {
      activate(target, null, null);
      return 'STRUCTURE_TARGET';
    }
  }

  const plusButtons = Array.from(document.querySelectorAll('button'))
    .filter(el => visible(el) && textOf(el) === '+');

  if (plusButtons.length) {
    plusButtons.sort((a, b) => centerDist(a) - centerDist(b));
    activate(plusButtons[0], null, null);
    return 'PLUS_FALLBACK';
  }

  return 'NO_TARGET';
})()