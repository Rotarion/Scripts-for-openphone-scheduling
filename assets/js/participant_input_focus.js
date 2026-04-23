copy(String((() => {
  const inputs = document.querySelectorAll('input[aria-label="participant input"]');
  const visibleInput = Array.from(inputs).find(el => el.offsetParent !== null);

  visibleInput?.focus();
  visibleInput?.click();

  return visibleInput ? 'OK_PARTICIPANT_INPUT' : 'NO_PARTICIPANT_INPUT';
})()));
