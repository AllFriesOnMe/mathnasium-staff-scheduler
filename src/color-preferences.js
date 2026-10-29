'use strict';

const ColorPreferences = (() => {
  const storageKey = 'mathnasium.schedule-colors.v1';
  const validColor = (value) =>
    typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : null;
  let indicators = [{ name: 'Assessment', color: null, behavior: 'assessment' }];
  const listeners = new Set();
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey));
    if (Array.isArray(saved?.colorIndicators)) {
      indicators = [];
      const used = new Set();
      for (const entry of saved.colorIndicators.slice(0, 50)) {
        if (!entry || typeof entry.name !== 'string') continue;
        const color = validColor(entry.color);
        if (color && used.has(color)) continue;
        indicators.push({
          name: entry.name.trim().slice(0, 80),
          color,
          behavior: entry.behavior === 'assessment' ? 'assessment' : 'instructor',
        });
        if (color) used.add(color);
      }
    }
    // Convert previous fixed settings into ordinary, removable indicators.
    const legacyAssessment = validColor(saved?.assessmentColor);
    const legacyDirector = validColor(saved?.directorColor);
    for (const [name, color, behavior] of [
      ['Assessment', legacyAssessment, 'assessment'],
      ['Center director', legacyDirector, 'instructor'],
    ]) {
      if (!color || indicators.some((entry) => entry.color === color)) continue;
      const unset = indicators.find(
        (entry) => entry.name === name && !entry.color && entry.behavior === behavior,
      );
      if (unset) unset.color = color;
      else if (indicators.length < 50) indicators.push({ name, color, behavior });
    }
  } catch {
    /* Storage may be disabled or contain an invalid value. */
  }

  function snapshot() {
    // Keep an explicit empty list so removed defaults stay removed on reload.
    return { colorIndicators: indicators.map((entry) => ({ ...entry })) };
  }

  function save() {
    const status = document.getElementById('colorPreferenceStatus');
    status.textContent = '';
    try {
      localStorage.setItem(storageKey, JSON.stringify(snapshot()));
    } catch {
      status.textContent = 'Colors apply now, but this browser could not save them for next time.';
    }
    listeners.forEach((listener) => listener());
  }
  function colorUsed(color, except = null) {
    return color && indicators.some((entry) => entry !== except && entry.color === color);
  }
  function paint(canvas, color) {
    const context = canvas.getContext('2d');
    context.clearRect(0, 0, 32, 32);
    context.fillStyle = color || '#FFFFFF';
    context.fillRect(0, 0, 32, 32);
    if (!color) {
      context.strokeStyle = '#AAAAAA';
      context.beginPath();
      context.moveTo(4, 28);
      context.lineTo(28, 4);
      context.stroke();
    }
  }
  function renderIndicators() {
    const list = document.getElementById('colorIndicators');
    list.replaceChildren();
    indicators.forEach((entry, index) => {
      const row = document.createElement('div');
      row.className = 'color-setting custom-indicator';
      const name = document.createElement('input');
      name.type = 'text';
      name.maxLength = 80;
      name.className = 'indicator-name';
      name.placeholder = 'Instructor or group name';
      name.value = entry.name;
      name.setAttribute('aria-label', `Color indicator ${index + 1} name`);
      name.addEventListener('input', () => {
        entry.name = name.value;
        save();
      });
      const picker = document.createElement('span');
      picker.className = 'color-picker';
      const swatch = document.createElement('canvas');
      swatch.width = swatch.height = 32;
      swatch.setAttribute('aria-hidden', 'true');
      const color = document.createElement('input');
      color.type = 'color';
      color.value = entry.color || '#000000';
      color.setAttribute('aria-label', `Color indicator ${index + 1} color`);
      const status = document.createElement('span');
      status.textContent = entry.color || 'Not set';
      paint(swatch, entry.color);
      color.addEventListener('change', () => {
        const next = validColor(color.value);
        if (colorUsed(next, entry)) {
          document.getElementById('colorPreferenceStatus').textContent =
            'Choose different colors for each indicator.';
          color.value = entry.color || '#000000';
          return;
        }
        entry.color = next;
        status.textContent = next || 'Not set';
        paint(swatch, next);
        save();
      });
      picker.append(swatch, color);
      const behavior = document.createElement('select');
      behavior.setAttribute('aria-label', `Color indicator ${index + 1} behavior`);
      for (const [value, label] of [
        ['instructor', 'Instructor / group'],
        ['assessment', 'Assessment'],
      ]) {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = label;
        behavior.append(option);
      }
      behavior.value = entry.behavior;
      behavior.addEventListener('change', () => {
        entry.behavior = behavior.value;
        save();
      });
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.textContent = 'Remove';
      remove.setAttribute('aria-label', `Remove color indicator ${index + 1}`);
      remove.addEventListener('click', () => {
        indicators.splice(index, 1);
        save();
        renderIndicators();
      });
      row.append(name, picker, status, behavior, remove);
      list.append(row);
    });
    document.getElementById('addColorIndicator').disabled = indicators.length >= 50;
  }

  renderIndicators();
  document.getElementById('addColorIndicator').addEventListener('click', () => {
    if (indicators.length >= 50) return;
    indicators.push({ name: '', color: null, behavior: 'instructor' });
    save();
    renderIndicators();
    document.querySelector('#colorIndicators .custom-indicator:last-child .indicator-name').focus();
  });
  return { get: snapshot, onChange: (listener) => listeners.add(listener) };
})();
