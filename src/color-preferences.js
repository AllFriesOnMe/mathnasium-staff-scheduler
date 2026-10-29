'use strict';

const ColorPreferences = (() => {
  const storageKey = 'mathnasium.schedule-colors.v1';
  const validColor = (value) =>
    typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : null;
  let preferences = { directorColor: null, assessmentColor: null };
  const listeners = new Set();
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey));
    preferences = {
      directorColor: validColor(saved?.directorColor),
      assessmentColor: validColor(saved?.assessmentColor),
    };
    if (preferences.directorColor && preferences.directorColor === preferences.assessmentColor)
      preferences.assessmentColor = null;
  } catch {
    /* Storage may be disabled or contain an invalid value. */
  }

  function render(kind) {
    const color = preferences[kind + 'Color'];
    const input = document.getElementById(kind + 'Color');
    input.value = color || '#000000';
    document.getElementById(kind + 'ColorStatus').textContent = color || 'Not set';
    const canvas = document.getElementById(kind + 'Swatch');
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

  function update(kind, value) {
    const next = validColor(value);
    const other = kind === 'director' ? 'assessmentColor' : 'directorColor';
    const status = document.getElementById('colorPreferenceStatus');
    if (next && next === preferences[other]) {
      status.textContent = 'Choose different colors for the director and assessment.';
      render(kind);
      return;
    }
    preferences[kind + 'Color'] = next;
    status.textContent = '';
    try {
      localStorage.setItem(storageKey, JSON.stringify(preferences));
    } catch {
      status.textContent = 'Colors apply now, but this browser could not save them for next time.';
    }
    render(kind);
    listeners.forEach((listener) => listener());
  }

  for (const kind of ['director', 'assessment']) {
    render(kind);
    document
      .getElementById(kind + 'Color')
      .addEventListener('change', (event) => update(kind, event.target.value));
    const clearId = kind === 'director' ? 'clearDirectorColor' : 'clearAssessmentColor';
    document.getElementById(clearId).addEventListener('click', () => update(kind, null));
  }
  return { get: () => ({ ...preferences }), onChange: (listener) => listeners.add(listener) };
})();
