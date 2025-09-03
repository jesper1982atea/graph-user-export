export function loadFieldLabelMap() {
  try {
    const raw = localStorage.getItem('field_label_map');
    const obj = raw ? JSON.parse(raw) : {};
    return obj && typeof obj === 'object' ? obj : {};
  } catch {
    return {};
  }
}

export function saveFieldLabelMap(map) {
  try {
    localStorage.setItem('field_label_map', JSON.stringify(map || {}));
  } catch {}
}

export function getFieldLabel(key) {
  if (!key) return '';
  const map = loadFieldLabelMap();
  // Try exact, then lowercase match to be forgiving
  return map[key] || map[key.toLowerCase()] || key;
}

export function formatFieldLabel(key) {
  const label = getFieldLabel(key);
  return label === key ? key : `${label} (${key})`;
}
