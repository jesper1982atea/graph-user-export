import React, { useMemo, useState } from 'react';
import { USER_FIELDS } from './userFields';
import { getFieldLabel } from './fieldLabelMap';
import { normalizeUser } from './userNormalize';
import { saveAs } from 'file-saver';

function CsvExportControls({ items = [], storageKey = 'default', defaultFileName = 'export.csv', buttonLabel = 'Exportera CSV', disabled = false, disabledTitle, fields }) {
  const [open, setOpen] = useState(false);
  const [selected, setSelected] = useState(() => {
    try {
      const raw = localStorage.getItem(`export_fields:${storageKey}`);
      if (raw) {
        const arr = JSON.parse(raw);
        if (Array.isArray(arr)) return arr;
      }
    } catch {}
    return [];
  });
  const candidates = useMemo(() => {
    if (Array.isArray(fields) && fields.length) return fields;
    const set = new Set();
    (items || []).slice(0, 25).forEach(raw => {
      const it = normalizeUser(raw);
      if (it && typeof it === 'object') {
        Object.keys(it).forEach(k => set.add(k));
      }
    });
    const all = Array.from(set).sort();
    // If the data looks like user objects, order with USER_FIELDS first, then extras
    const looksLikeUser = USER_FIELDS.some(f => all.includes(f));
    if (!looksLikeUser) return all;
    const base = USER_FIELDS.filter(f => all.includes(f));
    const extras = all.filter(k => !USER_FIELDS.includes(k));
    return base.concat(extras);
  }, [items, fields]);

  const toggleKey = (k) => setSelected(prev => prev.includes(k) ? prev.filter(x => x !== k) : [...prev, k]);
  const selectAll = () => setSelected(candidates);
  const clearAll = () => setSelected([]);
  const saveSelection = () => {
    try { localStorage.setItem(`export_fields:${storageKey}`, JSON.stringify(selected)); } catch {}
    setOpen(false);
  };
  const exportNow = () => {
    const keys = selected.length ? selected : candidates;
    if (!items || !Array.isArray(items) || items.length === 0 || keys.length === 0) return;
  const header = keys.map(k => getFieldLabel(k)).join(',');
    const rows = items.map(raw => {
      const it = normalizeUser(raw);
      return keys.map(k => {
      let v = it?.[k];
      if (typeof v === 'object' && v !== null) {
        try { v = JSON.stringify(v); } catch { v = ''; }
      }
      v = v == null ? '' : String(v);
      return '"' + v.replace(/"/g, '""') + '"';
    }).join(',');
    });
    const csv = [header, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, defaultFileName);
  };

  return (
    <div style={{ position:'relative', display:'inline-flex', gap:8, alignItems:'center' }}>
      <button className="btn btn-secondary" onClick={exportNow} disabled={disabled} title={disabled ? disabledTitle : undefined}>{buttonLabel}</button>
      <button className="btn btn-light" onClick={() => setOpen(o => !o)} title="Välj fält">⚙️</button>
      {open && (
        <div style={{ position:'absolute', right:0, top:'110%', zIndex:30, background:'var(--card-bg)', border:'1px solid var(--border)', borderRadius:10, boxShadow:'0 8px 30px rgba(0,0,0,.12)', padding:12, minWidth:260 }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:8 }}>
            <b>Välj fält</b>
            <div style={{ display:'flex', gap:6 }}>
              <button className="btn btn-light" onClick={selectAll}>Alla</button>
              <button className="btn btn-light" onClick={clearAll}>Rensa</button>
            </div>
          </div>
          <div style={{ maxHeight:220, overflow:'auto', paddingRight:6 }}>
            {candidates.length === 0 ? (
              <div className="muted">Inga fält tillgängliga.</div>
            ) : candidates.map(k => (
              <label key={k} style={{ display:'flex', alignItems:'center', gap:8, marginBottom:6 }}>
                <input type="checkbox" checked={selected.includes(k)} onChange={() => toggleKey(k)} />
                <span>{k}</span>
              </label>
            ))}
          </div>
          <div style={{ display:'flex', gap:8, marginTop:8, justifyContent:'flex-end' }}>
            <button className="btn btn-light" onClick={() => setOpen(false)}>Stäng</button>
            <button className="btn btn-primary" onClick={saveSelection}>Spara</button>
          </div>
          <small className="muted" style={{ display:'block', marginTop:6 }}>Dina val sparas i webbläsaren.</small>
        </div>
      )}
    </div>
  );
}

export default CsvExportControls;
