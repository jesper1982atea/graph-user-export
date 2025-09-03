import React, { useEffect, useMemo, useState } from 'react';
import { loadFieldLabelMap, saveFieldLabelMap, getFieldLabel } from './fieldLabelMap';
import { USER_FIELDS } from './userFields';
import { normalizeUser } from './userNormalize';

function uniqueKeysFromItems(items) {
  const set = new Set();
  (items || []).slice(0, 200).forEach(raw => {
    const it = normalizeUser(raw) || raw;
    if (it && typeof it === 'object') {
      Object.keys(it).forEach(k => set.add(k));
    }
  });
  return Array.from(set);
}

export default function FieldLabelMappingPage({ items = [], onBack }) {
  const [map, setMap] = useState(() => loadFieldLabelMap());
  const [query, setQuery] = useState('');
  const [onlyWithData, setOnlyWithData] = useState(false);

  useEffect(() => {
    // Ensure we reflect external changes if any
    setMap(loadFieldLabelMap());
  }, []);

  const discovered = useMemo(() => uniqueKeysFromItems(items), [items]);
  const extKeys = useMemo(() => Array.from({ length: 15 }, (_, i) => `extensionAttribute${i + 1}`), []);
  const derivedKeys = [
    'managerDisplayName', 'managerMail', 'managerUserPrincipalName', 'managerJobTitle',
    'assignedPlansCount', 'assignedPlansServices', 'provisionedPlansCount', 'provisionedPlansServices', 'managedDevicesCount',
  ];
  const candidates = useMemo(() => {
    const base = new Set([ ...USER_FIELDS, ...extKeys, ...derivedKeys, ...discovered ]);
    return Array.from(base).sort((a,b) => a.localeCompare(b));
  }, [discovered, extKeys]);

  const itemsNormalized = useMemo(() => (items || []).map(normalizeUser), [items]);

  const filteredKeys = useMemo(() => {
    let keys = candidates;
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      keys = keys.filter(k => k.toLowerCase().includes(q) || getFieldLabel(k).toLowerCase().includes(q));
    }
    if (onlyWithData && itemsNormalized.length) {
      keys = keys.filter(k => itemsNormalized.some(it => {
        const v = it?.[k];
        if (v == null) return false;
        if (typeof v === 'string') return v.trim().length > 0;
        if (Array.isArray(v)) return v.length > 0;
        if (typeof v === 'object') return Object.keys(v).length > 0;
        return String(v).length > 0;
      }));
    }
    return keys;
  }, [candidates, query, onlyWithData, itemsNormalized]);

  const upsert = (key, label) => {
    const next = { ...map };
    if (!label) delete next[key]; else next[key] = label;
    setMap(next);
    saveFieldLabelMap(next);
  };
  const clearAll = () => { setMap({}); saveFieldLabelMap({}); };

  return (
    <div className="card">
      <div style={{ display:'flex', alignItems:'center', gap:8 }}>
        <b style={{ fontSize: '1.1rem' }}>Rubrikmappning</b>
        <span className="muted">— Byt visningsnamn för attribut i tabeller och CSV.</span>
        <div className="spacer" />
        {onBack && <button className="btn btn-light" onClick={onBack}>Tillbaka</button>}
      </div>
      <div className="muted" style={{ marginTop:6 }}>Ändringar sparas lokalt i webbläsaren. Exempel: extensionAttribute1 → Kostnadsställe.</div>

      <div style={{ marginTop:12, display:'flex', gap:8, flexWrap:'wrap', alignItems:'center' }}>
        <input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Filtrera på namn…" style={{ minWidth:260 }} />
        <label style={{ display:'flex', alignItems:'center', gap:8 }}>
          <input type="checkbox" checked={onlyWithData} onChange={e=>setOnlyWithData(e.target.checked)} />
          Visa endast fält med data
        </label>
        <div className="spacer" />
        <button className="btn btn-light" onClick={clearAll}>Rensa alla mappningar</button>
      </div>

      <div style={{ marginTop:12, display:'grid', gap:8 }}>
        {filteredKeys.length === 0 ? (
          <div className="muted">Inga fält att visa.</div>
        ) : filteredKeys.map(k => (
          <div key={k} style={{ display:'flex', alignItems:'center', gap:10 }}>
            <code style={{ background:'var(--bg)', padding:'2px 6px', borderRadius:6 }}>{k}</code>
            <span>→</span>
            <input
              value={map[k] ?? ''}
              onChange={e => upsert(k, e.target.value)}
              placeholder={getFieldLabel(k)}
              style={{ minWidth:260 }}
            />
            {map[k] && (
              <button className="btn btn-light" onClick={() => upsert(k, '')}>Återställ</button>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
