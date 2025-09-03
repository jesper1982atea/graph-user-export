import React, { useEffect, useMemo, useState } from 'react';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { formatFieldLabel } from './fieldLabelMap';
import AvatarWithPresence from './AvatarWithPresence';
import { USER_FIELDS } from './userFields';
import { normalizeUser } from './userNormalize';

// users: array of user-like objects { id, displayName, mail, userPrincipalName, jobTitle, department, photoUrl }
// token: Graph token for presence/photo lookup
// mode: 'cards' | 'table'
// selectedFields: array of field names; if empty, use a default set
// onChangeSelectedFields: callback(newFields[]) to persist selection
// selectable: show checkboxes per row (for selection flows)
// selectedMap: { [id]: true }
// onToggleSelect: (user) => void
export default function UserList({ users = [], token, mode = 'cards', selectedFields, onChangeSelectedFields, selectable = false, selectedMap = {}, onToggleSelect, onUserClick, fieldsKey = 'default' }) {
  const defaultFields = useMemo(() => ['displayName','mail','userPrincipalName','jobTitle','department'], []);
  const normalize = (f) => {
    if (!f || typeof f !== 'string') return '';
    return f.charAt(0).toLowerCase() + f.slice(1);
  };
  const [fields, setFields] = useState(() => {
    const init = Array.isArray(selectedFields) && selectedFields.length ? selectedFields : defaultFields;
    return init.map(normalize);
  });
  const [showControls, setShowControls] = useState(() => {
    try { return (localStorage.getItem(`userlist_controls_open:${fieldsKey}`) === '1'); } catch { return false; }
  });
  useEffect(() => {
    try { setShowControls(localStorage.getItem(`userlist_controls_open:${fieldsKey}`) === '1'); } catch {}
  }, [fieldsKey]);
  const toggleControls = () => {
    setShowControls(prev => {
      const next = !prev;
      try { localStorage.setItem(`userlist_controls_open:${fieldsKey}`, next ? '1' : '0'); } catch {}
      return next;
    });
  };
  const [presenceMap, setPresenceMap] = useState({}); // { idOrUpn: presenceString }
  const data = useMemo(() => (users || []).map(u => normalizeUser(u)), [users]);

  const [photoBusy, setPhotoBusy] = useState(false);
  const [photoStatus, setPhotoStatus] = useState('');

  const toAsciiSlug = (s) => {
    try {
      return (s || '')
        .normalize('NFD')
        .replace(/\p{Diacritic}+/gu, '')
        .replace(/[^a-zA-Z0-9]+/g, '.')
        .replace(/^\.+|\.+$/g, '')
        .toLowerCase();
    } catch {
      return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
    }
  };

  const buildFileBase = (u) => {
    const gn = (u.givenName || '').trim();
    const sn = (u.surname || '').trim();
    if (gn && sn) return toAsciiSlug(`${gn}.${sn}`);
    const dn = (u.displayName || '').trim();
    if (dn) return toAsciiSlug(dn.replace(/\s+/g, '.'));
    const upn = (u.userPrincipalName || u.mail || '').trim();
    if (upn) return toAsciiSlug(upn.split('@')[0]);
    return toAsciiSlug(u.id || 'user');
  };

  const downloadPhotosZip = async (list) => {
    const source = Array.isArray(list) ? list : data;
    if (!token) { alert('Ingen token. Verifiera token först.'); return; }
    if (!source.length) { alert('Inga användare att hämta bilder för.'); return; }
    setPhotoBusy(true);
    setPhotoStatus('Förbereder…');
    try {
      const Authorization = (token || '').trim().toLowerCase().startsWith('bearer ') ? token : `Bearer ${token}`;
      const zip = new JSZip();
      const usedNames = new Set();
      let fetched = 0;
      // Limit concurrency to avoid throttling
      const concurrency = 4;
      const queue = [...source];
      const runWorker = async () => {
        while (queue.length) {
          const u = queue.shift();
          const id = u.id || u.userPrincipalName || u.mail;
          if (!id) { continue; }
          try {
            const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(id)}/photo/$value`, {
              headers: { Authorization },
            });
            if (!res.ok) { continue; }
            const blob = await res.blob();
            const ct = res.headers.get('content-type') || 'image/jpeg';
            const ext = ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg';
            let base = buildFileBase(u);
            let name = `${base}.${ext}`;
            let i = 2;
            while (usedNames.has(name)) { name = `${base}-${i++}.${ext}`; }
            usedNames.add(name);
            zip.file(name, blob);
            fetched++;
            if (fetched % 5 === 0) setPhotoStatus(`Hämtat ${fetched} bilder…`);
          } catch {
            // ignore this user
          }
        }
      };
      await Promise.all(Array.from({ length: concurrency }, runWorker));
      if (fetched === 0) {
        setPhotoStatus('Inga bilder hittades.');
        setPhotoBusy(false);
        return;
      }
      setPhotoStatus('Skapar ZIP…');
      const out = await zip.generateAsync({ type: 'blob' });
      saveAs(out, 'anvandar-bilder.zip');
      setPhotoStatus(`Klar. ${fetched} bilder sparade.`);
    } catch (e) {
      setPhotoStatus('Kunde inte skapa ZIP.');
    } finally {
      setPhotoBusy(false);
    }
  };

  const selectedCount = useMemo(() => {
    const m = selectedMap || {};
    return Object.keys(m).filter(id => m[id]).length;
  }, [selectedMap]);

  const downloadSelectedPhotosZip = async () => {
    const m = selectedMap || {};
    const subset = data.filter(u => u.id && m[u.id]);
    if (!subset.length) { alert('Inga valda användare.'); return; }
    return downloadPhotosZip(subset);
  };

  useEffect(() => {
    if (Array.isArray(selectedFields)) {
      setFields((selectedFields.length ? selectedFields : defaultFields).map(normalize));
    }
  }, [selectedFields, defaultFields]);

  // Fetch presence for up to 20 at a time using batch presence endpoint
  useEffect(() => {
    if (!token || !data.length) return;
    const guidRe = /^[0-9a-fA-F-]{36}$/;
    const ids = data.map(u => u.id || u.userPrincipalName || u.mail).filter(x => typeof x === 'string' && guidRe.test(x));
    if (!ids.length) return;
    const Authorization = (token || '').trim().toLowerCase().startsWith('bearer ') ? token : `Bearer ${token}`;
    const chunk = (arr, size) => arr.reduce((acc, _, i) => (i % size ? acc : acc.concat([arr.slice(i, i+size)])), []);
    let cancelled = false;
    const run = async () => {
      const newMap = {};
      for (const slice of chunk(ids, 20)) {
        try {
          const res = await fetch('https://graph.microsoft.com/v1.0/communications/getPresencesByUserId', {
            method: 'POST',
            headers: { Authorization, 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: slice }),
          });
          if (!res.ok) continue;
          const data = await res.json();
          (data.value || []).forEach(p => { newMap[p.id] = p.availability || p.activity || 'offline'; });
        } catch {}
      }
      if (!cancelled) setPresenceMap(prev => ({ ...prev, ...newMap }));
    };
    run();
    return () => { cancelled = true; };
  }, [token, data]);

  const allFieldOptions = useMemo(() => {
    const base = ['displayName','mail','userPrincipalName','jobTitle','department','companyName','mobilePhone','officeLocation','country','city',
      // manager flat fields
      'managerDisplayName','managerMail','managerUserPrincipalName','managerJobTitle',
      // extension attributes
  ...Array.from({ length: 15 }, (_, i) => `extensionAttribute${i+1}`),
  // derived summaries
  'assignedPlansCount','assignedPlansServices','provisionedPlansCount','provisionedPlansServices','managedDevicesCount'
    ];
    const userFieldsNorm = USER_FIELDS.map(f => normalize(f)).filter(Boolean);
    const all = base.concat(userFieldsNorm);
    const seen = new Set();
    const dedup = [];
    for (const f of all) {
      const k = f.toLowerCase();
      if (!seen.has(k)) { seen.add(k); dedup.push(f); }
    }
    return dedup;
  }, []);

  const handleFieldToggle = (f) => {
    const next = fields.includes(f) ? fields.filter(x => x !== f) : fields.concat([f]);
    setFields(next);
    onChangeSelectedFields && onChangeSelectedFields(next);
  };

  const moveField = (f, dir) => {
    const idx = fields.indexOf(f);
    if (idx < 0) return;
    const j = dir === 'up' ? idx - 1 : idx + 1;
    if (j < 0 || j >= fields.length) return;
    const next = fields.slice();
    const tmp = next[idx];
    next[idx] = next[j];
    next[j] = tmp;
    setFields(next);
    onChangeSelectedFields && onChangeSelectedFields(next);
  };

  const removeField = (f) => {
    if (!fields.includes(f)) return;
    const next = fields.filter(x => x !== f);
    setFields(next);
    onChangeSelectedFields && onChangeSelectedFields(next);
  };

  // Drag & drop reordering for selected fields
  const [dragIndex, setDragIndex] = useState(null);
  const [overIndex, setOverIndex] = useState(null);
  const reorder = (arr, from, to) => {
    if (from === to || from == null || to == null) return arr;
    const next = arr.slice();
    const [item] = next.splice(from, 1);
    next.splice(to, 0, item);
    return next;
  };

  const FieldControls = () => (
    <div style={{ display:'grid', gap:10, marginBottom:8 }}>
      <div>
        <b>Valda fält (ordning):</b>
        <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginTop:6 }}>
          {fields.length === 0 && <span className="muted">Inga valda fält.</span>}
          {fields.map((f, i) => (
            <div
              key={f+':'+i}
              draggable
              onDragStart={(e) => { setDragIndex(i); e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', String(i)); } catch {} }}
              onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; if (overIndex !== i) setOverIndex(i); }}
              onDragEnter={(e) => { e.preventDefault(); if (overIndex !== i) setOverIndex(i); }}
              onDrop={(e) => { e.preventDefault(); const to = i; const from = dragIndex != null ? dragIndex : (() => { try { return parseInt(e.dataTransfer.getData('text/plain')||'-1', 10); } catch { return -1; } })(); const next = reorder(fields, from, to); if (next !== fields) { setFields(next); onChangeSelectedFields && onChangeSelectedFields(next); } setDragIndex(null); setOverIndex(null); }}
              onDragEnd={() => { setDragIndex(null); setOverIndex(null); }}
              aria-grabbed={dragIndex === i}
              title="Dra för att flytta"
              style={{ display:'inline-flex', alignItems:'center', gap:6, border:'1px solid var(--border)', padding:'4px 8px', borderRadius:8, background:'var(--card-bg)', outline: overIndex===i ? '2px dashed var(--border)' : 'none' }}
            >
              <span style={{ fontWeight:600 }}>{formatFieldLabel(f)}</span>
              <div style={{ display:'inline-flex', gap:4 }}>
                <button className="btn btn-ghost" title="Flytta upp" onClick={() => moveField(f, 'up')} disabled={i === 0}>↑</button>
                <button className="btn btn-ghost" title="Flytta ner" onClick={() => moveField(f, 'down')} disabled={i === fields.length - 1}>↓</button>
                <button className="btn btn-ghost" title="Ta bort" onClick={() => removeField(f)}>✕</button>
              </div>
            </div>
          ))}
        </div>
      </div>
      <div>
        <b>Lägg till/ta bort fält:</b>
        <div style={{ display:'flex', flexWrap:'wrap', gap:8, alignItems:'center', marginTop:6 }}>
          {allFieldOptions.map(f => (
            <label key={f} className="muted" style={{ display:'inline-flex', alignItems:'center', gap:6, border:'1px solid var(--border)', padding:'2px 6px', borderRadius:6 }}>
              <input type="checkbox" checked={fields.includes(f)} onChange={() => handleFieldToggle(f)} />
              <span>{f}</span>
            </label>
          ))}
        </div>
      </div>
    </div>
  );

  if (!users.length) return <p className="list-empty">Inga användare.</p>;

  if (mode === 'table') {
    return (
      <div>
        <div style={{ display:'flex', justifyContent:'space-between', marginBottom:8, alignItems:'center', gap:8, flexWrap:'wrap' }}>
          <div className="muted">{photoStatus}</div>
          <div style={{ display:'flex', gap:8 }}>
            <button className="btn btn-light" onClick={toggleControls}>{showControls ? 'Dölj fält' : 'Visa fält'}</button>
            <button className="btn btn-secondary" onClick={() => downloadPhotosZip()} disabled={photoBusy}>Ladda ner bilder (ZIP)</button>
            {selectable && (
              <button className="btn btn-secondary" onClick={downloadSelectedPhotosZip} disabled={photoBusy || selectedCount === 0} title={selectedCount ? `Valda: ${selectedCount}` : 'Välj minst en användare'}>
                Ladda ner valda (ZIP)
              </button>
            )}
          </div>
        </div>
        {showControls && <FieldControls />}
        <div style={{ overflowX:'auto' }}>
          <table className="table" style={{ width:'100%', borderCollapse:'separate', borderSpacing:0 }}>
            <thead>
              <tr>
                {selectable && <th style={{ width:36 }} />}
                <th style={{ minWidth:220 }}>Användare</th>
                {fields.map(f => (
                  <th key={f} style={{ textTransform:'none' }}>{formatFieldLabel(f)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((u, idx) => {
                const key = u.id || u.userPrincipalName || u.mail || idx;
                const presence = presenceMap[u.id] || presenceMap[u.userPrincipalName] || presenceMap[u.mail];
                const handleClick = () => {
                  if (onUserClick) return onUserClick(u);
                  try { window.location.hash = `#/users/${encodeURIComponent(u.id || u.userPrincipalName || u.mail)}`; } catch {}
                };
                return (
                  <tr key={key} className="row-hover" onClick={handleClick} style={{ cursor:'pointer' }}>
                    {selectable && (
                      <td>
                        <input type="checkbox" checked={!!selectedMap[u.id]} onChange={(e) => { e.stopPropagation(); onToggleSelect && onToggleSelect(u); }} onClick={(e) => e.stopPropagation()} />
                      </td>
                    )}
                    <td>
                      <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                        <AvatarWithPresence photoUrl={u.photoUrl} name={u.displayName || u.mail || u.userPrincipalName} presence={presence} />
                        <div style={{ minWidth:0 }}>
                          <div style={{ fontWeight:600 }}>{u.displayName || '(okänd)'}</div>
                          <div className="muted" style={{ fontSize:'.9rem' }}>{u.mail || u.userPrincipalName || ''}</div>
                        </div>
                      </div>
                    </td>
                    {fields.map(f => (
                      <td key={f} className="muted" style={{ fontSize:'.92rem' }}>
                        {f === 'department' && u.department ? (
                          <button
                            className="linklike"
                            onClick={(e) => { e.stopPropagation(); try { window.location.hash = `#/search?department=${encodeURIComponent(u.department)}`; } catch {} }}
                            title="Visa alla användare i avdelningen"
                            style={{ background:'none', border:'none', color:'var(--link)', cursor:'pointer', padding:0 }}
                          >
                            {u.department}
                          </button>
                        ) : (
                          stringify(u[f])
                        )}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  // cards mode
  return (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', marginBottom:8, alignItems:'center', gap:8, flexWrap:'wrap' }}>
        <div className="muted">{photoStatus}</div>
        <div style={{ display:'flex', gap:8 }}>
          <button className="btn btn-light" onClick={toggleControls}>{showControls ? 'Dölj fält' : 'Visa fält'}</button>
          <button className="btn btn-secondary" onClick={() => downloadPhotosZip()} disabled={photoBusy}>Ladda ner bilder (ZIP)</button>
          {selectable && (
            <button className="btn btn-secondary" onClick={downloadSelectedPhotosZip} disabled={photoBusy || selectedCount === 0} title={selectedCount ? `Valda: ${selectedCount}` : 'Välj minst en användare'}>
              Ladda ner valda (ZIP)
            </button>
          )}
        </div>
      </div>
      {showControls && <FieldControls />}
      <div style={{ display:'flex', flexWrap:'wrap', gap:'1rem', marginTop:12 }}>
  {data.map((u, idx) => {
          const key = u.id || u.userPrincipalName || u.mail || idx;
          const presence = presenceMap[u.id] || presenceMap[u.userPrincipalName] || presenceMap[u.mail];
          const handleClick = () => {
            if (onUserClick) return onUserClick(u);
            try { window.location.hash = `#/users/${encodeURIComponent(u.id || u.userPrincipalName || u.mail)}`; } catch {}
          };
          return (
            <div key={key} className="card" style={{ width:320, cursor:'pointer' }} onClick={handleClick}>
              <div style={{ display:'flex', alignItems:'center', gap:12 }}>
                <AvatarWithPresence photoUrl={u.photoUrl} name={u.displayName || u.mail || u.userPrincipalName} presence={presence} size={56} />
                <div style={{ minWidth:0 }}>
                  <div className="grid-item-title">{u.displayName || '(okänd)'}</div>
                  <div className="grid-item-sub">{u.mail || u.userPrincipalName || ''}</div>
                  {u.jobTitle && <div className="grid-item-meta">{u.jobTitle}</div>}
                </div>
                {selectable && (
                  <label style={{ marginLeft:'auto' }}>
        <input type="checkbox" checked={!!selectedMap[u.id]} onChange={(e) => { e.stopPropagation(); onToggleSelect && onToggleSelect(u); }} onClick={(e) => e.stopPropagation()} />
                  </label>
                )}
              </div>
              {fields.length > 0 && (
                <div style={{ marginTop:8 }}>
                  {fields.map(f => (
                    <div key={f} className="muted" style={{ fontSize:'.92rem' }}>
                      <b style={{ textTransform:'none' }}>{formatFieldLabel(f)}:</b>{' '}
                      {f === 'department' && u.department ? (
                        <button
                          className="linklike"
                          onClick={(e) => { e.stopPropagation(); try { window.location.hash = `#/search?department=${encodeURIComponent(u.department)}`; } catch {} }}
                          title="Visa alla användare i avdelningen"
                          style={{ background:'none', border:'none', color:'var(--link)', cursor:'pointer', padding:0 }}
                        >
                          {u.department}
                        </button>
                      ) : (
                        stringify(u[f])
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function stringify(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.join(', ');
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}
