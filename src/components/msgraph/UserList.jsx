import React, { useEffect, useMemo, useState } from 'react';
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

  const FieldControls = () => (
    <div style={{ display:'flex', flexWrap:'wrap', gap:8, alignItems:'center', marginBottom:8 }}>
      <b>Visa fält:</b>
      {allFieldOptions.map(f => (
        <label key={f} className="muted" style={{ display:'inline-flex', alignItems:'center', gap:6, border:'1px solid var(--border)', padding:'2px 6px', borderRadius:6 }}>
          <input type="checkbox" checked={fields.includes(f)} onChange={() => handleFieldToggle(f)} />
          <span>{f}</span>
        </label>
      ))}
    </div>
  );

  if (!users.length) return <p className="list-empty">Inga användare.</p>;

  if (mode === 'table') {
    return (
      <div>
        <div style={{ display:'flex', justifyContent:'flex-end', marginBottom:8 }}>
          <button className="btn btn-light" onClick={toggleControls}>{showControls ? 'Dölj fält' : 'Visa fält'}</button>
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
      <div style={{ display:'flex', justifyContent:'flex-end', marginBottom:8 }}>
        <button className="btn btn-light" onClick={toggleControls}>{showControls ? 'Dölj fält' : 'Visa fält'}</button>
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
