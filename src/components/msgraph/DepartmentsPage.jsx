import React, { useEffect, useMemo, useRef, useState } from 'react';
import UserList from './UserList';
import { GRAPH_USER_SELECT_FIELDS } from './graphUserSelect';

// Generic attributes explorer: pick one or two fields, list unique values (or combos) + counts, lazy-loaded via Graph
export default function DepartmentsPage({ token }) {
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  const [filter, setFilter] = useState(() => { try { return localStorage.getItem('attr_filter') || ''; } catch { return ''; } });
  const [primary, setPrimary] = useState(() => { try { return localStorage.getItem('attr_primary') || 'department'; } catch { return 'department'; } });
  const [secondary, setSecondary] = useState(() => { try { return localStorage.getItem('attr_secondary') || ''; } catch { return ''; } });
  const [items, setItems] = useState(() => { try { return JSON.parse(localStorage.getItem('attr_cache') || '[]'); } catch { return []; } }); // [{ k1, k2?, count }]
  const nextLinkRef = useRef('');
  const abortRef = useRef(false);
  const [rowStates, setRowStates] = useState({}); // key => { open, loading, error, users, total, usedClientFilter }

  const authHeader = useMemo(() => {
    const t = (token || '').trim();
    return t ? (t.toLowerCase().startsWith('bearer ') ? t : `Bearer ${t}`) : '';
  }, [token]);

  const FIELD_OPTIONS = useMemo(() => {
    const base = [
      { key:'department', label:'Avdelning' },
      { key:'companyName', label:'Företag' },
      { key:'officeLocation', label:'Kontor' },
      { key:'city', label:'Stad' },
      { key:'country', label:'Land' },
      { key:'jobTitle', label:'Titel' },
      { key:'upnDomain', label:'Domän (UPN/mail, härledd)' },
    ];
    const ext = Array.from({ length: 15 }, (_, i) => ({ key:`extensionAttribute${i+1}`, label:`extensionAttribute${i+1}` }));
    return base.concat(ext);
  }, []);

  const FILTER_FIELD_MAP = useMemo(() => {
    const map = {};
    for (let i=1;i<=15;i++) map[`extensionAttribute${i}`] = `onPremisesExtensionAttributes/extensionAttribute${i}`;
    return map; // others use identity
  }, []);

  useEffect(() => { abortRef.current = false; return () => { abortRef.current = true; }; }, []);

  const saveCache = (arr) => {
    try { localStorage.setItem('attr_cache', JSON.stringify(arr || [])); } catch {}
  };

  const getSelectFor = (p, s) => {
    const set = new Set(['id']);
    const add = (k) => {
      if (!k) return;
      if (k.startsWith('extensionAttribute')) set.add('onPremisesExtensionAttributes');
      else if (k === 'upnDomain') { set.add('userPrincipalName'); set.add('mail'); }
      else set.add(k);
    };
    add(p); add(s);
    return Array.from(set).join(',');
  };

  const getValue = (u, key) => {
    if (!key) return '';
    if (key === 'upnDomain') {
      const upn = (u.userPrincipalName || '').split('@')[1] || '';
      const mail = (u.mail || '').split('@')[1] || '';
      return upn || mail || '';
    }
    if (key.startsWith('extensionAttribute')) {
      const idx = key.replace('extensionAttribute','');
      return u?.onPremisesExtensionAttributes?.[`extensionAttribute${idx}`] || '';
    }
    return u?.[key] || '';
  };

  const buildFilter = (field1, value1, field2, value2) => {
    const escape = (s) => String(s).replace(/'/g, "''");
    const mapField = (f, v) => {
      if (!f) return '';
      if (f.startsWith('extensionAttribute')) {
        const idx = f.replace('extensionAttribute','');
        return `onPremisesExtensionAttributes/extensionAttribute${idx}`;
      }
      if (f === 'upnDomain') {
        const dom = escape(v);
        return { custom: `endswith(userPrincipalName,'@${dom}') or endswith(mail,'@${dom}')` };
      }
      return f;
    };
    const v1 = escape(value1);
    const p1 = mapField(field1, value1);
    let filter = (typeof p1 === 'object' && p1.custom) ? `(${p1.custom})` : `${p1} eq '${v1}'`;
    if (field2 && value2) {
      const v2 = escape(value2);
      const p2 = mapField(field2, value2);
      filter += ' and ' + ((typeof p2 === 'object' && p2.custom) ? `(${p2.custom})` : `${p2} eq '${v2}'`);
    }
    return filter;
  };

  const loadUsersFor = async ({ k1, k2, rowKey }) => {
    if (!authHeader) {
      setRowStates(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey]||{}), error: 'Ingen token. Verifiera först.', loading: false, open: true } }));
      return;
    }
    setRowStates(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey]||{}), error: '', users: [], total: null, loading: true, open: true } }));
    try {
      const header = { Authorization: authHeader, 'Consistency-Level': 'eventual' };
      const select = encodeURIComponent(GRAPH_USER_SELECT_FIELDS.join(','));
      const expand = encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)");
      // Build server-side filter only for supported fields; filter extensionAttributeN locally
      const escape = (s) => String(s).replace(/'/g, "''");
      const clauses = [];
      const clientFilters = [];
      const addServerClause = (f, v) => {
        if (!f || v === undefined || v === null || v === '') return;
        if (f === 'upnDomain') {
          const dom = escape(v);
          clauses.push(`(endswith(userPrincipalName,'@${dom}') or endswith(mail,'@${dom}'))`);
        } else {
          clauses.push(`${f} eq '${escape(v)}'`);
        }
      };
      const addClientExtFilter = (f, v) => {
        const idx = f.replace('extensionAttribute','');
        clientFilters.push((u) => (u?.onPremisesExtensionAttributes?.[`extensionAttribute${idx}`] || '') === v);
      };
      if (primary.startsWith('extensionAttribute')) addClientExtFilter(primary, k1); else addServerClause(primary, k1);
      if (secondary) {
        if (secondary.startsWith('extensionAttribute')) addClientExtFilter(secondary, k2); else if (k2) addServerClause(secondary, k2);
      }
      const hasServerFilter = clauses.length > 0;
      const filter = hasServerFilter ? `&$filter=${encodeURIComponent(clauses.join(' and '))}` : '';
      let url = `https://graph.microsoft.com/v1.0/users?$count=true&$top=50${filter}&$select=${select}&$expand=${expand}`;
      const out = [];
      let guard = 0;
      let filteredCount = 0;
      while (url && guard < 50) {
        guard++;
        const res = await fetch(url, { headers: header });
        if (!res.ok) {
          let msg = `HTTP ${res.status}`;
          try { const j = await res.json(); msg += `: ${j?.error?.message || JSON.stringify(j)}`; } catch {}
          setRowStates(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey]||{}), loading: false, error: msg, open: true } }));
          return;
        }
  const j = await res.json();
        const list = Array.isArray(j.value) ? j.value : [];
        const accepted = clientFilters.length ? list.filter(u => clientFilters.every(fn => fn(u))) : list;
        filteredCount += accepted.length;
        accepted.forEach(u => out.push({
          ...u,
          managerDisplayName: u?.manager?.displayName || '',
          managerMail: u?.manager?.mail || '',
          managerUserPrincipalName: u?.manager?.userPrincipalName || '',
          managerJobTitle: u?.manager?.jobTitle || '',
        }));
        url = j['@odata.nextLink'] || '';
      }
      // Use client filtered count when applicable; else server-provided count if available on first page
      const total = filteredCount > 0 ? filteredCount : out.length;
      // Client-side sort by displayName (sv)
  out.sort((a,b) => (a.displayName || '').localeCompare(b.displayName || '', 'sv'));
      setRowStates(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey]||{}), loading: false, error: '', users: out, total, open: true, usedClientFilter: clientFilters.length > 0 } }));
    } catch (e) {
      setRowStates(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey]||{}), loading: false, error: (e?.message || 'Kunde inte hämta användare'), open: true } }));
    } finally {
      // noop, loading handled per row above
    }
  };

  const onToggleRow = (d) => {
    const key = `${d.k1}||${d.k2 || ''}`;
    setRowStates(prev => {
      const cur = prev[key];
      if (cur && cur.open) return { ...prev, [key]: { ...cur, open: false } };
      const next = { ...(cur || {}), open: true };
      const newState = { ...prev, [key]: next };
      // If no users fetched yet, load
      if (!cur || !Array.isArray(cur.users)) {
        // fire and forget
        setTimeout(() => loadUsersFor({ k1: d.k1, k2: d.k2, rowKey: key }), 0);
      }
      return newState;
    });
  };

  const loadAttributes = async (opts = { loadAll: true, pageLimit: null }) => {
    if (!authHeader) { setError('Ingen token. Verifiera först.'); return; }
    setLoading(true);
    setError('');
    setStatus('Hämtar…');
    abortRef.current = false;
    let url = nextLinkRef.current || `https://graph.microsoft.com/v1.0/users?$select=${encodeURIComponent(getSelectFor(primary, secondary))}&$top=999`;
    const map = new Map(); // key -> count, where key is JSON string of {k1,k2}
    // Seed from cache
    items.forEach(d => {
      const key = JSON.stringify({ k1:d.k1, k2:d.k2||'' });
      map.set(key, d.count);
    });
    let pages = 0;
    let usersSeen = 0;
    try {
      while (url && !abortRef.current) {
        const res = await fetch(url, { headers: { Authorization: authHeader } });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const j = await res.json();
        const list = Array.isArray(j.value) ? j.value : [];
        for (const u of list) {
          const v1 = (getValue(u, primary) || '').trim();
          const v2 = (secondary ? (getValue(u, secondary) || '').trim() : '');
          if (!v1) continue;
          const key = JSON.stringify({ k1:v1, k2:v2 });
          map.set(key, (map.get(key) || 0) + 1);
        }
        usersSeen += list.length;
        pages += 1;
        setStatus(`Hämtat ${usersSeen} användare… (${pages} sidor)`);
        url = j['@odata.nextLink'] || '';
        nextLinkRef.current = url;
        if (!opts.loadAll && opts.pageLimit && pages >= opts.pageLimit) break;
      }
      const arr = Array.from(map.entries()).map(([k, count]) => { const o = JSON.parse(k); return { k1:o.k1, k2:o.k2 || '', count }; });
      arr.sort((a, b) => b.count - a.count || a.k1.localeCompare(b.k1, 'sv') || a.k2.localeCompare(b.k2, 'sv'));
      setItems(arr);
      saveCache(arr);
      setStatus(url ? `Visar delvist resultat. Fler sidor finns…` : `Klar. ${arr.length} värden`);
    } catch (e) {
      setError(e.message || 'Fel vid hämtning');
    } finally {
      setLoading(false);
    }
  };

  const clearCache = () => {
    setItems([]);
    nextLinkRef.current = '';
    setStatus('');
    try { localStorage.removeItem('attr_cache'); } catch {}
  };

  const filtered = useMemo(() => {
    const q = (filter || '').trim().toLowerCase();
    try { localStorage.setItem('attr_filter', filter); } catch {}
    if (!q) return items;
    return items.filter(d => (d.k1.toLowerCase().includes(q) || (d.k2 || '').toLowerCase().includes(q)));
  }, [items, filter]);

  const onChangePrimary = (v) => { setPrimary(v); try { localStorage.setItem('attr_primary', v); } catch {} clearCache(); };
  const onChangeSecondary = (v) => { setSecondary(v); try { localStorage.setItem('attr_secondary', v); } catch {} clearCache(); };

  const labelFor = (k) => FIELD_OPTIONS.find(o => o.key === k)?.label || k;

  return (
    <div>
      <div className="section-header">
        <b>Utforska attribut</b>
        <span className="spacer" />
        <div className="muted">{status}</div>
      </div>
      <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
        <label className="muted">Primärt:</label>
        <select value={primary} onChange={e => onChangePrimary(e.target.value)}>
          {FIELD_OPTIONS.map(o => <option key={o.key} value={o.key}>{o.label}</option>)}
        </select>
        <label className="muted">Sekundärt:</label>
        <select value={secondary} onChange={e => onChangeSecondary(e.target.value)}>
          <option value="">(inget)</option>
          {FIELD_OPTIONS.map(o => <option key={o.key} value={o.key}>{o.label}</option>)}
        </select>
        <button className="btn btn-primary" disabled={loading} onClick={() => loadAttributes({ loadAll: true })}>
          {loading ? 'Hämtar…' : (items.length ? 'Ladda fler / fortsätt' : 'Ladda värden')}
        </button>
        <button className="btn btn-light" onClick={clearCache} disabled={loading || items.length===0}>Rensa</button>
        <input type="text" placeholder="Filtrera…" value={filter} onChange={e => setFilter(e.target.value)} />
      </div>
      {error && <div className="error" style={{ marginTop:8 }}>{error}</div>}
      {items.length > 0 && (
        <div style={{ marginTop:12 }}>
          <div className="muted" style={{ marginBottom:6 }}>Visar {filtered.length} av {items.length} värden</div>
          <div style={{ display:'grid', gridTemplateColumns:`minmax(200px,1fr) ${secondary? 'minmax(180px,1fr) ' : ''}120px 160px`, gap:'8px 12px', alignItems:'center' }}>
            <div className="muted" style={{ fontWeight:600 }}>{labelFor(primary)}</div>
            {secondary && <div className="muted" style={{ fontWeight:600 }}>{labelFor(secondary)}</div>}
            <div className="muted" style={{ fontWeight:600 }}>Antal</div>
            <div />
            {filtered.map(d => {
              const rowKey = `${d.k1}||${d.k2}`;
              const rs = rowStates[rowKey] || {};
              return (
              <React.Fragment key={rowKey}>
                <div>{d.k1}</div>
                {secondary && <div>{d.k2 || '—'}</div>}
                <div className="muted">{d.count}</div>
                <div>
                  <button className="btn btn-secondary" disabled={rs.loading} onClick={() => onToggleRow(d)}>{rs.open ? 'Dölj' : 'Visa användare'}</button>
                </div>
                {rs.open && (
                  <div style={{ gridColumn: '1 / -1' }}>
                    <div className="muted" style={{ margin: '6px 0' }}>{labelFor(primary)}: {d.k1}{secondary ? `, ${labelFor(secondary)}: ${d.k2 || '—'}` : ''} · {rs.loading ? 'Hämtar…' : (typeof rs.total === 'number' ? `${rs.total} träffar` : '')}</div>
                    {rs.usedClientFilter && (
                      <div className="muted" style={{ marginBottom:8 }}>Obs: filtrering på extensionAttribute sker klient-side.</div>
                    )}
                    {rs.error && <div className="error" style={{ marginBottom:8 }}>{rs.error}</div>}
                    {Array.isArray(rs.users) && rs.users.length > 0 && (
                      <UserList users={rs.users} token={token} mode="table" fieldsKey={`attr:${primary}:${secondary || 'none'}:${rowKey}`} />
                    )}
                    {Array.isArray(rs.users) && rs.users.length === 0 && !rs.loading && !rs.error && (
                      <div className="muted" style={{ marginBottom:8 }}>Inga användare.</div>
                    )}
                  </div>
                )}
              </React.Fragment>
            );})}
          </div>
        </div>
      )}
  {items.length === 0 && !loading && (
        <div className="muted" style={{ marginTop:8 }}>Välj attribut och klicka "Ladda värden" för att läsa in unika värden. Hämtning sker i bakgrunden med paginering.</div>
      )}
    </div>
  );
}
