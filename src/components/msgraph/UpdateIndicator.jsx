import React, { useEffect, useState } from 'react';
import { APP_VERSION } from './version';
import { checkForUpdate } from './updateChecker';

export default function UpdateIndicator({ openSettings }) {
  const [info, setInfo] = useState(null);
  const [open, setOpen] = useState(false);
  const [dismissed, setDismissed] = useState(() => {
    try { return localStorage.getItem('update_dismissed_version') || ''; } catch { return ''; }
  });

  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const res = await checkForUpdate(APP_VERSION);
        if (!mounted) return;
        setInfo(res);
      } catch {
        // ignore
      }
    })();
    return () => { mounted = false; };
  }, []);

  if (!info) return null;
  if (dismissed && info.latest && dismissed === info.latest) return null;
  if (!info.hasUpdate && !info.unknownVersion) return null;

  const label = info.hasUpdate ? `Ny version: v${info.latest}` : 'Ny version kan finnas';

  const dismiss = () => {
    try { localStorage.setItem('update_dismissed_version', info.latest || 'unknown'); } catch {}
    setDismissed(info.latest || 'unknown');
    setOpen(false);
  };

  return (
    <div style={{ position: 'relative', display: 'inline-block' }}>
      <button
        className="btn btn-ghost control-40"
        title={label}
        aria-label={label}
        onClick={() => setOpen(v => !v)}
        style={{ position:'relative' }}
      >
        {/* Simple arrow-up-circle icon */}
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <circle cx="12" cy="12" r="10" />
          <polyline points="16 12 12 8 8 12" />
          <line x1="12" y1="16" x2="12" y2="8" />
        </svg>
        {/* Dot indicator */}
        <span style={{ position:'absolute', top:4, right:4, width:8, height:8, borderRadius:'50%', background: info.hasUpdate ? '#ef4444' : '#f59e0b' }} />
      </button>
      {open && (
        <div style={{ position:'absolute', right:0, top:'110%', background:'var(--card-bg)', border:'1px solid var(--border)', borderRadius:10, boxShadow:'0 8px 30px rgba(0,0,0,.12)', padding:10, minWidth:280, zIndex:30 }}>
          <div style={{ fontWeight:600, marginBottom:4 }}>{info.hasUpdate ? 'Ny version tillgänglig' : 'Ny version kan finnas'}</div>
          <div className="muted" style={{ marginBottom:8 }}>
            Din: {APP_VERSION}{info.hasUpdate ? ` → Senaste: v${info.latest}` : ''}
          </div>
          <div style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
            {openSettings && <button className="btn btn-light" onClick={() => { setOpen(false); openSettings(); }}>Öppna i appen</button>}
            {info.assetUrl && <a className="btn btn-secondary" href={info.assetUrl} target="_blank" rel="noopener noreferrer">Ladda ner ZIP</a>}
            {info.releaseHtmlUrl && <a className="btn btn-ghost" href={info.releaseHtmlUrl} target="_blank" rel="noopener noreferrer">Mer info</a>}
            <button className="btn btn-ghost" onClick={dismiss}>Dölj</button>
          </div>
        </div>
      )}
    </div>
  );
}
