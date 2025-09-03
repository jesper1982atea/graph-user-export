import React, { useEffect, useState } from 'react';
import { APP_VERSION } from './version';
import { checkForUpdate } from './updateChecker';

export default function UpdateBanner() {
  const [info, setInfo] = useState(null);
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
  // If version unknown (API blocked), show a subtle banner offering direct download without claiming an update
  if (!info.hasUpdate && info.unknownVersion) {
    return (
      <div style={{ background:'#0b5', color:'#fff', borderRadius:8, padding:'10px 12px', display:'flex', alignItems:'center', gap:10 }}>
        <span style={{ fontWeight:600 }}>Ny version kan finnas</span>
        <span>Din version: {APP_VERSION}. GitHub API är begränsat, men du kan hämta senaste ZIP direkt.</span>
        <div className="spacer" />
        <a className="btn btn-light" href={info.assetUrl} target="_blank" rel="noopener noreferrer">Ladda ner ZIP</a>
        <a className="btn btn-ghost" href={info.releaseHtmlUrl} target="_blank" rel="noopener noreferrer">Visa release</a>
        <button className="btn btn-ghost" onClick={() => setDismissed('unknown')}>Dölj</button>
      </div>
    );
  }
  if (!info.hasUpdate) return null;
  if (dismissed && dismissed === info.latest) return null;

  const dismiss = () => {
    try { localStorage.setItem('update_dismissed_version', info.latest); } catch {}
    setDismissed(info.latest);
  };

  return (
    <div style={{ background:'#0b5', color:'#fff', borderRadius:8, padding:'10px 12px', display:'flex', alignItems:'center', gap:10 }}>
      <span style={{ fontWeight:600 }}>Ny version tillgänglig</span>
      <span>Din version: {APP_VERSION} → Senaste: v{info.latest}</span>
      <div className="spacer" />
      <a className="btn btn-light" href={info.assetUrl} target="_blank" rel="noopener noreferrer">Ladda ner ZIP</a>
      <a className="btn btn-ghost" href={info.releaseHtmlUrl} target="_blank" rel="noopener noreferrer">Visa release</a>
      <button className="btn btn-ghost" onClick={dismiss}>Dölj</button>
    </div>
  );
}
