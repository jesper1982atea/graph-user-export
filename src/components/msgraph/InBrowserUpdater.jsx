import React, { useCallback, useState } from 'react';
import JSZip from 'jszip';
import { checkForUpdate } from './updateChecker';
import { APP_VERSION } from './version';

export default function InBrowserUpdater() {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [files, setFiles] = useState([]);
  const [info, setInfo] = useState(null);

  const handleZipBuffer = useCallback(async (blobOrBuffer) => {
    setStatus('Packar upp…');
    const zip = await JSZip.loadAsync(blobOrBuffer);
    const entries = [];
    const promises = [];
    zip.forEach((relativePath, zipEntry) => {
      if (zipEntry.dir) return;
      const p = zipEntry.async('blob').then(b => {
        entries.push({ path: relativePath, blob: b });
      });
      promises.push(p);
    });
    await Promise.all(promises);
    setFiles(entries);
    setStatus(`Klar. ${entries.length} filer redo att sparas manuellt.`);
  }, []);

  const run = async () => {
    try {
      setBusy(true);
      setStatus('Söker efter ny version…');
  const upd = await checkForUpdate(APP_VERSION);
      setInfo(upd);
      const url = upd.assetUrl;
      if (!upd.hasUpdate && !upd.unknownVersion) { setStatus('Du har redan senaste versionen.'); setBusy(false); return; }
      setStatus('Hämtar ZIP…');
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error('Kunde inte ladda ner ZIP');
      const blob = await res.blob();
      await handleZipBuffer(blob);
    } catch (e) {
      // Vanligt orsakat av CORS på GitHubs asset‑domän. Erbjud fallback.
      setStatus('Kunde inte ladda ner (ofta p.g.a. CORS). Använd Direktlänk (ZIP) och välj filen nedan.');
    } finally {
      setBusy(false);
    }
  };

  const onChooseFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      await handleZipBuffer(file);
    } catch (err) {
      setStatus('Kunde inte läsa ZIP‑filen.');
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="card">
      <b>Uppdatera från GitHub</b>
      <div className="muted" style={{ marginTop:6 }}>Hämta senaste dist‑ZIP och extrahera i webbläsaren. Spara sedan filerna manuellt till din lokala mapp.</div>
      <div style={{ marginTop:10, display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
        <button className="btn btn-secondary" onClick={run} disabled={busy}>Sök och hämta</button>
        <a className="btn btn-light" href={`https://github.com/jesper1982atea/graph-user-export/releases/latest/download/graph-user-export-dist.zip`} target="_blank" rel="noopener noreferrer">Direktlänk (ZIP)</a>
        <label className="btn btn-light" style={{ cursor:'pointer' }}>
          Välj ZIP…
          <input type="file" accept=".zip" onChange={onChooseFile} style={{ display:'none' }} />
        </label>
        {status && <span className="muted">{status}</span>}
      </div>
      {info?.hasUpdate && (
        <div style={{ marginTop:8 }} className="muted">Ny version: v{info.latest}. Din: {APP_VERSION}.</div>
      )}
      {files.length > 0 && (
        <div style={{ marginTop:12 }}>
          <div style={{ marginBottom:6, fontWeight:600 }}>Filer</div>
          <ul style={{ maxHeight:220, overflow:'auto', border:'1px solid var(--border)', borderRadius:8, padding:8 }}>
            {files.slice(0, 200).map(f => (
              <li key={f.path} style={{ display:'flex', alignItems:'center', gap:8 }}>
                <code style={{ background:'var(--bg)', padding:'2px 6px', borderRadius:6 }}>{f.path}</code>
                <a className="btn btn-light" href={URL.createObjectURL(f.blob)} download={f.path}>Spara</a>
              </li>
            ))}
          </ul>
          {files.length > 200 && <div className="muted">…visar första 200</div>}
          <div className="muted" style={{ marginTop:8 }}>Tips: Spara filerna till din lokala `dist/` eller motsvarande mapp för att uppdatera.</div>
        </div>
      )}
    </div>
  );
}
