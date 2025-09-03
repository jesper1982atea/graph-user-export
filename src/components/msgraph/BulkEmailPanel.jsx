import React, { useEffect, useState } from 'react';

/**
 * BulkEmailPanel - compose emails to a group (all at once via BCC) or individually.
 * Props:
 * - emails?: string[]
 * - getEmails?: () => Promise<string[]>
 * - defaultSubject?: string
 * - defaultBody?: string
 * - contextKey?: string (for localStorage persistence)
 * - title?: string
 */
export default function BulkEmailPanel({
  token,
  emails,
  getEmails,
  defaultSubject = '',
  defaultBody = '',
  contextKey = 'email:default',
  title = 'E-post till deltagare',
}) {
  const sk = (s) => `${contextKey}:${s}`;
  const [subject, setSubject] = useState(() => lsGet(sk('subject'), defaultSubject));
  const [body, setBody] = useState(() => lsGet(sk('body'), defaultBody));
  const [busy, setBusy] = useState(false);
  const [lastInfo, setLastInfo] = useState('');
  const [isHtml, setIsHtml] = useState(() => lsGet(sk('isHtml'), '1') === '1');
  const [progress, setProgress] = useState({ done: 0, total: 0, errors: 0 });
  const [scopes, setScopes] = useState([]);
  const [hasMailSend, setHasMailSend] = useState(false);

  useEffect(() => { lsSet(sk('subject'), subject); }, [subject]);
  useEffect(() => { lsSet(sk('body'), body); }, [body]);
  useEffect(() => { lsSet(sk('isHtml'), isHtml ? '1' : '0'); }, [isHtml]);
  useEffect(() => {
    try {
      const s = token ? readTokenScopes(token) : [];
      setScopes(s);
      setHasMailSend(s.some(x => x.toLowerCase() === 'mail.send'));
    } catch { setScopes([]); setHasMailSend(false); }
  }, [token]);

  const resolveEmails = async () => {
    if (Array.isArray(emails) && emails.length) return dedupeEmails(emails);
    if (typeof getEmails === 'function') {
      try { const arr = await getEmails(); return dedupeEmails(arr || []); } catch { return []; }
    }
    return [];
  };

  const sendAllBcc = async () => {
    setBusy(true); setLastInfo(''); setProgress({ done: 0, total: 0, errors: 0 });
    try {
      const list = await resolveEmails();
      if (!list.length) { setLastInfo('Inga e‑postadresser.'); return; }
      if (!token || !hasMailSend) { setLastInfo('Saknar behörighet Mail.Send. Öppnar klient istället.'); openMailClientBcc(list); return; }
      const Authorization = buildAuth(token);
      // Chunk to avoid recipient limits (e.g., 100 per mail)
      const chunks = chunk(list, 90);
      setProgress({ done: 0, total: chunks.length, errors: 0 });
      let errors = 0;
      let lastErr = '';
      for (let i = 0; i < chunks.length; i++) {
        const bcc = chunks[i].map(addr => ({ emailAddress: { address: addr } }));
        const payload = {
          message: {
            subject: subject || '',
            body: { contentType: isHtml ? 'HTML' : 'Text', content: body || '' },
            toRecipients: [],
            bccRecipients: bcc,
          },
          saveToSentItems: true,
        };
        const { ok, status, text } = await postSendMail(Authorization, payload);
        if (!ok && (status === 401 || status === 403)) {
          setLastInfo('Token saknar rättighet att skicka (Mail.Send). Öppnar klient istället.');
          openMailClientBcc(list);
          return;
        }
        if (!ok) { errors += 1; lastErr = text || `HTTP ${status}`; }
        setProgress({ done: i + 1, total: chunks.length, errors });
        await delay(250);
      }
      setLastInfo(errors ? `Skickat med ${errors} fel (${chunks.length - errors}/${chunks.length}). ${hintFromStatus(lastErr)}` : `Skickat ${chunks.length} utskick.`);
    } finally { setBusy(false); }
  };

  const sendIndividually = async () => {
    setBusy(true); setLastInfo(''); setProgress({ done: 0, total: 0, errors: 0 });
    try {
      const list = await resolveEmails();
      if (!list.length) { setLastInfo('Inga e‑postadresser.'); return; }
      if (!token || !hasMailSend) { setLastInfo('Saknar behörighet Mail.Send. Öppnar klient istället.'); openMailClientIndividually(list); return; }
      const Authorization = buildAuth(token);
      setProgress({ done: 0, total: list.length, errors: 0 });
      let errors = 0;
      let lastErr = '';
      for (let i = 0; i < list.length; i++) {
        const addr = list[i];
        const payload = {
          message: {
            subject: subject || '',
            body: { contentType: isHtml ? 'HTML' : 'Text', content: body || '' },
            toRecipients: [{ emailAddress: { address: addr } }],
          },
          saveToSentItems: true,
        };
        const { ok, status, text } = await postSendMail(Authorization, payload);
        if (!ok && (status === 401 || status === 403)) {
          setLastInfo('Token saknar rättighet att skicka (Mail.Send). Öppnar klient istället.');
          openMailClientIndividually(list);
          return;
        }
        if (!ok) { errors += 1; lastErr = text || `HTTP ${status}`; }
        setProgress({ done: i + 1, total: list.length, errors });
        await delay(150);
      }
      setLastInfo(errors ? `Skickat med ${errors} fel (${list.length - errors}/${list.length}). ${hintFromStatus(lastErr)}` : `Skickade ${list.length} e‑post.`);
    } finally { setBusy(false); }
  };

  const openMailClientBcc = (list) => {
    const emails = dedupeEmails(list || []);
    if (!emails.length) { setLastInfo('Inga e‑postadresser.'); return; }
    const url = buildMailtoUrl({ bcc: emails, subject, body: isHtml ? stripTags(body) : body });
    window.location.href = url;
  };
  const openMailClientIndividually = (list) => {
    const emails = dedupeEmails(list || []);
    if (!emails.length) { setLastInfo('Inga e‑postadresser.'); return; }
    // Open the first one to avoid popup blockers; advise user to repeat for the rest if many.
    const first = emails[0];
    const url = buildMailtoUrl({ to: [first], subject, body: isHtml ? stripTags(body) : body });
    window.location.href = url;
    if (emails.length > 1) setLastInfo(`Öppnade första utkastet i klienten. Totalt ${emails.length} mottagare.`);
  };

  return (
    <div className="card" style={{ marginTop: 8 }}>
      <div className="section-header" style={{ gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        <b>{title}</b>
        <div className="spacer" />
        <label className="muted">Ämne:</label>
        <input type="text" value={subject} onChange={e => setSubject(e.target.value)} style={{ minWidth: 220 }} />
        <label className="muted" title="Skicka som HTML">
          <input type="checkbox" checked={isHtml} onChange={e => setIsHtml(e.target.checked)} /> HTML
        </label>
        {hasMailSend ? (
          <>
            <button className="btn btn-secondary" disabled={busy} onClick={sendAllBcc}>Skicka alla (bcc)</button>
            <button className="btn btn-light" disabled={busy} onClick={sendIndividually}>Skicka individuellt</button>
          </>
        ) : (
          <>
            <button className="btn btn-secondary" disabled={busy} onClick={async ()=>openMailClientBcc(await resolveEmails())}>Öppna i klient (bcc)</button>
            <button className="btn btn-light" disabled={busy} onClick={async ()=>openMailClientIndividually(await resolveEmails())}>Öppna i klient (individ.)</button>
          </>
        )}
        {progress.total > 0 && (
          <span className="muted">{progress.done}/{progress.total} · fel: {progress.errors}</span>
        )}
        {lastInfo && <span className="muted">{lastInfo}</span>}
      </div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center', marginBottom: 12 }}>
        <label className="muted">Meddelande ({isHtml ? 'HTML' : 'Text'}):</label>
        <textarea value={body} onChange={e => setBody(e.target.value)} rows={3} style={{ flex: 1, minWidth: 300 }} placeholder="Valfritt meddelande" />
      </div>
      <div className="muted" style={{ fontSize: '.9rem' }}>
        {hasMailSend ? 'Skickar via Microsoft Graph (/me/sendMail).' : 'Saknar Mail.Send—öppnar e‑postklient via mailto.'}
        {Array.isArray(scopes) && scopes.length > 0 && (
          <>
            {' '}Behörigheter i token: {scopes.join(', ')}
          </>
        )}
      </div>
    </div>
  );
}

function lsGet(k, f) { try { const v = localStorage.getItem(k); return v == null ? f : v; } catch { return f; } }
function lsSet(k, v) { try { localStorage.setItem(k, v); } catch {} }
function dedupeEmails(arr) {
  const s = new Set();
  const out = [];
  (arr || []).forEach(e => { const v = (e || '').trim().toLowerCase(); if (v && !s.has(v)) { s.add(v); out.push(v); } });
  return out;
}

function buildAuth(token) {
  const t = (token || '').trim();
  return t.toLowerCase().startsWith('bearer ') ? t : `Bearer ${t}`;
}
async function postSendMail(Authorization, payload) {
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: { Authorization, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (res.status === 202) return { ok: true, status: res.status };
    const text = await res.text().catch(() => '');
    return { ok: false, status: res.status, text };
  } catch (e) { return { ok: false, status: 0, text: String(e || '') }; }
}
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
function hintFromStatus(text) {
  if (!text) return '';
  try {
    const obj = JSON.parse(text);
    const code = obj?.error?.code || '';
    if (String(text).includes('AccessDenied') || code === 'ErrorAccessDenied') return 'Behörighet saknas (Mail.Send?).';
    if (String(text).includes('InvalidRecipients')) return 'Ogiltiga mottagare i listan.';
    if (String(text).includes('QuotaExceeded')) return 'Skickbegränsning nådd.';
  } catch {}
  return '';
}

// Token scope decoding and mailto helpers
function readTokenScopes(token) {
  try {
    const raw = (token || '').trim();
    const t = raw.toLowerCase().startsWith('bearer ') ? raw.slice(7) : raw;
    const parts = t.split('.');
    if (parts.length < 2) return [];
    const payload = JSON.parse(base64UrlDecode(parts[1]) || '{}');
    // AAD adds scopes in 'scp' (space-delimited) for delegated tokens; 'roles' for app perms
    const scp = typeof payload.scp === 'string' ? payload.scp.split(' ').filter(Boolean) : [];
    const roles = Array.isArray(payload.roles) ? payload.roles : [];
    return [...scp, ...roles];
  } catch { return []; }
}
function base64UrlDecode(s) {
  try {
    const pad = (str) => str + '='.repeat((4 - (str.length % 4)) % 4);
    const b64 = pad(s.replace(/-/g, '+').replace(/_/g, '/'));
    if (typeof atob === 'function') return decodeURIComponent(Array.prototype.map.call(atob(b64), c => '%'+('00'+c.charCodeAt(0).toString(16)).slice(-2)).join(''));
    // Node fallback
    return Buffer.from(b64, 'base64').toString('utf-8');
  } catch { return ''; }
}
function stripTags(html) {
  try { return String(html || '').replace(/<[^>]*>/g, ''); } catch { return String(html || ''); }
}
function buildMailtoUrl({ to = [], cc = [], bcc = [], subject = '', body = '' }) {
  const params = new URLSearchParams();
  if (cc.length) params.set('cc', cc.join(','));
  if (bcc.length) params.set('bcc', fitRecipients(bcc));
  if (subject) params.set('subject', subject);
  if (body) params.set('body', body);
  const toStr = to.join(',');
  const url = `mailto:${encodeURIComponent(toStr)}?${params.toString()}`;
  // Try to keep under ~2000 chars for broad client support.
  return url.length > 2000 ? `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}` : url;
}
function fitRecipients(list) {
  // Ensure the resulting URL likely stays <2000; trim recipients if too long
  const joined = list.join(',');
  if (joined.length < 1200) return joined; // leave space for subject/body
  let out = [];
  let len = 0;
  for (const r of list) {
    if (len + r.length + 1 > 1200) break;
    out.push(r); len += r.length + 1;
  }
  return out.join(',');
}
