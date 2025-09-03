import React, { useEffect, useMemo, useState } from 'react';
import AvatarWithPresence from './AvatarWithPresence';

// A polished contact card for a single user: avatar + presence, primary info, and quick actions.
// Props:
// - user: Graph user-like object
// - token: Graph token (optional, for photo/presence fetch)
// - onStartChat?: (user) => void
// - onScheduleMeeting?: (user) => void
// - onOpenDetails?: (user) => void (defaults to navigate #/users/{id|upn})
export default function UserCard({ user, token, onStartChat, onScheduleMeeting, onOpenDetails }) {
  if (!user) return null;

  const [photoUrl, setPhotoUrl] = useState(user.photoUrl || '');
  const [presence, setPresence] = useState('');

  const Authorization = useMemo(() => {
    if (!token) return '';
    const t = token.trim();
    return t.toLowerCase().startsWith('bearer ') ? t : `Bearer ${t}`;
  }, [token]);

  // Choose best identifier: id (GUID) -> userPrincipalName/mail -> fallback id
  const userKey = useMemo(() => {
    return user.id || user.userPrincipalName || user.mail || '';
  }, [user]);

  // Fetch photo if not provided
  useEffect(() => {
    let cancelled = false;
    if (!Authorization || photoUrl || !userKey) return;
    const run = async () => {
      try {
        // Try direct photo
        const p1 = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userKey)}/photo/$value`, { headers: { Authorization } });
        if (p1.ok) {
          const blob = await p1.blob();
          if (!cancelled) setPhotoUrl(URL.createObjectURL(blob));
          return;
        }
        // Fallback: resolve GUID then fetch
        const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userKey)}?$select=id`, { headers: { Authorization } });
        if (r.ok) {
          const j = await r.json();
          if (j?.id) {
            const p2 = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(j.id)}/photo/$value`, { headers: { Authorization } });
            if (p2.ok) {
              const blob = await p2.blob();
              if (!cancelled) setPhotoUrl(URL.createObjectURL(blob));
            }
          }
        }
      } catch {}
    };
    run();
    return () => { cancelled = true; };
  }, [Authorization, userKey, photoUrl]);

  // Fetch presence if token available
  useEffect(() => {
    let cancelled = false;
    if (!Authorization || !userKey) return;
    const run = async () => {
      try {
        const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userKey)}/presence`, { headers: { Authorization } });
        if (r.ok) {
          const j = await r.json();
          if (!cancelled) setPresence(j.availability || j.activity || '');
        }
      } catch {}
    };
    run();
    return () => { cancelled = true; };
  }, [Authorization, userKey]);

  const name = user.displayName || user.mail || user.userPrincipalName || '(okänd)';
  const email = user.mail || user.userPrincipalName || '';
  const title = [user.jobTitle, user.companyName].filter(Boolean).join(' · ');

  const handleOpenDetails = () => {
    if (onOpenDetails) return onOpenDetails(user);
    try { window.location.hash = `#/users/${encodeURIComponent(user.id || user.userPrincipalName || user.mail || '')}`; } catch {}
  };

  const copyEmail = async (e) => {
    e.stopPropagation();
    try { await navigator.clipboard.writeText(email); } catch {}
  };

  const ActionButton = ({ title, onClick, children }) => (
    <button className="btn btn-light" title={title} onClick={(e) => { e.stopPropagation(); onClick && onClick(); }}>
      {children}
    </button>
  );

  return (
    <div className="card" style={{ width: 360, cursor:'pointer' }} onClick={handleOpenDetails}>
      <div style={{ display:'flex', gap:12 }}>
        <AvatarWithPresence photoUrl={photoUrl} name={name} presence={presence} size={64} />
        <div style={{ minWidth:0, flex:1 }}>
          <div className="grid-item-title" style={{ lineHeight:1.15 }}>{name}</div>
          {email && <div className="grid-item-sub" style={{ wordBreak:'break-all' }}>{email}</div>}
          {title && <div className="grid-item-meta">{title}</div>}
        </div>
      </div>

      <div style={{ marginTop:10, display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
        {user.department && (
          <div className="muted">
            <b>Avdelning:</b>{' '}
            <button
              className="linklike"
              onClick={(e) => {
                e.stopPropagation();
                try { window.location.hash = `#/search?department=${encodeURIComponent(user.department)}`; } catch {}
              }}
              title="Visa alla användare i avdelningen"
              style={{ background:'none', border:'none', color:'var(--link)', cursor:'pointer', padding:0 }}
            >
              {user.department}
            </button>
          </div>
        )}
        {user.officeLocation && (
          <div className="muted"><b>Plats:</b> {user.officeLocation}</div>
        )}
        {user.mobilePhone && (
          <div className="muted"><b>Mobil:</b> {user.mobilePhone}</div>
        )}
        {user.country && (
          <div className="muted"><b>Land:</b> {user.country}</div>
        )}
        {user.managerDisplayName && (
          <div className="muted" style={{ gridColumn:'1 / -1' }}><b>Chef:</b> {user.managerDisplayName}</div>
        )}
      </div>

      <div style={{ marginTop:12, display:'flex', gap:8, flexWrap:'wrap' }}>
        {email && (
          <>
            <ActionButton title="Kopiera e‑post" onClick={() => navigator.clipboard.writeText(email)}>📋 Kopiera</ActionButton>
            <a
              className="btn btn-light"
              href={`mailto:${encodeURIComponent(email)}`}
              onClick={(e) => e.stopPropagation()}
              title="Skicka e‑post"
            >✉️ E‑post</a>
          </>
        )}
        {typeof onStartChat === 'function' && (
          <ActionButton title="Starta chatt" onClick={() => onStartChat(user)}>💬 Chatt</ActionButton>
        )}
        {typeof onScheduleMeeting === 'function' && (
          <ActionButton title="Boka möte" onClick={() => onScheduleMeeting(user)}>📅 Möte</ActionButton>
        )}
        <ActionButton title="Visa detaljer" onClick={handleOpenDetails}>🔎 Detaljer</ActionButton>
      </div>
    </div>
  );
}
