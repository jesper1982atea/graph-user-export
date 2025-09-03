import React, { useEffect, useState, useMemo } from 'react';
import AvatarWithPresence from './AvatarWithPresence';

// Renders two overlapping avatars for a 1-1 chat. Falls back to initials.
export default function OneOnOnePairAvatars({ a, b, token, size = 36, myPhotoUrl, overlapRatio = 0.72 }) {
  const [photos, setPhotos] = useState({}); // { key: objectUrl }
  const [presenceMap, setPresenceMap] = useState({}); // { key: presence }
  const Authorization = useMemo(() => {
    const t = (token || '').trim();
    return t.toLowerCase().startsWith('bearer ') ? t : `Bearer ${t}`;
  }, [token]);

  const users = useMemo(() => {
    // Prefer real user identifiers over membership ids from /chats/{id}/members
    // Order: userId (AAD GUID) -> email/UPN -> mail -> fallback id
    const toKey = (u) => u?.userId || u?.email || u?.userPrincipalName || u?.mail || u?.id || '';
    const keyA = toKey(a);
    const keyB = toKey(b);
    return [
      { key: keyA, name: a?.displayName || a?.name || keyA, preferUrl: myPhotoUrl },
      { key: keyB, name: b?.displayName || b?.name || keyB },
    ];
  }, [a, b, myPhotoUrl]);

  useEffect(() => {
    let cancelled = false;
    const run = async () => {
      const next = {};
      await Promise.all(users.map(async (u) => {
        if (!u.key) return;
        if (u.preferUrl) { next[u.key] = u.preferUrl; return; }
        try {
          let ok = false;
          // First try direct photo by key (UPN or GUID)
          try {
            const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(u.key)}/photo/$value`, { headers: { Authorization } });
            if (res.ok) {
              const blob = await res.blob();
              next[u.key] = URL.createObjectURL(blob);
              ok = true;
            }
          } catch {}
          if (!ok) {
            // Fallback: resolve GUID by key then fetch photo
            try {
              const r2 = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(u.key)}?$select=id`, { headers: { Authorization } });
              if (r2.ok) {
                const j = await r2.json();
                if (j?.id) {
                  const r3 = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(j.id)}/photo/$value`, { headers: { Authorization } });
                  if (r3.ok) {
                    const blob = await r3.blob();
                    next[u.key] = URL.createObjectURL(blob);
                    ok = true;
                  }
                }
              }
            } catch {}
          }
        } catch {}
      }));
      if (!cancelled) setPhotos(next);
    };
    run();
    return () => { cancelled = true; };
  }, [Authorization, users]);

  // Fetch presence for both users; batch with GUIDs, fallback to single GET for non-GUID keys
  useEffect(() => {
    const guidRe = /^[0-9a-fA-F-]{36}$/;
    let cancelled = false;
    const run = async () => {
      const keys = users.map(u => u.key).filter(Boolean);
      const ids = [];
      const upns = [];
      keys.forEach(k => guidRe.test(k) ? ids.push(k) : upns.push(k));
      const map = {};
      // Resolve UPNs to GUIDs to leverage batch when possible
      const resolvedMap = {}; // originalKey -> guid
      await Promise.all(upns.map(async (k) => {
        try {
          const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(k)}?$select=id`, { headers: { Authorization } });
          if (r.ok) {
            const j = await r.json();
            if (j?.id) { resolvedMap[k] = j.id; ids.push(j.id); }
          }
        } catch {}
      }));

      // Batch by GUIDs (combined direct GUIDs + resolved UPNs)
      if (ids.length) {
        try {
          const res = await fetch('https://graph.microsoft.com/v1.0/communications/getPresencesByUserId', {
            method: 'POST', headers: { Authorization, 'Content-Type': 'application/json' }, body: JSON.stringify({ ids })
          });
          if (res.ok) {
            const data = await res.json();
            (data.value || []).forEach(p => {
              // Map back to original keys where applicable
              const origKey = Object.keys(resolvedMap).find(k => resolvedMap[k] === p.id) || p.id;
              map[origKey] = p.availability || p.activity || 'offline';
            });
          }
        } catch {}
      }

      // As final fallback, GET presence for any remaining original keys missing from map
      const missing = keys.filter(k => map[k] == null);
      await Promise.all(missing.map(async (k) => {
        try {
          const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(k)}/presence`, { headers: { Authorization } });
          if (r.ok) {
            const j = await r.json();
            map[k] = j.availability || j.activity || 'offline';
          }
        } catch {}
      }));

      if (!cancelled) setPresenceMap(map);
    };
    run();
    return () => { cancelled = true; };
  }, [Authorization, users]);

  const circle = (url, name, presence, z) => (
    <div key={name + z} style={{ position:'relative', width: size, height: size, zIndex: z }}>
      <AvatarWithPresence photoUrl={url} name={name} presence={presence} size={size} showDot={Boolean(presence)} />
    </div>
  );

  const [u1, u2] = users;
  const url1 = photos[u1.key] || u1.preferUrl || '';
  const url2 = photos[u2.key] || '';
  const p1 = presenceMap[u1.key] || '';
  const p2 = presenceMap[u2.key] || '';

  const offset = Math.max(4, Math.round(size * overlapRatio));
  return (
    <div style={{ position:'relative', width: size + offset, height: size }}>
      <div style={{ position:'absolute', left: 0, top: 0, filter:'drop-shadow(0 0 0 var(--card-bg))' }}>
        <div style={{ position:'relative' }}>
          <div style={{ position:'absolute', inset:0, borderRadius:'50%', boxShadow:'0 0 0 2px #fff' }} />
          {circle(url1, u1.name, p1, 2)}
        </div>
      </div>
      <div style={{ position:'absolute', left: offset, top: 0 }}>
        <div style={{ position:'relative' }}>
          <div style={{ position:'absolute', inset:0, borderRadius:'50%', boxShadow:'0 0 0 2px #fff' }} />
          {circle(url2, u2.name, p2, 1)}
        </div>
      </div>
    </div>
  );
}
