import React, { useEffect, useMemo, useState } from 'react';

// Simple in-memory cache per session to avoid refetching
const chatAvatarCache = new Map(); // chatId -> { memberIds: string[], photoUrls: string[] }

/**
 * ChatAvatar
 * Renders a chat avatar that approximates Teams:
 * - oneOnOne: the other participant's profile photo
 * - group: a 2x2 mosaic of up to 4 member photos, or initials fallback
 */
export default function ChatAvatar({ chat, token, meId, size = 40, style }) {
  const [photoUrls, setPhotoUrls] = useState([]);
  const [loading, setLoading] = useState(true);

  const authHeader = useMemo(() => {
    if (!token) return '';
    const trimmed = token.trim();
    return trimmed.toLowerCase().startsWith('bearer ') ? trimmed : `Bearer ${trimmed}`;
  }, [token]);

  useEffect(() => {
    let disposed = false;
    if (!chat?.id || !authHeader) { setLoading(false); return; }

    const fromCache = chatAvatarCache.get(chat.id);
    if (fromCache && Array.isArray(fromCache.photoUrls) && fromCache.photoUrls.length) {
      setPhotoUrls(fromCache.photoUrls);
      setLoading(false);
      return;
    }

    const fetchAll = async () => {
      try {
        setLoading(true);
        // 0) Try native chat photo first (if available in Graph; fails gracefully if not supported)
        try {
          const cp = await fetch(`https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chat.id)}/photo/$value`, {
            headers: { Authorization: authHeader },
          });
          if (cp.ok) {
            const blob = await cp.blob();
            const url = URL.createObjectURL(blob);
            if (!disposed) {
              setPhotoUrls([url]);
              chatAvatarCache.set(chat.id, { memberIds: [], photoUrls: [url] });
              setLoading(false);
              return;
            }
          }
        } catch {}
        // 1) Get members of the chat
        const memRes = await fetch(`https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chat.id)}/members`, {
          headers: { Authorization: authHeader },
        });
        if (!memRes.ok) throw new Error('members fetch failed');
        const memData = await memRes.json();
        const members = Array.isArray(memData.value) ? memData.value : [];
        // Map to user identifiers; exclude me
        const ids = members
          .map(m => m.email || m.userId || m.id)
          .filter(Boolean)
          .filter(uid => !meId || uid !== meId);
        // oneOnOne: prefer the other participant first
        let candidateIds = ids;
        if (chat.chatType === 'oneOnOne') {
          candidateIds = ids.slice(0, 1);
        } else {
          // group: take up to 4 members
          candidateIds = Array.from(new Set(ids)).slice(0, 4);
        }

        // 2) Fetch photos
        const urls = [];
        await Promise.all(candidateIds.map(async (uid) => {
          try {
            const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(uid)}/photo/$value`, {
              headers: { Authorization: authHeader },
            });
            if (!res.ok) return;
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            urls.push(url);
          } catch {}
        }));
        if (!disposed) {
          setPhotoUrls(urls);
          chatAvatarCache.set(chat.id, { memberIds: candidateIds, photoUrls: urls });
        }
      } catch {
        if (!disposed) setPhotoUrls([]);
      } finally {
        if (!disposed) setLoading(false);
      }
    };
    fetchAll();
    return () => { disposed = true; };
  }, [chat?.id, chat?.chatType, authHeader, meId]);

  const initials = useMemo(() => {
    const topic = (chat?.topic || '').trim();
    if (topic) {
      const parts = topic.split(/\s+/).filter(Boolean);
      const a = (parts[0] || '').slice(0, 1).toUpperCase();
      const b = (parts[1] || '').slice(0, 1).toUpperCase();
      return (a + b).trim() || 'GC';
    }
    return 'GC';
  }, [chat?.topic]);

  const s = size;
  const tile = Math.floor(s / 2) - 1; // small overlap border

  if ((photoUrls.length === 1 && chat?.chatType === 'oneOnOne') || (photoUrls.length === 1 && s <= 32)) {
    return (
      <img
        src={photoUrls[0]}
        alt=""
        width={s}
        height={s}
        className="avatar"
        style={{ width: s, height: s, borderRadius: '50%', objectFit: 'cover', ...style }}
      />
    );
  }

  if (photoUrls.length > 1) {
    // 2x2 mosaic
    return (
      <div style={{ position: 'relative', width: s, height: s, borderRadius: 8, overflow: 'hidden', ...style }} aria-label="Gruppfoto">
        {photoUrls.slice(0, 4).map((u, idx) => {
          const row = Math.floor(idx / 2);
          const col = idx % 2;
          return (
            <img
              key={u + idx}
              src={u}
              alt=""
              style={{
                position: 'absolute',
                left: col * (s / 2),
                top: row * (s / 2),
                width: s / 2,
                height: s / 2,
                objectFit: 'cover',
                border: '1px solid var(--card-bg)'
              }}
            />
          );
        })}
      </div>
    );
  }

  // Fallback with initials (approximation of Teams group avatar)
  return (
    <div
      style={{
        width: s,
        height: s,
        borderRadius: '50%',
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        background: 'linear-gradient(135deg, #6366f1 0%, #22c55e 100%)',
        color: 'white',
        fontWeight: 700,
        fontSize: Math.max(10, Math.floor(s * 0.4)),
        boxShadow: '0 0 0 2px var(--card-bg)'
      }}
    >
      {initials}
    </div>
  );
}
