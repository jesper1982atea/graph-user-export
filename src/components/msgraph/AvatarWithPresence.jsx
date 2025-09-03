import React from 'react';

function presenceColor(p) {
  const v = (p || '').toLowerCase();
  if (!v) return '#9ca3af'; // gray
  if (v.includes('dnd') || v.includes('donotdisturb')) return '#7c3aed'; // purple
  if (v.includes('busy')) return '#ef4444'; // red
  if (v.includes('away') || v.includes('brb') || v.includes('idle')) return '#f59e0b'; // yellow
  if (v.includes('available') || v.includes('online')) return '#10b981'; // green
  if (v.includes('offline')) return '#9ca3af'; // gray
  return '#9ca3af';
}

export default function AvatarWithPresence({
  photoUrl,
  name,
  size = 48,
  presence,
  showDot = true,
}) {
  const url = photoUrl || `https://ui-avatars.com/api/?name=${encodeURIComponent(name || 'U')}&background=0ea5e9&color=fff&size=${size*2}`;
  const ring = presenceColor(presence);
  const dotSize = Math.max(8, Math.round(size * 0.22));
  return (
  <div style={{ position:'relative', width:size, height:size, borderRadius:'50%', display:'inline-block', boxShadow:'0 0 0 2px #fff' }} title={presence || ''}>
      <div style={{
        width:size, height:size, borderRadius:'50%',
        padding:2, background: ring,
        boxShadow: `0 0 0 2px ${ring} inset`,
      }}>
        <img
          src={url}
          alt={name || ''}
          style={{ width:'100%', height:'100%', borderRadius:'50%', objectFit:'cover', background:'#fff' }}
          onError={e => { e.target.onerror = null; e.target.src = `https://ui-avatars.com/api/?name=${encodeURIComponent((name||'U')[0])}&background=0ea5e9&color=fff&size=${size*2}`; }}
        />
      </div>
      {showDot && (
        <span style={{
          position:'absolute', right:0, bottom:0,
          width:dotSize, height:dotSize, borderRadius:'50%',
          background:ring, border:'2px solid #fff',
        }} />
      )}
    </div>
  );
}
