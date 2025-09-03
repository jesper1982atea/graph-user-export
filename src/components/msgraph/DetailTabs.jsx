import React from 'react';

export default function DetailTabs({ tabs, active, onChange }) {
  return (
    <div className="tabbar" style={{ display:'flex', gap:8, borderBottom:'1px solid var(--border)', marginBottom:12 }}>
      {(tabs || []).map(t => (
        <button
          key={t.key}
          className={`btn btn-light ${active===t.key?'active':''}`}
          onClick={() => onChange && onChange(t.key)}
          title={t.title}
        >{t.title}</button>
      ))}
    </div>
  );
}
