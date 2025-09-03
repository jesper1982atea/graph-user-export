import React, { useEffect, useMemo, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import AvatarWithPresence from './AvatarWithPresence';
import { GRAPH_USER_SELECT_FIELDS } from './graphUserSelect';
import { normalizeUser } from './userNormalize';

export default function UserDetailPage({ token }) {
  const { userId } = useParams();
  const navigate = useNavigate();
  const [user, setUser] = useState(null);
  const [photoUrl, setPhotoUrl] = useState('');
  const [presence, setPresence] = useState('');
  const [showPlans, setShowPlans] = useState(false);
  const [showProvisioned, setShowProvisioned] = useState(false);
  const [showDevices, setShowDevices] = useState(false);
  const Authorization = useMemo(() => (token || '').trim().toLowerCase().startsWith('bearer ') ? token : `Bearer ${token}`, [token]);

  useEffect(() => {
    if (!token || !userId) return;
    const load = async () => {
      try {
        const select = encodeURIComponent(GRAPH_USER_SELECT_FIELDS.join(','));
        const expandMgr = encodeURIComponent('manager($select=displayName,mail,userPrincipalName,jobTitle)');
        const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}?$select=${select}&$expand=${expandMgr}`, { headers: { Authorization } });
        if (!res.ok) throw new Error('Kunde inte hämta användare');
        const data = await res.json();
        setUser(normalizeUser(data));
      } catch {}
      try {
        const pr = await fetch('https://graph.microsoft.com/v1.0/communications/getPresencesByUserId', {
          method: 'POST', headers: { Authorization, 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: [userId] })
        });
        if (pr.ok) {
          const p = await pr.json();
          const v = (p.value && p.value[0]) || {};
          setPresence(v.availability || v.activity || '');
        }
      } catch {}
      try {
        const ph = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/photo/$value`, { headers: { Authorization } });
        if (ph.ok) { const blob = await ph.blob(); setPhotoUrl(URL.createObjectURL(blob)); }
      } catch {}
    };
    load();
  }, [token, userId, Authorization]);

  if (!user) return <div className="card">Laddar användare…</div>;

  const isEmpty = (v) => v == null || (typeof v === 'string' && v.trim() === '') || (Array.isArray(v) && v.length === 0) || (typeof v === 'object' && Object.keys(v || {}).length === 0);
  const chip = (label, value) => (!isEmpty(value) ? (
    <span key={label} className="badge badge-neutral" title={label} style={{ textTransform:'none' }}>{value}</span>
  ) : null);

  const identity = [
    ['ID', user.id],
    ['UPN', user.userPrincipalName],
    ['E-post', user.mail],
    ['SAM', user.onPremisesSamAccountName],
    ['On-Prem UPN', user.onPremisesUserPrincipalName],
    ['Employee ID', user.employeeId],
  ].filter(([,v]) => !isEmpty(v));

  const contact = [
    ['Mobil', user.mobilePhone],
    ['Telefon', Array.isArray(user.businessPhones) ? user.businessPhones.join(', ') : user.businessPhones],
  ].filter(([,v]) => !isEmpty(v));

  const org = [
    ['Titel', user.jobTitle],
    ['Avdelning', user.department],
    ['Företag', user.companyName],
    ['Kontor', user.officeLocation],
  ].filter(([,v]) => !isEmpty(v));

  const address = [
    ['Gatuadress', user.streetAddress],
    ['Postnummer', user.postalCode],
    ['Stad', user.city],
    ['Land', user.country],
  ].filter(([,v]) => !isEmpty(v));

  const extAttrs = Array.from({ length: 15 }, (_, i) => {
    const key = `extensionAttribute${i+1}`;
    return [key, user[key]];
  }).filter(([,v]) => !isEmpty(v));

  const hasAssigned = Array.isArray(user.assignedPlans) && user.assignedPlans.length > 0;
  const hasProvisioned = Array.isArray(user.provisionedPlans) && user.provisionedPlans.length > 0;
  const hasDevices = Array.isArray(user.managedDevices) && user.managedDevices.length > 0;

  return (
    <div className="card" style={{ maxWidth: 1100, margin: '2rem auto' }}>
      <div className="section-title" style={{ marginBottom: 8 }}>
        <b>Användardetaljer</b>
        <span className="spacer" />
        <button className="btn btn-light" onClick={() => navigate(-1)}>&larr; Tillbaka</button>
      </div>

      <div style={{ display:'flex', gap:16, alignItems:'center', marginBottom:12 }}>
        <AvatarWithPresence photoUrl={photoUrl} name={user.displayName || user.mail || user.userPrincipalName} presence={presence} size={64} />
        <div style={{ minWidth:0 }}>
          <div className="grid-item-title" title={user.displayName}>{user.displayName}</div>
          <div className="grid-item-sub">{user.mail || user.userPrincipalName}</div>
          {!isEmpty(user.jobTitle) && <div className="grid-item-meta">{user.jobTitle}</div>}
          <div style={{ display:'flex', gap:6, flexWrap:'wrap', marginTop:6 }}>
            {!isEmpty(user.department) ? (
              <button
                className="badge badge-neutral"
                onClick={() => { try { window.location.hash = `#/search?department=${encodeURIComponent(user.department)}`; } catch {} }}
                title="Visa alla användare i avdelningen"
                style={{ cursor:'pointer' }}
              >{user.department}</button>
            ) : null}
            {chip('Företag', user.companyName)}
            {chip('Kontor', user.officeLocation)}
            {chip('Stad', user.city)}
            {chip('Land', user.country)}
          </div>
        </div>
      </div>

      {/* Manager */}
      {!isEmpty(user.managerDisplayName) && (
        <div className="card" style={{ marginBottom:12 }}>
          <div className="section-header">
            <b>Chef</b>
            <span className="spacer" />
            <button className="btn btn-light" onClick={() => {
              const idOrUpn = user.managerUserPrincipalName || user.managerMail;
              if (idOrUpn) {
                try { window.location.hash = `#/users/${encodeURIComponent(idOrUpn)}`; } catch {}
              }
            }}>Visa chef</button>
          </div>
          <div className="muted" style={{ display:'flex', gap:16, flexWrap:'wrap' }}>
            <span><b>Namn:</b> {user.managerDisplayName}</span>
            {!isEmpty(user.managerMail) && <span><b>E-post:</b> {user.managerMail}</span>}
            {!isEmpty(user.managerUserPrincipalName) && <span><b>UPN:</b> {user.managerUserPrincipalName}</span>}
            {!isEmpty(user.managerJobTitle) && <span><b>Titel:</b> {user.managerJobTitle}</span>}
          </div>
        </div>
      )}

      {/* Details grid */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit, minmax(260px, 1fr))', gap:12 }}>
        {identity.length > 0 && (
          <div className="card">
            <div className="section-title"><b>Identitet</b></div>
            <DetailsList pairs={identity} />
          </div>
        )}
        {contact.length > 0 && (
          <div className="card">
            <div className="section-title"><b>Kontakt</b></div>
            <DetailsList pairs={contact} />
          </div>
        )}
        {org.length > 0 && (
          <div className="card">
            <div className="section-title"><b>Organisation</b></div>
            <DetailsList pairs={org} />
          </div>
        )}
        {address.length > 0 && (
          <div className="card">
            <div className="section-title"><b>Adress</b></div>
            <DetailsList pairs={address} />
          </div>
        )}
      </div>

      {/* Extension attributes */}
      {extAttrs.length > 0 && (
        <div className="card" style={{ marginTop:12 }}>
          <div className="section-title"><b>Extension Attributes</b></div>
          <table className="table" style={{ width:'100%', borderCollapse:'separate', borderSpacing:0 }}>
            <tbody>
              {extAttrs.map(([k,v]) => (
                <tr key={k}>
                  <th style={{ textAlign:'left', whiteSpace:'nowrap' }}>{k}</th>
                  <td className="muted" style={{ fontFamily:'ui-monospace, monospace' }}>{String(v)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Plans & devices */}
      {(hasAssigned || hasProvisioned || hasDevices) && (
        <div className="card" style={{ marginTop:12 }}>
          <div className="section-header">
            <b>Licenser & Enheter</b>
          </div>
          <div className="muted" style={{ display:'flex', gap:12, flexWrap:'wrap' }}>
            {!isEmpty(user.assignedPlansCount) && <span className="badge badge-neutral">Assigned plans: {user.assignedPlansCount}</span>}
            {!isEmpty(user.provisionedPlansCount) && <span className="badge badge-neutral">Provisioned: {user.provisionedPlansCount}</span>}
            {!isEmpty(user.managedDevicesCount) && <span className="badge badge-neutral">Enheter: {user.managedDevicesCount}</span>}
          </div>
          {hasAssigned && (
            <div style={{ marginTop:8 }}>
              <button className="btn btn-light" onClick={() => setShowPlans(v => !v)}>{showPlans ? 'Dölj Assigned plans' : 'Visa Assigned plans'}</button>
              {showPlans && (
                <pre style={{ background:'var(--bg-muted)', padding:8, borderRadius:8, overflow:'auto' }}>{JSON.stringify(user.assignedPlans, null, 2)}</pre>
              )}
            </div>
          )}
          {hasProvisioned && (
            <div style={{ marginTop:8 }}>
              <button className="btn btn-light" onClick={() => setShowProvisioned(v => !v)}>{showProvisioned ? 'Dölj Provisioned plans' : 'Visa Provisioned plans'}</button>
              {showProvisioned && (
                <pre style={{ background:'var(--bg-muted)', padding:8, borderRadius:8, overflow:'auto' }}>{JSON.stringify(user.provisionedPlans, null, 2)}</pre>
              )}
            </div>
          )}
          {hasDevices && (
            <div style={{ marginTop:8 }}>
              <button className="btn btn-light" onClick={() => setShowDevices(v => !v)}>{showDevices ? 'Dölj Enheter' : 'Visa Enheter'}</button>
              {showDevices && (
                <pre style={{ background:'var(--bg-muted)', padding:8, borderRadius:8, overflow:'auto' }}>{JSON.stringify(user.managedDevices, null, 2)}</pre>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function DetailsList({ pairs }) {
  return (
    <div style={{ overflowX:'auto' }}>
      <table className="table" style={{ width:'100%', borderCollapse:'separate', borderSpacing:0 }}>
        <tbody>
          {pairs.map(([k,v]) => (
            <tr key={k}>
              <th style={{ textAlign:'left', whiteSpace:'nowrap' }}>{k}</th>
              <td className="muted" style={{ fontFamily:'ui-monospace, monospace' }}>{String(v)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
