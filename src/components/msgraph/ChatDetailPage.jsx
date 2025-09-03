// src/components/msgraph/ChatDetailPage.jsx

import React, { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import UserCard from './UserCard';
import UserList from './UserList';
import CsvExportControls from './CsvExportControls';
import { USER_FIELDS } from './userFields';
import ChatAvatar from './ChatAvatar';
import OneOnOnePairAvatars from './OneOnOnePairAvatars';
import { findFirstAvailable, createMeeting } from './meetingScheduler';
import MeetingSchedulerPanel from './MeetingSchedulerPanel';
import DetailTabs from './DetailTabs';
import BulkEmailPanel from './BulkEmailPanel';
import { GRAPH_USER_SELECT_FIELDS } from './graphUserSelect';

function ChatDetailPage({ chats, token, meId, meUpn, mePhotoUrl }) {
  const { chatId } = useParams();
  const navigate = useNavigate();
  const [members, setMembers] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(true);
  const [photos, setPhotos] = useState({});
  // Scheduling state
  const [scheduling, setScheduling] = useState(false);
  const [scheduleDuration, setScheduleDuration] = useState(30);
  const [scheduleWindowStart, setScheduleWindowStart] = useState('');
  const [scheduleWindowEnd, setScheduleWindowEnd] = useState('');
  const [scheduleError, setScheduleError] = useState('');
  const [scheduleResult, setScheduleResult] = useState(null);
  const [membersViewMode, setMembersViewMode] = useState(() => {
    try { return localStorage.getItem('chat_members_view_mode') || 'cards'; } catch { return 'cards'; }
  });
  const setViewMode = (m) => { setMembersViewMode(m); try { localStorage.setItem('chat_members_view_mode', m); } catch {} };
  const [activeTab, setActiveTab] = useState(() => {
    try { return localStorage.getItem(`chat:${chatId}:active_tab`) || 'overview'; } catch { return 'overview'; }
  });
  useEffect(() => { try { localStorage.setItem(`chat:${chatId}:active_tab`, activeTab); } catch {} }, [chatId, activeTab]);

  const chat = chats.find(c => c.id === chatId);

  useEffect(() => {
    if (!chat || !token) return;
    setLoadingMembers(true);
    fetch(`https://graph.microsoft.com/v1.0/chats/${chatId}/members`, {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then(res => res.ok ? res.json() : Promise.reject('Kunde inte hämta medlemmar'))
      .then(data => {
        setMembers(data.value || []);
        setLoadingMembers(false);
      })
      .catch(() => {
        setMembers([]);
        setLoadingMembers(false);
      });
  }, [chat, chatId, token]);


  useEffect(() => {
    if (!token || members.length === 0) return;
    let isMounted = true;
    const fetchPhotos = async () => {
      const newPhotos = {};
      await Promise.all(members.map(async (m) => {
        const userId = m.email || m.userId || m.id;
        if (!userId) return;
        try {
          const res = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (res.ok) {
            const blob = await res.blob();
            newPhotos[userId] = URL.createObjectURL(blob);
          }
        } catch {}
      }));
      if (isMounted) setPhotos(newPhotos);
    };
    fetchPhotos();
    return () => { isMounted = false; };
  }, [token, members]);

  // Enrich members with full user profile for selected fields
  const [enrichedUsers, setEnrichedUsers] = useState([]);
  useEffect(() => {
    if (!token || members.length === 0) { setEnrichedUsers([]); return; }
    let cancelled = false;
    const Authorization = (token || '').trim().toLowerCase().startsWith('bearer ') ? token : `Bearer ${token}`;
    const run = async () => {
      const select = encodeURIComponent(GRAPH_USER_SELECT_FIELDS.join(','));
      const results = await Promise.all(members.map(async (m) => {
        const uid = m.userId || m.id || m.email; // prefer GUID userId when available
        const key = m.email || m.userId || m.id;
        const base = { displayName: m.displayName, mail: m.email, userPrincipalName: m.email, id: uid, photoUrl: key && photos[key] };
        if (!uid) return base;
        try {
          const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(uid)}?$select=${select}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`, { headers: { Authorization } });
          if (!res.ok) return base;
          const data = await res.json();
          return {
            ...data,
            photoUrl: base.photoUrl,
            managerDisplayName: data?.manager?.displayName || '',
            managerMail: data?.manager?.mail || '',
            managerUserPrincipalName: data?.manager?.userPrincipalName || '',
            managerJobTitle: data?.manager?.jobTitle || '',
          };
        } catch { return base; }
      }));
      if (!cancelled) setEnrichedUsers(results);
    };
    run();
    return () => { cancelled = true; };
  }, [token, members, photos]);

  // CSV export handled via CsvExportControls

  if (!chat) return <div className="card">Chatt hittades inte.</div>;

  const prettySubtitle = () => {
    if (chat.chatType === 'oneOnOne') {
      const others = members
        .map(m => ({ name: m.displayName, id: m.email || m.userId || m.id }))
        .filter(m => m.id && m.id !== meId);
      if (others.length) return others.map(o => o.name || o.id).join(', ');
      return '';
    }
    if (members.length) return `${chat.chatType} · ${members.length} deltagare`;
    return chat.chatType || '';
  };

  const resolveMemberEmails = async () => {
    const emails = [];
    const needLookup = [];
    members.forEach(m => {
      const mail = m.email || m.mail;
      if (mail) emails.push(mail);
      else if (m.userId) needLookup.push(m.userId);
    });
    if (!needLookup.length) return emails.filter(Boolean);
    // resolve missing via /users/{id}?$select=mail,userPrincipalName
    const resolved = await Promise.all(needLookup.map(async (uid) => {
      try {
        const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(uid)}?$select=mail,userPrincipalName`, {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (!res.ok) return null;
        const u = await res.json();
        return u.mail || u.userPrincipalName || null;
      } catch { return null; }
    }));
    return emails.concat(resolved.filter(Boolean));
  };

  const isOne = chat.chatType === 'oneOnOne';
  const selfMember = members.find(m => (m.userId || m.id) === meId || m.email === meUpn);
  const others = members.filter(m => ((m.userId || m.id) !== meId) && (m.email !== meUpn));
  const pairA = isOne ? (selfMember || members[0]) : null;
  const pairB = isOne ? (others[0] || members[1]) : null;

  return (
    <div className="card" style={{ maxWidth: 1000, margin: '2rem auto' }}>
      <div className="section-title" style={{ marginBottom: 8, display:'flex', alignItems:'center', gap:10 }}>
        {isOne && pairA && pairB ? (
          <OneOnOnePairAvatars a={pairA} b={pairB} token={token} size={44} overlapRatio={1.2} myPhotoUrl={mePhotoUrl} />
        ) : (
          <ChatAvatar chat={chat} token={token} meId={meId} size={44} />
        )}
        <div style={{ display:'flex', flexDirection:'column' }}>
          <b>{isOne ? (prettySubtitle() || '1‑1 chatt') : 'Detaljer för chatt'}</b>
          <span className="muted" style={{ fontSize:'.9rem' }}>{isOne ? '1‑1 chatt' : ((chat.topic || '').trim() || 'Ingen rubrik')} {isOne ? '' : `· ${prettySubtitle()}`}</span>
        </div>
        <span className="spacer" />
        <button className="btn btn-light" onClick={() => navigate(-1)}>&larr; Tillbaka</button>
      </div>
      <DetailTabs
        tabs={[
          { key:'overview', title:'Översikt' },
          { key:'members', title:`Medlemmar (${members.length})` },
          { key:'schedule', title:'Mötesbokning' },
          { key:'email', title:'E‑post' },
        ]}
        active={activeTab}
        onChange={setActiveTab}
      />

      {activeTab === 'overview' && (
        <div className="muted" style={{ marginBottom: 12 }}>
          <div><b>Chatt-id:</b> {chat.id}</div>
          {!isOne && <div><b>Rubrik:</b> {(chat.topic || '').trim() || 'Ingen rubrik'}</div>}
          <div><b>Info:</b> {prettySubtitle()}</div>
        </div>
      )}

      {activeTab === 'schedule' && members.length > 0 && (
        <MeetingSchedulerPanel
          token={token}
          getAttendeeEmails={resolveMemberEmails}
          defaultSubject={chat.topic || 'Möte'}
          defaultOnline={true}
          defaultDurationMinutes={30}
          tenantDomain={(meUpn || '').split('@')[1] || undefined}
          contextKey={`chat:${chat.id}:schedule`}
          title="Mötesbokning"
        />
      )}

      {activeTab === 'email' && (
        <BulkEmailPanel
          token={token}
          getEmails={resolveMemberEmails}
          defaultSubject={(chat.topic || 'Meddelande')}
          contextKey={`chat:${chat.id}:email`}
        />
      )}

      {activeTab === 'members' && (
        <>
          <div className="section-header">
            <b>Medlemmar ({members.length})</b>
            <div className="spacer" />
            <div className="muted" style={{ display:'inline-flex', gap:6, alignItems:'center', marginRight:8 }}>
              <span>Visning:</span>
              <button className={`btn btn-light ${membersViewMode==='table'?'active':''}`} onClick={() => setViewMode('table')}>Tabell</button>
              <button className={`btn btn-light ${membersViewMode==='cards'?'active':''}`} onClick={() => setViewMode('cards')}>Kort</button>
            </div>
            <CsvExportControls items={enrichedUsers} storageKey={`chat:${chat.id}:members`} defaultFileName={`chatt_${chat.id}_medlemmar.csv`} buttonLabel="Exportera CSV" />
          </div>
          {loadingMembers ? (
            <div className="muted">Laddar medlemmar...</div>
          ) : (
            <div style={{ marginTop:12 }}>
              {membersViewMode === 'table' ? (
                <UserList
                  users={enrichedUsers}
                  token={token}
                  mode={'table'}
                  fieldsKey={`chat:${chat.id}:members`}
                  selectedFields={(() => { try { return JSON.parse(localStorage.getItem('userlist_fields_chat')||'[]'); } catch { return []; } })()}
                  onChangeSelectedFields={(f) => { try { localStorage.setItem('userlist_fields_chat', JSON.stringify(f)); } catch {} }}
                />
              ) : (
                <div style={{ display:'flex', flexWrap:'wrap', gap:12 }}>
                  {enrichedUsers.map(u => (
                    <UserCard key={u.id || u.userPrincipalName || u.mail}
                      user={u}
                      token={token}
                      onOpenDetails={() => { try { window.location.hash = `#/users/${encodeURIComponent(u.id || u.userPrincipalName || u.mail)}`; } catch {} }}
                    />
                  ))}
                </div>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default ChatDetailPage;