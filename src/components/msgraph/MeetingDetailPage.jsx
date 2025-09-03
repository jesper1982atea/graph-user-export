import React, { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import UserCard from './UserCard';
import UserList from './UserList';
import CsvExportControls from './CsvExportControls';
import { USER_FIELDS } from './userFields';
import { GRAPH_USER_SELECT_FIELDS } from './graphUserSelect';
import { findFirstAvailable, createMeeting } from './meetingScheduler';
import MeetingSchedulerPanel from './MeetingSchedulerPanel';
import DetailTabs from './DetailTabs';
import BulkEmailPanel from './BulkEmailPanel';

function MeetingDetailPage({ meetings, token }) {
  const { meetingId } = useParams();
  const navigate = useNavigate();
  const [tokenVerified, setTokenVerified] = useState(false);
  const [loading, setLoading] = useState(true);
  const [attendees, setAttendees] = useState([]);
  const [loadingAttendees, setLoadingAttendees] = useState(true);
  const [photos, setPhotos] = useState({});
  // Scheduling state
  const [scheduling, setScheduling] = useState(false);
  const [scheduleDuration, setScheduleDuration] = useState(30);
  const [scheduleWindowStart, setScheduleWindowStart] = useState('');
  const [scheduleWindowEnd, setScheduleWindowEnd] = useState('');
  const [scheduleError, setScheduleError] = useState('');
  const [scheduleResult, setScheduleResult] = useState(null);
  const [scheduleSubject, setScheduleSubject] = useState('');
  const [scheduleBody, setScheduleBody] = useState('');
  const [scheduleOnline, setScheduleOnline] = useState(true);
  const [meUpn, setMeUpn] = useState('');
  const [attendeesViewMode, setAttendeesViewMode] = useState(() => {
    try { return localStorage.getItem('meeting_attendees_view_mode') || 'cards'; } catch { return 'cards'; }
  });
  const setViewMode = (m) => { setAttendeesViewMode(m); try { localStorage.setItem('meeting_attendees_view_mode', m); } catch {} };
  const [activeTab, setActiveTab] = useState(() => {
    try { return localStorage.getItem(`meeting:${meetingId}:active_tab`) || 'overview'; } catch { return 'overview'; }
  });
  useEffect(() => { try { localStorage.setItem(`meeting:${meetingId}:active_tab`, activeTab); } catch {} }, [meetingId, activeTab]);

  useEffect(() => {
    const verifyToken = async () => {
      if (!token) return;
      try {
        const res = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (res.ok) {
          const me = await res.json().catch(() => null);
          setMeUpn(me?.userPrincipalName || '');
          setTokenVerified(true);
        }
      } catch {
        setTokenVerified(false);
      } finally {
        setLoading(false);
      }
    };
    verifyToken();
  }, [token]);

  const meeting = meetings.find(m => m.id === meetingId);

  // Initialize default scheduling subject when meeting is known
  useEffect(() => {
    if (!meeting) return;
    const subj = (meeting.subject || '').trim();
    setScheduleSubject(subj ? `Uppföljning: ${subj}` : 'Uppföljningsmöte');
  }, [meeting]);

  useEffect(() => {
    if (!meeting || !token) return;
    // Fetch attendee info for each attendee
    const fetchAttendees = async () => {
      setLoadingAttendees(true);
  const selectFields = GRAPH_USER_SELECT_FIELDS.join(',');
      const attendeeEmails = (meeting.attendees || []).map(a => a.emailAddress?.address).filter(Boolean);
      const promises = attendeeEmails.map(async (email) => {
        try {
          const res = await fetch(`https://graph.microsoft.com/v1.0/users/${email}?$select=${selectFields}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (!res.ok) return null;
          const data = await res.json();
          // Try to fetch photo
          let photoUrl = '';
          try {
            const photoRes = await fetch(`https://graph.microsoft.com/v1.0/users/${email}/photo/$value`, {
              headers: { Authorization: `Bearer ${token}` },
            });
            if (photoRes.ok) {
              const blob = await photoRes.blob();
              photoUrl = URL.createObjectURL(blob);
            }
          } catch {}
          return {
            ...data,
            photoUrl,
            managerDisplayName: data?.manager?.displayName || '',
            managerMail: data?.manager?.mail || '',
            managerUserPrincipalName: data?.manager?.userPrincipalName || '',
            managerJobTitle: data?.manager?.jobTitle || '',
          };
        } catch {
          return null;
        }
      });
      const results = await Promise.all(promises);
      setAttendees(results.filter(Boolean));
      setLoadingAttendees(false);
    };
    fetchAttendees();
  }, [meeting, token]);

  // CSV export handled via CsvExportControls

  if (loading) return <div className="card">Laddar...</div>;
  if (!tokenVerified) return <div className="card">Ogiltig token. Gå tillbaka och logga in igen.</div>;
  if (!meeting) return <div className="card">Möte hittades inte.</div>;

  return (
    <div className="card" style={{ maxWidth: 1000, margin: '2rem auto' }}>
      <div className="section-title" style={{ marginBottom: 8 }}>
        <b>Detaljer för möte</b>
        <span className="spacer" />
        <button className="btn btn-light" onClick={() => navigate(-1)}>&larr; Tillbaka</button>
      </div>
      <div className="muted" style={{ marginBottom: 12 }}>
        <div><b>Ämne:</b> {meeting.subject}</div>
        <div><b>Start:</b> {meeting.start?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
        <div><b>Slut:</b> {meeting.end?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
        <div><b>Plats:</b> {meeting.location?.displayName || '-'}</div>
        {meeting.onlineMeeting?.joinUrl && (
          <div><b>Länk:</b> <a href={meeting.onlineMeeting.joinUrl} target="_blank" rel="noopener noreferrer">{meeting.onlineMeeting.joinUrl}</a></div>
        )}
      </div>
      <DetailTabs
        tabs={[
          { key:'overview', title:'Översikt' },
          { key:'attendees', title:`Deltagare (${attendees.length})` },
          { key:'schedule', title:'Mötesbokning' },
          { key:'email', title:'E‑post' },
        ]}
        active={activeTab}
        onChange={setActiveTab}
      />

      {activeTab === 'overview' && (
        <div className="muted" style={{ marginBottom: 12 }}>
          <div><b>Ämne:</b> {meeting.subject}</div>
          <div><b>Start:</b> {meeting.start?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
          <div><b>Slut:</b> {meeting.end?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
          <div><b>Plats:</b> {meeting.location?.displayName || '-'}</div>
          {meeting.onlineMeeting?.joinUrl && (
            <div><b>Länk:</b> <a href={meeting.onlineMeeting.joinUrl} target="_blank" rel="noopener noreferrer">{meeting.onlineMeeting.joinUrl}</a></div>
          )}
        </div>
      )}

      {activeTab === 'schedule' && (
        <MeetingSchedulerPanel
          token={token}
          attendeeEmails={(meeting.attendees || []).map(a => a.emailAddress?.address).filter(Boolean)}
          defaultSubject={((meeting.subject || '').trim() ? `Uppföljning: ${meeting.subject}` : 'Uppföljningsmöte')}
          defaultBody={scheduleBody}
          defaultOnline={true}
          defaultDurationMinutes={30}
          defaultWindowStart={(() => { const endIso = meeting.end?.dateTime; let s = endIso ? new Date(endIso) : new Date(); s = new Date(s.getTime() + 15*60*1000); return s.toISOString(); })()}
          defaultWindowEnd={(() => { const endIso = meeting.end?.dateTime; const e = (endIso ? new Date(endIso).getTime() : Date.now()) + 14*24*60*60*1000; return new Date(e).toISOString(); })()}
          tenantDomain={(meUpn || '').split('@')[1] || undefined}
          contextKey={`meeting:${meeting.id}:schedule`}
          title="Boka uppföljning"
        />
      )}

      {activeTab === 'email' && (
        <BulkEmailPanel
          token={token}
          emails={(meeting.attendees || []).map(a => a.emailAddress?.address).filter(Boolean)}
          defaultSubject={((meeting.subject || '').trim() ? `Uppföljning: ${meeting.subject}` : 'Uppföljning')}
          contextKey={`meeting:${meeting.id}:email`}
        />
      )}
      {activeTab === 'attendees' && (
        <>
          <div className="section-header">
            <b>Deltagare ({attendees.length})</b>
            <div className="spacer" />
            <div className="muted" style={{ display:'inline-flex', gap:6, alignItems:'center', marginRight:8 }}>
              <span>Visning:</span>
              <button className={`btn btn-light ${attendeesViewMode==='table'?'active':''}`} onClick={() => setViewMode('table')}>Tabell</button>
              <button className={`btn btn-light ${attendeesViewMode==='cards'?'active':''}`} onClick={() => setViewMode('cards')}>Kort</button>
            </div>
            <CsvExportControls items={attendees} storageKey={`meeting:${meeting.id}:attendees`} defaultFileName={`mote_${meeting.id}_deltagare.csv`} buttonLabel="Exportera CSV" />
          </div>
          {loadingAttendees ? (
            <div className="muted">Laddar deltagare...</div>
          ) : (
            <div style={{ marginTop:12 }}>
              {attendeesViewMode === 'table' ? (
                <UserList
                  users={attendees}
                  token={token}
                  mode={'table'}
                  fieldsKey={`meeting:${meeting.id}:attendees`}
                  selectedFields={(() => { try { return JSON.parse(localStorage.getItem('userlist_fields_meeting')||'[]'); } catch { return []; } })()}
                  onChangeSelectedFields={(f) => { try { localStorage.setItem('userlist_fields_meeting', JSON.stringify(f)); } catch {} }}
                />
              ) : (
                <div style={{ display:'flex', flexWrap:'wrap', gap:12 }}>
                  {attendees.map(u => (
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

export default MeetingDetailPage;
