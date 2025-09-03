import React, { useEffect, useMemo, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import CsvExportControls from './CsvExportControls';
import UserList from './UserList';
import UserCard from './UserCard';
import DetailTabs from './DetailTabs';
import MeetingSchedulerPanel from './MeetingSchedulerPanel';
import BulkEmailPanel from './BulkEmailPanel';

export default function TeamDetailPage({ token, teamId: propTeamId, inline = false, onClose }) {
  const { teamId: routeTeamId } = useParams();
  const teamId = propTeamId || routeTeamId;
  const navigate = useNavigate();
  const [team, setTeam] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [channels, setChannels] = useState([]);
  const [channelsLoading, setChannelsLoading] = useState(false);
  const [channelsError, setChannelsError] = useState('');
  const [selectedChannelId, setSelectedChannelId] = useState(null);
  const [selectedChannel, setSelectedChannel] = useState(null);
  const [selectedChannelMembers, setSelectedChannelMembers] = useState([]);
  const [selectedChannelLoading, setSelectedChannelLoading] = useState(false);
  const [selectedChannelError, setSelectedChannelError] = useState('');
  const [channelMembers, setChannelMembers] = useState({}); // { [channelId]: { loading, members } }
  const [members, setMembers] = useState([]);
  const [membersLoading, setMembersLoading] = useState(false);
  const [activeTab, setActiveTab] = useState(() => {
    try {
      const stored = localStorage.getItem(`team:${teamId}:active_tab`);
      if (stored) return stored;
    } catch {}
    return inline ? 'channels' : 'overview';
  });
  useEffect(() => { try { localStorage.setItem(`team:${teamId}:active_tab`, activeTab); } catch {} }, [teamId, activeTab]);

  const Authorization = useMemo(() => token ? ((token.trim().toLowerCase().startsWith('bearer ') ? token : `Bearer ${token}`)) : '', [token]);

  useEffect(() => {
    if (!token || !teamId) return;
    const load = async () => {
      setLoading(true); setError('');
      try {
        // Fetch team basic info (best-effort)
        try {
          const res = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}`, { headers: { Authorization } });
          if (res.ok) setTeam(await res.json());
        } catch {}
        // Channels
        setChannelsLoading(true);
        try {
          let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/`;
          const all = [];
          while (url) {
            const cres = await fetch(url, { headers: { Authorization } });
            if (!cres.ok) {
              try {
                const body = await cres.json();
                const msg = body?.error?.message || JSON.stringify(body);
                setChannelsError(`Kunde inte hämta kanaler (HTTP ${cres.status}): ${msg}`);
              } catch {
                const text = await cres.text();
                setChannelsError(text ? `Kunde inte hämta kanaler (HTTP ${cres.status}): ${text}` : `Kunde inte hämta kanaler (HTTP ${cres.status})`);
              }
              break;
            }
            const data = await cres.json();
            if (Array.isArray(data.value)) all.push(...data.value);
            url = data['@odata.nextLink'] || null;
          }
          setChannels(all);
        } finally { setChannelsLoading(false); }
        // Members
        setMembersLoading(true);
        try {
          const mres = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/members?$top=200`, { headers: { Authorization } });
          const data = mres.ok ? await mres.json() : { value: [] };
          const raw = (data.value || []);
          // Normalize: ensure we have an email when possible
          const withEmail = await Promise.all(raw.map(async (m) => {
            const email = m.email || m.mail || '';
            if (email) return normalizeMember(m, email);
            const userId = m.userId || m.id || '';
            if (!userId) return normalizeMember(m, '');
            try {
              const u = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}?$select=mail,userPrincipalName,displayName,jobTitle,department`, { headers: { Authorization } });
              if (u.ok) {
                const uj = await u.json();
                return normalizeMember(m, uj.mail || uj.userPrincipalName || '');
              }
            } catch {}
            return normalizeMember(m, '');
          }));
          setMembers(withEmail);
        } finally { setMembersLoading(false); }
      } catch (e) {
        setError(e.message || 'Kunde inte läsa team.');
      } finally {
        setLoading(false);
      }
    };
    load();
  }, [token, teamId, Authorization]);

  const toggleChannelMembers = async (channelId) => {
    setChannelMembers(prev => ({ ...prev, [channelId]: { ...(prev[channelId]||{}), loading: !(prev[channelId]?.loading), members: prev[channelId]?.members || [] } }));
    // If already loaded, just toggle visibility
    if (channelMembers[channelId]?.members?.length) {
      setChannelMembers(prev => ({ ...prev, [channelId]: { loading: false, members: prev[channelId].members, open: !(prev[channelId].open) } }));
      return;
    }
    try {
  let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/members/`;
      const all = [];
      while (url) {
        const res = await fetch(url, { headers: { Authorization } });
        if (!res.ok) break;
        const data = await res.json();
        if (Array.isArray(data.value)) all.push(...data.value);
        url = data['@odata.nextLink'] || null;
      }
      setChannelMembers(prev => ({ ...prev, [channelId]: { loading: false, members: all, open: true } }));
    } catch {
      setChannelMembers(prev => ({ ...prev, [channelId]: { loading: false, members: [], open: true } }));
    }
  };

  const getAllMemberEmails = async () => {
    const emails = members.map(m => m.mail).filter(Boolean);
    // Deduplicate
    const set = new Set(emails.map(e => e.toLowerCase()));
    return Array.from(set);
  };

  if (!token) return inline ? <div>Verifiera token.</div> : <div className="card">Verifiera token.</div>;
  if (loading) return inline ? <div>Laddar…</div> : <div className="card">Laddar…</div>;
  if (error) return (
    <div className={inline ? '' : 'card'}>
      <div className="section-title"><b>Team</b><span className="spacer" />
        {inline ? (
          <button className="btn btn-light" onClick={() => onClose && onClose()}>&larr; Till lista</button>
        ) : (
          <button className="btn btn-light" onClick={() => navigate(-1)}>&larr; Tillbaka</button>
        )}
      </div>
      <div className="muted" style={{ color:'#b91c1c' }}>{error}</div>
    </div>
  );

  return (
    <div className={inline ? '' : 'card'} style={inline ? {} : { maxWidth: 1100, margin: '2rem auto' }}>
      <div className="section-title" style={{ marginBottom: 8 }}>
        <b>{team?.displayName || 'Team'}</b>
        <span className="spacer" />
        {inline ? (
          <button className="btn btn-light" onClick={() => onClose && onClose()}>&larr; Till lista</button>
        ) : (
          <button className="btn btn-light" onClick={() => navigate(-1)}>&larr; Tillbaka</button>
        )}
      </div>
      <div className="muted" style={{ marginBottom: 12 }}>
        <div><b>Team-id:</b> {teamId}</div>
      </div>

      <DetailTabs
        tabs={[
          { key:'overview', title:'Översikt' },
          { key:'channels', title:`Kanaler (${channels.length})` },
          { key:'members', title:`Medlemmar (${members.length})` },
          { key:'schedule', title:'Mötesbokning' },
          { key:'email', title:'E‑post' },
        ]}
        active={activeTab}
        onChange={setActiveTab}
      />

      {activeTab === 'overview' && (
        <div className="muted">
          <div><b>Namn:</b> {team?.displayName || '-'}</div>
          <div><b>Beskrivning:</b> {team?.description || '-'}</div>
          <div style={{ marginTop:10 }}>
            <a className="btn btn-light" href={`https://teams.microsoft.com/l/team/${encodeURIComponent(teamId)}`} target="_blank" rel="noopener noreferrer">Öppna i Teams</a>
          </div>
        </div>
      )}

      {activeTab === 'channels' && (
        <div>
          {!selectedChannelId ? (
            <>
              <div className="section-header">
                <b>Kanaler ({channels.length})</b>
                <div className="spacer" />
                <a className="btn btn-light" href={`https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/`} target="_blank" rel="noopener noreferrer">Öppna i Graph</a>
                <CsvExportControls items={channels} storageKey={`team:${teamId}:channels`} defaultFileName={`team_${teamId}_kanaler.csv`} buttonLabel="Exportera kanaler (CSV)" />
              </div>
              {channelsLoading ? (
                <div className="muted">Laddar kanaler…</div>
              ) : channels.length === 0 ? (
                <div>
                  <div className="muted">Inga kanaler.</div>
                  {channelsError ? (
                    <div style={{ marginTop:8, color:'#b91c1c', background:'#fee2e2', border:'1px solid #fecaca', padding:8, borderRadius:8 }}>
                      {channelsError}
                      <div className="muted" style={{ marginTop:6 }}>
                        Tips: Behöver Channel.ReadBasic.All (eller Channel.Read.All) för att lista kanaler i ett team. Om du använder Graph Explorer, logga in och ge samtycke för dessa scopes.
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : (
                <ul style={{ paddingLeft: '1rem' }}>
                  {channels.map(ch => (
                    <li key={ch.id} style={{ marginBottom: 10 }}>
                      <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                        <button className="btn btn-link" style={{ padding:0, fontWeight:700 }} onClick={async () => {
                          setSelectedChannelId(ch.id);
                          setSelectedChannelLoading(true);
                          setSelectedChannelError('');
                          setSelectedChannel(null);
                          setSelectedChannelMembers([]);
                          try {
                            // Channel details
                            const cRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${ch.id}/`, { headers: { Authorization } });
                            if (cRes.ok) setSelectedChannel(await cRes.json());
                            // Members (paged)
                            let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${ch.id}/members/`;
                            const all = [];
                            while (url) {
                              const mRes = await fetch(url, { headers: { Authorization } });
                              if (!mRes.ok) break;
                              const data = await mRes.json();
                              if (Array.isArray(data.value)) all.push(...data.value);
                              url = data['@odata.nextLink'] || null;
                            }
                            setSelectedChannelMembers(all);
                          } catch (e) {
                            setSelectedChannelError(e?.message || 'Kunde inte läsa kanal');
                          } finally {
                            setSelectedChannelLoading(false);
                          }
                        }}>
                          {ch.displayName}
                        </button>
                        <span className="muted">{ch.membershipType || 'standard'}</span>
                        <button className="btn btn-light" onClick={() => toggleChannelMembers(ch.id)}>
                          Visa/Dölj medlemmar
                        </button>
                        <CsvExportControls items={channelMembers?.[ch.id]?.members || []} storageKey={`channel:${ch.id}:members`} defaultFileName={`kanal_${ch.id}_medlemmar.csv`} buttonLabel="Exportera medlemmar" />
                      </div>
                      {ch.description ? (
                        <div className="grid-item-meta" style={{ marginTop:4 }}>{ch.description}</div>
                      ) : null}
                      {channelMembers?.[ch.id]?.loading ? (
                        <div className="muted" style={{ marginTop:4 }}>Laddar medlemmar…</div>
                      ) : channelMembers?.[ch.id]?.open ? (
                        <ul style={{ marginTop:6, paddingLeft: '1rem' }}>
                          {(channelMembers?.[ch.id]?.members || []).map((m, i) => (
                            <li key={m.id || m.userId || i} className="muted">{m.displayName || m.email || m.userId}</li>
                          ))}
                        </ul>
                      ) : null}
                    </li>
                  ))}
                </ul>
              )}
            </>
          ) : (
            <div>
              <div className="section-header">
                <b>Kanal: {selectedChannel?.displayName || selectedChannelId}</b>
                <div className="spacer" />
                <button className="btn btn-light" onClick={() => { setSelectedChannelId(null); setSelectedChannel(null); setSelectedChannelMembers([]); }}>&larr; Till kanaler</button>
              </div>
              {selectedChannelLoading ? (
                <div className="muted">Laddar kanal…</div>
              ) : selectedChannelError ? (
                <div style={{ color:'#b91c1c', background:'#fee2e2', border:'1px solid #fecaca', padding:8, borderRadius:8 }}>{selectedChannelError}</div>
              ) : (
                <div>
                  <div className="muted" style={{ marginBottom:8 }}>
                    <div><b>Typ:</b> {selectedChannel?.membershipType || 'standard'}</div>
                    {selectedChannel?.description ? <div><b>Beskrivning:</b> {selectedChannel.description}</div> : null}
                    <div><b>Kanal‑id:</b> {selectedChannelId}</div>
                  </div>
                  <div className="section-header">
                    <b>Medlemmar ({selectedChannelMembers.length})</b>
                    <div className="spacer" />
                    <a className="btn btn-light" href={`https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(selectedChannelId)}/members/`} target="_blank" rel="noopener noreferrer">Öppna medlemmar i Graph</a>
                    <CsvExportControls items={selectedChannelMembers} storageKey={`channel:${selectedChannelId}:members`} defaultFileName={`kanal_${selectedChannelId}_medlemmar.csv`} buttonLabel="Exportera medlemmar" />
                  </div>
                  {selectedChannelMembers.length === 0 ? (
                    <div className="muted">Inga medlemmar eller saknar rättigheter.</div>
                  ) : (
                    <ul style={{ paddingLeft: '1rem' }}>
                      {selectedChannelMembers.map((m, i) => (
                        <li key={m.id || m.userId || i} className="muted">{m.displayName || m.email || m.userId}</li>
                      ))}
                    </ul>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {activeTab === 'members' && (
        <TeamMembersView token={token} teamId={teamId} members={members} loading={membersLoading} />
      )}

      {activeTab === 'schedule' && (
        <MeetingSchedulerPanel
          token={token}
          getAttendeeEmails={getAllMemberEmails}
          defaultSubject={team?.displayName ? `Möte: ${team.displayName}` : 'Team-möte'}
          defaultBody={''}
          defaultOnline={true}
          defaultDurationMinutes={45}
          defaultWindowStart={new Date().toISOString()}
          defaultWindowEnd={new Date(Date.now() + 14*24*60*60*1000).toISOString()}
          contextKey={`team:${teamId}:schedule`}
          title="Boka team‑möte"
        />
      )}

      {activeTab === 'email' && (
        <BulkEmailPanel
          token={token}
          getEmails={getAllMemberEmails}
          defaultSubject={team?.displayName ? `Info: ${team.displayName}` : 'Team‑info'}
          contextKey={`team:${teamId}:email`}
        />
      )}
    </div>
  );
}

function TeamMembersView({ token, teamId, members, loading }) {
  const [viewMode, setViewMode] = useState(() => {
    try { return localStorage.getItem(`team:${teamId}:members_view`) || 'cards'; } catch { return 'cards'; }
  });
  useEffect(() => { try { localStorage.setItem(`team:${teamId}:members_view`, viewMode); } catch {} }, [teamId, viewMode]);

  // Coerce minimal member objects into user-like for UserList/UserCard
  const normalized = useMemo(() => members.map(m => ({
    id: m.userId || m.id || m.mail || m.email || m.displayName,
    displayName: m.displayName || m['displayName'] || m.mail || m.email || 'Medlem',
    mail: m.mail || m.email || '',
    userPrincipalName: m.mail || m.email || '',
    jobTitle: m.jobTitle || '',
    department: m.department || '',
  })), [members]);

  return (
    <div>
      <div className="section-header">
        <b>Medlemmar ({members.length})</b>
        <div className="spacer" />
        <div className="muted" style={{ display:'inline-flex', gap:6, alignItems:'center', marginRight:8 }}>
          <span>Visning:</span>
          <button className={`btn btn-light ${viewMode==='table'?'active':''}`} onClick={() => setViewMode('table')}>Tabell</button>
          <button className={`btn btn-light ${viewMode==='cards'?'active':''}`} onClick={() => setViewMode('cards')}>Kort</button>
        </div>
        <CsvExportControls items={normalized} storageKey={`team:${teamId}:members`} defaultFileName={`team_${teamId}_medlemmar.csv`} buttonLabel="Exportera CSV" />
      </div>
      {loading ? (
        <div className="muted">Laddar medlemmar…</div>
      ) : (
        <div style={{ marginTop:12 }}>
          {viewMode === 'table' ? (
            <UserList
              users={normalized}
              token={token}
              mode={'table'}
              fieldsKey={`team:${teamId}:members`}
              selectedFields={(() => { try { return JSON.parse(localStorage.getItem('userlist_fields_team_members')||'[]'); } catch { return []; } })()}
              onChangeSelectedFields={(f) => { try { localStorage.setItem('userlist_fields_team_members', JSON.stringify(f)); } catch {} }}
            />
          ) : (
            <div style={{ display:'flex', flexWrap:'wrap', gap:12 }}>
              {normalized.map(u => (
                <UserCard key={u.id}
                  user={u}
                  token={token}
                  onOpenDetails={() => { try { window.location.hash = `#/users/${encodeURIComponent(u.id || u.userPrincipalName || u.mail)}`; } catch {} }}
                />
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function normalizeMember(m, email) {
  return {
    ...m,
    mail: email || '',
  };
}