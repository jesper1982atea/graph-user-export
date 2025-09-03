import React, { useState, useEffect, useRef } from 'react';
import UserCard from './UserCard';
import UserList from './UserList';
import ChatDetailPage from './ChatDetailPage';
import '../../styles.css';
import { saveAs } from 'file-saver';
import CsvExportControls from './CsvExportControls';
import { BrowserRouter as Router, Route, Routes, useNavigate, useParams, useLocation } from 'react-router-dom';
import FieldLabelMappingPage from './FieldLabelMappingPage';
import MeetingDetailPage from './MeetingDetailPage';
import TeamDetailPage from './TeamDetailPage';
import UserDetailPage from './UserDetailPage';
import DepartmentsPage from './DepartmentsPage';
import { loadFieldLabelMap, saveFieldLabelMap } from './fieldLabelMap';
import ChatAvatar from './ChatAvatar';
import OneOnOnePairAvatars from './OneOnOnePairAvatars';
import UpdateIndicator from './UpdateIndicator';
import InBrowserUpdater from './InBrowserUpdater';
import { getCachedChatMembers, fetchChatMembers } from './ChatMembersCache';
import { USER_FIELDS } from './userFields';
import { GRAPH_USER_SELECT_FIELDS } from './graphUserSelect';
import { findFirstAvailable, createMeeting } from './meetingScheduler';
// HomePage (index) som visar grundinnehållet från GraphMeetingsAndChats
const SunIcon = ({ size=18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>
);
const MoonIcon = ({ size=18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
);

// Enkel Error Boundary för att fånga render-fel och visa en vänlig sida
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    // Kan loggas till telemetri vid behov
    // console.error('ErrorBoundary caught:', error, info);
  }
  handleReset = () => {
    this.setState({ hasError: false, error: null });
    if (this.props.onReset) this.props.onReset();
  };
  render() {
    if (this.state.hasError) {
      const msg = this.state.error?.message || 'Ett oväntat fel inträffade.';
      return (
        <div className="app-container">
          <div className="card" style={{ borderColor: '#fecaca', background: 'var(--card-bg)' }}>
            <h2 style={{ marginTop: 0 }}>Något gick fel</h2>
            <p className="muted" style={{ whiteSpace: 'pre-wrap' }}>{msg}</p>
            <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button className="btn btn-primary" onClick={this.handleReset}>Gå till översikt</button>
              <button className="btn btn-light" onClick={() => window.location.reload()}>Ladda om sidan</button>
            </div>
            <small className="muted" style={{ display:'block', marginTop: 8 }}>Tips: Kontrollera att alla nödvändiga fält och props finns innan du försöker använda dem.</small>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

const CalendarIcon = ({ size=16 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
);

const ChatIcon = ({ size=16 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15a4 4 0 0 1-4 4H7l-4 4V7a4 4 0 0 1 4-4h10a4 4 0 0 1 4 4z"/></svg>
);

const LayersIcon = ({ size=16 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>
);

function FieldLabelMappingSettings() {
  const [map, setMap] = useState(() => loadFieldLabelMap());
  const [newKey, setNewKey] = useState('');
  const [newLabel, setNewLabel] = useState('');
  const keys = Object.keys(map || {}).sort();
  const upsert = (k, v) => {
    const next = { ...(map || {}) };
    if (!k) return;
    if (v == null || v === '') delete next[k]; else next[k] = v;
    setMap(next);
    saveFieldLabelMap(next);
  };
  return (
    <div style={{ marginTop:8 }}>
      <div style={{ fontWeight:600 }}>Rubrikmappning (visa egna etiketter för fält)</div>
      <div className="muted" style={{ marginTop:4 }}>Exempel: extensionAttribute1 → Kostnadsställe. Gäller i tabeller/kort och CSV-headers.</div>
      <div style={{ display:'flex', gap:8, alignItems:'center', marginTop:8, flexWrap:'wrap' }}>
        <input value={newKey} onChange={e=>setNewKey(e.target.value)} placeholder="fält-namn (t.ex. extensionAttribute1)" style={{ minWidth:260 }} />
        <input value={newLabel} onChange={e=>setNewLabel(e.target.value)} placeholder="visningsnamn (t.ex. Kostnadsställe)" style={{ minWidth:240 }} />
        <button className="btn btn-light" onClick={() => { if (newKey) { upsert(newKey, newLabel || newKey); setNewKey(''); setNewLabel(''); } }}>Lägg till/uppdatera</button>
        <button className="btn btn-light" onClick={() => { setMap({}); saveFieldLabelMap({}); }}>Rensa alla</button>
      </div>
      {keys.length > 0 && (
        <div style={{ marginTop:8, display:'grid', gap:6 }}>
          {keys.map(k => (
            <div key={k} style={{ display:'flex', gap:8, alignItems:'center' }}>
              <code style={{ background:'var(--bg)', padding:'2px 6px', borderRadius:6 }}>{k}</code>
              <span>→</span>
              <input value={map[k]} onChange={e => upsert(k, e.target.value)} style={{ minWidth:220 }} />
              <button className="btn btn-light" onClick={() => upsert(k, '')}>Ta bort</button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// (Removed legacy inline CsvExportControls; using shared component from './CsvExportControls')

const HomePage = (props) => {
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [avatarError, setAvatarError] = useState(false);
  const meCardRef = useRef(null);
  const {
  token,
    handleTokenChange,
  verifyToken,
  verifyLoading,
  tokenVerified,
    ThemeToggle,
  showTokenCard,
  onToggleTokenCard,
    searchInput,
    handleSearchInputChange,
    handleUserSearch,
    searchLoading,
    me,
    searchedUsers = [],
    error,
    meetings = [],
    navigate,
    chats = [],
    channels = [],
  // Selection & chat creation props (fix ReferenceError by pulling from props)
  selectedUsersMap = {},
  onToggleSelectUser,
  onCreateGroupChat,
  createChatLoading = false,
  createChatError,
  onExportMeetings,
  onExportChats,
  onExportChannels,
  onExportUsers,
  onReloadChats,
  onClearToken,
  loadingLists = false,
  currentPage = 'dashboard',
  chatCreateEnabled = false,
  onToggleChatCreateEnabled,
  } = props || {};
  // Search results view mode (table or cards)
  const [searchViewMode, setSearchViewMode] = useState(() => {
    try { return localStorage.getItem('search_view_mode') || 'table'; } catch { return 'table'; }
  });
  // Inline Team details state
  const [selectedTeamId, setSelectedTeamId] = useState(null);
  // Team photos cache
  const [teamPhotos, setTeamPhotos] = useState({}); // { [teamId]: objectUrl|null }
  const authHeaderForHome = React.useMemo(() => {
    const trimmed = (token || '').trim();
    return trimmed.toLowerCase().startsWith('bearer ') ? trimmed : trimmed ? `Bearer ${trimmed}` : '';
  }, [token]);
  const fetchTeamPhoto = React.useCallback(async (teamId) => {
    if (!authHeaderForHome) return;
    try {
      const res = await fetch(`https://graph.microsoft.com/v1.0/groups/${encodeURIComponent(teamId)}/photo/$value`, {
        headers: { Authorization: authHeaderForHome },
      });
      if (!res.ok) { setTeamPhotos(prev => ({ ...prev, [teamId]: null })); return; }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setTeamPhotos(prev => ({ ...prev, [teamId]: url }));
    } catch { setTeamPhotos(prev => ({ ...prev, [teamId]: null })); }
  }, [authHeaderForHome]);
  useEffect(() => {
    const list = props.joinedTeams || [];
    for (const t of list) {
      if (teamPhotos[t.id] === undefined) fetchTeamPhoto(t.id);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.joinedTeams, fetchTeamPhoto]);
  const setSearchMode = (mode) => {
    setSearchViewMode(mode);
    try { localStorage.setItem('search_view_mode', mode); } catch {}
  };
  const isActive = (page) => currentPage === page;
  const goto = (path) => props.navigate && props.navigate(path);
  const tokenScopes = React.useMemo(() => {
    try {
      const t = props.token || '';
      const parts = t.split('.');
      if (parts.length >= 2) {
        const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
        const json = JSON.parse(atob(base64));
        return json.scp || '';
      }
    } catch {}
    return '';
  }, [props.token]);
  const tokenExpired = typeof props.secondsLeft === 'number' && props.secondsLeft <= 0;
  const formatLeft = (s) => {
    if (s == null) return '';
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    const sec = s % 60;
    const two = (n) => String(n).padStart(2, '0');
    return h > 0 ? `${h}:${two(m)}:${two(sec)}` : `${m}:${two(sec)}`;
  };

  return (
    <div className="app-container">
  {/* Compact update indicator replaces large banner */}
      <div className="app-hero">
        <div className="hero-content">
          <div style={{ display:'flex', alignItems:'center', gap:12, position:'relative' }}>
            <h1 style={{ marginRight: 'auto' }}>Microsoft Graph – Möten & Chat Export</h1>
            <div className="header-actions">
              {ThemeToggle ? <div className="control-40 theme-holder"><ThemeToggle /></div> : null}
              <button className="btn btn-ghost control-40" onClick={onToggleTokenCard}>
                {tokenVerified && (
                  <img src={props.mePhotoUrl || ''} alt="" className="avatar avatar-sm" onError={(e)=>{ e.currentTarget.style.display='none'; }} />
                )}
                {showTokenCard || !tokenVerified ? 'Dölj token' : 'Visa token'}
              </button>
              {tokenVerified && (
              <div style={{ position:'relative' }} className="account-chip">
                <button
                  className="btn btn-circle"
                  onClick={() => setUserMenuOpen(v => !v)}
                  title={me?.displayName || me?.userPrincipalName}
                  aria-label="Konto"
                >
                  {props.mePhotoUrl && !avatarError ? (
                    <img
                      src={props.mePhotoUrl}
                      alt=""
                      className="avatar avatar-md"
                      onError={() => setAvatarError(true)}
                    />
                  ) : (
                    <span className="avatar-fallback">{(me?.displayName || me?.userPrincipalName || '?').slice(0,1).toUpperCase()}</span>
                  )}
                </button>
                <div className="account-name" title={me?.displayName || me?.userPrincipalName}>
                  {me?.displayName || me?.userPrincipalName}
                </div>
                {userMenuOpen && (
                  <div style={{ position:'absolute', right:0, top:'110%', background:'var(--card-bg)', border:'1px solid var(--border)', borderRadius:10, boxShadow:'0 8px 30px rgba(0,0,0,.12)', padding:8, minWidth:180, zIndex:20 }}>
                    <button className="btn btn-light" style={{ width:'100%' }} onClick={() => {
                      setUserMenuOpen(false);
                      if (props.navigate) props.navigate('/dashboard');
                      setTimeout(() => {
                        if (meCardRef.current) meCardRef.current.scrollIntoView({ behavior:'smooth', block:'start' });
                      }, 50);
                    }}>Visa profil</button>
                    <button className="btn btn-light" style={{ width:'100%', marginTop:6 }} onClick={() => { setUserMenuOpen(false); props.onClearToken && props.onClearToken(); }}>Logga ut</button>
                  </div>
                )}
              </div>
              )}
              {/* Subtle update icon with popover */}
              <UpdateIndicator openSettings={() => goto('/settings')} />
            </div>
          </div>
          <p>Verifiera din token, sök användare, filtrera datum och exportera listor till CSV. Allt körs lokalt i din webbläsare.</p>
          <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {tokenVerified ? (
              tokenExpired ? (
                <span className="badge badge-warn">Token utgången</span>
              ) : (
                <span className="badge badge-info">Token verifierad</span>
              )
            ) : (
              <span className="badge badge-neutral">Token behövs</span>
            )}
            {typeof props.secondsLeft === 'number' && props.secondsLeft >= 0 && (
              <span className="badge badge-neutral" title="Tid kvar tills token går ut">
                ⏳ {formatLeft(props.secondsLeft)} kvar
              </span>
            )}
            <span className="badge badge-neutral" style={{ cursor:'pointer' }} onClick={() => goto('/meetings')}><CalendarIcon />&nbsp;Möten: {meetings.length}</span>
            <span className="badge badge-neutral" style={{ cursor:'pointer' }} onClick={() => goto('/chats')}><ChatIcon />&nbsp;Chattar: {chats.length}</span>
            <span className="badge badge-neutral" style={{ cursor:'pointer' }} onClick={() => goto('/teams')}><LayersIcon />&nbsp;Teams: {(props.joinedTeams||[]).length}</span>
          </div>
        </div>
      </div>

      {/* Top navigation tabs */}
      <div className="tabs-bar">
        <div style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
          <button className={`btn tab-btn ${isActive('dashboard') ? 'active' : ''}`} onClick={() => goto('/dashboard')}>
            <span className="tab-icon" aria-hidden>🏠</span>
            <span className="tab-label">Översikt</span>
          </button>
          <button className={`btn tab-btn ${isActive('meetings') ? 'active' : ''}`} onClick={() => goto('/meetings')}>
            <span className="tab-icon" aria-hidden><CalendarIcon /></span>
            <span className="tab-label">Möten</span>
            <span className="tab-count">{meetings.length}</span>
          </button>
          <button className={`btn tab-btn ${isActive('chats') ? 'active' : ''}`} onClick={() => goto('/chats')}>
            <span className="tab-icon" aria-hidden><ChatIcon /></span>
            <span className="tab-label">Chattar</span>
            <span className="tab-count">{chats.length}</span>
          </button>
          <button className={`btn tab-btn ${isActive('teams') ? 'active' : ''}`} onClick={() => goto('/teams')}>
            <span className="tab-icon" aria-hidden><LayersIcon /></span>
            <span className="tab-label">Teams</span>
            <span className="tab-count">{(props.joinedTeams||[]).length}</span>
          </button>
          <button className={`btn tab-btn ${isActive('search') ? 'active' : ''}`} onClick={() => goto('/search')}>
            <span className="tab-icon" aria-hidden>🔎</span>
            <span className="tab-label">Sök</span>
          </button>
          <button className={`btn tab-btn ${isActive('departments') ? 'active' : ''}`} onClick={() => goto('/departments')}>
            <span className="tab-icon" aria-hidden>🏷️</span>
            <span className="tab-label">Attribut</span>
          </button>
          <button className={`btn tab-btn ${isActive('settings') ? 'active' : ''}`} onClick={() => goto('/settings')}>
            <span className="tab-icon" aria-hidden>⚙️</span>
            <span className="tab-label">Inställningar</span>
          </button>
        </div>
      </div>

  {(showTokenCard || !tokenVerified) && (
  <div className="card">
        <label style={{ fontWeight: 600, display: 'block' }}>Hämta Auth Token här <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank" rel="noopener noreferrer">https://developer.microsoft.com/en-us/graph/graph-explorer</a></label>
        <small style={{ color: '#6b7280' }}>Klistra in din token från Graph Explorer (med eller utan "Bearer")</small>
        {typeof props.secondsLeft === 'number' && (
          <div style={{ marginTop: 6 }} className={tokenExpired ? 'muted' : ''}>
            {tokenExpired ? 'Token har gått ut. Klistra in en ny token.' : `Token går ut om ${formatLeft(props.secondsLeft)}.`}
          </div>
        )}
        <textarea
          value={token || ''}
          onChange={(e) => (handleTokenChange ? handleTokenChange(e) : null)}
          rows={4}
          placeholder="eyJ0eXAiOiJK..."
          className="token-textarea"
        />
        <div style={{ marginTop: 8, display: 'flex', gap: 8 }}>
          <button
            className="btn btn-primary"
            onClick={verifyToken}
            disabled={verifyLoading || !token || token.trim() === ''}
          >
            {verifyLoading ? 'Verifierar...' : tokenVerified ? 'Verifierad' : 'Spara/Verifiera'}
          </button>
          {tokenVerified && <span className="status-ok" style={{ alignSelf: 'center' }}>Token verifierad</span>}
          {token && (
            <button className="btn btn-light" onClick={onClearToken} style={{ marginLeft: 'auto' }}>Rensa token</button>
          )}
        </div>
      </div>
  )}

  {isActive('dashboard') && tokenVerified && me ? (
        <div className="card" ref={meCardRef}>
          <div className="section-title" style={{ marginBottom: 8 }}>
            <b>Din profil</b>
          </div>
          <UserCard user={me} token={props.token} />
        </div>
      ) : null}

  {isActive('search') && (
        <div className="card">
          <b style={{ display: 'block', marginBottom: 6 }}>Sök användare</b>
          {props.searchDepartment && (
            <div className="muted" style={{ marginBottom:8, display:'flex', alignItems:'center', gap:8 }}>
              <span>Filter:</span>
              <span className="badge badge-info" title="Aktiv avdelningsfiltrering">Avdelning: {props.searchDepartment}</span>
              <button className="btn btn-light" onClick={() => {
                try {
                  const url = new URL(window.location.href);
                  url.searchParams.delete('department');
                  window.location.href = url.toString();
                } catch {}
              }}>Rensa</button>
            </div>
          )}
          <textarea
            value={searchInput || ''}
            onChange={(e) => (handleSearchInputChange ? handleSearchInputChange(e) : null)}
            rows={2}
            placeholder="anvandare1@domain.se, anvandare2@domain.se"
            className="search-textarea"
          />
          <button
            className="btn btn-primary"
            onClick={() => (handleUserSearch ? handleUserSearch() : null)}
            disabled={searchLoading}
            style={{ marginTop: 8 }}
          >
            {searchLoading ? 'Söker...' : 'Sök användare'}
          </button>
        </div>
      )}

  {/* Removed per-page 'Inloggad som' card since header shows user info */}

  {isActive('search') && !props.showUserDetail && searchedUsers.length > 0 && (
        <div className="card">
          <div className="section-header">
            <b>Sökresultat ({searchedUsers.length})</b>
            <div className="spacer" />
            <div className="muted" style={{ display:'inline-flex', gap:6, alignItems:'center', marginRight:8 }}>
              <span>Visning:</span>
              <button className={`btn btn-light ${searchViewMode==='table'?'active':''}`} onClick={() => setSearchMode('table')}>Tabell</button>
              <button className={`btn btn-light ${searchViewMode==='cards'?'active':''}`} onClick={() => setSearchMode('cards')}>Kort</button>
            </div>
            <label className="muted" style={{ display:'inline-flex', alignItems:'center', gap:6, marginRight:8 }} title="Tillåt skapande av nya chattar">
              <input type="checkbox" checked={!!chatCreateEnabled} onChange={() => onToggleChatCreateEnabled && onToggleChatCreateEnabled()} />
              Tillåt skapande
            </label>
            {Object.keys(selectedUsersMap || {}).filter(id => selectedUsersMap[id]).length > 0 && (
              <>
                <span className="muted" style={{ marginRight: 8 }}>Valda: {Object.keys(selectedUsersMap).filter(id => selectedUsersMap[id]).length}</span>
                <button className="btn btn-primary" onClick={() => onCreateGroupChat && onCreateGroupChat()} disabled={createChatLoading || !tokenVerified || !chatCreateEnabled} title={!chatCreateEnabled ? 'Avstängt i inställningar' : undefined}>
                  {createChatLoading ? 'Skapar…' : 'Skapa 1‑1 chatt'}
                </button>
                <div style={{ width:8 }} />
                <button className="btn btn-secondary" onClick={async () => {
                  setScheduleError('');
                  setScheduleResult(null);
                  setScheduling(true);
                  try {
                    const emails = searchedUsers.filter(u => selectedUsersMap[u.id]).map(u => u.mail || u.userPrincipalName).filter(Boolean);
                    const suggestion = await findFirstAvailable({
                      token: props.token,
                      attendeeEmails: emails,
                      durationMinutes: Number(scheduleDuration) || 30,
                      windowStart: scheduleWindowStart || undefined,
                      windowEnd: scheduleWindowEnd || undefined,
                      tenantDomain: (me?.userPrincipalName || '').split('@')[1] || undefined,
                    });
                    setScheduleResult(suggestion);
                    if (!suggestion) setScheduleError('Ingen gemensam tid hittades inom fönstret.');
                  } catch (e) {
                    setScheduleError(e.message || 'Kunde inte hitta tid.');
                  } finally {
                    setScheduling(false);
                  }
                }} disabled={scheduling || !tokenVerified}>
                  {scheduling ? 'Söker tid…' : 'Hitta första lediga tid'}
                </button>
              </>
            )}
                <CsvExportControls items={searchedUsers} storageKey="users" defaultFileName="anvandare.csv" />
          </div>
          {Object.keys(selectedUsersMap || {}).filter(id => selectedUsersMap[id]).length > 0 && (
            <div style={{ display:'flex', gap:8, flexWrap:'wrap', marginTop:8, alignItems:'center' }}>
              <label className="muted">Längd (min):</label>
              <input type="number" min="15" max="240" step="5" value={scheduleDuration} onChange={e => setScheduleDuration(e.target.value)} style={{ width:80 }} />
              <label className="muted">Från (UTC):</label>
              <input type="datetime-local" value={scheduleWindowStart} onChange={e => setScheduleWindowStart(e.target.value)} />
              <label className="muted">Till (UTC):</label>
              <input type="datetime-local" value={scheduleWindowEnd} onChange={e => setScheduleWindowEnd(e.target.value)} />
              <label className="muted">Ämne:</label>
              <input type="text" value={scheduleSubject || ''} onChange={e => setScheduleSubject(e.target.value)} placeholder="Möte" style={{ minWidth: 180 }} />
              <label className="muted">Beskrivning:</label>
              <input type="text" value={scheduleBody || ''} onChange={e => setScheduleBody(e.target.value)} placeholder="Valfri agenda/anteckningar" style={{ minWidth: 240 }} />
              <label className="muted" title="Säkerställ Teams-möte">
                <input type="checkbox" checked={!!scheduleOnline} onChange={e => setScheduleOnline(e.target.checked)} /> Teams‑möte
              </label>
              {scheduleResult && (
                <>
                  <span className="badge badge-info">Förslag: {scheduleResult.start.dateTime?.replace('T',' ').slice(0,16)}–{scheduleResult.end.dateTime?.replace('T',' ').slice(0,16)} (conf {Math.round((scheduleResult.confidence||0)*100)}%)</span>
                  <button className="btn btn-primary" onClick={async () => {
                    try {
                      const emails = searchedUsers.filter(u => selectedUsersMap[u.id]).map(u => u.mail || u.userPrincipalName).filter(Boolean);
                      await createMeeting({
                        token: props.token,
                        subject: scheduleSubject || 'Möte',
                        bodyHtml: scheduleBody || '',
                        attendeeEmails: emails,
                        start: scheduleResult.start,
                        end: scheduleResult.end,
                        isOnline: !!scheduleOnline,
                      });
                      alert('Möte skapat.');
                    } catch (e) { alert(e.message || 'Kunde inte skapa möte'); }
                  }}>Boka möte</button>
                </>
              )}
              {scheduleError && <span className="muted" style={{ color:'#b91c1c' }}>{scheduleError}</span>}
            </div>
          )}
          {!chatCreateEnabled && (
            <div className="muted" style={{ marginTop:6 }}>Tips: Token från Graph Explorer saknar ofta Chat.ReadWrite. Stäng/öppna denna inställning beroende på dina scopes.</div>
          )}
          {createChatError && <div style={{ color:'#b91c1c', marginTop:8 }}>{createChatError}</div>}
          <div style={{ marginTop:12 }}>
            {searchViewMode === 'table' ? (
              <>
        <UserList
      users={searchedUsers.filter(u => !u.error)}
      token={props.token}
                  mode={'table'}
                  fieldsKey={'search'}
                  selectable
                  selectedMap={props.selectedUsersMap || {}}
                  onToggleSelect={props.onToggleSelectUser}
                  onChangeSelectedFields={(f) => { try { localStorage.setItem('userlist_fields_search', JSON.stringify(f)); } catch {} }}
                  selectedFields={(() => { try { return JSON.parse(localStorage.getItem('userlist_fields_search')||'[]'); } catch { return []; } })()}
                />
              </>
            ) : (
              <div style={{ display:'flex', flexWrap:'wrap', gap:12 }}>
        {searchedUsers.filter(u => !u.error).map(u => (
                  <UserCard key={u.id || u.userPrincipalName || u.mail}
                    user={u}
          token={props.token}
                    onOpenDetails={() => { try { window.location.hash = `#/users/${encodeURIComponent(u.id || u.userPrincipalName || u.mail)}`; } catch {} }}
                  />
                ))}
              </div>
            )}
            {searchedUsers.some(u => u.error) && (
              <div style={{ color: '#b91c1c', background: '#fee2e2', border: '1px solid #fecaca', padding: 8, borderRadius: 8, marginTop:8 }}>
                {searchedUsers.filter(u => u.error).length} användare kunde inte hämtas fullt ut.
              </div>
            )}
          </div>
        </div>
      )}

      {error && <div style={{ color: 'red', marginBottom: '1rem' }}>{error}</div>}

      {isActive('search') && tokenVerified && props.showUserDetail ? (
        <UserDetailPage token={props.token} />
      ) : null}

          {isActive('departments') && tokenVerified ? (
            <div className="card">
              <DepartmentsPage token={props.token} />
            </div>
          ) : (
            isActive('departments') ? <div className="card"><span className="muted">Verifiera token för att se attribut.</span></div> : null
          )}

      {isActive('meetings') && tokenVerified ? (
        props.showMeetingDetail ? (
          <MeetingDetailPage meetings={meetings} token={props.token} />
        ) : (
        <div className="card">
          <div className="section-title" style={{ marginBottom: 8 }}>
            <b>Möten</b>
            <span className="spacer" />
          </div>
          <div className="section-header">
            <b>Möten ({meetings.length})</b>
            {meetings.length > 0 && (
              <>
                <div className="spacer" />
                <CsvExportControls items={meetings} storageKey="meetings" defaultFileName="moten.csv" />
              </>
            )}
          </div>
          <div className="controls" style={{ marginBottom: 8 }}>
            <input type="date" className="date-input" value={props.meetingsFrom || ''} onChange={e => props.onChangeFrom && props.onChangeFrom(e.target.value)} />
            <input type="date" className="date-input" value={props.meetingsTo || ''} onChange={e => props.onChangeTo && props.onChangeTo(e.target.value)} />
            <button className="btn btn-light" onClick={props.onApplyDateFilter}>Filtrera datum</button>
          </div>
          {verifyLoading || loadingLists ? (
            <ul className="grid-list">
              {Array.from({ length: 6 }).map((_, i) => (
                <li className="grid-item skeleton" key={i} style={{ height: 88 }} />
              ))}
            </ul>
          ) : meetings.length === 0 ? (
            <p className="list-empty">Inga möten hittades.</p>
          ) : (
            <ul className="grid-list">
              {meetings.map((m, i) => (
                <li className="grid-item" key={m.id || i} onClick={() => (navigate ? navigate(`/meetings/${m.id}`) : null)}>
                  <div className="grid-item-title">{m.subject || '(utan ämne)'}</div>
                  <div className="grid-item-sub">
                    {m.start?.dateTime?.replace('T', ' ').slice(0, 16)} — {m.end?.dateTime?.replace('T', ' ').slice(0, 16)}
                  </div>
                  {m.location?.displayName && <div className="grid-item-meta">Plats: {m.location.displayName}</div>}
                  <div style={{ marginTop:8 }}>
                    <button className="btn btn-secondary" title="Boka uppföljningsmöte" onClick={async (e) => {
                      e.stopPropagation();
                      try {
                        const emails = (m.attendees || []).map(a => a.emailAddress?.address).filter(Boolean);
                        if (!emails.length) { alert('Inga deltagare att föreslå uppföljning för.'); return; }
                        const suggestion = await findFirstAvailable({
                          token: authToken,
                          attendeeEmails: emails,
                          durationMinutes: 30,
                          windowStart: new Date(new Date(m.end?.dateTime || Date.now()).getTime() + 15*60*1000).toISOString(),
                          windowEnd: new Date(Date.now() + 14*24*60*60*1000).toISOString(),
                          tenantDomain: (me?.userPrincipalName || '').split('@')[1] || undefined,
                        });
                        if (!suggestion) { alert('Hittade ingen gemensam tid.'); return; }
                        await createMeeting({ token: props.token, subject: `Uppföljning: ${m.subject || ''}`, attendeeEmails: emails, start: suggestion.start, end: suggestion.end, isOnline: true });
                        alert('Uppföljningsmöte skapat.');
                      } catch (e2) { alert(e2.message || 'Kunde inte skapa uppföljning'); }
                    }}>Boka uppföljning</button>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
        )
      ) : (
        isActive('meetings') ? <div className="card"><span className="muted">Verifiera token för att se möten.</span></div> : null
      )}

      {isActive('chats') && tokenVerified ? (
        props.showChatDetail ? (
          <ChatDetailPage chats={chats} token={props.token} meId={me?.id} meUpn={me?.userPrincipalName} mePhotoUrl={props.mePhotoUrl} />
        ) : (
        <div className="card">
          <div className="section-title" style={{ marginBottom: 8 }}>
            <b>Chattar</b>
            <span className="spacer" />
          </div>
          <div className="section-header">
            <b>Teams-chattar ({chats.length})</b>
            {chats.length > 0 && (
              <>
                <div className="spacer" />
                <CsvExportControls items={chats} storageKey="chats" defaultFileName="chattar.csv" />
              </>
            )}
            <div className="spacer" />
            <button className="btn btn-light" onClick={() => onReloadChats && onReloadChats()} disabled={verifyLoading || loadingLists}>Ladda om</button>
          </div>
          {verifyLoading || loadingLists ? (
            <ul className="grid-list">
              {Array.from({ length: 6 }).map((_, i) => (
                <li className="grid-item skeleton" key={i} style={{ height: 72 }} />
              ))}
            </ul>
          ) : chats.length === 0 ? (
            <div className="muted">
              Inga chattar hittades.
              {props.chatsError ? (
                <div style={{ marginTop:6 }}>{props.chatsError}</div>
              ) : (
                <div style={{ marginTop:6 }}>Tips: Token behöver minst Chat.ReadBasic (eller Chat.Read/Chat.ReadWrite) för att lista chattar.</div>
              )}
            </div>
          ) : (
            <ul className="grid-list">
              {chats.map((c, i) => {
                // Derive subtitle/secondary text: for oneOnOne, show counterpart; for group, show member count
                const authHeader = (() => {
                  const trimmed = (props.token || '').trim();
                  return trimmed.toLowerCase().startsWith('bearer ') ? trimmed : `Bearer ${trimmed}`;
                })();
                const cached = getCachedChatMembers(c.id);
        let subtitle = c.chatType;
        let oneOnOneNames = '';
                if (c.chatType === 'oneOnOne') {
                  const members = cached || [];
                  const others = members
                    .map(m => ({ name: m.displayName, id: m.email || m.userId || m.id }))
                    .filter(m => m.id && m.id !== me?.id);
                  if (others.length) {
          oneOnOneNames = others.map(o => o.name || o.id).join(', ');
          subtitle = '1‑1 chatt';
                  }
                } else if (c.chatType && c.chatType !== 'oneOnOne') {
                  const members = cached || [];
                  if (members.length) subtitle = `${c.chatType} · ${members.length} deltagare`;
                }
                // Optimistic async fill of cache if empty
                if (!cached && props.token && c.id) {
                  fetchChatMembers(c.id, authHeader).then(() => { /* no-op re-render by state change elsewhere */ });
                }
    const isOne = c.chatType === 'oneOnOne';
                let a = null, b = null;
                if (isOne) {
                  const members = cached || [];
                  if (members.length >= 2) {
                    // Prefer to place current user first if present
                    const mine = members.find(m => (m.userId || m.id) === me?.id || m.email === me?.userPrincipalName);
                    const other = members.find(m => ((m.userId || m.id) !== me?.id) && (m.email !== me?.userPrincipalName));
                    if (mine && other) { a = mine; b = other; }
                    else { a = members[0]; b = members[1]; }
                  }
                }
                return (
                  <li className="grid-item" key={c.id || i} onClick={() => (navigate ? navigate(`/chats/${c.id}`) : null)}>
                    <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                      {isOne && a && b ? (
                        <OneOnOnePairAvatars a={a} b={b} token={props.token} size={36} myPhotoUrl={props.mePhotoUrl} overlapRatio={1.1} />
                      ) : (
                        <ChatAvatar chat={c} token={props.token} meId={me?.id} size={36} />
                      )}
                      <div>
      <div className="grid-item-title">{isOne ? (oneOnOneNames || '1‑1 chatt') : (c.topic || 'Ingen rubrik')}</div>
      <div className="grid-item-sub">{isOne ? '1‑1 chatt' : subtitle}</div>
                      </div>
                    </div>
                  </li>
                );
              })}
            </ul>
          )}
        </div>
        )
      ) : (
        isActive('chats') ? <div className="card"><span className="muted">Verifiera token för att se chattar.</span></div> : null
      )}

      {isActive('teams') && tokenVerified ? (
        <div className="card">
          <div className="section-title" style={{ marginBottom: 8 }}>
            <b>Teams</b>
            <span className="spacer" />
          </div>
          {!selectedTeamId && (
            <div className="section-header">
              <b>Dina Teams ({(props.joinedTeams||[]).length})</b>
              <div className="spacer" />
              <button className="btn btn-light" onClick={props.onReloadTeams} disabled={props.teamsLoading}>
                {props.teamsLoading ? 'Laddar…' : 'Ladda om'}
              </button>
            </div>
          )}
          {selectedTeamId ? (
            <div style={{ marginTop: 8 }}>
              <TeamDetailPage token={props.token} teamId={selectedTeamId} inline onClose={() => setSelectedTeamId(null)} />
            </div>
          ) : verifyLoading || props.teamsLoading ? (
            <ul className="grid-list">
              {Array.from({ length: 6 }).map((_, i) => (
                <li className="grid-item skeleton" key={i} style={{ height: 72 }} />
              ))}
            </ul>
          ) : (props.joinedTeams||[]).length === 0 ? (
            <div className="muted">
              Inga Teams hittades.
              {props.teamsError ? (
                <div style={{ marginTop:6 }}>{props.teamsError}</div>
              ) : (
                <div style={{ marginTop:6 }}>
                  Kontrollera att din token har rätt scopes (t.ex. Team.ReadBasic.All eller Team.Read.All). Om du använder Graph Explorer, logga in och ge samtycke.
                </div>
              )}
              <div style={{ marginTop:8 }}>
                <button className="btn btn-light" onClick={() => window.location.reload()}>Försök igen</button>
              </div>
            </div>
          ) : (
            <ul className="grid-list">
              {(props.joinedTeams||[]).map(t => (
                <li className="grid-item" key={t.id} onClick={() => setSelectedTeamId(t.id)} style={{ cursor: 'pointer' }}>
                  <div style={{ display:'flex', alignItems:'center', gap:12 }}>
                    {teamPhotos[t.id] ? (
                      <img src={teamPhotos[t.id]} alt="Team" style={{ width:48, height:48, borderRadius:8, objectFit:'cover' }} />
                    ) : (
                      <div style={{ width:48, height:48, borderRadius:8, background:'#eef2f7', display:'flex', alignItems:'center', justifyContent:'center', color:'#64748b', fontWeight:600 }}>
                        {(t.displayName || 'T').substring(0,1).toUpperCase()}
                      </div>
                    )}
                    <div>
                      <div className="grid-item-title">{t.displayName || 'Team'}</div>
                      <div className="grid-item-sub" title={t.id}>Id: {t.id}</div>
                    </div>
                    <div className="spacer" />
                    <div className="muted" style={{ fontSize:'.9rem' }}>
                      Klicka för detaljer →
                    </div>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : (
        isActive('teams') ? <div className="card"><span className="muted">Verifiera token för att se dina Teams.</span></div> : null
      )}

      {/* Dashboard (översikt) */}
  {isActive('dashboard') && (
        <div className="card">
          <b>Översikt</b>
          <p className="muted" style={{ marginTop: 6 }}>Välj en flik ovan för att visa Möten, Chattar, Teams eller Sök. Du kan också klicka på badges i hjälten.</p>
          <ul style={{ marginTop: 10, color: 'var(--muted)' }}>
            <li>✔️ Verifiera en Graph-token från Graph Explorer (Bearer-token) och kör allt lokalt i webbläsaren.</li>
            <li>📅 Visa och filtrera Möten på datumintervall och exportera till CSV.</li>
            <li>💬 Lista dina Teams-chattar och exportera till CSV.</li>
            <li>🧩 Se dina Teams, visa kanaler och medlemmar, och exportera listor till CSV.</li>
            <li>🔎 Sök användare via e-post eller namn (startswith) och exportera resultatet.</li>
            <li>🔐 Inget backend krävs – dina data lämnar inte din webbläsare.</li>
          </ul>
        </div>
      )}

      {isActive('dashboard') && (
        <details className="card">
          <summary style={{ cursor:'pointer', fontWeight:600 }}>Hämta token (Graph Explorer)</summary>
          <div>
            <ol style={{ marginTop: 8, color: 'var(--muted)' }}>
              <li>Gå till <a href="https://developer.microsoft.com/graph/graph-explorer" target="_blank" rel="noopener noreferrer">Graph Explorer</a>.</li>
              <li>Sign in och ge samtycke vid behov.</li>
              <li>Öppna fliken <i>Access token</i> och kopiera token.</li>
              <li>Klicka <i>Visa token</i> här i appen och klistra in (med eller utan "Bearer").</li>
            </ol>
            <div className="muted" style={{ marginTop: 8 }}>
              Obs: För Chat/Möten krävs rätt scopes och ofta admin‑samtycke. Av säkerhetsskäl kan appen inte läsa token automatiskt via inbäddad sida (iframe) eller liknande.
            </div>
            <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <a className="btn btn-secondary" href="https://developer.microsoft.com/graph/graph-explorer" target="_blank" rel="noopener noreferrer">Öppna Graph Explorer</a>
              <button className="btn btn-light" onClick={onToggleTokenCard}>Visa tokenfält</button>
            </div>
          </div>
        </details>
      )}

      {isActive('dashboard') && (
        <details className="card">
          <summary style={{ cursor:'pointer', fontWeight:600 }}>Uppdatera appen</summary>
          <div>
            <p className="muted" style={{ marginTop: 6 }}>Så uppdaterar du till senaste dist‑version:</p>
            <ul style={{ marginTop: 8, color: 'var(--muted)' }}>
              <li>
                Öppna Inställningar → "Uppdatera från GitHub" och klicka <i>Sök och hämta</i>.
                Vid CORS‑fel: använd <i>Direktlänk (ZIP)</i> och därefter <i>Välj ZIP…</i> i appen.
              </li>
              <li>
                Direktnedladdning: <a href="https://github.com/jesper1982atea/graph-user-export/releases/latest/download/graph-user-export-dist.zip" target="_blank" rel="noopener noreferrer">senaste dist‑ZIP</a>.
              </li>
            </ul>
            <div className="muted" style={{ marginTop: 8 }}>
              Tips: I Chrome/Edge kan du välja en målmapp och spara alla filer direkt. Safari saknar detta stöd – spara enskilda filer eller uppdatera manuellt.
            </div>
            <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button className="btn btn-light" onClick={() => goto('/settings')}>Öppna Inställningar →</button>
              <a className="btn btn-secondary" href="https://github.com/jesper1982atea/graph-user-export/releases/latest/download/graph-user-export-dist.zip" target="_blank" rel="noopener noreferrer">Direktlänk (ZIP)</a>
            </div>
          </div>
        </details>
      )}

      {isActive('settings') && (
        <div className="card">
          <b>Inställningar</b>
          <div style={{ marginTop:10, display:'grid', gap:12 }}>
            <label style={{ display:'flex', alignItems:'center', gap:8 }}>
              <input type="checkbox" checked={!!chatCreateEnabled} onChange={onToggleChatCreateEnabled} />
              Tillåt skapande av nya Teams‑chattar (1‑1)
            </label>
            <div className="muted">Scopes i nuvarande token (scp): {tokenScopes || 'okänt'}</div>
            <div className="muted">Dina inställningar sparas lokalt i webbläsaren.</div>
            <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
              <button className="btn btn-light" onClick={() => goto('/settings/labels')}>Hantera rubrikmappning →</button>
            </div>
            <details>
              <summary style={{ cursor:'pointer', fontWeight:600 }}>Hämta token (Graph Explorer)</summary>
              <div>
                <ol style={{ marginTop: 8, color: 'var(--muted)' }}>
                  <li>Gå till <a href="https://developer.microsoft.com/graph/graph-explorer" target="_blank" rel="noopener noreferrer">Graph Explorer</a>.</li>
                  <li>Sign in och ge samtycke vid behov.</li>
                  <li>Öppna fliken <i>Access token</i> och kopiera token.</li>
                  <li>Klicka <i>Visa tokenfält</i> här och klistra in (med eller utan "Bearer").</li>
                </ol>
                <div className="muted" style={{ marginTop: 8 }}>
                  Obs: För Chat/Möten krävs rätt scopes och ofta admin‑samtycke. Av säkerhetsskäl kan appen inte läsa token automatiskt via inbäddad sida (iframe).
                </div>
                <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <a className="btn btn-secondary" href="https://developer.microsoft.com/graph/graph-explorer" target="_blank" rel="noopener noreferrer">Öppna Graph Explorer</a>
                  <button className="btn btn-light" onClick={onToggleTokenCard}>Visa tokenfält</button>
                </div>
              </div>
            </details>
            <InBrowserUpdater />
          </div>
        </div>
      )}
    </div>
  );
};
// import ChatDetailPage from './msgraph/ChatDetailPage'; // Uncomment if you have a ChatDetailPage

// --- Helper components must be defined at top-level scope ---
function MeetingDetails({ meeting, token, onClose }) {
  if (!meeting) return null;
  const joinUrl = meeting.onlineMeeting ? meeting.onlineMeeting.joinUrl : meeting.joinUrl;
  return (
    <div style={{ background: '#fff', border: '1px solid #b3d8ff', borderRadius: 8, margin: '1rem 0', padding: 24 }}>
      <h4>Detaljer för möte</h4>
      <div><b>Ämne:</b> {meeting.subject}</div>
      <div><b>Start:</b> {meeting.start?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
      <div><b>Slut:</b> {meeting.end?.dateTime?.replace('T', ' ').slice(0, 16)}</div>
      <div><b>Plats:</b> {meeting.location?.displayName || '-'}</div>
      {joinUrl && <div><b>Länk till möte:</b> <a href={joinUrl} target="_blank" rel="noopener noreferrer">{joinUrl}</a></div>}
      <div style={{ marginTop: 12 }}>
        <b>Deltagare:</b>
        <AttendeeCards token={token} attendees={meeting.attendees || []} />
      </div>
      <button className="btn btn-secondary" style={{ marginTop: 12 }} onClick={onClose}>Stäng</button>
    </div>
  );
}

function ChatDetails({ chat, token, onClose }) {
  const [members, setMembers] = useState([]);
  useEffect(() => {
    if (!chat || !token) return;
    fetch(`https://graph.microsoft.com/v1.0/chats/${chat.id}/members`, {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then(res => res.ok ? res.json() : Promise.reject('Kunde inte hämta medlemmar'))
      .then(data => setMembers(data.value || []))
      .catch(() => setMembers([]));
  }, [chat, token]);
  return (
    <div style={{ background: '#fff', border: '1px solid #b3d8ff', borderRadius: 8, margin: '1rem 0', padding: 24 }}>
      <h4>Detaljer för chatt</h4>
      <div><b>Rubrik:</b> {chat.topic || 'Ingen rubrik'}</div>
      <div><b>Chatt-typ:</b> {chat.chatType}</div>
      <div><b>Chatt-id:</b> {chat.id}</div>
      <div><b>Länk till chatt:</b> <a href={`https://teams.microsoft.com/l/chat/0/0?users=${members.map(m => m.email || m.userId).join(',')}`} target="_blank" rel="noopener noreferrer">Öppna i Teams</a></div>
      <div style={{ marginTop: 12 }}>
        <b>Medlemmar:</b>
        <div style={{ marginTop: 8 }}>
          <ChatMembers token={token} chatId={chat.id} />
        </div>
      </div>
      <button className="btn btn-secondary" style={{ marginTop: 12 }} onClick={onClose}>Stäng</button>
    </div>
  );
}

function ChatMembers({ token, chatId }) {
  const [members, setMembers] = useState([]);
  const [error, setError] = useState('');
  const [photos, setPhotos] = useState({});
  useEffect(() => {
    if (!token || !chatId) return;
    fetch(`https://graph.microsoft.com/v1.0/chats/${chatId}/members`, {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then(res => res.ok ? res.json() : Promise.reject('Kunde inte hämta medlemmar'))
      .then(data => setMembers(data.value || []))
      .catch(err => setError(err.toString()));
  }, [token, chatId]);

  useEffect(() => {
    if (!token || members.length === 0) return;
    let isMounted = true;
    const fetchPhotos = async () => {
      const newPhotos = {};
      await Promise.all(members.map(async (m) => {
        const userId = m.email || m.userId;
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

  return (
    <div>
      {error && <div style={{ color: 'red' }}>{error}</div>}
      <ul style={{ paddingLeft: '1rem' }}>
        {members.map((m, idx) => {
          const userId = m.email || m.userId;
          return (
            <li key={userId || idx}>
              <img src={photos[userId]} alt="" style={{ width: 24, height: 24, borderRadius: '50%', marginRight: 8 }} />
              {userId}
            </li>
          );
        })}
      </ul>
    </div>
  );
}

function MainGraphMeetingsAndChats(props) {
  // All state and logic from GraphMeetingsAndChats
  // Vi använder intern state för token, möten och chattar

  // Token-hantering i lokal state så fältet inte blir låst
  const [authToken, setAuthToken] = useState(() => {
    try { return localStorage.getItem('graph_token') || ''; } catch { return ''; }
  });
  const [verifyLoading, setVerifyLoading] = useState(false);
  const [tokenVerified, setTokenVerified] = useState(false);
  const [me, setMe] = useState(null);
  const [mePhotoUrl, setMePhotoUrl] = useState('');
  const [error, setError] = useState(null);
  // Token expiry tracking
  const [tokenExp, setTokenExp] = useState(null); // epoch seconds
  const [secondsLeft, setSecondsLeft] = useState(null);
  const [meetings, setMeetings] = useState([]);
  const [chats, setChats] = useState([]);
  const [chatsError, setChatsError] = useState(null);
  const [joinedTeams, setJoinedTeams] = useState([]);
  const [teamsError, setTeamsError] = useState(null);
  const [teamsLoading, setTeamsLoading] = useState(false);
  const [teamChannels, setTeamChannels] = useState({}); // { [teamId]: Channel[] }
  const [teamChannelsLoading, setTeamChannelsLoading] = useState({}); // { [teamId]: boolean }
  const [channelMembers, setChannelMembers] = useState({}); // { [channelId]: { loading, members } }
  const [expandedTeams, setExpandedTeams] = useState({}); // { [teamId]: boolean }
  const [searchInput, setSearchInput] = useState('');
  const [searchLoading, setSearchLoading] = useState(false);
  const [searchedUsers, setSearchedUsers] = useState([]);
  const [selectedUsersMap, setSelectedUsersMap] = useState({}); // { [userId]: true }
  const [createChatLoading, setCreateChatLoading] = useState(false);
  const [createChatError, setCreateChatError] = useState('');
  // Meeting scheduling state
  const [scheduling, setScheduling] = useState(false);
  const [scheduleResult, setScheduleResult] = useState(null);
  const [scheduleError, setScheduleError] = useState('');
  const [scheduleDuration, setScheduleDuration] = useState(30);
  const [scheduleWindowStart, setScheduleWindowStart] = useState('');
  const [scheduleWindowEnd, setScheduleWindowEnd] = useState('');
  // Inställning för att tillåta chatt-skapande (persistens + auto-detektion av scopes)
  const [chatCreateEnabled, setChatCreateEnabled] = useState(() => {
    try {
      const saved = localStorage.getItem('chat_create_enabled');
      if (saved != null) return saved === '1';
    } catch {}
    // Försök inferera från token scopes (scp)
    try {
      const scp = (() => {
        const t = (typeof window !== 'undefined') ? (localStorage.getItem('graph_token') || '') : '';
        const parts = (t || '').split('.');
        if (parts.length < 2) return '';
        const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
        const json = JSON.parse(atob(base64));
        return json.scp || '';
      })();
      if (typeof scp === 'string' && /(^|\s)(Chat\.ReadWrite|Chat\.Create)(\s|$)/.test(scp)) return true;
    } catch {}
    return false;
  });
  const toggleChatCreateEnabled = () => {
    setChatCreateEnabled(prev => {
      const next = !prev;
      try { localStorage.setItem('chat_create_enabled', next ? '1' : '0'); } catch {}
      return next;
    });
  };
  const [theme, setTheme] = useState(() => {
    try {
      const saved = localStorage.getItem('theme');
      if (saved) return saved;
    } catch {}
    return (typeof document !== 'undefined' && document.documentElement.getAttribute('data-theme')) || 'light';
  });
  const [meetingsFrom, setMeetingsFrom] = useState('');
  const [meetingsTo, setMeetingsTo] = useState('');
  const [loadingLists, setLoadingLists] = useState(false);
  const [showTokenCard, setShowTokenCard] = useState(false);
  const handleTokenChange = (e) => {
    const val = e.target.value || '';
    const cleaned = val.replace(/^\s*Bearer\s+/i, '').trim();
    setAuthToken(cleaned);
    try { localStorage.setItem('graph_token', cleaned); } catch {}
  };
  // Parse token expiry on token change
  useEffect(() => {
    try {
      if (!authToken) { setTokenExp(null); setSecondsLeft(null); return; }
      const parts = (authToken || '').split('.');
      if (parts.length < 2) { setTokenExp(null); setSecondsLeft(null); return; }
      const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const json = JSON.parse(atob(base64));
      if (json && typeof json.exp === 'number') {
        setTokenExp(json.exp);
      } else {
        setTokenExp(null);
      }
    } catch { setTokenExp(null); setSecondsLeft(null); }
  }, [authToken]);
  // Update countdown every second when tokenExp is known
  useEffect(() => {
    if (!tokenExp) { setSecondsLeft(null); return; }
    const update = () => {
      const now = Math.floor(Date.now() / 1000);
      setSecondsLeft(Math.max(0, tokenExp - now));
    };
    update();
    const id = setInterval(update, 1000);
    return () => clearInterval(id);
  }, [tokenExp]);
  const verifyToken = async () => {
    if (!authToken) return;
    setVerifyLoading(true);
    setError(null);
    try {
      const trimmed = authToken.trim();
      const header = trimmed.toLowerCase().startsWith('bearer ') ? trimmed : `Bearer ${trimmed}`;
      const res = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: header },
      });
      if (!res.ok) throw new Error(`Verifiering misslyckades (${res.status})`);
      const data = await res.json();
      setMe(data);
      setTokenVerified(true);
      // Also set exp from the verified token (already parsed on change)
      try {
        const parts = (authToken || '').split('.');
        if (parts.length >= 2) {
          const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
          const json = JSON.parse(atob(base64));
          if (json && typeof json.exp === 'number') setTokenExp(json.exp);
        }
      } catch {}
      // Try to fetch profile photo
      try {
        const photoRes = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
          headers: { Authorization: header },
        });
        if (photoRes.ok) {
          const blob = await photoRes.blob();
          const url = URL.createObjectURL(blob);
          setMePhotoUrl(url);
        } else {
          setMePhotoUrl('');
        }
      } catch { setMePhotoUrl(''); }
      try { localStorage.setItem('graph_token_verified', '1'); } catch {}
      // Om användaren inte explicit satt inställningen: auto-sätt utifrån scopes
      try {
        const saved = localStorage.getItem('chat_create_enabled');
        if (saved == null) {
          const scp = (() => {
            const parts = (authToken || '').split('.');
            if (parts.length < 2) return '';
            const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
            const json = JSON.parse(atob(base64));
            return json.scp || '';
          })();
          const allowed = typeof scp === 'string' && /(^|\s)(Chat\.ReadWrite|Chat\.Create)(\s|$)/.test(scp);
          setChatCreateEnabled(allowed);
          try { localStorage.setItem('chat_create_enabled', allowed ? '1' : '0'); } catch {}
        }
      } catch {}
    } catch (err) {
      setTokenVerified(false);
      setMe(null);
  setMePhotoUrl('');
      setError(err.message || 'Verifiering misslyckades');
    } finally {
      setVerifyLoading(false);
    }
  };
  // Auto-verify on load if token exists
  useEffect(() => {
    if (authToken && !tokenVerified && !verifyLoading) {
      verifyToken();
    }
  }, []);

  // Theme toggle
  const ThemeToggle = () => (
    <button className="btn btn-light" onClick={() => {
      const next = theme === 'dark' ? 'light' : 'dark';
      setTheme(next);
      if (typeof document !== 'undefined') document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('theme', next); } catch {}
    }}>
      {theme === 'dark' ? <SunIcon /> : <MoonIcon />} {theme === 'dark' ? 'Light' : 'Dark'}
    </button>
  );
  useEffect(() => {
    if (typeof document !== 'undefined') {
      document.documentElement.setAttribute('data-theme', theme);
    }
  }, [theme]);
  const navigate = useNavigate();
  const location = useLocation();
  const currentPath = location.pathname || '/';
  const [searchDepartment, setSearchDepartment] = useState('');
  const currentPage = currentPath.startsWith('/meetings') ? 'meetings'
    : currentPath.startsWith('/chats') ? 'chats'
    : currentPath.startsWith('/teams') ? 'teams'
    : (currentPath.startsWith('/search') || currentPath.startsWith('/users/')) ? 'search'
    : currentPath.startsWith('/settings') ? 'settings'
    : 'dashboard';
    // Hjälpare för auth-header
    const getAuthHeader = () => {
      const trimmed = authToken.trim();
      return trimmed.toLowerCase().startsWith('bearer ') ? trimmed : `Bearer ${trimmed}`;
    };

    // Återanvändbar laddning av chattar (används av fliken och efter skapande)
    const reloadChats = async () => {
      if (!authToken) { setChats([]); setChatsError('Ingen token.'); return; }
      const header = getAuthHeader();
      try {
        setChatsError(null);
        const res = await fetch('https://graph.microsoft.com/v1.0/me/chats?$top=50', {
          headers: { Authorization: header },
        });
        if (!res.ok) {
          let msg = `Kunde inte hämta chattar (HTTP ${res.status})`;
          try {
            const body = await res.json();
            const detail = body?.error?.message || JSON.stringify(body);
            if (detail) msg += `: ${detail}`;
          } catch {
            try { const text = await res.text(); if (text) msg += `: ${text}`; } catch {}
          }
          setChats([]);
          setChatsError(msg);
          return;
        }
        const data = await res.json();
        setChats(Array.isArray(data.value) ? data.value : []);
      } catch (e) {
        console.warn(e);
        setChats([]);
        setChatsError(e?.message || 'Kunde inte hämta chattar');
      }
    };

    // Reagera på ?department= eller generiska ?field/value (& optional field2/value2) i URL och kör sökning
    useEffect(() => {
      // Support query params either before or after hash (#/search?...)
      let qs = location.search || '';
      if (!qs && typeof window !== 'undefined') {
        const h = window.location.hash || '';
        const qIndex = h.indexOf('?');
        if (qIndex >= 0) qs = h.substring(qIndex);
      }
      const params = new URLSearchParams(qs);
      const dept = params.get('department') || '';
      setSearchDepartment(dept);
      if (dept && tokenVerified && authToken) {
        // Navigera till /search om inte redan där
        try { if (!currentPath.startsWith('/search')) navigate('/search'); } catch {}
        handleUserSearchByDepartment(dept);
      }
      // Generic field search
      const f1 = params.get('field') || '';
      const v1 = params.get('value') || '';
      const f2 = params.get('field2') || '';
      const v2 = params.get('value2') || '';
      if (f1 && v1 && tokenVerified && authToken) {
        try { if (!currentPath.startsWith('/search')) navigate('/search'); } catch {}
        handleUserSearchByFields({ field1: f1, value1: v1, field2: f2, value2: v2 });
      }
  }, [location.search, tokenVerified, authToken]);

    // Händelser för sök
    const handleSearchInputChange = (e) => setSearchInput(e.target.value);
    const handleUserSearch = async () => {
      const selectHeaders = [
        'id','BusinessPhones','CompanyName','mail','jobTitle','UserPrincipalName','givenName','mobilePhone',
        'onPremisesDistinguishedName','onPremisesDomainName','surname','streetAddress','postalCode','physicalDeliveryOfficeName',
        'employeeId','department','country','city','state','onPremisesUserPrincipalName','onPremisesSamAccountName',
        'OfficeLocation','ManagedDevices','Manager','State','assignedPlans','onPremisesExtensionAttributes','provisionedPlans'
      ];
  const selectCamel = GRAPH_USER_SELECT_FIELDS;
      const inputs = (searchInput || '')
        .split(/[\n,\s]+/)
        .map(s => s.trim())
        .filter(Boolean);
      if (inputs.length === 0) return;
      setSearchLoading(true);
      setSearchedUsers([]);
      try {
        // If single non-email query, try a name search via $filter startswith
        if (inputs.length === 1 && !inputs[0].includes('@')) {
          const q = inputs[0].replace(/'/g, "");
          const filter = `startswith(displayName,'${q}') or startswith(givenName,'${q}') or startswith(surname,'${q}') or startswith(userPrincipalName,'${q}')`;
          const res = await fetch(`https://graph.microsoft.com/v1.0/users?$top=25&$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(selectCamel.join(','))}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`, {
            headers: { Authorization: getAuthHeader() },
          });
          if (!res.ok) throw new Error('Sökning misslyckades');
          const data = await res.json();
          const list = Array.isArray(data.value) ? data.value : [];
          const withMgr = list.map(u => ({
            ...u,
            managerDisplayName: u?.manager?.displayName || '',
            managerMail: u?.manager?.mail || '',
            managerUserPrincipalName: u?.manager?.userPrincipalName || '',
            managerJobTitle: u?.manager?.jobTitle || '',
          }));
          setSearchedUsers(withMgr);
        } else {
          const results = await Promise.all(inputs.map(async (val) => {
            try {
              const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(val)}?$select=${encodeURIComponent(selectCamel.join(','))}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`, {
                headers: { Authorization: getAuthHeader() },
              });
              if (!res.ok) throw new Error(`Hittade inte ${val}`);
              const data = await res.json();
              return {
                ...data,
                managerDisplayName: data?.manager?.displayName || '',
                managerMail: data?.manager?.mail || '',
                managerUserPrincipalName: data?.manager?.userPrincipalName || '',
                managerJobTitle: data?.manager?.jobTitle || '',
              };
            } catch (e) {
              return { mail: val, error: e.message || 'Fel vid sökning' };
            }
          }));
          setSearchedUsers(results);
        }
      } catch (e) {
        setSearchedUsers([{ error: e.message || 'Fel vid sökning' }]);
      } finally {
        setSearchLoading(false);
      }
    };

    // Sök användare via avdelning (från query param). Paginera via @odata.nextLink
    const handleUserSearchByDepartment = async (dept) => {
      if (!dept) return;
      setSearchLoading(true);
      setSearchedUsers([]);
      try {
        const header = { Authorization: getAuthHeader() };
        const selectCamel = GRAPH_USER_SELECT_FIELDS;
        const escaped = String(dept).replace(/'/g, "''");
        let url = `https://graph.microsoft.com/v1.0/users?$top=50&$filter=${encodeURIComponent(`department eq '${escaped}'`)}&$select=${encodeURIComponent(selectCamel.join(','))}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`;
        const all = [];
        let guard = 0;
        while (url && guard < 50) {
          guard++;
          const res = await fetch(url, { headers: header });
          if (!res.ok) {
            let msg = `HTTP ${res.status}`;
            try { const j = await res.json(); msg += `: ${j?.error?.message || JSON.stringify(j)}`; } catch {}
            throw new Error(`Sökning (avdelning) misslyckades ${msg}`);
          }
          const j = await res.json();
          const list = Array.isArray(j.value) ? j.value : [];
          list.forEach(u => all.push({
            ...u,
            managerDisplayName: u?.manager?.displayName || '',
            managerMail: u?.manager?.mail || '',
            managerUserPrincipalName: u?.manager?.userPrincipalName || '',
            managerJobTitle: u?.manager?.jobTitle || '',
          }));
          url = j['@odata.nextLink'] || '';
        }
        setSearchedUsers(all);
      } catch (e) {
        setSearchedUsers([{ error: e.message || 'Fel vid sökning (avdelning)' }]);
      } finally {
        setSearchLoading(false);
      }
    };

    // Sök användare via godtyckliga fält-värden (från attributes explorer)
  const handleUserSearchByFields = async ({ field1, value1, field2, value2 }) => {
      if (!field1 || !value1) return;
      setSearchLoading(true);
      setSearchedUsers([]);
      try {
        const header = { Authorization: getAuthHeader() };
        const selectCamel = GRAPH_USER_SELECT_FIELDS;
        const escape = (s) => String(s).replace(/'/g, "''");
    const mapField = (f, v) => {
          if (!f) return '';
          if (f.startsWith('extensionAttribute')) {
            const idx = f.replace('extensionAttribute','');
            return `onPremisesExtensionAttributes/extensionAttribute${idx}`;
          }
          if (f === 'upnDomain') {
            // match domain against UPN or mail domains
      const dom = escape(v);
      return { custom: `endswith(userPrincipalName,'@${dom}') or endswith(mail,'@${dom}')` };
          }
          return f;
        };
    const val1 = escape(value1);
    const part1 = mapField(field1, value1);
        let filter = '';
        if (typeof part1 === 'object' && part1.custom) filter = `(${part1.custom})`;
        else filter = `${part1} eq '${val1}'`;
        if (field2 && value2) {
      const val2 = escape(value2);
      const part2 = mapField(field2, value2);
      if (typeof part2 === 'object' && part2.custom) filter = `${filter} and (${part2.custom})`;
          else filter = `${filter} and ${part2} eq '${val2}'`;
        }
        let url = `https://graph.microsoft.com/v1.0/users?$top=50&$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(selectCamel.join(','))}&$expand=${encodeURIComponent("manager($select=displayName,mail,userPrincipalName,jobTitle)")}`;
        const all = [];
        let guard = 0;
        while (url && guard < 50) {
          guard++;
          const res = await fetch(url, { headers: header });
          if (!res.ok) {
            let msg = `HTTP ${res.status}`;
            try { const j = await res.json(); msg += `: ${j?.error?.message || JSON.stringify(j)}`; } catch {}
            throw new Error(`Sökning (fält) misslyckades ${msg}`);
          }
          const j = await res.json();
          const list = Array.isArray(j.value) ? j.value : [];
          list.forEach(u => all.push({
            ...u,
            managerDisplayName: u?.manager?.displayName || '',
            managerMail: u?.manager?.mail || '',
            managerUserPrincipalName: u?.manager?.userPrincipalName || '',
            managerJobTitle: u?.manager?.jobTitle || '',
          }));
          url = j['@odata.nextLink'] || '';
        }
        setSearchedUsers(all);
      } catch (e) {
        setSearchedUsers([{ error: e.message || 'Fel vid sökning (fält)' }]);
      } finally {
        setSearchLoading(false);
      }
    };

    const onToggleSelectUser = (user) => {
      if (!user?.id) return;
      setSelectedUsersMap(prev => ({ ...prev, [user.id]: !prev[user.id] }));
    };

    const onCreateGroupChat = async () => {
      setCreateChatError('');
      if (!tokenVerified || !authToken) {
        setCreateChatError('Verifiera token först.');
        return;
      }
      if (!chatCreateEnabled) {
        setCreateChatError('Skapande av chattar är avstängt i inställningar.');
        return;
      }
      if (!me?.id) {
        setCreateChatError('Kunde inte läsa din användarprofil (me.id). Verifiera token igen.');
        return;
      }
      const selectedIds = Object.keys(selectedUsersMap).filter(id => selectedUsersMap[id]);
      if (selectedIds.length === 0) {
        setCreateChatError('Välj minst en användare.');
        return;
      }
      const header = getAuthHeader();
      setCreateChatLoading(true);
      const successes = [];
      const failures = [];
      for (const uid of selectedIds) {
        const body = {
          chatType: 'oneOnOne',
          members: [
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: [],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${me.id}')`,
            },
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: [],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${uid}')`,
            },
          ],
        };
        try {
          const res = await fetch('https://graph.microsoft.com/v1.0/chats', {
            method: 'POST',
            headers: { Authorization: header, 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
          });
          if (!res.ok) {
            let msg = `HTTP ${res.status}`;
            try { const j = await res.json(); msg += `: ${j?.error?.message || JSON.stringify(j)}`; } catch {}
            throw new Error(msg);
          }
          await res.json();
          successes.push(uid);
        } catch (e) {
          failures.push({ uid, msg: e.message || 'Okänt fel' });
        }
      }
  // Refresh chats once
  try { await reloadChats(); } catch {}
      // Clear successful selections; keep failed for retry
      setSelectedUsersMap(prev => {
        const next = { ...prev };
        successes.forEach(id => { delete next[id]; });
        return next;
      });
      if (failures.length) {
        const example = failures[0].msg;
        setCreateChatError(`Klar: ${successes.length}/${selectedIds.length} skapade. Misslyckade: ${failures.length}. Exempel: ${example}`);
      } else {
        setCreateChatError('');
        try { if (navigate) navigate('/chats'); } catch {}
      }
      setCreateChatLoading(false);
    };

  // Hämta Teams som användaren är medlem i (med fallback via memberOf)
  const loadJoinedTeams = async () => {
    if (!authToken) return;
    const header = getAuthHeader();
    setTeamsLoading(true);
    try {
      const teamsRes = await fetch('https://graph.microsoft.com/v1.0/me/joinedTeams?$top=50', {
        headers: { Authorization: header },
      });
      if (teamsRes.ok) {
        const teamsData = await teamsRes.json();
        const teams = Array.isArray(teamsData.value) ? teamsData.value : [];
        if (teams.length > 0) {
          setJoinedTeams(teams);
          setTeamsError(null);
          return;
        }
      } else {
        // Read error details and continue to fallback
        try {
          const body = await teamsRes.json();
          const msg = body?.error?.message || JSON.stringify(body);
          setTeamsError(`Kunde inte hämta Teams (HTTP ${teamsRes.status}): ${msg}`);
        } catch {
          try {
            const text = await teamsRes.text();
            if (text) setTeamsError(`Kunde inte hämta Teams (HTTP ${teamsRes.status}): ${text}`);
            else setTeamsError(`Kunde inte hämta Teams (HTTP ${teamsRes.status})`);
          } catch {
            setTeamsError(`Kunde inte hämta Teams (HTTP ${teamsRes.status})`);
          }
        }
      }
      // Fallback: försök hämta grupper och filtrera på Team
      try {
        const groupsRes = await fetch('https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?$select=id,displayName,resourceProvisioningOptions&$top=999', {
          headers: { Authorization: header },
        });
        if (groupsRes.ok) {
          const groupsData = await groupsRes.json();
          const groups = Array.isArray(groupsData.value) ? groupsData.value : [];
          const asTeams = groups
            .filter(g => Array.isArray(g.resourceProvisioningOptions) && g.resourceProvisioningOptions.includes('Team'))
            .map(g => ({ id: g.id, displayName: g.displayName }));
          setJoinedTeams(asTeams);
          if (asTeams.length > 0) setTeamsError(null);
          else setTeamsError(prev => prev || 'Inga Teams hittades via fallback.');
        } else {
          setJoinedTeams([]);
          setTeamsError(prev => prev || `Fallback misslyckades (HTTP ${groupsRes.status})`);
        }
      } catch {
        setJoinedTeams([]);
        setTeamsError(prev => prev || 'Fallback misslyckades.');
      }
    } catch (e) {
      console.warn(e);
      setJoinedTeams([]);
      setTeamsError(prev => prev || (e?.message || 'Okänt fel vid hämtning av Teams'));
    } finally {
      setTeamsLoading(false);
    }
  };

  // Hämta möten, chattar och joined teams när token är verifierad
    useEffect(() => {
      if (!tokenVerified || !authToken) return;
      const header = getAuthHeader();
    const now = new Date();
    const start = meetingsFrom ? new Date(meetingsFrom + 'T00:00:00') : new Date(now);
    const end = meetingsTo ? new Date(meetingsTo + 'T23:59:59') : new Date(now);
    if (!meetingsTo) end.setDate(end.getDate() + 30);
      const startIso = start.toISOString();
      const endIso = end.toISOString();

    const loadMeetings = async () => {
        try {
        setLoadingLists(true);
    const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${encodeURIComponent(startIso)}&endDateTime=${encodeURIComponent(endIso)}&$top=100&$orderby=start/dateTime`;
          const res = await fetch(url, {
            headers: {
              Authorization: header,
              Prefer: 'outlook.timezone="UTC"',
            },
          });
          if (!res.ok) throw new Error('Kunde inte hämta möten');
          const data = await res.json();
          setMeetings(Array.isArray(data.value) ? data.value : []);
        } catch (e) {
          // lämna tomt men spara fel i konsolen
          console.warn(e);
      } finally {
        setLoadingLists(false);
        }
      };

  loadMeetings();
  reloadChats();
  loadJoinedTeams();
    }, [tokenVerified, authToken, meetingsFrom, meetingsTo]);

    // Explicit fetch by Team ID
    const fetchTeamChannels = async (teamId) => {
      if (!teamId) return;
      const header = getAuthHeader();
      setTeamChannelsLoading(prev => ({ ...prev, [teamId]: true }));
      try {
        // Get team info for display name
        let teamName = '';
        try {
          const tRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}?$select=displayName`, {
            headers: { Authorization: header },
          });
          if (tRes.ok) {
            const tData = await tRes.json();
            teamName = tData.displayName || '';
          }
        } catch {}
        // Page through channels (no $top to avoid 400 errors)
  let channelsUrl = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/`;
        const chs = [];
        while (channelsUrl) {
          const res = await fetch(channelsUrl, { headers: { Authorization: header } });
          if (!res.ok) throw new Error('Kunde inte hämta kanaler för team');
          const data = await res.json();
          if (Array.isArray(data.value)) chs.push(...data.value);
          channelsUrl = data['@odata.nextLink'] || null;
        }
        const myId = me?.id;
        const mapped = [];
        for (const ch of chs) {
          if (!ch.membershipType || ch.membershipType === 'standard') {
            mapped.push({ teamId, teamName, channelId: ch.id, channelName: ch.displayName, membershipType: ch.membershipType, description: ch.description });
            continue;
          }
          try {
            // Page through channel members for private/shared channels
            let membersUrl = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(ch.id)}/members`;
            const members = [];
            while (membersUrl) {
              const memRes = await fetch(membersUrl, { headers: { Authorization: header } });
              if (!memRes.ok) break;
              const memData = await memRes.json();
              if (Array.isArray(memData.value)) members.push(...memData.value);
              membersUrl = memData['@odata.nextLink'] || null;
            }
            const amMember = members.some(m => m.userId === myId);
            if (amMember) {
              mapped.push({ teamId, teamName, channelId: ch.id, channelName: ch.displayName, membershipType: ch.membershipType, description: ch.description });
            }
          } catch {}
        }
        setTeamChannels(prev => ({ ...prev, [teamId]: mapped }));
      } catch (e) {
        console.warn(e);
        setTeamChannels(prev => ({ ...prev, [teamId]: [] }));
      } finally {
        setTeamChannelsLoading(prev => ({ ...prev, [teamId]: false }));
      }
    };

    const toggleTeamExpand = async (teamId) => {
      setExpandedTeams(prev => ({ ...prev, [teamId]: !prev[teamId] }));
      const alreadyLoaded = Array.isArray(teamChannels[teamId]);
      if (!alreadyLoaded) {
        await fetchTeamChannels(teamId);
      }
    };

  const toggleChannelMembers = async (teamId, channelId) => {
      setChannelMembers(prev => {
        const curr = prev[channelId] || { loading: false, members: null };
        // If already loaded, collapse by setting members to null; if null, we will fetch
        return { ...prev, [channelId]: { ...curr, members: curr.members ? null : curr.members } };
      });
      // If members now null, fetch
      const state = channelMembers[channelId];
      const needFetch = !state || state.members == null;
      if (!needFetch) return;
      const header = getAuthHeader();
      setChannelMembers(prev => ({ ...prev, [channelId]: { loading: true, members: null } }));
      try {
        let url = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/members/`;
        const all = [];
        while (url) {
          const res = await fetch(url, { headers: { Authorization: header } });
          if (!res.ok) {
            try {
              const body = await res.json();
              console.warn('Kanalmedlemmar fel:', res.status, body?.error?.message || body);
            } catch {}
            break;
          }
          const data = await res.json();
          if (Array.isArray(data.value)) all.push(...data.value);
          url = data['@odata.nextLink'] || null;
        }
        setChannelMembers(prev => ({ ...prev, [channelId]: { loading: false, members: all } }));
      } catch (e) {
        console.warn(e);
        setChannelMembers(prev => ({ ...prev, [channelId]: { loading: false, members: [] } }));
      }
    };
    // Export helpers
    const toCsv = (items) => {
      if (!Array.isArray(items) || items.length === 0) return '';
      // collect keys
      const keys = Array.from(items.reduce((s, it) => {
        Object.keys(it || {}).forEach(k => s.add(k));
        return s;
      }, new Set()));
      const header = keys.join(',');
      const rows = items.map(it => keys.map(k => {
        let v = it[k];
        if (typeof v === 'object' && v !== null) {
          try { v = JSON.stringify(v); } catch { v = ''; }
        }
        v = v == null ? '' : String(v);
        return '"' + v.replace(/"/g, '""') + '"';
      }).join(','));
      return [header, ...rows].join('\n');
    };
    const exportMeetings = () => {
      const csv = toCsv(meetings);
      if (!csv) return;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, 'moten.csv');
    };
    const exportChats = () => {
      const csv = toCsv(chats);
      if (!csv) return;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, 'chattar.csv');
    };
    const exportTeamChannels = (teamId) => {
      const items = teamChannels[teamId] || [];
      const csv = toCsv(items);
      if (!csv) return;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, `team_${teamId}_kanaler.csv`);
    };
    const exportChannels = (items) => {
      const csv = toCsv(items || channels);
      if (!csv) return;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, 'kanaler.csv');
    };
    const exportUsers = (users) => {
      if (!Array.isArray(users) || users.length === 0) return;
      const headers = [
        'id','BusinessPhones','CompanyName','mail','jobTitle','UserPrincipalName','givenName','mobilePhone',
        'onPremisesDistinguishedName','onPremisesDomainName','surname','streetAddress','postalCode','physicalDeliveryOfficeName',
        'employeeId','department','country','city','state','onPremisesUserPrincipalName','onPremisesSamAccountName',
        'OfficeLocation','ManagedDevices','Manager','State','assignedPlans','onPremisesExtensionAttributes','provisionedPlans'
      ];
      const getVal = (u, h) => {
        switch (h) {
          case 'id': return u.id || '';
          case 'BusinessPhones': return Array.isArray(u.businessPhones) ? u.businessPhones.join('; ') : (u.businessPhones || '');
          case 'CompanyName': return u.companyName || '';
          case 'mail': return u.mail || u.email || u.userPrincipalName || '';
          case 'jobTitle': return u.jobTitle || '';
          case 'UserPrincipalName': return u.userPrincipalName || '';
          case 'givenName': return u.givenName || '';
          case 'mobilePhone': return u.mobilePhone || '';
          case 'onPremisesDistinguishedName': return u.onPremisesDistinguishedName || '';
          case 'onPremisesDomainName': return u.onPremisesDomainName || '';
          case 'surname': return u.surname || '';
          case 'streetAddress': return u.streetAddress || '';
          case 'postalCode': return u.postalCode || '';
          case 'physicalDeliveryOfficeName': return u.physicalDeliveryOfficeName || '';
          case 'employeeId': return u.employeeId || '';
          case 'department': return u.department || '';
          case 'country': return u.country || '';
          case 'city': return u.city || '';
          case 'state': return u.state || '';
          case 'onPremisesUserPrincipalName': return u.onPremisesUserPrincipalName || '';
          case 'onPremisesSamAccountName': return u.onPremisesSamAccountName || '';
          case 'OfficeLocation': return u.officeLocation || '';
          case 'ManagedDevices': return u.managedDevices ? JSON.stringify(u.managedDevices) : '';
          case 'Manager':
            return u.manager ? (u.manager.displayName || (typeof u.manager === 'string' ? u.manager : JSON.stringify(u.manager))) : '';
          case 'State': return u.state || '';
          case 'assignedPlans': return u.assignedPlans ? JSON.stringify(u.assignedPlans) : '';
          case 'onPremisesExtensionAttributes': return u.onPremisesExtensionAttributes ? JSON.stringify(u.onPremisesExtensionAttributes) : '';
          case 'provisionedPlans': return u.provisionedPlans ? JSON.stringify(u.provisionedPlans) : '';
          default: return '';
        }
      };
      const csvRows = [headers.map(h => '"' + h.replace(/"/g, '""') + '"').join(',')];
      users.forEach(u => {
        const row = headers.map(h => {
          const v = getVal(u, h);
          const s = v == null ? '' : String(v);
          return '"' + s.replace(/"/g, '""') + '"';
        }).join(',');
        csvRows.push(row);
      });
      const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, 'anvandare.csv');
    };
    const exportChannelMembers = (channelId) => {
      const entry = channelMembers[channelId];
      if (!entry || !Array.isArray(entry.members) || entry.members.length === 0) return;
      const csv = toCsv(entry.members);
      if (!csv) return;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, `kanal_${channelId}_medlemmar.csv`);
    };
    const clearToken = () => {
      setAuthToken('');
      setTokenVerified(false);
      setMe(null);
      setError(null);
  setTokenExp(null);
  setSecondsLeft(null);
      try { localStorage.removeItem('graph_token'); localStorage.removeItem('graph_token_verified'); } catch {}
    };
  const onApplyDateFilter = () => {
    // Triggered by changing meetingsFrom/meetingsTo
    // useEffect already reacts to state and reloads
  };
  const detailedMeeting = null;
  const detailedChat = null;

  // Settings page component (inline) to keep everything in this file
  // Settings rendered inline within HomePage's 'settings' tab

  // ...existing code for all logic and rendering...

  const resetToDashboard = () => {
    try { if (navigate) navigate('/dashboard'); } catch {}
    setError(null);
  };

  return (
    <ErrorBoundary onReset={resetToDashboard}>
    <Routes>
  <Route path="/teams/:teamId" element={<TeamDetailPage token={authToken} />} />
      <Route
        path="/users/:userId"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={null}
            navigate={navigate}
            chats={chats}
            detailedChat={null}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
            showUserDetail
          />
        }
      />
      <Route
        path="/settings"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
      <Route
        path="/settings/labels"
        element={
          <FieldLabelMappingPage
            items={searchedUsers && searchedUsers.length ? searchedUsers : (me ? [me] : [])}
            onBack={() => navigate && navigate('/settings')}
          />
        }
      />
      <Route
        path="/meetings/:meetingId"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={null}
            navigate={navigate}
            chats={chats}
            detailedChat={null}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
            showMeetingDetail
          />
        }
      />
      <Route
        path="/chats/:chatId"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={null}
            navigate={navigate}
            chats={chats}
            detailedChat={null}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
            showChatDetail
          />
        }
      />
      <Route
        path="/dashboard"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
      <Route
        path="/meetings"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
      <Route
        path="/chats"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
      <Route
        path="/teams"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
      <Route
        path="/search"
        element={
          <HomePage
            token={authToken}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            searchDepartment={searchDepartment}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            currentPage={currentPage}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
        <Route
          path="/departments"
          element={
            <HomePage
              token={authToken}
              handleTokenChange={handleTokenChange}
              verifyToken={verifyToken}
              verifyLoading={verifyLoading}
              tokenVerified={tokenVerified}
              secondsLeft={secondsLeft}
              ThemeToggle={ThemeToggle}
              showTokenCard={showTokenCard}
              onToggleTokenCard={() => setShowTokenCard(v => !v)}
              navigate={navigate}
              currentPage={'departments'}
              // Render DepartmentsPage within content below
            />
          }
        />
      {/* Lägg till motsvarande Route för ChatDetailPage om du har den */}
      <Route
        path="/"
        element={
          <HomePage
            token={authToken}
            currentPage={currentPage}
            handleTokenChange={handleTokenChange}
            verifyToken={verifyToken}
            verifyLoading={verifyLoading}
            tokenVerified={tokenVerified}
            secondsLeft={secondsLeft}
            ThemeToggle={ThemeToggle}
            showTokenCard={showTokenCard}
            onToggleTokenCard={() => setShowTokenCard(v => !v)}
            searchInput={searchInput}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            searchDepartment={searchDepartment}
            me={me}
            mePhotoUrl={mePhotoUrl}
            searchedUsers={searchedUsers}
            selectedUsersMap={selectedUsersMap}
            onToggleSelectUser={onToggleSelectUser}
            onCreateGroupChat={onCreateGroupChat}
            createChatLoading={createChatLoading}
            createChatError={createChatError}
            error={error}
            meetings={meetings}
            meetingsFrom={meetingsFrom}
            meetingsTo={meetingsTo}
            onChangeFrom={setMeetingsFrom}
            onChangeTo={setMeetingsTo}
            onApplyDateFilter={onApplyDateFilter}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
            chatsError={chatsError}
            onReloadChats={reloadChats}
            onExportMeetings={exportMeetings}
            onExportChats={exportChats}
            onExportUsers={exportUsers}
            onClearToken={clearToken}
            loadingLists={loadingLists}
            joinedTeams={joinedTeams}
            teamsError={teamsError}
            teamsLoading={teamsLoading}
            expandedTeams={expandedTeams}
            onToggleTeamExpand={toggleTeamExpand}
            teamChannels={teamChannels}
            onExportTeamChannels={exportTeamChannels}
            channelMembers={channelMembers}
            onToggleChannelMembers={toggleChannelMembers}
            onExportChannelMembers={exportChannelMembers}
            onReloadTeams={loadJoinedTeams}
            chatCreateEnabled={chatCreateEnabled}
            onToggleChatCreateEnabled={toggleChatCreateEnabled}
          />
        }
      />
  </Routes>
  </ErrorBoundary>
  );
}


export default MainGraphMeetingsAndChats;