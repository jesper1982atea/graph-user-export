// Enkel HomePage-komponent för routing
function HomePage(props) {
  return (
    <div style={{ padding: '2rem' }}>
      <h2>Välkommen till GraphMeetingsAndChats</h2>
      <p>Den här sidan är under utveckling. Logga in och navigera via menyn.</p>
      {/* Här kan du lägga till mer innehåll eller använda props */}
    </div>
  );
}
import React, { useState, useEffect } from 'react';
import UserCard from './components/msgraph/UserCard';
import { saveAs } from 'file-saver';
import { BrowserRouter as Router, Route, Routes, useNavigate, useParams } from 'react-router-dom';
import MeetingDetailPage from './components/msgraph/MeetingDetailPage';
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
  const {
    token,
    setToken,
    meetings,
    setMeetings,
    chats,
    setChats
  } = props;
  const navigate = useNavigate();

  // Stubbar för props till HomePage
  const handleTokenChange = () => {};
  const handleSearchInputChange = () => {};
  const handleUserSearch = () => {};
  const searchLoading = false;
  const me = null;
  const searchedUsers = [];
  const error = null;
  const detailedMeeting = null;
  const detailedChat = null;

  return (
    <Routes>
      <Route
        path="/meetings/:meetingId"
        element={<MeetingDetailPage meetings={meetings} token={token} />}
      />
      {/* Lägg till motsvarande Route för ChatDetailPage om du har den */}
      <Route
        path="/"
        element={
          <HomePage
            token={token}
            handleTokenChange={handleTokenChange}
            searchInput={""}
            handleSearchInputChange={handleSearchInputChange}
            handleUserSearch={handleUserSearch}
            searchLoading={searchLoading}
            me={me}
            searchedUsers={searchedUsers}
            error={error}
            meetings={meetings}
            detailedMeeting={detailedMeeting}
            navigate={navigate}
            chats={chats}
            detailedChat={detailedChat}
          />
        }
      />
    </Routes>
  );
}


export default MainGraphMeetingsAndChats;