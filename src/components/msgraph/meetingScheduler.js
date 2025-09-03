// Utilities for scheduling meetings with Microsoft Graph

function buildAuthHeader(token) {
  const trimmed = (token || '').trim();
  return trimmed.toLowerCase().startsWith('bearer ') ? trimmed : `Bearer ${trimmed}`;
}

function toIso(dt) {
  if (typeof dt === 'string') return dt;
  return new Date(dt).toISOString();
}

function splitAttendeesByDomain(emails, tenantDomain) {
  const internal = [];
  const external = [];
  (emails || []).forEach(e => {
    const addr = (e || '').trim();
    if (!addr) return;
    const domain = addr.split('@')[1] || '';
    if (tenantDomain && domain.toLowerCase() === tenantDomain.toLowerCase()) internal.push(addr);
    else external.push(addr);
  });
  return { internal, external };
}

export async function findFirstAvailable({ token, attendeeEmails, durationMinutes = 30, windowStart, windowEnd, tenantDomain, workHoursOnly = false }) {
  if (!token || !Array.isArray(attendeeEmails) || attendeeEmails.length === 0) {
    throw new Error('Saknar token eller deltagare');
  }
  const Authorization = buildAuthHeader(token);

  const start = windowStart ? new Date(windowStart) : new Date();
  const end = windowEnd ? new Date(windowEnd) : new Date(Date.now() + 7*24*60*60*1000);
  const timeConstraint = {
    timeslots: [
      {
        start: { dateTime: toIso(start), timeZone: 'UTC' },
        end: { dateTime: toIso(end), timeZone: 'UTC' },
      },
    ],
  };
  if (workHoursOnly) {
    timeConstraint.activityDomain = 'work';
  }
  const { internal, external } = splitAttendeesByDomain(attendeeEmails, tenantDomain);

  const makeAttendees = (reqList = [], optList = []) => [
    ...reqList.map(a => ({ emailAddress: { address: a }, type: 'required' })),
    ...optList.map(a => ({ emailAddress: { address: a }, type: 'optional' })),
  ];

  const meetingDuration = `PT${Math.max(15, Math.min(240, durationMinutes))}M`;

  // Attempt 1: everyone required, 100% attendance
  const body1 = {
    attendees: makeAttendees(attendeeEmails, []),
    timeConstraint,
    meetingDuration,
    maxCandidates: 10,
    isOrganizerOptional: false,
    returnSuggestionReasons: true,
    minimumAttendeePercentage: 1.0,
  };

  const postFind = async (payload) => {
    const res = await fetch('https://graph.microsoft.com/v1.0/me/findMeetingTimes', {
      method: 'POST',
      headers: { Authorization, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      throw new Error(`findMeetingTimes misslyckades (${res.status}): ${t}`);
    }
    const data = await res.json();
    const list = Array.isArray(data.meetingTimeSuggestions) ? data.meetingTimeSuggestions : [];
    return list;
  };

  let suggestions = [];
  try {
    suggestions = await postFind(body1);
  } catch (e) {
    // Continue to fallbacks
  }

  // Fallback 1: externals optional, require internals (if any). Lower minimum attendee percentage slightly
  if (!suggestions.length && (external.length && internal.length)) {
    const body2 = {
      attendees: makeAttendees(internal, external),
      timeConstraint,
      meetingDuration,
      maxCandidates: 10,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
      minimumAttendeePercentage: Math.min(1.0, Math.max(0.5, internal.length / (internal.length + external.length))),
    };
    try { suggestions = await postFind(body2); } catch {}
  }

  // Fallback 2: only internal participants
  if (!suggestions.length && internal.length) {
    const body3 = {
      attendees: makeAttendees(internal, []),
      timeConstraint,
      meetingDuration,
      maxCandidates: 10,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
      minimumAttendeePercentage: 1.0,
    };
    try { suggestions = await postFind(body3); } catch {}
  }

  if (!suggestions.length) return null;
  // Pick earliest by start time
  suggestions.sort((a, b) => new Date(a.meetingTimeSlot.start.dateTime) - new Date(b.meetingTimeSlot.start.dateTime));
  const s = suggestions[0];
  return {
    start: s.meetingTimeSlot.start, // {dateTime, timeZone}
    end: s.meetingTimeSlot.end,
    confidence: s.confidence,
    suggestion: s,
  };
}

export async function createMeeting({ token, subject, attendeeEmails, start, end, bodyHtml = '', isOnline = true, locationDisplayName = '', recurrence = null }) {
  if (!token) throw new Error('Saknar token');
  if (!start || !end) throw new Error('Saknar start/slut');
  const Authorization = buildAuthHeader(token);
  const attendees = (attendeeEmails || []).map(a => ({ emailAddress: { address: a }, type: 'required' }));
  const payload = {
    subject: subject || 'Möte',
    body: { contentType: 'HTML', content: bodyHtml || '' },
    start: { dateTime: toIso(start.dateTime || start), timeZone: start.timeZone || 'UTC' },
    end: { dateTime: toIso(end.dateTime || end), timeZone: end.timeZone || 'UTC' },
    location: locationDisplayName ? { displayName: locationDisplayName } : undefined,
    attendees,
    isOnlineMeeting: !!isOnline,
    onlineMeetingProvider: isOnline ? 'teamsForBusiness' : undefined,
  };
  if (recurrence) {
    payload.recurrence = recurrence;
  }
  const res = await fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: { Authorization, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`Skapande av möte misslyckades (${res.status}): ${t}`);
  }
  return res.json();
}
