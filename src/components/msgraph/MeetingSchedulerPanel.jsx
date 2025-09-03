import React, { useEffect, useMemo, useState } from 'react';
import { findFirstAvailable, createMeeting } from './meetingScheduler';

/**
 * Reusable meeting scheduling panel
 * Props:
 * - token: Graph token
 * - attendeeEmails?: string[]
 * - getAttendeeEmails?: () => Promise<string[]>
 * - defaultSubject?: string
 * - defaultBody?: string
 * - defaultOnline?: boolean
 * - defaultDurationMinutes?: number
 * - defaultWindowStart?: string (ISO or datetime-local)
 * - defaultWindowEnd?: string
 * - tenantDomain?: string
 * - contextKey?: string (persist UI state per context)
 * - title?: string (label)
 * - onBooked?: (eventJson) => void
 */
export default function MeetingSchedulerPanel({
  token,
  attendeeEmails,
  getAttendeeEmails,
  defaultSubject = 'Möte',
  defaultBody = '',
  defaultOnline = true,
  defaultDurationMinutes = 30,
  defaultWindowStart = '',
  defaultWindowEnd = '',
  tenantDomain,
  contextKey = 'schedule:default',
  title = 'Mötesbokning',
  onBooked,
}) {
  const localTz = useMemo(() => Intl.DateTimeFormat().resolvedOptions().timeZone || 'Local', []);
  const storageKey = (suffix) => `${contextKey}:${suffix}`;
  const [subject, setSubject] = useState(() => getLs(storageKey('subject'), defaultSubject));
  const [body, setBody] = useState(() => getLs(storageKey('body'), defaultBody));
  const [online, setOnline] = useState(() => getLs(storageKey('online'), defaultOnline ? '1' : '0') === '1');
  const [duration, setDuration] = useState(() => parseInt(getLs(storageKey('duration'), String(defaultDurationMinutes)), 10) || defaultDurationMinutes);
  const [windowStart, setWindowStart] = useState(() => normalizeToLocalInput(getLs(storageKey('winStart'), defaultWindowStart)));
  const [windowEnd, setWindowEnd] = useState(() => normalizeToLocalInput(getLs(storageKey('winEnd'), defaultWindowEnd)));
  const [scheduling, setScheduling] = useState(false);
  const [scheduleResult, setScheduleResult] = useState(null);
  const [error, setError] = useState('');
  const [workHoursOnly, setWorkHoursOnly] = useState(() => getLs(storageKey('workHours'),'1')==='1');
  const [recurrenceMode, setRecurrenceMode] = useState(() => getLs(storageKey('recurMode'),'none'));
  const [recurInterval, setRecurInterval] = useState(() => parseInt(getLs(storageKey('recurInt'),'2'),10)||2);
  const [recurCount, setRecurCount] = useState(() => parseInt(getLs(storageKey('recurCount'),'6'),10)||6);

  useEffect(() => { setLs(storageKey('subject'), subject); }, [subject]);
  useEffect(() => { setLs(storageKey('body'), body); }, [body]);
  useEffect(() => { setLs(storageKey('online'), online ? '1' : '0'); }, [online]);
  useEffect(() => { setLs(storageKey('duration'), String(duration)); }, [duration]);
  useEffect(() => { setLs(storageKey('winStart'), windowStart || ''); }, [windowStart]);
  useEffect(() => { setLs(storageKey('winEnd'), windowEnd || ''); }, [windowEnd]);
  useEffect(() => { setLs(storageKey('workHours'), workHoursOnly?'1':'0'); }, [workHoursOnly]);
  useEffect(() => { setLs(storageKey('recurMode'), recurrenceMode||'none'); }, [recurrenceMode]);
  useEffect(() => { setLs(storageKey('recurInt'), String(recurInterval)); }, [recurInterval]);
  useEffect(() => { setLs(storageKey('recurCount'), String(recurCount)); }, [recurCount]);

  const resolveEmails = async () => {
    if (Array.isArray(attendeeEmails) && attendeeEmails.length) return attendeeEmails;
    if (typeof getAttendeeEmails === 'function') {
      const list = await getAttendeeEmails();
      return (list || []).filter(Boolean);
    }
    return [];
  };

  const onFind = async () => {
    setError('');
    setScheduleResult(null);
    setScheduling(true);
    try {
      const emails = await resolveEmails();
      if (!emails.length) { setError('Inga deltagare att boka med.'); return; }
      const suggestion = await findFirstAvailable({
        token,
        attendeeEmails: emails,
        durationMinutes: Number(duration) || 30,
        windowStart: windowStart ? toUTCISOStringFromLocalInput(windowStart) : undefined,
        windowEnd: windowEnd ? toUTCISOStringFromLocalInput(windowEnd) : undefined,
        tenantDomain,
        workHoursOnly,
      });
      setScheduleResult(suggestion);
      if (!suggestion) setError('Ingen gemensam tid hittades.');
    } catch (e) {
      setError(e.message || 'Kunde inte hitta tid.');
    } finally { setScheduling(false); }
  };

  const onBook = async () => {
    if (!scheduleResult) return;
    try {
      const emails = await resolveEmails();
      const recurrence = makeRecurrence(recurrenceMode, recurInterval, recurCount, scheduleResult);
      const ev = await createMeeting({
        token,
        subject: subject || 'Möte',
        attendeeEmails: emails,
        start: scheduleResult.start,
        end: scheduleResult.end,
        bodyHtml: body || '',
        isOnline: !!online,
        recurrence,
      });
      onBooked && onBooked(ev);
      alert('Möte skapat.');
    } catch (e) {
      alert(e.message || 'Kunde inte skapa möte');
    }
  };

  return (
    <div className="card" style={{ marginTop: 8 }}>
      <div className="section-title" style={{ marginBottom: 8 }}>
        <b>{title}</b>
        <span className="spacer" />
        <span className="muted" style={{ fontSize: '.9rem' }}>Tidszon: {localTz}</span>
      </div>

      <div style={{ display:'grid', gridTemplateColumns:'minmax(280px,1fr) minmax(280px,1fr)', gap:16, alignItems:'start' }}>
        {/* Vänster kolumn: Ämne & Detaljer */}
        <div>
          <div className="muted" style={{ fontWeight:600, marginBottom:6 }}>Detaljer</div>
          <div className="muted" style={{ marginBottom:6 }}>Ämne</div>
          <input type="text" value={subject} onChange={e => setSubject(e.target.value)} style={{ width:'100%' }} placeholder="Mötets ämne" />
          <div className="muted" style={{ marginTop:12, marginBottom:6 }}>Mötesdetaljer (HTML)</div>
          <textarea value={body} onChange={e => setBody(e.target.value)} rows={4} style={{ width:'100%' }} placeholder="Meddelande/beskrivning" />
          <label className="muted" title="Skapa som Teams-möte" style={{ display:'inline-flex', alignItems:'center', gap:6, marginTop:10 }}>
            <input type="checkbox" checked={online} onChange={e => setOnline(e.target.checked)} /> Skapa som Teams‑möte
          </label>
          <div style={{ marginTop:10, display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
            <div>
              <div className="muted" style={{ marginBottom:6 }}>Återkommande</div>
              <select value={recurrenceMode} onChange={e=>setRecurrenceMode(e.target.value)} style={{ width:'100%' }}>
                <option value="none">Ingen</option>
                <option value="weekly">Veckovis</option>
                <option value="biweekly">Varrannan vecka</option>
                <option value="monthly">Månadsvis</option>
                <option value="quarterly">Kvartalsvis</option>
              </select>
              <div className="muted" style={{ marginTop:6, fontSize:'.9rem' }}>{recurrenceSummarySv(recurrenceMode, recurInterval, recurCount, scheduleResult)}</div>
            </div>
            <div>
              <div className="muted" style={{ marginBottom:6 }}>Intervall / Antal</div>
              <div style={{ display:'flex', gap:8, alignItems:'center' }}>
                <input
                  type="number"
                  min="1"
                  max={recurrenceMode==='weekly'||recurrenceMode==='biweekly'?"8":"12"}
                  value={recurInterval}
                  onChange={e=>setRecurInterval(parseInt(e.target.value||'1',10))}
                  style={{ width:90 }}
                  title="Intervall"
                  disabled={recurrenceMode==='none'||recurrenceMode==='biweekly'||recurrenceMode==='quarterly'}
                />
                <span className="muted" style={{ fontSize:'.9rem' }}>
                  {recurrenceMode==='weekly'||recurrenceMode==='biweekly' ? 'veckor' : 'månader'}
                </span>
                <input
                  type="number"
                  min="1"
                  max="52"
                  value={recurCount}
                  onChange={e=>setRecurCount(parseInt(e.target.value||'1',10))}
                  style={{ width:110 }}
                  title="Antal förekomster"
                  disabled={recurrenceMode==='none'}
                />
              </div>
              {(recurrenceMode==='biweekly'||recurrenceMode==='quarterly') && (
                <div className="muted" style={{ marginTop:4, fontSize:'.85rem' }}>
                  {recurrenceMode==='biweekly' ? 'Fast intervall: 2 veckor' : 'Fast intervall: 3 månader'}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Höger kolumn: Tidfönster */}
        <div>
          <div className="muted" style={{ fontWeight:600, marginBottom:6 }}>Tid & tillgänglighet</div>
          <div className="muted" style={{ marginBottom:6 }}>Längd (minuter)</div>
          <div style={{ display:'flex', gap:8, alignItems:'center' }}>
            <input type="number" min="15" max="240" step="5" value={duration} onChange={e => setDuration(e.target.value)} style={{ width: 100 }} />
            <div className="muted">Snabbval:</div>
            {[15,30,45,60].map(v => (
              <button key={v} className={`btn btn-light ${Number(duration)===v?'active':''}`} onClick={() => setDuration(v)}>{v}</button>
            ))}
          </div>

          <div style={{ marginTop:12 }}>
            <div className="muted" style={{ marginBottom:6, display:'flex', alignItems:'center', gap:8 }}>
              <span>Tidfönster (lokal tid)</span>
              <label className="muted" title="Endast arbetstid (Graph findMeetingTimes)">
                <input type="checkbox" checked={workHoursOnly} onChange={e=>setWorkHoursOnly(e.target.checked)} /> Arbetstid 08–17
              </label>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
              <div>
                <div className="muted" style={{ marginBottom:4 }}>Från</div>
                <input type="datetime-local" value={windowStart} onChange={e => setWindowStart(e.target.value)} style={{ width:'100%' }} />
              </div>
              <div>
                <div className="muted" style={{ marginBottom:4 }}>Till</div>
                <input type="datetime-local" value={windowEnd} onChange={e => setWindowEnd(e.target.value)} style={{ width:'100%' }} />
              </div>
            </div>
            <div style={{ display:'flex', gap:8, flexWrap:'wrap', marginTop:8 }}>
              <button className="btn btn-light" onClick={() => applyWindowPreset('today', setWindowStart, setWindowEnd)}>Idag</button>
              <button className="btn btn-light" onClick={() => applyWindowPreset('tomorrowAM', setWindowStart, setWindowEnd)}>Imorgon (förmiddag)</button>
              <button className="btn btn-light" onClick={() => applyWindowPreset('nextWeek', setWindowStart, setWindowEnd)}>Nästa vecka</button>
              <button className="btn" onClick={() => { setWindowStart(''); setWindowEnd(''); }}>Rensa</button>
            </div>
          </div>

          <div style={{ marginTop:16, display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
            <button className="btn btn-secondary" disabled={scheduling} onClick={onFind}>{scheduling ? 'Söker tid…' : 'Hitta första lediga tid'}</button>
            {error && <span className="muted" style={{ color:'#b91c1c' }}>{error}</span>}
          </div>

          {scheduleResult && (
            <div className="card" style={{ marginTop:12, padding:10 }}>
              <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                <span className="badge badge-info">Förslag</span>
                <b>{formatSuggestionLocal(scheduleResult)}</b>
                <span className="muted">({localTz})</span>
                {typeof scheduleResult.confidence === 'number' && (
                  <span className="muted">· Träffsäkerhet {Math.round(scheduleResult.confidence*100)}%</span>
                )}
              </div>
              <div style={{ display:'flex', gap:12, alignItems:'center', flexWrap:'wrap', marginTop:8 }}>
                <span className="badge badge-light">{Number(duration)||30} min</span>
                {online && <span className="badge badge-light">Teams</span>}
                {recurrenceMode!=="none" && (
                  <span className="muted">{recurrenceSummarySv(recurrenceMode, recurInterval, recurCount, scheduleResult)}</span>
                )}
                <span className="spacer" />
                <button className="btn btn-primary" onClick={onBook}>Boka möte</button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function getLs(key, fallback) {
  try { const v = localStorage.getItem(key); return v == null ? fallback : v; } catch { return fallback; }
}
function setLs(key, value) {
  try { localStorage.setItem(key, value); } catch {}
}

// Helpers for local time handling
function pad(n) { return String(n).padStart(2, '0'); }
function toLocalInputValue(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = pad(d.getMonth() + 1);
  const da = pad(d.getDate());
  const h = pad(d.getHours());
  const mi = pad(d.getMinutes());
  return `${y}-${m}-${da}T${h}:${mi}`;
}
function normalizeToLocalInput(val) {
  if (!val) return '';
  // If looks like ISO (ends with Z or has seconds), convert to local input.
  if (/Z$/.test(val) || /T\d{2}:\d{2}:\d{2}/.test(val)) {
    try { return toLocalInputValue(val); } catch { return ''; }
  }
  return val; // assume already in datetime-local format
}
function toUTCISOStringFromLocalInput(localStr) {
  try { return new Date(localStr).toISOString(); } catch { return undefined; }
}
function formatSuggestionLocal(sugg) {
  const s = sugg?.start; const e = sugg?.end;
  if (!s || !e) return '';
  // If Graph provided UTC (likely as configured), append Z to parse correctly
  const sIso = s.timeZone && s.timeZone.toUpperCase() === 'UTC' && !/Z$/.test(s.dateTime) ? `${s.dateTime}Z` : s.dateTime;
  const eIso = e.timeZone && e.timeZone.toUpperCase() === 'UTC' && !/Z$/.test(e.dateTime) ? `${e.dateTime}Z` : e.dateTime;
  try {
    const sd = new Date(sIso);
    const ed = new Date(eIso);
    const y = sd.getFullYear(); const m = pad(sd.getMonth()+1); const d = pad(sd.getDate());
    const sh = pad(sd.getHours()); const sm = pad(sd.getMinutes());
    const eh = pad(ed.getHours()); const em = pad(ed.getMinutes());
    return `${y}-${m}-${d} ${sh}:${sm}–${eh}:${em}`;
  } catch {
    return `${s.dateTime?.replace('T',' ').slice(0,16)}–${e.dateTime?.replace('T',' ').slice(0,16)}`;
  }
}

function applyWindowPreset(preset, setStart, setEnd) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  if (preset === 'today') {
    start.setHours(8,0,0,0);
    end.setHours(17,0,0,0);
  } else if (preset === 'tomorrowAM') {
    const t = new Date(now.getFullYear(), now.getMonth(), now.getDate()+1);
    start.setFullYear(t.getFullYear(), t.getMonth(), t.getDate()); start.setHours(8,0,0,0);
    end.setFullYear(t.getFullYear(), t.getMonth(), t.getDate()); end.setHours(12,0,0,0);
  } else if (preset === 'nextWeek') {
    const day = now.getDay(); // 0=Sun
    const daysToMon = (8 - day) % 7 || 1; // next Monday
    const mon = new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysToMon);
    start.setFullYear(mon.getFullYear(), mon.getMonth(), mon.getDate()); start.setHours(8,0,0,0);
    const fri = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 4);
    end.setFullYear(fri.getFullYear(), fri.getMonth(), fri.getDate()); end.setHours(17,0,0,0);
  }
  setStart(toLocalInputValue(start));
  setEnd(toLocalInputValue(end));
}

// Recurrence helpers
function graphDateToLocalDate(dt) {
  if (!dt) return null;
  const tz = dt.timeZone || '';
  const raw = dt.dateTime;
  if (!raw) return null;
  const iso = tz.toUpperCase() === 'UTC' && !/Z$/.test(raw) ? `${raw}Z` : raw;
  const d = new Date(iso);
  return isNaN(d) ? null : d;
}
function dayNameSv(date) {
  const names = ['söndag','måndag','tisdag','onsdag','torsdag','fredag','lördag'];
  return names[date.getDay()];
}
function localYMD(date) {
  if (!date) return '';
  const y = date.getFullYear();
  const m = pad(date.getMonth()+1);
  const d = pad(date.getDate());
  return `${y}-${m}-${d}`;
}
function dayName(date) {
  const names = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  return names[date.getDay()];
}
function makeRecurrence(mode, interval, count, scheduleResult) {
  try {
    if (!mode || mode === 'none') return undefined;
    const startLocal = graphDateToLocalDate(scheduleResult?.start);
    if (!startLocal) return undefined;
    const recurTz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    const startDate = localYMD(startLocal);
    const safeInterval = Math.max(1, Number(interval) || 1);
    const safeCount = Math.max(1, Number(count) || 1);
    const monday = 'Monday';

    if (mode === 'weekly' || mode === 'biweekly') {
      const useInterval = mode === 'biweekly' ? 2 : safeInterval;
      return {
        pattern: {
          type: 'weekly',
          interval: useInterval,
          daysOfWeek: [dayName(startLocal)],
          firstDayOfWeek: monday,
        },
        range: {
          type: 'numbered',
          startDate,
          numberOfOccurrences: safeCount,
          recurrenceTimeZone: recurTz,
        },
      };
    }
    if (mode === 'monthly' || mode === 'quarterly') {
      const useInterval = mode === 'quarterly' ? 3 : safeInterval;
      return {
        pattern: {
          type: 'absoluteMonthly',
          interval: useInterval,
          dayOfMonth: startLocal.getDate(),
        },
        range: {
          type: 'numbered',
          startDate,
          numberOfOccurrences: safeCount,
          recurrenceTimeZone: recurTz,
        },
      };
    }
    return undefined;
  } catch {
    return undefined;
  }
}

function recurrenceSummarySv(mode, interval, count, scheduleResult) {
  if (!mode || mode==='none') return 'Engångsmöte';
  const startLocal = graphDateToLocalDate(scheduleResult?.start) || new Date();
  const weekday = dayNameSv(startLocal);
  const dom = startLocal.getDate();
  const safeInterval = Math.max(1, Number(interval) || 1);
  const safeCount = Math.max(1, Number(count) || 1);
  if (mode==='weekly') return `Veckovis på ${weekday}, ${safeCount} gånger (var ${safeInterval} v)`;
  if (mode==='biweekly') return `Varannan vecka på ${weekday}, ${safeCount} gånger`;
  if (mode==='monthly') return `Månadsvis den ${dom}:e, ${safeCount} gånger (var ${safeInterval} mån)`;
  if (mode==='quarterly') return `Kvartalsvis den ${dom}:e, ${safeCount} gånger`;
  return '';
}
