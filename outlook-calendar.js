// ============================================================
// DELETED ITEMS — RESTORE & PERMANENT DELETE
// ============================================================
async function restoreEmail(emailId){
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/move`,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({destinationId:'inbox'})
    });
    if(!res.ok)throw new Error('Restore failed');
    if(outlookFolderEmails['deleted']){
      outlookFolderEmails['deleted']=outlookFolderEmails['deleted'].filter(e=>e.id!==emailId);
    }
    delete emailCache[emailId];
    const readerEl=document.getElementById('ol-email-reader');
    if(readerEl)readerEl.innerHTML=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#9ca3af;text-align:center;padding:40px"><div style="font-size:13px;font-weight:600;color:#15803d">Email restored to Inbox</div></div>`;
    renderEmailRows('deleted',outlookFolderEmails['deleted']||[]);
    toast('Email restored to Inbox');
  }catch(err){toast('Could not restore email');}
}

async function permanentDeleteEmail(emailId){
  if(!confirm('Permanently delete this email? This cannot be undone.'))return;
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}`,{method:'DELETE',headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok&&res.status!==204)throw new Error('Delete failed');
    if(outlookFolderEmails['deleted']){
      outlookFolderEmails['deleted']=outlookFolderEmails['deleted'].filter(e=>e.id!==emailId);
    }
    delete emailCache[emailId];
    const readerEl=document.getElementById('ol-email-reader');
    if(readerEl)readerEl.innerHTML=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#9ca3af;text-align:center;padding:40px"><div style="font-size:13px;font-weight:600;color:#6b7280">Email permanently deleted</div></div>`;
    renderEmailRows('deleted',outlookFolderEmails['deleted']||[]);
    toast('Email permanently deleted');
  }catch(err){toast('Could not delete email');}
}

// ============================================================
// CALENDAR INTEGRATION
// ============================================================
let calendarEvents=[];

async function loadScheduleFolder(){
  document.querySelectorAll('.ol-folder').forEach(f=>f.classList.remove('active'));
  const btn=document.getElementById('ol-folder-schedule');
  if(btn)btn.classList.add('active');
  currentFolder='schedule';
  if(!olMidExpanded){olMidExpanded=true;const mp=document.getElementById('ol-mid-panel');if(mp)mp.style.width='300px';updateOlDividerIcons();}
  const titleEl=document.getElementById('ol-folder-title');
  if(titleEl)titleEl.textContent='Schedule Meeting';
  const listEl=document.getElementById('ol-email-list');
  if(listEl)listEl.innerHTML='<div style="padding:20px;text-align:center;font-size:12px;color:#9ca3af">Loading meetings...</div>';
  const readerEl=document.getElementById('ol-email-reader');
  if(readerEl)readerEl.innerHTML=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--muted);text-align:center;padding:40px"><svg width="48" height="48" fill="none" stroke="#d1d5db" viewBox="0 0 24 24" style="margin-bottom:14px"><path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" stroke-width="1.5" stroke-linecap="round"/></svg><div style="font-size:14px;font-weight:600;color:var(--body);margin-bottom:4px">Schedule a Meeting</div><div style="font-size:12px;color:var(--muted)">Click an event to view details, or + New Meeting to create one</div></div>`;
  const newMeetBtn=()=>{const d=document.createElement('div');d.style.cssText='padding:10px 12px 8px;border-bottom:1px solid #f0f0f0';d.innerHTML=`<button onclick="openNewMeetingFromEmail('')" style="display:flex;align-items:center;justify-content:center;gap:6px;width:100%;padding:10px;border-radius:6px;border:none;background:#0E3416;color:#fff;font-family:Inter,sans-serif;font-size:12px;font-weight:600;cursor:pointer" onmouseover="this.style.background='#1A5C2A'" onmouseout="this.style.background='#0E3416'"><svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>+ New Meeting</button>`;return d;};
  try{
    const token=await getAccessToken();
    const today=new Date();today.setHours(0,0,0,0);
    const end=new Date(today);end.setDate(end.getDate()+30);
    const url=`https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${today.toISOString()}&endDateTime=${end.toISOString()}&$select=id,subject,start,end,location,organizer,attendees,body,bodyPreview&$top=50&$orderby=start/dateTime`;
    const res=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error();
    const data=await res.json();
    calendarEvents=data.value||[];
    renderCalendarList(calendarEvents);
    if(listEl)listEl.insertBefore(newMeetBtn(),listEl.firstChild);
  }catch{
    if(listEl){listEl.innerHTML='';const b=newMeetBtn();listEl.appendChild(b);const e=document.createElement('div');e.style.cssText='padding:16px;text-align:center;font-size:12px;color:#b91c1c';e.textContent='Could not load meetings';listEl.appendChild(e);}
  }
}

async function loadCalendar(){
  // Update folder header and expand mid panel
  document.querySelectorAll('.ol-folder').forEach(f=>f.classList.remove('active'));
  const calBtn=document.getElementById('ol-folder-calendar');
  if(calBtn)calBtn.classList.add('active');
  currentFolder='calendar';
  if(!olMidExpanded){olMidExpanded=true;const mp=document.getElementById('ol-mid-panel');if(mp)mp.style.width='300px';updateOlDividerIcons();}
  const titleEl=document.getElementById('ol-folder-title');
  if(titleEl)titleEl.textContent='Calendar';
  const listEl=document.getElementById('ol-email-list');
  if(listEl)listEl.innerHTML='<div style="padding:20px;text-align:center;font-size:12px;color:#9ca3af">Loading events...</div>';
  try{
    const calToken=await getAccessToken();
    const today=new Date();today.setHours(0,0,0,0);
    const end=new Date(today);end.setDate(end.getDate()+30);
    const url=`https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${today.toISOString()}&endDateTime=${end.toISOString()}&$select=id,subject,start,end,location,organizer,attendees,body,bodyPreview&$top=50&$orderby=start/dateTime`;
    const res=await fetch(url,{headers:{Authorization:`Bearer ${calToken}`}});
    if(!res.ok)throw new Error('Calendar fetch failed');
    const data=await res.json();
    calendarEvents=data.value||[];
    renderCalendarList(calendarEvents);
  }catch(err){
    const listEl=document.getElementById('ol-email-list');
    if(listEl)listEl.innerHTML=`<div style="padding:20px;text-align:center;font-size:12px;color:#b91c1c">Could not load calendar events</div>`;
    toast('Could not load calendar');
  }
}

function renderCalendarList(events){
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  if(!events.length){listEl.innerHTML='<div class="empty-state"><div class="es-text">No upcoming events</div></div>';return;}
  // Group by date
  const groups={};
  events.forEach(ev=>{
    const dt=ev.start?.dateTime?graphDateToCentral(ev.start):new Date(ev.start?.date||'');
    const key=isNaN(dt)?'Unknown':centralDate(dt,{weekday:'long',month:'long',day:'numeric',year:'numeric'});
    if(!groups[key])groups[key]=[];
    groups[key].push(ev);
  });
  let html='';
  Object.entries(groups).forEach(([date,evs])=>{
    html+=`<div style="padding:5px 10px 2px;font-size:9px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.8px;background:#fafafa;border-bottom:1px solid #f0f0f0">${escapeHtml(date)}</div>`;
    evs.forEach(ev=>{
      const start=ev.start?.dateTime?graphDateToCentral(ev.start):null;
      const end=ev.end?.dateTime?graphDateToCentral(ev.end):null;
      const timeStr=start&&!isNaN(start)?centralTime(start,{hour:'2-digit',minute:'2-digit'})+(end&&!isNaN(end)?` — ${centralTime(end,{hour:'2-digit',minute:'2-digit'})}`:''): 'All day';
      const loc=ev.location?.displayName||'';
      const attCount=(ev.attendees||[]).length;
      html+=`<div class="ol-email-item" onclick="readCalendarEvent('${escapeHtml(ev.id)}')">
        <div class="ol-email-top"><span class="ol-email-sender">${escapeHtml(ev.subject||'(no title)')}</span></div>
        <div class="ol-email-subject">${escapeHtml(timeStr)}</div>
        ${loc?`<div class="ol-email-preview">${escapeHtml(loc)}</div>`:''}
        ${attCount?`<div style="font-size:10px;color:#9ca3af;margin-top:1px">${attCount} attendee${attCount!==1?'s':''}</div>`:''}
      </div>`;
    });
  });
  listEl.innerHTML=html;
}

function readCalendarEvent(eventId){
  const ev=calendarEvents.find(e=>e.id===eventId);
  if(!ev)return;
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  const start=ev.start?.dateTime?graphDateToCentral(ev.start):null;
  const end=ev.end?.dateTime?graphDateToCentral(ev.end):null;
  const fmt=(dt)=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric'});
  const fmtT=(dt)=>centralTime(dt,{hour:'2-digit',minute:'2-digit'});
  const dur=start&&end&&!isNaN(start)&&!isNaN(end)?Math.round((end-start)/60000):0;
  const durStr=dur?`(${Math.floor(dur/60)}h${dur%60?` ${dur%60}m`:''})`:'';
  const org=ev.organizer?.emailAddress?.name||ev.organizer?.emailAddress?.address||'';
  const attendees=ev.attendees||[];
  const joinUrl=ev.onlineMeeting?.joinUrl||ev.onlineMeetingUrl||extractMeetingJoinLink(ev.body?.content||ev.bodyPreview||'');
  const location=ev.location?.displayName||ev.location?.address?.street||'';
  const pBtn=(label,icon,onclick)=>`<button onclick="${onclick}" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #e5e7eb;background:#fff;color:#374151;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s;white-space:nowrap" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='#fff'">${icon?`<svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24">${icon}</svg>`:''}${label}</button>`;
  readerEl.innerHTML=`
    <div style="padding:12px 16px 10px;border-bottom:1px solid #e5e7eb;flex-shrink:0;background:#fff">
      <div style="font-size:16px;font-weight:600;color:#111;margin-bottom:8px">${escapeHtml(ev.subject||'(no title)')}</div>
      <div style="display:flex;gap:5px;flex-wrap:wrap;align-items:center;margin-bottom:10px">
        ${pBtn('Add to Tasks','<path d="M12 5v14M5 12h14" stroke-width="2" stroke-linecap="round"/>',`addCalendarEventToTasks('${escapeHtml(eventId)}')`)}
      </div>
    </div>
    <div style="flex:1;overflow-y:auto;padding:20px 16px">
      <div style="background:#f8fdf9;border:1px solid #d1e9d4;border-radius:6px;padding:14px">
        <div style="display:grid;grid-template-columns:auto 1fr;gap:6px 12px;font-size:12px">
          ${start?`<span style="color:#6b7280;font-weight:600">Date</span><span>${escapeHtml(fmt(start))}</span>`:''}
          ${start?`<span style="color:#6b7280;font-weight:600">Time</span><span>${fmtT(start)}${end?` — ${fmtT(end)}`:''} <span style="color:#9ca3af">${durStr}</span></span>`:''}
          <span style="color:#6b7280;font-weight:600">Room / Location</span><span>${location?escapeHtml(location):'<span style="color:#9ca3af">No location specified</span>'}</span>
          ${joinUrl?`<span style="color:#6b7280;font-weight:600">Teams Link</span><span><a href="${escapeHtml(joinUrl)}" target="_blank" rel="noopener" style="color:#0E3416;font-weight:700">Join Teams Meeting</a></span>`:''}
          ${org?`<span style="color:#6b7280;font-weight:600">Organiser</span><span>${escapeHtml(org)}</span>`:''}
          ${attendees.length?`<span style="color:#6b7280;font-weight:600">Attendees</span><span style="line-height:1.8">${renderExpandablePeople(attendees,`cal-att-${eventId}`,4)}</span>`:''}
        </div>
        ${ev.bodyPreview?`<div style="margin-top:12px;padding-top:10px;border-top:1px solid #d1e9d4;font-size:12.5px;color:#374151;line-height:1.6">${escapeHtml(ev.bodyPreview.slice(0,700))}</div>`:'' }
      </div>
    </div>`;
}

async function addCalendarEventToTasks(eventId){
  const ev=calendarEvents.find(e=>e.id===eventId);
  if(!ev)return;
  const start=ev.start?.dateTime?graphDateToCentral(ev.start):null;
  const fmt=(dt)=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric'});
  const fmtT=(dt)=>centralTime(dt,{hour:'2-digit',minute:'2-digit'});
  const org=ev.organizer?.emailAddress?.name||ev.organizer?.emailAddress?.address||'';
  const orgEmail=ev.organizer?.emailAddress?.address||'';
  const summary=`Meeting: ${ev.subject||''} on ${fmt(start)} at ${fmtT(start)}.${ev.bodyPreview?' '+ev.bodyPreview.slice(0,300):''}`;
  const task={
    id:Date.now()+Math.random(),
    assignedAt:new Date().toISOString(),
    createdAt:new Date().toISOString(),
    title:ev.subject||'(no title)',
    date:start&&!isNaN(start)?start.toISOString().split('T')[0]:new Date().toISOString().split('T')[0],
    person:org,email:orgEmail,dept:'Unknown',
    summary,threadSummary:summary,status:'Pending',priority:'Normal',
    wednesday:false,followup:'',weekOffset:0,replyCount:1,
    lastReplyAt:new Date().toISOString(),calendarEventId:eventId
  };
  tasks.unshift(task);
  syncBadges();refreshAll();
  toast('Calendar event added to Action Log');
  await saveTasksToOneDrive();
}
