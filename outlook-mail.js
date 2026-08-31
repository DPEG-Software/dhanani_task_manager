// ============================================================
// OUTLOOK INTEGRATION — Full featured with CC/BCC/Untracked
// ============================================================

/*
  POWER AUTOMATE SETUP (optional, enhances notifications):

  Create a flow with these two steps only:

  Step 1 — Trigger: When a new email arrives (V3)
  - Folder: Inbox
  - Subject filter: DONE

  Step 2 — Action: Update file (OneDrive for Business)
  - File path: /DPEGTaskManager/[recipient]/notifications.json
  - Append a new notification JSON object to the file

  This flow runs for each recipient separately and
  gives real-time notifications without polling.
*/

let currentFolder = 'inbox';
let currentEmailId = null;
let emailCache = {};
let trackedEmailIds = new Set();
let outlookFolderEmails = {};
let outlookNextLinks = {};
let outlookFolderTotals = {};
let _loadingMoreEmails = {};
let olListFilter = localStorage.getItem('dpeg_outlook_list_filter') || 'all';
if(!['all','unread','flagged'].includes(olListFilter))olListFilter='all';
let deletedFolderID = null;
let pinnedEmailIds = new Set(JSON.parse(localStorage.getItem('dpeg_pinned_email_ids')||'[]'));

const FOLDER_MAP = {inbox:'inbox',sent:'sentitems',drafts:'drafts',archive:'archive',deleted:'deleteditems',deleteditems:'deleteditems'};
const FOLDER_LABELS = {inbox:'Inbox',sent:'Sent Items',drafts:'Drafts',archive:'Archive',flagged:'Flagged Items',untracked:'Untracked Emails',deleted:'Deleted Items',deleteditems:'Deleted Items',calendar:'Calendar'};
function outlookFolderCanBeUnread(folder){return !['sent','drafts','deleted','deleteditems'].includes(String(folder||''));}

function buildTrackedSet(){
  trackedEmailIds = new Set(tasks.flatMap(t=>[t.emailId,t.lastMessageId,t.conversationId]).filter(Boolean));
}

function extractEmailAddress(value){
  const raw=String(value||'').trim();
  const angle=raw.match(/<([^<>@\s]+@[^<>\s]+)>/);
  if(angle)return angle[1].trim().toLowerCase();
  const email=raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return email?email[0].trim().toLowerCase():raw.toLowerCase();
}

function getDpegRecipients(toStr,ccStr){
  const bccStr=document.getElementById('compose-bcc')?.value||'';
  const all=[...(toStr||'').split(','),...(ccStr||'').split(','),...(bccStr||'').split(',')]
    .map(extractEmailAddress).filter(e=>e.includes('@dhananipeg.com'));
  return [...new Set(all)];
}

function showRecipientPreview(){
  const to=collectChipEmails('compose-to-chips','compose-to-input');
  const cc=collectChipEmails('compose-cc-chips','compose-cc-input');
  const dpeg=getDpegRecipients(to,cc);
  const preview=document.getElementById('compose-recipients-preview');
  const list=document.getElementById('compose-dpeg-list');
  if(!preview||!list)return;
  if(dpeg.length){
    preview.style.display='block';
    list.innerHTML=dpeg.map(e=>{
      const p=findPersonByEmail(e);
      return `<span style="display:inline-flex;align-items:center;gap:4px;margin-right:6px;margin-bottom:3px;background:#fff;border:1px solid var(--sage);border-radius:4px;padding:2px 8px;font-size:11.5px;font-weight:600;color:var(--forest)"><span style="width:6px;height:6px;border-radius:50%;background:var(--fern);display:inline-block;flex-shrink:0"></span>${p?.name||e}</span>`;
    }).join('');
  } else {
    preview.style.display='none';
  }
}

// ============================================================
// COMPOSE AUTOCOMPLETE (Outlook-style people picker)
// ============================================================
function showComposeAC(inputId, acId){
  const input=document.getElementById(inputId);
  const ac=document.getElementById(acId);
  if(!input||!ac)return;
  const val=input.value;
  const lastSep=Math.max(val.lastIndexOf(','),val.lastIndexOf(';'));
  const token=(lastSep>=0?val.slice(lastSep+1):val).trim().toLowerCase();
  if(!token.length){ac.style.display='none';return;}
  const seen=new Set();
  const allContacts=[...Object.values(staffConfig||{}),...(outlookContacts||[])];
  const matches=allContacts.filter(p=>p?.email&&p?.name)
    .filter(p=>{
      const k=normEmail(p.email||'');
      if(!k||seen.has(k))return false;
      seen.add(k);
      return (p.name||'').toLowerCase().includes(token)||(k).includes(token)||(p.role||'').toLowerCase().includes(token);
    }).slice(0,8);
  if(!matches.length){ac.style.display='none';return;}
  ac.innerHTML=matches.map(p=>`
    <div class="compose-ac-item"
      onmousedown="event.preventDefault();selectComposeAC('${inputId}','${acId}','${(p.email||'').replace(/'/g,"\\'")}')"
      onmouseover="document.querySelectorAll('#${acId} .compose-ac-item').forEach(x=>x.classList.remove('ac-focused'));this.classList.add('ac-focused')">
      ${av(p.name||'?',30)}
      <div style="flex:1;min-width:0;overflow:hidden">
        <div class="compose-ac-name">${p.name||p.email}</div>
        <div class="compose-ac-email">${p.email||''}</div>
        ${p.role||p.dept?`<div class="compose-ac-role">${p.role||p.dept}</div>`:''}
      </div>
    </div>`).join('');
  const rect=input.getBoundingClientRect();
  ac.style.position='fixed';
  ac.style.left=rect.left+'px';
  ac.style.top=(rect.bottom+2)+'px';
  ac.style.width=rect.width+'px';
  ac.style.right='auto';
  ac.style.display='block';
}

// ── Chip input helpers ────────────────────────────────────────────────────────
const _CHIP_MAP={
  'compose-to-input':['compose-to-chips','compose-to'],
  'compose-cc-input':['compose-cc-chips','compose-cc'],
  'compose-bcc-input':['compose-bcc-chips','compose-bcc'],
  'meet-required-input':['meet-required-chips','meet-required'],
  'meet-optional-input':['meet-optional-chips','meet-optional'],
};

function addEmailChip(chipsId,syncId,email,label){
  email=(email||'').trim();
  if(!email)return;
  label=(label||email).trim();
  const chips=document.getElementById(chipsId);
  if(!chips)return;
  const dup=[...chips.querySelectorAll('.chip-pill')].some(c=>c.dataset.email===email.toLowerCase());
  if(dup)return;
  const pill=document.createElement('span');
  pill.className='chip-pill';
  pill.dataset.email=email.toLowerCase();
  pill.title=email;
  pill.innerHTML=`<span class="chip-pill-label">${escapeHtml(label)}</span><button class="chip-x" type="button" onclick="event.stopPropagation();removeEmailChip(this,'${escapeHtml(chipsId)}','${escapeHtml(syncId)}')">&#215;</button>`;
  chips.appendChild(pill);
  _syncChips(chipsId,syncId);
}

function removeEmailChip(btn,chipsId,syncId){
  btn.closest('.chip-pill')?.remove();
  _syncChips(chipsId,syncId);
}

function _syncChips(chipsId,syncId){
  const chips=document.getElementById(chipsId);
  const inp=document.getElementById(syncId);
  if(inp)inp.value=[...( chips?.querySelectorAll('.chip-pill')||[])].map(c=>c.dataset.email).join(', ');
  showRecipientPreview();
}

function collectChipEmails(chipsId,textInputId){
  const chips=[...(document.getElementById(chipsId)?.querySelectorAll('.chip-pill')||[])].map(c=>c.dataset.email);
  const typed=(document.getElementById(textInputId)?.value||'').trim();
  if(typed)chips.push(typed);
  return chips.filter(Boolean).join(', ');
}

function chipKeyDown(e,chipsId,syncId,textInputId,acId){
  const inp=document.getElementById(textInputId);
  if(!inp)return;
  const ac=document.getElementById(acId);
  const focused=ac?.querySelector('.ac-focused');
  if((e.key==='Enter'||e.key===','||e.key===';')&&!focused){
    e.preventDefault();
    const val=inp.value.replace(/[,;]$|^\s+|\s+$/g,'').trim();
    if(val){addEmailChip(chipsId,syncId,val);inp.value='';}
    if(ac)ac.style.display='none';
    return;
  }
  if(e.key==='Backspace'&&!inp.value){
    const chips=document.getElementById(chipsId);
    chips?.querySelector('.chip-pill:last-of-type')?.remove();
    _syncChips(chipsId,syncId);
    return;
  }
  acKeyNav(e,acId,textInputId);
}

function chipBlur(chipsId,syncId,textInputId){
  const inp=document.getElementById(textInputId);
  if(!inp)return;
  const val=inp.value.replace(/[,;]$|^\s+|\s+$/g,'').trim();
  if(val){addEmailChip(chipsId,syncId,val);inp.value='';}
  _syncChips(chipsId,syncId);
}

function clearChipField(chipsId,syncId,textInputId){
  document.getElementById(chipsId)?.querySelectorAll('.chip-pill').forEach(c=>c.remove());
  const ti=document.getElementById(textInputId);if(ti)ti.value='';
  const si=document.getElementById(syncId);if(si)si.value='';
}

function setChipField(chipsId,syncId,textInputId,emailsStr){
  clearChipField(chipsId,syncId,textInputId);
  (emailsStr||'').split(/[,;]+/).map(s=>s.trim()).filter(Boolean).forEach(e=>addEmailChip(chipsId,syncId,e));
}
// ─────────────────────────────────────────────────────────────────────────────

function selectComposeAC(inputId,acId,email){
  const ac=document.getElementById(acId);
  if(ac)ac.style.display='none';
  const map=_CHIP_MAP[inputId];
  if(map){
    const [chipsId,syncId]=map;
    addEmailChip(chipsId,syncId,email);
    const inp=document.getElementById(inputId);if(inp){inp.value='';inp.focus();}
  }else{
    const input=document.getElementById(inputId);
    if(input){const v=input.value;const ls=Math.max(v.lastIndexOf(','),v.lastIndexOf(';'));input.value=(ls>=0?v.slice(0,ls+1)+' ':'')+email+', ';input.focus();}
  }
  showRecipientPreview();
  if((inputId==='compose-to-input'||inputId==='compose-to')&&currentComposeType==='Forward'){showForwardRelatedTasks(email);checkRecipientReplied(email);}
}

function hideComposeAC(acId){
  setTimeout(()=>{const ac=document.getElementById(acId);if(ac)ac.style.display='none';},180);
}

function acKeyNav(e,acId,inputId){
  const ac=document.getElementById(acId);
  if(!ac||ac.style.display==='none')return;
  const items=ac.querySelectorAll('.compose-ac-item');
  if(!items.length)return;
  const focused=ac.querySelector('.ac-focused');
  if(e.key==='ArrowDown'){
    e.preventDefault();
    const next=focused?focused.nextElementSibling||items[0]:items[0];
    items.forEach(i=>i.classList.remove('ac-focused'));
    if(next)next.classList.add('ac-focused');
  }else if(e.key==='ArrowUp'){
    e.preventDefault();
    const prev=focused?focused.previousElementSibling||items[items.length-1]:items[items.length-1];
    items.forEach(i=>i.classList.remove('ac-focused'));
    if(prev)prev.classList.add('ac-focused');
  }else if((e.key==='Enter'||e.key==='Tab')&&focused){
    e.preventDefault();
    focused.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));
  }else if(e.key==='Escape'){
    ac.style.display='none';
  }
}

function toggleBcc(){
  const row=document.getElementById('bcc-row');
  if(!row)return;
  const showing=row.style.display!=='none'&&row.style.display!=='';
  row.style.display=showing?'none':'block';
  const btn=document.getElementById('bcc-toggle-btn');
  if(btn)btn.style.color=showing?'':'var(--ink)';
}

function emailTimestamp(e){
  const t=new Date(e?.receivedDateTime||e?.sentDateTime||0).getTime();
  return Number.isFinite(t)?t:0;
}

function isEmailPinned(email){
  if(!email)return false;
  return pinnedEmailIds.has(email.id)||pinnedEmailIds.has(email.conversationId);
}

function savePinnedEmails(){
  localStorage.setItem('dpeg_pinned_email_ids',JSON.stringify([...pinnedEmailIds]));
}

function outlookHoverActionsHTML(email){
  const flagOn=outlookFlagStatus(email)==='flagged';
  const pinOn=isEmailPinned(email);
  return `<span class="ol-email-actions">
    <button class="ol-email-action-btn${flagOn?' on':''}" title="${flagOn?'Remove flag':'Flag'}" onclick="event.stopPropagation();toggleOutlookFlag(this.closest('.ol-email-item').dataset.emailId,event)"><svg width="11" height="11" viewBox="0 0 24 24" fill="${flagOn?'currentColor':'none'}" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 22V4"/><path d="M4 4h12l-1.5 5L16 14H4"/></svg></button>
    <button class="ol-email-action-btn${pinOn?' on':''}" title="${pinOn?'Unpin':'Pin'}" onclick="event.stopPropagation();toggleEmailPin(this.closest('.ol-email-item').dataset.emailId,event)"><svg width="11" height="11" viewBox="0 0 24 24" fill="${pinOn?'currentColor':'none'}" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5"/><path d="M8 3h8"/><path d="M9 3l1 8-3 4h10l-3-4 1-8"/></svg></button>
  </span>`;
}

function listAISummaryButtonHTML(folder){
  const safeFolder=escapeHtml(folder||'inbox');
  return `<button class="ol-sum-btn" title="AI Summary" onclick="event.stopPropagation();toggleListAISummary(this.closest('.ol-email-item').dataset.emailId,this,'${safeFolder}')" style="width:20px;height:20px;display:inline-flex;align-items:center;justify-content:center;background:none;border:1px solid #e9d5ff;border-radius:4px;cursor:pointer;color:#7c3aed;padding:0;flex-shrink:0;transition:all .12s" onmouseover="this.style.background='#f5f3ff'" onmouseout="this.style.background='none'"><svg width="10" height="10" viewBox="0 0 24 24" fill="currentColor"><path d="M12 1l2.7 8.3H23l-7 5.1 2.7 8.3L12 18l-7.7 4.7 2.7-8.3-7-5.1h8.3z"/></svg></button>`;
}

function emailDateGroupLabel(dt){
  const d=new Date(dt);
  if(isNaN(d))return 'Older';
  const today=new Date();
  const yesterday=new Date();
  yesterday.setDate(today.getDate()-1);
  if(d.toDateString()===today.toDateString())return 'Today';
  if(d.toDateString()===yesterday.toDateString())return 'Yesterday';
  return d.toLocaleDateString('en-US',{weekday:'long',month:'long',day:'numeric'});
}

function groupEmailsByConversation(emails){
  const groups=new Map();
  (emails||[]).forEach(e=>{
    const key=e.conversationId||e.id;
    if(!groups.has(key))groups.set(key,{key,messages:[],latest:e});
    const g=groups.get(key);
    g.messages.push(e);
    if(emailTimestamp(e)>emailTimestamp(g.latest))g.latest=e;
  });
  return Array.from(groups.values()).sort((a,b)=>{
    const ap=a.messages.some(isEmailPinned)?0:1;
    const bp=b.messages.some(isEmailPinned)?0:1;
    if(ap!==bp)return ap-bp;
    return emailTimestamp(b.latest)-emailTimestamp(a.latest);
  });
}

function applyOlListFilter(emails,folder=currentFolder){
  const list=emails||[];
  if(olListFilter==='unread')return list.filter(e=>!e.isRead);
  if(olListFilter==='flagged')return list.filter(e=>outlookFlagStatus(e)==='flagged');
  return list;
}

function updateOlFilterButtons(){
  ['all','unread','flagged'].forEach(mode=>{
    document.getElementById(`ol-filter-${mode}`)?.classList.toggle('active',olListFilter===mode);
  });
  document.getElementById('ol-filter-icon')?.classList.toggle('active',olListFilter!=='all');
}

function setOutlookEmailFilterVisible(visible){
  const menu=document.getElementById('ol-list-filter');
  if(menu)menu.style.display=visible?'':'none';
  if(!visible)document.getElementById('ol-filter-pop')?.classList.remove('open');
}

function toggleOlFilterMenu(event){
  event?.stopPropagation();
  const pop=document.getElementById('ol-filter-pop');
  if(pop)pop.classList.toggle('open');
}

function setOlListFilter(mode){
  olListFilter=['all','unread','flagged'].includes(mode)?mode:'all';
  localStorage.setItem('dpeg_outlook_list_filter',olListFilter);
  document.getElementById('ol-filter-pop')?.classList.remove('open');
  updateOlFilterButtons();
  if(currentFolder==='untracked'){loadUntracked();return;}
  renderEmailRows(currentFolder,outlookFolderEmails[currentFolder]||[]);
}

document.addEventListener('click',e=>{
  const menu=document.getElementById('ol-list-filter');
  if(menu&&!menu.contains(e.target))document.getElementById('ol-filter-pop')?.classList.remove('open');
});

function updateFlaggedCount(){
  const all=Object.values(outlookFolderEmails||{}).flat();
  const ids=new Set();
  all.forEach(e=>{if(e?.id&&outlookFlagStatus(e)==='flagged')ids.add(e.id);});
  const el=document.getElementById('ol-flagged-count');
  if(el){
    el.textContent=ids.size?String(ids.size):'';
    el.style.display=ids.size?'':'none';
  }
}

async function readConversationOrEmail(emailId,folder,el){
  if(_activeEmailEl){_activeEmailEl.classList.remove('active');}
  if(el){el.classList.add('active');_activeEmailEl=el;}
  currentEmailId=emailId;
  window._activeOutlookConversationMessageIds=typeof parseOutlookMessageIds==='function'
    ?parseOutlookMessageIds(el?.dataset?.messageIds,emailId)
    :[emailId];
  if(!emailCache[emailId]){
    const folderEmail=(outlookFolderEmails[folder]||[]).find(e=>e.id===emailId);
    if(folderEmail)emailCache[emailId]=folderEmail;
  }
  await viewFullThread(emailId);
}

function renderEmailRows(folder,emails){
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  updateOlFilterButtons();
  updateFlaggedCount();
  const loadedEmails=emails||[];
  if(!loadedEmails.length){listEl.innerHTML='<div class="empty-state"><div class="es-text">No emails</div></div>';return;}
  const filteredEmails=applyOlListFilter(loadedEmails,folder);
  if(!filteredEmails.length){
    const labels={unread:'No unread emails',flagged:'No flagged emails',all:'No emails'};
    const hint=outlookNextLinks[folder]?'Load more to check older messages':'Try a different filter';
    const more=outlookNextLinks[folder]?`<button class="btn btn-ghost btn-sm" style="margin-top:10px;justify-content:center;font-size:11px" onclick="loadMoreEmails('${folder}')">Load more</button>`:'';
    listEl.innerHTML=`<div class="empty-state"><div class="es-text">${labels[olListFilter]||'No emails'}</div><div class="es-sub">${hint}</div>${more}</div>`;
    return;
  }
  const groups=groupEmailsByConversation(filteredEmails);
  let lastGroup='';
  const rows=groups.map(g=>{
    const e=g.latest;
    const threadCount=g.messages.length;
    const isTracked=g.messages.some(m=>trackedEmailIds.has(m.id)||trackedEmailIds.has(m.conversationId));
    const isUnread=outlookFolderCanBeUnread(folder)&&g.messages.some(m=>!m.isRead);
    const isHigh=String(e.importance||'normal').toLowerCase()==='high';
    const isFlagged=g.messages.some(m=>outlookFlagStatus(m)==='flagged');
    const isPinned=g.messages.some(isEmailPinned);
    const hasAtt=g.messages.some(m=>!!m.hasAttachments);
    const safeId=escapeHtml(e.id||'');
    const safeFolder=escapeHtml(folder||'inbox');
    const safeMessageIds=escapeHtml(encodeURIComponent(JSON.stringify(g.messages.map(message=>message.id).filter(Boolean))));
    const canDrag=!['flagged','untracked','search'].includes(String(folder||''));
    // Sender
    const sender=folder==='sent'
      ?(e.toRecipients?.[0]?.emailAddress?.name||e.toRecipients?.[0]?.emailAddress?.address||'Unknown')
      :(e.from?.emailAddress?.name||e.from?.emailAddress?.address||'Unknown');
    // Date
    const dt=new Date(e.receivedDateTime||e.sentDateTime);
    const now=new Date();
    const isToday=dt.toDateString()===now.toDateString();
    const dateStr=isToday?dt.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'}):dt.toLocaleDateString('en-US',{month:'short',day:'numeric'});
    const groupLabel=emailDateGroupLabel(dt);
    const groupHeader=groupLabel!==lastGroup?`<div class="ol-date-group">${groupLabel}</div>`:'';
    lastGroup=groupLabel;
    // Badges
    const trackedBadge=isTracked?`<span style="background:#f0fdf4;color:#15803d;font-size:9px;font-weight:700;padding:1px 6px 1px 5px;border-radius:10px;border:1px solid #bbf7d0;white-space:nowrap;flex-shrink:0">✓</span>`:'';
    const highBadge=isHigh?`<span style="background:#fef2f2;color:#b91c1c;font-size:9px;font-weight:700;padding:0 5px;border-radius:10px;border:1px solid #fecaca;white-space:nowrap;flex-shrink:0">!</span>`:'';
    const attIcon=hasAtt?`<svg width="10" height="10" fill="none" stroke="#9ca3af" viewBox="0 0 24 24" style="flex-shrink:0"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`:'';
    const sumBtn=listAISummaryButtonHTML(safeFolder);
    const hoverActions=outlookHoverActionsHTML(e);
    const threadBadge=threadCount>1?`<span class="ol-thread-count">${threadCount}</span>`:'';
    return `${groupHeader}<div class="ol-email-item${isUnread?' unread':''}${isFlagged?' flagged':''}${isPinned?' pinned':''}" id="email-item-${safeId}" data-email-id="${safeId}" data-email-folder="${safeFolder}" data-message-ids="${safeMessageIds}" draggable="${canDrag?'true':'false'}"${canDrag?` ondragstart="startOutlookEmailDrag(event,this.dataset.emailId,this.dataset.emailFolder,this.dataset.messageIds)" ondragend="endOutlookEmailDrag(event)"`:''} onclick="readConversationOrEmail(this.dataset.emailId,'${safeFolder}',this)" oncontextmenu="showEmailCtxMenu(event,'${safeId}','${safeFolder}')">
      <div class="ol-email-top">
        <span class="ol-email-sender">${escapeHtml(sender)}</span>
        <span style="display:flex;align-items:center;gap:3px;flex-shrink:0">${threadBadge}${trackedBadge}${highBadge}${attIcon}${sumBtn}${hoverActions}<span class="ol-email-date">${dateStr}</span></span>
      </div>
      <div class="ol-email-subject">${escapeHtml(e.subject||'(no subject)')}</div>
      <div class="ol-email-preview">${escapeHtml(e.bodyPreview||'')}</div>
      <div id="list-ai-sum-${safeId}" style="display:none;margin-top:5px;padding:6px 9px;background:#faf8ff;border-left:2px solid #7c3aed;border-radius:0 4px 4px 0;font-size:11px;color:#4b5563;line-height:1.5"></div>
    </div>`;
  }).join('');
  const outlookTotal=outlookFolderTotals[folder];
  const footer=outlookNextLinks[folder]
    ?`<div id="ol-list-loadmore" style="padding:10px;border-top:1px solid #f0f0f0;text-align:center;font-size:11px;color:#9ca3af">${_loadingMoreEmails[folder]?'Loading more…':(outlookTotal?`${loadedEmails.length} of ${outlookTotal} loaded — scroll for more`:'Scroll for more')}</div>`
    :`<div style="padding:8px 10px;border-top:1px solid #f0f0f0;font-size:10px;color:#9ca3af;text-align:center">${filteredEmails.length}${filteredEmails.length!==loadedEmails.length?` of ${loadedEmails.length}`:''} emails${groups.length!==filteredEmails.length?` · ${groups.length} conversations`:''}${outlookTotal?` (Outlook: ${outlookTotal})`:''}</div>`;
  listEl.innerHTML=rows+footer;
}

async function loadFolder(folder){
  currentFolder=folder;
  setOutlookEmailFilterVisible(true);
  if(folder==='flagged'&&olListFilter!=='all'){
    olListFilter='all';
    localStorage.setItem('dpeg_outlook_list_filter',olListFilter);
  }
  updateOlFilterButtons();
  olMobileBackToList();
  _activeEmailEl=null;
  buildTrackedSet();
  const titleEl=document.getElementById('ol-folder-title');
  if(titleEl)titleEl.textContent=typeof outlookFolderLabel==='function'?outlookFolderLabel(folder):(FOLDER_LABELS[folder]||folder);
  if(!olMidExpanded){olMidExpanded=true;const mp=document.getElementById('ol-mid-panel');if(mp)mp.style.width='300px';updateOlDividerIcons();}
  document.querySelectorAll('.ol-folder').forEach(f=>f.classList.remove('active'));
  const af=document.getElementById('ol-folder-'+folder)||document.querySelector(`.ol-folder[data-folder-key="${typeof CSS!=='undefined'&&CSS.escape?CSS.escape(folder):folder}"]`);
  if(af)af.classList.add('active');
  const readerEl=document.getElementById('ol-email-reader');
  if(readerEl)readerEl.innerHTML=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--muted);text-align:center;padding:40px"><svg width="48" height="48" fill="none" stroke="#d1d5db" viewBox="0 0 24 24" style="margin-bottom:14px"><path d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke-width="1.5" stroke-linecap="round"/></svg><div style="font-size:14px;font-weight:600;color:var(--body);margin-bottom:4px">Select an email</div><div style="font-size:12px;color:var(--muted)">Click any email to read it here</div></div>`;
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  listEl.innerHTML=`<div class="empty-state"><div class="es-text">Loading...</div></div>`;
  if(folder==='flagged'){await loadFlaggedFolder();return;}
  if(folder==='untracked'){await loadUntracked();return;}
  try{
    const token=await getAccessToken();
    const orderField=folder==='sent'?'sentDateTime':'receivedDateTime';
    const graphFolderId=typeof resolveOutlookFolderId==='function'?resolveOutlookFolderId(folder):FOLDER_MAP[folder];
    if(!graphFolderId)throw new Error('Folder is unavailable');
    const [res,totalRes]=await Promise.all([
      fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(graphFolderId)}/messages?$top=100&$select=id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview,isRead,conversationId,importance,hasAttachments,flag&$orderby=${orderField} desc`,{headers:{Authorization:`Bearer ${token}`}}),
      fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(graphFolderId)}?$select=totalItemCount,unreadItemCount`,{headers:{Authorization:`Bearer ${token}`}}).catch(()=>null)
    ]);
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const emails=data.value||[];
    outlookFolderEmails[folder]=emails;
    outlookNextLinks[folder]=data['@odata.nextLink']||'';
    let realUnread=null;
    if(totalRes&&totalRes.ok){
      const totalData=await totalRes.json();
      outlookFolderTotals[folder]=totalData.totalItemCount;
      realUnread=totalData.unreadItemCount;
    }else{
      delete outlookFolderTotals[folder];
    }
    if(folder==='inbox'){
      const unread=realUnread!=null?realUnread:emails.filter(e=>!e.isRead).length;
      const ce=document.getElementById('ol-inbox-count');
      if(ce){ce.textContent=unread>0?unread:'';ce.style.display=unread>0?'':'none';}
      const nb=document.getElementById('nb-inbox');
      if(nb)nb.textContent=unread>0?unread:'0';
    }
    renderEmailRows(folder,emails);
    return;
    if(!emails.length){listEl.innerHTML='<div class="empty-state"><div class="es-text">No emails found</div></div>';return;}
    listEl.innerHTML=emails.map(e=>{
      const isTracked=trackedEmailIds.has(e.id)||trackedEmailIds.has(e.conversationId);
      const sender=folder==='sent'?(e.toRecipients?.[0]?.emailAddress?.name||e.toRecipients?.[0]?.emailAddress?.address||'Unknown'):(e.from?.emailAddress?.name||e.from?.emailAddress?.address||'Unknown');
      const dt=new Date(e.receivedDateTime||e.sentDateTime);
      const now=new Date();
      const isToday=dt.toDateString()===now.toDateString();
      const dateStr=isToday?dt.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'}):dt.toLocaleDateString('en-US',{month:'short',day:'numeric'});
      const isUnread=!e.isRead&&outlookFolderCanBeUnread(folder);
      return `<div class="ol-email-item${isUnread?' unread':''}" id="email-item-${e.id}" data-email-id="${e.id}" onclick="readEmail('${e.id}','${folder}',this)">
        <div class="ol-email-top">
          <div class="ol-email-sender">${sender}</div>
          <div style="display:flex;align-items:center;gap:4px">
            ${isTracked?`<span style="background:var(--leaf-bg);color:var(--leaf);font-size:9px;font-weight:700;padding:1px 5px;border-radius:3px;border:1px solid var(--leaf-bd)">✓ Tracked</span>`:''}
            <div class="ol-email-date">${dateStr}</div>
          </div>
        </div>
        <div class="ol-email-subject">${e.subject||'(no subject)'}</div>
        <div class="ol-email-preview">${e.bodyPreview||''}</div>
      </div>`;
    }).join('');
  }catch(err){
    console.error('Load folder error:',err);
    listEl.innerHTML='<div class="empty-state"><div class="es-text">Could not load emails</div><div class="es-sub">Check your connection</div></div>';
  }
}

async function loadMoreEmails(folder){
  const next=outlookNextLinks[folder];
  if(!next||_loadingMoreEmails[folder])return;
  _loadingMoreEmails[folder]=true;
  const listEl=document.getElementById('ol-email-list');
  const loadingEl=document.getElementById('ol-list-loadmore');
  if(loadingEl)loadingEl.textContent='Loading more…';
  const scrollPos=listEl?.scrollTop||0;
  try{
    const token=await getAccessToken();
    const res=await fetch(next,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const more=data.value||[];
    outlookFolderEmails[folder]=[...(outlookFolderEmails[folder]||[]),...more];
    if(folder==='flagged')outlookFolderEmails[folder].sort((a,b)=>emailTimestamp(b)-emailTimestamp(a));
    outlookNextLinks[folder]=data['@odata.nextLink']||'';
    _loadingMoreEmails[folder]=false;
    renderEmailRows(folder,outlookFolderEmails[folder]);
    if(listEl)listEl.scrollTop=scrollPos;
  }catch(err){
    console.error('Load more email error:',err);
    _loadingMoreEmails[folder]=false;
    if(listEl)listEl.insertAdjacentHTML('beforeend','<div style="padding:10px 14px;font-size:11px;color:#b91c1c;text-align:center">Could not load more emails</div>');
  }
}

async function loadFlaggedFolder(){
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  try{
    const token=await getAccessToken();
    const select='id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview,isRead,conversationId,importance,hasAttachments,flag';
    const filter=encodeURIComponent("flag/flagStatus eq 'flagged'");
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=100&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const byId=new Map();
    Object.entries(outlookFolderEmails||{}).forEach(([folder,items])=>{
      if(folder==='flagged')return;
      (items||[]).forEach(e=>{if(e?.id&&outlookFlagStatus(e)==='flagged')byId.set(e.id,e);});
    });
    (data.value||[]).forEach(e=>{if(e?.id)byId.set(e.id,e);});
    outlookFolderEmails.flagged=[...byId.values()].sort((a,b)=>emailTimestamp(b)-emailTimestamp(a));
    outlookNextLinks.flagged=data['@odata.nextLink']||'';
    updateFlaggedCount();
    renderEmailRows('flagged',outlookFolderEmails.flagged);
  }catch(err){
    console.error('Load flagged emails error:',err);
    listEl.innerHTML='<div class="empty-state"><div class="es-text">Could not load flagged emails</div></div>';
  }
}

function isUntrackedDpegEmail(e){
  const isDpeg=e.toRecipients?.some(r=>r.emailAddress?.address?.toLowerCase().includes('@dhananipeg.com'));
  const isTracked=trackedEmailIds.has(e.id)||trackedEmailIds.has(e.conversationId);
  return isDpeg&&!isTracked;
}
function sortUntracked(list){
  return [...list].sort((a,b)=>{
    const ap=isEmailPinned(a)?0:1;
    const bp=isEmailPinned(b)?0:1;
    if(ap!==bp)return ap-bp;
    return emailTimestamp(b)-emailTimestamp(a);
  });
}
function renderUntrackedList(){
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  const untracked=outlookFolderEmails.untracked||[];
  const untrackedCount=untracked.length;
  const uc=document.getElementById('ol-untracked-count');
  if(uc){uc.textContent=untrackedCount>0?untrackedCount:'';uc.style.display=untrackedCount>0?'':'none';}
  if(!untracked.length){listEl.innerHTML='<div class="empty-state"><div class="es-text">All caught up</div><div class="es-sub">All emails to DPEG staff are tracked</div></div>';return;}
  const visibleUntracked=applyOlListFilter(untracked,'untracked');
  if(!visibleUntracked.length){
    const labels={unread:'No unread untracked emails',flagged:'No flagged untracked emails',all:'No untracked emails'};
    listEl.innerHTML=`<div class="empty-state"><div class="es-text">${labels[olListFilter]||'No untracked emails'}</div><div class="es-sub">Try a different filter</div></div>`;
    return;
  }
  listEl.innerHTML=`<div style="padding:10px 14px;background:var(--amber-bg);border-bottom:1px solid var(--amber-bd)"><div style="font-size:11.5px;font-weight:600;color:var(--amber)">${untracked.length} email${untracked.length!==1?'s':''} sent to DPEG staff not yet in Action Log</div></div>`+
  visibleUntracked.map(e=>{
    const to=e.toRecipients?.map(r=>r.emailAddress?.name||r.emailAddress?.address).join(', ')||'';
    const dt=new Date(e.sentDateTime);
    const dateStr=dt.toLocaleDateString('en-US',{month:'short',day:'numeric'});
    const isHigh=String(e.importance||'normal').toLowerCase()==='high';
    const isFlagged=outlookFlagStatus(e)==='flagged';
    const isPinned=isEmailPinned(e);
    const hasAtt=!!e.hasAttachments;
    const safeId=escapeHtml(e.id||'');
    const sumBtn=listAISummaryButtonHTML('sent');
    const hoverActions=outlookHoverActionsHTML(e);
    return `<div class="ol-email-item${isFlagged?' flagged':''}${isPinned?' pinned':''}" id="email-item-${safeId}" data-email-id="${safeId}" onclick="readEmail(this.dataset.emailId,'sent',this)" style="display:flex;align-items:flex-start;gap:8px">
      <div style="flex:1;min-width:0">
        <div class="ol-email-top"><div class="ol-email-sender">To: ${escapeHtml(to)}</div><span style="display:flex;align-items:center;gap:3px;flex-shrink:0">${isHigh?`<span style="background:#fef2f2;color:#b91c1c;font-size:9px;font-weight:700;padding:0 5px;border-radius:10px;border:1px solid #fecaca;white-space:nowrap;flex-shrink:0">!</span>`:''}${sumBtn}${hoverActions}<div class="ol-email-date">${dateStr}</div></span></div>
        <div class="ol-email-subject">${escapeHtml(e.subject||'(no subject)')}</div>
        <div class="ol-email-preview">${escapeHtml(e.bodyPreview||'')}</div>
        ${hasAtt?`<div class="ol-email-attachments"><svg width="12" height="12" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Attachment</span></div>`:''}
        <div id="list-ai-sum-${safeId}" style="display:none;margin-top:5px;padding:6px 9px;background:#faf8ff;border-left:2px solid #7c3aed;border-radius:0 4px 4px 0;font-size:11px;color:#4b5563;line-height:1.5"></div>
      </div>
      <button onclick="event.stopPropagation();quickAddTask(this.closest('.ol-email-item').dataset.emailId)" style="flex-shrink:0;margin-top:8px;padding:4px 8px;background:var(--ink);color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;white-space:nowrap">+ Add</button>
    </div>`;
  }).join('');
}
async function loadUntracked(){
  const listEl=document.getElementById('ol-email-list');
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=100&$select=id,subject,from,toRecipients,sentDateTime,bodyPreview,conversationId,importance,hasAttachments,flag&$orderby=sentDateTime desc`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    outlookNextLinks.untracked=data['@odata.nextLink']||'';
    outlookFolderEmails.untracked=sortUntracked((data.value||[]).filter(isUntrackedDpegEmail));
    renderUntrackedList();
  }catch(err){
    console.error('Load untracked error:',err);
    if(listEl)listEl.innerHTML='<div class="empty-state"><div class="es-text">Could not load untracked emails</div><div class="es-sub">Check your connection</div></div>';
  }
}
async function loadMoreUntracked(){
  const next=outlookNextLinks.untracked;
  if(!next||_loadingMoreEmails.untracked)return;
  _loadingMoreEmails.untracked=true;
  const listEl=document.getElementById('ol-email-list');
  const scrollPos=listEl?.scrollTop||0;
  try{
    const token=await getAccessToken();
    const res=await fetch(next,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const more=(data.value||[]).filter(isUntrackedDpegEmail);
    outlookFolderEmails.untracked=sortUntracked([...(outlookFolderEmails.untracked||[]),...more]);
    outlookNextLinks.untracked=data['@odata.nextLink']||'';
    _loadingMoreEmails.untracked=false;
    renderUntrackedList();
    if(listEl)listEl.scrollTop=scrollPos;
  }catch(err){
    console.error('Load more untracked error:',err);
    _loadingMoreEmails.untracked=false;
    toast('Could not load more emails — try again');
  }
}

async function quickAddTask(emailId){
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=id,subject,from,toRecipients,sentDateTime,bodyPreview,conversationId,importance,flag`,{headers:{Authorization:`Bearer ${token}`}});
    const email=await res.json();
    const recipient=email.toRecipients?.[0]?.emailAddress;
    const p=findPersonByEmail(recipient?.address);
    const person=p?.name||recipient?.name||recipient?.address||'Unknown';
    const address=recipient?.address||'';
    await attachFreeThreadSummary(email);
    const result=upsertTaskFromEmail(email,person,address,personDept(address,person),email.sentDateTime);
    trackedEmailIds.add(emailId);
    await saveTasksToOneDrive();
    refreshAll();
    toast(result.created?'Added to Action Log':'Updated existing thread in Action Log');
    loadUntracked();
  }catch(err){toast('Could not add task');}
}

function updateUnreadBadges(delta){
  ['ol-inbox-count','nb-inbox'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    const next=Math.max(0,(parseInt(el.textContent,10)||0)+delta);
    el.textContent=next>0?String(next):(id==='nb-inbox'?'0':'');
    if(id==='ol-inbox-count')el.style.display=next>0?'':'none';
  });
}

async function openTaskEmail(emailId){
  closeMo('mo-detail');
  nav('outlook');
  await loadFolder('inbox');
  await readEmail(emailId,'inbox');
}


let _pwaPrompt=null;
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();_pwaPrompt=e;const b=document.getElementById('pwa-install-btn');if(b)b.style.display='flex';});
window.addEventListener('appinstalled',()=>{_pwaPrompt=null;const b=document.getElementById('pwa-install-btn');if(b)b.style.display='none';});
function pwaInstall(){if(!_pwaPrompt)return;_pwaPrompt.prompt();_pwaPrompt.userChoice.then(()=>{_pwaPrompt=null;const b=document.getElementById('pwa-install-btn');if(b)b.style.display='none';});}

// Keeps an email iframe's height matched to its real content. A single
// measurement taken on 'load' is too early — images without cid: (so not
// inlined as data URIs) and web fonts still finish loading/reflowing after
// that, and the iframe was left frozen at the too-small pre-image height,
// clipping the rest of the email. ResizeObserver keeps it in sync as those
// finish, instead of guessing a single "safe" moment to measure.
const _iframeSizeObservers=new WeakMap();
function syncIframeHeight(fr){
  try{
    const doc=fr.contentDocument;
    if(!doc?.body)return;
    fr.style.height=Math.max(doc.body.scrollHeight,160)+'px';
    if('ResizeObserver' in window && !_iframeSizeObservers.has(fr)){
      const ro=new ResizeObserver(()=>{try{fr.style.height=Math.max(doc.body.scrollHeight,160)+'px';}catch(e){}});
      ro.observe(doc.body);
      _iframeSizeObservers.set(fr,ro);
    }
  }catch(e){}
}

function tvToggle(msgId){
  const body=document.getElementById(msgId+'-body');
  if(!body)return;
  const wasHidden=body.classList.contains('tv-body--collapsed');
  body.classList.toggle('tv-body--collapsed');
  if(wasHidden){body.querySelectorAll('iframe').forEach(fr=>syncIframeHeight(fr));}
}

let threadViewState={subject:'',messages:[],backEmailId:'',backFolder:'inbox',webLink:'',showBack:false,actionsHTML:''};

function stripQuotedPlainText(text){
  let out=String(text||'').replace(/\r\n/g,'\n');
  const patterns=[
    /\n\s*_{5,}\s*\n[\s\S]*$/i,
    /\n\s*-{2,}\s*Original Message\s*-{2,}[\s\S]*$/i,
    /\n\s*Begin forwarded message:[\s\S]*$/i,
    /\n\s*On .{0,180}wrote:\s*\n[\s\S]*$/i,
    /\n\s*From:\s.+\n\s*(Sent|Date):\s.+\n\s*To:\s.+[\s\S]*$/i,
    /\n\s*________________________________\s*\n\s*From:\s[\s\S]*$/i
  ];
  patterns.forEach(re=>{out=out.replace(re,'');});
  return out.trim()||String(text||'').trim();
}

function removeNodeAndFollowing(node){
  if(!node||!node.parentNode)return;
  let cur=node;
  while(cur){
    const next=cur.nextSibling;
    cur.parentNode.removeChild(cur);
    cur=next;
  }
}

// Remove a quoted-history boundary and everything after it, even when Outlook
// nests the boundary several levels deep inside Word-generated divs/tables.
// The old helper only removed siblings inside the immediate parent, leaving
// the rest of the quoted thread behind when the wrapper was nested.
function removeNodeAndFollowingThroughRoot(node,root){
  if(!node||!root)return;
  const immediateParent=node.parentNode;
  if(!immediateParent)return;
  let tail=node;
  while(tail){
    const next=tail.nextSibling;
    immediateParent.removeChild(tail);
    tail=next;
  }
  // At higher levels, remove only siblings after the branch containing the
  // boundary. Removing the branch itself would also delete the authored text
  // that appeared before the nested Outlook quote marker.
  let branch=immediateParent;
  while(branch&&branch!==root){
    const parent=branch.parentNode;
    if(!parent)break;
    let sibling=branch.nextSibling;
    while(sibling){
      const next=sibling.nextSibling;
      parent.removeChild(sibling);
      sibling=next;
    }
    branch=parent;
  }
}

function findOutlookQuoteBoundary(root){
  if(!root)return null;
  const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);
  let node;
  while((node=walker.nextNode())){
    const text=String(node.nodeValue||'').replace(/\u00a0/g,' ').trim();
    if(!text)continue;
    const startsHeader=/^(from|sent|to|cc|subject|date)\s*:/i.test(text);
    const replyLead=/^on .{5,220} wrote:\s*$/i.test(text);
    const originalLead=/^(begin forwarded message:|-{2,}\s*original message\s*-{2,}|_{5,})/i.test(text);
    if(replyLead||originalLead)return node;
    if(startsHeader){
      const containerText=String(node.parentNode?.parentNode?.textContent||node.parentNode?.textContent||'')
        .replace(/\u00a0/g,' ').slice(0,1200);
      const headerMatches=containerText.match(/\b(?:from|sent|date|to|cc|subject)\s*:/gi)||[];
      if(headerMatches.length>=3)return node;
    }
  }
  return null;
}

function stripQuotedHTML(html){
  const raw=String(html||'');
  if(!raw)return raw;
  try{
    const wrap=document.createElement('div');
    wrap.innerHTML=raw;
    const structural=wrap.querySelector([
      'blockquote','.gmail_quote','.x_gmail_quote','.yahoo_quoted','.moz-cite-prefix',
      '[id^="divRplyFwdMsg"]','[id*="divRplyFwdMsg"]','[class*="OutlookMessageHeader"]',
      '[id*="appendonsend"]','[class*="replyForwardMsg"]','[class*="ms-outlook-mobile-reference-message"]'
    ].join(','));
    if(structural)removeNodeAndFollowingThroughRoot(structural,wrap);
    wrap.querySelectorAll('hr').forEach(hr=>{
      const tail=(hr.nextSibling?.textContent||hr.parentNode?.textContent||'').slice(0,400);
      if(/From:\s|Sent:\s|To:\s|Subject:\s/i.test(tail))removeNodeAndFollowingThroughRoot(hr,wrap);
    });
    const textBoundary=findOutlookQuoteBoundary(wrap);
    if(textBoundary)removeNodeAndFollowingThroughRoot(textBoundary,wrap);
    const text=wrap.textContent||'';
    if(/^\s*$/.test(text))return raw;
    return wrap.innerHTML;
  }catch{
    return raw;
  }
}

function threadMessageBody(m,isHtml){
  const body=m.body?.content||m.bodyPreview||'';
  // A true forwarded message intentionally contains the forwarded payload.
  // Ordinary replies should show only the text authored for this message.
  if(/^\s*(fw|fwd)\s*:/i.test(String(m.subject||'')))return body;
  return isHtml?stripQuotedHTML(body):stripQuotedPlainText(body);
}

function threadHeaderHTML(){
  const s=threadViewState;
  const count=s.messages.length;
  const latest=[...s.messages].sort((a,b)=>new Date(b.receivedDateTime||b.sentDateTime)-new Date(a.receivedDateTime||a.sentDateTime))[0];
  const sender=latest?.from?.emailAddress;
  const toAll=latest?.toRecipients||[];
  const ccAll=latest?.ccRecipients||[];
  const dt=latest?new Date(latest.receivedDateTime||latest.sentDateTime):null;
  const dateStr=dt?centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'}):'';
  const toStr=toAll.length?renderPeopleChips(toAll,'thread-hdr-to',4):'—';
  const ccStr=ccAll.length?renderPeopleChips(ccAll,'thread-hdr-cc',4):'';
  return `<div class="tv-hdr">
    <button type="button" class="ol-mobile-back-btn" onclick="olMobileBackToList()">← Back to emails</button>
    <div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:6px">
      <div style="flex:1;min-width:0">
        <div class="tv-subject">${escapeHtml(s.subject||'(no subject)')}</div>
        <div class="tv-count">${count} message${count!==1?'s':''}</div>
      </div>
      ${s.showBack&&s.backEmailId?`<button class="btn btn-ghost btn-sm" style="flex-shrink:0" onclick="readEmail('${s.backEmailId}','${s.backFolder||'inbox'}')">← Back</button>`:''}
    </div>
    ${latest?`<div style="margin-bottom:7px">
      <div style="display:flex;align-items:baseline;gap:6px;margin-bottom:3px"><span style="font-size:11px;color:var(--muted);font-weight:600;width:32px;flex-shrink:0;padding-top:2px">From</span><span style="font-size:12px;color:var(--body)">${escapeHtml(sender?.name||'')}${sender?.address?' <span style="color:var(--muted);font-size:11px">&lt;'+escapeHtml(sender.address)+'&gt;</span>':''}</span></div>
      <div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:3px;position:relative"><span style="font-size:11px;color:var(--muted);font-weight:600;width:32px;flex-shrink:0;padding-top:4px">To</span><div style="display:flex;flex-wrap:wrap;flex:1;min-width:0">${toStr}</div></div>
      ${ccStr?`<div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:3px;position:relative"><span style="font-size:11px;color:var(--muted);font-weight:600;width:32px;flex-shrink:0;padding-top:4px">CC</span><div style="display:flex;flex-wrap:wrap;flex:1;min-width:0">${ccStr}</div></div>`:''}
      <div style="display:flex;align-items:baseline;gap:6px"><span style="font-size:11px;color:var(--muted);font-weight:600;width:32px;flex-shrink:0">Date</span><span style="font-size:11px;color:var(--muted)">${dateStr}</span></div>
    </div>`:''}
    ${s.actionsHTML?`<div style="display:flex;gap:5px;flex-wrap:wrap">${s.actionsHTML}</div>`:''}
  </div>`;
}

async function hydrateThreadFramesAndAttachments(token){
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  const useToken=token||await getAccessToken();
  readerEl.querySelectorAll('iframe[data-bidx]').forEach(function(fr){
    var c=window._emailFrameBodies[+fr.dataset.bidx];
    if(c){
      fr.addEventListener('load',function(){syncIframeHeight(fr);});
      resolveInlineImages(useToken,fr.dataset.msgid,c).then(function(resolved){fr.srcdoc=resolved;});
    }
  });
  loadThreadAttachmentChips(threadViewState.messages);
}

async function renderStoredThread(token){
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  readerEl.innerHTML=threadHeaderHTML()+
    `<div id="email-ai-summary-panel" style="display:none;padding:10px 16px;background:#f5f3ff;border-bottom:1px solid #ddd6fe;flex-shrink:0"><div style="font-size:10px;font-weight:700;color:#7c3aed;margin-bottom:5px;letter-spacing:.04em">✦ AI SUMMARY</div><div id="email-ai-summary-content" style="font-size:12.5px;line-height:1.6;color:#111"></div></div>`+
    `<div class="tv-list">${buildThreadHTML(threadViewState.messages)}</div>`;
  await hydrateThreadFramesAndAttachments(token);
}

function buildThreadHTML(messages){
  // Newest first — latest lands at the top, history collapses below
  const sorted=[...messages].sort((a,b)=>new Date(b.receivedDateTime||b.sentDateTime)-new Date(a.receivedDateTime||a.sentDateTime));
  window._emailFrameBodies=[];
  const avColors=['#0078d4','#107c41','#b91c1c','#7c3aed','#b45309','#0f766e','#c2185b','#455a64'];
  const avColor=name=>{const h=Array.from(name||'?').reduce((s,c)=>s+c.charCodeAt(0),0);return avColors[h%avColors.length];};
  const avInit=(name,addr)=>{if(name&&name.trim().length>1)return name.trim().split(/\s+/).map(w=>w[0]).slice(0,2).join('').toUpperCase();return((addr||name||'?')[0]||'?').toUpperCase();};
  const fmtDt=dt=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});
  return sorted.map((m,i)=>{
    if(m?.id)emailCache[m.id]=m;
    const from=m.from?.emailAddress;
    const fromName=from?.name||from?.address||'Unknown';
    const to=renderExpandablePeople(m.toRecipients||[],`thread-to-${m.id}`,2);
    const cc=(m.ccRecipients||[]).length?renderExpandablePeople(m.ccRecipients||[],`thread-cc-${m.id}`,2):'';
    const dt=new Date(m.receivedDateTime||m.sentDateTime);
    const isHtml=m.body?.contentType==='html';
    const isLatest=i===0;
    // Every reply gets the same extraction rule. Each Graph message already
    // exists as its own block, so quoted copies of older messages must not be
    // repeated inside the newer block.
    const displayBody=threadMessageBody(m,isHtml);
    const bIdx=isHtml?(window._emailFrameBodies.push(displayBody)-1):-1;
    const preview=escapeHtml((m.bodyPreview||'').slice(0,140).replace(/\s+/g,' '));
    const msgId=`tv-${m.id.replace(/[^a-zA-Z0-9]/g,'_')}`;
    return `<div class="tv-msg${isLatest?' tv-msg--latest':''}" id="${msgId}">
      <div class="tv-msg-hdr" onclick="tvToggle('${msgId}')">
        <div class="tv-avatar" style="background:${avColor(fromName)}">${avInit(fromName,from?.address)}</div>
        <div class="tv-msg-meta">
          <div class="tv-from">${escapeHtml(fromName)}${isLatest?'<span class="tv-badge-latest">Latest</span>':''}</div>
          <div class="tv-to">To: ${to||'—'}${cc?` · CC: ${cc}`:''}</div>
        </div>
        <div class="tv-date">${fmtDt(dt)}</div>
      </div>
      <div class="tv-msg-actions" onclick="event.stopPropagation()">
        <button onclick="replyEmail('${m.id}')">Reply</button>
        <button onclick="replyAllEmail('${m.id}')">Reply All</button>
        <button onclick="forwardEmail('${m.id}')">Forward</button>
      </div>
      ${!isLatest&&preview?`<div class="tv-preview">${preview}…</div>`:''}
      <div class="tv-body${isLatest?'':' tv-body--collapsed'}" id="${msgId}-body">
        ${m.hasAttachments?`<div class="tv-att" id="thread-att-${m.id}"></div>`:''}
        <div class="tv-body-inner">${isHtml?`<iframe data-bidx="${bIdx}" data-msgid="${m.id}" style="width:100%;border:none;display:block;min-height:160px" sandbox="allow-same-origin allow-popups allow-popups-to-escape-sandbox"></iframe>`:`<pre style="font-family:inherit;white-space:pre-wrap;font-size:13px;line-height:1.6;margin:0">${escapeHtml(displayBody)}</pre>`}</div>
      </div>
    </div>`;
  }).join('');
}

async function openTaskThread(taskId){
  const t=tasks.find(x=>x.id===taskId);
  if(!t){return;}
  if(!t.conversationId){
    await openTaskEmail(taskEmailId(t));
    return;
  }
  closeMo('mo-detail');
  nav('outlook');
  await loadFolder('inbox');
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  readerEl.innerHTML='<div class="empty-state"><div class="es-text">Loading thread...</div></div>';
  try{
    const token=await getAccessToken();
    const filter=encodeURIComponent(`conversationId eq '${t.conversationId.replace(/'/g,"''")}'`);
    const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,conversationId,webLink,hasAttachments,flag,importance';
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const messages=(data.value||[]).sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime)-new Date(b.receivedDateTime||b.sentDateTime));
    refreshTaskFromThreadMessages(messages,emailSubject(t)).catch(()=>{});
    threadViewState={subject:emailSubject(t),messages,backEmailId:taskEmailId(t)||'',backFolder:'inbox',webLink:'',showBack:true,actionsHTML:''};
    await renderStoredThread(token);
  }catch(err){
    toast('Could not open email thread');
  }
}

function openEmailInOutlook(emailId){
  const email=emailCache[emailId];
  if(email?.webLink)window.open(email.webLink,'_blank','noopener');
}

function attachmentChipHTML(emailId,a){
  const safeName=String(a.name||'file').replace(/'/g,"\\'");
  return `<button class="btn btn-ghost btn-sm" style="font-size:11.5px;display:flex;align-items:center;gap:5px;max-width:240px" onclick="downloadAttachment('${emailId}','${a.id}','${safeName}')" title="Download ${escapeHtml(a.name||'file')}">
    <svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
    <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(a.name||'file')}</span>
    <span style="color:var(--muted);font-weight:400;flex-shrink:0">(${formatFileSize(a.size)})</span>
  </button>`;
}

async function fetchVisibleAttachments(emailId){
  const token=await getAccessToken();
  const attRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments?$select=id,name,contentType,size,isInline`,{headers:{Authorization:`Bearer ${token}`}});
  if(!attRes.ok)return [];
  const attData=await attRes.json();
  return (attData.value||[]).filter(a=>!a.isInline);
}

async function loadThreadAttachmentChips(messages){
  for(const m of messages){
    if(!m.hasAttachments)continue;
    const host=document.getElementById(`thread-att-${m.id}`);
    if(!host)continue;
    try{
      const atts=await fetchVisibleAttachments(m.id);
      if(atts.length){
        host.innerHTML=`<div style="display:flex;gap:7px;flex-wrap:wrap">${atts.map(a=>attachmentChipHTML(m.id,a)).join('')}</div>`;
      }
    }catch{}
  }
}

async function viewFullThread(emailId){
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  readerEl.innerHTML='<div class="empty-state"><div class="es-text">Loading...</div></div>';
  try{
    const token=await getAccessToken();
    const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,conversationId,webLink,hasAttachments,flag,importance';
    // Ensure we have full email data (conversationId may not be set in folder cache)
    let email=emailCache[emailId];
    if(!email?.conversationId||!email?.body){
      const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=${select}`,{headers:{Authorization:`Bearer ${token}`}});
      if(r.ok){email=await r.json();emailCache[emailId]=email;}
    }
    if(!email)throw new Error('Email not found');
    let messages;
    if(email.conversationId){
      const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
      const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
      if(res.ok){const data=await res.json();messages=data.value||[];}
    }
    if(!messages||!messages.length)messages=[email];
    messages.forEach(m=>{if(m?.id)emailCache[m.id]=m;});
    refreshTaskFromThreadMessages(messages,email.subject||'(no subject)').catch(()=>{});
    // Mark unread messages as read
    if(outlookFolderCanBeUnread(currentFolder)){
      const unread=messages.filter(m=>!m.isRead);
      let markedCount=0;
      for(const m of unread){
        try{
          const r=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+m.id,{method:'PATCH',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:'{"isRead":true}'});
          if(r.ok){m.isRead=true;const fc=(outlookFolderEmails[currentFolder]||[]).find(e=>e.id===m.id);if(fc)fc.isRead=true;markedCount++;}
        }catch{}
      }
      if(markedCount){
        document.getElementById('email-item-'+emailId)?.classList.remove('unread');
        updateUnreadBadges(-markedCount);
      }
    }
    // Process meeting/Teams invite emails so they render as cards
    messages=await Promise.all(messages.map(m=>processMeetingMsg(m,token)));
    messages.forEach(m=>{if(m?.id)emailCache[m.id]=m;});
    // Latest = first after newest-first sort; use its id for reply actions
    const latestMsg=[...messages].sort((a,b)=>new Date(b.receivedDateTime||b.sentDateTime)-new Date(a.receivedDateTime||a.sentDateTime))[0];
    const latestId=latestMsg?.id||emailId;
    const isTracked=trackedEmailIds.has(emailId)||(email.conversationId&&trackedEmailIds.has(email.conversationId));
    const isDeleted=currentFolder==='deleted'||currentFolder==='deleteditems';
    const pBtn=(label,icon,oc,extra)=>'<button onclick="'+oc+'" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #e5e7eb;background:#fff;color:#374151;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s;white-space:nowrap'+(extra?';'+extra:'')+'" onmouseover="this.style.background=\'#f5f5f5\'" onmouseout="this.style.background=\'#fff\'">'+(icon?'<svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24">'+icon+'</svg>':'')+label+'</button>';
    const aBtns=[
      pBtn('Reply','<path d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','replyEmail(\''+latestId+'\')'),
      pBtn('Reply All','<path d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6M7 6h8a7 7 0 017 7" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','replyAllEmail(\''+latestId+'\')'),
      pBtn('Forward All','<path d="M21 10H11a8 8 0 00-8 8v2M21 10l-6 6m6-6l-6-6" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','forwardFullThread(\''+emailId+'\')'),
      pBtn('New Meeting','<path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','openNewMeetingFromEmail(\''+emailId+'\')'),
      pBtn('Mark Unread','<path d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','markEmailUnread(\''+latestId+'\')'),
      pBtn('Move','<path d="M3 7h7l2 2h9v10a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" stroke-width="1.8" stroke-linejoin="round"/><path d="M12 14h6m-2-2l2 2-2 2" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','openMoveEmailMenu(\''+latestId+'\')'),
      email.webLink?pBtn('Outlook','<path d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>','openEmailInOutlook(\''+latestId+'\')',''):'',
      '<button onclick="toggleEmailAISummary(\''+latestId+'\')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #e9d5ff;background:#fff;color:#7c3aed;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s" onmouseover="this.style.background=\'#f5f3ff\'" onmouseout="this.style.background=\'#fff\'"><svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor"><path d="M12 1l2.7 8.3H23l-7 5.1 2.7 8.3L12 18l-7.7 4.7 2.7-8.3-7-5.1h8.3z"/></svg>Summarize</button>',
      isTracked
        ?'<span style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #bbf7d0;background:#f0fdf4;color:#15803d;font-size:11px;font-weight:500">✓ In Tasks</span><button onclick="removeTaskByEmail(\''+emailId+'\')" style="display:inline-flex;align-items:center;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer" onmouseover="this.style.background=\'#fff1f2\'" onmouseout="this.style.background=\'#fff\'">Remove</button>'
        :'<button onclick="addTaskFromEmail(\''+emailId+'\')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:none;background:#0E3416;color:#fff;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer" onmouseover="this.style.background=\'#1A5C2A\'" onmouseout="this.style.background=\'#0E3416\'">+ Add to Tasks</button>',
      isDeleted
        ?'<button onclick="restoreEmail(\''+emailId+'\')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #86efac;background:#f0fdf4;color:#166534;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer">Restore</button><button onclick="permanentDeleteEmail(\''+emailId+'\')" style="display:inline-flex;align-items:center;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer" onmouseover="this.style.background=\'#fff1f2\'" onmouseout="this.style.background=\'#fff\'">Delete Permanently</button>'
        :'<button onclick="deleteEmailItem(\''+emailId+'\')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer" onmouseover="this.style.background=\'#fff1f2\'" onmouseout="this.style.background=\'#fff\'"><svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>Delete</button>'
    ];
    const actionsHTML=aBtns.filter(Boolean).join('');
    threadViewState={subject:email.subject||'(no subject)',messages,backEmailId:emailId,backFolder:currentFolder,webLink:email.webLink||'',showBack:false,actionsHTML};
    await renderStoredThread(token);
  }catch(err){
    console.error('Thread load error:',err);
    toast('Could not load email');
  }
}

function sanitizeEmailFrameHTML(html){
  if(!html)return html;
  const wrap=document.createElement('div');
  wrap.innerHTML=html;
  wrap.querySelectorAll('script,noscript,iframe,object,embed,form,input,button').forEach(el=>el.remove());
  wrap.querySelectorAll('*').forEach(el=>{
    Array.from(el.attributes).forEach(attr=>{
      const name=attr.name.toLowerCase();
      const value=attr.value||'';
      if(name.startsWith('on'))el.removeAttribute(attr.name);
      if((name==='href'||name==='src'||name==='xlink:href')&&/^\s*javascript:/i.test(value))el.removeAttribute(attr.name);
      if((name==='href'||name==='src'||name==='xlink:href')&&/^\s*(cid|file|ms-appx|mhtml):/i.test(value))el.removeAttribute(attr.name);
    });
  });
  return wrap.innerHTML;
}

async function resolveInlineImages(token,emailId,html){
  if(!html)return html;
  if(_resolvedBodyCache.has(emailId))return _resolvedBodyCache.get(emailId);
  // Upgrade HTTP image URLs to HTTPS to avoid mixed-content blocking on the HTTPS app
  html=html.replace(/(<img\b[^>]*?\bsrc=(["']))http:\/\//gi,'$1https://');
  if(html.toLowerCase().indexOf('cid:')===-1){const r=sanitizeEmailFrameHTML(html);_resolvedBodyCache.set(emailId,r);return r;}
  // Convert Outlook VML image elements to regular img tags so browsers render them
  html=html.replace(/<v:imagedata\b[^>]*\bsrc=(["'])(cid:[^"']+)\1[^>]*>/gi,function(m,q,src){
    return '<img src="'+src+'" style="max-width:100%;height:auto">';
  });
  try{
    var atts=[];
    var nextUrl='https://graph.microsoft.com/v1.0/me/messages/'+emailId+'/attachments?$select=id,contentType,contentId&$top=50';
    while(nextUrl){
      var res=await fetch(nextUrl,{headers:{Authorization:'Bearer '+token}});
      if(!res.ok)break;
      var page=await res.json();
      (page.value||[]).forEach(function(a){if(a.contentId)atts.push(a);});
      nextUrl=page['@odata.nextLink']||'';
    }
    for(var i=0;i<atts.length;i++){
      var att=atts[i];
      if(!att.contentBytes&&att.id){
        try{
          var r2=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+emailId+'/attachments/'+att.id,{headers:{Authorization:'Bearer '+token}});
          if(r2.ok){var d2=await r2.json();if(d2.contentBytes){att.contentBytes=d2.contentBytes;att.contentType=d2.contentType||att.contentType;}}
        }catch(e){}
      }
      if(!att.contentBytes)continue;
      var rawCid=String(att.contentId||'').replace(/[<>]/g,'').trim();
      if(!rawCid)continue;
      var shortCid=rawCid.indexOf('@')>-1?rawCid.split('@')[0]:'';
      var dataUrl='data:'+(att.contentType||'image/png')+';base64,'+att.contentBytes;
      [rawCid,shortCid].filter(Boolean).forEach(function(v){
        var esc=v.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
        html=html.replace(new RegExp('cid:<'+esc+'>','gi'),dataUrl);
        html=html.replace(new RegExp('cid:'+esc,'gi'),dataUrl);
        try{html=html.replace(new RegExp('cid:'+encodeURIComponent(v),'gi'),dataUrl);}catch(e){}
      });
    }
    const resolved=sanitizeEmailFrameHTML(html);
    _resolvedBodyCache.set(emailId,resolved);
    return resolved;
  }catch(e){
    const resolved=sanitizeEmailFrameHTML(html);
    _resolvedBodyCache.set(emailId,resolved);
    return resolved;
  }
}

var _activeEmailEl=null;
const _resolvedBodyCache=new Map(); // msgId → resolved HTML string
const CENTRAL_TZ=Intl.DateTimeFormat().resolvedOptions().timeZone||'UTC';
function tzOffsetMs(timeZone,date){
  const parts=new Intl.DateTimeFormat('en-US',{timeZone,hour12:false,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'}).formatToParts(date);
  const vals=Object.fromEntries(parts.filter(p=>p.type!=='literal').map(p=>[p.type,Number(p.value)]));
  const asUtc=Date.UTC(vals.year,vals.month-1,vals.day,vals.hour%24,vals.minute,vals.second);
  return asUtc-date.getTime();
}
function dateFromZoneParts(y,m,d,h=0,min=0,s=0,timeZone=CENTRAL_TZ){
  const utcGuess=new Date(Date.UTC(y,m-1,d,h,min,s));
  return new Date(utcGuess.getTime()-tzOffsetMs(timeZone,utcGuess));
}
function graphDateToCentral(obj){
  if(!obj?.dateTime)return null;
  const raw=String(obj.dateTime).replace(/\.\d+$/,'');
  const m=raw.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  const zone=String(obj.timeZone||'').toLowerCase();
  if(m&&(zone.includes('central')||zone.includes('chicago')))return dateFromZoneParts(+m[1],+m[2],+m[3],+m[4],+m[5],m[6]?+m[6]:0);
  const dt=new Date(obj.dateTime.endsWith('Z')?obj.dateTime:obj.dateTime+'Z');
  return isNaN(dt)?null:dt;
}
function centralDate(dt,opts){
  return (!dt||isNaN(dt))?'':dt.toLocaleDateString('en-US',{timeZone:CENTRAL_TZ,...opts});
}
function centralTime(dt,opts){
  return (!dt||isNaN(dt))?'':dt.toLocaleTimeString('en-US',{timeZone:CENTRAL_TZ,...opts});
}
async function processMeetingMsg(m,token){
  const bodyContent=m.body?.content||m.bodyPreview||'';
  const isHtml=m.body?.contentType==='html';
  const isEventMsg=String(m['@odata.type']||'').toLowerCase().includes('eventmessage');
  // Meeting detection must inspect only this message's authored content.
  // A Teams link in quoted history does not make a normal reply a meeting.
  const authoredBody=isHtml?stripQuotedHTML(bodyContent):stripQuotedPlainText(bodyContent);
  const rawBody=isHtml?(()=>{try{const d=document.createElement('div');d.innerHTML=authoredBody;return d.textContent||'';}catch{return authoredBody;}})():authoredBody;
  const icalText=extractICalText(rawBody);
  const teamsJoinUrl=extractMeetingJoinLink(authoredBody);
  if(!isEventMsg&&!icalText)return m;
  const toAll=m.toRecipients||[];
  const ccAll=m.ccRecipients||[];
  let meetingData=icalText?parseICalData(icalText):(isEventMsg?{
    summary:m.subject||'Teams meeting',location:'',description:m.bodyPreview||'',
    organizer:m.from?.emailAddress?.name||m.from?.emailAddress?.address||'',
    dtstart:null,dtend:null,status:'',
    attendees:[...toAll,...ccAll].map(r=>r.emailAddress?.name||r.emailAddress?.address||'').filter(Boolean),
    joinUrl:teamsJoinUrl
  }:null);
  if(isEventMsg&&m.id){
    if(!m.startDateTime){
      try{
        const r=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+m.id+'?$select=meetingRequestType,startDateTime,endDateTime,location,attendees',{headers:{Authorization:'Bearer '+token}});
        if(r.ok){const d=await r.json();Object.assign(m,d);if(m.id)emailCache[m.id]=m;}
      }catch{}
    }
    const gm=buildMeetingDataFromGraph(m);
    meetingData={...gm,joinUrl:gm.joinUrl||meetingData?.joinUrl||teamsJoinUrl,description:gm.description||meetingData?.description||''};
  }
  if(meetingData)return{...m,body:{contentType:'html',content:renderMeetingCard(meetingData,m.subject,isEventMsg?m.id:'')},bodyPreview:m.bodyPreview||meetingData.description||''};
  return m;
}

function extractICalText(str){
  const m=String(str||'').match(/BEGIN:VCALENDAR[\s\S]*?END:VCALENDAR/i);
  return m?m[0]:'';
}
function parseICalData(raw){
  const text=String(raw||'').replace(/\r\n[ \t]/g,'').replace(/\r\n/g,'\n').replace(/\n[ \t]/g,'');
  const lines=text.split('\n');
  const get=(key)=>{
    const line=lines.find(l=>new RegExp('^'+key+'(;[^:]*)?:','i').test(l));
    if(!line)return '';
    return line.replace(/^[^:]+:/,'').replace(/\\n/g,'\n').replace(/\\,/g,',').replace(/\\;/g,';').trim();
  };
  const dtParse=(raw,key)=>{
    if(!raw)return null;
    const v=raw.includes(':')?raw.split(':').pop():raw;
    const m=v.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?)?(Z)?/);
    if(!m)return null;
    if(m[7])return new Date(Date.UTC(+m[1],+m[2]-1,+m[3],m[4]?+m[4]:0,m[5]?+m[5]:0,m[6]?+m[6]:0));
    const line=lines.find(l=>new RegExp('^'+key+'(;[^:]*)?:','i').test(l))||'';
    const tzid=(line.match(/TZID=([^;:]+)/i)?.[1]||'').toLowerCase();
    if(tzid.includes('central')||tzid.includes('chicago'))return dateFromZoneParts(+m[1],+m[2],+m[3],m[4]?+m[4]:0,m[5]?+m[5]:0,m[6]?+m[6]:0);
    return new Date(+m[1],+m[2]-1,+m[3],m[4]?+m[4]:0,m[5]?+m[5]:0,m[6]?+m[6]:0);
  };
  const orgLine=lines.find(l=>/^ORGANIZER/i.test(l))||'';
  const cnM=orgLine.match(/CN=([^;:\r\n]+)/i);
  const mailM=orgLine.match(/mailto:([^\r\n\s;]+)/i);
  // Parse all ATTENDEE lines
  const attendees=lines.filter(l=>/^ATTENDEE/i.test(l)).map(l=>{
    const cn=l.match(/CN=([^;:\r\n]+)/i);
    const mail=l.match(/mailto:([^\r\n\s;]+)/i);
    return cn?cn[1].trim():(mail?mail[1].trim():'');
  }).filter(Boolean);
  const desc=get('DESCRIPTION');
  return{
    summary:get('SUMMARY'),location:get('LOCATION'),description:get('DESCRIPTION'),
    organizer:cnM?cnM[1].trim():(mailM?mailM[1].trim():''),
    dtstart:dtParse(get('DTSTART'),'DTSTART'),dtend:dtParse(get('DTEND'),'DTEND'),
    status:get('STATUS')||get('METHOD'),attendees,
    joinUrl:extractMeetingJoinLink(desc),
  };
}
function renderMeetingCard(d,emailSubjectHint,emailId){
  const fmt=(dt)=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric'});
  const fmtT=(dt)=>centralTime(dt,{hour:'2-digit',minute:'2-digit'});
  const dur=(d.dtstart&&d.dtend&&!isNaN(d.dtstart)&&!isNaN(d.dtend))?Math.round((d.dtend-d.dtstart)/60000):0;
  // Detect status from METHOD or subject hint
  const subj=String(emailSubjectHint||'').toLowerCase();
  const method=String(d.status||'').toUpperCase();
  let statusBadge='';
  if(method==='CANCEL'||subj.startsWith('canceled:')||subj.startsWith('cancelled:'))
    statusBadge=`<span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;background:#fff7ed;color:#c2410c;border:1px solid #fed7aa">Cancelled</span>`;
  else if(subj.startsWith('accepted:'))
    statusBadge=`<span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0">Accepted</span>`;
  else if(subj.startsWith('declined:'))
    statusBadge=`<span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;background:#fef2f2;color:#b91c1c;border:1px solid #fecaca">Declined</span>`;
  const rows=[];
  if(d.dtstart&&!isNaN(d.dtstart)){
    const timeStr=dur>0?`${fmtT(d.dtstart)} — ${fmtT(d.dtend)}`:'All day';
    rows.push(`<span style="color:#6b7280;font-weight:600">Date</span><span>${escapeHtml(fmt(d.dtstart))}</span>`);
    rows.push(`<span style="color:#6b7280;font-weight:600">Time</span><span>${timeStr}${dur?` <span style="color:#9ca3af">(${Math.floor(dur/60)}h${dur%60?` ${dur%60}m`:''})</span>`:''}</span>`);
  }
  rows.push(`<span style="color:#6b7280;font-weight:600">Room / Location</span><span>${d.location?escapeHtml(d.location):'<span style="color:#9ca3af">No location specified</span>'}</span>`);
  if(d.joinUrl)rows.push(`<span style="color:#6b7280;font-weight:600">Teams Link</span><span><a href="${escapeHtml(d.joinUrl)}" target="_blank" rel="noopener" style="color:#0E3416;font-weight:700">Join Teams Meeting</a></span>`);
  if(d.organizer)rows.push(`<span style="color:#6b7280;font-weight:600">Organiser</span><span>${escapeHtml(d.organizer)}</span>`);
  if(d.attendees.length)rows.push(`<span style="color:#6b7280;font-weight:600">Attendees</span><span style="line-height:1.6">${renderExpandablePeople(d.attendees,'meet-att-'+Math.random().toString(36).slice(2),4)}</span>`);
  return`<div style="padding:20px 16px"><div style="display:flex;align-items:center;gap:8px;margin-bottom:14px"><svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="#0E3416" stroke-width="2"><path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" stroke-linecap="round" stroke-linejoin="round"/></svg><span style="font-size:11px;font-weight:700;color:#0E3416;text-transform:uppercase;letter-spacing:.5px">Meeting Request</span>${statusBadge?` ${statusBadge}`:''}</div>${emailId&&!method?`<div data-rsvp-email="${emailId}" style="display:flex;gap:8px;margin-bottom:10px"><button onclick="rsvpMeeting('${emailId}','accept')" style="padding:5px 14px;background:#0E3416;color:#fff;border:none;border-radius:4px;font-size:12px;cursor:pointer;font-weight:600">Accept</button><button onclick="rsvpMeeting('${emailId}','tentative')" style="padding:5px 14px;background:#f3f4f6;color:#374151;border:1px solid #d1d5db;border-radius:4px;font-size:12px;cursor:pointer;font-weight:600">Tentative</button><button onclick="rsvpMeeting('${emailId}','decline')" style="padding:5px 14px;background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;border-radius:4px;font-size:12px;cursor:pointer;font-weight:600">Decline</button></div>`:''}<div style="background:#f8fdf9;border:1px solid #d1e9d4;border-radius:6px;padding:14px">${d.summary?`<div style="font-size:14px;font-weight:700;color:#111;margin-bottom:12px">${escapeHtml(d.summary)}</div>`:''}<div style="display:grid;grid-template-columns:auto 1fr;gap:5px 12px;font-size:12px">${rows.join('')}</div>${d.description&&d.description.trim()?`<div style="margin-top:12px;padding-top:10px;border-top:1px solid #d1e9d4;font-size:12.5px;color:#374151;line-height:1.6;white-space:pre-wrap">${escapeHtml(d.description.trim().slice(0,800))}</div>`:''}</div></div>`;
}
function buildMeetingDataFromGraph(email){
  const attendees=(email.attendees||[]).map(a=>a.emailAddress?.name||a.emailAddress?.address||'').filter(Boolean);
  const method=String(email.meetingRequestType||'').toLowerCase();
  let status='';
  if(method.includes('cancel'))status='CANCEL';
  else if(method.includes('accept'))status='ACCEPTED';
  else if(method.includes('decline'))status='DECLINED';
  return{
    summary:email.subject||'',
    location:(email.location?.displayName||email.location?.address?.street||''),
    description:email.bodyPreview||'',
    organizer:email.from?.emailAddress?.name||email.from?.emailAddress?.address||'',
    dtstart:graphDateToCentral(email.startDateTime),
    dtend:graphDateToCentral(email.endDateTime),
    status,
    attendees,
    joinUrl:email.onlineMeeting?.joinUrl||email.onlineMeetingUrl||extractMeetingJoinLink(email.body?.content||email.bodyPreview||''),
  };
}
function icalToReadableText(bodyText){
  const ical=extractICalText(bodyText);
  if(!ical)return null;
  const d=parseICalData(ical);
  const parts=[];
  if(d.summary)parts.push('Meeting: '+d.summary);
  if(d.organizer)parts.push('Organiser: '+d.organizer);
  const fmt=(dt)=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});
  if(d.dtstart)parts.push('Date: '+fmt(d.dtstart));
  if(d.dtend&&d.dtend!==d.dtstart)parts.push('Ends: '+fmt(d.dtend));
  if(d.location)parts.push('Location: '+d.location);
  if(d.attendees&&d.attendees.length)parts.push(`Attendees (${d.attendees.length}): `+d.attendees.slice(0,8).join(', '));
  if(d.description&&d.description.trim())parts.push('Notes: '+d.description.trim().slice(0,500));
  return parts.length?parts.join('\n'):null;
}

function safeDomId(s){
  return String(s||'x').replace(/[^a-zA-Z0-9_-]/g,'_');
}

function togglePeopleList(id){
  const el=document.getElementById(id);
  const btn=document.getElementById(id+'-btn');
  if(!el)return;
  const open=el.style.display!=='none';
  el.style.display=open?'none':'block';
  if(btn)btn.textContent=open?btn.dataset.closed:'Show less';
}

function renderExpandablePeople(items,id,limit=2){
  const people=(items||[]).map(x=>{
    if(typeof x==='string')return x;
    return x?.emailAddress?.name||x?.emailAddress?.address||x?.name||x?.address||'';
  }).filter(Boolean);
  if(!people.length)return '—';
  const shown=people.slice(0,limit).map(escapeHtml).join(', ');
  if(people.length<=limit)return shown;
  const hidden=people.slice(limit).map(p=>`<div style="padding:2px 0">${escapeHtml(p)}</div>`).join('');
  const safe=safeDomId(id);
  const closed=`+${people.length-limit} others`;
  return `${shown} <button id="${safe}-btn" data-closed="${closed}" onclick="event.stopPropagation();togglePeopleList('${safe}')" style="border:none;background:none;color:#0E3416;font-size:11.5px;font-weight:700;cursor:pointer;padding:0 2px">${closed}</button><div id="${safe}" style="display:none;margin-top:4px;padding:6px 8px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:5px;line-height:1.5">${hidden}</div>`;
}

function renderPeopleChips(items,id,limit=3){
  const avColors=['#0078d4','#107c41','#b91c1c','#7c3aed','#b45309','#0f766e','#c2185b','#455a64'];
  const chipColor=n=>{const h=Array.from(n||'?').reduce((s,c)=>s+c.charCodeAt(0),0);return avColors[h%avColors.length];};
  const chipInit=n=>(n||'?').trim().split(/\s+/).map(w=>w[0]).slice(0,2).join('').toUpperCase()||'?';
  const people=(items||[]).map(x=>{
    const name=x?.emailAddress?.name||x?.emailAddress?.address||x?.name||x?.address||'';
    const addr=x?.emailAddress?.address||x?.address||'';
    return {name,addr};
  }).filter(p=>p.name||p.addr);
  if(!people.length)return '—';
  const chipHTML=p=>{
    const label=p.name||p.addr;
    const col=chipColor(label);
    const init=chipInit(label);
    return `<span title="${escapeHtml(p.addr||p.name)}" style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px 2px 3px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:11px;font-size:11px;color:#374151;white-space:nowrap;margin:1px 3px 1px 0;line-height:1.4"><span style="display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;background:${col};color:#fff;font-size:7.5px;font-weight:700;flex-shrink:0">${init}</span>${escapeHtml(label)}</span>`;
  };
  const shown=people.slice(0,limit);
  const rest=people.slice(limit);
  const safe=safeDomId(id);
  if(!rest.length)return shown.map(chipHTML).join('');
  const hiddenHTML=rest.map(p=>`<div style="padding:2px 0">${chipHTML(p)}</div>`).join('');
  const closed=`+${rest.length}`;
  return `${shown.map(chipHTML).join('')}<button id="${safe}-btn" data-closed="${closed}" onclick="event.stopPropagation();togglePeopleList('${safe}')" style="display:inline-flex;align-items:center;padding:2px 7px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:11px;font-size:11px;color:#0E3416;font-weight:700;cursor:pointer;margin:1px 3px 1px 0">${closed}</button><div id="${safe}" style="display:none;position:absolute;left:40px;margin-top:2px;padding:8px 10px;background:#fff;border:1px solid #e5e7eb;border-radius:7px;box-shadow:0 4px 14px rgba(0,0,0,.12);z-index:200;min-width:180px">${hiddenHTML}</div>`;
}

function extractMeetingJoinLink(raw){
  const text=String(raw||'');
  if(!text)return '';
  try{
    const wrap=document.createElement('div');
    wrap.innerHTML=text;
    const anchors=[...wrap.querySelectorAll('a[href]')];
    const hit=anchors.find(a=>/teams\.microsoft\.com|meetup-join|join/i.test(a.href||a.textContent||''));
    if(hit?.href)return hit.href;
  }catch{}
  const m=text.match(/https?:\/\/[^\s"'<>]*(?:teams\.microsoft\.com|meetup-join|join\.microsoft\.com)[^\s"'<>]*/i);
  return m?m[0].replace(/&amp;/g,'&'):'';
}

async function readEmail(emailId,folder,el){
  currentEmailId=emailId;
  document.getElementById('page-outlook')?.classList.add('ol-mobile-reading');
  if(_activeEmailEl){_activeEmailEl.classList.remove('active');}
  var _ael=el||document.getElementById('email-item-'+emailId);
  if(_ael){_ael.classList.add('active');_activeEmailEl=_ael;}
  const readerEl=document.getElementById('ol-email-reader');
  if(!readerEl)return;
  readerEl.innerHTML='<div style="flex:1;display:flex;align-items:center;justify-content:center;color:#9ca3af;font-size:12px">Loading...</div>';
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,isRead,conversationId,webLink,importance,hasAttachments,flag`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Failed');
    const email=await res.json();
    emailCache[emailId]=email;
    const sender=email.from?.emailAddress;
    const toAll=email.toRecipients||[];
    const toStr=renderExpandablePeople(toAll,`email-to-${emailId}`,2);
    const ccAll=email.ccRecipients||[];
    const ccStr=ccAll.length?renderExpandablePeople(ccAll,`email-cc-${emailId}`,2):'';
    const dt=new Date(email.receivedDateTime||email.sentDateTime);
    const dateStr=dt.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});
    const bodyContent=email.body?.content||email.bodyPreview||'';
    const isHtml=email.body?.contentType==='html';
    // Detect meeting: @odata.type tells us it's an EventMessage; also scan body for iCal
    const isEventMsg=String(email['@odata.type']||'').toLowerCase().includes('eventmessage');
    const authoredBody=isHtml?stripQuotedHTML(bodyContent):stripQuotedPlainText(bodyContent);
    const rawBodyForIcal=isHtml?(()=>{const d=document.createElement('div');d.innerHTML=authoredBody;return d.textContent||'';})():authoredBody;
    const icalText=extractICalText(rawBodyForIcal);
    const teamsJoinUrl=extractMeetingJoinLink(authoredBody);
    const isMeeting=isEventMsg||!!icalText;
    let meetingData=icalText?parseICalData(icalText):(isEventMsg?{
      summary:email.subject||'Teams meeting',
      location:'',
      description:email.bodyPreview||'',
      organizer:sender?.name||sender?.address||'',
      dtstart:null,
      dtend:null,
      status:'',
      attendees:[...toAll,...ccAll].map(r=>r.emailAddress?.name||r.emailAddress?.address||'').filter(Boolean),
      joinUrl:teamsJoinUrl
    }:null);
    // For EventMessage, fetch the EventMessage-specific fields such as time/location/attendees.
    if(isEventMsg){
      try{
        const emRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=meetingRequestType,startDateTime,endDateTime,location,attendees`,{headers:{Authorization:`Bearer ${token}`}});
        if(emRes.ok){const emData=await emRes.json();Object.assign(email,emData);emailCache[emailId]=email;}
      }catch{}
      const graphMeeting=buildMeetingDataFromGraph(email);
      meetingData={...graphMeeting,joinUrl:graphMeeting.joinUrl||meetingData?.joinUrl||teamsJoinUrl,description:graphMeeting.description||meetingData?.description||''};
    }
    const isTracked=trackedEmailIds.has(emailId)||trackedEmailIds.has(email.conversationId);
    // Helper: small pill button
    const pBtn=(label,icon,onclick,extra='')=>`<button onclick="${onclick}" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #e5e7eb;background:#fff;color:#374151;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s;white-space:nowrap${extra?';'+extra:''}" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='#fff'">${icon?`<svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24">${icon}</svg>`:''}${label}</button>`;
    const actionsHTML=`
      ${pBtn('Reply','<path d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`replyEmail('${emailId}')`)}
      ${pBtn('Reply All','<path d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6M7 6h8a7 7 0 017 7" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`replyAllEmail('${emailId}')`)}
      ${pBtn('Forward','<path d="M21 10H11a8 8 0 00-8 8v2M21 10l-6 6m6-6l-6-6" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`forwardEmail('${emailId}')`)}
      ${pBtn('New Meeting','<path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`openNewMeetingFromEmail('${emailId}')`)}
      ${pBtn('Mark Unread','<path d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`markEmailUnread('${emailId}')`)}
      ${pBtn('Move','<path d="M3 7h7l2 2h9v10a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" stroke-width="1.8" stroke-linejoin="round"/><path d="M12 14h6m-2-2l2 2-2 2" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`openMoveEmailMenu('${emailId}')`)}
      <button onclick="toggleEmailAISummary('${emailId}')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #e9d5ff;background:#fff;color:#7c3aed;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s" onmouseover="this.style.background='#f5f3ff'" onmouseout="this.style.background='#fff'"><svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor"><path d="M12 1l2.7 8.3H23l-7 5.1 2.7 8.3L12 18l-7.7 4.7 2.7-8.3-7-5.1h8.3z"/></svg>Summarize</button>
      ${isTracked
        ?`<span style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #bbf7d0;background:#f0fdf4;color:#15803d;font-size:11px;font-weight:500">✓ In Tasks</span><button onclick="removeTaskByEmail('${emailId}')" style="display:inline-flex;align-items:center;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s" onmouseover="this.style.background='#fff1f2'" onmouseout="this.style.background='#fff'">Remove</button>`
        :`<button onclick="addTaskFromEmail('${emailId}')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:none;background:#0E3416;color:#fff;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:background .12s" onmouseover="this.style.background='#1A5C2A'" onmouseout="this.style.background='#0E3416'">+ Add to Tasks</button>`
      }
      ${(folder==='deleted'||folder==='deleteditems')
        ?`${pBtn('Restore','<path d="M3 12a9 9 0 019-9 9 9 0 016.7 3.1M21 3v6h-6M21 12a9 9 0 01-9 9 9 0 01-6.7-3.1" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',`restoreEmail('${emailId}')`,'background:#f0fdf4;border-color:#86efac;color:#166534;font-weight:700')}<button onclick="permanentDeleteEmail('${emailId}')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s" onmouseover="this.style.background='#fff1f2'" onmouseout="this.style.background='#fff'">Delete Permanently</button>`
        :`<button onclick="deleteEmailItem('${emailId}')" style="display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border-radius:4px;border:1px solid #fca5a5;background:#fff;color:#dc2626;font-family:Inter,sans-serif;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s" onmouseover="this.style.background='#fff1f2'" onmouseout="this.style.background='#fff'"><svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>Delete</button>`
      }`;
    const displayEmail=isMeeting
      ?{...email,body:{contentType:'html',content:renderMeetingCard(meetingData,email.subject)},bodyPreview:email.bodyPreview||meetingData?.description||''}
      :email;
    threadViewState={subject:email.subject||'(no subject)',messages:[displayEmail],backEmailId:emailId,backFolder:folder||currentFolder,webLink:email.webLink||'',showBack:false,actionsHTML};
    readerEl.innerHTML=threadHeaderHTML()+`
      <div id="email-ai-summary-panel" style="display:none;padding:10px 16px;background:#f5f3ff;border-bottom:1px solid #ddd6fe;flex-shrink:0">
        <div style="font-size:10px;font-weight:700;color:#7c3aed;margin-bottom:5px;letter-spacing:.04em">✦ AI SUMMARY</div>
        <div id="email-ai-summary-content" style="font-size:12.5px;line-height:1.6;color:#111"></div>
      </div>
      <div class="tv-list">${buildThreadHTML([displayEmail])}</div>`;
    await hydrateThreadFramesAndAttachments(token);
    // Check .ics attachments if no iCal found in body (always check for EventMessage emails)
    if(!isMeeting&&(email.hasAttachments||isEventMsg)){
      try{
        const atts=await fetchVisibleAttachments(emailId);
        const icsAtt=atts.find(a=>(a.name||'').toLowerCase().endsWith('.ics')||(a.contentType||'').toLowerCase().includes('calendar'));
        if(icsAtt){
          const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments/${icsAtt.id}`,{headers:{Authorization:`Bearer ${token}`}});
          if(r.ok){
            const d=await r.json();
            if(d.contentBytes){
              const icsRaw=atob(d.contentBytes);
              const icalParsed=parseICalData(icsRaw);
              displayEmail.body={contentType:'html',content:renderMeetingCard(icalParsed,email.subject)};
              threadViewState.messages=[displayEmail];
              await renderStoredThread(token);
            }
          }
        }
      }catch{}
    }
    if(!email.isRead&&outlookFolderCanBeUnread(folder)){
      const readRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}`,{method:'PATCH',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({isRead:true})});
      if(readRes.ok){
        email.isRead=true;
        var _unreadEl=el||document.getElementById('email-item-'+emailId);if(_unreadEl)_unreadEl.classList.remove('unread');
        const cached=(outlookFolderEmails[folder]||[]).find(e=>e.id===emailId);
        if(cached)cached.isRead=true;
        updateUnreadBadges(-1);
      }
    }
  }catch(err){
    readerEl.innerHTML='<div style="flex:1;display:flex;align-items:center;justify-content:center;color:#9ca3af;font-size:12px;padding:40px;text-align:center">Could not load email</div>';
  }
}

async function downloadAttachment(emailId,attId,filename){
  toast('Downloading...');
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments/${attId}/$value`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Download failed');
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=filename;a.click();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
    toast(`Downloaded: ${filename}`);
  }catch(err){toast('Download failed');}
}

// ── Right-click context menu ──────────────────────────────────────────────────
let _ctxEmailId=null,_ctxFolder=null;
function showEmailCtxMenu(e,emailId,folder){
  e.preventDefault();e.stopPropagation();
  _ctxEmailId=emailId;_ctxFolder=folder||currentFolder;
  const m=document.getElementById('email-ctx-menu');
  if(!m)return;
  const x=Math.min(e.clientX,window.innerWidth-200);
  const y=Math.min(e.clientY,window.innerHeight-180);
  m.style.left=x+'px';m.style.top=y+'px';m.style.display='block';
}
function hideEmailCtxMenu(){const m=document.getElementById('email-ctx-menu');if(m)m.style.display='none';}
document.addEventListener('click',hideEmailCtxMenu);
document.addEventListener('keydown',e=>{if(e.key==='Escape')hideEmailCtxMenu();});
async function ctxMarkUnread(){hideEmailCtxMenu();if(_ctxEmailId)await markEmailUnread(_ctxEmailId);}
async function ctxMarkRead(){
  hideEmailCtxMenu();
  if(!_ctxEmailId)return;
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${_ctxEmailId}`,{method:'PATCH',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({isRead:true})});
    if(res.ok){
      const em=emailCache[_ctxEmailId];if(em)em.isRead=true;
      const cached=(outlookFolderEmails[_ctxFolder||currentFolder]||[]).find(e=>e.id===_ctxEmailId);
      if(cached)cached.isRead=true;
      document.getElementById('email-item-'+_ctxEmailId)?.classList.remove('unread');
      toast('Marked as read');
    }
  }catch{toast('Could not mark as read');}
}
function ctxFlag(){hideEmailCtxMenu();if(_ctxEmailId)toggleOutlookFlag(_ctxEmailId,null);}
function ctxNewMeeting(){hideEmailCtxMenu();if(_ctxEmailId)openNewMeetingFromEmail(_ctxEmailId);}
function ctxMoveEmail(){hideEmailCtxMenu();if(_ctxEmailId)openMoveEmailMenu(_ctxEmailId);}
async function ctxDelete(){hideEmailCtxMenu();if(_ctxEmailId)await deleteEmailItem(_ctxEmailId);}
// ─────────────────────────────────────────────────────────────────────────────

async function markEmailUnread(emailId){
  if(!emailId)return;
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}`,{
      method:'PATCH',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({isRead:false})
    });
    if(res.ok){
      const email=emailCache[emailId];
      if(email)email.isRead=false;
      const folder=currentFolder||'inbox';
      const cached=(outlookFolderEmails[folder]||[]).find(e=>e.id===emailId);
      if(cached)cached.isRead=false;
      document.getElementById('email-item-'+emailId)?.classList.add('unread');
      updateUnreadBadges(1);
      toast('Marked as unread');
    }
  }catch{toast('Could not mark as unread');}
}

function openNewMeetingFromEmail(emailId){
  const email=emailCache[emailId]||{};
  const toRecipients=email.toRecipients||[];
  const ccRecipients=email.ccRecipients||[];
  const required=[...toRecipients,...ccRecipients].map(r=>r?.emailAddress?.address||'').filter(Boolean).join(', ');
  document.getElementById('meet-subject').value=email.subject?'Re: '+email.subject:'';
  setChipField('meet-required-chips','meet-required','meet-required-input',required);
  clearChipField('meet-optional-chips','meet-optional','meet-optional-input');
  document.getElementById('meet-location').value='';
  document.getElementById('meet-notes').value='';
  const tomorrow=new Date();tomorrow.setDate(tomorrow.getDate()+1);
  const y=tomorrow.getFullYear();
  const mo=String(tomorrow.getMonth()+1).padStart(2,'0');
  const d=String(tomorrow.getDate()).padStart(2,'0');
  document.getElementById('meet-date').value=`${y}-${mo}-${d}`;
  document.getElementById('meet-start').value='09:00';
  document.getElementById('meet-end').value='09:30';
  document.getElementById('meet-teams').checked=true;
  document.getElementById('mo-new-meeting').classList.add('open');
}

async function scheduleTeamsMeeting(){
  const subject=(document.getElementById('meet-subject').value||'').trim();
  const dateVal=document.getElementById('meet-date').value;
  const startVal=document.getElementById('meet-start').value;
  const endVal=document.getElementById('meet-end').value;
  if(!subject||!dateVal||!startVal||!endVal){toast('Please fill in subject, date, and time');return;}
  const emailOk=e=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
  const rawAttendees=[
    ...collectChipEmails('meet-required-chips','meet-required-input').split(/[,;]+/).map(address=>({address:address.trim(),type:'required'})),
    ...collectChipEmails('meet-optional-chips','meet-optional-input').split(/[,;]+/).map(address=>({address:address.trim(),type:'optional'}))
  ].filter(a=>a.address);
  const invalid=rawAttendees.find(a=>!emailOk(a.address));
  if(invalid){toast(`Check attendee email: ${invalid.address}`);return;}
  const unique=new Map();
  rawAttendees.forEach(a=>unique.set(a.address.toLowerCase(),a));
  const attendees=[...unique.values()].map(a=>({emailAddress:{address:a.address},type:a.type}));
  const location=(document.getElementById('meet-location').value||'').trim();
  const notes=(document.getElementById('meet-notes').value||'').trim();
  const teamsOn=document.getElementById('meet-teams').checked;
  // Parse as local time then convert to UTC so Graph stores the correct absolute time
  // regardless of what timezone the browser is in
  const startDT=new Date(`${dateVal}T${startVal}:00`);
  const endDT=new Date(`${dateVal}T${endVal}:00`);
  if(Number.isNaN(startDT.getTime())||Number.isNaN(endDT.getTime())){toast('Please enter a valid date and time');return;}
  if(endDT<=startDT){toast('End time must be after start time');return;}
  if(startDT.getTime()<Date.now()-60000){toast('Meeting start time cannot be in the past');return;}
  const toUTCStr=dt=>dt.toISOString().replace('Z','');
  const event={
    subject,
    start:{dateTime:toUTCStr(startDT),timeZone:'UTC'},
    end:{dateTime:toUTCStr(endDT),timeZone:'UTC'},
    attendees,
    responseRequested:true,
    allowNewTimeProposals:false,
    isOnlineMeeting:teamsOn,
    onlineMeetingProvider:teamsOn?'teamsForBusiness':'unknown',
    body:{contentType:'HTML',content:notes?`<p>${escapeHtml(notes)}</p>`:''}
  };
  if(globalThis.crypto?.randomUUID)event.transactionId=crypto.randomUUID();
  if(location)event.location={displayName:location};
  const btn=document.getElementById('meet-send-btn');
  if(btn){btn.disabled=true;btn.textContent='Sending...';}
  try{
    const token=await getAccessToken();
    const res=await fetch('https://graph.microsoft.com/v1.0/me/events',{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify(event)
    });
    if(!res.ok){const err=await res.text();throw new Error(`${res.status}: ${err}`);}
    closeMo('mo-new-meeting');
    toast('Meeting invite sent!');
    if(currentFolder==='schedule'||currentFolder==='calendar')await loadScheduleFolder();
  }catch(err){
    toast(String(err.message||'').startsWith('403:')?'Calendar write permission is required':'Could not send meeting invite');
    console.error('Meeting invite error:',err);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Send invitation';}
  }
}

async function rsvpMeeting(emailId,action){
  if(!emailId||!action)return;
  const map={accept:'accept',tentative:'tentativelyAccept',decline:'decline'};
  const endpoint=map[action];if(!endpoint)return;
  const labels={accept:'Accepted',tentative:'Tentative',decline:'Declined'};
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/${endpoint}`,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({sendResponse:true,comment:''})
    });
    if(!res.ok)throw new Error(await res.text());
    const container=document.querySelector(`[data-rsvp-email="${emailId}"]`);
    const colors={accept:'#15803d',tentative:'#b45309',decline:'#b91c1c'};
    if(container)container.innerHTML=`<span style="font-size:12px;font-weight:600;color:${colors[action]||'#374151'}">✓ ${labels[action]||action}</span>`;
    toast('Meeting '+( labels[action]||action));
  }catch(err){
    toast('Could not update meeting response');
    console.error('RSVP error:',err);
  }
}
async function toggleOutlookFlag(emailId,ev){
  ev?.stopPropagation();
  if(!emailId)return;
  const email=emailCache[emailId]||(outlookFolderEmails[currentFolder]||[]).find(e=>e.id===emailId)||{};
  const next=outlookFlagStatus(email)==='flagged'?'notFlagged':'flagged';
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}`,{
      method:'PATCH',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({flag:{flagStatus:next}})
    });
    if(!res.ok)throw new Error('Flag update failed');
    email.flag={...(email.flag||{}),flagStatus:next};
    if(emailCache[emailId])emailCache[emailId].flag=email.flag;
    Object.keys(outlookFolderEmails||{}).forEach(folder=>{
      const cached=(outlookFolderEmails[folder]||[]).find(e=>e.id===emailId);
      if(cached)cached.flag=email.flag;
    });
    if(currentFolder==='flagged'&&next!=='flagged'){
      outlookFolderEmails.flagged=(outlookFolderEmails.flagged||[]).filter(e=>e.id!==emailId);
    }
    updateFlaggedCount();
    if(currentFolder==='untracked')loadUntracked();
    else renderEmailRows(currentFolder,outlookFolderEmails[currentFolder]||[]);
    if(currentEmailId===emailId)readEmail(emailId,currentFolder==='untracked'?'sent':currentFolder);
    toast(next==='flagged'?'Flagged in Outlook':'Flag removed');
  }catch(err){
    toast('Could not update Outlook flag');
  }
}

function toggleEmailPin(emailId,ev){
  ev?.stopPropagation();
  if(!emailId)return;
  const email=emailCache[emailId]||(outlookFolderEmails[currentFolder]||[]).find(e=>e.id===emailId)||{id:emailId};
  const keys=[email.id||emailId,email.conversationId].filter(Boolean);
  const pinned=keys.some(k=>pinnedEmailIds.has(k));
  if(pinned){
    keys.forEach(k=>pinnedEmailIds.delete(k));
  }else{
    keys.forEach(k=>pinnedEmailIds.add(k));
  }
  savePinnedEmails();
  if(currentFolder==='untracked')loadUntracked();
  else renderEmailRows(currentFolder,outlookFolderEmails[currentFolder]||[]);
  toast(pinned?'Unpinned':'Pinned');
}

function openCompose(to,subject,body,type,emailId){
  setChipField('compose-to-chips','compose-to','compose-to-input',to||'');
  clearChipField('compose-cc-chips','compose-cc','compose-cc-input');
  clearChipField('compose-bcc-chips','compose-bcc','compose-bcc-input');
  document.getElementById('compose-subject').value=subject||'';
  document.getElementById('compose-body').value=body||'';
  document.getElementById('compose-title').textContent=type||'New Email';
  const pr=document.getElementById('compose-priority');if(pr)pr.value='normal';
  const dl=document.getElementById('compose-deadline');if(dl)dl.value='';
  const bccRow=document.getElementById('bcc-row');
  if(bccRow)bccRow.style.display='none';
  resetComposeAttachments();
  // Reset AI panels
  currentComposeType=type||'New Email';
  currentComposeEmailId=emailId||null;
  const aiPanel=document.getElementById('compose-ai-brief');
  const relPanel=document.getElementById('compose-related-tasks');
  if(aiPanel)aiPanel.style.display='none';
  if(relPanel)relPanel.style.display='none';
  applyComposeSignature();
  document.getElementById('mo-compose').classList.add('open');
  showRecipientPreview();
  // Load AI brief async when forwarding
  if(type==='Forward'&&emailId)loadForwardAIBrief(emailId);
}

function formatFileSize(bytes){
  if(!bytes)return '0 B';
  const units=['B','KB','MB','GB'];
  let size=bytes,idx=0;
  while(size>=1024&&idx<units.length-1){size/=1024;idx++;}
  return `${size.toFixed(size>=10||idx===0?0:1)} ${units[idx]}`;
}

function resetComposeAttachments(){
  currentComposeFiles=[];
  const input=document.getElementById('compose-file-input');
  if(input)input.value='';
  renderComposeAttachments();
}

function selectComposeAttachments(fileList){
  const incoming=Array.from(fileList||[]);
  for(const file of incoming){
    const key=`${file.name}|${file.size}|${file.lastModified}`;
    if(!currentComposeFiles.some(f=>`${f.name}|${f.size}|${f.lastModified}`===key)){
      currentComposeFiles.push(file);
    }
  }
  renderComposeAttachments();
}

function removeComposeAttachment(index){
  currentComposeFiles.splice(index,1);
  renderComposeAttachments();
}

function renderComposeAttachments(){
  const box=document.getElementById('compose-attachments');
  if(!box)return;
  if(!currentComposeFiles.length){box.style.display='none';box.innerHTML='';return;}
  box.style.display='block';
  box.innerHTML=`
    <div style="padding:8px 10px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:8px;background:#f8fafc">
      <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px">${currentComposeFiles.length} attachment${currentComposeFiles.length!==1?'s':''}</div>
      <button class="btn btn-ghost btn-xs" onclick="resetComposeAttachments()">Clear</button>
    </div>
    ${currentComposeFiles.map((file,i)=>`
      <div style="display:flex;align-items:center;gap:10px;padding:8px 10px;border-bottom:${i===currentComposeFiles.length-1?'none':'1px solid var(--border)'}">
        <svg width="15" height="15" fill="none" stroke="var(--muted)" viewBox="0 0 24 24" style="flex-shrink:0"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        <div style="flex:1;min-width:0">
          <div style="font-size:12.5px;font-weight:600;color:var(--body);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(file.name)}</div>
          <div style="font-size:11px;color:var(--muted)">${formatFileSize(file.size)}</div>
        </div>
        <button class="btn btn-ghost btn-xs" onclick="removeComposeAttachment(${i})" title="Remove attachment">Remove</button>
      </div>`).join('')}`;
}

async function replyEmail(emailId){
  const email=await ensureFullEmailForCompose(emailId);if(!email)return;
  const to=email.from?.emailAddress?.address||'';
  const subject=email.subject?.startsWith('Re:')?email.subject:`Re: ${email.subject}`;
  const orig=`\n\n${repliedMessageText(email)}`;
  openCompose(to,subject,orig,'Reply',emailId);
}

async function replyAllEmail(emailId){
  const email=await ensureFullEmailForCompose(emailId);if(!email)return;
  const me=normEmail(currentUser?.email||currentAccount?.username||'');
  const sender=email.from?.emailAddress?.address||'';
  const uniq=[];
  const add=addr=>{
    const e=extractEmailAddress(addr||'');
    if(!e||!e.includes('@')||normEmail(e)===me||uniq.some(x=>normEmail(x)===normEmail(e)))return;
    uniq.push(e);
  };
  add(sender);
  (email.toRecipients||[]).forEach(r=>add(r.emailAddress?.address));
  const ccList=[];
  (email.ccRecipients||[]).forEach(r=>{
    const e=extractEmailAddress(r.emailAddress?.address||'');
    if(!e||!e.includes('@')||normEmail(e)===me||uniq.some(x=>normEmail(x)===normEmail(e))||ccList.some(x=>normEmail(x)===normEmail(e)))return;
    ccList.push(e);
  });
  const subject=email.subject?.startsWith('Re:')?email.subject:`Re: ${email.subject}`;
  const orig=`\n\n${repliedMessageText(email)}`;
  openCompose(uniq.join(', '),subject,orig,'Reply All',emailId);
  document.getElementById('compose-cc').value=ccList.join(', ');
  showRecipientPreview();
}


// ============================================================
// PRE-SEND UNREAD THREAD CHECK
// ============================================================
let _pendingSendCallback=null;

function proceedSend(){
  if(_pendingSendCallback){const fn=_pendingSendCallback;_pendingSendCallback=null;fn();}
}

async function checkUnreadThreadsForRecipients(emails){
  // Find open tasks in the action log that match any of the recipient emails
  const lower=emails.map(e=>e.toLowerCase());
  return tasks.filter(t=>{
    if(nstt(t.status)==='Done')return false;
    return lower.some(e=>normEmail(t.email||'')===e);
  });
}

async function showUnreadCheckIfNeeded(recipientStr, onProceed){
  const emails=recipientStr.split(',').map(e=>e.trim().toLowerCase()).filter(Boolean);
  if(!emails.length){onProceed();return;}
  const openThreads=await checkUnreadThreadsForRecipients(emails);
  if(!openThreads.length){onProceed();return;}

  // Build gist cards for each open thread
  const listEl=document.getElementById('unread-check-list');
  document.getElementById('unread-check-title').textContent=
    `${openThreads.length} open thread${openThreads.length!==1?'s':''} with this recipient`;
  document.getElementById('unread-check-sub').textContent=
    'Review before sending — you may want to address these first.';

  listEl.innerHTML=openThreads.map(t=>`
    <div class="unread-thread-card">
      <div class="unread-thread-head">
        <div class="unread-thread-title">${emailSubject(t)}</div>
        ${sbadge(t)}
      </div>
      <div class="unread-thread-meta">
        <span style="font-weight:600">${t.person}</span>
        <span style="color:var(--muted)">${t.dept}</span>
        <span style="color:var(--muted)">${fmtD(t.date)}</span>
      </div>
      <div class="unread-thread-summary" id="uc-sum-${t.id}">
        ${t.summary?formatSummaryHTML(t.summary):'<span style="color:var(--muted);font-style:italic;font-size:12px">Loading AI gist…</span>'}
      </div>
      ${threadLinkButton(t)?`<div style="margin-top:6px">${threadLinkButton(t)}</div>`:''}
    </div>`).join('');

  document.getElementById('mo-unread-check').classList.add('open');
  _pendingSendCallback=onProceed;

  // If any thread has no summary yet, generate one in the background
  for(const t of openThreads){
    if(!t.summary&&(t.conversationId||t.emailId)){
      const emailProxy={id:t.emailId||t.lastMessageId,subject:t.emailSubject||t.title,conversationId:t.conversationId,bodyPreview:t.summary||''};
      attachFreeThreadSummary(emailProxy).then(()=>{
        t.summary=emailProxy.threadSummary||t.summary;
        t.threadSummary=emailProxy.threadSummary||t.threadSummary||'';
        const el=document.getElementById(`uc-sum-${t.id}`);
        if(el&&t.summary)el.innerHTML=formatSummaryHTML(t.summary);
      }).catch(()=>{});
    }
  }
}

async function sendEmail(){
  const to=collectChipEmails('compose-to-chips','compose-to-input');
  const cc=collectChipEmails('compose-cc-chips','compose-cc-input');
  const bcc=collectChipEmails('compose-bcc-chips','compose-bcc-input');
  const subject=document.getElementById('compose-subject').value.trim();
  const body=document.getElementById('compose-body').value.trim();
  const importance=document.getElementById('compose-priority')?.value||'normal';
  const deadline=document.getElementById('compose-deadline')?.value||'';
  if(!to){toast('Please add at least one recipient');return;}
  if(!subject){toast('Please add a subject');return;}

  // Check for open threads with any recipient before sending
  const allRecipients=[...to.split(','),...(cc?cc.split(','):[]),...(bcc?bcc.split(','):[])]
    .map(e=>normEmail(extractEmailAddress(e))).filter(Boolean);
  const createTask=document.getElementById('compose-create-task')?.checked!==false;
  await showUnreadCheckIfNeeded(allRecipients.join(','), ()=>doSendEmail(to,cc,bcc,subject,body,importance,deadline,createTask));
}

async function getForwardAttachmentNames(emailId){
  if(!emailId)return [];
  try{
    const token=await getAccessToken();
    const listRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments?$select=name,isInline`,{headers:{Authorization:`Bearer ${token}`}});
    if(!listRes.ok)return [];
    const list=await listRes.json();
    return (list.value||[]).filter(a=>!a.isInline).map(a=>a.name||'attachment');
  }catch(err){
    console.warn('Forward attachment lookup failed:',err.message);
    return [];
  }
}

async function getDraftAccessToken(){
  try{
    const result=await msalInstance.acquireTokenSilent({scopes:SCOPES_DRAFTS,account:currentAccount});
    return result.accessToken;
  }catch(err){
    try{
      const result=await msalInstance.acquireTokenPopup({scopes:SCOPES_DRAFTS});
      return result.accessToken;
    }catch(popupErr){
      console.error('Draft permission error:',popupErr);
      throw popupErr;
    }
  }
}

function fileToBase64(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=()=>resolve(String(reader.result||'').split(',')[1]||'');
    reader.onerror=()=>reject(reader.error||new Error('Could not read attachment'));
    reader.readAsDataURL(file);
  });
}

async function addSmallAttachmentToMessage(token,messageId,file){
  const contentBytes=await fileToBase64(file);
  const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({
      '@odata.type':'#microsoft.graph.fileAttachment',
      name:file.name,
      contentType:file.type||'application/octet-stream',
      contentBytes
    })
  });
  if(!res.ok){
    const errText=await res.text().catch(()=>'');
    throw new Error(`Attachment failed (${res.status}) ${errText}`);
  }
}

async function addInlineAttachmentToMessage(token,messageId,att){
  const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(att)
  });
  if(!res.ok){
    const errText=await res.text().catch(()=>'');
    throw new Error(`Inline attachment failed (${res.status}) ${errText}`);
  }
}

async function uploadLargeAttachmentToMessage(token,messageId,file){
  const sessionRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/createUploadSession`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({AttachmentItem:{attachmentType:'file',name:file.name,size:file.size,contentType:file.type||'application/octet-stream'}})
  });
  if(!sessionRes.ok){
    const errText=await sessionRes.text().catch(()=>'');
    throw new Error(`Attachment upload session failed (${sessionRes.status}) ${errText}`);
  }
  const session=await sessionRes.json();
  const chunkSize=3276800;
  let start=0;
  while(start<file.size){
    const end=Math.min(start+chunkSize,file.size)-1;
    const chunk=await file.slice(start,end+1).arrayBuffer();
    const uploadRes=await fetch(session.uploadUrl,{
      method:'PUT',
      headers:{
        'Content-Length':String(chunk.byteLength),
        'Content-Range':`bytes ${start}-${end}/${file.size}`
      },
      body:chunk
    });
    if(!uploadRes.ok&&uploadRes.status!==201&&uploadRes.status!==202){
      const errText=await uploadRes.text().catch(()=>'');
      throw new Error(`Attachment upload failed (${uploadRes.status}) ${errText}`);
    }
    start=end+1;
  }
}

async function addComposeAttachmentsToDraft(token,messageId,files){
  for(const file of files||[]){
    if(file.size>150*1024*1024)throw new Error(`${file.name} is larger than Outlook's 150 MB attachment limit`);
    if(file.size<3*1024*1024)await addSmallAttachmentToMessage(token,messageId,file);
    else await uploadLargeAttachmentToMessage(token,messageId,file);
  }
}

async function addInlineAttachmentsToDraft(token,messageId,inlineAttachments=[]){
  for(const att of inlineAttachments||[]){
    await addInlineAttachmentToMessage(token,messageId,att);
  }
}

async function sendDraftMessage(token,message,toAttach=[],inlineAttachments=[]){
  const draftRes=await fetch('https://graph.microsoft.com/v1.0/me/messages',{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(message)
  });
  if(!draftRes.ok){
    const errText=await draftRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph draft failed (${draftRes.status}) ${errText}`);
  }
  const draft=await draftRes.json();
  await addInlineAttachmentsToDraft(token,draft.id,inlineAttachments);
  await addComposeAttachmentsToDraft(token,draft.id,toAttach);
  const sendRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`}
  });
  if(!sendRes.ok){
    const errText=await sendRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph draft send failed (${sendRes.status}) ${errText}`);
  }
}

function forwardComposeInstruction(body){
  const text=String(body||'');
  const markerIndex=text.search(/(?:^|\n)\s*(?:---|———)\s*Forwarded Message\s*(?:---|———)/i);
  const instruction=markerIndex>=0?text.slice(0,markerIndex):text;
  return plainizeEmailBody(instruction);
}

function taskSummaryWithInstruction(instruction,context){
  const cleanInstruction=plainizeEmailBody(instruction);
  const cleanContext=plainizeEmailBody(context);
  if(!cleanInstruction)return cleanContext;
  return [
    `Action needed:\n${cleanInstruction}`,
    cleanContext?`Email context:\n${cleanContext}`:''
  ].filter(Boolean).join('\n\n');
}

async function sendForwardedOutlookMessage(token,emailId,toRecipients,ccRecipients,bccRecipients,subject,body,importance,toAttach=[],composed=null){
  const markerIndex=String(body||'').search(/\n\s*(?:---|———)\s*Forwarded Message\s*(?:---|———)/i);
  const commentBody=markerIndex>=0?String(body||'').slice(0,markerIndex).trim():body;
  const prepared=markerIndex>=0?prepareComposedEmail(commentBody):(composed||prepareComposedEmail(body));
  const createRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/createForward`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({message:{toRecipients},comment:prepared.html})
  });
  if(!createRes.ok){
    const errText=await createRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph createForward failed (${createRes.status}) ${errText}`);
  }
  const draft=await createRes.json();
  const patch={subject,importance};
  if(ccRecipients.length)patch.ccRecipients=ccRecipients;
  if(bccRecipients.length)patch.bccRecipients=bccRecipients;
  // Explicitly set from+sender to the forwarder's account so Exchange never
  // shows "on behalf of [original sender]" in the recipient's inbox
  if(currentUser?.email){
    const me={emailAddress:{address:currentUser.email,name:currentUser.name||currentUser.email}};
    patch.from=me;
    patch.sender=me;
  }
  const patchRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}`,{
    method:'PATCH',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(patch)
  });
  if(!patchRes.ok){
    const errText=await patchRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph forward draft update failed (${patchRes.status}) ${errText}`);
  }
  await addInlineAttachmentsToDraft(token,draft.id,prepared.inlineAttachments);
  await addComposeAttachmentsToDraft(token,draft.id,toAttach);
  const sendRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`}
  });
  if(!sendRes.ok){
    const errText=await sendRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph forward send failed (${sendRes.status}) ${errText}`);
  }
}

async function sendReplyOutlookMessage(token,emailId,toRecipients,ccRecipients,bccRecipients,body,importance,toAttach=[],composed=null){
  const markerIndex=String(body||'').search(/\n\s*(?:---|———)\s*Original Message\s*(?:---|———)/i);
  const commentBody=markerIndex>=0?String(body||'').slice(0,markerIndex).trim():body;
  const prepared=markerIndex>=0?prepareComposedEmail(commentBody):(composed||prepareComposedEmail(body));
  const createRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/createReply`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({})
  });
  if(!createRes.ok){
    const errText=await createRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph createReply failed (${createRes.status}) ${errText}`);
  }
  const draft=await createRes.json();
  const patch={importance};
  const replyIntro=prepared.html||'';
  let originalBody=draft.body?.content||'';
  let originalBodyType=(draft.body?.contentType||'html').toLowerCase();
  if(!originalBody){
    const bodyRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}?$select=body`,{headers:{Authorization:`Bearer ${token}`}});
    if(bodyRes.ok){const bodyData=await bodyRes.json();originalBody=bodyData.body?.content||'';originalBodyType=(bodyData.body?.contentType||originalBodyType).toLowerCase();}
  }
  if(originalBody&&originalBodyType==='text'){originalBody=escapeHtml(originalBody).replace(/\r?\n/g,'<br>');}
  if(replyIntro&&originalBody){
    const bIdx=originalBody.indexOf('<body');
    if(bIdx>=0){const bEnd=originalBody.indexOf('>',bIdx)+1;patch.body={contentType:'HTML',content:originalBody.slice(0,bEnd)+replyIntro+'<br><br>'+originalBody.slice(bEnd)};}
    else{patch.body={contentType:'HTML',content:replyIntro+'<br><br>'+originalBody};}
  }else if(replyIntro||originalBody){patch.body={contentType:'HTML',content:replyIntro||originalBody};}
  if(toRecipients.length)patch.toRecipients=toRecipients;
  if(ccRecipients.length)patch.ccRecipients=ccRecipients;
  if(bccRecipients.length)patch.bccRecipients=bccRecipients;
  const patchRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}`,{
    method:'PATCH',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(patch)
  });
  if(!patchRes.ok){
    const errText=await patchRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph reply draft update failed (${patchRes.status}) ${errText}`);
  }
  await addInlineAttachmentsToDraft(token,draft.id,prepared.inlineAttachments);
  await addComposeAttachmentsToDraft(token,draft.id,toAttach);
  const sendRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`}
  });
  if(!sendRes.ok){
    const errText=await sendRes.text().catch(()=>'');
    throw new Error(`Microsoft Graph reply send failed (${sendRes.status}) ${errText}`);
  }
}

async function doSendEmail(to,cc,bcc,subject,body,importance='normal',deadline='',createTask=true){
  const btn=document.getElementById('compose-send-btn');
  if(btn){btn.disabled=true;btn.innerHTML='Sending...';}
  try{
    const toRecipients=to.split(',').map(e=>({emailAddress:{address:extractEmailAddress(e)}})).filter(r=>r.emailAddress.address&&r.emailAddress.address.includes('@'));
    const ccRecipients=cc?cc.split(',').map(e=>({emailAddress:{address:extractEmailAddress(e)}})).filter(r=>r.emailAddress.address&&r.emailAddress.address.includes('@')):[];
    const bccRecipients=bcc?bcc.split(',').map(e=>({emailAddress:{address:extractEmailAddress(e)}})).filter(r=>r.emailAddress.address&&r.emailAddress.address.includes('@')):[];
    const isOutlookForward=currentComposeType==='Forward'&&currentComposeEmailId;
    const isOutlookReply=(currentComposeType==='Reply'||currentComposeType==='Reply All')&&currentComposeEmailId;
    const needsDraftPermission=isOutlookForward||isOutlookReply||currentComposeFiles.length>0;
    const token=needsDraftPermission?await getDraftAccessToken():await getAccessToken();
    const forwardAttachmentNames=isOutlookForward?await getForwardAttachmentNames(currentComposeEmailId):[];
    const manualAttachmentNames=currentComposeFiles.map(f=>f.name);
    const composed=prepareComposedEmail(body);
    if(isOutlookForward){
      await sendForwardedOutlookMessage(token,currentComposeEmailId,toRecipients,ccRecipients,bccRecipients,subject,body,importance,currentComposeFiles,composed);
    }else if(isOutlookReply){
      await sendReplyOutlookMessage(token,currentComposeEmailId,toRecipients,ccRecipients,bccRecipients,body,importance,currentComposeFiles,composed);
    }else{
      const message={subject,importance,body:{contentType:'HTML',content:composed.html},toRecipients};
      if(ccRecipients.length)message.ccRecipients=ccRecipients;
      if(bccRecipients.length)message.bccRecipients=bccRecipients;
      if(currentComposeFiles.length){
        await sendDraftMessage(token,message,currentComposeFiles,composed.inlineAttachments);
      }else{
        if(composed.inlineAttachments.length)message.attachments=composed.inlineAttachments;
        const sendRes = await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
          method:'POST',
          headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
          body:JSON.stringify({message,saveToSentItems:true})
        });
        if(!sendRes.ok){
          const errText=await sendRes.text().catch(()=>"");
          throw new Error(`Microsoft Graph sendMail failed (${sendRes.status}) ${errText}`);
        }
      }
    }
    closeMo('mo-compose');
    if(btn){btn.disabled=false;btn.innerHTML='<svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg> Send';}
    if(createTask){
      const dpegRecipients=getDpegRecipients(to,cc);
      if(dpegRecipients.length){
        let forwardedEmail=null;
        let forwardedFullBody='';
        let taskInstruction='';
        if(currentComposeType==='Forward'&&currentComposeEmailId){
          taskInstruction=forwardComposeInstruction(body);
          forwardedEmail=emailCache[currentComposeEmailId]||null;
          if(forwardedEmail){
            await attachFreeThreadSummary(forwardedEmail).catch(()=>{});
            const rawBody=forwardedEmail.body?.content||'';
            const isHtml=(forwardedEmail.body?.contentType||'').toLowerCase()==='html';
            forwardedFullBody=plainizeEmailBody(rawBody?(isHtml?stripEmailHtml(rawBody):rawBody):'');
          }
        }
        for(const email of dpegRecipients){
          const p=findPersonByEmail(email);
          const person=p?.name||email.split('@')[0];
          const emailContext=forwardedFullBody||forwardedEmail?.threadSummary||forwardedEmail?.bodyPreview||body;
          const taskSummary=[
            taskSummaryWithInstruction(taskInstruction,emailContext),
            [...forwardAttachmentNames,...manualAttachmentNames].length?`Attachments: ${[...forwardAttachmentNames,...manualAttachmentNames].join(', ')}`:''
          ].filter(Boolean).join('\n\n');
          const newTask={
            id:Date.now()+Math.random(),
            assignedAt:new Date().toISOString(),
            createdAt:new Date().toISOString(),
            title:forwardedEmail?.subject||subject,
            emailSubject:forwardedEmail?.subject||subject,
            emailId:forwardedEmail?.id||'',
            conversationId:forwardedEmail?.conversationId||'',
            lastMessageId:forwardedEmail?.id||'',
            summary:taskSummary,
            taskInstruction,
            threadSummary:forwardedEmail?.threadSummary||'',
            attachmentNames:[...forwardAttachmentNames,...manualAttachmentNames],
            aiGenerated:!!forwardedEmail?.aiGenerated,
            person,email,dept:personDept(email,person),date:deadline||new Date().toISOString().split('T')[0],deadline:deadline||'',status:'Pending',priority:importance==='high'?'High':'Normal',wednesday:false,followup:'',weekOffset:0,
            replyCount:forwardedEmail?.conversationId?undefined:1
          };
          tasks.unshift(newTask);
          await sendTaskNotification(newTask).catch(err=>console.warn('Task notification failed:',err.message));
          await createToDoTask(newTask).catch(err=>console.warn('To Do creation failed:',err.message));
        }
        await saveTasksToOneDrive();
        refreshAll();
        toast(`Email sent. ${dpegRecipients.length} task${dpegRecipients.length!==1?'s':''} created in Action Log.`);
      }else{
        toast('Email sent');
      }
    }else{
      toast('Email sent');
    }
    if(currentFolder==='sent')loadFolder('sent');
  }catch(err){
    console.error('Send error:',err);
    const msg=String(err?.errorMessage||err?.message||err||'');
    if(msg.includes('AADSTS90094')||msg.toLowerCase().includes('admin')){
      toast('Admin approval is needed only for Outlook-style forwards/attachments. Email was not sent.');
    }else{
      toast('Email was not sent. No Action Log task was created.');
    }
    if(btn){btn.disabled=false;btn.innerHTML='<svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg> Send';}
  }
}

// ============================================================
// FORWARD THREAD GAP CHECK
// ============================================================
let _pendingForwardFn=null;
function proceedForward(){closeMo('mo-fwd-gap');if(_pendingForwardFn){const fn=_pendingForwardFn;_pendingForwardFn=null;fn();}}

async function ensureFullEmailForCompose(emailId){
  let email=emailCache[emailId];
  if(email?.body?.content)return email;
  const token=await getAccessToken();
  const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,isRead,conversationId,webLink,importance,hasAttachments,flag';
  const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=${select}`,{headers:{Authorization:`Bearer ${token}`}});
  if(!res.ok)return email||null;
  email=await res.json();
  emailCache[emailId]=email;
  return email;
}

function emailFullText(email){
  const body=email?.body?.content||'';
  if(body){
    const type=String(email?.body?.contentType||'').toLowerCase();
    return type==='html'?plainizeEmailBody(stripEmailHtml(body)):plainizeEmailBody(body);
  }
  return plainizeEmailBody(email?.bodyPreview||'');
}

function emailAuthoredText(email){
  const body=email?.body?.content||email?.bodyPreview||'';
  if(!body)return '';
  const isHtml=String(email?.body?.contentType||'').toLowerCase()==='html';
  const authored=threadMessageBody(email,isHtml);
  if(!isHtml)return plainizeEmailBody(authored);
  // The compose editor is plain text. Convert Outlook's remaining authored
  // HTML after quote removal so tags/styles never appear in Forward or
  // Forward All. Do not use stripEmailHtml here because it intentionally
  // removes some Outlook classes that can also contain the authored reply.
  try{
    const div=document.createElement('div');
    div.innerHTML=authored;
    div.querySelectorAll('style,script,meta,head,title').forEach(el=>el.remove());
    div.querySelectorAll('br').forEach(el=>el.replaceWith('\n'));
    div.querySelectorAll('p,div,li,tr,h1,h2,h3,h4,h5,h6').forEach(el=>{
      if(el.previousSibling)el.insertAdjacentText('beforebegin','\n');
      el.insertAdjacentText('beforeend','\n');
    });
    div.querySelectorAll('td,th').forEach(el=>el.insertAdjacentText('beforeend',' '));
    return plainizeEmailBody(div.textContent||div.innerText||'');
  }catch{
    return plainizeEmailBody(authored.replace(/<[^>]*>/g,' '));
  }
}

function forwardedMessageText(email){
  const from=email?.from?.emailAddress;
  const toLine=forwardRecipientsLine('To',email?.toRecipients||[]);
  const ccLine=forwardRecipientsLine('CC',email?.ccRecipients||[]);
  const dateLine=forwardThreadDate(email);
  const lines=[
    '--- Forwarded Message ---',
    forwardPersonLine('From',from),
    toLine,
    ccLine,
    dateLine?`Date: ${dateLine}`:'',
    `Subject: ${email?.subject||''}`,
  ].filter(Boolean);
  return `\n\n${lines.join('\n')}\n\n${emailAuthoredText(email)}`.trimEnd();
}

function repliedMessageText(email){
  const from=email?.from?.emailAddress;
  const toLine=forwardRecipientsLine('To',email?.toRecipients||[]);
  const ccLine=forwardRecipientsLine('CC',email?.ccRecipients||[]);
  const dateLine=forwardThreadDate(email);
  const lines=[
    '——— Original Message ———',
    forwardPersonLine('From',from),
    dateLine?`Sent: ${dateLine}`:'',
    toLine,
    ccLine,
    `Subject: ${email?.subject||''}`,
  ].filter(Boolean);
  return `${lines.join('\n')}\n\n${emailFullText(email)}`.trimEnd();
}

async function forwardEmail(emailId){
  const email=await ensureFullEmailForCompose(emailId);if(!email)return;
  const subject=email.subject?.startsWith('Fw:')?email.subject:`Fw: ${email.subject}`;
  const orig=`\n\n${forwardedMessageText(email)}`;
  const doForward=()=>openCompose('',subject,orig,'Forward',emailId);
  if(!email.conversationId){doForward();return;}
  try{
    const token=await getAccessToken();
    const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=id,subject,from,receivedDateTime,sentDateTime,bodyPreview&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok){doForward();return;}
    const data=await res.json();
    const allMsgs=(data.value||[]).sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime)-new Date(b.receivedDateTime||b.sentDateTime));
    const emailDate=new Date(email.receivedDateTime||email.sentDateTime);
    const missed=allMsgs.filter(m=>new Date(m.receivedDateTime||m.sentDateTime)>emailDate);
    if(!missed.length){doForward();return;}
    _pendingForwardFn=doForward;
    const cnt=missed.length;
    document.getElementById('fwd-gap-sub').textContent=`${cnt} newer message${cnt!==1?'s':''} in this thread won't be forwarded`;
    const bodyEl=document.getElementById('fwd-gap-body');
    bodyEl.innerHTML=`
      <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:6px;padding:12px;margin-bottom:14px">
        <div style="font-size:12px;font-weight:700;color:#c2410c;margin-bottom:3px">You are forwarding an older message</div>
        <div style="font-size:12px;color:#9a3412">The recipient will not see the ${cnt} newer repl${cnt!==1?'ies':'y'} below.</div>
      </div>
      <div style="font-size:10.5px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.7px;margin-bottom:8px">Messages that will be skipped</div>
      ${missed.map(m=>{const from=m.from?.emailAddress;const dt=new Date(m.receivedDateTime||m.sentDateTime);return `<div style="border:1px solid var(--border);border-radius:6px;padding:10px 12px;margin-bottom:7px"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px"><span style="font-size:12px;font-weight:600;color:var(--body)">${from?.name||from?.address||'Unknown'}</span><span style="font-size:11px;color:var(--muted)">${dt.toLocaleDateString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'})}</span></div><div style="font-size:12px;color:var(--sub);line-height:1.5">${(m.bodyPreview||'').slice(0,160)}${(m.bodyPreview||'').length>160?'…':''}</div></div>`;}).join('')}
      <div id="fwd-gap-gist"></div>`;
    document.getElementById('mo-fwd-gap').classList.add('open');
    // Generate AI gist of missed messages in background
    const fnUrl=localStorage.getItem('dpeg_ai_fn_url');
    if(fnUrl){
      const gistEl=document.getElementById('fwd-gap-gist');
      if(gistEl){
        gistEl.innerHTML='<div style="margin-top:10px;padding:10px 12px;background:var(--leaf-bg,#f0fdf4);border:1px solid var(--leaf-bd,#86efac);border-radius:6px"><div style="font-size:11px;font-weight:700;color:var(--fern,#166534);margin-bottom:4px">✦ AI Gist of skipped messages</div><div style="font-size:12px;color:var(--muted);font-style:italic">Generating…</div></div>';
        callAISummary(missed,`Missed replies in: ${email.subject||''}`).then(gist=>{
          if(gist&&gistEl.isConnected)gistEl.innerHTML=`<div style="margin-top:10px;padding:10px 12px;background:var(--leaf-bg,#f0fdf4);border:1px solid var(--leaf-bd,#86efac);border-radius:6px"><div style="font-size:11px;font-weight:700;color:var(--fern,#166534);margin-bottom:4px">✦ AI Gist of skipped messages</div>${formatSummaryHTML(gist)}</div>`;
          else if(gistEl.isConnected){const localG=localThreadSummary(missed);gistEl.innerHTML=`<div style="margin-top:10px;padding:10px 12px;background:#f8fafc;border:1px solid var(--border);border-radius:6px"><div style="font-size:11px;font-weight:700;color:var(--muted);margin-bottom:4px">Gist of skipped messages</div>${formatSummaryHTML(localG)}</div>`;}
        }).catch(()=>{if(gistEl.isConnected){const localG=localThreadSummary(missed);gistEl.innerHTML=`<div style="margin-top:10px;padding:10px 12px;background:#f8fafc;border:1px solid var(--border);border-radius:6px"><div style="font-size:11px;font-weight:700;color:var(--muted);margin-bottom:4px">Gist of skipped messages</div>${formatSummaryHTML(localG)}</div>`;}});
      }
    }
  }catch(err){
    console.error('Forward gap check:',err);
    doForward();
  }
}

function forwardThreadDate(m){
  const dt=m?.receivedDateTime||m?.sentDateTime;
  if(!dt)return '';
  try{return centralDate(new Date(dt),{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});}
  catch{return new Date(dt).toLocaleString();}
}

function forwardPersonLine(label,person){
  const e=person?.emailAddress||person||{};
  const name=e.name||'';
  const addr=e.address||'';
  if(!name&&!addr)return '';
  return `${label}: ${name}${addr&&addr!==name?` <${addr}>`:''}`;
}

function forwardRecipientsLine(label,recipients){
  const list=(recipients||[]).map(r=>{
    const e=r?.emailAddress||{};
    return e.name&&e.address&&e.name!==e.address?`${e.name} <${e.address}>`:(e.address||e.name||'');
  }).filter(Boolean);
  return list.length?`${label}: ${list.join(', ')}`:'';
}

async function getMessagesForForwardAll(emailId){
  const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,conversationId,webLink,hasAttachments';
  let email=emailCache[emailId];
  if(!email?.conversationId||!email?.body){
    const token=await getAccessToken();
    const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}?$select=${select}`,{headers:{Authorization:`Bearer ${token}`}});
    if(r.ok){email=await r.json();emailCache[emailId]=email;}
  }
  if(!email)return [];
  let messages=[];
  const visibleThread=threadViewState?.messages||[];
  if(email.conversationId&&visibleThread.some(m=>m.id===email.id||m.conversationId===email.conversationId)){
    messages=visibleThread;
  }else if(email.conversationId){
    const token=await getAccessToken();
    const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=50&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
    if(res.ok){const data=await res.json();messages=data.value||[];}
  }
  if(!messages.length)messages=[email];
  messages.forEach(m=>{if(m?.id)emailCache[m.id]=m;});
  return messages.sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime)-new Date(b.receivedDateTime||b.sentDateTime));
}

async function forwardFullThread(emailId){
  try{
    const messages=await getMessagesForForwardAll(emailId);
    if(!messages.length){toast('Could not load thread to forward');return;}
    const latest=messages[messages.length-1];
    const subject=String(latest.subject||'').startsWith('Fw:')?latest.subject:`Fw: ${latest.subject||'Email thread'}`;
    const cleanThread=messages.map(m=>forwardedMessageText(m)).join('\n\n');
    openCompose('',subject,`\n\n${cleanThread}`,'Forward',latest.id);
  }catch(err){
    console.error('Forward all error:',err);
    toast('Could not prepare full thread forward');
  }
}

// ============================================================
// FORWARD AI ASSIST
// ============================================================
async function loadForwardAIBrief(emailId){
  const panel=document.getElementById('compose-ai-brief');
  const content=document.getElementById('compose-ai-brief-content');
  if(!panel||!content)return;
  panel.style.display='block';
  content.innerHTML='<span style="color:var(--muted);font-style:italic">Generating AI brief…</span>';
  window._fwdThreadMessages=[];
  window._fwdEmailId=emailId;
  try{
    const email=emailCache[emailId];
    if(!email){panel.style.display='none';return;}
    let messages=[email];
    if(email.conversationId){
      const token=await getAccessToken();
      const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
      const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=id,subject,from,toRecipients,receivedDateTime,body,bodyPreview,flag&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
      if(res.ok){const d=await res.json();if(d.value?.length)messages=d.value;}
    }
    messages.sort((a,b)=>new Date(a.receivedDateTime)-new Date(b.receivedDateTime));
    window._fwdThreadMessages=messages;
    const fwdMsg=messages.find(m=>m.id===emailId);
    const fwdTime=fwdMsg?new Date(fwdMsg.receivedDateTime):null;
    const newerMsgs=fwdTime?messages.filter(m=>new Date(m.receivedDateTime)>fwdTime):[];
    const ai=await callAISummary(messages,email.subject||'');
    const summary=ai||localThreadSummary(messages);
    if(!panel.isConnected)return;
    let html=formatSummaryHTML(summary);
    if(newerMsgs.length){
      const latest=newerMsgs[newerMsgs.length-1];
      const latestName=latest?.from?.emailAddress?.name||latest?.from?.emailAddress?.address||'someone';
      const latestDate=new Date(latest?.receivedDateTime);
      const latestStr=latestDate.toLocaleDateString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
      html=`<div style="margin-bottom:8px;padding:7px 10px;background:#fffbeb;border:1px solid #fcd34d;border-radius:5px;display:flex;align-items:center;gap:8px">
        <svg width="14" height="14" fill="none" stroke="#d97706" viewBox="0 0 24 24" style="flex-shrink:0"><path d="M12 9v4m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        <span style="font-size:11.5px;color:#92400e;flex:1"><strong>${newerMsgs.length} newer message${newerMsgs.length>1?'s':''}</strong> in this thread after this email — last from <strong>${latestName}</strong> on ${latestStr}. View before forwarding?</span>
        <button class="btn btn-ghost btn-xs" onclick="viewFullThread('${emailId}')" style="font-size:10.5px;white-space:nowrap;flex-shrink:0">View Thread</button>
      </div>`+html;
    }
    content.innerHTML=html;
  }catch(e){
    if(panel?.isConnected)panel.style.display='none';
  }
}

function checkRecipientReplied(recipientEmail){
  const msgs=window._fwdThreadMessages||[];
  const fwdId=window._fwdEmailId||'';
  if(!msgs.length||!recipientEmail||!fwdId)return;
  const norm=normEmail(recipientEmail);
  const fwdMsg=msgs.find(m=>m.id===fwdId);
  const fwdTime=fwdMsg?new Date(fwdMsg.receivedDateTime):new Date(0);
  const replied=msgs.find(m=>normEmail(m.from?.emailAddress?.address||'')===norm&&new Date(m.receivedDateTime)>fwdTime);
  if(!replied)return;
  const panel=document.getElementById('compose-ai-brief');
  const content=document.getElementById('compose-ai-brief-content');
  if(!panel||!content)return;
  if(content.querySelector('.recipient-replied-warn'))return;
  const name=replied.from?.emailAddress?.name||recipientEmail;
  const dt=new Date(replied.receivedDateTime);
  const dtStr=dt.toLocaleDateString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
  const warn=document.createElement('div');
  warn.className='recipient-replied-warn';
  warn.style.cssText='margin-bottom:8px;padding:7px 10px;background:#fef3c7;border:1px solid #f59e0b;border-radius:5px;display:flex;align-items:center;gap:8px';
  warn.innerHTML=`<svg width="14" height="14" fill="none" stroke="#b45309" viewBox="0 0 24 24" style="flex-shrink:0"><path d="M12 9v4m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span style="font-size:11.5px;color:#78350f;flex:1"><strong>${name}</strong> already replied to this thread on <strong>${dtStr}</strong>. See their reply before forwarding?</span><button class="btn btn-ghost btn-xs" onclick="viewFullThread('${fwdId}')" style="font-size:10.5px;white-space:nowrap;flex-shrink:0">View Reply</button>`;
  content.insertBefore(warn,content.firstChild);
  panel.style.display='block';
}

function emailSubjectMatch(task,subject){
  if(!subject||(!(task.emailSubject||task.title)))return false;
  const clean=s=>String(s||'').toLowerCase().replace(/^(re:|fw:|fwd:)\s*/gi,'').trim();
  const ts=clean(task.emailSubject||task.title);
  const es=clean(subject);
  if(!ts||!es)return false;
  const esSlice=es.slice(0,25);const tsSlice=ts.slice(0,25);
  return ts.includes(esSlice)||es.includes(tsSlice);
}

function showForwardRelatedTasks(recipientEmail){
  const panel=document.getElementById('compose-related-tasks');
  const content=document.getElementById('compose-related-tasks-content');
  if(!panel||!content)return;
  const fwdSubject=document.getElementById('compose-subject')?.value||'';
  const norm=normEmail(recipientEmail);
  // Primary: subject match only (most relevant)
  let related=fwdSubject?tasks.filter(t=>nstt(t.status)!=='Done'&&emailSubjectMatch(t,fwdSubject)):[];
  // Fallback: same person, recent tasks within 60 days, only if no subject match
  if(!related.length&&norm){
    const cutoff=Date.now()-60*24*60*60*1000;
    related=tasks.filter(t=>
      nstt(t.status)!=='Done'&&
      normEmail(t.email||'')===norm&&
      (!t.createdAt||new Date(t.createdAt).getTime()>cutoff)
    );
  }
  related=related.slice(0,3);
  if(!related.length){panel.style.display='none';return;}
  content.innerHTML=related.map(t=>`
    <div style="display:flex;align-items:flex-start;gap:8px;padding:5px 0;border-bottom:1px solid #fed7aa">
      <div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:600;color:var(--body);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${emailSubject(t)}</div>
        <div style="font-size:11px;color:var(--muted);margin-top:1px">${t.person||''} · ${sbadge(t)}</div>
      </div>
      ${taskEmailId(t)?`<button class="btn btn-ghost btn-xs" onclick="openTaskThread(${t.id})" style="flex-shrink:0;font-size:10.5px;padding:3px 7px">View</button>`:''}
    </div>`).join('');
  panel.style.display='block';
}

// Per-email AI summary in list view
function renderAIBullets(summary){
  if(!summary)return'<span style="color:var(--muted);font-size:11px;font-style:italic">No summary available.</span>';
  const lines=summary.split('\n')
    .map(l=>l.replace(/^[•\-\*▾▲◆]\s*/,'').replace(/^(About|Latest|Action needed|Action):\s*/i,'').trim())
    .filter(l=>l.length>8)
    .slice(0,3);
  if(!lines.length)return'<span style="color:var(--muted);font-size:11px;font-style:italic">No summary available.</span>';
  return lines.map(l=>'<div style="display:flex;gap:6px;align-items:flex-start;margin-bottom:4px"><span style="color:#7c3aed;font-size:7px;margin-top:5px;flex-shrink:0">●</span><span style="font-size:12px;line-height:1.5;color:#1e1b4b">'+escapeHtml(l)+'</span></div>').join('');
}
async function toggleListAISummary(emailId,btn,folder){
  const panel=document.getElementById('list-ai-sum-'+emailId);
  if(!panel)return;
  if(panel.style.display!=='none'){
    panel.style.display='none';
    btn.style.background='none';
    btn.style.borderColor='#e9d5ff';
    return;
  }
  panel.style.display='block';
  btn.style.background='#f5f3ff';
  btn.style.borderColor='#c4b5fd';
  if(panel.dataset.loaded)return;
  panel.innerHTML='<span style="color:var(--muted);font-style:italic;font-size:11px;display:flex;align-items:center;gap:5px"><svg width="10" height="10" viewBox="0 0 24 24" fill="#7c3aed"><path d="M12 1l2.7 8.3H23l-7 5.1 2.7 8.3L12 18l-7.7 4.7 2.7-8.3-7-5.1h8.3z"/></svg>Summarizing…</span>';
  try{
    let email=emailCache[emailId];
    if(!email){
      const token=await getAccessToken();
      const res=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+emailId+'?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,conversationId,flag,importance,hasAttachments',{headers:{Authorization:'Bearer '+token}});
      if(!res.ok)throw new Error('fetch failed');
      email=await res.json();
      emailCache[emailId]=email;
    }
    const messages=await fetchThreadMessagesForSummary(email,15);
    const atts=[];
    for(const m of messages){
      if(m.hasAttachments){try{atts.push(...await fetchVisibleAttachments(m.id));}catch{}}
    }
    const attachmentNames=atts.map(a=>a.name);
    const ai=await callAISummary(messages,email.subject||'',{attachmentNames});
    const summary=ai||localThreadSummary(messages);
    let html=renderAIBullets(summary);
    if(atts.length){
      html+=`<div style="margin-top:5px;font-size:10px;color:#6b7280"><span style="font-weight:600">${atts.length} attachment${atts.length>1?'s':''}</span>: ${attachmentNames.join(', ')}</div>`;
    }
    panel.innerHTML=html;
    panel.dataset.loaded='1';
  }catch(e){
    panel.innerHTML='<span style="color:#b91c1c;font-size:11px">Could not summarize.</span>';
  }
}
// Per-email AI summary in reading pane
async function toggleEmailAISummary(emailId){
  const panel=document.getElementById('email-ai-summary-panel');
  const content=document.getElementById('email-ai-summary-content');
  if(!panel||!content)return;
  if(panel.style.display!=='none'){panel.style.display='none';return;}
  panel.style.display='block';
  if(panel.dataset.loaded)return;
  content.innerHTML='<span style="color:var(--muted);font-style:italic;font-size:12px">Summarizing…</span>';
  try{
    const email=emailCache[emailId];
    if(!email){content.innerHTML='<span style="color:#b91c1c">Email not loaded.</span>';return;}
    // For meeting request emails, build summary from structured meeting data
    if(String(email['@odata.type']||'').toLowerCase().includes('eventmessage')){
      const mtg=buildMeetingDataFromGraph(email);
      const fmtDt=(dt)=>centralDate(dt,{weekday:'short',month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});
      const lines=[];
      if(mtg.summary)lines.push('Meeting: '+mtg.summary);
      if(mtg.organizer)lines.push('Organiser: '+mtg.organizer);
      if(mtg.dtstart)lines.push('Date: '+fmtDt(mtg.dtstart)+(mtg.dtend?' – '+fmtDt(mtg.dtend):''));
      if(mtg.location)lines.push('Location: '+mtg.location);
      if(mtg.attendees.length)lines.push('Attendees: '+mtg.attendees.join(', '));
      const meetText=lines.join('\n');
      const fakeMsgs=[{body:{contentType:'text',content:meetText},receivedDateTime:email.receivedDateTime,sentDateTime:email.sentDateTime,from:email.from}];
      const ai=await callAISummary(fakeMsgs,email.subject||'');
      content.innerHTML=formatSummaryHTML(ai||meetText);
      panel.dataset.loaded='1';
      return;
    }
    let messages=await fetchThreadMessagesForSummary(email,15);
    // Collect visible attachment names from thread (don't fetch content yet)
    const allAtts=[];
    for(const m of messages){
      if(m.hasAttachments){try{const atts=await fetchVisibleAttachments(m.id);allAtts.push(...atts);}catch{}}
    }
    const attachmentNames=allAtts.map(a=>a.name);
    const ai=await callAISummary(messages,email.subject||'',{attachmentNames});
    const summary=ai||localThreadSummary(messages);
    let html=formatSummaryHTML(summary);
    if(allAtts.length){
      const nameList=attachmentNames.map(n=>`<span style="font-style:italic">${escapeHtml(n)}</span>`).join(', ');
      html+=`<div id="att-footer-${emailId}" style="margin-top:8px;padding-top:7px;border-top:1px solid #f0f0f0;font-size:11px;color:#6b7280">
        📎 This email has ${allAtts.length} attachment${allAtts.length>1?'s':''}: ${nameList}. Would you like me to summarise them?
        <div style="display:flex;gap:6px;margin-top:5px">
          <button id="att-sum-btn-${emailId}" onclick="summarizeAttachments('${emailId}')" style="padding:3px 10px;font-size:10px;border:none;border-radius:3px;background:#0E3416;color:#fff;cursor:pointer;font-weight:600">Yes, summarise attachments</button>
          <button onclick="document.getElementById('att-footer-${emailId}').style.display='none'" style="padding:3px 10px;font-size:10px;border:1px solid #d1d5db;border-radius:3px;background:#fff;cursor:pointer;color:#374151">No thanks</button>
        </div>
      </div>
      <div id="att-sum-${emailId}" style="display:none"></div>`;
    }
    content.innerHTML=html;
    panel.dataset.loaded='1';
  }catch(e){
    content.innerHTML='<span style="color:#b91c1c;font-size:12px">Could not summarize.</span>';
  }
}

async function summarizeAttachments(emailId){
  const btn=document.getElementById('att-sum-btn-'+emailId);
  const container=document.getElementById('att-sum-'+emailId);
  if(!container)return;
  if(btn){btn.disabled=true;btn.textContent='Reading…';}
  container.style.display='block';
  container.innerHTML='<span style="color:var(--muted);font-style:italic;font-size:11px">Reading attachments…</span>';
  try{
    const token=await getAccessToken();
    const atts=await fetchVisibleAttachments(emailId);
    const textAtts=[];
    const unsupported=[];
    const SUPPORTED_TYPES=['text/plain','text/html','text/csv','application/json','application/xml','application/pdf','application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
    for(const a of atts){
      const ct=(a.contentType||'').toLowerCase();
      const isSupported=SUPPORTED_TYPES.some(t=>ct.startsWith(t)||ct===t);
      if(!isSupported){unsupported.push(a.name);continue;}
      try{
        const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments/${a.id}`,{headers:{Authorization:`Bearer ${token}`}});
        if(!r.ok)continue;
        const d=await r.json();
        if(!d.contentBytes)continue;
        let text='';
        const raw=atob(d.contentBytes);
        if(ct.includes('pdf')){
          // Extract text from BT...ET blocks (basic PDF text extraction)
          const btMatches=raw.match(/BT[\s\S]*?ET/g)||[];
          text=btMatches.map(block=>block.replace(/BT|ET/g,'').replace(/\(([^)]+)\)/g,'$1').replace(/[\r\n]+/g,' ').replace(/Tf|Td|TD|TJ|Tj|Tm|Ts|Tc|Tw|T\*|\d+(\.\d+)?\s/g,' ')).join(' ');
          if(!text.trim())text='[PDF content could not be extracted as text]';
        }else if(ct.includes('html')){
          text=raw.replace(/<[^>]+>/g,' ');
        }else{
          text=raw;
        }
        text=text.replace(/\s+/g,' ').trim();
        textAtts.push({name:a.name,text:text.slice(0,800)});
      }catch{}
    }
    if(!textAtts.length){
      const unsuppMsg=unsupported.length?` (${unsupported.map(n=>`"${n}"`).join(', ')} — unsupported type)`:'';
      container.innerHTML=`<div style="font-size:11px;color:#6b7280;padding:6px 0">Cannot read attachment content automatically${unsuppMsg}.</div>`;
      if(btn){btn.disabled=false;btn.textContent='Summarize attachments';}
      return;
    }
    const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL||'').replace(/\/?$/,'');
    const token2=await getAccessToken();
    const email=emailCache[emailId];
    const res=await fetch(fnUrl+'/attachment-summary',{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':`Bearer ${token2}`},
      body:JSON.stringify({subject:email?.subject||'',attachmentContents:textAtts})
    });
    if(!res.ok)throw new Error('Failed');
    const data=await res.json();
    const summary=data.summary||'No summary available.';
    const unsuppHtml=unsupported.length?`<div style="margin-top:6px;font-size:10px;color:#9ca3af">${unsupported.map(n=>`Cannot summarise "${escapeHtml(n)}" — unsupported file type`).join('<br>')}</div>`:'';
    container.innerHTML=`<div style="margin-top:6px;padding:8px 10px;background:#f9fafb;border-radius:4px;border:1px solid #e5e7eb">`+
      `<div style="font-size:10px;font-weight:700;color:#6b7280;margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px">Attachment Summary</div>`+
      formatSummaryHTML(summary)+unsuppHtml+'</div>';
    if(btn){btn.disabled=false;btn.textContent='Re-summarize';}
  }catch(e){
    container.innerHTML='<span style="color:#b91c1c;font-size:11px">Could not summarize attachments.</span>';
    if(btn){btn.disabled=false;btn.textContent='Summarize attachments';}
  }
}

// ============================================================
// TASK ASSIGNMENT NOTIFICATION
// ============================================================
function graphDueDate(date){
  const raw=String(date||'').trim();
  const m=raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(!m)return null;
  return {dateTime:`${m[1]}-${m[2]}-${m[3]}T17:00:00.0000000`,timeZone:'Central Standard Time'};
}

async function createToDoTask(task){
  const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  const attachmentNote=Array.isArray(task.attachmentNames)&&task.attachmentNames.length
    ? `\n\nAttachments from email:\n${task.attachmentNames.map(n=>`- ${n}`).join('\n')}`
    : '';
  const dueDate=task.deadline||((task.deadline===undefined)?task.date:'');
  const assignerEmail=currentUser?.email||currentAccount?.username||'';
  const assignerName=currentUser?.name||currentAccount?.name||assignerEmail||'DPEG Manager';
  task.assignedByName=task.assignedByName||assignerName;
  task.assignedByEmail=task.assignedByEmail||assignerEmail;
  try{
    await ensureTaskProofFolder(task);
  }catch(err){
    console.warn('Proof folder creation failed:',err.message);
  }

  // ── 1. Current user's To Do via Tasks.ReadWrite (now in SCOPES_GRAPH) ──
  try{
    const token=await getAccessToken();
    const due=graphDueDate(dueDate);
    const cleanSummary=(task.summary||'').replace(/[•*▾▲◆]/g,'').slice(0,4000);
    const deadlineLine=dueDate?`\nDeadline: ${dueDate}`:'';
    const taskBody={
      title:`[${task.person}] ${task.title||'Task'}`,
      body:{content:`Assigned to: ${task.person}\nEmail: ${task.email||''}\nDept: ${task.dept||''}\nPriority: ${task.priority||'Normal'}${deadlineLine}\n\n${cleanSummary}${attachmentNote}`,contentType:'text'},
      importance:task.priority==='High'?'high':'normal',
      status:'notStarted',
    };
    if(due)taskBody.dueDateTime=due;
    const listsRes=await fetch('https://graph.microsoft.com/v1.0/me/todo/lists',{headers:{Authorization:`Bearer ${token}`}});
    if(listsRes.ok){
      const ld=await listsRes.json();
      let listId=null;
      const existing=(ld.value||[]).find(l=>l.displayName==='DPEG Assigned Tasks');
      if(existing){listId=existing.id;}
      else{
        const cr=await fetch('https://graph.microsoft.com/v1.0/me/todo/lists',{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({displayName:'DPEG Assigned Tasks'})});
        if(cr.ok){const cd=await cr.json();listId=cd.id;}
      }
      if(listId){
        const taskRes=await fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks`,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify(taskBody)});
        if(taskRes.ok){const td=await taskRes.json();task.todoListId=listId;task.todoTaskId=td.id;}
      }
    }
  }catch(err){console.warn('Current user To Do failed:',err.message);}

  // ── 2. Recipient's To Do via Cloudflare Worker app credentials ──
  if(!task.email||!task.email.includes('@dhananipeg.com')||!fnUrl)return;
  try{
    const userToken=await getAccessToken();
    const todoRes=await fetch(`${fnUrl}/todo`,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':`Bearer ${userToken}`},
      body:JSON.stringify({
        recipientEmail:task.email,
        title:task.title||'Task',
        summary:[(task.summary||'').replace(/[•*▾▲◆]/g,'').slice(0,4000),attachmentNote.trim()].filter(Boolean).join('\n\n'),
        priority:task.priority||'Normal',
        date:dueDate||'',
        deadline:dueDate||'',
        assignedByName:task.assignedByName,
        assignedByEmail:task.assignedByEmail,
        appTaskId:String(task.id||''),
        proofShareUrl:task.proofShareUrl||'',
        proofInstructions:task.proofInstructions||'',
        proofBaseUrl:location.origin+location.pathname,
      }),
    });
    if(!todoRes.ok){
      const err=await todoRes.text().catch(()=>'');
      throw new Error(`Recipient To Do failed (${todoRes.status}) ${err}`);
    }
    const todoData=await todoRes.json().catch(()=>({}));
    if(todoData.listId)task.recipientTodoListId=todoData.listId;
    if(todoData.taskId)task.recipientTodoTaskId=todoData.taskId;
  }catch(err){
    console.warn('Recipient To Do (Worker) failed:',err.message);
  }

  // ── 3. Shared assignment record (D1) for the in-app Tasks tab ──
  recordAssignment(task);
}

// Send a message via draft (POST /me/messages → send). Returns {messageId, conversationId}.
// Requires Mail.ReadWrite scope (getDraftAccessToken).
async function sendMailAndGetId(token, message){
  const draftRes=await fetch('https://graph.microsoft.com/v1.0/me/messages',{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(message),
  });
  if(!draftRes.ok)throw new Error(`Draft create failed (${draftRes.status})`);
  const draft=await draftRes.json();
  const sendRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`},
  });
  if(!sendRes.ok)throw new Error(`Draft send failed (${sendRes.status})`);
  return {messageId:draft.id,conversationId:draft.conversationId};
}

// Reply to the task's original notification email in the same thread.
// Uses conversationId to find the sent message and createReply so Exchange sets proper
// In-Reply-To / References headers. Falls back to a new standalone email.
async function replyToTaskEmail(task, htmlContent, addr){
  const address=String(addr||task.email||'').trim().toLowerCase();
  if(!address.includes('@'))return;
  const token=await getDraftAccessToken();
  const convId=task.notifConversationId;
  if(convId){
    try{
      // Find any sent message in this conversation — draft id is stale after send, conversationId persists
      const filterParam=`conversationId eq '${convId}'`;
      const searchRes=await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$filter=${encodeURIComponent(filterParam)}&$top=1&$select=id`,
        {headers:{Authorization:`Bearer ${token}`}}
      );
      if(searchRes.ok){
        const searchData=await searchRes.json();
        const msgId=searchData.value?.[0]?.id;
        if(msgId){
          // createReply creates a draft with proper In-Reply-To / References headers set by Exchange
          const replyDraftRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${msgId}/createReply`,{
            method:'POST',
            headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
            body:JSON.stringify({message:{toRecipients:[{emailAddress:{address}}]}}),
          });
          if(replyDraftRes.ok){
            const rd=await replyDraftRes.json();
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${rd.id}`,{
              method:'PATCH',
              headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
              body:JSON.stringify({body:{contentType:'HTML',content:htmlContent}}),
            });
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${rd.id}/send`,{
              method:'POST',
              headers:{Authorization:`Bearer ${token}`},
            });
            return;
          }
        }
      }
    }catch(err){
      console.warn('Thread reply failed, falling back to standalone:',err.message);
    }
  }
  // Fallback: standalone email for old tasks without conversationId
  await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({message:{subject:`Re: Task Assigned: ${task.title||'(no subject)'}`,body:{contentType:'HTML',content:htmlContent},toRecipients:[{emailAddress:{address}}]},saveToSentItems:true}),
  });
}

// A proof is submitted from the assignee's account, which normally does not
// have the assignor's saved Sent Items conversation id. Locate the original
// assignment notification in the assignee's inbox and reply to that message
// so Outlook keeps this important alert in the same task conversation.
async function sendProofSubmittedEmail(params,proofs,note){
  const assigner=String(params?.assignedByEmail||'').trim().toLowerCase();
  const title=String(params?.title||'Task').trim();
  if(!assigner.includes('@'))return;
  const token=await getDraftAccessToken();
  const expectedSubject=`Task Assigned: ${title}`;
  let originalId='';
  try{
    const res=await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100&$select=id,subject,from,receivedDateTime&$orderby=receivedDateTime%20desc',{
      headers:{Authorization:`Bearer ${token}`},
    });
    if(res.ok){
      const data=await res.json();
      const match=(data.value||[]).find(message=>
        String(message.subject||'').trim()===expectedSubject&&
        String(message.from?.emailAddress?.address||'').trim().toLowerCase()===assigner
      );
      originalId=match?.id||'';
    }
  }catch(err){console.warn('Original assignment email lookup failed:',err.message);}
  const attachmentNames=(proofs||[]).map(file=>file?.name).filter(Boolean);
  const html=`<div style="font-family:Arial,sans-serif;max-width:620px;color:#111">
    <div style="background:#0E3416;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;font-size:13px;font-weight:700">Proof submitted — ${escapeHtml(title)}</div>
    <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">
      <p style="margin:0 0 8px"><strong>Submitted by:</strong> ${escapeHtml(currentUser?.name||currentUser?.email||'Assignee')}</p>
      ${note?`<p style="margin:0 0 8px"><strong>Note:</strong> ${escapeHtml(note)}</p>`:''}
      ${attachmentNames.length?`<p style="margin:0 0 8px"><strong>Files:</strong> ${attachmentNames.map(escapeHtml).join(', ')}</p>`:''}
      <p style="margin:10px 0 0">Open DPEG Task Manager and select <strong>Delegated → View Proof</strong> to review it.</p>
    </div>
  </div>`;
  if(originalId){
    try{
      const draftRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${originalId}/createReply`,{
        method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({message:{toRecipients:[{emailAddress:{address:assigner}}]}}),
      });
      if(draftRes.ok){
        const draft=await draftRes.json();
        const patchRes=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}`,{
          method:'PATCH',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
          body:JSON.stringify({body:{contentType:'HTML',content:html}}),
        });
        const sendRes=patchRes.ok?await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`,{method:'POST',headers:{Authorization:`Bearer ${token}`}}):null;
        if(sendRes?.ok)return;
      }
    }catch(err){console.warn('Proof thread reply failed:',err.message);}
  }
  await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
    method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({message:{subject:`Proof submitted: ${title}`,body:{contentType:'HTML',content:html},toRecipients:[{emailAddress:{address:assigner}}]},saveToSentItems:true}),
  });
}

async function sendTaskNotification(task){
  if(!task.email)return;
  const addr=(task.email||'').trim().toLowerCase();
  if(!addr||!addr.includes('@'))return;
  if(!addr.includes('@dhananipeg.com')&&!addr.includes('@dpeg'))return;
  try{
    const token=await getDraftAccessToken();
    const assignedBy=currentUser?.name||currentUser?.email||'DPEG Manager';
    const dueDate=task.deadline||((task.deadline===undefined)?task.date:'');
    const summaryClean=(task.summary||'').replace(/[•*▾▲◆]/g,'').trim().slice(0,600);
    const attHtml=Array.isArray(task.attachmentNames)&&task.attachmentNames.length
      ?`<p><strong>Attachments:</strong> ${task.attachmentNames.map(n=>escapeHtml(n)).join(', ')}</p>`:'';
    const htmlBody=`<div style="font-family:Arial,sans-serif;max-width:600px;color:#111">
      <div style="background:#0E3416;color:#fff;padding:14px 18px;border-radius:6px 6px 0 0">
        <div style="font-size:15px;font-weight:700">You have been assigned a task</div>
        <div style="font-size:12px;opacity:.8;margin-top:2px">DPEG Task Manager</div>
      </div>
      <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:16px 18px">
        <p style="margin:0 0 10px"><strong>Task:</strong> ${escapeHtml(task.title||'')}</p>
        ${summaryClean?`<p style="margin:0 0 10px"><strong>Summary:</strong> ${escapeHtml(summaryClean)}</p>`:''}
        <p style="margin:0 0 6px"><strong>Assigned by:</strong> ${escapeHtml(assignedBy)}</p>
        ${dueDate?`<p style="margin:0 0 6px"><strong>Deadline:</strong> ${escapeHtml(dueDate)}</p>`:''}
        <p style="margin:0 0 6px"><strong>Priority:</strong> ${escapeHtml(task.priority||'Normal')}</p>
        ${attHtml}
        <hr style="border:none;border-top:1px solid #e5e7eb;margin:12px 0">
        <p style="color:#9ca3af;font-size:11px;margin:0">Task messages and updates will appear as replies to this email.</p>
      </div>
    </div>`;
    const subject=`Task Assigned: ${task.title||'(no subject)'}`;
    const msg={subject,body:{contentType:'HTML',content:htmlBody},toRecipients:[{emailAddress:{address:addr}}]};
    const {conversationId}=await sendMailAndGetId(token,msg);
    task.notifConversationId=conversationId;
  }catch(err){
    console.warn('Task notification failed:',err.message);
  }
}

async function updateTodoTask(task,changes){
  const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  const token=await getAccessToken();

  // 1. Update current user's To Do task
  if(task.todoListId&&task.todoTaskId){
    try{
      const cleanSummary=(task.summary||'').replace(/[•*▾▲◆]/g,'').slice(0,4000);
      const due=graphDueDate(task.date||'');
      const bodyLines=[
        `Assigned to: ${task.person}`,
        task.email?`Email: ${task.email}`:'',
        `Dept: ${task.dept||''}`,
        `Priority: ${task.priority||'Normal'}`,
        task.date?`Date: ${task.date}`:'',
        cleanSummary?`\n${cleanSummary}`:'',
        task.followup?`\nFollow-up: ${task.followup}`:'',
        changes.length?`\nUpdated: ${changes.join(' | ')}`:'',
      ].filter(Boolean).join('\n');
      const patch={
        title:`[${task.person}] ${task.title||'Task'}`,
        importance:task.priority==='High'?'high':'normal',
        body:{content:bodyLines,contentType:'text'},
      };
      if(due)patch.dueDateTime=due;
      await fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${task.todoListId}/tasks/${task.todoTaskId}`,{
        method:'PATCH',
        headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify(patch),
      });
    }catch(err){console.warn('Update own To Do failed:',err.message);}
  }

  // 2. Update recipient's To Do via Worker (needs app credentials)
  if(task.recipientTodoListId&&task.recipientTodoTaskId&&task.email&&fnUrl){
    try{
      await fetch(`${fnUrl}/todo-update`,{
        method:'POST',
        headers:{'Content-Type':'application/json','Authorization':`Bearer ${token}`},
        body:JSON.stringify({
          recipientEmail:task.email,
          todoListId:task.recipientTodoListId,
          todoTaskId:task.recipientTodoTaskId,
          title:task.title||'Task',
          priority:task.priority||'Normal',
          date:task.date||'',
          followupNote:task.followup||'',
          changes,
          assignedByName:task.assignedByName||currentUser?.name||'',
        }),
      });
    }catch(err){console.warn('Update recipient To Do failed:',err.message);}
  }
}

async function sendTaskUpdateNotification(task,changes){
  if(!task.email)return;
  const addr=(task.email||'').trim().toLowerCase();
  if(!addr||!addr.includes('@'))return;
  if(!addr.includes('@dhananipeg.com')&&!addr.includes('@dpeg'))return;
  try{
    const updatedBy=currentUser?.name||currentUser?.email||'DPEG Manager';
    const changesHtml=changes.map(c=>`<li style="margin-bottom:4px;font-size:13px">${escapeHtml(c)}</li>`).join('');
    const htmlBody=`<div style="font-family:Arial,sans-serif;max-width:600px;color:#111">
      <div style="background:#1e40af;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;font-size:13px;font-weight:700">Task Update — ${escapeHtml(task.title||'')}</div>
      <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">
        <p style="margin:0 0 8px"><strong>Updated by:</strong> ${escapeHtml(updatedBy)}</p>
        ${changesHtml?`<p style="margin:0 0 4px"><strong>Changes:</strong></p><ul style="margin:0 0 10px;padding-left:20px">${changesHtml}</ul>`:''}
        ${task.followup?`<div style="background:#f0fdf4;border-left:3px solid #16a34a;padding:8px 12px;margin-top:8px"><strong>Note from manager:</strong><br>${escapeHtml(task.followup)}</div>`:''}
      </div>
    </div>`;
    await replyToTaskEmail(task,htmlBody,addr);
  }catch(err){
    console.warn('Task update notification failed:',err.message);
  }
}

// ============================================================
// REMOVE TASK
// ============================================================
// Shared by cancelActionTask and bulk delete: cancels the D1 assignment
// linked to this task (if any), so a task that's cancelled or deleted from
// the Action Log doesn't leave an orphaned row in the recipient's Tasks tab.
// Returns true if there was nothing to cancel, cancellation succeeded, or
// the assignment was already in a terminal state (409) — false only on a
// genuine failure to reach/update the Worker.
async function cancelLinkedAssignment(task){
  if(!task.assignmentId){
    await window.renderMyTasks?.(true);
    task.assignmentId=window.findDelegatedAssignmentByAppTaskId?.(task.id)?.id||'';
  }
  if(!task.assignmentId){task.cancelledAt=new Date().toISOString();return true;}
  try{
    const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL||'').replace(/\/?$/,'');
    const token=await getAccessToken();
    const res=await fetch(`${fnUrl}/assignment-cancel`,{method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${token}`},body:JSON.stringify({id:task.assignmentId,reason:''})});
    if(res.ok){
      const data=await res.json().catch(()=>({}));
      task.cancelledAt=data.cancelledAt||new Date().toISOString();
      return true;
    }
    // Already Done or already Cancelled server-side — nothing left to clean up.
    if(res.status===409){task.cancelledAt=new Date().toISOString();return true;}
    return false;
  }catch{return false;}
}

async function cancelActionTask(id){
  const task=tasks.find(t=>t.id===id);if(!task)return;
  if(!confirm(`Cancel "${task.title||'this task'}"?\n\nIt will be removed from active tasks and moved to Cancelled History.`))return;
  try{
    if(!await cancelLinkedAssignment(task))throw new Error('Could not cancel the linked assignment');
    task.status='Cancelled';task.cancelReason='';
    closeMo('mo-detail');syncBadges();refreshAll();await saveTasksToOneDrive();
    await window.sendTaskCancelledEmail?.({
      appTaskId:String(task.id||''),title:task.title||'Task',recipientEmail:task.email||'',cancelReason:'',
    });
    await window.renderMyTasks?.(false);
    toast('Task cancelled and moved to Cancelled History');
  }catch(err){toast('Could not cancel task: '+(err.message||err));}
}

async function removeTask(id){
  if(!confirm('Remove this task from the Action Log?'))return;
  tasks=tasks.filter(t=>t.id!==id);
  selectedTaskIds.delete(id);
  buildTrackedSet();
  closeMo('mo-detail');
  syncBadges();refreshAll();
  toast('Task removed from Action Log');
  await saveTasksToOneDrive();
}

async function removeTaskByEmail(emailId){
  const email=emailCache[emailId];if(!email)return;
  const idx=tasks.findIndex(t=>t.conversationId===email.conversationId||t.emailId===emailId||t.lastMessageId===emailId);
  if(idx<0){toast('Task not found in Action Log');return;}
  if(!confirm('Remove this task from the Action Log?'))return;
  tasks.splice(idx,1);
  trackedEmailIds.delete(emailId);
  if(email.conversationId)trackedEmailIds.delete(email.conversationId);
  syncBadges();refreshAll();
  toast('Task removed from Action Log');
  await saveTasksToOneDrive();
  readEmail(emailId,currentFolder);
}

async function addTaskFromEmail(emailId){
  const email=emailCache[emailId];if(!email)return;
  const sender=email.from?.emailAddress;
  const p=findPersonByEmail(sender?.address);
  const person=p?.name||sender?.name||sender?.address||'Unknown';
  const address=sender?.address||'';
  const emailDate=email.receivedDateTime||email.sentDateTime||'';
  const assignedDate=new Date().toISOString();
  await attachFreeThreadSummary(email,{assignedDate,emailDate});
  const result=upsertTaskFromEmail(email,person,address,personDept(address,person),assignedDate);
  trackedEmailIds.add(emailId);
  refreshAll();
  if(result.created){
    await sendTaskNotification(result.task).catch(()=>{});
    await createToDoTask(result.task).catch(()=>{});
  }
  await saveTasksToOneDrive();
  toast(result.created?'Task added to Action Log':'Updated existing thread in Action Log');
  readEmail(emailId,currentFolder);
}

async function updateTrackedThreadFromEmail(emailId){
  const email=emailCache[emailId];if(!email)return;
  const sender=email.from?.emailAddress;
  const p=findPersonByEmail(sender?.address);
  const person=p?.name||sender?.name||sender?.address||'Unknown';
  const address=sender?.address||'';
  await attachFreeThreadSummary(email);
  const result=upsertTaskFromEmail(email,person,address,personDept(address,person),email.receivedDateTime||email.sentDateTime||new Date().toISOString());
  trackedEmailIds.add(emailId);
  await saveTasksToOneDrive();
  refreshAll();
  toast(result.created?'Task added to Action Log':'Thread summary updated');
  readEmail(emailId,currentFolder);
}

function refreshFolder(){
  if(currentFolder==='schedule'||currentFolder==='calendar'){loadScheduleFolder();return;}
  loadFolder(currentFolder);
}
