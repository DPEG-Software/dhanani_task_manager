// ============================================================
// NOTIFICATIONS SYSTEM
// ============================================================
let notifications=[];
let highlightedNotifId=null;

// "Seen" used to live only as notifications[].seen, persisted through an
// async OneDrive save shared with tasks/staffConfig/etc — a slow round-trip
// that could lose a just-set seen:true if the page was refreshed before it
// completed (and old data that predates seen tracking entirely never had
// the field set, so it always read as unseen). Tracked locally instead,
// same instant/race-free localStorage pattern as the Tasks tab follow-up
// read-state. Keyed by kvNotifId (stable across polls) falling back to the
// local id, and only pending notifications ever count as "new" — once
// approved/dismissed it's resolved regardless of whether it was "seen".
const NOTIF_SEEN_STORAGE_KEY='dpeg_notif_seen_ids';
function loadNotifSeenIds(){
  try{ return new Set(JSON.parse(localStorage.getItem(NOTIF_SEEN_STORAGE_KEY)||'[]')); }
  catch{ return new Set(); }
}
function saveNotifSeenIds(set){
  try{ localStorage.setItem(NOTIF_SEEN_STORAGE_KEY,JSON.stringify([...set].slice(-1000))); }catch{}
}
let notifSeenIds=loadNotifSeenIds();
function notifSeenKey(n){ return String(n.kvNotifId||n.id); }
function isNotifUnseen(n){ return n.status==='pending' && !notifSeenIds.has(notifSeenKey(n)); }
let processedEmailIds=new Set(JSON.parse(localStorage.getItem('dpeg_processed_done_emails')||'[]'));
let _notifPollTimer=null;
let _notifPollCycle=0;
let pendingProofFiles=[];
let pendingFollowupFiles=[];
let _taskMessagePollInitialized=false;
let _taskMessagePollCursor='';

function renderPendingProofFileList(){
  const list=document.getElementById('proof-file-list');
  if(!list)return;
  if(!pendingProofFiles.length){list.innerHTML='';return;}
  list.innerHTML=pendingProofFiles.map((f,i)=>`<span style="display:inline-flex;align-items:center;gap:4px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:4px;padding:2px 6px;font-size:10.5px;color:#166534;white-space:nowrap;max-width:160px">
    <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(f.name)}</span>
    <button type="button" onclick="pendingProofFiles.splice(${i},1);renderPendingProofFileList()" style="border:none;background:none;color:#dc2626;font-size:12px;font-weight:700;cursor:pointer;padding:0;line-height:1;flex-shrink:0" title="Remove">&#10005;</button>
  </span>`).join(' ');
}

// ── Proof notification system ─────────────────────────────────────────────────
async function checkAndLoadProofNotifications(){
  const fnUrl=workerBaseUrl();
  if(!fnUrl||!currentUser?.email)return;
  const userToken=await getAccessToken();
  const messageQuery=_taskMessagePollInitialized&&_taskMessagePollCursor
    ?`messagesSince=${encodeURIComponent(_taskMessagePollCursor)}`
    :'includeMessages=1';
  const res=await fetch(`${fnUrl}/notify?${messageQuery}`,{headers:{Authorization:`Bearer ${userToken}`}});
  if(!res.ok)return;
  const data=await res.json();
  _taskMessagePollInitialized=true;
  if(data.messageCursor)_taskMessagePollCursor=String(data.messageCursor);
  const kvNotifs=Array.isArray(data.notifications)?data.notifications:[];
  const myEmail=(currentUser?.email||'').toLowerCase();
  tasks.forEach(t=>delete t._proofNotif);

  // Keep every submission for each task. The review queue still uses only
  // the latest pending submission, but earlier attempts remain available in
  // Changes Requested and completed History views.
  const proofHistory={};
  kvNotifs.filter(n=>n.type==='proof_submitted'&&(
      String(n.senderEmail||'').toLowerCase()===myEmail||
      String(n.recipientEmail||'').toLowerCase()===myEmail||
      window.isDelegatedTaskProof?.(n.appTaskId,n.recipientEmail)
    )).forEach(n=>{
      const key=String(n.appTaskId||'');
      if(!key)return;
      (proofHistory[key]||(proofHistory[key]=[])).push(n);
    });
  Object.values(proofHistory).forEach(list=>list.sort((a,b)=>new Date(b.submittedAt||b.createdAt||0)-new Date(a.submittedAt||a.createdAt||0)));
  window._proofSubmissionHistory=proofHistory;

  // Keep the latest approval/change-request result available to My Tasks.
  // The assignment row carries the status, but the KV result is where the
  // assignor's reason lives. Without this bridge the assignee only sees
  // "declined" and has to discover the explanation in email.
  const proofResults={};
  kvNotifs.filter(n=>n.type==='proof_result'&&String(n.recipientEmail||'').toLowerCase()===myEmail)
    .forEach(n=>{
      const key=String(n.appTaskId||'');
      if(!key)return;
      const current=proofResults[key];
      const nextTime=new Date(n.createdAt||n.updatedAt||0).getTime();
      const currentTime=current?new Date(current.createdAt||current.updatedAt||0).getTime():0;
      if(!current||nextTime>=currentTime)proofResults[key]=n;
    });
  const proofResultSig=Object.keys(proofResults).sort().map(k=>{
    const n=proofResults[k];return `${k}:${n.result}:${n.reason||''}:${n.createdAt||''}`;
  }).join('|');
  window._proofResultState=proofResults;
  if(proofResultSig!==window._proofResultStateSig){
    window._proofResultStateSig=proofResultSig;
    window.renderTasksTabList?.();
  }

  // Track the latest follow-up thread per (appTaskId, recipientEmail) that
  // involves the current user (as assignor or assignee), so the Tasks tab
  // can show an unread dot on the Follow-up button without a separate fetch.
  window._taskFollowupState=window._taskFollowupState||{};
  kvNotifs.filter(n=>(n.type==='proof_submitted'||n.type==='task_followup')
      &&Array.isArray(n.thread)&&n.thread.length
      &&(String(n.senderEmail||'').toLowerCase()===myEmail||String(n.recipientEmail||'').toLowerCase()===myEmail))
    .forEach(n=>{
      const key=`${n.appTaskId}::${String(n.recipientEmail||'').toLowerCase()}`;
      const existing=window._taskFollowupState[key];
      if(isBetterFollowupCandidate(n,existing)){
        let thread=n.thread;
        if(n.type==='task_followup'&&existing&&Array.isArray(existing.thread)){
          const byId=new Map();
          [...existing.thread,...n.thread].forEach(item=>byId.set(String(item.id||''),item));
          thread=[...byId.values()].sort((a,b)=>new Date(a.createdAt||0)-new Date(b.createdAt||0)||String(a.id||'').localeCompare(String(b.id||'')));
        }
        window._taskFollowupState[key]={notifId:n.id,appTaskId:n.appTaskId,recipientEmail:n.recipientEmail,thread,type:n.type,updatedAt:n.updatedAt||n.submittedAt||n.createdAt||0};
      }
    });
  // renderMyTasks(true) below only redraws the Tasks tab when the D1
  // assignment row itself changed (status/proof) — a new follow-up message
  // only touches this KV thread state, so without this it would silently
  // never repaint until the user manually leaves and reopens the tab.
  const followupSig=Object.keys(window._taskFollowupState).sort()
    .map(k=>{const s=window._taskFollowupState[k];return `${k}:${s.thread.length}:${s.updatedAt}`;}).join('|');
  if(followupSig!==window._taskFollowupStateSig){
    window._taskFollowupStateSig=followupSig;
    window.renderTasksTabList?.();
  }

  const existingKvIds=new Set((notifications||[]).filter(n=>n.kvNotifId).map(n=>n.kvNotifId));
  let newAdded=0;
  let existingUpdated=0;

  // Deduplicate by appTaskId — only the latest proof_submitted per task matters
  const latestByTask=new Map();
  kvNotifs.filter(n=>n.type==='proof_submitted'&&n.status==='pending'&&(
      String(n.senderEmail||'').toLowerCase()===myEmail||
      window.isDelegatedTaskProof?.(n.appTaskId,n.recipientEmail)
    ))
    .forEach(pn=>{
      const key=String(pn.appTaskId);
      const existing=latestByTask.get(key);
      const pnTime=new Date(pn.updatedAt||pn.submittedAt||0).getTime();
      const exTime=existing?new Date(existing.updatedAt||existing.submittedAt||0).getTime():0;
      if(!existing||pnTime>=exTime)latestByTask.set(key,pn);
    });

  latestByTask.forEach((pn,taskKey)=>{
    // A proof submission for a task assigned purely through the Tasks Hub
    // (D1) doesn't always have a matching entry in this user's own Action
    // Log — that array is per-account OneDrive data, not shared — so a
    // missing local match must not stop the notification (and the Review
    // Proof modal it powers) from surfacing. Fall back to the appTaskId
    // itself when there's no local task to enrich the record with.
    const task=tasks.find(t=>String(t.id)===taskKey)||null;
    if(task)task._proofNotif=pn;
    const localData={
      type:'proof_submitted',
      message:`${pn.recipientName||pn.recipientEmail||'Someone'} submitted proof for "${pn.taskTitle||'a task'}"`,
      taskTitle:pn.taskTitle||task?.title||'',
      taskId:task?task.id:taskKey,
      proofs:Array.isArray(pn.proofs)?pn.proofs:[],
      note:pn.note||'',
      thread:Array.isArray(pn.thread)?pn.thread:[],
      followupStatus:pn.followupStatus||'',
      submittedAt:pn.submittedAt||'',
      timestamp:pn.updatedAt||pn.submittedAt||new Date().toISOString(),
      status:'pending',
      recipientEmail:pn.recipientEmail||'',
      recipientName:pn.recipientName||'',
      assignmentGroupId:task?.assignmentGroupId||'',
    };
    // Remove any stale local notifications for this same task before adding/updating
    const staleIndices=notifications.reduce((acc,n,i)=>{
      if(n.type==='proof_submitted'&&String(n.taskId)===taskKey&&n.kvNotifId!==pn.id)acc.push(i);
      return acc;
    },[]);
    for(let i=staleIndices.length-1;i>=0;i--)notifications.splice(staleIndices[i],1);

    const existingLocal=notifications.find(n=>n.kvNotifId===pn.id);
    if(existingLocal){
      // Only count as changed if meaningful fields actually differ — prevents refreshAll() every poll
      const sig=n=>[n.proofs,n.note,n.thread,n.followupStatus,n.status].map(v=>JSON.stringify(v)).join('|');
      const before=sig(existingLocal);
      Object.assign(existingLocal,localData);
      if(sig(existingLocal)!==before)existingUpdated++;
    }else{
      // seen must only be set here, at creation — never inside localData
      // (which also gets merged into already-existing entries on every poll
      // via Object.assign below), or an already-viewed notification would
      // flip back to "unseen" the moment any other field on it changed.
      notifications.unshift({
        id:Date.now()+Math.random(),
        kvNotifId:pn.id,
        seen:false,
        ...localData,
      });
      newAdded++;
    }
  });

  if(newAdded>0||existingUpdated>0){
    updateNotifBadge();
    renderNotifications();
    window.updateNotificationCenter?.();
    await saveNotifications();
    refreshAll();
    if(newAdded>0)toast(`${newAdded} proof submission${newAdded>1?'s':''} ready for review in Assigned by Me`);
  }
}

// Opens the same focused review surface used by Delegated Tasks so proof,
// conversation and approval stay in one workflow.
function goToProofInTasks(taskId){
  const task=tasks.find(x=>String(x.id)===String(taskId));
  closeMo('mo-detail');
  window.openTaskProofReview?.({
    appTaskId:String(taskId),
    title:task?.title||'Review submitted proof',
    recipientName:task?.person||'',
    recipientEmail:task?.email||'',
  });
}

function renderProofThread(n){
  const thread=Array.isArray(n?.thread)?n.thread:[];
  if(!thread.length)return '';
  const visible=[...thread];
  const threadKey=String(n?.id||n?.kvNotifId||n?.appTaskId||'thread').replace(/[^a-zA-Z0-9_-]/g,'_');
  return `<div style="margin:8px 0 10px;padding:8px 10px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:6px">
    <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px">
      <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:#6b7280">Follow-up thread</div>
      ${thread.length>2?`<div style="font-size:10px;color:#9ca3af">${thread.length} messages</div>`:''}
    </div>
    <div data-proof-thread-scroll="1" data-thread-key="${threadKey}" data-thread-count="${thread.length}" style="max-height:154px;overflow-y:auto;padding-right:3px">
    ${visible.map(item=>{
      const isAnswer=item.by==='assignee';
      const label=isAnswer?'Assignee reply':'Question';
      const color=isAnswer?'#14532d':'#1f2937';
      const bg=isAnswer?'#f0fdf4':'#fff';
      const border=isAnswer?'#bbf7d0':'#e5e7eb';
      const dt=item.createdAt?new Date(item.createdAt).toLocaleString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'}):'';
      const attachments=Array.isArray(item.attachments)?item.attachments:[];
      const attachmentHtml=attachments.length?`<div style="margin-top:6px;display:grid;gap:4px">
        ${attachments.map(a=>`<div style="display:flex;align-items:center;gap:7px;font-size:11.5px;background:#fff;border:1px solid #e5e7eb;border-radius:5px;padding:5px 7px;min-width:0">
          <span style="font-weight:700;color:#111;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(a.name||'Attachment')}</span>
          <span style="color:#9ca3af;flex-shrink:0">${a.size?escapeHtml(formatFileSize(a.size)):''}</span>
          ${a.webUrl?`<a href="${escapeHtml(a.webUrl)}" target="_blank" rel="noopener" style="color:#0E3416;font-weight:700;flex-shrink:0">Open</a>`:''}
        </div>`).join('')}
      </div>`:'';
      return `<div style="background:${bg};border:1px solid ${border};border-radius:5px;padding:7px 8px;margin-bottom:6px">
        <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:3px">
          <span style="font-size:10px;font-weight:800;color:${color};text-transform:uppercase;letter-spacing:.4px">${label}</span>
          <span style="font-size:10px;color:#9ca3af">${escapeHtml(item.name||item.email||'')}</span>
          ${dt?`<span style="font-size:10px;color:#9ca3af">${dt}</span>`:''}
        </div>
        <div style="font-size:12px;line-height:1.5;color:${color};white-space:pre-wrap">${escapeHtml(item.message||'')}</div>
        ${attachmentHtml}
      </div>`;
    }).join('')}
    </div>
  </div>`;
}

function scrollProofThreadsToBottom(root=document){
  window._proofThreadScrollState=window._proofThreadScrollState||{};
  setTimeout(()=>{
    (root||document).querySelectorAll?.('[data-proof-thread-scroll="1"]').forEach(el=>{
      const key=el.dataset.threadKey||'thread';
      const count=el.dataset.threadCount||'0';
      const stateKey=`${key}:${count}`;
      if(window._proofThreadScrollState[key]===stateKey)return;
      el.scrollTop=el.scrollHeight;
      window._proofThreadScrollState[key]=stateKey;
    });
  },0);
}

async function sendProofFollowupEmail(recipientEmail,task,question){
  const addr=String(recipientEmail||task?.email||'').trim();
  if(!addr||!addr.includes('@'))return;
  try{
    const link=buildProofSubmitLinkForTask(task,{recipientEmail:addr,taskTitle:task?.title||'Assigned task'});
    const sender=currentUser?.name||currentUser?.email||'DPEG Task Manager';
    const html=`<div style="font-family:Arial,sans-serif;max-width:620px;color:#111">
      <div style="background:#d97706;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;font-size:13px;font-weight:700">Follow-up Question — ${escapeHtml(task?.title||'Assigned task')}</div>
      <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">
        <p style="margin:0 0 8px"><strong>Asked by:</strong> ${escapeHtml(sender)}</p>
        <div style="background:#fffbeb;border-left:3px solid #d97706;padding:10px 12px;margin:10px 0;font-size:14px;line-height:1.5">${escapeHtml(question).replace(/\n/g,'<br>')}</div>
        <p style="margin:10px 0 4px"><a href="${escapeHtml(link)}" style="display:inline-block;padding:8px 16px;background:#0E3416;color:#fff;text-decoration:none;border-radius:4px;font-size:13px;font-weight:600">Open proof form to reply →</a></p>
        <p style="color:#9ca3af;font-size:11px;margin:8px 0 0">You can also reply via the link in your Microsoft To Do task.</p>
      </div>
    </div>`;
    await replyToTaskEmail(task,html,addr);
  }catch(err){
    console.warn('Follow-up email failed:',err.message);
  }
}

// ── Generic task follow-up (Tasks tab — Received & Delegated cards) ──────────
// Unlike askProofFollowup/askNotificationFollowup above, this does not require
// a proof to have been submitted first: either party can start or continue
// the conversation about a task at any time. New messages are separate D1
// rows; the Worker merges them with legacy KV history into the same thread
// shape so the UI stays visually and structurally consistent.
let _taskFollowupCtx=null;

// A dedicated task_followup record always wins over a proof_submitted one,
// even if the proof_submitted has a fresher timestamp — a proof resubmission
// creates a brand-new, empty proof_submitted record, and picking on recency
// alone would make that look like "the" thread and hide the real
// conversation. Recency only breaks ties within the same type. Mirrors the
// same rule enforced server-side in task_followup_message.
function isBetterFollowupCandidate(candidate,current){
  if(!current)return true;
  const candIsThread=candidate.type==='task_followup';
  const curIsThread=current.type==='task_followup';
  if(candIsThread!==curIsThread)return candIsThread;
  const candTime=new Date(candidate.updatedAt||candidate.submittedAt||candidate.createdAt||0).getTime();
  const curTime=new Date(current.updatedAt||current.submittedAt||current.createdAt||0).getTime();
  return candTime>=curTime;
}

async function fetchTaskFollowupThread(appTaskId,recipientEmail){
  try{
    const fnUrl=workerBaseUrl();
    const userToken=await getAccessToken();
    const query=new URLSearchParams({
      taskId:String(appTaskId||''),
      recipientEmail:String(recipientEmail||''),
    });
    const res=await fetch(`${fnUrl}/notify?${query}`,{headers:{Authorization:`Bearer ${userToken}`}});
    if(!res.ok)return null;
    const data=await res.json();
    const notifs=Array.isArray(data.notifications)?data.notifications:[];
    const email=String(recipientEmail||'').toLowerCase();
    return notifs
      .filter(n=>(n.type==='proof_submitted'||n.type==='task_followup')&&String(n.appTaskId)===String(appTaskId||'')&&String(n.recipientEmail||'').toLowerCase()===email)
      .reduce((best,n)=>isBetterFollowupCandidate(n,best)?n:best,null);
  }catch{return null;}
}

async function showTaskFollowupModal(params){
  _taskFollowupCtx=params;
  const titleEl=document.getElementById('tf-modal-title');
  const subEl=document.getElementById('tf-modal-sub');
  const sendBtn=document.getElementById('tf-modal-send');
  const otherName=params.role==='assignee'?(params.assignerName||params.assignerEmail):(params.recipientName||params.recipientEmail);
  if(titleEl)titleEl.textContent=params.requestChanges?'Ask for Changes':`Messages — ${otherName||'the other party'}`;
  if(subEl)subEl.textContent=params.requestChanges?`Tell ${otherName||'the assignee'} exactly what needs to be changed.`:(params.title||'');
  if(sendBtn)sendBtn.textContent=params.requestChanges?'Send Request':'Send';
  const input=document.getElementById('tf-modal-input');
  if(input){
    input.value=params.requestChanges?'Please make these changes: ':'';
    input.placeholder=params.requestChanges?'What needs to be changed?':'Type a message...';
  }
  const statusEl=document.getElementById('tf-modal-status');
  if(statusEl)statusEl.textContent='';
  document.getElementById('mo-task-followup')?.classList.add('open');
  await refreshTaskFollowupThread();
  if(params.requestChanges){input?.focus();input?.setSelectionRange(input.value.length,input.value.length);}
}

async function refreshTaskFollowupThread(){
  if(!_taskFollowupCtx)return;
  const body=document.getElementById('tf-modal-body');
  if(!body)return;
  body.innerHTML='<div style="font-size:12px;color:var(--muted)">Loading conversation...</div>';
  const n=await fetchTaskFollowupThread(_taskFollowupCtx.appTaskId,_taskFollowupCtx.recipientEmail);
  const thread=n&&Array.isArray(n.thread)?n.thread:[];
  body.innerHTML=thread.length?renderProofThread(n):'<div style="font-size:12px;color:var(--muted);padding:6px 0">No messages yet — send one to start the conversation.</div>';
  scrollProofThreadsToBottom(body);
  window.markTaskFollowupSeen?.(_taskFollowupCtx.assignmentId,thread.length,_taskFollowupCtx.role);
}

function closeTaskFollowup(){
  document.getElementById('mo-task-followup')?.classList.remove('open');
  _taskFollowupCtx=null;
}

async function sendTaskFollowupEmail(recipientEmail,ctx,message){
  const addr=String(recipientEmail||'').trim();
  if(!addr||!addr.includes('@'))return;
  try{
    const sender=currentUser?.name||currentUser?.email||'DPEG Task Manager';
    const appUrl=location.origin+location.pathname;
    const html=`<div style="font-family:Arial,sans-serif;max-width:620px;color:#111">
      <div style="background:#0E3416;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;font-size:13px;font-weight:700">Task Message — ${escapeHtml(ctx.title||'Assigned task')}</div>
      <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">
        <p style="margin:0 0 8px"><strong>From:</strong> ${escapeHtml(sender)}</p>
        <div style="background:#f0fdf4;border-left:3px solid #0E3416;padding:10px 12px;margin:10px 0;font-size:14px;line-height:1.5">${escapeHtml(message).replace(/\n/g,'<br>')}</div>
        <p style="color:#9ca3af;font-size:11px;margin:8px 0 0">Reply from the Tasks tab in DPEG Task Manager: <a href="${escapeHtml(appUrl)}">${escapeHtml(appUrl)}</a></p>
      </div>
    </div>`;
    await replyToTaskEmail({title:ctx.title||'Assigned task',email:addr},html,addr);
  }catch(err){
    console.warn('Task message email failed:',err.message);
  }
}

async function sendTaskFollowupMessage(){
  if(!_taskFollowupCtx)return;
  const ctx=_taskFollowupCtx;
  const input=document.getElementById('tf-modal-input');
  const statusEl=document.getElementById('tf-modal-status');
  const btn=document.getElementById('tf-modal-send');
  const message=String(input?.value||'').trim();
  if(!message){toast('Type a message first');return;}
  if(btn)btn.disabled=true;
  if(statusEl)statusEl.textContent='Sending...';
  try{
    const fnUrl=workerBaseUrl();
    const userToken=await getAccessToken();
    const res=await fetch(`${fnUrl}/notify`,{
      method:'POST',
      headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
      body:JSON.stringify({
        type:'task_followup_message',
        appTaskId:ctx.appTaskId||'',
        taskTitle:ctx.title||'',
        assignerEmail:ctx.assignerEmail||'',
        recipientEmail:ctx.recipientEmail||'',
        recipientName:ctx.recipientName||'',
        by:ctx.role,
        senderEmail:currentUser?.email||'',
        senderName:currentUser?.name||'',
        message,
      }),
    });
    if(!res.ok){const d=await res.text().catch(()=>'');throw new Error(d||'Could not send message');}
    if(input)input.value='';
    if(statusEl)statusEl.textContent='';
    if(ctx.requestChanges){
      await dismissNotification(ctx.notificationId,message);
      const proofNotice=notifications.find(n=>n.id===ctx.notificationId);
      if(proofNotice?.status!=='dismissed')throw new Error('The message was sent, but the proof status could not be updated');
      toast('Changes requested — the assignee was notified');
      closeTaskFollowup();
    }else{
      toast('Message sent — the recipient was notified in the app');
    }
    await refreshTaskFollowupThread();
    window.renderMyTasks?.(true);
  }catch(err){
    if(statusEl)statusEl.innerHTML=`<span style="color:#b91c1c">${escapeHtml(err.message||'Could not send message')}</span>`;
  }finally{if(btn)btn.disabled=false;}
}

async function sendProofDeclineEmail(recipientEmail,taskTitle,reason,senderName,task){
  const addr=String(recipientEmail||'').trim();
  if(!addr||!addr.includes('@'))return;
  try{
    const html=`<div style="font-family:Arial,sans-serif;max-width:600px;color:#111">
      <div style="background:#b91c1c;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;font-size:13px;font-weight:700">Changes Requested — Resubmission Required</div>
      <div style="border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">
        <p style="margin:0 0 8px"><strong>Task:</strong> ${escapeHtml(taskTitle||'Task')}</p>
        <p style="margin:0 0 8px"><strong>Changes requested by:</strong> ${escapeHtml(senderName||'The assignor')}</p>
        <div style="background:#fff1f2;border-left:3px solid #b91c1c;padding:10px 12px;margin:10px 0;font-size:14px;line-height:1.5"><strong>Reason:</strong> ${escapeHtml(reason||'No reason provided')}</div>
        <p style="margin:10px 0 0">Please review the feedback above and resubmit your proof using the link in your Microsoft To Do task.</p>
      </div>
    </div>`;
    if(task){
      await replyToTaskEmail(task,html,addr);
    }else{
      const token=await getDraftAccessToken();
      await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({message:{subject:`Changes requested: ${taskTitle||'Task'}`,body:{contentType:'HTML',content:html},toRecipients:[{emailAddress:{address:addr}}]},saveToSentItems:true})});
    }
  }catch(err){console.warn('Proof decline email failed:',err.message);}
}

async function loadNotifications(){
  try{
    notifications=Array.isArray(notifications)?notifications:[];
  }catch{notifications=[];}
  updateNotifBadge();
}

async function saveNotifications(){
  try{
    await saveTasksToOneDrive();
  }catch(err){console.warn('Notifications save failed:',err.message);}
}

async function addNotification(n){
  notifications.unshift(n);
  updateNotifBadge();
  await saveNotifications();
}

function updateNotifBadge(){
  const pending=notifications.filter(n=>n.status==='pending').length;
  const badge=document.getElementById('nb-notif');
  if(badge){
    badge.textContent=pending;
    badge.style.display=pending>0?'':'none';
  }
  window.renderTasksTabList?.();
  window.updateNotificationCenter?.();
}

function markNotifSeen(id){
  const n=notifications.find(x=>x.id===id);
  if(!n)return;
  const key=notifSeenKey(n);
  if(!notifSeenIds.has(key)){
    notifSeenIds.add(key);
    saveNotifSeenIds(notifSeenIds);
    n.seen=true;
    renderNotifications();
  }
}

let taskProofReviewContext=null;

function taskProofReviewNotification(appTaskId,pendingOnly=true){
  const local=notifications.find(n=>
    n.type==='proof_submitted'&&
    String(n.taskId)===String(appTaskId||'')&&
    (!pendingOnly||n.status==='pending')
  );
  if(local)return local;
  const remote=(window._proofSubmissionHistory?.[String(appTaskId||'')]||[])
    .find(n=>!pendingOnly||n.status==='pending');
  if(!remote)return null;
  return {
    ...remote,
    taskId:String(remote.appTaskId||appTaskId||''),
    kvNotifId:remote.id,
    timestamp:remote.updatedAt||remote.submittedAt||remote.createdAt||'',
    status:remote.status==='declined'?'dismissed':remote.status,
  };
}

window.hasPendingTaskProofReview=function hasPendingTaskProofReview(appTaskId){
  return !!taskProofReviewNotification(appTaskId,true);
};

window.openTaskProofReview=async function openTaskProofReview(assignment){
  const appTaskId=String(assignment?.appTaskId||assignment?.taskId||'');
  if(!appTaskId){toast('Proof submission could not be matched to this task');return;}
  const modal=document.getElementById('mo-task-proof-review');
  const body=document.getElementById('task-review-body');
  if(!modal||!body)return;
  taskProofReviewContext={assignment,appTaskId,notificationId:null};
  document.getElementById('task-review-title').textContent=assignment?.title||'Review submitted proof';
  document.getElementById('task-review-sub').textContent=`Submitted by ${assignment?.recipientName||assignment?.recipientEmail||'the assignee'}`;
  document.getElementById('task-review-status').textContent='Loading submission...';
  body.innerHTML='<div style="padding:28px;text-align:center;font-size:12px;color:var(--muted)">Loading proof and conversation...</div>';
  modal.classList.add('open');
  await checkAndLoadProofNotifications().catch(err=>console.warn('Proof notification refresh failed:',err.message));
  const n=taskProofReviewNotification(appTaskId,false);
  if(!n){
    body.innerHTML='<div class="empty-state"><div class="es-text">Submission not available yet</div><div class="es-sub">It may still be syncing. Close this window and try again shortly.</div></div>';
    document.getElementById('task-review-status').textContent='';
    setTaskReviewActionsEnabled(false);
    return;
  }
  taskProofReviewContext.notificationId=n.id;
  markNotifSeen(n.id);
  renderTaskProofReview(n);
};

function renderTaskProofReview(n){
  const body=document.getElementById('task-review-body');
  if(!body||!n)return;
  const pending=n.status==='pending';
  const stateLabel=n.status==='approved'?'Approved':n.status==='dismissed'?'Changes requested':'Awaiting your review';
  const stateClass=n.status==='approved'?'is-approved':n.status==='dismissed'?'is-changes':'is-pending';
  const task=tasks.find(t=>String(t.id)===String(n.taskId||''));
  const instructions=String(taskProofReviewContext?.assignment?.proofInstructions||task?.proofInstructions||'').trim();
  const instructionsHtml=instructions?`<div style="margin:10px 0;padding:10px 12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px"><div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:#166534;margin-bottom:4px">Proof instructions</div><div style="font-size:12px;color:#14532d;line-height:1.5;white-space:pre-wrap">${escapeHtml(instructions)}</div></div>`:'';
  body.innerHTML=`<div class="task-review-summary ${stateClass}">
      <div><div class="task-review-kicker">Proof submission</div><div class="task-review-state">${stateLabel}</div></div>
      <div class="task-review-person">${escapeHtml(n.recipientName||n.recipientEmail||'Assignee')}</div>
    </div>${instructionsHtml}${renderNotificationProofs(n)}`;
  document.getElementById('task-review-status').textContent=pending?'Review the files and note before deciding.':'';
  setTaskReviewActionsEnabled(pending);
  scrollProofThreadsToBottom(body);
}

function setTaskReviewActionsEnabled(enabled){
  ['task-review-changes-btn','task-review-approve-btn'].forEach(id=>{
    const el=document.getElementById(id);if(el)el.disabled=!enabled;
  });
}

function setTaskReviewBusy(busy,message=''){
  setTaskReviewActionsEnabled(!busy);
  const status=document.getElementById('task-review-status');
  if(status)status.textContent=message;
}

function closeTaskProofReview(){
  document.getElementById('mo-task-proof-review')?.classList.remove('open');
  taskProofReviewContext=null;
}

async function approveTaskProof(){
  const id=taskProofReviewContext?.notificationId;if(id==null)return;
  setTaskReviewBusy(true,'Approving proof...');
  await approveNotification(id);
  const n=notifications.find(x=>x.id===id);
  if(n?.status==='approved'){
    closeTaskProofReview();
    await window.renderMyTasks?.(true);
  }else setTaskReviewBusy(false,'Approval failed. Please try again.');
}

async function requestTaskProofChanges(){
  const id=taskProofReviewContext?.notificationId;if(id==null)return;
  const n=notifications.find(x=>x.id===id);
  const assignment=taskProofReviewContext?.assignment||{};
  if(n){
    n.recipientEmail=n.recipientEmail||assignment.recipientEmail||'';
    n.recipientName=n.recipientName||assignment.recipientName||'';
    n.taskTitle=n.taskTitle||assignment.title||'';
  }
  closeTaskProofReview();
  await showTaskFollowupModal({
    assignmentId:assignment.id||'',appTaskId:assignment.appTaskId||n?.taskId||'',
    title:assignment.title||n?.taskTitle||'',assignerEmail:assignment.assignerEmail||currentUser?.email||'',
    assignerName:assignment.assignerName||currentUser?.name||'',recipientEmail:assignment.recipientEmail||n?.recipientEmail||'',
    recipientName:assignment.recipientName||n?.recipientName||'',role:'assignor',requestChanges:true,notificationId:id,
  });
}

async function approveNotification(id){
  markNotifSeen(id);
  if(highlightedNotifId===id)highlightedNotifId=null;
  const n=notifications.find(x=>x.id===id);
  if(!n)return;
  // KV-based proof_submitted — route through KV approve endpoint
  if(n.type==='proof_submitted'&&n.kvNotifId){
    const task=n.taskId!=null?tasks.find(t=>String(t.id)===String(n.taskId)):null;
    try{
      const fnUrl=workerBaseUrl();
      const userToken=await getAccessToken();
      const res=await fetch(`${fnUrl}/notify`,{
        method:'POST',
        headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
        body:JSON.stringify({type:'proof_result',notifId:n.kvNotifId,appTaskId:String(n.taskId||''),taskTitle:task?.title||n.message||'',senderEmail:currentUser?.email||'',senderName:currentUser?.name||'',recipientEmail:n.recipientEmail||'',result:'approved',reason:''}),
      });
      if(!res.ok){const detail=await res.text().catch(()=>'');throw new Error(detail||`Approval failed (${res.status})`);}
      window.updateTasksTabProofState?.(String(n.taskId||''), 'approved');
      if(task){task.status='Done';task.completedAt=new Date().toISOString();task.proofSubmittedAt=n.submittedAt||n.timestamp||task._proofNotif?.submittedAt||'';task.approvedBy=currentUser?.email||'';delete task._proofNotif;if(Array.isArray(n.proofs)&&n.proofs.length)task.proofs=n.proofs;}
      n.status='approved';
      updateNotifBadge();renderNotifications();
      await Promise.all([saveNotifications(),saveTasksToOneDrive()]);
      syncBadges();refreshAll();
      toast('Proof approved — task marked Done');
    }catch(err){toast('Could not approve: '+err.message);}
    return;
  }
  // Legacy task_completion type
  n.status='approved';
  if(n.taskId){
    const task=tasks.find(t=>t.id===n.taskId);
    if(task){
      task.status='Done';
      task.completedAt=new Date().toISOString();
      task.proofSubmittedAt=n.submittedAt||n.timestamp||task.proofSubmittedAt||'';
      task.approvedBy=currentUser?.email||'';
      if(Array.isArray(n.proofs)&&n.proofs.length)task.proofs=n.proofs;
      syncBadges();refreshAll();
    }
  }
  updateNotifBadge();
  renderNotifications();
  await saveNotifications();
  await saveTasksToOneDrive();
  toast('Task approved as Done');
}

async function dismissNotification(id,providedReason){
  markNotifSeen(id);
  if(highlightedNotifId===id)highlightedNotifId=null;
  const n=notifications.find(x=>x.id===id);
  if(!n)return;
  // KV-based proof_submitted — route through KV decline endpoint + email
  if(n.type==='proof_submitted'&&n.kvNotifId){
    const task=n.taskId!=null?tasks.find(t=>String(t.id)===String(n.taskId)):null;
    const assignment=taskProofReviewContext?.assignment||{};
    const recipientEmail=n.recipientEmail||assignment.recipientEmail||task?.email||'';
    const reason=providedReason!==undefined?providedReason:prompt('Reason for declining this proof?');
    if(reason===null)return;
    try{
      const fnUrl=workerBaseUrl();
      const userToken=await getAccessToken();
      const res=await fetch(`${fnUrl}/notify`,{
        method:'POST',
        headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
        body:JSON.stringify({type:'proof_result',notifId:n.kvNotifId,appTaskId:String(n.taskId||assignment.appTaskId||''),taskTitle:task?.title||n.taskTitle||assignment.title||'',senderEmail:currentUser?.email||'',senderName:currentUser?.name||'',recipientEmail,result:'declined',reason:reason||'No reason provided',todoListId:task?.recipientTodoListId||assignment.recipientTodoListId||'',todoTaskId:task?.recipientTodoTaskId||assignment.recipientTodoTaskId||''}),
      });
      if(!res.ok){const detail=await res.text().catch(()=>'');throw new Error(detail||`Request failed (${res.status})`);}
      window.updateTasksTabProofState?.(String(n.taskId||''), 'declined');
      if(task)delete task._proofNotif;
      n.status='dismissed';n.dismissReason=String(reason||'').trim();
      updateNotifBadge();renderNotifications();
      await Promise.all([saveNotifications(),saveTasksToOneDrive()]);
      await sendProofDeclineEmail(recipientEmail,task?.title||n.taskTitle||assignment.title||'',reason,currentUser?.name||'',task||{
        id:n.taskId||assignment.appTaskId||'',title:n.taskTitle||assignment.title||'',email:recipientEmail,
      });
      refreshAll();
      toast('Changes requested — assignee notified');
    }catch(err){toast('Could not decline: '+err.message);}
    return;
  }
  const reason=prompt('Reason for dismissing this completed task?');
  if(reason===null)return;
  n.status='dismissed';
  n.dismissReason=String(reason||'').trim();
  if(n.taskId){
    const task=tasks.find(t=>t.id===n.taskId);
    if(task){
      task.status='Pending';
      task.dismissReason=n.dismissReason;
      task.lastDismissedAt=new Date().toISOString();
    }
  }
  if(n.todoTaskId)processedEmailIds.delete('todo-'+n.todoTaskId);
  try{localStorage.setItem('dpeg_processed_done_emails',JSON.stringify([...processedEmailIds].slice(-200)));}catch{}
  updateNotifBadge();
  renderNotifications();
  await saveNotifications();
  await sendDismissalEmail(n).catch(err=>console.warn('Dismissal email failed:',err.message));
  toast('Dismissed and emailed assignee');
}

async function markAllNotificationsRead(){
  notifications.filter(n=>n.status==='pending').forEach(n=>n.status='dismissed');
  updateNotifBadge();
  renderNotifications();
  await saveNotifications();
}

// Notifications used to be a single flat list — with several pending
// submissions from different people it read as a wall of tiny near-identical
// text. Grouped by person (submitter/assignee) the same way the Tasks tab
// groups Received/Delegated Assignments, collapsible, so it's scannable at
// a glance. notifManualToggles overrides the default (pending groups start
// open, resolved-only groups start collapsed) once a user actually clicks one.
let notifManualToggles={};

function notifGroupKey(n){
  return String(n.recipientEmail||n.recipientName||'unknown').toLowerCase();
}
function notifGroupLabel(n){
  // recipientName is only as reliable as currentUser.name was at the moment
  // that person submitted proof, which isn't always populated — fall back
  // to the app's own contact directory (staffConfig/userContacts) before
  // giving up and showing the raw email.
  return String(n.recipientName||findPersonByEmail(n.recipientEmail)?.name||n.recipientEmail||'Unknown').trim();
}
function groupNotificationsByPerson(list){
  const grouped=new Map();
  list.forEach(n=>{
    const key=notifGroupKey(n);
    if(!grouped.has(key))grouped.set(key,{key,name:notifGroupLabel(n),items:[]});
    grouped.get(key).items.push(n);
  });
  const groups=[...grouped.values()];
  groups.forEach(g=>g.items.sort((a,b)=>new Date(b.timestamp||0)-new Date(a.timestamp||0)));
  groups.sort((a,b)=>new Date(b.items[0]?.timestamp||0)-new Date(a.items[0]?.timestamp||0));
  return groups;
}
// Collapsed by default, same as the Tasks tab's Received/Delegated groups —
// only expands once the user actually clicks a name.
function isNotifGroupOpen(group){
  return notifManualToggles[group.key]===true;
}
function toggleNotifGroup(key){
  const group=groupNotificationsByPerson(notifications).find(g=>g.key===key);
  const opening=!(group&&isNotifGroupOpen(group));
  notifManualToggles[key]=opening;
  // Same as the Tasks tab's toggleTasksGroup — expanding a group is what
  // clears its "new" badge, marking everything currently in it as seen.
  if(opening&&group){
    let changed=false;
    group.items.forEach(n=>{
      const k=notifSeenKey(n);
      if(!notifSeenIds.has(k)){notifSeenIds.add(k);n.seen=true;changed=true;}
    });
    if(changed)saveNotifSeenIds(notifSeenIds);
  }
  renderNotifications();
}

function renderNotificationCard(n){
  const ts=n.timestamp?new Date(n.timestamp).toLocaleString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'}):'';
  const statusColor=n.status==='approved'?'#15803d':n.status==='dismissed'?'#9ca3af':'#92400e';
  const statusLabel=n.status==='approved'?'Approved':n.status==='dismissed'?'Dismissed':'Pending';
  const icon=n.type==='task_completion'
    ?`<svg width="16" height="16" fill="none" stroke="#15803d" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`
    :`<svg width="16" height="16" fill="none" stroke="#6b7280" viewBox="0 0 24 24"><path d="M15 17h5l-1.405-1.405A2.032 2.032 0 0118 14.158V11a6 6 0 00-4-5.659V5a2 2 0 10-4 0v.341C7.67 6.165 6 8.388 6 11v3.159c0 .538-.214 1.055-.595 1.436L4 17h5m6 0v1a3 3 0 11-6 0v-1m6 0H9" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const proofHtml=renderNotificationProofs(n);
  const isUnseen=isNotifUnseen(n);
  const newTag=isUnseen?`<span style="display:inline-flex;align-items:center;padding:1px 6px;background:#fef3c7;border:1px solid #f59e0b;border-radius:10px;font-size:9px;font-weight:800;color:#92400e;letter-spacing:.4px;margin-left:5px;vertical-align:middle">NEW</span>`:'';
  // The task title used to only exist buried inside the prose message
  // ("X submitted proof for 'task title'"), rendered as one flat gray
  // sentence — easy to miss which task a card was even about. Pulled out
  // as its own bold line so it's the first thing the eye lands on.
  const taskTitleHtml=n.taskTitle?`<div style="font-size:13.5px;font-weight:800;color:${isUnseen?'#92400e':'var(--body)'};margin-bottom:3px">${escapeHtml(n.taskTitle)}</div>`:'';
  return`<div class="card${n.id===highlightedNotifId?' notif-highlight':''}${isUnseen?' notif-unseen':''}" id="notif-card-${n.id}" style="padding:12px 14px">
    <div style="display:flex;align-items:flex-start;gap:10px">
      <div style="flex-shrink:0;margin-top:1px">${icon}</div>
      <div style="flex:1;min-width:0">
        ${taskTitleHtml}
        <div style="font-size:12.5px;color:var(--muted);line-height:1.4;margin-bottom:4px">${escapeHtml(n.message||'')}${newTag}</div>
        ${proofHtml}
        ${n.dismissReason?`<div style="margin-top:6px;font-size:11px;color:#991b1b">Dismissed reason: ${escapeHtml(n.dismissReason)}</div>`:''}
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <span style="font-size:10px;color:var(--muted)">${ts}</span>
          <span style="font-size:10px;font-weight:600;color:${statusColor}">${statusLabel}</span>
        </div>
      </div>
      <div style="display:flex;gap:5px;flex-shrink:0;flex-wrap:wrap;justify-content:flex-end">
        ${n.taskId!=null?`<button class="btn btn-ghost btn-sm" onclick="goToNotificationTask(${n.id})">Go to Action Log</button>`:''}
        ${n.status==='pending'?`
        ${(n.type==='task_completion'||n.type==='proof_submitted')?`<button class="btn btn-primary btn-sm" onclick="approveNotification(${n.id})">Approve</button>`:''}
        <button class="btn btn-ghost btn-sm" onclick="dismissNotification(${n.id})">${n.type==='proof_submitted'?'Request Changes':'Dismiss'}</button>
        `:''}
      </div>
    </div>
  </div>`;
}

// Resolved (approved/dismissed) notifications used to stay mixed in with
// pending ones inside each person's group forever, so the list only ever
// grew. Split into two tabs — Needs Approval / History — same
// directory-toggle pattern as the Tasks tab's Received/Delegated switch,
// rather than stacking both on the page at once.
let notifTabMode='current'; // 'current' | 'history'
window.setNotifTabMode=function setNotifTabMode(mode){
  notifTabMode=mode;
  document.getElementById('notif-current-btn')?.classList.toggle('active',mode==='current');
  document.getElementById('notif-history-btn')?.classList.toggle('active',mode==='history');
  renderNotifications();
};
function updateNotifTabBadges(pendingCount,resolvedCount){
  const setBadge=(id,count)=>{
    const badgeEl=document.getElementById(id);
    if(!badgeEl)return;
    badgeEl.textContent=count>99?'99+':count;
    badgeEl.style.display=count>0?'':'none';
  };
  setBadge('notif-current-badge',pendingCount||0);
  setBadge('notif-history-badge',resolvedCount||0);
}

function renderNotifGroupBlock(group){
  const open=isNotifGroupOpen(group);
  const noun=group.items.length===1?'notification':'notifications';
  // Full breakdown, not just the pending count — "3 notifications · 1
  // pending" left the other 2 unexplained; spell out what they actually are.
  const statusCounts={};
  group.items.forEach(n=>{
    const label=n.status==='approved'?'approved':n.status==='dismissed'?'dismissed':'pending';
    statusCounts[label]=(statusCounts[label]||0)+1;
  });
  const summaryText=['pending','approved','dismissed']
    .filter(label=>statusCounts[label])
    .map(label=>`${statusCounts[label]} ${label}`)
    .join(' · ');
  const unseenCount=group.items.filter(isNotifUnseen).length;
  // Readable red pill beside the name, same pattern as the Tasks tab's
  // "+1 follow-up" badge, instead of a small dot tucked into the avatar
  // corner that's easy to miss. Clears once you've merely opened/viewed it.
  const newBadge=unseenCount>0?`<span class="notif-new-badge">+${unseenCount>9?'9+':unseenCount} new</span>`:'';
  // Separate yellow pill for "still needs action" — unlike newBadge, this
  // does NOT clear just from viewing it; only approving/dismissing removes
  // it, so it stays as a persistent reminder that something is unresolved.
  const pendingCount=statusCounts.pending||0;
  const pendingBadge=pendingCount>0?`<span class="notif-pending-badge">${pendingCount} pending</span>`:'';
  const safeGroupKey=escapeHtml(JSON.stringify(group.key));
  const cards=open?`<div class="assign-cards">${group.items.map(renderNotificationCard).join('')}</div>`:'';
  return `<div class="assign-group">
    <div class="assign-group-head" onclick="toggleNotifGroup(${safeGroupKey})">
      <span class="assign-group-toggle">${open?'−':'+'}</span>
      <span class="assign-avatar-wrap">${av(group.name,24)}</span>
      <span class="assign-group-name" style="color:var(--forest)">${escapeHtml(group.name)}</span>
      ${pendingBadge}
      ${newBadge}
      <span class="assign-group-summary">${group.items.length} ${noun} · ${summaryText}</span>
    </div>
    ${cards}
  </div>`;
}

function renderNotifications(){
  const el=document.getElementById('notif-list');
  if(!el)return;
  const pending=notifications.filter(n=>n.status==='pending');
  const resolved=notifications.filter(n=>n.status!=='pending');
  updateNotifTabBadges(pending.length,resolved.length);
  if(!notifications.length){
    el.innerHTML='<div style="padding:40px;text-align:center;font-size:12px;color:var(--muted)">No notifications yet</div>';
    return;
  }
  // Live count, not a fixed number — always reflects however many are
  // actually unseen right now, and always equals the sum of the per-group
  // badges below it since both are computed from isNotifUnseen. Only pending
  // items ever count, so resolved History items can't inflate this anymore.
  const totalNew=notifications.filter(isNotifUnseen).length;
  const summaryBanner=(notifTabMode==='current'&&totalNew>0)?`<div style="display:flex;align-items:center;gap:8px;padding:10px 14px;margin-bottom:10px;background:var(--sage3);border:1px solid var(--sage);border-radius:var(--rlg)">
    <span style="font-size:13.5px;font-weight:800;color:var(--forest)">${totalNew} new notification${totalNew===1?'':'s'}</span>
    <span style="font-size:11px;color:var(--muted)">— divided by person below</span>
  </div>`:'';

  const list=notifTabMode==='current'?pending:resolved;
  const listHtml=list.length
    ? groupNotificationsByPerson(list).map(renderNotifGroupBlock).join('')
    : notifTabMode==='current'
      ? `<div style="padding:20px 14px;text-align:center;font-size:12px;color:var(--muted);background:var(--sage3);border:1px solid var(--sage);border-radius:var(--rlg)">✓ All caught up — nothing needs action</div>`
      : `<div style="padding:20px 14px;text-align:center;font-size:12px;color:var(--muted)">No resolved notifications yet</div>`;

  el.innerHTML=summaryBanner+listHtml;
  scrollProofThreadsToBottom(el);
}

// Jumping here from a Delegated card's "Review Proof" button used to just
// dump the user into a flat Notifications list they had to search through
// manually. This scrolls straight to the matching card and flashes it, so
// the right one is obvious even with many pending notifications.
function highlightProofNotificationForTask(appTaskId){
  const n=notifications.find(x=>x.type==='proof_submitted'&&x.status==='pending'&&String(x.taskId)===String(appTaskId||''));
  if(!n)return;
  // Baked into renderNotifications() (keyed off highlightedNotifId) rather
  // than a one-off classList mutation, so it survives the 15s poll's
  // re-render instead of vanishing after a couple seconds. Cleared once the
  // user actually resolves this notification (see approveNotification/
  // dismissNotification), not on a timer.
  highlightedNotifId=n.id;
  const groupKey=notifGroupKey(n);
  notifManualToggles[groupKey]=true; // force that person's group open so the card is actually in the DOM to scroll to
  // Forcing it open is the same as clicking to expand it — clears that
  // person's "new" badges too, same as toggleNotifGroup.
  const group=groupNotificationsByPerson(notifications).find(g=>g.key===groupKey);
  if(group){
    let changed=false;
    group.items.forEach(item=>{
      const k=notifSeenKey(item);
      if(!notifSeenIds.has(k)){notifSeenIds.add(k);item.seen=true;changed=true;}
    });
    if(changed)saveNotifSeenIds(notifSeenIds);
  }
  window.setNotifTabMode?.('current'); // target is always pending — make sure that tab is the one showing
  renderNotifications();
  setTimeout(()=>{
    document.getElementById(`notif-card-${n.id}`)?.scrollIntoView({behavior:'smooth',block:'center'});
  },80);
}
window.highlightProofNotificationForTask=highlightProofNotificationForTask;

function goToNotificationTask(id){
  markNotifSeen(id);
  const n=notifications.find(x=>x.id===id);
  if(!n||n.taskId==null){toast('Task not found for this notification');return;}
  const task=tasks.find(t=>String(t.id)===String(n.taskId));
  if(!task){toast('Task not found in Action Log');return;}
  nav('master');
  setTimeout(()=>openDetail(task.id),0);
}

function renderNotificationProofs(n){
  const history=(window._proofSubmissionHistory?.[String(n.taskId||n.appTaskId||'')]||[]);
  if(history.length>1){
    const relatedHtml=renderRelatedProofSubmissions(n);
    return relatedHtml+`<div style="display:grid;gap:9px;margin-top:9px">${history.map((submission,index)=>{
      const number=history.length-index;
      const status=submission.status==='approved'?'Approved':submission.status==='declined'?'Changes requested':'Submitted';
      const when=submission.submittedAt||submission.createdAt||'';
      return `<div style="border:1px solid #dbe4dc;border-radius:7px;background:#fff;padding:9px 10px">
        <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:5px">
          <div style="font-size:11px;font-weight:800;color:var(--forest)">Submission ${number}${index===0?' · Latest':''}</div>
          <div style="font-size:10.5px;color:var(--muted)">${escapeHtml(status)}${when?` · ${escapeHtml(new Date(when).toLocaleString())}`:''}</div>
        </div>${renderSingleProofContents(submission)}</div>`;
    }).join('')}</div>`;
  }
  return renderRelatedProofSubmissions(n)+renderSingleProofContents(n);
}

function renderSingleProofContents(n){
  const proofs=Array.isArray(n.proofs)?n.proofs:[];
  const note=String(n.note||'').trim();
  const noteHtml=note?`<div style="margin:7px 0;padding:8px 10px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:5px">
    <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:#166534;margin-bottom:4px">Submitted note</div>
    <div style="font-size:12px;color:#14532d;line-height:1.5;white-space:pre-wrap">${escapeHtml(note)}</div>
  </div>`:'';
  const filesHtml=proofs.length?`<div style="margin:7px 0;padding:7px 8px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:5px">
    <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:#6b7280;margin-bottom:5px">Attached files</div>
    ${proofs.map(p=>`<div style="display:flex;align-items:center;gap:7px;font-size:11.5px;margin:3px 0;min-width:0">
      <span style="font-weight:700;color:#111;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(p.name||'Proof file')}</span>
      <span style="color:#9ca3af;flex-shrink:0">${p.size?escapeHtml(formatFileSize(p.size)):''}</span>
      ${p.webUrl?`<a href="${escapeHtml(p.webUrl)}" target="_blank" rel="noopener" style="color:#0E3416;font-weight:700;flex-shrink:0">Open / Download</a>`:''}
    </div>`).join('')}
  </div>`:(note?'':`<div style="margin:6px 0;font-size:11px;color:#92400e;background:#fffbeb;border:1px solid #fde68a;border-radius:5px;padding:6px 8px">No proof files were attached.</div>`);
  return noteHtml+filesHtml;
}

function renderRelatedProofSubmissions(n){
  const groupId=String(n.assignmentGroupId||tasks.find(t=>String(t.id)===String(n.taskId))?.assignmentGroupId||'');
  if(!groupId)return '';
  const groupTasks=tasks.filter(t=>String(t.assignmentGroupId||'')===groupId);
  if(groupTasks.length<2)return '';
  const rows=groupTasks.map(t=>{
    const notif=notifications.find(x=>x.type==='proof_submitted'&&String(x.taskId)===String(t.id));
    const state=notif?notif.status||'pending':(nstt(t.status)==='Done'?'submitted':'waiting');
    const color=state==='approved'?'#15803d':state==='dismissed'?'#991b1b':state==='waiting'?'#6b7280':'#92400e';
    const label=state==='approved'?'Approved':state==='dismissed'?'Dismissed':state==='waiting'?'Waiting':'Pending review';
    return `<button type="button" onclick="goToNotificationTask(${notif?notif.id:JSON.stringify('')})" ${notif?'':'disabled'} style="display:flex;align-items:center;gap:6px;border:1px solid #e5e7eb;background:${String(t.id)===String(n.taskId)?'#f0fdf4':'#fff'};border-radius:999px;padding:4px 8px;font-size:11px;color:#374151;${notif?'cursor:pointer':'opacity:.7'}">
      <span style="font-weight:800;color:#111">${escapeHtml(t.person||t.email||'Assignee')}</span>
      <span style="font-weight:700;color:${color}">${label}</span>
    </button>`;
  }).join('');
  return `<div style="margin:7px 0;padding:8px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:6px">
    <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:#6b7280;margin-bottom:6px">Related assignees for this task</div>
    <div style="display:flex;gap:5px;flex-wrap:wrap">${rows}</div>
  </div>`;
}

async function askNotificationFollowup(id,providedMessage){
  markNotifSeen(id);
  const n=notifications.find(x=>x.id===id);
  if(!n)return false;
  const question=providedMessage!==undefined?providedMessage:prompt('Send the assignee a message:');
  if(question===null)return;
  const message=String(question||'').trim();
  if(!message){toast('Please type a message');return;}
  try{
    const task=n.taskId!=null?tasks.find(t=>t.id===n.taskId):null;
    if(n.kvNotifId){
      const fnUrl=workerBaseUrl();
      const userToken=await getAccessToken();
      const res=await fetch(`${fnUrl}/notify`,{
        method:'POST',
        headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
        body:JSON.stringify({type:'proof_followup_question',notifId:n.kvNotifId,appTaskId:String(n.taskId||''),taskTitle:task?.title||'',senderEmail:currentUser?.email||'',senderName:currentUser?.name||'',recipientEmail:n.recipientEmail||'',message}),
      });
      if(!res.ok){const d=await res.text().catch(()=>'');throw new Error(d||'Could not send message');}
      await checkAndLoadProofNotifications();
      renderNotifications();
    }
    if(providedMessage===undefined)toast('Message sent — the recipient was notified in the app');
    return true;
  }catch(err){toast('Could not send message: '+(err.message||err));return false;}
}

async function sendDismissalEmail(n){
  const task=tasks.find(t=>String(t.id)===String(n.taskId))||{};
  const recipient=n.recipientEmail||task.email;
  if(!recipient)return;
  const reason=n.dismissReason||'No reason provided.';
  const html=`<div style="font-family:Arial,sans-serif;max-width:620px">
    <h2 style="color:#991b1b">Task completion dismissed</h2>
    <p><strong>Task:</strong> ${escapeHtml(task.title||'Task')}</p>
    <p><strong>Dismissed by:</strong> ${escapeHtml(currentUser?.name||currentUser?.email||'DPEG Task Manager')}</p>
    <p><strong>Reason:</strong> ${escapeHtml(reason)}</p>
    <p>Please update the task/proof and mark it complete again when ready.</p>
    <hr style="border:none;border-top:1px solid #e5e7eb;margin:16px 0">
    <p style="color:#9ca3af;font-size:12px">DPEG Task Manager - automated notification</p>
  </div>`;
  await replyToTaskEmail(task,html,recipient);
}

// Notifies the recipient that a Delegated task (tasks-hub.js: cancelAssignmentPrompt)
// was called off. Built straight from the D1 assignment record rather than a local
// Action Log task — an assignment made purely through the Tasks Hub doesn't always
// have a matching entry in the assigner's own per-account Action Log.
async function sendTaskCancelledEmail(a){
  const recipient=a?.recipientEmail;
  if(!recipient)return;
  try{
    const reason=String(a.cancelReason||'').trim();
    const html=`<div style="font-family:Arial,sans-serif;max-width:620px">
      <h2 style="color:#991b1b">Task cancelled</h2>
      <p><strong>Task:</strong> ${escapeHtml(a.title||'Task')}</p>
      <p><strong>Cancelled by:</strong> ${escapeHtml(currentUser?.name||currentUser?.email||'DPEG Task Manager')}</p>
      ${reason?`<p><strong>Reason:</strong> ${escapeHtml(reason)}</p>`:''}
      <p>No further action is needed on this task.</p>
      <hr style="border:none;border-top:1px solid #e5e7eb;margin:16px 0">
      <p style="color:#9ca3af;font-size:12px">DPEG Task Manager - automated notification</p>
    </div>`;
    const localTask=tasks.find(t=>String(t.id)===String(a.appTaskId||''));
    await replyToTaskEmail(localTask||{title:a.title||'Task',email:recipient},html,recipient);
  }catch(err){console.warn('Task-cancelled email failed:',err.message);}
}
window.sendTaskCancelledEmail=sendTaskCancelledEmail;

async function listTaskProofFiles(task){
  if(!task?.proofFolderItemId)return [];
  const token=await getAccessToken();
  const res=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(task.proofFolderItemId)}/children?$select=id,name,size,file,folder,webUrl,createdDateTime,lastModifiedDateTime,createdBy`,{
    headers:{Authorization:`Bearer ${token}`}
  });
  if(!res.ok)return [];
  const data=await res.json();
  return (data.value||[]).filter(i=>i.file).map(i=>({
    name:i.name||'Proof file',
    size:i.size||0,
    type:i.file?.mimeType||'',
    uploadedBy:i.createdBy?.user?.email||i.createdBy?.user?.displayName||'',
    uploadedByName:i.createdBy?.user?.displayName||'',
    uploadedAt:i.createdDateTime||i.lastModifiedDateTime||'',
    webUrl:i.webUrl||'',
    driveItemId:i.id||'',
    note:'',
  }));
}

// ============================================================
// TASK COMPLETION POLLING (via Cloudflare Worker /poll-completions)
// ============================================================
async function pollToDoCompletions(manual=false){
  if(!currentUser)return;
  const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL||'').replace(/\/?$/,'');
  if(!fnUrl){if(manual)toast('Worker URL not configured — check Settings');return;}

  const assignments=tasks.filter(t=>
    t.status!=='Done'&&t.recipientTodoListId&&t.recipientTodoTaskId&&
    t.email&&t.email.includes('@dhananipeg.com')
  ).map(t=>({taskId:t.id,recipientEmail:t.email,todoListId:t.recipientTodoListId,todoTaskId:t.recipientTodoTaskId}));

  // Update the monitor count label in the notifications panel
  const countEl=document.getElementById('notif-monitor-count');
  if(countEl)countEl.textContent=assignments.length?`(monitoring ${assignments.length} task${assignments.length>1?'s':''})` :'(no tasks being monitored)';

  if(!assignments.length){
    if(manual)toast('No tasks with To Do IDs to poll. Assign a new task first.');
    return;
  }

  try{
    const userToken=await getAccessToken();
    const res=await fetch(`${fnUrl}/poll-completions`,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':`Bearer ${userToken}`},
      body:JSON.stringify({assignments}),
    });
    if(!res.ok){
      const errText=await res.text().catch(()=>'');
      if(manual)toast(`Poll failed (${res.status}): ${errText.slice(0,120)}`);
      console.warn('Poll failed:',res.status,errText);
      return;
    }
    const data=await res.json();
    let added=0;
    for(const item of (data.completed||[])){
      const key='todo-'+item.todoTaskId;
      if(processedEmailIds.has(key))continue;
      processedEmailIds.add(key);
      const matchTask=tasks.find(t=>t.id===item.taskId);
      if(matchTask&&matchTask.status!=='Done'){
        const alreadyPending=notifications.some(n=>n.taskId===matchTask.id&&n.status==='pending');
        if(alreadyPending)continue;
        const completedOn=item.completedDateTime
          ?new Date(item.completedDateTime.endsWith('Z')?item.completedDateTime:item.completedDateTime+'Z').toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})
          :new Date().toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
        const deadlineStr=matchTask.deadline||matchTask.date?` — Deadline: ${matchTask.deadline||matchTask.date}`:'';
        let proofFiles=Array.isArray(item.proofs)?item.proofs:[];
        try{
          const folderProofs=await listTaskProofFiles(matchTask);
          if(folderProofs.length)proofFiles=folderProofs;
        }catch(err){console.warn('Proof folder scan failed:',err.message);}
        notifications.unshift({
          id:Date.now()+Math.random(),
          type:'task_completion',
          message:`${matchTask.person||'Someone'} completed "${matchTask.title||'task'}"${deadlineStr}. Marked done on ${completedOn}. Do you want to approve?`,
          taskTitle:matchTask.title||'',
          taskId:matchTask.id,
          todoTaskId:item.todoTaskId,
          recipientEmail:item.recipientEmail||matchTask.email||'',
          recipientName:matchTask.person||'',
          proofs:proofFiles,
          timestamp:new Date().toISOString(),
          status:'pending',
          seen:false
        });
        added++;
      }
    }
    if(added>0){
      updateNotifBadge();
      await saveNotifications();
      renderNotifications();
      toast(`${added} task completion${added>1?'s':''} need your approval in Assigned by Me`);
    }else if(manual){
      toast(`Polled ${assignments.length} task${assignments.length>1?'s':''} — no completions yet`);
    }
    try{localStorage.setItem('dpeg_processed_done_emails',JSON.stringify([...processedEmailIds].slice(-200)));}catch{}
  }catch(err){
    if(manual)toast(`Poll error: ${err.message}`);
    console.warn('To Do poll failed:',err.message);
  }
}

async function pollNow(){
  const btn=document.getElementById('btn-poll-now');
  if(btn){btn.disabled=true;btn.textContent='Polling...';}
  await pollToDoCompletions(true);
  await checkAndLoadProofNotifications().catch(()=>{});
  if(btn){btn.disabled=false;btn.textContent='Poll Now';}
}

function startNotifPolling(){
  if(_notifPollTimer)return;
  _notifPollTimer=setInterval(()=>{
    if(document.hidden)return;
    _notifPollCycle++;
    if(_notifPollCycle%4===0)pollToDoCompletions();
    checkAndLoadProofNotifications().catch(()=>{});
    // Keep the Tasks tab's status/proof stages live without a manual reopen,
    // and keep the sidebar/tab alert badges current from any page — silent=true
    // so it only redraws the (possibly-hidden) tab DOM when something changed
    renderMyTasks(true);
  },8000);
  document.addEventListener('visibilitychange',()=>{
    if(!document.hidden){
      checkAndLoadProofNotifications().catch(()=>{});
      renderMyTasks(true);
    }
  });
}
