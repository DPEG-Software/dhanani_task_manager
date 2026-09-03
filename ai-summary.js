// ============================================================
// AI SUMMARIZATION
// ============================================================
function redactAITextForPreview(value,maxLength=12000){
  return String(value||'').slice(0,maxLength)
    .replace(/\b\d{3}-\d{2}-\d{4}\b/g,'[REDACTED SSN]')
    .replace(/\b(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}\b/g,'[REDACTED PHONE]')
    .replace(/\b(?:\d[ -]*?){13,19}\b/g,'[REDACTED PAYMENT NUMBER]')
    .replace(/\b(routing|account|bank account|tax id|ein)\s*(?:number|no\.?|#)?\s*[:=-]?\s*[A-Z0-9-]{4,}\b/gi,'$1: [REDACTED]')
    .replace(/\b(password|passcode|pin|access code|security code|secret)\s*[:=-]\s*\S+/gi,'$1: [REDACTED]')
    .replace(/\bBearer\s+[A-Za-z0-9._~-]+/gi,'Bearer [REDACTED]');
}

const SENSITIVE_AI_BLOCK_MESSAGE='Not safe — contains sensitive info.';
function containsSensitiveAIText(value){
  const text=String(value||'');
  return /\b\d{3}-\d{2}-\d{4}\b/.test(text)
    || /\b(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}\b/.test(text)
    || /\b(?:\d[ -]*?){13,19}\b/.test(text)
    || /\b\d{8,12}\b/.test(text)
    || /\b(routing|account|bank account|tax id|ein|social security)\s*(?:number|no\.?|#)?\s*[:=-]?\s*[A-Z0-9-]{3,}\b/i.test(text)
    || /\b(password|passcode|pin|access code|security code|secret|api key|token)\s*[:=-]\s*\S+/i.test(text)
    || /\bBearer\s+[A-Za-z0-9._~-]+/i.test(text);
}

async function callAISummary(messages, subject, opts={}){
  if(localStorage.getItem('dpeg_ai_enabled')==='false')return null;
  const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL||'').replace(/\/?$/,'');
  if(!fnUrl)return null;
  try{
    const token=await getAccessToken();
    const ordered=[...(messages||[])].sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime||0)-new Date(b.receivedDateTime||b.sentDateTime||0));
    const sender=(()=>{
      for(const m of ordered){
        const addr=normEmail(m.from?.emailAddress?.address||'');
        if(addr&&!addr.includes('@dhananipeg.com'))
          return (m.from?.emailAddress?.name||addr.split('@')[0]).replace(/["'<>]/g,'').trim();
      }
      return null;
    })();
    // Format each message with clear FROM/DATE attribution so the AI knows context
    let emailText=ordered.map(m=>{
      const from=m.from?.emailAddress?.name||m.from?.emailAddress?.address||'Unknown';
      const dt=new Date(m.receivedDateTime||m.sentDateTime||0);
      const dateStr=isNaN(dt)?'':dt.toLocaleDateString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
      const rawHtml=m.body?.contentType==='html'?m.body?.content:'';
      const rawText=rawHtml?'':m.body?.content||m.bodyPreview||'';
      const text=rawHtml?cleanEmailBodyForAI(rawHtml):(icalToReadableText(rawText)||cleanEmailText(rawText));
      return `[${from}${dateStr?' — '+dateStr:''}]\n${text}`;
    }).join('\n---\n');
    // Append task assignment date context if provided
    if(opts.emailDate&&opts.assignedDate){
      const eDt=new Date(opts.emailDate);
      const aDt=new Date(opts.assignedDate);
      const eStr=isNaN(eDt)?opts.emailDate:eDt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
      const aStr=isNaN(aDt)?opts.assignedDate:aDt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
      emailText+=`\n\n[TASK CONTEXT] This email was received on ${eStr}. The task is being assigned today: ${aStr}.`;
    }
    // Latest message details sent separately so the worker can anchor Action needed to it
    const latest=ordered[ordered.length-1];
    const latestSender=latest?.from?.emailAddress?.name||latest?.from?.emailAddress?.address||'';
    const latestDt=new Date(latest?.receivedDateTime||latest?.sentDateTime||0);
    const latestDate=isNaN(latestDt)?'':latestDt.toLocaleDateString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
    const latestRawHtml=latest?.body?.contentType==='html'?latest?.body?.content:'';
    const latestMessageText=latestRawHtml?cleanEmailBodyForAI(latestRawHtml):cleanEmailText(latest?.body?.content||latest?.bodyPreview||'');
    if(containsSensitiveAIText(`${subject}\n${emailText}\n${latestMessageText}`))return SENSITIVE_AI_BLOCK_MESSAGE;
    if(localStorage.getItem('dpeg_ai_preview_enabled')!=='false'){
      const preview=`Subject: ${redactAITextForPreview(subject,300)}\n\n${redactAITextForPreview(emailText,3500)}`;
      if(!confirm(`The following redacted content will be sent to Groq for summarization:\n\n${preview}\n\nContinue?`))return null;
    }
    const res=await fetch(fnUrl,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':`Bearer ${token}`},
      body:JSON.stringify({subject:subject||'Email thread',emailText,senderName:sender||'',messageCount:ordered.length,latestMessageText,latestSender,latestDate,...(opts.attachmentNames?.length?{attachmentNames:opts.attachmentNames}:{})})
    });
    if(!res.ok)return null;
    const data=await res.json();
    return data.summary||null;
  }catch{
    return null;
  }
}

async function attachFreeThreadSummary(email, opts={}){
  try{
    let messages=[email];
    if(email.conversationId){
      const token=await getAccessToken();
      const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
      const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview';
      const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
      if(res.ok){const d=await res.json();if(d.value?.length)messages=d.value;}
    }
    // Try AI first, fall back to local algorithm
    const subject=email.subject||email.emailSubject||'';
    const aiSummary=await callAISummary(messages,subject,opts);
    email.threadSummary=aiSummary||localThreadSummary(messages);
    if(aiSummary)email.aiGenerated=true;
  }catch(err){
    email.threadSummary=localThreadSummary([email]);
  }
  return email;
}

async function fetchThreadMessagesForSummary(email,limit=15){
  if(!email?.conversationId)return [email].filter(Boolean);
  try{
    const token=await getAccessToken();
    const filter=encodeURIComponent(`conversationId eq '${email.conversationId.replace(/'/g,"''")}'`);
    const select='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,flag,importance,hasAttachments';
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=${limit}&$select=${select}&$filter=${filter}`,{headers:{Authorization:`Bearer ${token}`}});
    if(res.ok){
      const d=await res.json();
      if(d.value?.length)return d.value;
    }
  }catch{}
  return [email].filter(Boolean);
}

function upsertTaskFromEmail(email,person,address,dept,date){
  const conversationId=email.conversationId||"";
  const emailPriority=String(email.importance||"normal").toLowerCase()==="high"||outlookFlagStatus(email)==='flagged'?"High":"Normal";
  const existing=conversationId?tasks.find(t=>t.conversationId===conversationId):null;
  if(existing){
    existing.emailId=email.id||existing.emailId;
    existing.lastMessageId=email.id||existing.lastMessageId||"";
    existing.emailSubject=email.subject||existing.emailSubject||existing.title;
    existing.title=existing.title||email.subject||"(no subject)";
    const refreshedContext=email.threadSummary||email.bodyPreview||existing.threadSummary||existing.summary||"";
    existing.summary=taskSummaryWithInstruction(existing.taskInstruction,refreshedContext);
    existing.threadSummary=email.threadSummary||existing.threadSummary||"";
    existing.person=person||existing.person;
    existing.email=address||existing.email||"";
    existing.dept=dept||existing.dept;
    existing.priority=emailPriority==="High"?"High":(existing.priority||"Normal");
    existing.lastReplyAt=date||new Date().toISOString();
    existing.replyCount=(existing.replyCount||1)+1;
    return {task:existing,created:false};
  }
  const rawEmailDate=email.receivedDateTime||email.sentDateTime||'';
  const assignedAt=new Date().toISOString();
  const task={id:Date.now()+Math.random(),assignedAt,createdAt:assignedAt,emailId:email.id,emailSubject:email.subject||'(no subject)',conversationId,lastMessageId:email.id||'',title:email.subject||'(no subject)',summary:email.threadSummary||email.bodyPreview||'',threadSummary:email.threadSummary||'',person,email:address,dept,date:(date?new Date(date):new Date()).toISOString().split('T')[0],emailDate:rawEmailDate?new Date(rawEmailDate).toISOString().split('T')[0]:'',status:'Pending',priority:emailPriority,wednesday:false,followup:'',weekOffset:0,replyCount:1,lastReplyAt:date||assignedAt};
  tasks.unshift(task);
  return {task,created:true};
}

async function refreshTaskFromThreadMessages(messages,subject=''){
  if(!Array.isArray(messages)||!messages.length)return;
  const conversationId=messages.find(m=>m.conversationId)?.conversationId;
  if(!conversationId)return;
  const t=tasks.find(x=>x.conversationId===conversationId&&nstt(x.status)!=="Done");
  if(!t)return;
  const ordered=[...messages].sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime||0)-new Date(b.receivedDateTime||b.sentDateTime||0));
  const latest=ordered[ordered.length-1];
  const hasHigh=ordered.some(m=>String(m.importance||"normal").toLowerCase()==="high"||outlookFlagStatus(m)==='flagged');
  t.lastMessageId=latest.id||t.lastMessageId||t.emailId||"";
  t.emailId=t.emailId||latest.id||"";
  t.emailSubject=latest.subject||subject||t.emailSubject||t.title;
  t.title=t.title||t.emailSubject||"(no subject)";
  t.replyCount=ordered.length;
  t.lastReplyAt=latest.receivedDateTime||latest.sentDateTime||t.lastReplyAt;
  if(hasHigh)t.priority="High";
  const local=localThreadSummary(ordered);
  if(local&&!t.aiGenerated)t.summary=taskSummaryWithInstruction(t.taskInstruction,local);
  // Only call AI if: no summary yet, or new messages arrived since last summary was generated
  const lastSumAt=t.summaryGeneratedAt?new Date(t.summaryGeneratedAt).getTime():0;
  const hasNewMessages=ordered.some(m=>new Date(m.receivedDateTime||m.sentDateTime||0).getTime()>lastSumAt);
  if(!t.aiGenerated||hasNewMessages){
    try{
      const ai=await callAISummary(ordered,t.emailSubject||subject||'');
      if(ai){t.summary=taskSummaryWithInstruction(t.taskInstruction,ai);t.threadSummary=ai;t.aiGenerated=true;t.summaryGeneratedAt=new Date().toISOString();}
    }catch{}
    syncBadges();
    renderMaster();
    if(document.querySelector(".page.active")?.id==="page-wednesday")renderWed();
    await saveTasksToOneDrive();
  }
}

function sbadge(t) {
  const st=nstt(t.status)==="Done"?"Done":(t._proofNotif?"Awaiting Approval":(isOverdueTask(t)?"Overdue":"Pending"));
  const m={Pending:"bp",Done:"bd",Overdue:"bo","Awaiting Approval":"bo"};
  return `<span class="badge ${m[st]||"bp"}">${st}</span>`;
}
function av(name,size) {
  const sz=size||32;
  return `<div style="width:${sz}px;height:${sz}px;border-radius:50%;background:${dcolor(personDept("",name))};display:flex;align-items:center;justify-content:center;font-size:${Math.round(sz*.34)}px;font-weight:700;color:#fff;flex-shrink:0">${ini(name)}</div>`;
}
function pBadge(p) {
  const pr=String(p||"Normal").toLowerCase()==="high"?"High":"Normal";
  return `<span class="${pr==="High"?"p-high":"p-low"}">${pr}</span>`;
}
function deptAbbr(dept){
  const map={"Property Management":"PM","Investor Relations":"IR","Legal and Title":"LT","Outside DPEG":"EXT"};
  return map[dept]||String(dept||"").split(/\s|-/).filter(Boolean).map(x=>x[0]).join("").toUpperCase().slice(0,3)||"--";
}
function peopleOptions(){
  const map=new Map();
  Object.values(staffConfig).forEach(p=>{if(p?.name)map.set(p.name.toLowerCase(),p);});
  tasks.forEach(t=>{
    if(t.person&&(isInternalEmail(t.email)||staffConfig[staffKey(t.email,t.person)])){
      map.set(t.person.toLowerCase(),{name:t.person,email:t.email||"",dept:t.dept||"Needs Department"});
    }
  });
  return [...map.values()].sort((a,b)=>String(a.name).localeCompare(String(b.name)));
}
function openDepartmentSettings(){
  if(!isAdmin()){toast('Admin access only');return;}
  initSelects();
  renderDepartmentSettingsList();
  renderDepartmentSettingsCatalog();
  document.getElementById('mo-department-settings')?.classList.add('open');
}
function renderDepartmentSettingsCatalog(){
  const el=document.getElementById('department-settings-catalog');
  if(!el)return;
  el.innerHTML=allDepartments().map(dept=>`<span class="dept-pill" style="font-size:11px;padding:4px 8px"><span class="dept-dot" style="background:${dcolor(dept)}"></span>${escapeHtml(dept)}</span>`).join('');
}
async function createDepartmentSetting(){
  if(!isAdmin()){toast('Admin access only');return;}
  const input=document.getElementById('dept-new-name');
  const name=String(input?.value||'').trim().replace(/\s+/g,' ');
  if(!name){toast('Enter a department name');return;}
  if(allDepartments().some(d=>d.toLowerCase()===name.toLowerCase())){toast('Department already exists');return;}
  customDepartments=[...(customDepartments||[]),name].sort((a,b)=>a.localeCompare(b));
  if(input)input.value='';
  initSelects();
  renderDepartmentSettingsCatalog();
  renderDepartmentSettingsList();
  const sharedSaved=await saveSharedDepartmentSettings();
  await saveTasksToOneDrive();
  toast(sharedSaved?`${name} added for everyone`:`${name} added locally; shared save failed`);
}
function renderDepartmentSettingsList(){
  const el=document.getElementById('department-settings-list');
  if(!el)return;
  const people=departmentAssignmentContacts().filter(p=>p.name||p.email).slice(0,80);
  if(!people.length){el.innerHTML='<div style="padding:20px;text-align:center;font-size:12px;color:var(--muted)">No saved contacts yet. Sync contacts or assign a person above.</div>';return;}
  el.innerHTML=people.map((p,i)=>`
    <div style="display:grid;grid-template-columns:1fr 170px;gap:10px;align-items:center;padding:9px 10px;border:1px solid var(--border);border-radius:7px;background:#fff">
      <div style="display:flex;align-items:center;gap:8px;min-width:0">
        ${av(p.name||p.email||'?',28)}
        <div style="min-width:0">
          <div style="font-size:12.5px;font-weight:700;color:var(--body);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(p.name||p.email)}</div>
          <div style="font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(p.dept||'Needs Department')}</div>
        </div>
      </div>
      <select class="form-sel" style="padding:6px 8px;font-size:12px" onchange="quickReassignDepartment(${i},this.value)">${deptOptions(p.dept||'Needs Department')}</select>
    </div>`).join('');
  window._departmentSettingsPeople=people;
}
async function quickReassignDepartment(index,dept){
  const p=(window._departmentSettingsPeople||[])[index];
  if(!p)return;
  const nameEl=document.getElementById('dept-assign-name');
  const emailEl=document.getElementById('dept-assign-email');
  if(nameEl)nameEl.value=p.name||'';
  if(emailEl)emailEl.value=p.email||'';
  setDeptAssignDepartment(dept);
  await saveDepartmentAssignmentSetting();
}
