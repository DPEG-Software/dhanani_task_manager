function renderNtAssigneeChips(){
  const box=document.getElementById('nt-assignee-chips');
  if(!box)return;
  box.innerHTML=ntAssignees.map((p,i)=>`
    <button type="button" class="chip-pill" data-email="${escapeHtml(p.email||'')}" onclick="removeNtAssignee(${i})" title="Remove ${escapeHtml(p.name||p.email||'assignee')}">
      <span class="chip-pill-label">${escapeHtml(p.name||p.email||'Assignee')}</span>
      <span style="font-size:10px;color:#6b7280">${escapeHtml(p.dept||'')}</span>
      <span class="chip-x">&#215;</span>
    </button>`).join('');
}
function addNtAssignee(person){
  const name=String(person?.name||'').trim().replace(/\s+/g,' ');
  const email=normEmail(person?.email||'');
  if(!name&&!email)return;
  const key=email||name.toLowerCase();
  if(ntAssignees.some(p=>(p.email&&email&&normEmail(p.email)===email)||(!email&&String(p.name||'').toLowerCase()===name.toLowerCase()))){
    toast('Assignee already added');
    return;
  }
  // The shared admin mapping is authoritative. Microsoft/personal contact
  // caches may contain stale values (including "Unknown"), so use them only
  // when there is no shared mapping and they contain a real department.
  const configured=personDept(email,name);
  const contactDept=String(person?.dept||'').trim();
  const dept=hasAssignedDepartment(configured)
    ?configured
    :(hasAssignedDepartment(contactDept)?contactDept:'Needs Department');
  ntAssignees.push({name:name||email.split('@')[0],email,dept});
  renderNtAssigneeChips();
  // The visible Department <select> was never synced to the picked person —
  // it just sat on its default first option ("Investor Relations") no
  // matter who got selected, since only ntAssignees[].dept (the small chip
  // subtitle) was ever set. Mirrors setDeptAssignDepartment's sel.value=val
  // pattern used by the sibling Department Assignment picker.
  syncAddTaskDepartmentControl();
  clearAddTaskPerson(false);
}
function findDepartmentAssignmentContact(text){
  const raw=String(text||'').trim().toLowerCase();
  if(!raw)return null;
  return departmentAssignmentContacts().find(p=>
    (raw.includes('@')&&normEmail(p.email)===normEmail(raw)) ||
    String(p.name||'').trim().toLowerCase()===raw
  )||null;
}
function removeNtAssignee(index){
  ntAssignees.splice(index,1);
  renderNtAssigneeChips();
  syncAddTaskDepartmentControl();
}
function syncAddTaskDepartmentControl(){
  const sel=document.getElementById('nt-dept');
  const hint=document.getElementById('nt-dept-hint');
  if(!sel)return;
  const person=ntAssignees.length===1?ntAssignees[0]:null;
  sel.disabled=true;
  const dept=person?.dept||'Needs Department';
  sel.value=[...sel.options].some(o=>o.value===dept)?dept:'Needs Department';
  if(hint)hint.textContent='(managed in Department Settings)';
}
function addTypedNtAssignee(){
  const input=document.getElementById('nt-person');
  const raw=String(input?.value||'').trim();
  if(!raw)return false;
  const contact=findDepartmentAssignmentContact(raw)||{};
  addNtAssignee({
    name:contact.name||(raw.includes('@')?raw.split('@')[0]:raw),
    email:contact.email||(raw.includes('@')?raw:''),
    dept:contact.dept||personDept(contact.email||(raw.includes('@')?raw:''),contact.name||raw)
  });
  return true;
}
function showAddTaskAC(val){
  const ac=document.getElementById('nt-ac');if(!ac)return;
  const input=document.getElementById('nt-person');
  const token=(val||'').trim().toLowerCase();
  const clearBtn=document.getElementById('nt-clear-btn');
  if(clearBtn)clearBtn.style.display=token?'block':'none';
  if(!token){ac.style.display='none';return;}
  const seen=new Set();
  const matches=departmentAssignmentContacts().filter(p=>p?.email||p?.name)
    .filter(p=>{const k=normEmail(p.email||'')||String(p.name||'').toLowerCase();if(!k||seen.has(k))return false;seen.add(k);
      return (p.name||'').toLowerCase().includes(token)||k.includes(token)||(p.role||'').toLowerCase().includes(token)||(p.dept||'').toLowerCase().includes(token);
    }).slice(0,8);
  if(!matches.length){ac.style.display='none';return;}
  ac.innerHTML=matches.map(p=>`
    <div class="compose-ac-item"
      onmousedown="event.preventDefault();selectAddTaskAC('${(p.name||'').replace(/'/g,"\\'")}','${(p.email||'').replace(/'/g,"\\'")}','${(p.dept||'').replace(/'/g,"\\'")}')"
      onmouseover="document.querySelectorAll('#nt-ac .compose-ac-item').forEach(x=>x.classList.remove('ac-focused'));this.classList.add('ac-focused')">
      ${av(p.name||'?',28)}
      <div style="flex:1;min-width:0">
        <div class="compose-ac-name">${p.name||p.email}</div>
        <div class="compose-ac-email">${p.email||''}</div>
        ${p.dept||p.role?`<div class="compose-ac-role">${p.dept||p.role}</div>`:''}
      </div>
    </div>`).join('');
  if(input){const rect=input.getBoundingClientRect();ac.style.position='fixed';ac.style.left=rect.left+'px';ac.style.top=(rect.bottom+2)+'px';ac.style.width=rect.width+'px';ac.style.right='auto';}
  ac.style.display='block';
}
function selectAddTaskAC(name,email,dept){
  addNtAssignee({name,email,dept});
  const ac=document.getElementById('nt-ac');if(ac)ac.style.display='none';
}
function hideAddTaskAC(){setTimeout(()=>{const ac=document.getElementById('nt-ac');if(ac)ac.style.display='none';},180);}
function addTaskACNav(e){
  const ac=document.getElementById('nt-ac');
  if((e.key==='Enter'||e.key==='Tab')&&(!ac||ac.style.display==='none')){
    const added=addTypedNtAssignee();
    if(added)e.preventDefault();
    return;
  }
  if(!ac||ac.style.display==='none')return;
  const items=ac.querySelectorAll('.compose-ac-item');if(!items.length)return;
  const focused=ac.querySelector('.ac-focused');
  if(e.key==='ArrowDown'){e.preventDefault();const next=focused?focused.nextElementSibling||items[0]:items[0];items.forEach(i=>i.classList.remove('ac-focused'));if(next)next.classList.add('ac-focused');}
  else if(e.key==='ArrowUp'){e.preventDefault();const prev=focused?focused.previousElementSibling||items[items.length-1]:items[items.length-1];items.forEach(i=>i.classList.remove('ac-focused'));if(prev)prev.classList.add('ac-focused');}
  else if((e.key==='Enter'||e.key==='Tab')&&focused){e.preventDefault();focused.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));}
  else if(e.key==='Escape'){ac.style.display='none';}
}
function clearAddTaskPerson(focus=true){
  document.getElementById('nt-person').value='';
  document.getElementById('nt-email').value='';
  const disp=document.getElementById('nt-email-display');if(disp)disp.style.display='none';
  const clearBtn=document.getElementById('nt-clear-btn');if(clearBtn)clearBtn.style.display='none';
  document.getElementById('nt-ac').style.display='none';
  if(focus)document.getElementById('nt-person').focus();
}
function setDeadlineBtn(btn,days){
  document.querySelectorAll('.deadline-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const di=document.getElementById('nt-date');
  if(days==='custom'){di.style.display='block';di.focus();return;}
  di.style.display='none';
  if(!days){di.value='';return;}
  const d=new Date();
  if(days==='friday'){while(d.getDay()!==5)d.setDate(d.getDate()+1);}
  else if(days==='monday'){d.setDate(d.getDate()+1);while(d.getDay()!==1)d.setDate(d.getDate()+1);}
  else d.setDate(d.getDate()+Number(days));
  di.value=d.toISOString().split('T')[0];
}
function initSelects() {
  ["sf-dept","sum-d"].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML='<option value="all">All Departments</option>';});
  allDepartments().forEach(d=>{
    ["sf-dept","sum-d"].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML+=`<option>${d}</option>`;});
  });
  document.getElementById("nt-dept").innerHTML=['Needs Department',...allDepartments().filter(d=>d!=='Outside DPEG')].map(d=>`<option>${d}</option>`).join("");
  const assignDept=document.getElementById("dept-assign-dept");
  if(assignDept)assignDept.innerHTML=allDepartments().map(d=>`<option>${d}</option>`).join("");
}
// ============================================================
// DETAIL MODAL
// ============================================================
let detailEditLockTaskId='';
let detailEditLockGranted=false;
let detailEditLockTimer=null;
function setDetailEditAvailability(enabled,message=''){
  ['mo-edit-date','mo-edit-priority','mo-save-detail-btn'].forEach(id=>{
    const el=document.getElementById(id);if(el)el.disabled=!enabled;
  });
  const notice=document.getElementById('mo-edit-lock-notice');
  if(notice){notice.style.display=message?'block':'none';notice.textContent=message;}
}
async function acquireDetailEditLock(task){
  const taskId=String(task?.assignmentId||task?.id||'');
  detailEditLockTaskId=taskId;
  detailEditLockGranted=false;
  setDetailEditAvailability(false,'Checking whether this task is being edited…');
  try{
    const token=await getAccessToken();
    const res=await fetch(`${workerBaseUrl()}/task-edit-lock`,{
      method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({taskId,title:task?.title||''})
    });
    const data=await res.json().catch(()=>({}));
    if(res.status===423){
      setDetailEditAvailability(false,`${data.editorName||'Another user'} is currently editing this task. Please wait.`);
      return false;
    }
    if(!res.ok||!data.success)throw new Error(data.error||`HTTP ${res.status}`);
    if(data.version!=null)task.assignmentVersion=Number(data.version);
    detailEditLockGranted=true;
    setDetailEditAvailability(true,'');
    clearInterval(detailEditLockTimer);
    detailEditLockTimer=setInterval(()=>{
      if(detailEditLockGranted&&detailEditLockTaskId===taskId){
        fetch(`${workerBaseUrl()}/task-edit-lock`,{
          method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({taskId})
        }).catch(()=>{});
      }
    },45000);
    return true;
  }catch(err){
    console.warn('Task edit presence check unavailable:',err.message);
    detailEditLockGranted=true;
    setDetailEditAvailability(true,'Live editing check is temporarily unavailable. Save carefully.');
    return true;
  }
}
function releaseDetailEditLock(){
  clearInterval(detailEditLockTimer);detailEditLockTimer=null;
  const taskId=detailEditLockTaskId;
  const held=detailEditLockGranted;
  detailEditLockTaskId='';detailEditLockGranted=false;
  if(!taskId||!held)return;
  getAccessToken().then(token=>fetch(`${workerBaseUrl()}/task-edit-lock`,{
    method:'DELETE',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({taskId})
  })).catch(()=>{});
}
async function openDetail(id){
  const t=tasks.find(x=>x.id===id);if(!t)return;curTaskId=id;
  document.getElementById("mo-title").textContent=emailSubject(t);
  const personBadge=document.getElementById("mo-person-badge");
  if(personBadge){
    if(t.person){personBadge.textContent=t.person;personBadge.style.display="inline-flex";}
    else personBadge.style.display="none";
  }
  document.getElementById("mo-sub").textContent=`${t.dept} • ${fmtD(t.date)}`;
  document.getElementById("mo-sum").innerHTML=renderSumBox(t);
  document.getElementById("mo-edit-title").value=t.title||"";
  document.getElementById("mo-edit-person").value=t.person||"";
  document.getElementById("mo-edit-email").value=t.email||"";
  document.getElementById("mo-edit-date").value=t.date||"";
  document.getElementById("mo-edit-priority").value=t.priority||"Normal";
  document.getElementById("mo-dept").innerHTML=deptOptions(t.dept);
  const statusDisplay=document.getElementById("mo-status-display");
  if(statusDisplay)statusDisplay.innerHTML=sbadge(t);
  const rs=document.getElementById("mo-refresh-summary");
  if(rs)rs.style.display=t.conversationId||t.emailId?"inline-flex":"none";
  renderProofPanel(t);
  document.getElementById("mo-detail").classList.add("open");
  await acquireDetailEditLock(t);
}
async function saveDetail(){
  const t=tasks.find(x=>x.id===curTaskId);
  if(!t){closeMo("mo-detail");return;}
  if(!detailEditLockGranted){toast('This task is currently being edited by another user');return;}
  const newTitle=(document.getElementById("mo-edit-title").value||"").trim()||t.title;
  const newPerson=(document.getElementById("mo-edit-person").value||"").trim()||t.person;
  const newEmail=(document.getElementById("mo-edit-email").value||"").trim();
  const newDate=document.getElementById("mo-edit-date").value||t.date;
  const newPriority=document.getElementById("mo-edit-priority").value||t.priority||"Normal";
  const newStatus=nstt(t.status);
  const newDept=t.dept;
  const changes=[];
  if(newTitle!==t.title)changes.push(`Title: "${t.title}" → "${newTitle}"`);
  if(newPerson!==t.person)changes.push(`Assigned to: ${t.person} → ${newPerson}`);
  if(newEmail&&newEmail!==t.email)changes.push(`Email updated`);
  if(newDate&&newDate!==t.date)changes.push(`Date: ${fmtD(t.date)} → ${fmtD(newDate)}`);
  if(newPriority!==(t.priority||"Normal"))changes.push(`Priority: ${t.priority||"Normal"} → ${newPriority}`);
  const previous={title:t.title,person:t.person,email:t.email,date:t.date,priority:t.priority,status:t.status};
  t.title=newTitle;t.person=newPerson;if(newEmail)t.email=newEmail;
  t.date=newDate;t.priority=newPriority;t.status=newStatus;
  saveStaffDeptForTask(t,newDept);
  if(changes.length){
    if(t.assignmentId&&await recordAssignment(t)===false){
      Object.assign(t,previous);
      refreshAll();
      toast('This task was updated elsewhere. Close and reopen it before saving.');
      return;
    }
    closeMo("mo-detail");refreshAll();
    updateTodoTask(t,changes).catch(()=>{});
    if(t.email){
      toast('Changes saved');
    }else{
      toast("Changes saved — To Do updated");
    }
  }else{
    closeMo("mo-detail");refreshAll();
    toast("Changes saved");
  }
  await saveTasksToOneDrive();
}
async function refreshTaskSummary(){
  const t=tasks.find(x=>x.id===curTaskId);
  if(!t)return;
  const btn=document.getElementById("mo-refresh-summary");
  if(btn){btn.disabled=true;btn.textContent="Refreshing...";}
  try{
    const email={id:t.emailId||t.lastMessageId,subject:t.emailSubject||t.title,conversationId:t.conversationId,bodyPreview:t.summary||""};
    await attachFreeThreadSummary(email);
    t.summary=taskSummaryWithInstruction(t.taskInstruction,email.threadSummary||t.threadSummary||t.summary);
    t.threadSummary=email.threadSummary||t.threadSummary||"";
    document.getElementById("mo-sum").innerHTML=renderSumBox(t);
    refreshAll();
    toast("Summary refreshed");
    await saveTasksToOneDrive();
  }catch(err){
    toast("Could not refresh summary");
  }finally{
    if(btn){btn.disabled=false;btn.textContent="Refresh Summary";}
  }
}
async function movToWed(){const t=tasks.find(x=>x.id===curTaskId);if(t)t.wednesday=true;closeMo("mo-detail");refreshAll();toast(isWednesdayUser?"Moved to Wednesday notes":"Moved to Discussion Notes");await saveTasksToOneDrive();}
function closeMo(id){
  if(id==='mo-detail')releaseDetailEditLock();
  document.getElementById(id).classList.remove("open");
}
document.querySelectorAll(".mo,.drill-overlay").forEach(m=>m.addEventListener("click",e=>{
  if(e.target!==m)return;
  // Add Task holds an in-progress, unsaved form — a stray click outside it
  // must not silently discard what's been typed. Only its own X/Cancel
  // controls (which call closeMo directly) may close it.
  if(m.id==='mo-add')return;
  if(m.classList.contains('mo'))closeMo(m.id);else m.classList.remove("open");
}));

// ============================================================
// ADD TASK
// ============================================================
function openAdd(){
  document.getElementById("nt-date").value='';
  document.getElementById("nt-date").style.display='none';
  document.querySelectorAll('.deadline-btn').forEach((b,i)=>b.classList.toggle('active',i===0));
  ntAssignees=[];
  renderNtAssigneeChips();
  const deptSel=document.getElementById('nt-dept');
  if(deptSel)deptSel.value='Needs Department';
  syncAddTaskDepartmentControl();
  clearAddTaskPerson(false);
  document.getElementById("mo-add").classList.add("open");
}
async function addTask(){
  const title=document.getElementById("nt-title").value.trim();
  if(!title){toast("Please enter a task title");return;}
  addTypedNtAssignee();
  const fallbackDept=document.getElementById("nt-dept").value;
  const assignees=ntAssignees.length?ntAssignees:[{name:"Unassigned",email:"",dept:fallbackDept}];
  const summary=document.getElementById("nt-summary").value||"";
  const proofInstructions=document.getElementById("nt-proof-instructions")?.value.trim()||"";
  const date=document.getElementById("nt-date").value;
  const status=document.getElementById("nt-status").value;
  const priority=document.getElementById("nt-priority").value;
  const groupId=assignees.length>1?`grp-${Date.now()}-${Math.random().toString(36).slice(2,7)}`:'';
  const baseId=Date.now();
  const assignedAt=new Date(baseId).toISOString();
  const newTasks=assignees.map((a,i)=>{
    const email=normEmail(a.email||"");
    const person=String(a.name||email.split('@')[0]||"Unassigned").trim();
    const configured=personDept(email,person);
    const dept=isInternalEmail(email)
      ?(hasAssignedDepartment(a.dept)?a.dept:(hasAssignedDepartment(configured)?configured:'Needs Department'))
      :(email?'Outside DPEG':fallbackDept||'Needs Department');
    return {id:baseId+i,title,summary,proofInstructions,person,email,dept,date,status,priority,assignedAt,createdAt:assignedAt,wednesday:false,followup:"",weekOffset:0,assignmentGroupId:groupId,assignmentGroupSize:assignees.length};
  });
  tasks.unshift(...newTasks);
  newTasks.forEach(t=>saveStaffDeptForTask(t,t.dept));
  closeMo("mo-add");
  ["nt-title","nt-summary","nt-proof-instructions"].forEach(i=>{if(document.getElementById(i))document.getElementById(i).value="";});
  ntAssignees=[];
  renderNtAssigneeChips();
  clearAddTaskPerson();
  refreshAll();
  await Promise.allSettled(newTasks.map(t=>sendTaskNotification(t)));
  await Promise.allSettled(newTasks.map(t=>createToDoTask(t)));
  toast(`${newTasks.length} task${newTasks.length>1?'s':''} saved${newTasks.some(t=>t.email)?' and assigned':''}`);
  await saveTasksToOneDrive();
}

// ============================================================
// DRILL DOWN
// ============================================================
function drillStat(type){
  let list,title,sub;
  if(type==="high"){list=tasks.filter(t=>String(t.priority||"Normal").toLowerCase()==="high"&&isOpenTask(t));title="High Priority Tasks";sub=`${list.length} tasks needing immediate attention`;}
  if(type==="overdue"){list=tasks.filter(isOverdueTask).sort((a,b)=>new Date(a.date)-new Date(b.date));title="Overdue Tasks";sub=`${list.length} tasks past their deadline`;}
  if(!list||!list.length){toast("No tasks in this category");return;}
  document.getElementById("drill-title").textContent=title;document.getElementById("drill-sub").textContent=sub;
  document.getElementById("drill-tbody").innerHTML=list.map(t=>`
    <tr style="border-bottom:1px solid #f3f4f6;cursor:pointer" onclick="closeMo('mo-drill');openDetail(${t.id})">
      <td style="padding:11px 14px"><div style="font-size:13px;font-weight:600;color:var(--body)">${t.title}</div></td>
      <td style="padding:11px 14px;font-size:12.5px">${t.person}</td>
      <td style="padding:11px 14px"><span class="dept-pill"><span class="dept-dot" style="background:${dcolor(t.dept)}"></span>${t.dept}</span></td>
      <td style="padding:11px 14px">${sbadge(t)}</td><td style="padding:11px 14px">${pBadge(t.priority)}</td>
      <td style="padding:11px 14px;font-size:12px;color:var(--muted);white-space:nowrap">${fmtD(t.date)}</td>
    </tr>`).join("");
  document.getElementById("mo-drill").classList.add("open");
}
function doSearch(v){
  curSearch=String(v||"").toLowerCase();
  if(curSearch&& !document.getElementById("page-master").classList.contains("active")){
    nav("master");
    curSearch=String(v||"").toLowerCase();
    document.getElementById("srch").value=v;
  }
  if(document.getElementById("page-master").classList.contains("active"))renderMaster();
}
function renderProofPanel(t){
  const panel=document.getElementById('mo-proof-panel');
  if(!panel)return;
  if(!t?._proofNotif){panel.innerHTML='';return;}
  const n=t._proofNotif;
  const proofFiles=(n.proofs||[]).filter(p=>p.webUrl||p.shareId);
  const summary=proofFiles.length
    ?`${proofFiles.length} file${proofFiles.length>1?'s':''} submitted`
    :'A message was submitted';
  panel.innerHTML=`<div class="mo-section-lbl">Proof Submission</div>
    <button type="button" onclick="goToProofInTasks(${t.id})" style="display:flex;align-items:center;justify-content:space-between;gap:10px;width:100%;text-align:left;padding:11px 13px;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;cursor:pointer;font-family:inherit">
      <span style="font-size:12.5px;color:#92400e;line-height:1.4">
        <strong>${escapeHtml(n.recipientName||n.recipientEmail||'Someone')}</strong> submitted proof — ${summary}. Review the submission and conversation.
      </span>
      <span style="font-size:15px;color:#92400e;flex-shrink:0">→</span>
    </button>`;
}
