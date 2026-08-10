const WL=["This Week","Last Week","2 Weeks Ago","3 Weeks Ago","4 Weeks Ago"];
function weekRange(o){
  const n=new Date(),d=n.getDay();
  const m=new Date(n);m.setDate(n.getDate()-d+1-o*7);
  const s=new Date(m);s.setDate(m.getDate()+6);
  return `${m.toLocaleDateString("en-US",{month:"short",day:"numeric"})} to ${s.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"})}`;
}
function chWeek(d){curWeek=Math.max(0,curWeek-d);document.getElementById("wn-l").textContent=WL[curWeek]||`${curWeek} weeks ago`;document.getElementById("wn-r").textContent=weekRange(curWeek);syncBadges();renderMaster();}

// ============================================================
// ACTION LOG
// ============================================================
// Buckets by when the task was CREATED, not its due date. Every task id in
// this app is Date.now()-based (see addTask/createTaskFromEmail/etc.), so
// it doubles as a reliable creation timestamp without needing a dedicated
// createdAt field (which most tasks don't actually have set).
function taskWeekOffset(t){
  const d=new Date(t.id);
  const toMon=x=>{const r=new Date(x);r.setDate(r.getDate()-(r.getDay()||7)+1);r.setHours(0,0,0,0);return r;};
  return Math.max(0,Math.round((toMon(new Date())-toMon(d))/(7*24*60*60*1000)));
}
function taskCompletedWeekOffset(t){
  const raw=t.completedAt||t.approvedAt||t.proofReviewedAt||t.date||t.id;
  const d=new Date(raw);
  if(Number.isNaN(d.getTime()))return taskWeekOffset(t);
  const toMon=x=>{const r=new Date(x);r.setDate(r.getDate()-(r.getDay()||7)+1);r.setHours(0,0,0,0);return r;};
  return Math.max(0,Math.round((toMon(new Date())-toMon(d))/(7*24*60*60*1000)));
}
function taskAssignedAt(t){
  if(t.assignedAt||t.createdAt)return t.assignedAt||t.createdAt;
  const numericId=Number(t.id);
  return Number.isFinite(numericId)?new Date(Math.floor(numericId)).toISOString():'';
}
function taskDeadlineAt(t){return Object.prototype.hasOwnProperty.call(t,'deadline')?(t.deadline||''):(t.date||'');}
function taskSubmittedAt(t){
  if(t.proofSubmittedAt||t.submittedAt)return t.proofSubmittedAt||t.submittedAt;
  const uploads=(t.proofs||[]).map(p=>p?.uploadedAt).filter(Boolean).sort();
  return uploads.length?uploads[uploads.length-1]:'';
}
function actionTimelineDate(value){
  if(!value)return '';
  const d=new Date(value);
  return Number.isNaN(d.getTime())?'':d.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
}
function submittedAfterDeadline(t){
  const submitted=taskSubmittedAt(t),deadline=taskDeadlineAt(t);
  if(!submitted||!deadline)return false;
  const m=String(deadline).match(/^(\d{4})-(\d{2})-(\d{2})/);
  const due=m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),23,59,59,999):new Date(deadline);
  const sent=new Date(submitted);
  return !Number.isNaN(due.getTime())&&!Number.isNaN(sent.getTime())&&sent>due;
}
function completedTaskTimelineHTML(t){
  const assigned=actionTimelineDate(taskAssignedAt(t));
  const deadline=actionTimelineDate(taskDeadlineAt(t));
  const submitted=actionTimelineDate(taskSubmittedAt(t));
  const approved=actionTimelineDate(t.completedAt||t.approvedAt||t.proofReviewedAt);
  const late=submittedAfterDeadline(t);
  return `<div style="display:grid;gap:2px;font-size:10.5px;line-height:1.35">
    <div><span style="color:var(--muted)">Assigned:</span> ${assigned||'Unavailable'}</div>
    <div><span style="color:var(--muted)">Deadline:</span> ${deadline||'No deadline'}</div>
    <div><span style="color:var(--muted)">Submitted:</span> ${submitted||'Unavailable'}${late?' <span style="color:#b91c1c;font-weight:800">• Submitted overdue</span>':''}</div>
    <div style="color:var(--leaf);font-weight:700">Approved: ${approved||'Unavailable'}</div>
  </div>`;
}
function completedActionCount(){
  return tasks.filter(t=>nstt(t.status)==="Done"&&taskCompletedWeekOffset(t)===curWeek).length;
}
function syncMasterCompletedToggle(){
  const btn=document.getElementById("master-done-toggle");
  if(!btn)return;
  const count=completedActionCount();
  btn.textContent=`${showMasterCompleted?'Hide':'Show'} completed (${count})`;
  btn.setAttribute('aria-pressed',showMasterCompleted?'true':'false');
  btn.style.background=showMasterCompleted?'var(--leaf-bg)':'';
  btn.style.borderColor=showMasterCompleted?'var(--leaf-bd)':'';
  btn.style.color=showMasterCompleted?'var(--leaf)':'';
  const statusFilter=document.getElementById('sf-status');
  if(statusFilter){if(showMasterCompleted)statusFilter.value='all';statusFilter.style.display=showMasterCompleted?'none':'';}
}
function toggleMasterCompleted(){
  showMasterCompleted=!showMasterCompleted;
  renderMaster();
}
function sortActionTasks(list){
  return [...list].sort((a,b)=>{
    const ad=nstt(a.status)==="Done",bd=nstt(b.status)==="Done";
    if(ad!==bd)return ad?1:-1;
    if(ad&&bd){
      const at=new Date(a.completedAt||a.approvedAt||a.proofReviewedAt||a.date||a.id).getTime()||0;
      const bt=new Date(b.completedAt||b.approvedAt||b.proofReviewedAt||b.date||b.id).getTime()||0;
      if(at!==bt)return bt-at;
    }
    const ao=isOverdueTask(a)?0:1;
    const bo=isOverdueTask(b)?0:1;
    if(ao!==bo)return ao-bo;
    const ap=String(a.priority||"Normal").toLowerCase()==="high"?0:1;
    const bp=String(b.priority||"Normal").toLowerCase()==="high"?0:1;
    if(ap!==bp)return ap-bp;
    return b.id-a.id;
  });
}
function getVis(){
  const sf=document.getElementById("sf-status")?.value||"all";
  const df=document.getElementById("sf-dept")?.value||"all";
  const pf=document.getElementById("sf-priority")?.value||"all";
  return sortActionTasks(tasks.filter(t=>{
    const isDone=nstt(t.status)==="Done";
    const isCancelled=nstt(t.status)==="Cancelled";
    const inRelevantWeek=isDone?taskCompletedWeekOffset(t)===curWeek:taskWeekOffset(t)===curWeek;
    const wm=curSearch?true:inRelevantWeek;
    const sm=!curSearch||[t.title,t.emailSubject||"",t.person,t.email||"",t.dept,t.summary||""].some(x=>String(x).toLowerCase().includes(curSearch));
    const stm=sf==="all"||(sf==="Overdue"?isOverdueTask(t):nstt(t.status)===sf);
    const completionVisible=showMasterCompleted?isDone:!isDone;
    return !isCancelled&&wm&&completionVisible&&sm&&stm&&(df==="all"||t.dept===df)&&(pf==="all"||(String(t.priority||"Normal").toLowerCase()==="high"?"High":"Normal")===pf);
  }));
}
function renderMaster(){
  syncMasterCompletedToggle();
  const list=getVis(),tb=document.getElementById("master-tbody");
  if(!list.length){
    tb.innerHTML=showMasterCompleted
      ?`<tr><td colspan="8"><div class="empty-state"><div class="es-text">No completed tasks this week</div><div class="es-sub">Approved tasks will appear here based on their approval date</div></div></td></tr>`
      :`<tr><td colspan="8"><div class="empty-state"><div class="es-text">No active tasks yet</div><div class="es-sub">Tasks appear automatically when you add emails or forward messages to team members</div></div></td></tr>`;
    return;
  }
  tb.innerHTML=list.map(t=>{
    const priorityNote=String(t.priority||"Normal").toLowerCase()==="high"?`<div style="margin-top:4px;color:#b91c1c;font-size:10.5px;font-weight:700">Prioritized from Outlook</div>`:'';
    return `
    <tr onclick="openDetail(${t.id})">
      <td style="padding:10px 14px 10px 16px">
        <div style="font-size:13px;font-weight:600;color:var(--body);line-height:1.3">${emailSubject(t)}</div>
        ${t.replyCount>1?`<div style="font-size:11px;color:var(--muted);margin-top:2px">${t.replyCount} msgs</div>`:""}
        ${priorityNote}
      </td>
      <td style="padding:10px 10px">
        <div style="display:flex;align-items:center;gap:6px">${av(t.person,22)}<span style="font-size:12px;color:var(--body);max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${t.person}</span></div>
      </td>
      <td style="padding:10px 8px"><span class="dept-pill" style="font-size:10.5px;padding:2px 7px"><span class="dept-dot" style="background:${dcolor(t.dept)}"></span>${t.dept}</span></td>
      <td style="padding:10px 8px">${sbadge(t)}</td>
      <td style="padding:10px 8px">${pBadge(t.priority)}</td>
      <td onclick="event.stopPropagation()" style="padding:10px 8px">
        <div class="wed-btn${t.wednesday?" on":""}" style="width:auto;padding:0 8px;font-size:10.5px;white-space:nowrap" onclick="togWed(${t.id})">${t.wednesday?(isWednesdayUser?"On Wed":"In Discussion"):(isWednesdayUser?"+ Wed":"+ Discussion")}</div>
      </td>
      <td onclick="event.stopPropagation()" style="padding:10px 8px;font-size:11.5px;color:var(--body);white-space:normal;line-height:1.35">${nstt(t.status)==="Done"?completedTaskTimelineHTML(t):(t.date?fmtD(t.date):'No deadline')}</td>
      <td onclick="event.stopPropagation()" style="padding:6px 10px 6px 4px;text-align:right;white-space:nowrap">
        <div style="display:inline-flex;gap:5px;align-items:center">
          ${nstt(t.status)!=="Done"?`<button class="btn btn-ghost btn-sm" style="color:#dc2626;border-color:#fca5a5;padding:4px 10px;font-size:11px;min-width:58px;line-height:1;font-weight:700;white-space:nowrap;display:inline-flex;align-items:center;justify-content:center" onclick="cancelActionTask(${t.id})" title="Cancel task">Cancel</button>`:''}
        </div>
      </td>
    </tr>`}).join("");
}
function selectedVisibleTasks(){
  const visible=new Set(getVis().map(t=>t.id));
  selectedTaskIds.forEach(id=>{if(!visible.has(id))selectedTaskIds.delete(id);});
  return tasks.filter(t=>selectedTaskIds.has(t.id));
}
function toggleTaskSelected(id,checked){
  if(checked)selectedTaskIds.add(id);
  else selectedTaskIds.delete(id);
}
function toggleAllSelected(checked){
  getVis().forEach(t=>{
    if(checked)selectedTaskIds.add(t.id);
    else selectedTaskIds.delete(t.id);
  });
  renderMaster();
}
async function moveSelectedToWed(){
  const selected=selectedVisibleTasks();
  if(!selected.length){toast("Select at least one task");return;}
  selected.forEach(t=>t.wednesday=true);
  selectedTaskIds.clear();
  refreshAll();
  toast(`${selected.length} task${selected.length!==1?"s":""} sent to review`);
  await saveTasksToOneDrive();
}
async function deleteSelectedTasks(){
  const selected=selectedVisibleTasks();
  if(!selected.length){toast("Select at least one task");return;}
  if(!confirm(`Remove ${selected.length} selected task${selected.length!==1?"s":""} from the Action Log?`))return;
  const ids=new Set(selected.map(t=>t.id));
  tasks=tasks.filter(t=>!ids.has(t.id));
  selectedTaskIds.clear();
  buildTrackedSet();
  refreshAll();
  toast(`${selected.length} task${selected.length!==1?"s":""} removed`);
  await saveTasksToOneDrive();
}
async function togWed(id){
  const t=tasks.find(x=>x.id===id);if(!t)return;
  t.wednesday=!t.wednesday;
  toast(t.wednesday?(isWednesdayUser?"Added to Wednesday notes":"Added to Discussion Notes"):(isWednesdayUser?"Removed from Wednesday notes":"Removed from Discussion Notes"));
  syncBadges();renderMaster();await saveTasksToOneDrive();
}
