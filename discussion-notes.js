
// ============================================================
// WEDNESDAY / DISCUSSION NOTES
// ============================================================
function allSummaryBullets(summary){
  if(!summary)return[];
  return summary.split('\n').map(l=>l.replace(/^[•*▾▲◆-]\s*/,'').trim()).filter(Boolean);
}
function summaryFieldMap(summary){
  const map={};
  String(summary||'').split('\n').map(l=>l.replace(/^[•*▾▲◆]\s*/,'').trim()).filter(Boolean).forEach(line=>{
    const ci=line.indexOf(':');
    if(ci>0&&ci<40){
      const key=line.slice(0,ci).trim().toLowerCase();
      const val=line.slice(ci+1).trim();
      if(val&&!map[key])map[key]=val;
    }else if(!map.note){
      map.note=line;
    }
  });
  return map;
}
function firstSummaryField(fields,names){
  for(const n of names){
    const hit=Object.keys(fields).find(k=>k===n||k.includes(n));
    if(hit&&fields[hit])return fields[hit];
  }
  return '';
}
function openNotesProofReview(taskId){
  const task=tasks.find(t=>String(t.id)===String(taskId));
  if(!task||!task._proofNotif){toast('No proof is currently waiting for review');return;}
  if(typeof window.openTaskProofReview!=='function'){toast('Proof review is still loading — try again');return;}
  window.openTaskProofReview({
    appTaskId:String(task.id),title:task.title||task.emailSubject||'Task',
    recipientName:task.person||'',recipientEmail:task.email||'',
    proofInstructions:task.proofInstructions||''
  });
}
function renderWed(){
  const wt=tasks.filter(t=>t.wednesday);
  const total=wt.length+customNotes.length;
  document.getElementById('ws-n').textContent=total;
  document.getElementById('ws-p').textContent=wt.filter(t=>nstt(t.status)==='Pending'&&!isOverdueTask(t)).length;
  document.getElementById('ws-o').textContent=wt.filter(isOverdueTask).length;
  const el=document.getElementById('wed-list');
  if(!total){
    el.innerHTML=`<div class="empty-state"><div class="es-text">Nothing added yet</div><div class="es-sub">Use the ${isWednesdayUser?'+ Wed':'+ Discussion'} button in the Action Log to add tasks, or type a note above</div></div>`;
    return;
  }
  let html='';

  if(wt.length){
    // Group by department
    const groups={};
    wt.forEach(t=>{
      const dept=t.dept&&t.dept!=='Unknown'?t.dept:'Miscellaneous';
      if(!groups[dept])groups[dept]=[];
      groups[dept].push(t);
    });

    // Sort departments — Miscellaneous always last
    const depts=Object.keys(groups).sort((a,b)=>{
      if(a==='Miscellaneous')return 1;
      if(b==='Miscellaneous')return -1;
      return a.localeCompare(b);
    });

    // Department tiles navigation strip
    html+=`<div style="display:flex;gap:8px;flex-wrap:wrap;padding:4px 0 16px;border-bottom:1px solid var(--border);margin-bottom:16px">`;
    depts.forEach(dept=>{
      const dt=groups[dept];
      const dc=DEPT_COLORS[dept]||'#374151';
      const dn=dt.filter(t=>nstt(t.status)==='Done').length;
      const ov=dt.filter(isOverdueTask).length;
      const pct=dt.length?Math.round(dn/dt.length*100):0;
      html+=`<button onclick="document.getElementById('wed-dept-${dept.replace(/[^a-zA-Z0-9]/g,'-')}')?.scrollIntoView({behavior:'smooth',block:'nearest'})"
        style="display:flex;flex-direction:column;gap:3px;padding:10px 13px;border:1px solid var(--border);border-radius:var(--r);background:var(--white);cursor:pointer;transition:all .15s;text-align:left;min-width:100px;border-top:3px solid ${dc}"
        onmouseover="this.style.boxShadow='0 2px 12px rgba(0,0,0,.09)';this.style.borderColor='${dc}'"
        onmouseout="this.style.boxShadow='';this.style.borderColor='var(--border)';this.style.borderTopColor='${dc}'">
        <div style="font-size:12px;font-weight:700;color:var(--body);white-space:nowrap">${dept}</div>
        <div style="display:flex;gap:5px;align-items:center">
          <span style="font-size:17px;font-weight:800;color:${dc}">${dt.length}</span>
          <span style="font-size:10px;color:var(--muted)">${dt.length===1?'task':'tasks'}</span>
        </div>
        <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:1px">
          ${ov?`<span style="font-size:9.5px;color:#b91c1c;font-weight:600">${ov} overdue</span>`:''}
          ${dn?`<span style="font-size:9.5px;color:var(--forest);font-weight:600">${pct}% done</span>`:''}
        </div>
      </button>`;
    });
    html+=`</div>`;

    depts.forEach(dept=>{
      const deptTasks=groups[dept];
      const dcolor_val=DEPT_COLORS[dept]||'#374151';
      const safeId='wed-dept-'+dept.replace(/[^a-zA-Z0-9]/g,'-');
      html+=`<div class="wed-date-group" id="${safeId}">
        <div class="wed-date" style="display:flex;align-items:center;gap:8px">
          <span style="width:10px;height:10px;border-radius:50%;background:${dcolor_val};display:inline-block;flex-shrink:0"></span>
          ${dept} <span style="font-size:10px;opacity:.6;font-weight:500;text-transform:none;letter-spacing:0">${deptTasks.length} item${deptTasks.length!==1?'s':''}</span>
        </div>`;
      html+=deptTasks.map(t=>`
        <div class="wed-card">
          <div class="wed-card-head">
            <div class="wed-card-title">${emailSubject(t)}</div>
            <div style="display:flex;align-items:center;gap:5px;flex-shrink:0">
              ${sbadge(t)}
              <button class="rm-btn" onclick="removeFromWed(${t.id})" title="Remove from agenda">&#215;</button>
            </div>
          </div>
          <div class="wed-card-compact-body">
            <div class="wed-card-meta" style="margin-bottom:6px">
              <span style="font-weight:600;color:var(--body)">${t.person}</span>
              ${pBadge(t.priority||'Normal')}
              ${t.replyCount>1?`<span style="font-size:11px;color:var(--muted)">&bull; ${t.replyCount} msgs</span>`:''}
            </div>
            ${(()=>{const lines=allSummaryBullets(t.summary);return lines.length?`<ul style="margin:0 0 7px;padding-left:14px">${lines.map(l=>`<li style="font-size:12px;color:var(--sub);line-height:1.6;padding:1px 0">${escapeHtml(l)}</li>`).join('')}</ul>`:''})()}
            ${t.followup?`<div style="font-size:11.5px;color:var(--forest);font-weight:500;margin-bottom:7px;padding:5px 9px;background:var(--sage3);border-radius:5px">Note: ${t.followup}</div>`:''}
            <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:7px">
              ${taskEmailId(t)?`<button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();openTaskThread(${t.id})" style="font-size:10.5px;padding:3px 9px;display:flex;align-items:center;gap:4px"><svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke-width="2" stroke-linecap="round"/></svg>Open Thread</button>`:''}
            </div>
            <textarea class="ag-ta" placeholder="Meeting notes — what was discussed, decided or deferred..." id="wn-${t.id}" rows="1" style="min-height:32px">${t.meetingNote||''}</textarea>
            <div class="ag-foot" style="margin-top:6px">
              ${t._proofNotif?`<button class="btn btn-primary btn-xs" onclick="openNotesProofReview(${t.id})" style="font-size:10.5px">Review Proof</button>`:`<span title="Status is updated from the Tasks workflow">${sbadge(t)}</span>`}
              <button class="btn btn-ghost btn-xs" onclick="removeFromWed(${t.id})" style="margin-left:auto;font-size:10.5px">Remove</button>
            </div>
          </div>
        </div>`).join('');
      html+=`</div>`;
    });
  }

  // Custom notes section
  if(customNotes.length){
    html+=`<div class="wed-date-group">
      <div class="wed-date" style="display:flex;align-items:center;gap:8px">
        <span style="width:10px;height:10px;border-radius:50%;background:#6b7280;display:inline-block;flex-shrink:0"></span>
        ${isWednesdayUser?'Wednesday Review Notes':'Discussion Notes'} <span style="font-size:10px;opacity:.6;font-weight:500;text-transform:none;letter-spacing:0">${customNotes.length} note${customNotes.length!==1?'s':''}</span>
      </div>`;
    html+=customNotes.map((n,i)=>`
      <div class="wed-card">
        <div class="wed-card-head">
          <div class="wed-card-title">${n.text}</div>
          <div style="display:flex;align-items:center;gap:6px;flex-shrink:0">
            <span style="font-size:11px;color:var(--muted)">${n.time}</span>
            <button class="rm-btn" onclick="removeCN(${i})" title="Remove note">&#215;</button>
          </div>
        </div>
        <div class="wed-card-compact-body">
          <textarea class="ag-ta" placeholder="${isWednesdayUser?'Add Wednesday review notes here...':'Add discussion notes here...'}" id="cn-note-${i}" rows="1" style="min-height:32px">${n.meetingNote||''}</textarea>
        </div>
      </div>`).join('');
    html+=`</div>`;
  }

  el.innerHTML=html;
}



async function addCN(){const v=document.getElementById("cn-in").value.trim();if(!v)return;customNotes.unshift({text:v,time:new Date().toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"}),meetingNote:""});document.getElementById("cn-in").value="";renderWed();syncBadges();toast("Note added");await saveTasksToOneDrive();}
async function removeCN(i){const el=document.getElementById("cn-note-"+i);if(el)customNotes[i].meetingNote=el.value;customNotes.splice(i,1);renderWed();syncBadges();await saveTasksToOneDrive();}
async function removeFromWed(id){const t=tasks.find(x=>x.id===id);if(t)t.wednesday=false;renderWed();toast(isWednesdayUser?"Removed from Wednesday notes":"Removed from Discussion Notes");await saveTasksToOneDrive();}
async function clearWed(){tasks.forEach(t=>t.wednesday=false);customNotes=[];renderWed();syncBadges();toast("Cleared");await saveTasksToOneDrive();}
function dpegPrint(){
  const dateLabel=document.getElementById('print-date-label');
  const pageLabel=document.getElementById('print-page-label');
  const userLabel=document.getElementById('print-user-label');
  const now=new Date();
  const dateStr=now.toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
  const activePage=document.querySelector('.page.active');
  const pageTitle=activePage?.id==='page-wednesday'?(isWednesdayUser?'Wednesday Review':'Discussion Notes'):'Summary Report';
  const isSummary=activePage?.id==='page-summary';
  if(dateLabel)dateLabel.textContent=dateStr;
  if(pageLabel)pageLabel.textContent=pageTitle;
  if(userLabel&&currentUser)userLabel.textContent=`${currentUser.name} • ${currentUser.email} • Confidential Internal Report`;
  if(isSummary){
    buildSummaryPrintReport();
    document.body.classList.add('printing-summary');
  }
  window.print();
}
window.addEventListener('afterprint',()=>document.body.classList.remove('printing-summary'));
