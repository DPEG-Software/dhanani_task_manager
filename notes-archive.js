// Notes Archive
// Keeps the existing global API used by inline onclick handlers and navigation.

async function saveMeeting(){
  const wt=tasks.filter(t=>t.wednesday);
  wt.forEach(t=>{const el=document.getElementById("wn-"+t.id);if(el)t.meetingNote=el.value;});
  customNotes.forEach((n,i)=>{const el=document.getElementById("cn-note-"+i);if(el)n.meetingNote=el.value;});
  archives.unshift({id:Date.now(),date:new Date().toISOString().split("T")[0],
    tasks:wt.map(t=>({title:t.title,person:t.person,dept:t.dept,status:isOverdueTask(t)?"Overdue":nstt(t.status),note:t.meetingNote||""})),
    customNotes:customNotes.map(n=>({text:n.text,note:n.meetingNote||""}))
  });
  tasks.forEach(t=>{
    if(isOverdueTask(t)){
      t.wednesday=false;
      t.carryForwardFrom=t.carryForwardFrom||new Date().toISOString().split("T")[0];
      t.followup=t.meetingNote||t.followup||"Pending from Wednesday review.";
      delete t.meetingNote;
      return;
    }
    t.wednesday=false;
    delete t.meetingNote;
  });
  customNotes=[];renderWed();renderArc();syncBadges();
  await saveTasksToOneDrive();
  toast("Saved and archived.");
  nav("archive");
}

function renderArc(){
  const el=document.getElementById("arc-list");
  if(!archives.length){el.innerHTML=`<div class="empty-state"><div class="es-text">No archived meetings yet</div></div>`;return;}
  el.innerHTML=archives.map(a=>{
    const d=new Date(a.date);
    return `<div class="arc-item" onclick="showArcD(${a.id})">
      <div class="arc-date"><div class="arc-day">${d.getDate()}</div><div class="arc-mon">${d.toLocaleDateString("en-US",{month:"short"})}</div></div>
      <div style="flex:1"><div class="arc-ttl">${isWednesdayUser?"Wednesday Meeting":"Discussion Notes"} ${d.toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}</div><div class="arc-meta">${a.tasks.length} tasks &bull; ${(a.customNotes||[]).length} notes &bull; ${a.tasks.filter(t=>t.status==="Done").length} completed</div></div>
      <span style="color:var(--muted);font-size:18px">&#8250;</span>
    </div>`;
  }).join("");
}

function showArcD(id){
  const a=archives.find(x=>x.id===id);if(!a)return;
  const d=new Date(a.date);
  document.getElementById("arc-lv").style.display="none";
  document.getElementById("arc-dv").style.display="block";
  let idx=1;
  document.getElementById("arc-dc").innerHTML=`
    <div class="tbl-wrap">
      <div class="tbl-hdr" style="background:var(--ink)">
        <div><div class="tbl-title" style="color:#fff">${isWednesdayUser?"Wednesday Meeting":"Discussion Notes"} ${d.toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"})}</div>
        <div style="font-size:11.5px;color:rgba(210,223,212,.6);margin-top:2px">${a.tasks.length} tasks &bull; ${a.tasks.filter(t=>t.status==="Done").length} completed &bull; ${(a.customNotes||[]).length} notes</div></div>
      </div>
      ${a.tasks.map(t=>`<div style="padding:14px 18px;border-bottom:1px solid #f3f4f6">
        <div style="display:flex;align-items:center;gap:9px;margin-bottom:6px">
          <div style="width:20px;height:20px;border-radius:50%;background:var(--ink);color:#fff;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0">${idx++}</div>
          <div style="flex:1"><div style="font-size:13px;font-weight:700;color:var(--body)">${t.title}</div><div style="font-size:11.5px;color:var(--muted)">${t.person} &bull; ${t.dept}</div></div>
          ${sbadge(t)}
        </div>
        ${t.note?`<div style="font-size:12px;color:var(--sub);padding:8px 12px;background:#f9fafb;border-radius:var(--r);border-left:3px solid #e5e7eb;margin-top:4px;line-height:1.55">${t.note}</div>`:`<div style="font-size:12px;color:var(--muted);font-style:italic">No notes recorded</div>`}
      </div>`).join("")}
      ${(a.customNotes||[]).length?`<div style="padding:14px 18px;background:#f9fafb">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:10px">Notes</div>
        ${a.customNotes.map(n=>`<div style="margin-bottom:10px"><div style="font-size:13px;font-weight:600;color:var(--body);margin-bottom:3px">${typeof n==="string"?n:n.text}</div>${(typeof n==="object"&&n.note)?`<div style="font-size:12px;color:var(--sub);padding:6px 10px;background:#fff;border-radius:var(--r);border-left:3px solid #e5e7eb">${n.note}</div>`:""}</div>`).join("")}
      </div>`:""}`
    +`</div>`;
}

function backArc(){document.getElementById("arc-lv").style.display="block";document.getElementById("arc-dv").style.display="none";}
