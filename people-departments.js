// ============================================================
// PEOPLE
// ============================================================
function renderPplList(filter){
  const pp={};
  tasks.forEach(t=>{
    const key=staffKey(t.email,t.person);
    if(!isInternalEmail(t.email)&&!staffConfig[key])return;
    if(!pp[t.person])pp[t.person]={name:t.person,email:t.email||"",dept:personDept(t.email,t.person),tasks:[]};
    pp[t.person].tasks.push(t);
  });
  const list=Object.values(pp).filter(p=>!filter||p.name.toLowerCase().includes(filter.toLowerCase())||p.dept.toLowerCase().includes(filter.toLowerCase()));
  const el=document.getElementById("ppl-list");
  if(!list.length){el.innerHTML=`<div style="padding:14px;text-align:center;font-size:12px;color:var(--muted)">No results</div>`;return;}
  el.innerHTML=list.map(p=>{
    const dn=p.tasks.filter(t=>nstt(t.status)==="Done").length;
    const pe=p.tasks.filter(t=>nstt(t.status)!=="Done").length;
    const ov=p.tasks.filter(isOverdueTask).length;
    return `<div class="person-row${curPplSel===p.name?" sel":""}" onclick="selectPpl('${p.name}')">
      ${av(p.name,32)}
      <div style="flex:1;min-width:0">
        <div class="pr-nm">${p.name}${p.principal?` <span class="principal-tag">Principal</span>`:""}</div>
        <div class="pr-role">${p.role||p.dept}</div>
        <div class="pr-stats"><span class="pr-stat pending">${pe} pending</span><span class="pr-stat done">${dn} done</span>${ov?`<span class="pr-stat" style="color:#b91c1c">${ov} overdue</span>`:""}</div>
      </div>
    </div>`;
  }).join("");
}
function filterPpl(v){curPplFilter=v;renderPplList(v);}
function setDirectoryMode(mode){
  directoryMode=mode;
  document.getElementById("dir-people-btn")?.classList.toggle("active",mode==="people");
  document.getElementById("dir-dept-btn")?.classList.toggle("active",mode==="departments");
  document.getElementById("dir-people-pane")?.classList.toggle("active",mode==="people");
  document.getElementById("dir-dept-pane")?.classList.toggle("active",mode==="departments");
  if(mode==="people")renderPplList(curPplFilter);
  if(mode==="departments")renderDeptList();
}
function selectPpl(name){
  curPplSel=name;renderPplList(curPplFilter);
  const pt=tasks.filter(t=>t.person===name);
  const pdb=findPersonByName(name);
  const primaryTask=pt[0]||{};
  const dept=personDept(primaryTask.email,name);
  const email=pdb?.email||primaryTask.email||"";
  const dn=pt.filter(t=>nstt(t.status)==="Done").length;
  const pct=pt.length?Math.round(dn/pt.length*100):0;
  document.getElementById("ppl-right").innerHTML=`
    <div style="display:flex;align-items:center;gap:14px;margin-bottom:18px;padding-bottom:16px;border-bottom:1px solid var(--border)">
      ${av(name,44)}<div style="flex:1">
        <div style="font-size:17px;font-weight:700;color:var(--body);margin-bottom:2px">${name}</div>
        <div style="font-size:12px;color:var(--muted)">${pdb?.role||""} &bull; ${dept} &bull; <span id="ppl-email-label">${email||"No email saved"}</span></div>
        <div style="display:flex;gap:12px;margin-top:5px"><span style="font-size:12px;font-weight:600;color:#15803d">${dn} completed</span><span style="font-size:12px;font-weight:600;color:${pt.filter(isOverdueTask).length?"#b91c1c":"var(--muted)"}">${pt.filter(t=>nstt(t.status)!=="Done").length} open</span></div>
      </div>
      <div style="text-align:right"><div style="font-size:22px;font-weight:700;color:var(--body)">${pct}%</div><div class="prog-track" style="width:72px;margin-top:4px"><div class="prog-fill" style="width:${pct}%"></div></div></div>
    </div>
    ${pt.length?`<table style="width:100%;border-collapse:collapse"><thead><tr>${["Task","Status","Priority","Date","Notes"].map(h=>`<th style="padding:8px 10px;text-align:left;font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);background:#f9fafb;border-bottom:1px solid var(--border)">${h}</th>`).join("")}</tr></thead><tbody>${pt.map(t=>`<tr style="border-bottom:1px solid #f3f4f6;cursor:pointer" onclick="openDetail(${t.id})"><td style="padding:10px"><div style="font-size:13px;font-weight:600;color:var(--body)">${t.title}</div>${threadLinkButton(t)}</td><td style="padding:10px">${sbadge(t)}</td><td style="padding:10px">${pBadge(t.priority)}</td><td style="padding:10px;font-size:12px;color:var(--muted);white-space:nowrap">${fmtD(t.date)}</td><td style="padding:10px;font-size:12px;color:var(--muted);font-style:${t.followup?"normal":"italic"}">${t.followup||"None"}</td></tr>`).join("")}</tbody></table>`:
    `<div class="empty-state"><div class="es-text">No tasks assigned yet</div></div>`}`;
}
async function updatePersonDept(name,dept){
  const t=tasks.find(x=>x.person===name);
  if(!t)return;
  saveStaffDeptForTask(t,dept);
  refreshAll();
  toast("Department saved");
  await saveTasksToOneDrive();
}
function renderDeptList(){
  const dp={};allDepartments().forEach(d=>dp[d]=[]);
  tasks.forEach(t=>{if(dp[t.dept])dp[t.dept].push(t);else dp[t.dept]=[t];});
  document.getElementById("dept-list").innerHTML=Object.entries(dp).map(([d,dt])=>{
    const dn=dt.filter(t=>nstt(t.status)==="Done").length;
    const ov=dt.filter(isOverdueTask).length;
    const initials=d.split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2);
    return `<div class="dept-row${curDeptSel===d?" sel":""}" onclick="selectDept('${d}')">
      <div class="dept-icon" style="background:${dcolor(d)}">${initials}</div>
      <div><div class="dr-nm">${d}</div><div class="dr-cnt">${dt.length} tasks &bull; ${dn} done${ov?` &bull; <span style="color:#b91c1c">${ov} overdue</span>`:""}</div></div>
    </div>`;
  }).join("");
}
function selectDept(dept){
  curDeptSel=dept;renderDeptList();
  const dt=tasks.filter(t=>t.dept===dept);
  const dn=dt.filter(t=>nstt(t.status)==="Done").length;
  const pct=dt.length?Math.round(dn/dt.length*100):0;
  document.getElementById("dept-right").innerHTML=`
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:18px;padding-bottom:16px;border-bottom:1px solid var(--border)">
      <div style="width:42px;height:42px;border-radius:8px;background:${dcolor(dept)};color:#fff;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700">${dept.split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2)}</div>
      <div style="flex:1"><div style="font-size:17px;font-weight:700;color:var(--body);margin-bottom:1px">${dept}</div><div style="font-size:12px;color:var(--muted)">${dt.length} tasks &bull; ${pct}% complete</div></div>
      <div style="text-align:right"><div style="font-size:22px;font-weight:700;color:var(--body)">${pct}%</div><div class="prog-track" style="width:72px;margin-top:4px"><div class="prog-fill" style="width:${pct}%"></div></div></div>
    </div>
    ${dt.length?`<table style="width:100%;border-collapse:collapse"><thead><tr>${["Task","Assigned To","Status","Date"].map(h=>`<th style="padding:8px 10px;text-align:left;font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);background:#f9fafb;border-bottom:1px solid var(--border)">${h}</th>`).join("")}</tr></thead><tbody>${dt.map(t=>`<tr style="border-bottom:1px solid #f3f4f6;cursor:pointer" onclick="openDetail(${t.id})"><td style="padding:10px"><div style="font-size:13px;font-weight:600;color:var(--body)">${t.title}</div>${threadLinkButton(t)}</td><td style="padding:10px"><div style="display:flex;align-items:center;gap:7px">${av(t.person,24)}<span style="font-size:12.5px">${t.person}</span></div></td><td style="padding:10px">${sbadge(t)}</td><td style="padding:10px;font-size:12px;color:var(--muted)">${fmtD(t.date)}</td></tr>`).join("")}</tbody></table>`:
    `<div class="empty-state"><div class="es-text">No tasks yet</div></div>`}`;
}
