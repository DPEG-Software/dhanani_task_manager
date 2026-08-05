// ============================================================
// NAVIGATION
// ============================================================
const PT={dashboard:"Dashboard",master:"Action Log",people:"People & Departments",wednesday:isWednesdayUser?"Wednesday Review":"Discussion Notes",archive:"Notes Archive",summary:"Summary Sheet",outlook:"Outlook",notifications:"Notifications",tasks:"Tasks"};
const PS={dashboard:"Real-time overview of task activity",master:"All dispatched tasks",people:"Team members and departments",wednesday:"Weekly agenda and notes",archive:"Past meeting records",summary:"Task report and CSV export",outlook:"Send and receive emails. Tasks are created automatically for emails sent to DPEG staff.",notifications:"Task completion alerts and approvals",tasks:"Tasks assigned to you and by you"};
function nav(name,mode){
  if(name==="departments"){name="people";mode="departments";}
  document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".ni").forEach(n=>n.classList.remove("active"));
  document.getElementById("page-"+name)?.classList.add("active");
  document.getElementById("tb-title").textContent=PT[name]||"";
  document.getElementById("tb-sub").textContent=PS[name]||"";
  document.querySelectorAll(".ni").forEach(btn=>{
    const t=btn.textContent.trim().toLowerCase();
    if((name==="dashboard"&&t.startsWith("dash"))||(name==="master"&&t.startsWith("action"))||(name==="people"&&(t.startsWith("people")||t.startsWith("dept")))||(name==="wednesday"&&(t.startsWith("wed")||t.startsWith("disc")))||(name==="archive"&&t.startsWith("notes archive"))||(name==="summary"&&t.startsWith("sum"))||(name==="outlook"&&t.startsWith("out"))||(name==="notifications"&&t.startsWith("notif"))||(name==="tasks"&&t.startsWith("tasks")))
      btn.classList.add("active");
  });
  curSearch="";document.getElementById("srch").value="";
  if(name==="master")renderMaster();
  if(name==="tasks")renderMyTasks();
  if(name==="people"){directoryMode=mode||"people";setDirectoryMode(directoryMode);if(directoryMode==="people"){renderPplList();curPplSel=null;document.getElementById("ppl-srch").value="";curPplFilter="";}if(directoryMode==="departments"){renderDeptList();curDeptSel=null;}}
  if(name==="wednesday")renderWed();
  if(name==="archive")renderArc();
  if(name==="dashboard"){syncPulse();renderCharts();renderActivity();}
  if(name==="summary")renderSum();
  if(name==="outlook"){initOlPanels();fetchOutlookContacts();loadFolder("inbox");}
  if(name==="notifications"){renderNotifications();}
  // Prevent outer body scrollbar while Outlook fills the full viewport
  const mainEl=document.querySelector('.main');
  if(mainEl)mainEl.style.overflowY=name==='outlook'?'hidden':'';
}


function syncBadges(){
  const open=tasks.filter(t=>nstt(t.status)!=="Done").length;
  const actionWeekOpen=tasks.filter(t=>nstt(t.status)!=="Done"&&taskWeekOffset(t)===curWeek).length;
  const high=tasks.filter(t=>String(t.priority||"Normal").toLowerCase()==="high"&&nstt(t.status)!=="Done").length;
  const wed=tasks.filter(t=>t.wednesday).length+customNotes.length;
  const dn=tasks.filter(t=>nstt(t.status)==="Done").length;
  const pct=tasks.length?Math.round(dn/tasks.length*100):0;
  document.getElementById("nb-m").textContent=actionWeekOpen;
  document.getElementById("nb-w").textContent=wed;
  document.getElementById("d-tot").textContent=tasks.length;
  document.getElementById("d-high").textContent=high;
  document.getElementById("d-wed").textContent=wed;
  document.getElementById("d-rate").textContent=pct+"%";
  document.getElementById("d-rate-sub").textContent=`${dn} of ${tasks.length} tasks completed`;
  const pctL=document.getElementById("d-pct-l");if(pctL)pctL.textContent=pct+"%";
  syncPulse();
}
function refreshAll(){
  syncBadges();
  const a=document.querySelector(".page.active");
  if(a?.id==="page-master")renderMaster();
  if(a?.id==="page-people"&&directoryMode==="people"&&curPplSel)selectPpl(curPplSel);
  if(a?.id==="page-people"&&directoryMode==="departments"&&curDeptSel)selectDept(curDeptSel);
  if(a?.id==="page-wednesday")renderWed();
  if(a?.id==="page-dashboard"){renderCharts();renderActivity();}
  if(a?.id==="page-summary")renderSum();
}

function toggleAccountMenu(e){
  e?.stopPropagation();
  document.getElementById("acct-menu")?.classList.toggle("open");
}
function closeAccountMenu(){document.getElementById("acct-menu")?.classList.remove("open");}
document.addEventListener("click",e=>{if(!e.target.closest("#acct-menu")&&!e.target.closest("#acct-btn"))closeAccountMenu();});

function toast(msg,dur=2600){
  const t=document.getElementById("toast");
  t.textContent=msg;t.classList.add("show");
  clearTimeout(t._t);t._t=setTimeout(()=>t.classList.remove("show"),dur);
}
