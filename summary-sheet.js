// Summary Sheet
// Reporting and CSV export functions, kept global for the current static app.

function summaryWeekLabel(v){
  return v==="all"?"All Time":(v==="0"?"This Week":(v==="1"?"Last Week":`${v} Weeks Ago`));
}

function getSummaryFilters(){
  const w=document.getElementById("sum-w").value,dept=document.getElementById("sum-d").value,status=document.getElementById("sum-s").value;
  return {w,dept,status};
}

function getSummaryFilteredTasks(){
  const {w,dept,status}=getSummaryFilters();
  return tasks.filter(t=>{
    const wm=w==="all"||t.weekOffset===parseInt(w)||(isOpenTask(t)&&t.weekOffset<=parseInt(w));
    const stm=status==="all"||(status==="Overdue"?isOverdueTask(t):nstt(t.status)===status);
    return (w==="all"?true:wm)&&(dept==="all"||t.dept===dept)&&stm;
  });
}

function buildSummaryPrintReport(){
  const el=document.getElementById("summary-print-report");
  if(!el)return;
  const {w,dept,status}=getSummaryFilters();
  const list=getSummaryFilteredTasks();
  const done=list.filter(t=>nstt(t.status)==="Done").length;
  const open=list.length-done;
  const high=list.filter(t=>(t.priority||"").toLowerCase()==="high").length;
  const printed=new Date().toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"});
  el.innerHTML=`
    <div class="summary-print-meta">
      <div class="summary-print-box"><div class="summary-print-label">Prepared For</div><div class="summary-print-value">${escapeHtml(currentUser?.name||"DPEG User")}</div></div>
      <div class="summary-print-box"><div class="summary-print-label">Email</div><div class="summary-print-value">${escapeHtml(currentUser?.email||"")}</div></div>
      <div class="summary-print-box"><div class="summary-print-label">Printed</div><div class="summary-print-value">${escapeHtml(printed)}</div></div>
      <div class="summary-print-box"><div class="summary-print-label">Tasks</div><div class="summary-print-value">${list.length}</div></div>
    </div>
    <p class="summary-print-note">Filters applied: ${escapeHtml(summaryWeekLabel(w))}; ${escapeHtml(dept==="all"?"All Departments":dept)}; ${escapeHtml(status==="all"?"All Statuses":status)}. Open: ${open}. Completed: ${done}. High priority: ${high}.</p>
    <table>
      <thead><tr><th>#</th><th>Task</th><th>Assigned To</th><th>Department</th><th>Status</th><th>Priority</th><th>Date</th><th>Notes</th></tr></thead>
      <tbody>${list.length?list.map((t,i)=>`
        <tr>
          <td>${i+1}</td>
          <td><strong>${escapeHtml(emailSubject(t))}</strong></td>
          <td>${escapeHtml(t.person||"")}</td>
          <td>${escapeHtml(t.dept||"")}</td>
          <td>${escapeHtml(isOverdueTask(t)?"Overdue":nstt(t.status))}</td>
          <td>${escapeHtml(t.priority||"Normal")}</td>
          <td>${escapeHtml(fmtD(t.date))}</td>
          <td>${escapeHtml(t.followup||t.summary||"")}</td>
        </tr>`).join(""):`<tr><td colspan="8" style="text-align:center;color:#6b7280;padding:18pt">No tasks match the selected filters.</td></tr>`}</tbody>
    </table>`;
}

function renderSum(){
  let list=getSummaryFilteredTasks();
  document.getElementById("sum-cnt").textContent=`${list.length} task${list.length!==1?"s":""}`;
  document.getElementById("sum-tbody").innerHTML=list.map((t,i)=>`
    <tr><td style="color:var(--muted)">${i+1}</td><td><div style="font-weight:600;color:var(--body)">${t.title}</div></td><td>${t.person}</td>
    <td><span class="dept-pill"><span class="dept-dot" style="background:${dcolor(t.dept)}"></span>${t.dept}</span></td>
    <td>${sbadge(t)}</td><td>${pBadge(t.priority)}</td>
    <td style="color:var(--muted);font-size:12px;white-space:nowrap">${fmtD(t.date)}</td>
    <td style="color:var(--muted);font-size:12px;font-style:${t.followup?"normal":"italic"}">${t.followup||"None"}</td></tr>`).join("");
}

function dlCSV(){
  let list=getSummaryFilteredTasks();
  const rows=[["Number","Task","Assigned To","Email","Department","Status","Priority","Date","Notes"],...list.map((t,i)=>[i+1,t.title,t.person,t.email||"",t.dept,isOverdueTask(t)?"Overdue":nstt(t.status),t.priority||"",fmtD(t.date),t.followup||""])];
  const csv=rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  const a=document.createElement("a");a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(csv);
  a.download=`DPEG_Tasks_${new Date().toISOString().split("T")[0]}.csv`;a.click();toast("CSV downloaded");
}
