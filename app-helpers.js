// ============================================================
// HELPERS
// ============================================================
function avc(name) { return DEPT_COLORS[findPersonByName(name)?.dept] || "#374151"; }
function dcolor(dept) { return DEPT_COLORS[dept] || "#374151"; }
function allDepartments(){
  const seen=new Set();
  const list=[];
  [...DEPARTMENTS,...(customDepartments||[])].forEach(d=>{
    const name=String(d||'').trim();
    const key=name.toLowerCase();
    if(name&&!seen.has(key)){seen.add(key);list.push(name);}
  });
  Object.values(staffConfig||{}).forEach(p=>{
    const name=String(p?.dept||'').trim();
    const key=name.toLowerCase();
    if(name&&!seen.has(key)){seen.add(key);list.push(name);}
  });
  tasks.forEach(t=>{
    const name=String(t?.dept||'').trim();
    const key=name.toLowerCase();
    if(name&&!seen.has(key)){seen.add(key);list.push(name);}
  });
  return list;
}
function ini(n) { return n.split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2); }
function fmtD(d) {
  if(!d)return "";
  const p=String(d).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(!p)return "";
  const dt=new Date(Number(p[1]),Number(p[2])-1,Number(p[3]));
  return dt.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});
}
function nstt(s) { return s==="In Progress"?"Pending":(s||"Pending"); }
// Overdue is computed from the due date, not a stored status string — nothing
// in the app ever actually sets status to the literal "Overdue" value, so
// every nstt(t.status)==="Overdue" check was silently always false.
function isOverdueTask(t) {
  if (nstt(t.status) === "Done" || !t.date) return false;
  // Parse as a local calendar date, not new Date(t.date) — that parses
  // plain "YYYY-MM-DD" as UTC midnight, which in US timezones lands hours
  // before local midnight and wrongly flags tasks due *today* as overdue.
  const m = String(t.date).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return false;
  const due = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const startOfToday = new Date(); startOfToday.setHours(0, 0, 0, 0);
  return due < startOfToday;
}
function normEmail(email) { return String(email||"").trim().toLowerCase(); }
function staffKey(email,name) { return normEmail(email) || String(name||"").trim().toLowerCase(); }
function findPersonByEmail(email) {
  const e=normEmail(email);
  const configured=Object.values(staffConfig).find(x=>normEmail(x.email)===e);
  const personal=Object.values(userContacts||{}).find(x=>normEmail(x.email)===e);
  return configured || personal || null;
}
function findPersonByName(name) {
  const n=String(name||"").trim().toLowerCase();
  return staffConfig[n] || Object.values(staffConfig).find(x=>String(x.name||"").trim().toLowerCase()===n) || Object.values(userContacts||{}).find(x=>String(x.name||"").trim().toLowerCase()===n) || null;
}
function isInternalEmail(email) {
  const e=normEmail(email);
  return e.endsWith("@dhananipeg.com") || e.endsWith("@dpeg.com");
}
function configuredDept(email,name) {
  const key = staffKey(email,name);
  const exact=String((key&&staffConfig[key]?.dept)||'').trim();
  if(hasAssignedDepartment(exact))return exact;
  const normalized=normEmail(email);
  const normalizedName=String(name||'').trim().toLowerCase();
  const matches=Object.values(staffConfig||{}).filter(person=>
    (normalized&&normEmail(person?.email)===normalized)||
    (!normalized&&normalizedName&&String(person?.name||'').trim().toLowerCase()===normalizedName)
  );
  const assigned=matches.find(person=>hasAssignedDepartment(person?.dept));
  return String(assigned?.dept||matches[0]?.dept||exact||'').trim();
}
function personDept(email,name) {
  const configured = configuredDept(email,name);
  if (configured) return configured;
  const p = email ? findPersonByEmail(email) : findPersonByName(name);
  return p?.dept || "Needs Department";
}
function deptOptions(selected) {
  const base=allDepartments();
  const list = base.some(d=>d===selected) || !selected ? base : [selected,...base];
  return list.map(d=>`<option${d===selected?" selected":""}>${d}</option>`).join("");
}
function saveStaffDeptForTask(t,dept) {
  const key = staffKey(t.email,t.person);
  if (!key) return;
  staffConfig[key] = { ...(staffConfig[key]||{}), name:t.person, email:t.email||"", dept };
  tasks.forEach(x=>{
    if ((t.email && normEmail(x.email)===normEmail(t.email)) || (!t.email && x.person===t.person)) x.dept = dept;
  });
}
function hasAssignedDepartment(dept){
  const value=String(dept||'').trim();
  return !!value&&!['Needs Department','Unknown','Outside DPEG'].includes(value);
}
function applyDepartmentAssignment(email,name,dept){
  const normalized=normEmail(email);
  const key=staffKey(normalized,name);
  if(!key)return;
  const existing=(normalized&&findPersonByEmail(normalized))||findPersonByName(name)||{};
  staffConfig[key]={...(staffConfig[key]||{}),name:name||existing.name||normalized,email:normalized,dept,role:staffConfig[key]?.role||existing.role||''};
  tasks.forEach(t=>{if(normalized&&normEmail(t.email)===normalized)t.dept=dept;});
}
function requestDepartmentSelection(person,currentDept='',proposedDept=''){
  return new Promise(resolve=>{
    const modal=document.getElementById('mo-department-required');
    const select=document.getElementById('department-required-select');
    const title=document.getElementById('department-required-title');
    const personEl=document.getElementById('department-required-person');
    const message=document.getElementById('department-required-message');
    const save=document.getElementById('department-required-save');
    const cancel=document.getElementById('department-required-cancel');
    if(!modal||!select){resolve(null);return;}
    const changing=hasAssignedDepartment(currentDept)&&hasAssignedDepartment(proposedDept)&&currentDept!==proposedDept;
    title.textContent=changing?'Change Department?':'Select Department';
    personEl.textContent=[person.name,person.email].filter(Boolean).join(' • ');
    message.innerHTML=changing
      ?`This email is currently assigned to <strong>${escapeHtml(currentDept)}</strong>. Do you want to change it to <strong>${escapeHtml(proposedDept)}</strong>? The change will apply across the entire app.`
      :'Select the department for this email. The selection will stay the same for future emails and tasks until someone changes it.';
    const departments=allDepartments().filter(hasAssignedDepartment);
    select.innerHTML=departments.map(d=>`<option${d===(proposedDept||currentDept)?' selected':''}>${escapeHtml(d)}</option>`).join('');
    select.style.display=changing?'none':'';
    save.textContent=changing?'Yes, Change Department':'Save Department';
    cancel.textContent=changing?'No, Keep Current':'Cancel';
    let settled=false;
    const finish=value=>{if(settled)return;settled=true;modal.classList.remove('open');save.onclick=null;cancel.onclick=null;resolve(value);};
    save.onclick=()=>finish(changing?proposedDept:select.value);
    cancel.onclick=()=>finish(changing?currentDept:null);
    modal.onclick=e=>{if(e.target===modal)finish(null);};
    modal.classList.add('open');
  });
}
async function ensureDepartmentForPerson(person,proposedDept=''){
  const email=normEmail(person?.email||'');
  const name=String(person?.name||email.split('@')[0]||'').trim();
  if(!email){
    if(name&&name!=='Unassigned'){
      toast('Select the person from Microsoft contacts so their exact email can be saved');
      return null;
    }
    return person?.dept||'Needs Department';
  }
  if(!isInternalEmail(email))return person?.dept||'Outside DPEG';
  const current=configuredDept(email,name);
  let selected=current;
  if(!hasAssignedDepartment(current))selected=await requestDepartmentSelection({name,email},'',hasAssignedDepartment(proposedDept)?proposedDept:'');
  else if(hasAssignedDepartment(proposedDept)&&proposedDept!==current)selected=await requestDepartmentSelection({name,email},current,proposedDept);
  if(!selected)return null;
  if(selected!==current){
    const saved=await saveSharedDepartmentAssignment(email,name,selected);
    if(!saved){toast(window.lastDepartmentSaveError||'Department could not be saved. Please try again.');return null;}
    applyDepartmentAssignment(email,name,selected);
    initSelects();
    // The shared registry and D1 assignment rows are authoritative. Do not
    // leave the underlying edit window hanging while OneDrive performs a
    // secondary personal backup.
    saveTasksToOneDrive().catch(err=>console.warn('Department backup save skipped:',err.message));
    toast(`${name}'s department changed to ${selected} across the app`);
  }
  person.dept=selected;
  return selected;
}
async function ensureDepartmentsForPeople(people,proposedDept=''){
  for(const person of people){
    const dept=await ensureDepartmentForPerson(person,people.length===1?proposedDept:'');
    if(dept===null)return false;
  }
  return true;
}
function saveStaffEmail(name,email){
  const key=String(name||"").trim().toLowerCase();
  if(!key)return;
  const existing=findPersonByName(name)||{};
  staffConfig[key]={...(staffConfig[key]||{}),name,email,dept:staffConfig[key]?.dept||existing.dept||"Needs Department",role:existing.role||""};
  tasks.forEach(t=>{if(t.person===name)t.email=email;});
}
function departmentAssignmentContacts(){
  const map=new Map();
  const add=p=>{
    if(!p)return;
    const name=String(p.name||p.displayName||p.fullName||"").trim().replace(/\s+/g," ");
    const email=normEmail(p.email||p.mail||p.userPrincipalName||p.address||"");
    if(!name&&!email)return;
    const key=email||name.toLowerCase();
    const existing=map.get(key)||{};
    const sharedDept=personDept(email,name);
    const incomingDept=String(p.dept||p.department||'').trim();
    map.set(key,{
      name: name||existing.name||email,
      email: email||existing.email||"",
      dept: hasAssignedDepartment(sharedDept)
        ?sharedDept
        :(hasAssignedDepartment(existing.dept)?existing.dept:(hasAssignedDepartment(incomingDept)?incomingDept:'Needs Department')),
      role: String(p.role||p.jobTitle||p.title||existing.role||"").trim()
    });
  };
  Object.values(staffConfig||{}).forEach(add);
  Object.values(userContacts||{}).forEach(add);
  (outlookContacts||[]).forEach(add);
  (window.companyDirectoryContacts||[]).forEach(add);
  tasks.forEach(t=>add({name:t.person,email:t.email,dept:t.dept}));
  // Keep the shared admin account discoverable by its employee name even
  // when another user's personal Outlook directory has not cached it yet.
  add({name:'Nikhil Kumar',email:'propertymanagement2@dhananipeg.com',dept:'Property Management',role:'Admin'});
  add({name:'Aishwarya Sai',email:'systemmanager1@dhananipeg.com',dept:'Software Development',role:'System Manager'});
  return [...map.values()].sort((a,b)=>String(a.name||a.email).localeCompare(String(b.name||b.email)));
}
function setDeptAssignDepartment(dept){
  const sel=document.getElementById('dept-assign-dept');
  if(!sel)return;
  const val=String(dept||"").trim()||"Needs Department";
  if(![...sel.options].some(o=>o.value===val)){
    const opt=document.createElement('option');
    opt.textContent=val;
    opt.value=val;
    sel.insertBefore(opt,sel.firstChild);
  }
  sel.value=val;
}
function fillDeptAssignContact(p){
  if(!p)return;
  const nameEl=document.getElementById('dept-assign-name');
  const emailEl=document.getElementById('dept-assign-email');
  if(nameEl)nameEl.value=p.name||"";
  if(emailEl)emailEl.value=p.email||"";
  setDeptAssignDepartment(p.dept||personDept(p.email,p.name));
}
function showDeptAssignAC(source){
  const ac=document.getElementById('dept-assign-ac');
  if(!ac)return;
  const nameEl=document.getElementById('dept-assign-name');
  const emailEl=document.getElementById('dept-assign-email');
  const token=String((source==='email'?emailEl?.value:nameEl?.value)||"").trim().toLowerCase();
  if(!token){ac.style.display='none';return;}
  const matches=departmentAssignmentContacts().filter(p=>{
    const hay=[p.name,p.email,p.dept,p.role].map(x=>String(x||"").toLowerCase());
    return hay.some(x=>x.includes(token));
  }).slice(0,10);
  window._deptAssignMatches=matches;
  if(!matches.length){ac.style.display='none';return;}
  ac.innerHTML=matches.map((p,i)=>`
    <div class="compose-ac-item"
      onmousedown="event.preventDefault();selectDeptAssignAC(${i})"
      onmouseover="document.querySelectorAll('#dept-assign-ac .compose-ac-item').forEach(x=>x.classList.remove('ac-focused'));this.classList.add('ac-focused')">
      ${av(p.name||p.email||'?',28)}
      <div style="flex:1;min-width:0">
        <div class="compose-ac-name">${escapeHtml(p.name||p.email)}</div>
        <div class="compose-ac-email">${escapeHtml(p.email||'')}</div>
        ${p.dept||p.role?`<div class="compose-ac-role">${escapeHtml(p.dept||p.role)}</div>`:''}
      </div>
    </div>`).join('');
  ac.style.display='block';
}
function selectDeptAssignAC(index){
  const p=(window._deptAssignMatches||[])[index];
  if(!p)return;
  fillDeptAssignContact(p);
  const ac=document.getElementById('dept-assign-ac');
  if(ac)ac.style.display='none';
}
function completeDeptAssignFromText(){
  const nameEl=document.getElementById('dept-assign-name');
  const emailEl=document.getElementById('dept-assign-email');
  const name=String(nameEl?.value||"").trim().toLowerCase();
  const email=normEmail(emailEl?.value||"");
  if(!name&&!email)return null;
  const exact=departmentAssignmentContacts().find(p=>
    (email&&normEmail(p.email)===email) ||
    (name&&String(p.name||"").trim().toLowerCase()===name)
  );
  if(exact)fillDeptAssignContact(exact);
  return exact||null;
}
function hideDeptAssignAC(){
  setTimeout(()=>{completeDeptAssignFromText();const ac=document.getElementById('dept-assign-ac');if(ac)ac.style.display='none';},160);
}
function deptAssignACNav(e){
  const ac=document.getElementById('dept-assign-ac');
  if(!ac||ac.style.display==='none')return;
  const items=ac.querySelectorAll('.compose-ac-item');if(!items.length)return;
  const focused=ac.querySelector('.ac-focused');
  if(e.key==='ArrowDown'){e.preventDefault();const next=focused?focused.nextElementSibling||items[0]:items[0];items.forEach(i=>i.classList.remove('ac-focused'));if(next)next.classList.add('ac-focused');}
  else if(e.key==='ArrowUp'){e.preventDefault();const prev=focused?focused.previousElementSibling||items[items.length-1]:items[items.length-1];items.forEach(i=>i.classList.remove('ac-focused'));if(prev)prev.classList.add('ac-focused');}
  else if((e.key==='Enter'||e.key==='Tab')&&focused){e.preventDefault();focused.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));}
  else if(e.key==='Escape'){ac.style.display='none';}
}
async function saveDepartmentAssignmentSetting(){
  if(!isAdmin()){toast('Admin access only');return;}
  // Read the admin's selection before resolving typed contact text. Contact
  // resolution fills the person's existing department into the dropdown;
  // doing it first silently replaced a newly selected value (for example,
  // Maintenance) with the old one (commonly Investor Relations).
  const selectedDept=String(document.getElementById('dept-assign-dept')?.value||'').trim();
  if(!document.getElementById('dept-assign-email')?.value)completeDeptAssignFromText();
  const name=String(document.getElementById('dept-assign-name')?.value||'').trim().replace(/\s+/g,' ');
  const email=String(document.getElementById('dept-assign-email')?.value||'').trim().toLowerCase();
  const dept=selectedDept;
  if(!name&&!email){toast('Enter a person name or email');return;}
  if(email&&!email.includes('@')){toast('Enter a valid email');return;}
  if(!dept){toast('Select a department');return;}
  const existing=(email&&findPersonByEmail(email))||findPersonByName(name)||{};
  const finalName=name||existing.name||email.split('@')[0];
  const finalEmail=normEmail(email||existing.email||'');
  if(!finalEmail){toast('Select a Microsoft contact with an email address');return;}
  const previousDept=configuredDept(finalEmail,finalName);
  if(hasAssignedDepartment(previousDept)&&previousDept!==dept){
    const confirmed=await requestDepartmentSelection({name:finalName,email:finalEmail},previousDept,dept);
    if(confirmed===previousDept){
      setDeptAssignDepartment(previousDept);
      renderDepartmentSettingsList();
      toast(`Kept ${finalName} in ${previousDept}`);
      return;
    }
  }
  // Save just this email mapping. The previous bulk save could include a
  // duplicate stale Outlook contact later in the object and overwrite the
  // department the admin had just selected.
  const sharedSaved=await saveSharedDepartmentAssignment(finalEmail,finalName,dept);
  if(!sharedSaved){toast(window.lastDepartmentSaveError||'Department could not be saved. Please try again.');return;}
  applyDepartmentAssignment(finalEmail,finalName,dept);
  initSelects();
  if(directoryMode==='people')renderPplList(curPplFilter);
  if(directoryMode==='departments')renderDeptList();
  renderDepartmentSettingsList();
  syncBadges();
  if(document.getElementById("page-master")?.classList.contains("active"))renderMaster();
  if(document.getElementById("page-dashboard")?.classList.contains("active")){syncPulse();renderCharts();renderActivity();}
  await saveTasksToOneDrive();
  toast(`Department saved for ${finalName} across the app`);
}
function emailSubject(t){return t.emailSubject||t.subject||t.title||"(no subject)";}
function taskEmailId(t){return t.lastMessageId||t.emailId||"";}
function emailLinkButton(t,label="Open Full Email"){
  const id=taskEmailId(t);
  return id?`<button class="email-link-btn" onclick="event.stopPropagation();openTaskEmail('${id}')">${label}</button>`:"";
}
function formatSummaryHTML(summary){
  if(!summary)return '<span class="sum-empty">No summary available.</span>';
  const lines=summary.split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const rows=lines.map(line=>{
    const inner=line.replace(/^(?:•|â€¢)\s*/,'').trim();
    const ci=inner.indexOf(':');
    if(ci>0&&ci<30){
      const lbl=inner.slice(0,ci).trim();
      const val=inner.slice(ci+1).trim();
      const key=lbl.toLowerCase();
      const cls=key.includes('action')?'s-action':key.includes('latest')?'s-facts':'';
      return `<div class="sum-row${cls?' '+cls:''}"><span class="lbl">${lbl}</span><span class="val">${val}</span></div>`;
    }
    return `<div class="sum-row"><span class="val" style="grid-column:1/-1">${inner}</span></div>`;
  });
  const rankSummaryLine=line=>{
    const txt=String(line).toLowerCase();
    if(txt.includes('about:'))return 1;
    if(txt.includes('latest:'))return 2;
    if(txt.includes('action needed:')||txt.includes('action:'))return 3;
    return 4;
  };
  return `<div class="sum-rows">${rows.map((html,i)=>({html,i,rank:rankSummaryLine(lines[i])})).sort((a,b)=>a.rank-b.rank||a.i-b.i).map(x=>x.html).join('')}</div>`;
}
function renderSumBox(t){
  const badge=t.replyCount>1?`<span class="sum-thread-badge">${t.replyCount} messages</span>`:'';
  const aiBadge=t.aiGenerated?`<span class="sum-ai-badge">✦ AI</span>`:'';
  const hdr=`<div class="sum-hdr">${emailSubject(t)}${badge}${aiBadge}</div>`;
  const body=t.summary?formatSummaryHTML(t.summary):'<span class="sum-empty">No summary yet — click Refresh Summary to generate one.</span>';
  const lnk=threadLinkButton(t);
  const linkRow=lnk?`<div class="sum-link-row">${lnk}</div>`:'';
  return hdr+body+linkRow;
}
function threadLinkButton(t,label="Open Email Thread"){
  return (t.conversationId||taskEmailId(t))?`<button class="email-link-btn" onclick="event.stopPropagation();openTaskThread(${t.id})">${label}</button>`:"";
}
function stripEmailHtml(html){
  const div=document.createElement("div");
  div.innerHTML=html||"";
  // Remove quoted/forwarded blocks — Outlook, Gmail, Apple Mail
  div.querySelectorAll([
    "style","script","blockquote","hr",
    "[id='divRplyFwdMsg']","[id='divTaggedMessage']",
    "[class*='gmail_quote']","[class*='gmail_extra']",
    "[class*='x_gmail_quote']","[class*='OutlookMessageHeader']",
    "[class*='WordSection']","[class*='MsoNormal']"
  ].join(",")).forEach(x=>x.remove());
  // Insert newlines before block/cell elements so adjacent cells don't merge
  div.querySelectorAll("p,div,br,tr,td,li,h1,h2,h3,h4,h5,h6").forEach(x=>{
    x.insertAdjacentText("afterbegin","\n");
  });
  return div.textContent||div.innerText||"";
}
function plainizeEmailBody(text){
  return String(text||"")
    .replace(/\r/g,"")
    .replace(/[ \t]+\n/g,"\n")
    .replace(/\n{3,}/g,"\n\n")
    .trim();
}
function cleanEmailText(text){
  return String(text||"")
    .replace(/\r/g,"\n")
    .replace(/https?:\/\/\S+/gi," ")
    .replace(/\b[\w.-]+(\/[\w.-]+){2,}\b/g," ")
    .replace(/\S+@\S+\.\S+/g," ")
    .split("\n")
    .map(x=>x.trim())
    .filter(x=>x.length>3)
    .filter(x=>!/^>/.test(x))
    .filter(x=>!(/^(from|sent|to|cc|bcc|subject|date)\s*:/i.test(x)))
    .filter(x=>!(/^(thanks|thank you|regards|best|sincerely|cheers|yours|sent from my|confidentiality|this message|external email|unsubscribe|disclaimer)/i.test(x)))
    .filter(x=>!(/^\+?1?[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/.test(x)))
    .filter(x=>!(/^\d+\s+[A-Z][a-z]/.test(x)))
    .filter(x=>!(/^(main|direct|cell|fax|phone|tel|mobile|office)\s*:/i.test(x)))
    .filter(x=>!(/\b[A-Z]{2}\s+\d{5}\b/.test(x)))
    .join(" ")
    .replace(/\s+/g," ")
    .trim();
}
function cleanSubject(subject){
  return String(subject||"(no subject)").replace(/^\s*(re|fw|fwd):\s*/ig,"").trim();
}
function cleanEmailBodyForAI(htmlContent){
  const div=document.createElement('div');
  div.innerHTML=htmlContent||'';
  // Remove reply/forward headers, signatures, quoted blocks
  div.querySelectorAll([
    'style','script','blockquote','hr',
    '[id="divRplyFwdMsg"]','[id="divTaggedMessage"]',
    '[class*="gmail_quote"]','[class*="gmail_extra"]',
    '[class*="x_gmail_quote"]','[class*="OutlookMessageHeader"]',
    '[class*="WordSection"]','[class*="MsoNormal"]'
  ].join(',')).forEach(x=>x.remove());
  div.querySelectorAll('p,div,br,tr,td,li,h1,h2,h3,h4,h5,h6').forEach(x=>{
    x.insertAdjacentText('afterbegin','\n');
  });
  const raw=div.textContent||div.innerText||'';
  const lines=raw.replace(/\r/g,'\n').split('\n').map(l=>l.trim()).filter(l=>l.length>2);
  const result=[];
  for(const line of lines){
    // Stop at common signature/forward separators
    if(/^_{5,}|-{5,}|={5,}|\*{5,}/.test(line))break;
    if(/^(from|sent|to|cc|bcc|subject|date)\s*:/i.test(line))break;
    if(/^(on .{5,} wrote:)/i.test(line))break;
    // Skip disclaimer/footer lines
    if(/^(thanks|thank you|regards|best|sincerely|cheers|yours truly|sent from my|confidential|this (message|email)|external email|unsubscribe|disclaimer|privileged)/i.test(line))continue;
    if(/^(\+?1?[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})/.test(line))continue;
    if(/^(main|direct|cell|fax|phone|tel|mobile|office)\s*:/i.test(line))continue;
    if(/https?:\/\/\S+/.test(line)&&line.length<80&&!/\w{4,}/.test(line.replace(/https?:\/\/\S+/g,'')))continue;
    result.push(line);
  }
  return result.join(' ').replace(/\s+/g,' ').trim();
}
function uniqueSentences(sentences){
  const seen=new Set();
  return sentences.filter(s=>{
    const key=s.toLowerCase().replace(/[^a-z0-9 ]/g,"").replace(/\s+/g," ").slice(0,120);
    if(!key||seen.has(key))return false;
    seen.add(key);
    return true;
  });
}
function shortenSentence(s,max=190){
  s=String(s||"").replace(/\s+/g," ").trim();
  return s.length>max?s.slice(0,max-3).replace(/\s+\S*$/,"")+"...":s;
}
function humanSubject(subject){
  const s=cleanSubject(subject).replace(/\([^)]*\d{4,}[^)]*\)/g,"").replace(/\s+-\s+/g,": ").replace(/\s+/g," ").trim();
  return s||"Email thread";
}
function dpegPeopleInThread(messages){
  const people=new Map();
  (messages||[]).forEach(m=>{
    const contacts=[m.from?.emailAddress,...(m.toRecipients||[]).map(r=>r.emailAddress),...(m.ccRecipients||[]).map(r=>r.emailAddress)].filter(Boolean);
    contacts.forEach(c=>{
      const email=normEmail(c.address);
      if(!email.includes("@dhananipeg.com"))return;
      const p=findPersonByEmail(email);
      const name=p?.name||c.name||email.split("@")[0];
      people.set(name.toLowerCase(),name);
    });
  });
  return [...people.values()].sort((a,b)=>a.localeCompare(b));
}
function localThreadSummary(messages){
  const ordered=[...(messages||[])].sort((a,b)=>new Date(a.receivedDateTime||a.sentDateTime||0)-new Date(b.receivedDateTime||b.sentDateTime||0));
  const subject=humanSubject(ordered.at(-1)?.subject||ordered[0]?.subject);
  const involved=dpegPeopleInThread(ordered);

  // Find the primary external sender name
  const sender=(()=>{
    for(const m of ordered){
      const addr=normEmail(m.from?.emailAddress?.address||'');
      if(addr&&!addr.includes('@dhananipeg.com'))
        return (m.from?.emailAddress?.name||addr.split('@')[0]).replace(/["'<>]/g,'').trim();
    }
    return null;
  })();

  function getClean(m){return cleanEmailText(m.body?.contentType==="html"?stripEmailHtml(m.body?.content):m.body?.content||m.bodyPreview||"");}
  const allText=ordered.map(getClean).join(" ");
  const latestText=getClean(ordered.at(-1));

  function pickSentences(text){
    return uniqueSentences(
      (text.match(/[^.!?]+[.!?]+|[^.!?]+$/g)||[])
        .map(s=>s.trim())
        .filter(s=>s.length>22&&s.length<260)
        .filter(s=>s.split(/\s+/).length>=5)
        .filter(s=>!/(?:^|\s)(?:com)?from\s*:/i.test(s))
        .filter(s=>!/\b\d{5}\b/.test(s))
        .filter(s=>!/(main|cell|fax|direct)\s*:\s*[\d.]/i.test(s))
    );
  }

  const allSentences=pickSentences(allText);
  const latestSentences=pickSentences(latestText);

  // Sentence classifiers — distinct purposes so bullets don't repeat each other
  const ACTION=/\bplease\b|\bneed(?:s)?\s+(you|us|dpeg|the\s+team)\b|\bcan\s+you\b|\bcould\s+you\b|\bwould\s+you\b|\brequest(?:ing)?\b|\brequire(?:s|d)?\b|\bconfirm\b|\bapprove\b|\bprovide\b|\bfollow[\s-]up\b|\bdeadline\b|\bdue\s+(by|on|date)\b|\burgent\b|\baction\s+required\b|\brespond\b|\bpending\s+your\b/i;
  const CONTEXT=/\battach(?:ed|ment)\b|\bproposal\b|\bagreement\b|\blease\b|\bcontract\b|\binvoice\b|\bquote\b|\bpresentation\b|\breport\b|\bschedule\b|\bas\s+discussed\b|\bper\s+our\b|\bfollowing\s+(our|up)\b|\bregarding\b|\bpursuant\b|\bplease\s+find\b|\b(am|is)\s+(writing|reaching|following)\b|\bwanted\s+to\b|\bwriting\s+to\b|\bthis\s+(email|message|note)\b/i;

  const actionSentence=allSentences.find(s=>ACTION.test(s));
  const contextSentence=allSentences.find(s=>CONTEXT.test(s)&&s!==actionSentence);

  // Entity extraction — concrete facts to surface
  const amounts=(allText.match(/\$[\d,]+(?:\.\d{2})?(?:\s*(?:million|M|billion|B|thousand|k))?\b/gi)||[]).slice(0,2);
  const deadline=(allText.match(/\b(?:by|before|deadline|due(?:\s+by)?)\s+[A-Z][a-z]+(?:\s+\d{1,2})?(?:,?\s*\d{4})?\b/i)||[])[0];
  const sqft=(allText.match(/[\d,]+\s*(?:sq\.?\s*ft\.?|square\s+feet|sqft|SF)\b/gi)||[])[0];
  const pct=(allText.match(/\b\d+(?:\.\d+)?%/g)||[])[0];

  // Latest meaningful sentence from the most recent email only
  const latestSentence=latestSentences.find(s=>s!==actionSentence&&s!==contextSentence&&s.split(/\s+/).length>=6);

  const bullets=[];
  bullets.push(`• Topic: ${subject}`);

  if(contextSentence){
    bullets.push(`• About: ${shortenSentence(contextSentence,175)}`);
  }else if(sender){
    const verb=ordered.length>1?'discussed':'sent an email about';
    bullets.push(`• About: ${sender} ${verb} ${subject.toLowerCase()}.`);
  }

  if(actionSentence){
    bullets.push(`• Action needed: ${shortenSentence(actionSentence,195)}`);
  }else{
    bullets.push(`• Action needed: Review and determine next steps.`);
  }

  const facts=[];
  if(deadline)facts.push(deadline);
  if(amounts.length)facts.push(amounts.join(', '));
  if(sqft)facts.push(sqft);
  if(pct&&!amounts.length)facts.push(pct);
  if(facts.length)bullets.push(`• Key info: ${facts.join(' · ')}`);
  else if(involved.length)bullets.push(`• Assigned to: ${involved.slice(0,2).join(", ")}`);

  return bullets.slice(0,4).join("\n");
}
