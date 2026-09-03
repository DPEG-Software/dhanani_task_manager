// ============================================================
// AI SETTINGS
// ============================================================
const WORKER_URL=window.DPEG_STAGING_MODE
  ?`${window.DPEG_STAGING_WORKER}/`
  :'https://dpeg-ai-summarize.systemmanager1.workers.dev/';
function loadAIConfig(){
  const url=localStorage.getItem('dpeg_ai_fn_url')||'';
  const el=document.getElementById('ai-fn-url');
  if(el)el.value=url||WORKER_URL;
  const enabled=localStorage.getItem('dpeg_ai_enabled')!=='false';
  const enabledEl=document.getElementById('ai-enabled');if(enabledEl)enabledEl.checked=enabled;
  const attachmentsEl=document.getElementById('ai-attachments-enabled');if(attachmentsEl)attachmentsEl.checked=localStorage.getItem('dpeg_ai_attachments_enabled')==='true';
  const previewEl=document.getElementById('ai-preview-enabled');if(previewEl)previewEl.checked=localStorage.getItem('dpeg_ai_preview_enabled')!=='false';
  updateAIStatusDot(enabled);
  return url;
}
function saveAISettings(){
  const url=(document.getElementById('ai-fn-url')?.value||'').trim();
  const enabled=!!document.getElementById('ai-enabled')?.checked;
  const attachmentsEnabled=!!document.getElementById('ai-attachments-enabled')?.checked;
  const previewEnabled=!!document.getElementById('ai-preview-enabled')?.checked;
  localStorage.setItem('dpeg_ai_fn_url',url);
  localStorage.setItem('dpeg_ai_enabled',String(enabled));
  localStorage.setItem('dpeg_ai_attachments_enabled',String(attachmentsEnabled));
  localStorage.setItem('dpeg_ai_preview_enabled',String(previewEnabled));
  updateAIStatusDot(enabled);
  closeMo('mo-ai-settings');
  toast(enabled?'AI summaries enabled with server-side redaction':'AI summarization disabled — using local processing');
}
function openAISettings(){
  loadAIConfig();
  document.getElementById('ai-settings-status').innerHTML='';
  document.getElementById('mo-ai-settings').classList.add('open');
}
function updateAIStatusDot(active){
  const dot=document.getElementById('ai-status-dot');
  if(dot)dot.style.background=active?'#22c55e':'#d1d5db';
}

// ============================================================
// EMAIL SIGNATURE
// ============================================================
function signatureStorageKey(){
  return `dpeg_email_signature_${normEmail(currentUser?.email||'default')}`;
}
function getSignatureHTML(){
  return compactSignatureHTML(localStorage.getItem(signatureStorageKey())||'');
}
function compactSignatureHTML(html){
  const wrap=document.createElement('div');
  wrap.innerHTML=html||'';
  wrap.querySelectorAll('p,div').forEach(el=>{
    if(!el.textContent.trim()&&!el.querySelector('img'))el.remove();
    else el.style.margin='0';
  });
  wrap.querySelectorAll('br+br').forEach(br=>br.remove());
  return wrap.innerHTML.trim();
}
function openSignatureSettings(){
  const editor=document.getElementById('signature-editor');
  if(editor)editor.innerHTML=getSignatureHTML();
  document.getElementById('mo-signature').classList.add('open');
}
function insertSignatureLogo(file){
  if(!file)return;
  if(!file.type?.startsWith('image/')){toast('Please choose an image file');return;}
  const reader=new FileReader();
  reader.onload=()=>{
    const editor=document.getElementById('signature-editor');
    if(!editor)return;
    editor.focus();
    const img=`<div><img src="${reader.result}" alt="DPEG logo" style="max-width:190px;height:auto;display:block;margin:4px 0 8px"></div>`;
    try{
      document.execCommand('insertHTML',false,img);
    }catch(e){
      editor.insertAdjacentHTML('beforeend',img);
    }
    const input=document.getElementById('signature-logo-input');
    if(input)input.value='';
  };
  reader.readAsDataURL(file);
}
function saveSignatureSettings(){
  const html=compactSignatureHTML(document.getElementById('signature-editor')?.innerHTML||'');
  localStorage.setItem(signatureStorageKey(),html);
  closeMo('mo-signature');
  applyComposeSignature();
  toast('Signature saved');
}
function clearSignatureSettings(){
  localStorage.removeItem(signatureStorageKey());
  const editor=document.getElementById('signature-editor');
  if(editor)editor.innerHTML='';
  applyComposeSignature();
  toast('Signature cleared');
}
function applyComposeSignature(){
  const preview=document.getElementById('compose-signature-preview');
  if(!preview)return;
  const html=getSignatureHTML();
  preview.innerHTML=html;
  preview.style.display=html?'block':'none';
}
function escapeHtml(s){
  return String(s||'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
function composeBodyHTML(text){
  const typed=escapeHtml(text).replace(/\r?\n/g,'<br>');
  const sig=getSignatureHTML();
  return `${typed}${sig?`<br><br><div style="line-height:1.25">${sig}</div>`:''}`;
}
function prepareComposedEmail(text){
  let html=composeBodyHTML(text);
  const inlineAttachments=[];
  let idx=0;
  html=html.replace(/<img\b([^>]*?)\bsrc=(["'])data:(image\/[^;,]+);base64,([^"']+)\2([^>]*)>/gi,(match,before,q,mime,b64,after)=>{
    const ext=(mime.split('/')[1]||'png').replace(/[^a-z0-9]/gi,'')||'png';
    const cid=`dpeg-signature-logo-${Date.now()}-${idx}@dpeg`;
    inlineAttachments.push({
      '@odata.type':'#microsoft.graph.fileAttachment',
      name:`signature-logo-${++idx}.${ext}`,
      contentType:mime,
      contentBytes:b64,
      isInline:true,
      contentId:cid
    });
    return `<img${before}src="cid:${cid}"${after}>`;
  });
  return {html,inlineAttachments};
}
function outlookFlagStatus(email){
  return String(email?.flag?.flagStatus||'notFlagged');
}
function flagBadge(email,interactive=false){
  const status=outlookFlagStatus(email);
  const color=status==='complete'?'#15803d':status==='flagged'?'#dc2626':'#9ca3af';
  const glyph=status==='complete'?'✓':'⚑';
  if(!interactive&&status==='notFlagged')return '';
  const title=status==='flagged'?'Remove Outlook flag':status==='complete'?'Completed in Outlook':'Flag in Outlook';
  const attrs=interactive?`role="button" onclick="toggleOutlookFlag('${email?.id||''}',event)"`:'';
  return `<span ${attrs} title="${title}" style="color:${color};font-size:13px;font-weight:900;line-height:1;cursor:${interactive?'pointer':'default'}">${glyph}</span>`;
}

// ============================================================
// ADMIN SETTINGS
// ============================================================
function openAdminSettings(){
  if(!isAdmin()){toast('Admin access only');return;}
  renderAdminPeopleList();
  renderAdminDeptEditor();
  document.getElementById('mo-admin').classList.add('open');
}
function switchAdminTab(tab){
  document.querySelectorAll('.admin-tab').forEach(b=>b.classList.remove('active'));
  document.getElementById('atab-'+tab)?.classList.add('active');
  document.getElementById('admin-contacts-pane').style.display=tab==='contacts'?'block':'none';
  document.getElementById('admin-depts-pane').style.display=tab==='depts'?'block':'none';
}
function allKnownPeople(){
  const map={};
  Object.values(staffConfig).forEach(p=>{if(!p?.email)return;const k=normEmail(p.email);map[k]={...map[k],...p,src:'custom'};});
  return Object.values(map).sort((a,b)=>String(a.name).localeCompare(String(b.name)));
}
function adminSearchTerm(){
  return String(document.getElementById('admin-search')?.value||'').trim().toLowerCase();
}
function filterAdminPeople(people){
  const q=adminSearchTerm();
  if(!q)return people;
  return people.filter(p=>[p.name,p.email,p.dept,p.role].some(v=>String(v||'').toLowerCase().includes(q)));
}
function renderAdminPeopleList(){
  const people=filterAdminPeople(allKnownPeople());
  const el=document.getElementById('admin-people-list');
  if(!el)return;
  if(!people.length){el.innerHTML=`<div style="color:var(--muted);font-size:12px;padding:12px 0">${adminSearchTerm()?'No contacts match this search.':'No contacts yet. Click "Sync from Outlook" to import.'}</div>`;return;}
  el.innerHTML=`<div class="admin-person-row" style="font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);border-bottom:2px solid var(--border)"><span>Name / Email</span><span>Department</span><span>Role</span><span></span></div>`+
    people.map((p,i)=>{const currentDept=hasAssignedDepartment(p.dept)?p.dept:'Needs Department';return `
    <div class="admin-person-row" id="apr-${i}">
      <div>${av(p.name||'?',26)} <div style="display:inline-block;vertical-align:middle;margin-left:6px"><div style="display:flex;align-items:center;gap:4px"><span style="font-size:12.5px;font-weight:600;color:var(--body)">${p.name||'Unknown'}</span><button title="Edit name" onclick="renameAdminPerson('${(p.email||'').replace(/'/g,"\\'")}','${(p.name||'').replace(/'/g,"\\'")}')" style="background:none;border:none;cursor:pointer;padding:1px 3px;color:var(--muted);font-size:11px;line-height:1" onmouseover="this.style.color='var(--body)'" onmouseout="this.style.color='var(--muted)'">&#9998;</button></div><div style="font-size:11px;color:var(--muted)">${p.email||''}</div></div></div>
      <select class="form-sel" style="font-size:11.5px;padding:3px 6px" onchange="updateAdminPersonDept('${(p.email||'').replace(/'/g,"\\'")}','${(p.name||'').replace(/'/g,"\\'")}',this.value)">
        <option${currentDept==='Needs Department'?' selected':''}>Needs Department</option>
        ${allDepartments().map(d=>`<option${d===currentDept?' selected':''}>${d}</option>`).join('')}
      </select>
      <div style="font-size:11.5px;color:var(--muted)">${p.role||''}</div>
      <button class="btn btn-ghost btn-sm" style="padding:2px 6px;font-size:11px;color:#dc2626;border-color:#fca5a5" onclick="removeAdminPerson('${(p.email||'').replace(/'/g,"\\'")}')">&#10005;</button>
    </div>`}).join('');
}
function renderAdminDeptEditor(){
  const people=filterAdminPeople(allKnownPeople());
  const el=document.getElementById('admin-dept-editor');if(!el)return;
  const byDept={};
  allDepartments().forEach(d=>{if(!byDept[d])byDept[d]=[];});
  people.forEach(p=>{const d=p.dept||'Needs Department';if(!byDept[d])byDept[d]=[];byDept[d].push(p);});
  el.innerHTML=Object.entries(byDept).sort(([a],[b])=>a.localeCompare(b)).map(([dept,pp])=>`
    <div style="margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
        <span style="width:10px;height:10px;border-radius:50%;background:${dcolor(dept)};display:inline-block"></span>
        <span style="font-size:12px;font-weight:700;color:var(--body)">${dept}</span>
        <span style="font-size:11px;color:var(--muted)">${pp.length} ${pp.length===1?'person':'people'}</span>
      </div>
      <div style="padding-left:18px">${pp.map(p=>`<div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f3f4f6;font-size:12.5px">${av(p.name||'?',22)}<span>${p.name}</span><span style="color:var(--muted);font-size:11px">${p.email||''}</span></div>`).join('')}</div>
    </div>`).join('');
}
async function updateAdminPersonDept(email,name,dept){
  const selected=await ensureDepartmentForPerson({email,name,dept:configuredDept(email,name)},dept);
  if(selected===null){renderAdminPeopleList();return;}
  applyDepartmentAssignment(email,name,selected);
  renderAdminPeopleList();
  renderAdminDeptEditor();
}
function createAdminDepartment(){
  const input=document.getElementById('admin-new-dept');
  const name=String(input?.value||'').trim().replace(/\s+/g,' ');
  if(!name){toast('Enter a department name');return;}
  if(allDepartments().some(d=>d.toLowerCase()===name.toLowerCase())){toast('Department already exists');if(input)input.value='';return;}
  customDepartments=[...(customDepartments||[]),name].sort((a,b)=>a.localeCompare(b));
  if(input)input.value='';
  initSelects();
  renderAdminPeopleList();
  renderAdminDeptEditor();
  toast(`${name} department added`);
}
function removeAdminPerson(email){
  const key=Object.keys(staffConfig).find(k=>normEmail(staffConfig[k]?.email||'')=== normEmail(email));
  if(key)delete staffConfig[key];
  renderAdminPeopleList();renderAdminDeptEditor();
}
function renameAdminPerson(email,currentName){
  const newName=(prompt('Display name:',currentName)||'').trim();
  if(!newName||newName===currentName)return;
  const key=Object.keys(staffConfig).find(k=>normEmail(staffConfig[k]?.email||'')===normEmail(email));
  if(key){staffConfig[key].name=newName;renderAdminPeopleList();renderAdminDeptEditor();toast(`Renamed to ${newName}`);}
}
function addPersonManually(){
  const name=prompt('Full name:');if(!name)return;
  const email=prompt('Email address:');if(!email)return;
  const dept=prompt('Department (or leave blank):')||'Needs Department';
  const key=staffKey(email,name);
  staffConfig[key]={name,email,dept,role:''};
  renderAdminPeopleList();renderAdminDeptEditor();
  toast(`${name} added`);
}
async function saveAdminSettings(){
  await saveTasksToOneDrive();
  closeMo('mo-admin');
  initSelects();
  toast('Admin settings saved');
}
function graphEmailValue(item){
  if(!item)return "";
  if(typeof item==="string")return item;
  return item.address||item.emailAddress?.address||item.email||item.mail||item.userPrincipalName||"";
}
function graphNameValue(item,email){
  if(item?.displayName)return item.displayName;
  if(item?.name)return item.name;
  if(item?.emailAddress?.name)return item.emailAddress.name;
  if(item?.givenName&&item?.surname)return `${item.givenName} ${item.surname}`;
  return email.split("@")[0];
}
function upsertSyncedContact(raw,defaultDept="Needs Department"){
  const candidates=[
    ...(raw.emailAddresses||[]),
    ...(raw.scoredEmailAddresses||[]),
    raw.mail,
    raw.userPrincipalName,
    raw.email,
    raw.emailAddress
  ];
  let email="";
  for(const c of candidates){
    email=normEmail(graphEmailValue(c));
    if(email&&email.includes("@"))break;
  }
  if(!email||!email.includes("@"))return false;
  const name=(raw.displayName||raw.name||raw.emailAddress?.name||graphNameValue(raw,email)||email.split("@")[0]).trim();
  const dept=raw.department||defaultDept;
  const role=raw.jobTitle||raw.role||"";
  const key=staffKey(email,name);
  const before=JSON.stringify(staffConfig[key]||null);
  staffConfig[key]={...(staffConfig[key]||{}),name:staffConfig[key]?.name||name,email,dept:staffConfig[key]?.dept||dept,role:staffConfig[key]?.role||role};
  return JSON.stringify(staffConfig[key])!==before;
}
function contactsFromMessage(m){
  return [
    m.from?.emailAddress,
    m.sender?.emailAddress,
    ...(m.toRecipients||[]).map(r=>r.emailAddress),
    ...(m.ccRecipients||[]).map(r=>r.emailAddress),
    ...(m.bccRecipients||[]).map(r=>r.emailAddress),
  ].filter(Boolean);
}
async function syncContactsFromMailbox(){
  let added=0;
  try{
    const token=await getAccessToken();
    const urls=[
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100&$select=from,sender,toRecipients,ccRecipients,receivedDateTime&$orderby=receivedDateTime desc",
      "https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=100&$select=from,sender,toRecipients,ccRecipients,bccRecipients,sentDateTime&$orderby=sentDateTime desc"
    ];
    for(const url of urls){
      const res=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
      if(!res.ok)continue;
      const data=await res.json();
      (data.value||[]).forEach(m=>{
        contactsFromMessage(m).forEach(c=>{
          if(upsertSyncedContact({name:c.name,displayName:c.name,emailAddress:c},isInternalEmail(c.address)?"Needs Department":"Outside DPEG"))added++;
        });
      });
    }
  }catch(err){
    console.warn("Mailbox contact sync skipped:",err.message);
  }
  return added;
}
async function syncContacts(){
  if(!isAdmin()){toast('Admin access only');return;}
  const btn=document.getElementById('sync-contacts-btn');
  const status=document.getElementById('sync-status');
  btn.disabled=true;btn.textContent='Syncing...';
  status.textContent='';
  try{
    let token;
    try{
      const r=await msalInstance.acquireTokenSilent({scopes:SCOPES_CONTACTS,account:currentAccount});
      token=r.accessToken;
    }catch{
      const r=await msalInstance.acquireTokenPopup({scopes:SCOPES_CONTACTS});
      token=r.accessToken;
    }
    let synced=0;
    // Fetch Outlook contacts
    const cRes=await fetch('https://graph.microsoft.com/v1.0/me/contacts?$top=150&$select=displayName,emailAddresses,department,jobTitle',{headers:{Authorization:`Bearer ${token}`}});
    if(cRes.ok){
      const cd=await cRes.json();
      (cd.value||[]).forEach(c=>{
        if(upsertSyncedContact(c))synced++;
      });
    }
    // Fetch People (colleagues)
    const pRes=await fetch('https://graph.microsoft.com/v1.0/me/people?$top=100&$select=displayName,emailAddresses,scoredEmailAddresses,department,jobTitle',{headers:{Authorization:`Bearer ${token}`}});
    if(pRes.ok){
      const pd=await pRes.json();
      (pd.value||[]).forEach(p=>{
        if(upsertSyncedContact(p))synced++;
      });
    }
    // Fetch ALL org users from Azure AD directory (requires User.ReadBasic.All)
    let usersUrl='https://graph.microsoft.com/v1.0/users?$select=displayName,mail,jobTitle,department&$filter=accountEnabled eq true&$top=999';
    while(usersUrl){
      const uRes=await fetch(usersUrl,{headers:{Authorization:`Bearer ${token}`}});
      if(!uRes.ok)break;
      const ud=await uRes.json();
      (ud.value||[]).forEach(u=>{
        if(u.mail&&isInternalEmail(u.mail)){if(upsertSyncedContact(u))synced++;}
      });
      usersUrl=ud['@odata.nextLink']||null;
    }
    synced+=await syncContactsFromMailbox();
    await saveTasksToOneDrive();
    renderAdminPeopleList();renderAdminDeptEditor();
    initSelects();
    status.textContent=`✓ Synced ${synced} contacts`;
    const available=allKnownPeople().length;
    status.textContent=`Synced ${synced} new/updated contact${synced!==1?'s':''}. ${available} available in autocomplete.`;
    status.style.color='var(--forest)';
    toast(`${available} contacts available for autocomplete`);
  }catch(err){
    status.textContent=`✗ Sync failed: ${err.message}`;
    status.style.color='#dc2626';
  }finally{
    btn.disabled=false;btn.textContent='Sync from Outlook';
  }
}

async function autoSyncContacts(){
  try{
    let token;
    try{
      const r=await msalInstance.acquireTokenSilent({scopes:SCOPES_CONTACTS,account:currentAccount});
      token=r.accessToken;
    }catch{return;}
    let changed=false;
    const pRes=await fetch('https://graph.microsoft.com/v1.0/me/people?$top=150&$select=displayName,emailAddresses,scoredEmailAddresses,department,jobTitle',{headers:{Authorization:`Bearer ${token}`}});
    if(pRes.ok){
      const pd=await pRes.json();
      (pd.value||[]).forEach(p=>{if(upsertSyncedContact(p))changed=true;});
    }
    const cRes=await fetch('https://graph.microsoft.com/v1.0/me/contacts?$top=150&$select=displayName,emailAddresses,department,jobTitle',{headers:{Authorization:`Bearer ${token}`}});
    if(cRes.ok){
      const cd=await cRes.json();
      (cd.value||[]).forEach(c=>{if(upsertSyncedContact(c))changed=true;});
    }
    // Fetch ALL org users from Azure AD directory (requires User.ReadBasic.All)
    let usersUrl='https://graph.microsoft.com/v1.0/users?$select=displayName,mail,jobTitle,department&$filter=accountEnabled eq true&$top=999';
    while(usersUrl){
      const uRes=await fetch(usersUrl,{headers:{Authorization:`Bearer ${token}`}});
      if(!uRes.ok)break;
      const ud=await uRes.json();
      (ud.value||[]).forEach(u=>{if(u.mail&&isInternalEmail(u.mail)){if(upsertSyncedContact(u))changed=true;}});
      usersUrl=ud['@odata.nextLink']||null;
    }
    if(changed)await saveTasksToOneDrive();
  }catch{}
}
async function testAIConnection(){
  const url=(document.getElementById('ai-fn-url')?.value||'').trim();
  const statusEl=document.getElementById('ai-settings-status');
  const btn=document.getElementById('ai-test-btn');
  if(!url){statusEl.innerHTML='<div class="ai-status-bar ai-err">Please enter a URL first.</div>';return;}
  btn.disabled=true;btn.textContent='Testing...';
  statusEl.innerHTML='';
  try{
    const token=await getAccessToken();
    const res=await fetch(url,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':`Bearer ${token}`},
      body:JSON.stringify({subject:'Test connection',emailText:'This is a test message to verify the connection is working.',senderName:'Test',messageCount:1})
    });
    if(res.ok){
      const d=await res.json();
      statusEl.innerHTML=`<div class="ai-status-bar ai-ok">✓ Connected — AI returned a summary successfully.</div>`;
    }else{
      const errBody=await res.json().catch(()=>({}));
      const detail=String(errBody.detail||errBody.error||'').trim();
      statusEl.innerHTML=`<div class="ai-status-bar ai-err">✗ Function returned ${res.status}. Check deployment and env vars.${detail?`<div style="margin-top:6px;font-size:11px;font-family:monospace;white-space:pre-wrap;word-break:break-word">${escapeHtml(detail).slice(0,500)}</div>`:''}</div>`;
    }
  }catch(err){
    statusEl.innerHTML=`<div class="ai-status-bar ai-err">✗ Could not reach function: ${err.message}</div>`;
  }finally{
    btn.disabled=false;btn.textContent='Test Connection';
  }
}
