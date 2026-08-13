// ============================================================
// ONEDRIVE DATA SYNC
// ============================================================
function setSyncStatus(state, label) {
  const dot = document.getElementById("sync-dot");
  const lbl = document.getElementById("sync-label");
  if (dot) { dot.className = "sync-dot " + state; }
  if (lbl) { lbl.textContent = label; }
}

async function getAccessToken() {
  try {
    const silentPromise = msalInstance.acquireTokenSilent({ scopes: SCOPES_GRAPH, account: currentAccount });
    const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error('silent_timeout')), 10000));
    const result = await Promise.race([silentPromise, timeoutPromise]);
    return result.accessToken;
  } catch (err) {
    try {
      const result = await msalInstance.acquireTokenPopup({ scopes: SCOPES_GRAPH });
      return result.accessToken;
    } catch (popupErr) {
      console.error("Token error:", popupErr);
      throw popupErr;
    }
  }
}

function isProofUploadRoute(){
  return new URLSearchParams(location.search).get('proof')==='1';
}

let proofParamsOverride=null;

function proofRouteParams(){
  if(proofParamsOverride)return proofParamsOverride;
  const p=new URLSearchParams(location.search);
  return {
    appTaskId:p.get('taskId')||'',
    recipientEmail:p.get('recipientEmail')||'',
    assignedByName:p.get('assignedByName')||'',
    assignedByEmail:p.get('assignedByEmail')||'',
    title:p.get('title')||'Assigned task',
    proofShareUrl:p.get('proofShareUrl')||'',
    todoListId:p.get('todoListId')||'',
    todoTaskId:p.get('todoTaskId')||'',
    proofInstructions:p.get('proofInstructions')||'',
  };
}

function workerBaseUrl(){
  return (localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL||'').replace(/\/?$/,'');
}

function buildProofSubmitLinkForTask(task,n={}){
  const url=new URL(location.origin+location.pathname);
  url.searchParams.set('proof','1');
  url.searchParams.set('taskId',String(task?.id||n.taskId||''));
  url.searchParams.set('recipientEmail',String(n.recipientEmail||task?.email||''));
  url.searchParams.set('assignedByName',String(task?.assignedByName||currentUser?.name||''));
  url.searchParams.set('assignedByEmail',String(task?.assignedByEmail||currentUser?.email||''));
  url.searchParams.set('title',String(task?.title||n.taskTitle||'Assigned task'));
  url.searchParams.set('proofShareUrl',String(task?.proofShareUrl||''));
  url.searchParams.set('proofInstructions',String(task?.proofInstructions||''));
  url.searchParams.set('todoListId',String(task?.recipientTodoListId||''));
  url.searchParams.set('todoTaskId',String(task?.recipientTodoTaskId||''));
  return url.toString();
}

function safeProofFileName(name){
  const cleaned=String(name||'proof-file').replace(/[\\/:*?"<>|#%{}~&]/g,'_').replace(/\s+/g,' ').trim();
  return cleaned.slice(0,120)||'proof-file';
}

async function ensureDriveFolderPath(token,segments){
  let path='';
  for(const name of segments){
    const parentPath=path;
    path+=`/${name}`;
    const check=await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(path)}`,{headers:{Authorization:`Bearer ${token}`}});
    if(check.ok)continue;
    const endpoint=parentPath
      ?`https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(parentPath)}:/children`
      :'https://graph.microsoft.com/v1.0/me/drive/root/children';
    const create=await fetch(endpoint,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({name,folder:{},'@microsoft.graph.conflictBehavior':'rename'})
    });
    if(!create.ok){
      const err=await create.text().catch(()=>'');
      throw new Error(`Could not create OneDrive folder ${name}: ${err}`);
    }
  }
}

async function ensureDriveFolderPathItem(token,segments){
  let path='';
  let last=null;
  for(const name of segments){
    const parentPath=path;
    path+=`/${name}`;
    const check=await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(path)}`,{headers:{Authorization:`Bearer ${token}`}});
    if(check.ok){last=await check.json();continue;}
    const endpoint=parentPath
      ?`https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(parentPath)}:/children`
      :'https://graph.microsoft.com/v1.0/me/drive/root/children';
    const create=await fetch(endpoint,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({name,folder:{},'@microsoft.graph.conflictBehavior':'rename'})
    });
    if(!create.ok){
      const err=await create.text().catch(()=>'');
      throw new Error(`Could not create OneDrive folder ${name}: ${err}`);
    }
    last=await create.json();
  }
  return last;
}

async function createOrgEditLinkForDriveItem(token,itemId){
  const res=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/createLink`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({type:'edit',scope:'organization'})
  });
  if(!res.ok){
    const detail=await res.text().catch(()=>'');
    throw new Error(`Could not create proof upload link: ${detail}`);
  }
  const data=await res.json();
  return data.link?.webUrl||'';
}

function proofFolderName(task){
  const id=String(task?.id||Date.now()).replace(/[^a-z0-9_.-]/gi,'_').slice(0,40);
  const title=safeProofFileName(task?.title||'Task').slice(0,60);
  return `${id}-${title}`;
}

async function ensureTaskProofFolder(task){
  if(task.proofShareUrl&&task.proofFolderItemId)return task.proofShareUrl;
  const token=await getAccessToken();
  const folder=await ensureDriveFolderPathItem(token,[ONEDRIVE_FOLDER,'Proofs',proofFolderName(task)]);
  const link=await createOrgEditLinkForDriveItem(token,folder.id);
  task.proofFolderItemId=folder.id||'';
  task.proofFolderWebUrl=folder.webUrl||'';
  task.proofShareUrl=link;
  return link;
}

function parseProofBlock(text){
  const raw=String(text||'');
  const start=raw.indexOf('DPEG_PROOF_START');
  const end=raw.indexOf('DPEG_PROOF_END');
  if(start<0||end<0||end<=start)return {proofs:[],base:raw.trim()};
  const before=raw.slice(0,start).trim();
  const after=raw.slice(end+'DPEG_PROOF_END'.length).trim();
  try{
    const parsed=JSON.parse(raw.slice(start+'DPEG_PROOF_START'.length,end).trim());
    return {proofs:Array.isArray(parsed.proofs)?parsed.proofs:[],base:[before,after].filter(Boolean).join('\n\n')};
  }catch{
    return {proofs:[],base:[before,after].filter(Boolean).join('\n\n')};
  }
}

function buildProofBlock(base,proofs){
  return `${String(base||'').trim()}\n\nDPEG_PROOF_START\n${JSON.stringify({proofs},null,2)}\nDPEG_PROOF_END`.trim();
}

async function createProofViewLink(token,itemId){
  // Try org-scoped first, fall back to anonymous view link
  for(const scope of['organization','anonymous']){
    try{
      const res=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/createLink`,{
        method:'POST',
        headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({type:'view',scope})
      });
      if(res.ok){
        const data=await res.json();
        const url=data.link?.webUrl||'';
        if(url)return url;
      }
    }catch{}
  }
  // Last resort: return the direct drive item webUrl
  try{
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}?$select=id,webUrl`,{headers:{Authorization:`Bearer ${token}`}});
    if(res.ok){const d=await res.json();return d.webUrl||'';}
  }catch{}
  return '';
}

function graphShareIdFromUrl(url){
  const b64=btoa(unescape(encodeURIComponent(String(url||''))))
    .replace(/=+$/,'')
    .replace(/\//g,'_')
    .replace(/\+/g,'-');
  return 'u!'+b64;
}

// Upload a single file to the RECIPIENT's own OneDrive under /DPEG Task Proofs/{taskId}/
// Handles large files (>4MB) via upload session automatically
async function uploadFileToOwnOneDrive(token,taskId,file,onProgress){
  const folderSegments=['DPEG Task Proofs',String(taskId).replace(/[^a-z0-9_-]/gi,'_').slice(0,40)];
  await ensureDriveFolderPath(token,folderSegments);
  const timestamp=new Date().toISOString().replace(/[:.]/g,'-');
  const filename=`${timestamp}-${safeProofFileName(file.name)}`;
  const encodedUploadPath=[...folderSegments,filename].map(s=>encodeURIComponent(s)).join('/');
  const THRESHOLD=4*1024*1024;
  let item;
  if(file.size<=THRESHOLD){
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodedUploadPath}:/content`,{
      method:'PUT',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':file.type||'application/octet-stream'},
      body:file,
    });
    if(!res.ok){const d=await res.text().catch(()=>'');throw new Error(`Upload failed: ${d}`.slice(0,200));}
    item=await res.json();
    if(onProgress)onProgress(file.size,file.size);
  }else{
    // Large file: upload session in 3.125MB aligned chunks
    const sessionRes=await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodedUploadPath}:/createUploadSession`,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({item:{'@microsoft.graph.conflictBehavior':'rename'}}),
    });
    if(!sessionRes.ok)throw new Error('Could not start upload session for large file');
    const {uploadUrl}=await sessionRes.json();
    const chunkSize=327680*10; // 3.125MB — must be multiple of 320KB
    let start=0,last=null;
    while(start<file.size){
      const end=Math.min(start+chunkSize,file.size);
      const chunkRes=await fetch(uploadUrl,{
        method:'PUT',
        headers:{'Content-Range':`bytes ${start}-${end-1}/${file.size}`,'Content-Length':String(end-start)},
        body:file.slice(start,end),
      });
      if(!chunkRes.ok&&chunkRes.status!==202){const d=await chunkRes.text().catch(()=>'');throw new Error(`Chunk upload failed: ${d}`.slice(0,200));}
      if(chunkRes.status===200||chunkRes.status===201)last=await chunkRes.json();
      if(onProgress)onProgress(end,file.size);
      start=end;
    }
    item=last;
  }
  if(!item?.id)throw new Error('Upload completed but no file ID returned');
  const viewUrl=await createProofViewLink(token,item.id);
  return{name:file.name,size:file.size,type:file.type||'application/octet-stream',shareId:graphShareIdFromUrl(viewUrl),webUrl:viewUrl,driveItemId:item.id||''};
}

async function saveProofNotificationToKV(params,proofs,note){
  const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  const userToken=await getAccessToken();
  const res=await fetch(`${fnUrl}/notify`,{
    method:'POST',
    headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
    body:JSON.stringify({type:'proof_submitted',appTaskId:params.appTaskId||'',taskTitle:params.title||'',senderEmail:params.assignedByEmail||'',recipientEmail:currentUser?.email||'',recipientName:currentUser?.name||'',proofs,note}),
  });
  if(!res.ok){const d=await res.text().catch(()=>'');throw new Error(`Could not save notification: ${d}`.slice(0,180));}
}

async function checkProofResult(appTaskId,recipientEmail){
  try{
    const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
    const userToken=await getAccessToken();
    const res=await fetch(`${fnUrl}/notify`,{headers:{Authorization:`Bearer ${userToken}`}});
    if(!res.ok)return null;
    const data=await res.json();
    const notifs=Array.isArray(data.notifications)?data.notifications:[];
    return notifs.filter(n=>n.type==='proof_result'&&String(n.appTaskId)===String(appTaskId)&&n.recipientEmail===recipientEmail).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt))[0]||null;
  }catch{return null;}
}

async function checkProofFollowup(params){
  try{
    const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
    const userToken=await getAccessToken();
    const res=await fetch(`${fnUrl}/notify`,{headers:{Authorization:`Bearer ${userToken}`}});
    if(!res.ok)return null;
    const data=await res.json();
    const notifs=Array.isArray(data.notifications)?data.notifications:[];
    const myEmail=(currentUser?.email||params.recipientEmail||'').toLowerCase();
    return notifs
      .filter(n=>n.type==='proof_submitted'&&n.status==='pending'&&String(n.appTaskId)===String(params.appTaskId||'')&&String(n.recipientEmail||'').toLowerCase()===myEmail)
      .sort((a,b)=>new Date(b.updatedAt||b.submittedAt||0)-new Date(a.updatedAt||a.submittedAt||0))[0]||null;
  }catch{return null;}
}

async function submitProofFollowupAnswer(notifId){
  const params=proofRouteParams();
  const answerEl=document.getElementById('proof-followup-answer');
  const statusEl=document.getElementById('proof-followup-status');
  const btn=document.getElementById('proof-followup-btn');
  const message=String(answerEl?.value||'').trim();
  const files=[...pendingFollowupFiles];
  if(!message&&!files.length){if(statusEl)statusEl.textContent='Please type a reply or attach a file.';return;}
  if(btn)btn.disabled=true;
  if(statusEl)statusEl.textContent='Sending...';
  try{
    const fnUrl=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
    const userToken=await getAccessToken();
    const attachments=[];
    if(files.length){
      const token=await getAccessToken();
      for(let i=0;i<files.length;i++){
        if(statusEl)statusEl.textContent=`Uploading ${i+1} of ${files.length}: ${files[i].name}`;
        const proofItem=await uploadFileToOwnOneDrive(token,params.appTaskId||`ext-${Date.now()}`,files[i]);
        proofItem.uploadedBy=currentUser?.email||'';
        proofItem.uploadedByName=currentUser?.name||'';
        proofItem.uploadedAt=new Date().toISOString();
        attachments.push(proofItem);
      }
      if(statusEl)statusEl.textContent='Sending reply...';
    }
    const res=await fetch(`${fnUrl}/notify`,{
      method:'POST',
      headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
      body:JSON.stringify({type:'proof_followup_answer',notifId,appTaskId:params.appTaskId||'',recipientEmail:currentUser?.email||params.recipientEmail||'',recipientName:currentUser?.name||'',message,attachments}),
    });
    if(!res.ok){const d=await res.text().catch(()=>'');throw new Error(d||'Could not send answer');}
    if(answerEl)answerEl.value='';
    pendingFollowupFiles=[];
    renderProofFollowupFileList();
    if(statusEl)statusEl.innerHTML='<span style="color:#166534;font-weight:600">✓ Reply sent. The assignor can now review it.</span>';
    await renderProofFollowupBox(params);
  }catch(err){
    if(statusEl)statusEl.innerHTML=`<span style="color:#b91c1c">${escapeHtml(err.message||'Could not send answer')}</span>`;
  }finally{if(btn)btn.disabled=false;}
}

async function renderProofFollowupBox(params){
  const host=document.getElementById('proof-followup-area');
  if(!host||!params.appTaskId)return;
  const n=await checkProofFollowup(params);
  const formArea=document.getElementById('proof-form-area');

  if(!n){
    // No active submission: initial submission or a declined proof that can
    // now be resubmitted.
    host.innerHTML='';
    if(formArea)formArea.style.display='';
    return;
  }
  showProofSentState();
}

function showProofSentState(){
  const host=document.getElementById('proof-followup-area');
  const formArea=document.getElementById('proof-form-area');
  if(host)host.innerHTML=`<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:7px;padding:16px;margin-bottom:4px">
    <div style="font-size:14px;font-weight:800;color:#166534;margin-bottom:5px">✓ Proof sent</div>
    <div style="font-size:12px;color:#166534;line-height:1.5">Your proof is waiting for review. Use <strong>Messages</strong> on the task for any additional information.</div>
  </div>`;
  if(formArea)formArea.style.display='none';
}

function renderProofFollowupFileList(){
  const list=document.getElementById('proof-followup-file-list');
  if(!list)return;
  if(!pendingFollowupFiles.length){list.innerHTML='';return;}
  list.innerHTML=pendingFollowupFiles.map((f,i)=>`<span style="display:inline-flex;align-items:center;gap:4px;background:#fef9c3;border:1px solid #fcd34d;border-radius:4px;padding:2px 6px;font-size:10.5px;color:#92400e;white-space:nowrap;max-width:160px">
    <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(f.name)}</span>
    <button type="button" onclick="pendingFollowupFiles.splice(${i},1);renderProofFollowupFileList()" style="border:none;background:none;color:#92400e;font-size:12px;font-weight:700;cursor:pointer;padding:0;line-height:1;flex-shrink:0" title="Remove">&#10005;</button>
  </span>`).join(' ');
}

async function uploadProofFiles(){
  const params=proofRouteParams();
  const statusEl=document.getElementById('proof-status');
  const note=document.getElementById('proof-note')?.value.trim()||'';
  const files=[...pendingProofFiles];
  const btn=document.getElementById('proof-upload-btn');
  if(!files.length&&!note){statusEl.textContent='Please add a file or write a note.';return;}
  btn.disabled=true;
  const progressWrap=document.getElementById('proof-progress');
  const progressBar=document.getElementById('proof-progress-bar');
  const progressLabel=document.getElementById('proof-progress-label');
  if(progressWrap&&files.length)progressWrap.style.display='block';
  statusEl.textContent='';
  try{
    const token=await getAccessToken();
    const taskId=params.appTaskId||`ext-${Date.now()}`;
    const uploaded=[];
    if(files.length){
      for(let i=0;i<files.length;i++){
        const file=files[i];
        if(progressLabel)progressLabel.textContent=`Uploading ${i+1} of ${files.length}: ${file.name}`;
        const proofItem=await uploadFileToOwnOneDrive(token,taskId,file,(done,total)=>{
          if(progressBar)progressBar.style.width=`${Math.round((done/total)*100)}%`;
        });
        proofItem.uploadedBy=currentUser?.email||'';
        proofItem.uploadedByName=currentUser?.name||'';
        proofItem.uploadedAt=new Date().toISOString();
        proofItem.note=note;
        uploaded.push(proofItem);
        if(progressBar)progressBar.style.width=`${Math.round(((i+1)/files.length)*100)}%`;
      }
    }
    if(progressLabel)progressLabel.textContent='Saving...';
    await saveProofNotificationToKV(params,uploaded,note);
    window.updateTasksTabProofState?.(params.appTaskId||'', 'submitted');
    if(progressWrap)progressWrap.style.display='none';
    pendingProofFiles=[];
    renderPendingProofFileList();
    const noteEl=document.getElementById('proof-note');
    if(noteEl)noteEl.value='';
    showProofSentState();
  }catch(err){
    if(progressWrap)progressWrap.style.display='none';
    statusEl.innerHTML=`<span style="color:#b91c1c">${escapeHtml(err.message||'Upload failed. Please try again.')}</span>`;
  }finally{btn.disabled=false;}
}

async function showProofUploadMode(paramsOverride){
  proofParamsOverride=paramsOverride||null;
  const fromTasksTab=!!paramsOverride;
  if(!fromTasksTab){
    document.querySelector('aside')?.remove();
    const main=document.querySelector('main');
    if(main)main.style.display='none';
  }
  const params=proofRouteParams();
  const screen=document.createElement('div');
  screen.id='proof-upload-screen';
  screen.style.cssText='position:fixed;inset:0;z-index:99999;min-height:100vh;background:#f8fafc;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Inter,sans-serif;box-sizing:border-box;overflow:auto';
  const instrHtml=params.proofInstructions?`<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;padding:9px 12px;margin-bottom:14px;font-size:12px;color:#166534;line-height:1.5"><strong>Instructions:</strong> ${escapeHtml(params.proofInstructions)}</div>`:'';
  const closeBtnHtml=fromTasksTab?`<button onclick="closeProofUploadScreen()" style="background:transparent;border:none;color:#fff;opacity:.8;font-size:13px;cursor:pointer;padding:2px 4px">✕ Close</button>`:'';
  screen.innerHTML=`<div style="width:min(540px,100%);background:#fff;border:1px solid #e5e7eb;border-radius:10px;box-shadow:0 8px 24px rgba(15,23,42,.09);overflow:hidden">
    <div style="padding:16px 20px;border-bottom:1px solid #e5e7eb;background:#0E3416;color:#fff;display:flex;align-items:flex-start;justify-content:space-between;gap:10px">
      <div>
      <div style="font-size:15px;font-weight:800;margin-bottom:2px">Submit Completion Proof</div>
      <div style="font-size:11.5px;opacity:.8">${escapeHtml(params.title||'Assigned task')}</div>
      </div>
      ${closeBtnHtml}
    </div>
    <div style="padding:18px 20px">
      <div id="proof-history-area"></div>
      <div id="proof-result-area"></div>
      <div id="proof-followup-area"></div>
      <div id="proof-form-area">
        <div style="font-size:11.5px;color:#6b7280;margin-bottom:12px">Signed in as <strong style="color:#374151">${escapeHtml(currentUser?.email||'')}</strong></div>
        ${instrHtml}
        <div style="border:1px solid #d1d5db;border-radius:8px;overflow:hidden">
          <textarea id="proof-note" placeholder="Describe what was completed, include any relevant details..." style="display:block;width:100%;min-height:90px;border:none;padding:12px 14px;font-family:Inter,sans-serif;font-size:13px;resize:vertical;box-sizing:border-box;outline:none;color:#111;line-height:1.5"></textarea>
          <div style="border-top:1px solid #e5e7eb;padding:8px 10px;display:flex;align-items:center;gap:8px;background:#fafafa;flex-wrap:wrap">
            <input id="proof-files" type="file" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.zip" style="display:none">
            <button type="button" onclick="document.getElementById('proof-files').click()" style="display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border:1px solid #d1d5db;border-radius:5px;background:#fff;color:#374151;font-size:11.5px;font-weight:600;cursor:pointer">📎 Attach files</button>
            <div id="proof-file-list" style="flex:1;font-size:11px;color:#374151;min-width:0"></div>
          </div>
        </div>
        <div id="proof-progress" style="display:none;margin-top:10px">
          <div style="font-size:11px;color:#6b7280;margin-bottom:4px" id="proof-progress-label">Uploading...</div>
          <div style="background:#e5e7eb;border-radius:4px;height:5px"><div id="proof-progress-bar" style="background:#0E3416;height:5px;border-radius:4px;transition:width .2s;width:0%"></div></div>
        </div>
        <div id="proof-status" style="margin-top:8px;font-size:12px"></div>
        <div style="display:flex;gap:8px;align-items:center;margin-top:12px">
          <button class="btn btn-primary btn-sm" onclick="uploadProofFiles()" id="proof-upload-btn" style="background:#0E3416;border-color:#0E3416">Submit Proof</button>
          <button class="btn btn-ghost btn-sm" onclick="signOut()" style="color:#9ca3af;border-color:#e5e7eb">Sign out</button>
        </div>
      </div>
    </div>
  </div>`;
  document.body.appendChild(screen);
  pendingProofFiles=[];
  document.getElementById('proof-files')?.addEventListener('change',function(){
    const newFiles=Array.from(this.files||[]);
    newFiles.forEach(f=>{
      if(!pendingProofFiles.some(p=>p.name===f.name&&p.size===f.size))pendingProofFiles.push(f);
    });
    this.value='';
    renderPendingProofFileList();
  });
  // Check for existing proof result
  if(params.appTaskId){
    await checkAndLoadProofNotifications().catch(()=>{});
    const prior=window._proofSubmissionHistory?.[String(params.appTaskId)]||[];
    const historyArea=document.getElementById('proof-history-area');
    if(historyArea&&prior.length){
      historyArea.innerHTML=`<div style="margin-bottom:14px"><div style="font-size:11px;font-weight:800;color:var(--body);margin-bottom:6px">Previous proof submission${prior.length===1?'':'s'}</div>${renderNotificationProofs({...prior[0],taskId:String(params.appTaskId)})}</div>`;
    }
    await renderProofFollowupBox(params);
    const result=await checkProofResult(params.appTaskId,currentUser?.email||'');
    const resultArea=document.getElementById('proof-result-area');
    if(result&&resultArea){
      if(result.result==='approved'){
        resultArea.innerHTML=`<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:7px;padding:14px 16px;margin-bottom:16px"><div style="font-size:14px;font-weight:700;color:#166534;margin-bottom:4px">✓ Proof Approved</div><div style="font-size:12px;color:#166534">Approved by ${escapeHtml(result.senderName||result.senderEmail||'the assignor')}. Task is complete.</div></div>`;
        document.getElementById('proof-form-area').style.display='none';
      }else if(result.result==='declined'){
        resultArea.innerHTML=`<div style="background:#fff1f2;border:1px solid #fca5a5;border-radius:7px;padding:14px 16px;margin-bottom:16px"><div style="font-size:14px;font-weight:700;color:#b91c1c;margin-bottom:4px">✗ Proof Declined — Please Resubmit</div><div style="font-size:12px;color:#b91c1c"><strong>Reason:</strong> ${escapeHtml(result.reason||'No reason provided')}</div><div style="font-size:11px;color:#9ca3af;margin-top:4px">Address the feedback above and upload updated proof below.</div></div>`;
      }
    }
  }
}

function closeProofUploadScreen(){
  proofParamsOverride=null;
  document.getElementById('proof-upload-screen')?.remove();
  renderMyTasks();
}

async function loadTasksFromOneDrive() {
  if(window.DPEG_STAGING_MODE){
    tasks=[];archives=[];staffConfig={};customDepartments=[];customNotes=[];notifications=[];
    sharedDataActive=false;
    setSyncStatus('synced','D1 staging · Live');
    finishDataLoad();
    setTimeout(()=>nav('tasks'),0);
    return;
  }
  setSyncStatus("syncing", "Loading tasks...");
  try {
    sharedDataActive=false;
    const loaded=await loadLegacyOneDriveData();
    let savedNewFile=true;
    if(!loaded){
      tasks=[];archives=[];staffConfig={};customDepartments=[];customNotes=[];notifications=[];
      savedNewFile=await saveTasksToOneDrive();
    }
    await loadSharedDepartmentSettings();
    if(loaded||savedNewFile)setSyncStatus("synced", "Synced with OneDrive");
    if(loaded)await compareCanaryD1Tasks(tasks);
    finishDataLoad();
  } catch (err) {
    sharedDataActive=false;
    console.error("OneDrive load error:", err);
    try{
      const backup=localStorage.getItem(`dpeg_local_task_backup_${normEmail(currentUser?.email||'unknown')}`);
      if(backup)applyLoadedData(JSON.parse(backup));
    }catch{}
    setSyncStatus("error","OneDrive unavailable • tasks still online");
    toast("OneDrive is unavailable. Shared tasks and messages are still online.");
    refreshAll();
    renderCharts();
    renderActivity();
    autoSyncContacts();
  }
}

async function compareCanaryD1Tasks(legacyTasks){
  const base=workerBaseUrl();
  if(!base||!currentUser?.email)return {enabled:false};
  try{
    const token=await getAccessToken();
    const res=await fetch(`${base}/shared-workflow-read`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error(`D1 canary read failed (${res.status})`);
    let result=await res.json();
    if(!result.enabled)return result;
    // During the protected canary, refresh this user's shadow rows from the
    // already-loaded OneDrive source before comparing. This backfills newly
    // added metadata without changing the visible legacy data or OneDrive.
    await shadowSyncTasksToD1(legacyTasks);
    const refreshed=await fetch(`${base}/shared-workflow-read`,{headers:{Authorization:`Bearer ${token}`}});
    if(!refreshed.ok)throw new Error(`D1 canary refresh failed (${refreshed.status})`);
    result=await refreshed.json();
    const legacyById=new Map((Array.isArray(legacyTasks)?legacyTasks:[]).map(task=>[String(task.id||''),task]));
    const d1ById=new Map((Array.isArray(result.tasks)?result.tasks:[]).map(task=>[String(task.id||''),task]));
    const missingInD1=[...legacyById.keys()].filter(id=>id&&!d1ById.has(id));
    const extraInD1=[...d1ById.keys()].filter(id=>id&&!legacyById.has(id));
    const changed=[],differenceDetails=[];
    const fieldChanges={title:[],status:[],summary:[],taskInstruction:[],proofInstructions:[],department:[],priority:[],dueDate:[],completedAt:[],cancelledAt:[],sourceMessageId:[],sourceConversationId:[]};
    for(const [id,legacy] of legacyById){
      const d1=d1ById.get(id);if(!d1)continue;
      const canonicalStatus=value=>{
        const status=String(value||'Pending').trim().toLowerCase();
        // Overdue is a date-derived display condition, not a separate
        // workflow state. D1 correctly stores the underlying state Pending.
        if(status==='assigned'||status==='pending'||status==='overdue')return 'pending';
        if(status==='completed')return 'done';
        if(status==='inprogress')return 'in progress';
        if(status==='canceled')return 'cancelled';
        return status;
      };
      const legacyStatus=canonicalStatus(legacy.status);
      const d1Status=canonicalStatus(d1.status);
      // OneDrive/Outlook titles can contain non-breaking spaces or repeated
      // whitespace that render identically. Compare their visible text while
      // preserving the original title in both stores.
      const canonicalTitle=value=>String(value||'').normalize('NFKC').replace(/\s+/g,' ').trim();
      const titleDiff=canonicalTitle(legacy.title)!==canonicalTitle(d1.title);
      const statusDiff=legacyStatus!==d1Status;
      const canonicalText=value=>String(value||'').normalize('NFKC').replace(/\s+/g,' ').trim();
      const canonicalLower=value=>canonicalText(value).toLowerCase();
      const canonicalDate=value=>{
        if(!value)return '';
        const parsed=new Date(value);
        return Number.isNaN(parsed.getTime())?canonicalText(value):parsed.toISOString().slice(0,10);
      };
      const comparisons={
        title:{legacy:String(legacy.title||''),d1:String(d1.title||''),different:titleDiff},
        status:{legacy:String(legacy.status||'Pending'),d1:String(d1.status||'Pending'),different:statusDiff},
        summary:{legacy:String(legacy.summary||legacy.description||''),d1:String(d1.summary||''),different:canonicalText(legacy.summary||legacy.description)!==canonicalText(d1.summary)},
        taskInstruction:{legacy:String(legacy.taskInstruction||''),d1:String(d1.taskInstruction||''),different:canonicalText(legacy.taskInstruction)!==canonicalText(d1.taskInstruction)},
        proofInstructions:{legacy:String(legacy.proofInstructions||''),d1:String(d1.proofInstructions||''),different:canonicalText(legacy.proofInstructions)!==canonicalText(d1.proofInstructions)},
        department:{legacy:String(legacy.dept||legacy.department||'Needs Department'),d1:String(d1.dept||'Needs Department'),different:canonicalLower(legacy.dept||legacy.department||'Needs Department')!==canonicalLower(d1.dept||'Needs Department')},
        priority:{legacy:String(legacy.priority||'Normal'),d1:String(d1.priority||'Normal'),different:canonicalLower(legacy.priority||'Normal')!==canonicalLower(d1.priority||'Normal')},
        dueDate:{legacy:String(Object.prototype.hasOwnProperty.call(legacy,'deadline')?(legacy.deadline||''):(legacy.date||legacy.dueDate||'')),d1:String(d1.date||''),different:canonicalDate(Object.prototype.hasOwnProperty.call(legacy,'deadline')?(legacy.deadline||''):(legacy.date||legacy.dueDate||''))!==canonicalDate(d1.date)},
        completedAt:{legacy:String(legacy.completedAt||legacy.approvedAt||''),d1:String(d1.completedAt||''),different:canonicalDate(legacy.completedAt||legacy.approvedAt)!==canonicalDate(d1.completedAt)},
        cancelledAt:{legacy:String(legacy.cancelledAt||''),d1:String(d1.cancelledAt||''),different:canonicalDate(legacy.cancelledAt)!==canonicalDate(d1.cancelledAt)},
        sourceMessageId:{legacy:String(legacy.lastMessageId||legacy.emailId||''),d1:String(d1.sourceMessageId||''),different:canonicalText(legacy.lastMessageId||legacy.emailId)!==canonicalText(d1.sourceMessageId)},
        sourceConversationId:{legacy:String(legacy.conversationId||''),d1:String(d1.sourceConversationId||''),different:canonicalText(legacy.conversationId)!==canonicalText(d1.sourceConversationId)},
      };
      const fieldDetails={};
      Object.entries(comparisons).forEach(([field,comparison])=>{
        if(!comparison.different)return;
        fieldChanges[field].push(id);
        fieldDetails[field]={legacy:comparison.legacy,d1:comparison.d1};
      });
      const hasDifference=Object.keys(fieldDetails).length>0;
      if(hasDifference)differenceDetails.push({
        id,
        task:String(legacy.title||d1.title||'Task'),
        fields:fieldDetails,
      });
      if(hasDifference)changed.push(id);
    }
    const report={
      checkedAt:new Date().toISOString(),enabled:true,
      legacyCount:legacyById.size,d1Count:d1ById.size,
      missingInD1:missingInD1.slice(0,100),extraInD1:extraInD1.slice(0,100),changed:changed.slice(0,100),
      fieldChanges:Object.fromEntries(Object.entries(fieldChanges).map(([field,ids])=>[field,ids.slice(0,100)])),
      differenceDetails:differenceDetails.slice(0,20),
      safeToRender:missingInD1.length===0&&changed.length===0,
    };
    try{localStorage.setItem('dpeg_d1_canary_report',JSON.stringify(report));}catch{}
    console.info('D1 canary comparison',report);
    const changedFieldSummary=Object.entries(fieldChanges).filter(([,ids])=>ids.length).map(([field,ids])=>`${ids.length} ${field}`).join(' · ');
    const mismatchSummary=`${missingInD1.length} missing${changedFieldSummary?` · ${changedFieldSummary}`:''} · ${extraInD1.length} extra`;
    setSyncStatus('synced',report.safeToRender?'D1 canary verified · Legacy view':`D1 mismatch (${mismatchSummary}) · Legacy protected`);
    if(result.visibleRead&&report.safeToRender){
      tasks=Array.isArray(result.tasks)?result.tasks:tasks;
      sharedDataActive=true;
      setSyncStatus('synced','D1 controlled view');
    }
    const syncLabel=document.getElementById('sync-label');
    if(syncLabel&&!report.safeToRender){
      syncLabel.title='Click to view the protected D1 comparison details';
      syncLabel.style.cursor='pointer';
      syncLabel.onclick=()=>{
        const lines=differenceDetails.map(detail=>{
          const changes=Object.entries(detail.fields).map(([field,value])=>`${field}\nLegacy: ${value.legacy||'(empty)'}\nD1: ${value.d1||'(empty)'}`);
          return `${detail.task}\n${changes.join('\n')}`;
        });
        alert(`D1 canary differences\n\n${lines.join('\n\n')}`);
      };
    }
    return report;
  }catch(error){
    console.warn('D1 canary comparison skipped:',error.message);
    return {enabled:false,error:error.message};
  }
}

async function loadLegacyOneDriveData() {
    const token = await getAccessToken();
    const folder = currentUser.folder;
    const path = `/me/drive/root:/${ONEDRIVE_FOLDER}/${folder}/tasks.json:/content`;
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    // A genuine 404 means this user has no task file yet. Authentication,
    // permission, throttling, and Graph service errors must not be treated as
    // an empty account, because loadTasksFromOneDrive would otherwise attempt
    // to initialize and save an empty task file.
    if (res.status===404)return false;
    if (!res.ok){
      const detail=await res.text().catch(()=>'');
      throw new Error(`OneDrive task load failed (${res.status})${detail?`: ${detail.slice(0,180)}`:''}`);
    }
    applyLoadedData(JSON.parse(await res.text()));
    return true;
}

async function ensureLegacyOneDriveFolder(token) {
  async function ensureFolder(path, createEndpoint, name) {
    const check = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (check.ok) return true;
    const create = await fetch(`https://graph.microsoft.com/v1.0${createEndpoint}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "replace" })
    });
    return create.ok;
  }
  if (!await ensureFolder(`/me/drive/root:/${ONEDRIVE_FOLDER}`, "/me/drive/root/children", ONEDRIVE_FOLDER)) return false;
  return ensureFolder(`/me/drive/root:/${ONEDRIVE_FOLDER}/${currentUser.folder}`, `/me/drive/root:/${ONEDRIVE_FOLDER}:/children`, currentUser.folder);
}

function finishDataLoad(){
  initSelects();
  refreshAll();
  renderCharts();
  renderActivity();
  autoSyncContacts();
  updateNotifBadge();
  startNotifPolling();
  pollToDoCompletions().catch(()=>{});
  checkAndLoadProofNotifications().catch(()=>{});
}

function companyDataIsEmpty(){
  return !tasks.length && !archives.length && !Object.keys(staffConfig||{}).length && !(customDepartments||[]).length && !customNotes.length;
}

async function migrateLegacyOneDriveToShared(){
  if(!await loadLegacyOneDriveData())return false;
  if(companyDataIsEmpty())return false;
  if(!await saveSharedCompanyData())return false;
  sharedDataActive=true;
  setSyncStatus("synced","Migrated and synced company data");
  finishDataLoad();
  return true;
}

async function saveTasksToOneDrive() {
  if(window.DPEG_STAGING_MODE){
    setSyncStatus('synced','D1 staging · Live');
    return true;
  }
  setSyncStatus("syncing", "Saving...");
  const backupKey=`dpeg_local_task_backup_${normEmail(currentUser?.email||'unknown')}`;
  const payload=JSON.stringify({tasks,archives,staffConfig,customDepartments,customNotes,notifications},null,2);
  try{localStorage.setItem(backupKey,payload);}catch{}
  try {
    const token = await getAccessToken();
    if(!await ensureLegacyOneDriveFolder(token))throw new Error('OneDrive folder could not be created');
    const path = `/me/drive/root:/${ONEDRIVE_FOLDER}/${currentUser.folder}/tasks.json:/content`;
    let res;
    for(let attempt=0;attempt<2;attempt++){
      res=await fetch(`https://graph.microsoft.com/v1.0${path}`,{
        method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:payload
      });
      if(res.ok||![429,500,502,503,504].includes(res.status))break;
      await new Promise(resolve=>setTimeout(resolve,700));
    }
    if (!res.ok) throw new Error(`OneDrive save failed (${res.status})`);
    sharedDataActive=false;
    setSyncStatus("synced","Saved to OneDrive");
    await shadowSyncTasksToD1(tasks);
    return true;
  } catch (err) {
    setSyncStatus("error", "OneDrive unavailable • tasks still online");
    console.error("OneDrive save error:", err);
    return false;
  }
}

async function shadowSyncTasksToD1(sourceTasks){
  const base=workerBaseUrl();
  if(!base||!currentUser?.email)return {enabled:false,written:0};
  // Only the OneDrive owner/assigner may shadow a task. Recipient copies are
  // deliberately excluded so they cannot overwrite the assigner's row.
  const allTasks=Array.isArray(sourceTasks)?sourceTasks:[];
  let written=0,enabled=false;
  try{
    const token=await getAccessToken();
    for(let offset=0;offset<allTasks.length;offset+=100){
      const res=await fetch(`${base}/shared-workflow-sync`,{
        method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({tasks:allTasks.slice(offset,offset+100)}),
      });
      if(!res.ok)throw new Error(`Shadow sync failed (${res.status})`);
      const result=await res.json().catch(()=>({}));
      enabled=!!result.enabled;
      written+=Number(result.written||0);
      if(!enabled)break;
    }
    const finalRes=await fetch(`${base}/shared-workflow-sync`,{
      method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({tasks:[],finalizeIds:allTasks.map(task=>String(task.id||'')).filter(Boolean)}),
    });
    if(!finalRes.ok)throw new Error(`Shadow finalize failed (${finalRes.status})`);
    const finalResult=await finalRes.json().catch(()=>({}));
    enabled=!!finalResult.enabled;
  }catch(error){
    // OneDrive has already saved successfully. Shadow-sync failures are
    // observable but must never roll back or disguise that primary save.
    console.warn('D1 shadow sync skipped:',error.message);
  }
  return {enabled,written};
}

function applyLoadedData(data){
  if(Array.isArray(data)){
    tasks=data;archives=[];staffConfig={};customDepartments=[];customNotes=[];
  }else if(data?.tasks){
    tasks=data.tasks||[];
    archives=data.archives||[];
    staffConfig=data.staffConfig||{};
    customDepartments=Array.isArray(data.customDepartments)?data.customDepartments.filter(Boolean):[];
    customNotes=data.customNotes||[];
    notifications=Array.isArray(data.notifications)?data.notifications:notifications||[];
  }else if(data?.id){
    tasks=[data];archives=[];staffConfig={};customDepartments=[];customNotes=[];
  }else{
    tasks=[];archives=[];staffConfig={};customDepartments=[];customNotes=[];
  }
}

async function loadSharedCompanyData(){
  const base=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  if(!base)return false;
  try{
    const token=await getAccessToken();
    const res=await fetch(`${base}/data`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)return false;
    const data=await res.json();
    applyLoadedData(data);
    sharedDataVersion=data?.updatedAt||null;
    sharedDataActive=true;
    setSyncStatus("synced","Synced company data");
    finishDataLoad();
    return true;
  }catch(err){
    console.warn("Shared data load skipped:",err.message);
    return false;
  }
}

async function saveSharedCompanyData(){
  const base=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  if(!base)return false;
  try{
    const token=await getAccessToken();
    const payload=JSON.stringify({tasks,archives,staffConfig,customDepartments,customNotes,notifications,baseUpdatedAt:sharedDataVersion});
    const res=await fetch(`${base}/data`,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:payload});
    if(res.status===409){
      setSyncStatus("error","Someone else made changes — reload to see them, then retry");
      console.warn("Shared data save conflict: local view is stale, refusing to overwrite.");
      return false;
    }
    if(!res.ok)return false;
    const result=await res.json().catch(()=>null);
    sharedDataVersion=result?.updatedAt||sharedDataVersion;
    sharedDataActive=true;
    setSyncStatus("synced","Saved company data");
    return true;
  }catch(err){
    console.warn("Shared data save skipped:",err.message);
    return false;
  }
}

async function loadSharedDepartmentSettings(){
  const base=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  if(!base)return false;
  try{
    const token=await getAccessToken();
    const res=await fetch(`${base}/departments`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)return false;
    const data=await res.json();
    const sharedList=Array.isArray(data.departments)?data.departments.filter(Boolean):[];
    // The server registry is authoritative. Personal OneDrive files may still
    // contain older department choices, but they must not differ by user.
    customDepartments=[...new Set(sharedList)];
    const sharedAssignments=data.assignments||{};
    Object.values(staffConfig||{}).forEach(person=>{
      const email=normEmail(person?.email||'');
      if(email&&isInternalEmail(email)&&!sharedAssignments[email])person.dept='Needs Department';
    });
    Object.values(sharedAssignments).forEach(row=>{
      const email=normEmail(row?.email||'');
      if(!email||!row?.dept)return;
      const key=staffKey(email,row.name||email);
      staffConfig[key]={...(staffConfig[key]||{}),name:row.name||staffConfig[key]?.name||email.split('@')[0],email,dept:row.dept,role:staffConfig[key]?.role||''};
    });
    tasks.forEach(task=>{
      const email=normEmail(task?.email||'');
      if(!email||!isInternalEmail(email)||nstt(task.status)==='Done')return;
      task.dept=sharedAssignments[email]?.dept||'Needs Department';
    });
    sharedDepartmentsVersion=data.updatedAt||null;
    return true;
  }catch(err){console.warn('Shared departments load skipped:',err.message);return false;}
}

async function saveSharedDepartmentSettings(){
  if(!isAdmin())return false;
  const base=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  if(!base)return false;
  const assignments={};
  Object.values(staffConfig||{}).forEach(p=>{
    const email=normEmail(p?.email||'');
    const dept=String(configuredDept(email,p?.name)||'').trim();
    if(email&&isInternalEmail(email)&&dept&&dept!=='Needs Department')assignments[email]={email,name:p.name||'',dept};
  });
  try{
    const token=await getAccessToken();
    const res=await fetch(`${base}/departments`,{
      method:'PUT',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({departments:customDepartments,assignments,baseUpdatedAt:sharedDepartmentsVersion})
    });
    if(res.status===409){toast('Departments changed elsewhere. Reopen Department Settings and try again.');return false;}
    if(!res.ok)return false;
    const data=await res.json().catch(()=>({}));
    sharedDepartmentsVersion=data.updatedAt||sharedDepartmentsVersion;
    return true;
  }catch(err){console.warn('Shared departments save skipped:',err.message);return false;}
}

async function saveSharedDepartmentAssignment(email,name,dept){
  const normalized=normEmail(email);
  if(!normalized||!isInternalEmail(normalized)||!dept||dept==='Needs Department')return false;
  const base=(localStorage.getItem('dpeg_ai_fn_url')||WORKER_URL).replace(/\/?$/,'');
  if(!base)return false;
  window.lastDepartmentSaveError='';
  try{
    const token=await getAccessToken();
    const res=await fetch(`${base}/department-assignment`,{
      method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({assignment:{email:normalized,name:String(name||'').trim(),dept}})
    });
    if(!res.ok){
      const data=await res.json().catch(()=>({}));
      throw new Error(data.error||`Department save failed (${res.status})`);
    }
    const data=await res.json();
    if(!data.success||!data.assignment)throw new Error('The department service is not updated yet');
    sharedDepartmentsVersion=data.updatedAt||sharedDepartmentsVersion;
    return true;
  }catch(err){
    window.lastDepartmentSaveError=err.message||'Department save failed';
    console.warn('Shared department assignment save failed:',err.message);
    return false;
  }
}
