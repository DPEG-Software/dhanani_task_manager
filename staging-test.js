const STAGING_WORKER='https://dpeg-task-manager-staging.systemmanager1.workers.dev';
const TENANT_ID='9152bf5c-22ff-4e4a-8624-784a2d243006';
const CLIENT_ID='8d523e65-0163-49c7-881b-407c0222527e';
const REDIRECT_URI=window.location.origin+window.location.pathname;
const loginRequest={scopes:['User.Read'],redirectUri:REDIRECT_URI,prompt:'select_account'};
let authClient=null,account=null,workflow={tasks:[],assignments:[],proofs:[],reminders:[],messageThreads:[]};
let refreshInFlight=false,refreshQueued=false,autoRefreshTimer=null;
let realtimeSocket=null,reconnectTimer=null,reconnectAttempt=0;
const expandedThreads=new Set();

const $=id=>document.getElementById(id);
const esc=value=>String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const norm=value=>String(value||'').trim().toLowerCase();

function setStatus(message,error=false){$('form-status').className='status '+(error?'error':'ok');$('form-status').textContent=message||'';}
async function token(){
  if(!account)throw new Error('Sign in first');
  try{return (await authClient.acquireTokenSilent({scopes:['User.Read'],account})).accessToken;}
  catch{return (await authClient.acquireTokenPopup({scopes:['User.Read'],account,redirectUri:REDIRECT_URI})).accessToken;}
}
async function api(body=null){
  const accessToken=await token();
  const options={headers:{Authorization:`Bearer ${accessToken}`}};
  if(body){options.method='POST';options.headers['Content-Type']='application/json';options.body=JSON.stringify(body);}
  const response=await fetch(`${STAGING_WORKER}/staging/tasks`,options);
  const data=await response.json().catch(()=>({}));
  if(!response.ok)throw new Error(data.message||data.error||`Request failed (${response.status})`);
  return data;
}
async function checkSafety(){
  try{
    const response=await fetch(`${STAGING_WORKER}/environment`);const data=await response.json();
    const safe=data.environment==='staging'&&data.externalEffectsEnabled===false;
    $('safety').className='safety '+(safe?'':'bad');
    $('safety').textContent=safe?'✓ Safe staging environment confirmed. Email, Microsoft To Do, OneDrive and AI effects are disabled.':'STOP: staging safety could not be confirmed.';
    $('delegate').disabled=!safe||!account;$('refresh').disabled=!safe||!account;
  }catch{$('safety').className='safety bad';$('safety').textContent='STOP: could not reach the staging environment.';}
}
async function signIn(){await authClient.loginRedirect(loginRequest);}
async function signOut(){if(account)await authClient.logoutRedirect({account,postLogoutRedirectUri:REDIRECT_URI});}

async function loadWorkflow(options={}){
  if(!account)return;
  // A realtime event can arrive while an earlier snapshot is still loading.
  // Never discard that event: the earlier snapshot may have been read between
  // two concurrent message commits. Coalesce any overlapping events into one
  // guaranteed follow-up read after the current request finishes.
  if(refreshInFlight){refreshQueued=true;return;}
  refreshInFlight=true;
  if(!options.silent)setStatus('Loading…');
  try{
    workflow=await api();render();
    const time=new Date().toLocaleTimeString([],{hour:'numeric',minute:'2-digit',second:'2-digit'});
    if($('sync-status'))$('sync-status').textContent=`Auto-synced ${time}`;
    if(!options.silent)setStatus('Staging data refreshed.');
  }catch(error){
    if($('sync-status'))$('sync-status').textContent='Sync paused';
    if(!options.silent)setStatus(error.message,true);
  }finally{
    refreshInFlight=false;
    if(refreshQueued){
      refreshQueued=false;
      queueMicrotask(()=>loadWorkflow({silent:true}));
    }
  }
}
function startAutoRefresh(){
  clearInterval(autoRefreshTimer);
  autoRefreshTimer=setInterval(()=>{
    if(account&&!document.hidden)loadWorkflow({silent:true});
  },30000);
}
async function connectRealtime(){
  if(!account||realtimeSocket?.readyState===WebSocket.OPEN||realtimeSocket?.readyState===WebSocket.CONNECTING)return;
  clearTimeout(reconnectTimer);
  try{
    const accessToken=await token();
    const ticketResponse=await fetch(`${STAGING_WORKER}/staging/realtime-ticket`,{
      method:'POST',headers:{Authorization:`Bearer ${accessToken}`}
    });
    const ticketData=await ticketResponse.json().catch(()=>({}));
    if(!ticketResponse.ok)throw new Error(ticketData.error||'Could not create realtime ticket');
    const socketUrl=STAGING_WORKER.replace(/^https:/,'wss:').replace(/^http:/,'ws:')+`/staging/realtime?ticket=${encodeURIComponent(ticketData.ticket)}`;
    realtimeSocket=new WebSocket(socketUrl);
    realtimeSocket.onopen=()=>{
      reconnectAttempt=0;
      if($('sync-status'))$('sync-status').textContent='Live connected';
    };
    realtimeSocket.onmessage=event=>{
      if(event.data==='pong')return;
      try{
        const update=JSON.parse(event.data);
        if(update.type==='workflow_changed')loadWorkflow({silent:true});
      }catch{}
    };
    realtimeSocket.onclose=()=>{
      realtimeSocket=null;
      if($('sync-status'))$('sync-status').textContent='Live reconnecting…';
      scheduleRealtimeReconnect();
    };
    realtimeSocket.onerror=()=>realtimeSocket?.close();
  }catch{
    if($('sync-status'))$('sync-status').textContent='Live reconnecting…';
    scheduleRealtimeReconnect();
  }
}
function scheduleRealtimeReconnect(){
  if(!account)return;
  clearTimeout(reconnectTimer);
  const delay=Math.min(15000,1000*(2**Math.min(reconnectAttempt++,4)));
  reconnectTimer=setTimeout(connectRealtime,delay);
}
function relatedActivity(a){
  const proofs=(workflow.proofs||[]).filter(p=>p.assignment_id===a.id);
  const reminders=(workflow.reminders||[]).filter(r=>r.assignment_id===a.id);
  const thread=(workflow.messageThreads||[]).find(t=>t.appTaskId===a.app_task_id&&norm(t.recipientEmail)===norm(a.recipient_email));
  const lines=[];
  if(reminders.length)lines.push(`<div>${reminders.length} reminder${reminders.length===1?'':'s'}</div>`);
  if(thread?.thread?.length){
    const messages=thread.thread;
    const expanded=expandedThreads.has(a.id);
    const messageRows=expanded?messages.map(message=>{
      const mine=norm(message.email)===norm(account?.username);
      const sender=mine?'You':message.name||message.email||'Employee';
      const time=message.createdAt?new Date(message.createdAt).toLocaleString([],{
        month:'short',day:'numeric',hour:'numeric',minute:'2-digit'
      }):'';
      return `<div class="thread-message${mine?' mine':''}"><div class="thread-meta">${esc(sender)}${time?` · ${esc(time)}`:''}</div><div>${esc(message.message||'')}</div></div>`;
    }).join(''):'';
    lines.push(`<div class="thread-summary"><button type="button" class="thread-toggle" onclick="toggleThread('${a.id}')">${messages.length} message${messages.length===1?'':'s'} · ${expanded?'Hide conversation':'View conversation'}</button></div>${expanded?`<div class="thread-list">${messageRows}</div>`:''}`);
  }
  proofs.forEach(p=>lines.push(`<div>Proof: <b>${esc(p.status)}</b>${p.note?` — ${esc(p.note)}`:''}</div>`));
  return {proofs,html:lines.length?`<div class="activity">${lines.join('')}</div>`:''};
}
function toggleThread(assignmentId){
  if(expandedThreads.has(assignmentId))expandedThreads.delete(assignmentId);
  else expandedThreads.add(assignmentId);
  render();
}
function card(a,role){
  const activity=relatedActivity(a);const version=Number(a.version||1);const pending=activity.proofs.find(p=>p.status==='pending');
  const buttons=[`<button onclick="sendMessage('${a.id}')">Message</button>`];
  if(role==='assigner'){
    buttons.push(`<button onclick="sendReminder('${a.id}',${version})">Send Reminder</button>`);
    if(pending){buttons.push(`<button class="primary" onclick="reviewProof('${pending.id}',${version},'approved')">Approve Proof</button>`,`<button class="danger" onclick="reviewProof('${pending.id}',${version},'changes_requested')">Request Changes</button>`);}
  }else{
    if(a.status!=='Submitted'&&a.status!=='Done')buttons.push(`<button onclick="changeStatus('${a.id}',${version},'In Progress')">Start</button>`,`<button class="primary" onclick="submitProof('${a.id}',${version})">Submit Fake Proof</button>`);
  }
  const person=role==='assigner'?a.recipient_name||a.recipient_email:a.assigner_name||a.assigner_email;
  return `<article class="card"><div class="card-head"><div><div class="title">${esc(a.title)}</div><div class="meta">${esc(person)} · ${esc(a.dept||'Needs Department')}${a.due_date?` · Due ${esc(a.due_date)}`:''}</div></div><span class="badge">${esc(a.status)}</span></div>${activity.html}<div class="actions">${buttons.join('')}</div></article>`;
}
function render(){
  const email=norm(account?.username);const assignments=workflow.assignments||[];
  const by=assignments.filter(a=>norm(a.assigner_email)===email);const to=assignments.filter(a=>norm(a.recipient_email)===email);
  $('by-me').innerHTML=by.length?by.map(a=>card(a,'assigner')).join(''):'<div class="empty">No fake tasks assigned by you.</div>';
  $('to-me').innerHTML=to.length?to.map(a=>card(a,'recipient')).join(''):'<div class="empty">No fake tasks assigned to you.</div>';
}
async function action(payload,success){try{setStatus('Saving…');await api(payload);setStatus(success);await loadWorkflow();}catch(error){setStatus(error.message,true);await loadWorkflow();}}
async function delegate(){
  const recipientEmail=$('recipient').value.trim(),title=$('title').value.trim();
  if(!recipientEmail||!title)return setStatus('Recipient email and task title are required.',true);
  await action({action:'delegate',recipientEmail,title,summary:$('summary').value,departmentName:$('department').value,priority:$('priority').value,dueDate:$('due-date').value},'Fake task created.');
}
async function sendMessage(assignmentId){const message=prompt('Type a staging-only message:');if(message?.trim())await action({action:'message',assignmentId,message:message.trim()},'Message saved in staging D1.');}
async function sendReminder(assignmentId,expectedVersion){await action({action:'remind',assignmentId,expectedVersion,idempotencyKey:`ui-rem-${assignmentId}-${crypto.randomUUID()}`},'Reminder saved in staging D1. No email was sent.');}
async function changeStatus(assignmentId,expectedVersion,status){await action({action:'assignment_status',assignmentId,expectedVersion,status},'Status updated.');}
async function submitProof(assignmentId,expectedVersion){const note=prompt('Describe the fake proof:','Staging proof completed.');if(note!==null)await action({action:'submit_proof',assignmentId,expectedVersion,note,idempotencyKey:`ui-proof-${assignmentId}-${crypto.randomUUID()}`,files:[{fileName:'staging-proof.txt',mimeType:'text/plain',sizeBytes:10,webUrl:'about:blank'}]},'Fake proof submitted.');}
async function reviewProof(proofId,expectedVersion,decision){const reason=decision==='changes_requested'?prompt('Reason for requesting changes:','Please update the proof.'):'';if(reason!==null)await action({action:'review_proof',proofId,expectedVersion,decision,reason:reason||''},decision==='approved'?'Proof approved.':'Changes requested.');}
Object.assign(window,{sendMessage,sendReminder,changeStatus,submitProof,reviewProof,toggleThread});

async function init(){
  authClient=new msal.PublicClientApplication({auth:{clientId:CLIENT_ID,authority:`https://login.microsoftonline.com/${TENANT_ID}`,redirectUri:REDIRECT_URI,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage',storeAuthStateInCookie:true}});
  await authClient.initialize();const response=await authClient.handleRedirectPromise();account=response?.account||authClient.getAllAccounts()[0]||null;
  $('sign-in').hidden=!!account;$('sign-out').hidden=!account;$('account').textContent=account?`${account.name||''} (${account.username})`:'Not signed in';
  $('sign-in').onclick=signIn;$('sign-out').onclick=signOut;$('delegate').onclick=delegate;$('refresh').onclick=loadWorkflow;
  document.addEventListener('visibilitychange',()=>{if(account&&!document.hidden)loadWorkflow({silent:true});});
  await checkSafety();if(account){await loadWorkflow();startAutoRefresh();await connectRealtime();}
}
init().catch(error=>setStatus(error.message,true));
