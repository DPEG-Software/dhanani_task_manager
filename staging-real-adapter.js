(function(){
  if(!window.DPEG_STAGING_MODE)return;

  const base=window.DPEG_STAGING_WORKER.replace(/\/$/,'');
  let socket=null,reconnectTimer=null,reconnectAttempt=0,refreshQueued=false,refreshing=false;
  let reviewContext=null;

  async function stagingToken(){
    if(!currentAccount)throw new Error('Sign in first');
    try{return (await msalInstance.acquireTokenSilent({scopes:['User.Read'],account:currentAccount})).accessToken;}
    catch{return (await msalInstance.acquireTokenPopup({scopes:['User.Read'],account:currentAccount})).accessToken;}
  }

  async function stagingApi(body=null){
    const token=await stagingToken();
    const options={headers:{Authorization:`Bearer ${token}`}};
    if(body){options.method='POST';options.headers['Content-Type']='application/json';options.body=JSON.stringify(body);}
    const response=await fetch(`${base}/staging/tasks`,options);
    const data=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(data.message||data.error||`Request failed (${response.status})`);
    return data;
  }

  async function refreshRealTasks(){
    if(refreshing){refreshQueued=true;return;}
    refreshing=true;
    try{
      await window.renderMyTasks?.(true);
      await window.checkAndLoadProofNotifications?.();
      setSyncStatus('synced','D1 staging · Live');
    }finally{
      refreshing=false;
      if(refreshQueued){refreshQueued=false;queueMicrotask(refreshRealTasks);}
    }
  }

  function scheduleReconnect(){
    clearTimeout(reconnectTimer);
    const delay=Math.min(15000,1000*(2**Math.min(reconnectAttempt++,4)));
    reconnectTimer=setTimeout(connectRealtime,delay);
  }

  async function connectRealtime(){
    if(!currentAccount||socket?.readyState===WebSocket.OPEN||socket?.readyState===WebSocket.CONNECTING)return;
    try{
      const token=await stagingToken();
      const ticketResponse=await fetch(`${base}/staging/realtime-ticket`,{method:'POST',headers:{Authorization:`Bearer ${token}`}});
      const ticket=await ticketResponse.json().catch(()=>({}));
      if(!ticketResponse.ok)throw new Error(ticket.error||'Realtime ticket failed');
      socket=new WebSocket(`${base.replace(/^https:/,'wss:')}/staging/realtime?ticket=${encodeURIComponent(ticket.ticket)}`);
      socket.onopen=()=>{reconnectAttempt=0;setSyncStatus('synced','D1 staging · Live');};
      socket.onmessage=event=>{
        if(event.data==='pong')return;
        try{if(JSON.parse(event.data).type==='workflow_changed')refreshRealTasks();}catch{}
      };
      socket.onclose=()=>{socket=null;setSyncStatus('syncing','D1 staging · Reconnecting');scheduleReconnect();};
      socket.onerror=()=>socket?.close();
    }catch{setSyncStatus('syncing','D1 staging · Reconnecting');scheduleReconnect();}
  }

  // Staging never sends email or changes Microsoft data from the browser.
  window.sendTaskFollowupEmail=async()=>{};
  window.sendTaskCancelledEmail=async()=>{};
  window.sendProofDeclineEmail=async()=>{};
  window.saveTasksToOneDrive=async()=>{setSyncStatus('synced','D1 staging · Live');return true;};
  window.openAdd=()=>toast('Create fake tasks from the D1 Workflow Test during this staging phase.');

  window.testStagingDualWrite=async function(){
    const button=document.getElementById('staging-dual-write-test');
    if(button)button.disabled=true;
    try{
      const email=String(currentUser?.email||'').toLowerCase();
      const token=await stagingToken();
      const response=await fetch(`${base}/shared-workflow-sync`,{
        method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({tasks:[
          {id:`dual-write-check-${email}`,title:'Dual-write staging check',summary:'Synthetic staging-only record',assignedByEmail:email,status:'Pending',updatedAt:new Date().toISOString()},
          {id:`dual-write-foreign-${email}`,title:'Must be rejected',assignedByEmail:'foreign-owner@dhananipeg.com',status:'Pending'},
        ]}),
      });
      const result=await response.json().catch(()=>({}));
      if(!response.ok)throw new Error(result.error||`Request failed (${response.status})`);
      if(!result.enabled)throw new Error('Staging dual-write flag is off');
      if(Number(result.skipped)!==1)throw new Error('Ownership guard did not reject exactly one task');
      toast(`Dual-write passed: ${result.written} owned write, ${result.skipped} foreign task rejected`);
    }catch(error){toast(`Dual-write test failed: ${error.message}`);}
    finally{if(button)button.disabled=false;}
  };

  window.openProofFromTasksTab=async function(id){
    try{
      const workflow=await stagingApi();
      const assignment=(workflow.assignments||[]).find(row=>row.id===id);
      if(!assignment)throw new Error('Assignment not found');
      const note=prompt('Describe what was completed:','Work completed for staging review.');
      if(note===null)return;
      await stagingApi({
        action:'submit_proof',assignmentId:id,expectedVersion:Number(assignment.version||1),
        note,idempotencyKey:`real-ui-proof-${id}-${crypto.randomUUID()}`,files:[],
      });
      toast('Proof submitted for review');
      await refreshRealTasks();
    }catch(error){toast(`Could not submit proof: ${error.message}`);}
  };

  window.openProofReviewFromTasksTab=async function(id){
    try{
      const workflow=await stagingApi();
      const assignment=(workflow.assignments||[]).find(row=>row.id===id);
      const proofs=(workflow.proofs||[]).filter(row=>row.assignment_id===id);
      const pending=proofs.find(row=>row.status==='pending');
      if(!assignment||!pending)throw new Error('No proof is waiting for review');
      reviewContext={assignment,proof:pending};
      document.getElementById('task-review-title').textContent='Review submitted proof';
      document.getElementById('task-review-sub').textContent=assignment.title||'';
      document.getElementById('task-review-body').innerHTML=`<div class="task-review-proof-card"><div class="task-review-proof-label">Submitted proof</div><div style="white-space:pre-wrap">${escapeHtml(pending.note||'No written note provided.')}</div><div style="margin-top:8px;font-size:11px;color:var(--muted)">${escapeHtml(pending.submitter_name||pending.submitter_email||'Assignee')} · ${escapeHtml(new Date(pending.submitted_at).toLocaleString())}</div></div>`;
      document.getElementById('task-review-status').textContent='';
      document.getElementById('mo-task-proof-review').classList.add('open');
    }catch(error){toast(error.message);}
  };

  window.approveTaskProof=async function(){
    if(!reviewContext)return;
    try{
      await stagingApi({action:'review_proof',proofId:reviewContext.proof.id,expectedVersion:Number(reviewContext.assignment.version||1),decision:'approved',reason:''});
      closeMo('mo-task-proof-review');reviewContext=null;toast('Proof approved');await refreshRealTasks();
    }catch(error){toast(`Could not approve proof: ${error.message}`);}
  };

  window.requestTaskProofChanges=async function(){
    if(!reviewContext)return;
    const reason=prompt('What needs to be changed?','Please update and resubmit the proof.');
    if(reason===null)return;
    try{
      await stagingApi({action:'review_proof',proofId:reviewContext.proof.id,expectedVersion:Number(reviewContext.assignment.version||1),decision:'changes_requested',reason});
      closeMo('mo-task-proof-review');reviewContext=null;toast('Changes requested');await refreshRealTasks();
    }catch(error){toast(`Could not request changes: ${error.message}`);}
  };

  const originalLoadUser=window.loadUser;
  window.loadUser=async function(account){
    await originalLoadUser(account);
    await connectRealtime();
    await refreshRealTasks();
  };
})();
