// Outlook mailbox folders are read directly from Microsoft Graph. Nothing in
// this module is stored in D1 or OneDrive; changes are changes to the signed-in
// user's real Outlook mailbox.
let outlookMailboxFolders=[];
let outlookFolderByKey=new Map();
let outlookEmailDragState=null;
let outlookMoveUndoTimer=null;

const OUTLOOK_STANDARD_FOLDER_NAMES=new Set([
  'archive','clutter','conversation history','deleted items','drafts','inbox',
  'junk email','outbox','sent items','sync issues','rss feeds','recoverable items deletions'
]);

function outlookCustomFolderKey(id){return `custom:${encodeURIComponent(String(id||''))}`;}
function resolveOutlookFolderId(key){
  if(FOLDER_MAP[key])return FOLDER_MAP[key];
  if(String(key||'').startsWith('custom:')){
    try{return decodeURIComponent(String(key).slice(7));}catch{return '';}
  }
  return '';
}
function outlookFolderLabel(key){
  return outlookFolderByKey.get(key)?.displayName||FOLDER_LABELS[key]||key;
}

async function outlookFolderWriteToken(){
  if(typeof getDraftAccessToken==='function')return getDraftAccessToken();
  const account=currentAccount||msalInstance.getActiveAccount();
  try{return (await msalInstance.acquireTokenSilent({scopes:SCOPES_DRAFTS,account})).accessToken;}
  catch{return (await msalInstance.acquireTokenPopup({scopes:SCOPES_DRAFTS})).accessToken;}
}

async function fetchOutlookFolderLevel(token,parentId='',depth=0){
  if(depth>5)return [];
  const base=parentId
    ?`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(parentId)}/childFolders`
    :'https://graph.microsoft.com/v1.0/me/mailFolders';
  let next=`${base}?$top=100&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount`;
  const folders=[];
  while(next){
    const res=await fetch(next,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error(`Could not load Outlook folders (${res.status})`);
    const data=await res.json();
    folders.push(...(data.value||[]));
    next=data['@odata.nextLink']||'';
  }
  const withChildren=await Promise.all(folders.map(async folder=>({
    ...folder,
    children:Number(folder.childFolderCount||0)>0
      ?await fetchOutlookFolderLevel(token,folder.id,depth+1)
      :[],
  })));
  return withChildren;
}

function flattenOutlookFolders(folders,depth=0,output=[]){
  (folders||[]).forEach(folder=>{
    output.push({...folder,depth});
    flattenOutlookFolders(folder.children,depth+1,output);
  });
  return output;
}

function customOutlookFolders(){
  return outlookMailboxFolders.filter(folder=>{
    const name=String(folder.displayName||'').trim().toLowerCase();
    return name&&!OUTLOOK_STANDARD_FOLDER_NAMES.has(name);
  });
}

function renderOutlookFolders(){
  const host=document.getElementById('ol-custom-folders');
  if(!host)return;
  const folders=customOutlookFolders();
  outlookFolderByKey=new Map(folders.map(folder=>[outlookCustomFolderKey(folder.id),folder]));
  if(!folders.length){host.innerHTML='<div class="ol-folder-empty">No personal folders</div>';return;}
  host.innerHTML=folders.map(folder=>{
    const key=outlookCustomFolderKey(folder.id);
    const indent=Math.min(Number(folder.depth||0),5)*12;
    const count=Number(folder.unreadItemCount||0);
    return `<div class="ol-folder ol-custom-folder" data-folder-key="${escapeHtml(key)}" onclick="loadFolder('${escapeHtml(key)}')" title="${escapeHtml(folder.displayName||'Folder')}" style="padding-left:${8+indent}px">
      <svg width="12" height="12" fill="none" stroke="currentColor" viewBox="0 0 24 24" style="flex-shrink:0"><path d="M3 7h7l2 2h9v10a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" stroke-width="2" stroke-linejoin="round"/></svg>
      <span style="overflow:hidden;text-overflow:ellipsis">${escapeHtml(folder.displayName||'Folder')}</span>
      ${count?`<span class="ol-folder-count">${count>999?'999+':count}</span>`:''}
      <button type="button" class="ol-folder-rename" onclick="event.stopPropagation();renameOutlookFolder('${escapeHtml(key)}')" title="Rename folder">✎</button>
    </div>`;
  }).join('');
}

function outlookFolderKeyFromElement(element){
  if(!element)return '';
  if(element.dataset.folderKey)return element.dataset.folderKey;
  const id=String(element.id||'');
  return id.startsWith('ol-folder-')?id.slice('ol-folder-'.length):'';
}

function validOutlookDropFolder(element){
  const key=outlookFolderKeyFromElement(element);
  return key&&!['flagged','untracked','schedule','calendar','sent','drafts'].includes(key)&&!!resolveOutlookFolderId(key);
}

function parseOutlookMessageIds(encoded,fallback=''){
  try{
    const parsed=JSON.parse(decodeURIComponent(String(encoded||'')));
    if(Array.isArray(parsed))return [...new Set(parsed.map(String).filter(Boolean))];
  }catch{}
  return fallback?[String(fallback)]:[];
}

function outlookConversationMessageIds(emailId){
  const card=document.getElementById(`email-item-${emailId}`);
  const fromCard=parseOutlookMessageIds(card?.dataset?.messageIds,emailId);
  const active=Array.isArray(window._activeOutlookConversationMessageIds)?window._activeOutlookConversationMessageIds:[];
  return card?fromCard:(active.includes(emailId)?active:[emailId]);
}

function startOutlookEmailDrag(event,emailId,sourceFolder,messageIds){
  if(!emailId||['flagged','untracked','search'].includes(String(sourceFolder||''))){event.preventDefault();return;}
  outlookEmailDragState={emailId,messageIds:parseOutlookMessageIds(messageIds,emailId),sourceFolder};
  event.dataTransfer.effectAllowed='move';
  event.dataTransfer.setData('text/plain',emailId);
  event.currentTarget?.classList.add('is-dragging');
  document.getElementById('page-outlook')?.classList.add('is-dragging-email');
}

function endOutlookEmailDrag(event){
  event.currentTarget?.classList.remove('is-dragging');
  document.getElementById('page-outlook')?.classList.remove('is-dragging-email');
  document.querySelectorAll('.ol-folder.drag-over').forEach(folder=>folder.classList.remove('drag-over'));
  outlookEmailDragState=null;
}

document.addEventListener('dragover',event=>{
  if(!outlookEmailDragState)return;
  const folder=event.target.closest?.('.ol-folder');
  document.querySelectorAll('.ol-folder.drag-over').forEach(item=>{if(item!==folder)item.classList.remove('drag-over');});
  if(!validOutlookDropFolder(folder))return;
  const destinationKey=outlookFolderKeyFromElement(folder);
  if(destinationKey===outlookEmailDragState.sourceFolder)return;
  event.preventDefault();
  event.dataTransfer.dropEffect='move';
  folder.classList.add('drag-over');
});

document.addEventListener('dragleave',event=>{
  const folder=event.target.closest?.('.ol-folder');
  if(folder&&!folder.contains(event.relatedTarget))folder.classList.remove('drag-over');
});

document.addEventListener('drop',event=>{
  if(!outlookEmailDragState)return;
  const folder=event.target.closest?.('.ol-folder');
  if(!validOutlookDropFolder(folder))return;
  const destinationKey=outlookFolderKeyFromElement(folder);
  if(destinationKey===outlookEmailDragState.sourceFolder)return;
  event.preventDefault();
  folder.classList.remove('drag-over');
  const move={...outlookEmailDragState};
  moveEmailIdsToOutlookFolder(move.messageIds,resolveOutlookFolderId(destinationKey),{
    sourceFolder:move.sourceFolder,
    destinationKey,
    allowUndo:true,
  });
});

async function loadOutlookFolders(options={}){
  const host=document.getElementById('ol-custom-folders');
  if(host&&!options.silent)host.innerHTML='<div class="ol-folder-loading">Loading folders…</div>';
  try{
    const token=await getAccessToken();
    outlookMailboxFolders=flattenOutlookFolders(await fetchOutlookFolderLevel(token));
    renderOutlookFolders();
    return outlookMailboxFolders;
  }catch(error){
    console.warn('Outlook folders unavailable:',error);
    if(host)host.innerHTML='<button class="ol-folder-empty" style="border:none;background:none;cursor:pointer" onclick="loadOutlookFolders()">Could not load · Retry</button>';
    return [];
  }
}

async function createOutlookFolder(){
  const name=String(prompt('Name the new Outlook folder:')||'').trim();
  if(!name)return;
  if(name.length>120){toast('Folder name is too long');return;}
  try{
    const token=await outlookFolderWriteToken();
    const res=await fetch('https://graph.microsoft.com/v1.0/me/mailFolders',{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({displayName:name}),
    });
    if(!res.ok){const detail=await res.text().catch(()=>'');throw new Error(detail||`Outlook returned ${res.status}`);}
    toast(`Folder “${name}” created`);
    await loadOutlookFolders({silent:true});
  }catch(error){console.error(error);toast('Could not create the Outlook folder');}
}

async function renameOutlookFolder(key){
  const folder=outlookFolderByKey.get(key);
  if(!folder)return;
  const name=String(prompt('Rename Outlook folder:',folder.displayName||'')||'').trim();
  if(!name||name===folder.displayName)return;
  try{
    const token=await outlookFolderWriteToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folder.id)}`,{
      method:'PATCH',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({displayName:name}),
    });
    if(!res.ok)throw new Error(`Outlook returned ${res.status}`);
    toast('Folder renamed');
    await loadOutlookFolders({silent:true});
    if(currentFolder===key){
      const title=document.getElementById('ol-folder-title');
      if(title)title.textContent=name;
    }
  }catch(error){console.error(error);toast('Could not rename the Outlook folder');}
}

function closeMoveEmailMenu(){document.getElementById('ol-folder-picker')?.remove();}

async function openMoveEmailMenu(emailId){
  if(!emailId)return;
  if(!outlookMailboxFolders.length)await loadOutlookFolders({silent:true});
  closeMoveEmailMenu();
  const destinations=[
    {id:'inbox',displayName:'Inbox',depth:0},
    {id:'archive',displayName:'Archive',depth:0},
    ...customOutlookFolders(),
  ];
  const backdrop=document.createElement('div');
  backdrop.id='ol-folder-picker';
  backdrop.className='ol-folder-picker-backdrop';
  backdrop.onclick=e=>{if(e.target===backdrop)closeMoveEmailMenu();};
  backdrop.innerHTML=`<div class="ol-folder-picker" role="dialog" aria-modal="true" aria-label="Move email">
    <div class="ol-folder-picker-head"><span>Move email to</span><button type="button" onclick="closeMoveEmailMenu()" aria-label="Close">×</button></div>
    <div class="ol-folder-picker-list">
      ${destinations.map(folder=>`<button type="button" class="ol-folder-picker-item" style="padding-left:${10+Math.min(Number(folder.depth||0),5)*12}px" data-destination-id="${escapeHtml(folder.id)}">
        <svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M3 7h7l2 2h9v10a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" stroke-width="2" stroke-linejoin="round"/></svg>
        <span>${escapeHtml(folder.displayName||'Folder')}</span>
      </button>`).join('')}
    </div>
  </div>`;
  backdrop.querySelectorAll('[data-destination-id]').forEach(button=>button.addEventListener('click',()=>moveEmailIdsToOutlookFolder(outlookConversationMessageIds(emailId),button.dataset.destinationId)));
  document.body.appendChild(backdrop);
}

function showOutlookMoveUndo(move){
  document.getElementById('ol-move-undo')?.remove();
  clearTimeout(outlookMoveUndoTimer);
  const bar=document.createElement('div');
  bar.id='ol-move-undo';
  bar.className='ol-move-undo';
  bar.innerHTML=`<span>Moved to ${escapeHtml(move.destinationLabel||'folder')}</span><button type="button">Undo</button><button type="button" class="ol-move-undo-close" aria-label="Dismiss">×</button>`;
  bar.querySelector('button')?.addEventListener('click',()=>undoOutlookEmailMove(move));
  bar.querySelector('.ol-move-undo-close')?.addEventListener('click',()=>bar.remove());
  document.body.appendChild(bar);
  outlookMoveUndoTimer=setTimeout(()=>bar.remove(),10000);
}

async function undoOutlookEmailMove(move){
  if(!move?.movedMessages?.length||!move?.sourceFolderId)return;
  const bar=document.getElementById('ol-move-undo');
  const undoButton=bar?.querySelector('button');
  if(undoButton)undoButton.disabled=true;
  try{
    const token=await outlookFolderWriteToken();
    for(const message of move.movedMessages){
      const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(message.newEmailId)}/move`,{
        method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({destinationId:move.sourceFolderId}),
      });
      if(!res.ok)throw new Error(`Outlook returned ${res.status}`);
    }
    bar?.remove();
    delete outlookFolderEmails[move.sourceFolder];
    if(currentFolder===move.sourceFolder)await loadFolder(move.sourceFolder);
    await loadOutlookFolders({silent:true});
    toast('Move undone');
  }catch(error){console.error(error);if(undoButton)undoButton.disabled=false;toast('Could not undo the move');}
}

async function moveEmailToOutlookFolder(emailId,destinationId,options={}){
  return moveEmailIdsToOutlookFolder([emailId],destinationId,options);
}

async function moveEmailIdsToOutlookFolder(messageIds,destinationId,options={}){
  const ids=[...new Set((Array.isArray(messageIds)?messageIds:[messageIds]).map(String).filter(Boolean))];
  if(!ids.length||!destinationId)return;
  const button=document.querySelector('#ol-folder-picker [data-destination-id]:focus');
  if(button)button.disabled=true;
  try{
    const token=await outlookFolderWriteToken();
    const source=options.sourceFolder||currentFolder;
    const sourceFolderId=resolveOutlookFolderId(source);
    const movedMessages=[];
    try{
      for(const id of ids){
        const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(id)}/move`,{
          method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
          body:JSON.stringify({destinationId}),
        });
        if(!res.ok)throw new Error(`Outlook returned ${res.status}`);
        const moved=await res.json().catch(()=>({}));
        movedMessages.push({oldEmailId:id,newEmailId:moved.id||''});
      }
    }catch(error){
      // Avoid leaving half a conversation in each folder if one Graph move
      // fails midway. Best-effort rollback of messages already moved.
      if(sourceFolderId){
        for(const moved of movedMessages.filter(item=>item.newEmailId)){
          await fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(moved.newEmailId)}/move`,{
            method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
            body:JSON.stringify({destinationId:sourceFolderId}),
          }).catch(()=>null);
        }
      }
      throw error;
    }
    const movedIds=new Set(ids);
    outlookFolderEmails[source]=(outlookFolderEmails[source]||[]).filter(email=>!movedIds.has(String(email.id)));
    const destinationEntry=[...outlookFolderByKey.entries()].find(([,folder])=>folder.id===destinationId);
    if(destinationEntry)delete outlookFolderEmails[destinationEntry[0]];
    if(destinationId==='inbox'||destinationId==='archive')delete outlookFolderEmails[destinationId];
    ids.forEach(id=>delete emailCache[id]);
    closeMoveEmailMenu();
    renderEmailRows(source,outlookFolderEmails[source]||[]);
    const reader=document.getElementById('ol-email-reader');
    if(reader)reader.innerHTML='<div style="flex:1;display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:12px">Email moved. Select another email.</div>';
    const customDestination=[...outlookFolderByKey.entries()].find(([,folder])=>folder.id===destinationId);
    const destinationKey=options.destinationKey||customDestination?.[0]||destinationId;
    const destinationLabel=outlookFolderLabel(destinationKey);
    if(options.allowUndo!==false&&movedMessages.every(message=>message.newEmailId)&&sourceFolderId){
      showOutlookMoveUndo({movedMessages,sourceFolder:source,sourceFolderId,destinationLabel});
    }else toast(`Email moved to ${destinationLabel||'folder'}`);
    await loadOutlookFolders({silent:true});
  }catch(error){
    console.error(error);
    if(button)button.disabled=false;
    toast('Could not move the email');
  }
}
