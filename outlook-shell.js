let olLeftExpanded=false;
let olMidExpanded=false;
let outlookContacts=[];
let olSearchTimer=null;
let currentComposeType='';
let currentComposeEmailId=null;
let currentComposeFiles=[];

function initOlPanels(){
  olLeftExpanded=false;
  olMidExpanded=false;
  const lp=document.getElementById('ol-left-panel');
  const mp=document.getElementById('ol-mid-panel');
  if(lp)lp.style.width='0';
  if(mp)mp.style.width='0';
  updateOlDividerIcons();
  setupEmailListInfiniteScroll();
}

function setupEmailListInfiniteScroll(){
  const listEl=document.getElementById('ol-email-list');
  if(!listEl||listEl._infiniteScrollBound)return;
  listEl._infiniteScrollBound=true;
  listEl.addEventListener('scroll',()=>{
    if(listEl.scrollTop+listEl.clientHeight<listEl.scrollHeight-200)return;
    // Search results page independently of currentFolder — currentFolder still
    // reflects whichever real folder the user was in before searching, and a
    // lot of other logic (mark-as-read, delete, back-navigation) depends on
    // that staying accurate, so search must not overwrite it.
    if(_olSearchActive){
      if(outlookNextLinks.search&&!_loadingMoreEmails.search)loadMoreEmails('search');
      return;
    }
    const folder=currentFolder;
    if(!folder||!outlookNextLinks[folder]||_loadingMoreEmails[folder])return;
    if(folder==='untracked'){loadMoreUntracked();return;}
    loadMoreEmails(folder);
  });
}

function toggleOlLeft(){
  olLeftExpanded=!olLeftExpanded;
  const lp=document.getElementById('ol-left-panel');
  if(lp)lp.style.width=olLeftExpanded?'190px':'0';
  updateOlDividerIcons();
}

function toggleOlMid(){
  olMidExpanded=!olMidExpanded;
  const mp=document.getElementById('ol-mid-panel');
  if(mp)mp.style.width=olMidExpanded?'300px':'0';
  updateOlDividerIcons();
}

function updateOlDividerIcons(){
  const li=document.getElementById('ol-divider-left-icon');
  const mi=document.getElementById('ol-divider-mid-icon');
  if(li)li.innerHTML=olLeftExpanded?'<path d="M6 2l-4 5 4 5" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>'
    :'<path d="M2 2l4 5-4 5" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>';
  if(mi)mi.innerHTML=olMidExpanded?'<path d="M6 2l-4 5 4 5" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>'
    :'<path d="M2 2l4 5-4 5" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>';
}

function toggleOlSidebar(){toggleOlLeft();}
function olMobileBackToList(){
  document.getElementById('page-outlook')?.classList.remove('ol-mobile-reading');
}

async function fetchOutlookContacts(){
  try{
    const token=await getAccessToken();
    const res=await fetch('https://graph.microsoft.com/v1.0/me/people?$top=100&$select=displayName,scoredEmailAddresses,jobTitle,department',{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)return;
    const data=await res.json();
    const people=data.value||[];
    outlookContacts=people.filter(p=>p.scoredEmailAddresses?.length).map(p=>({
      name:p.displayName||'',
      email:(p.scoredEmailAddresses[0]?.address||'').toLowerCase(),
      role:p.jobTitle||'',
      dept:p.department||'Unknown'
    })).filter(p=>p.email);
    // Merge into staffConfig if not present
    outlookContacts.forEach(c=>{
      const existing=Object.values(staffConfig||{}).find(s=>normEmail(s.email)===normEmail(c.email));
      if(!existing&&c.email){staffConfig[c.email]=c;}
    });
    const bar=document.getElementById('ol-contacts-sync-bar');
    const txt=document.getElementById('ol-contacts-sync-text');
    if(bar&&txt){
      txt.textContent=`${outlookContacts.length} contacts synced`;
      bar.style.display='flex';
      setTimeout(()=>{if(bar)bar.style.display='none';},4000);
    }
  }catch(err){console.log('Contacts sync skipped:',err);}
}

let _olSearchActive=false;
function handleOlSearch(val){
  const clearBtn=document.getElementById('ol-search-clear');
  if(clearBtn)clearBtn.style.display=val.length>0?'block':'none';
  clearTimeout(olSearchTimer);
  if(val.length<3){
    if(_olSearchActive){_olSearchActive=false;renderEmailRows(currentFolder,outlookFolderEmails[currentFolder]||[]);}
    return;
  }
  olSearchTimer=setTimeout(()=>searchEmails(val),400);
}

function clearOlSearch(){
  const inp=document.getElementById('ol-search-input');
  const clearBtn=document.getElementById('ol-search-clear');
  if(inp)inp.value='';
  if(clearBtn)clearBtn.style.display='none';
  _olSearchActive=false;
  renderEmailRows(currentFolder,outlookFolderEmails[currentFolder]||[]);
  const titleEl=document.getElementById('ol-folder-title');
  if(titleEl)titleEl.textContent=FOLDER_LABELS[currentFolder]||currentFolder;
}

async function searchEmails(query){
  _olSearchActive=true;
  const listEl=document.getElementById('ol-email-list');
  if(!listEl)return;
  // Auto-expand mid panel for search results
  if(!olMidExpanded){olMidExpanded=true;const mp=document.getElementById('ol-mid-panel');if(mp)mp.style.width='300px';updateOlDividerIcons();}
  const titleEl=document.getElementById('ol-folder-title');
  if(titleEl)titleEl.textContent='Searching...';
  listEl.innerHTML='<div class="empty-state"><div class="es-text">Searching...</div></div>';
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages?$search="${encodeURIComponent(query)}"&$top=100&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,conversationId,importance,hasAttachments,flag`,{headers:{Authorization:`Bearer ${token}`}});
    if(!res.ok)throw new Error('Search failed');
    const data=await res.json();
    const results=data.value||[];
    outlookFolderEmails.search=results;
    outlookNextLinks.search=data['@odata.nextLink']||'';
    if(titleEl)titleEl.textContent=`Results (${results.length})`;
    if(!results.length){listEl.innerHTML='<div class="empty-state"><div class="es-text">No results</div><div class="es-sub">Try different keywords</div></div>';return;}
    // Rendered under its own 'search' key (not 'inbox') so pagination/load-more
    // tracks the search results instead of corrupting the real inbox's cursor.
    renderEmailRows('search',results);
  }catch(err){
    if(titleEl)titleEl.textContent='Search';
    if(listEl)listEl.innerHTML='<div class="empty-state"><div class="es-text">Search failed</div></div>';
  }
}

async function deleteEmailItem(emailId){return deleteEmail(emailId);}

async function deleteEmail(emailId){
  if(!emailId)return;
  try{
    const token=await getAccessToken();
    const res=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/move`,{
      method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
      body:JSON.stringify({destinationId:'deleteditems'})
    });
    if(!res.ok)throw new Error('Delete failed');
    // Remove from current folder list immediately
    const emailEl=document.getElementById('email-item-'+emailId);
    if(emailEl)emailEl.remove();
    if(outlookFolderEmails[currentFolder]){
      outlookFolderEmails[currentFolder]=outlookFolderEmails[currentFolder].filter(e=>e.id!==emailId);
    }
    delete emailCache[emailId];
    const readerEl=document.getElementById('ol-email-reader');
    if(readerEl)readerEl.innerHTML=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#9ca3af;text-align:center;padding:40px"><div style="font-size:13px;font-weight:600;color:#6b7280;margin-bottom:3px">Email moved to Deleted Items</div><div style="font-size:11px;color:#9ca3af">Select another email to read</div></div>`;
    toast('Email moved to Deleted Items');
    // Refresh deleted folder cache if it was open
    if(currentFolder==='deleted')loadFolder('deleted');
  }catch(err){toast('Could not delete email: '+err.message);}
}
