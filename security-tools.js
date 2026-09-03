// Proof-link privacy audit. Production cleanup remains explicitly user-driven;
// staging uses synthetic permissions and never calls Microsoft Graph.
function isDirectAnonymousPermission(permission){
  return permission?.link?.scope==='anonymous'&&!permission?.inheritedFrom&&Boolean(permission?.id);
}

function summarizeProofPermissions(rows){
  const permissions=Array.isArray(rows)?rows:[];
  return {
    directAnonymous:permissions.filter(isDirectAnonymousPermission).length,
    inheritedAnonymous:permissions.filter(p=>p?.link?.scope==='anonymous'&&Boolean(p?.inheritedFrom)).length,
    organization:permissions.filter(p=>p?.link?.scope==='organization').length,
    other:permissions.filter(p=>!['anonymous','organization'].includes(p?.link?.scope||'')).length,
  };
}

function showProofSecurityResult(result,{staging=false}={}){
  const message=[
    staging?'STAGING SIMULATION — no files or permissions were changed.':'Proof-link security scan complete.',
    `Anonymous links eligible for removal: ${result.directAnonymous}`,
    `Inherited anonymous permissions requiring owner review: ${result.inheritedAnonymous}`,
    `Organization-only links preserved: ${result.organization}`,
    `Other/private permissions preserved: ${result.other}`,
  ].join('\n');
  alert(message);
}

async function ownedProofRecords(){
  const token=await getAccessToken();
  const base=workerBaseUrl();
  const response=await fetch(`${base}/notify`,{headers:{Authorization:`Bearer ${token}`}});
  if(!response.ok)throw new Error(`Could not load proof inventory (${response.status})`);
  const data=await response.json();
  const owner=normEmail(currentUser?.email||'');
  const items=new Map();
  for(const notification of data.notifications||[]){
    for(const proof of notification.proofs||[]){
      const itemId=String(proof?.driveItemId||'').trim();
      if(!itemId||normEmail(proof?.uploadedBy||'')!==owner)continue;
      if(!items.has(itemId))items.set(itemId,{itemId,name:String(proof.name||'Proof file')});
    }
  }
  return {token,items:[...items.values()]};
}

async function scanOwnedProofPermissions(){
  const {token,items}=await ownedProofRecords();
  const result={directAnonymous:0,inheritedAnonymous:0,organization:0,other:0,filesChecked:0,failures:0,unsafe:[],token};
  for(const item of items){
    try{
      const response=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(item.itemId)}/permissions?$select=id,link,inheritedFrom`,{headers:{Authorization:`Bearer ${token}`}});
      if(!response.ok){result.failures++;continue;}
      const permissions=(await response.json()).value||[];
      const summary=summarizeAnonymousPermissions(permissions);
      result.directAnonymous+=summary.directAnonymous;
      result.inheritedAnonymous+=summary.inheritedAnonymous;
      result.organization+=summary.organization;
      result.other+=summary.other;
      result.filesChecked++;
      const removable=permissions.filter(isDirectAnonymousPermission);
      if(removable.length)result.unsafe.push({...item,permissions:removable});
    }catch{result.failures++;}
  }
  return result;
}

function summarizeAnonymousPermissions(permissions){return summarizeProofPermissions(permissions);}

async function replaceStoredProofLink(token,itemId,webUrl){
  const response=await fetch(`${workerBaseUrl()}/notify`,{
    method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({type:'proof_security_link_update',driveItemId:itemId,webUrl,shareId:graphShareIdFromUrl(webUrl)}),
  });
  if(!response.ok)throw new Error(`Could not save replacement secure link (${response.status})`);
}

async function removeAnonymousProofPermissions(scan){
  let removed=0,failed=0;
  for(const item of scan.unsafe){
    try{
      const secureUrl=await createProofViewLink(scan.token,item.itemId);
      await replaceStoredProofLink(scan.token,item.itemId,secureUrl);
      for(const permission of item.permissions){
        const response=await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(item.itemId)}/permissions/${encodeURIComponent(permission.id)}`,{method:'DELETE',headers:{Authorization:`Bearer ${scan.token}`}});
        if(response.status===204||response.status===404)removed++;else failed++;
      }
    }catch{failed+=item.permissions.length;}
  }
  return {removed,failed};
}

async function openProofSecurityCheck(){
  closeAccountMenu();
  if(window.DPEG_STAGING_MODE){
    const synthetic=[
      {id:'test-anonymous',link:{scope:'anonymous'}},
      {id:'test-organization',link:{scope:'organization'}},
      {id:'test-inherited',link:{scope:'anonymous'},inheritedFrom:{id:'parent'}},
      {id:'test-private'},
    ];
    showProofSecurityResult(summarizeProofPermissions(synthetic),{staging:true});
    return;
  }
  try{
    toast('Scanning proof-link permissions…');
    const scan=await scanOwnedProofPermissions();
    if(!scan.filesChecked&&!scan.failures){alert('No historical proof files owned by your account were found. Nothing was changed.');return;}
    const summary=`Proof files checked: ${scan.filesChecked}\nAnonymous permissions found: ${scan.directAnonymous}\nOrganization-only permissions preserved: ${scan.organization}\nInherited permissions requiring manual review: ${scan.inheritedAnonymous}${scan.failures?`\nFiles that could not be checked: ${scan.failures}`:''}`;
    if(!scan.directAnonymous){alert(`${summary}\n\nNo unsafe direct links were found. Nothing was changed.`);return;}
    if(!confirm(`${summary}\n\nRemove only the direct anonymous permissions now? Files will not be deleted and secure organization-only links will be created first.`))return;
    const outcome=await removeAnonymousProofPermissions(scan);
    alert(`Proof-link cleanup finished.\nAnonymous permissions removed: ${outcome.removed}\nFailed removals: ${outcome.failed}\nNo files were deleted.`);
  }catch(error){alert(`Proof-link security check could not finish: ${error.message}`);}
}
