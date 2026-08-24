// ============================================================
// MICROSOFT AUTH
// ============================================================
async function initMsal() {
  showSignIn();
  try {
    msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    await msalInstance.initialize();

    // Handle redirect response from Microsoft
    let response = null;
    try {
      response = await msalInstance.handleRedirectPromise();
    } catch (redirectErr) {
      console.error("Redirect error:", redirectErr);
      setMsStatus("Sign in error: " + redirectErr.message);
      return;
    }

    if (response && response.account) {
      restoreProofReturnSearch();
      currentAccount = response.account;
      await loadUser(currentAccount);
      return;
    }

    // Check existing session
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      restoreProofReturnSearch();
      currentAccount = accounts[0];
      await loadUser(currentAccount);
      return;
    }

    // Ready for sign in
    setMsStatus("");

  } catch (err) {
    console.error("MSAL init error:", err);
    setMsStatus("Error: " + err.message + " — Please refresh.");
  }
}

function showSignIn() {
  const el = document.getElementById("ms-signin-screen");
  if (el) el.style.display = "flex";
}

function hideSignIn() {
  const el = document.getElementById("ms-signin-screen");
  if (el) el.style.display = "none";
}

function setMsStatus(msg) {
  const el = document.getElementById("ms-status");
  if (el) el.textContent = msg;
}

function restoreProofReturnSearch(){
  const saved=sessionStorage.getItem(PROOF_RETURN_KEY)||'';
  if(saved&&saved.includes('proof=1')&&!location.search.includes('proof=1')){
    history.replaceState(null,'',location.pathname+saved);
  }
  if(location.search.includes('proof=1'))sessionStorage.removeItem(PROOF_RETURN_KEY);
}

async function signInWithMicrosoft() {
  const btn = document.getElementById("ms-login-btn");
  if (btn) { btn.disabled = true; btn.textContent = "Connecting..."; }
  setMsStatus("Redirecting to Microsoft...");
  try {
    if(isProofUploadRoute())sessionStorage.setItem(PROOF_RETURN_KEY,location.search);
    const loginRequest = {
      scopes: SCOPES,
      // Use the address the app is currently running on. Both production
      // GitHub Pages and http://localhost:8765/ are registered as SPA
      // redirect URIs in Microsoft Entra, so local testing returns locally
      // while production sign-in continues returning to production.
      redirectUri: window.location.origin + window.location.pathname,
      prompt: "select_account"
    };
    await msalInstance.loginRedirect(loginRequest);
  } catch (err) {
    console.error("Login error:", err);
    setMsStatus("Sign in failed. Please try again.");
    if (btn) { btn.disabled = false; btn.innerHTML = '<svg width="20" height="20" viewBox="0 0 21 21" fill="none"><rect x="1" y="1" width="9" height="9" fill="#F25022"/><rect x="11" y="1" width="9" height="9" fill="#7FBA00"/><rect x="1" y="11" width="9" height="9" fill="#00A4EF"/><rect x="11" y="11" width="9" height="9" fill="#FFB900"/></svg> Sign in with Microsoft'; }
  }
}

async function getSignedInProfile(account){
  const fallbackEmail=normEmail(account.username);
  try{
    const r=await msalInstance.acquireTokenSilent({scopes:["User.Read"],account});
    const res=await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle",{headers:{Authorization:`Bearer ${r.accessToken}`}});
    if(res.ok){
      const me=await res.json();
      const candidates=[me.mail,me.userPrincipalName,fallbackEmail].map(normEmail).filter(Boolean);
      const principalEmail=candidates.find(e=>PRINCIPALS[e])||candidates[0]||fallbackEmail;
      return {email:principalEmail,name:me.displayName||account.name||principalEmail,role:me.jobTitle||""};
    }
  }catch(err){console.warn("Profile lookup skipped:",err.message);}
  return {email:fallbackEmail,name:account.name||fallbackEmail,role:""};
}

async function loadUser(account) {
  const profile = await getSignedInProfile(account);
  const email = profile.email;
  const principal = PRINCIPALS[email];
  currentUser = {
    email,
    name: principal?.name || profile.name || account.username,
    role: principal?.role || profile.role || "Team Member",
    wednesday: principal?.wednesday || false,
    folder: email.split("@")[0],
  };
  isWednesdayUser = currentUser.wednesday;
  loadUserContacts();

  // Hide sign-in screen immediately
  hideSignIn();

  if(isProofUploadRoute()){
    await showProofUploadMode();
    syncD1ShadowInBackground();
    return;
  }

  updateUI();
  await loadTasksFromOneDrive();
}

// setStatus removed - no longer needed

async function signOut() {
  closeAccountMenu();
  if (msalInstance && currentAccount) {
    showSignIn();
    await msalInstance.logoutRedirect({
      account: currentAccount,
      postLogoutRedirectUri: window.location.origin + window.location.pathname
    });
  }
}

function updateUI() {
  const u = currentUser;
  document.getElementById("user-name").textContent = u.name;
  document.getElementById("user-role").textContent = u.role;
  document.getElementById("user-avatar").textContent = u.name.split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2);
  const adminBtn=document.getElementById("admin-settings-btn");
  if(adminBtn)adminBtn.style.display=isAdmin()?'flex':'none';
  const departmentBtn=document.getElementById("department-settings-btn");
  if(departmentBtn)departmentBtn.style.display=isAdmin()?'flex':'none';

  if(window.DPEG_STAGING_MODE)applyStagingUiBoundary();

  // Wednesday vs Discussion Notes
  const wedLabel = document.getElementById("nav-wednesday-label");
  const wedPageTitle = document.getElementById("wed-page-title");
  const agendaTitle = document.getElementById("agenda-title");
  const wedSaveBtn = document.getElementById("wed-save-btn");
  const dWedLbl = document.getElementById("d-wed-lbl");
  const actionNotesHint=document.getElementById('action-notes-hint');
  const modalNotesBtn=document.getElementById('mo-notes-btn');

  if (isWednesdayUser) {
    PT.wednesday = "Wednesday Review";
    if (wedLabel) wedLabel.textContent = "Wednesday Review";
    if (wedPageTitle) wedPageTitle.textContent = "Wednesday Review";
    if (agendaTitle) agendaTitle.textContent = "Agenda";
    if (wedSaveBtn) wedSaveBtn.textContent = "Save Agenda";
    if (dWedLbl) dWedLbl.textContent = "Wednesday Queue";
    if(actionNotesHint)actionNotesHint.textContent='Use + Wed to add tasks to your private Wednesday notes';
    if(modalNotesBtn)modalNotesBtn.textContent='Move to Wednesday';
  } else {
    PT.wednesday = "Discussion Notes";
    if (wedLabel) wedLabel.textContent = "Discussion Notes";
    if (wedPageTitle) wedPageTitle.textContent = "Discussion Notes";
    if (agendaTitle) agendaTitle.textContent = "Notes";
    if (wedSaveBtn) wedSaveBtn.textContent = "Save Notes";
    if (dWedLbl) dWedLbl.textContent = "Discussion Notes";
    if(actionNotesHint)actionNotesHint.textContent='Use + Discussion to add tasks to your private notes';
    if(modalNotesBtn)modalNotesBtn.textContent='Move to Discussion';
  }
}

function applyStagingUiBoundary(){
  document.body.classList.add('dpeg-staging-app');
  document.querySelectorAll('.sb-nav .ni').forEach(button=>{
    const handler=button.getAttribute('onclick')||'';
    button.style.display=handler.includes("nav('tasks')")?'':'none';
  });
  document.querySelectorAll('button[onclick="openAdd()"],#pwa-install-btn').forEach(button=>button.style.display='none');
  document.getElementById('admin-settings-btn')?.style.setProperty('display','none');
  document.getElementById('department-settings-btn')?.style.setProperty('display','none');
  if(!document.getElementById('staging-app-banner')){
    const banner=document.createElement('div');
    banner.id='staging-app-banner';
    banner.innerHTML='D1 STAGING — Fake tasks only. Email, OneDrive, Microsoft To Do and AI are disabled. <button type="button" id="staging-dual-write-test" onclick="testStagingDualWrite()">Test dual-write</button>';
    document.querySelector('.main')?.prepend(banner);
  }
}
