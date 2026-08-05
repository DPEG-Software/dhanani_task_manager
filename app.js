// ============================================================
// DARK MODE
// ============================================================
function applyDark(on){
  document.documentElement.setAttribute('data-theme',on?'dark':'light');
  const moon=document.getElementById('theme-icon-moon');
  const sun=document.getElementById('theme-icon-sun');
  if(moon)moon.style.display=on?'none':'';
  if(sun)sun.style.display=on?'':'none';
  if(document.getElementById('page-dashboard')?.classList.contains('active'))renderCharts();
}
function toggleDarkMode(){
  const isDark=document.documentElement.getAttribute('data-theme')==='dark';
  applyDark(!isDark);
  localStorage.setItem('dpeg_theme',isDark?'light':'dark');
}
function initDarkMode(){
  const saved=localStorage.getItem('dpeg_theme');
  if(saved==='dark'||(!saved&&window.matchMedia('(prefers-color-scheme:dark)').matches)){
    applyDark(true);
  }
}
// Start only after every feature module has loaded.
initDarkMode();
initSelects();
chWeek(0);
syncBadges();
loadAIConfig();

if (typeof msal !== 'undefined' && typeof msal.PublicClientApplication === 'function') {
  initMsal();
} else {
  showSignIn();
  setMsStatus('Authentication library not loaded. Please refresh the page.');
  console.error('MSAL check failed. msal:', typeof msal);
}
