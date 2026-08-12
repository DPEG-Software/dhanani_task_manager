// The full staging UI is hosted on a separate Cloudflare Pages project. Keep
// its data and side effects isolated even though it reuses the production UI
// modules. Preview deployment hostnames include the project name as a suffix.
window.DPEG_STAGING_MODE=location.hostname==='dpeg-task-manager-staging-test.pages.dev'
  ||location.hostname.endsWith('.dpeg-task-manager-staging-test.pages.dev');
window.DPEG_STAGING_WORKER='https://dpeg-task-manager-staging.systemmanager1.workers.dev';
if(window.DPEG_STAGING_MODE){
  // This origin is staging-only, so force every legacy helper that still reads
  // the saved Worker setting to the isolated staging Worker.
  localStorage.setItem('dpeg_ai_fn_url',window.DPEG_STAGING_WORKER);
}

if('serviceWorker' in navigator){
  if(window.DPEG_STAGING_MODE){
    navigator.serviceWorker.getRegistrations().then(rows=>rows.forEach(row=>row.unregister())).catch(()=>{});
  }else{
    navigator.serviceWorker.register('sw.js').catch(()=>{});
  }
}

// Scale the whole app to the actual window size. This runs in the document
// head to avoid a visible resize after the main feature modules load.
(function(){
  var BASELINE_WIDTH=1440;
  var MIN_ZOOM=0.9,MAX_ZOOM=1.1;
  var MOBILE_BREAKPOINT=760;
  var timer=null;
  function apply(){
    var w=window.outerWidth||window.innerWidth;
    if(w<=MOBILE_BREAKPOINT){document.documentElement.style.zoom='';window.__dpegZoomFactor=1;return;}
    var factor=Math.min(MAX_ZOOM,Math.max(MIN_ZOOM,w/BASELINE_WIDTH));
    document.documentElement.style.zoom=factor;
    window.__dpegZoomFactor=factor;
  }
  window.addEventListener('resize',function(){clearTimeout(timer);timer=setTimeout(apply,150);});
  apply();
})();
