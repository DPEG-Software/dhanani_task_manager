if('serviceWorker' in navigator){navigator.serviceWorker.register('sw.js').catch(()=>{});}

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
