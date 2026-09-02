// ============================================================
// ACTIVITY
// ============================================================
let curActTab='recent';
function setActTab(tab){
  curActTab=tab;
  document.getElementById('act-tab-recent')?.classList.toggle('active',tab==='recent');
  document.getElementById('act-tab-overdue')?.classList.toggle('active',tab==='overdue');
  renderActivity();
}
function renderActivity(){
  const el=document.getElementById("activity-list");
  let list;
  if(curActTab==='overdue'){
    list=tasks.filter(isOverdueTask).sort((a,b)=>new Date(a.date)-new Date(b.date)).slice(0,12);
  } else {
    list=[...tasks].sort((a,b)=>b.id-a.id).slice(0,10);
  }
  if(!list.length){el.innerHTML=`<div style="padding:20px;text-align:center;font-size:12px;color:var(--muted)">${curActTab==='overdue'?'No overdue tasks ✔':'No activity yet'}</div>`;return;}
  el.innerHTML=list.map(t=>{
    const dc=nstt(t.status)==="Done"?"done":isOverdueTask(t)?"overdue":"";
    const dateLbl=fmtD(t.date)?`Added: ${fmtD(t.date)}`:'No deadline';
    return `<div class="act-item" style="cursor:pointer" onclick="openDetail(${t.id})"><div class="act-dot ${dc}"></div><div><div class="act-title">${t.title}</div><div class="act-meta">${t.dept} &bull; ${t.person} &bull; ${dateLbl}</div></div></div>`;
  }).join("");
}
// ============================================================
// CHARTS
// ============================================================
function chartTheme(){
  const dark=document.documentElement.getAttribute('data-theme')==='dark';
  return {
    dark,
    text:dark?'#e5e7eb':'#374151',
    muted:dark?'#9ca3af':'#6b7280',
    grid:dark?'rgba(148,163,184,.18)':'#eef2f7',
    tooltipBg:dark?'#020617':'#111827',
    tooltipBody:dark?'#d1d5db':'#d1d5db',
    surface:dark?'#1f2937':'#fff'
  };
}
function mkChart(id,type,labels,datasets,opts={}){
  if(CH[id])CH[id].destroy();
  const ctx=document.getElementById(id);if(!ctx)return;
  const th=chartTheme();
  const interaction=opts.interaction||{mode:"index",intersect:false};
  CH[id]=new Chart(ctx,{type,data:{labels,datasets},options:{
    responsive:true,maintainAspectRatio:false,interaction,
    devicePixelRatio:(window.devicePixelRatio||1)*(window.__dpegZoomFactor||1),
    plugins:{
      legend:{display:!!opts.legend,position:"bottom",labels:{font:{family:"Inter",size:11,weight:"600"},padding:12,boxWidth:9,color:th.text,usePointStyle:true,pointStyle:"rectRounded"}},
      tooltip:{backgroundColor:th.tooltipBg,titleColor:"#f9fafb",bodyColor:th.tooltipBody,titleFont:{family:"Inter",size:12,weight:"700"},bodyFont:{family:"Inter",size:12},cornerRadius:8,padding:10,displayColors:!!opts.legend,filter:(item)=>!item.dataset.hidden,...(opts.tooltipCallbacks?{callbacks:opts.tooltipCallbacks}:(type==="doughnut"||type==="pie")?{callbacks:{title:items=>items.length?[items[0].label]:[],label:item=>`  ${item.parsed} tasks`}}:{})}
    },
    scales:type!=="pie"&&type!=="doughnut"?{
      x:{grid:{display:false},ticks:{font:{family:"Inter",size:11,weight:"600"},color:th.muted},border:{display:false}},
      y:{grid:{color:th.grid},ticks:{font:{family:"Inter",size:11,weight:"600"},color:th.muted},beginAtZero:true,border:{display:false}}
    }:{},animation:{duration:650,easing:"easeOutQuart"},events:["mousemove","mouseout","click","touchstart","touchmove"],onHover:(evt,elements)=>{ctx.style.cursor=elements.length?"pointer":"default";},onClick:opts.onClick||null,...opts.extra
  }});
}

function daysSince(date){
  const t=new Date(date||0).getTime();
  if(!Number.isFinite(t)||!t)return 0;
  return Math.max(0,Math.floor((Date.now()-t)/(24*60*60*1000)));
}

function executiveAttentionItems(){
  const open=tasks.filter(isOpenTask);
  const ranked=open.map(t=>{
    const overdue=isOverdueTask(t);
    const high=String(t.priority||"Normal").toLowerCase()==="high";
    const carry=!!t.carryForwardFrom;
    const replied=!!t.lastReplyAt&&daysSince(t.lastReplyAt)<=3;
    const age=daysSince(t.date);
    let score=0;
    if(overdue)score+=70+Math.min(age,21);
    if(high)score+=42;
    if(carry)score+=28+Math.min(daysSince(t.carryForwardFrom),21);
    if(replied)score+=18;
    if(age>=7)score+=8;
    if(score<18)return null;
    let reason='Watch';
    let tone='info';
    if(overdue){reason=age>7?`${Math.max(1,Math.floor(age/7))} wk late`:'Overdue';tone='hot';}
    else if(high){reason='High';tone='hot';}
    else if(carry){reason='Carryover';tone='warn';}
    else if(replied){reason='New reply';tone='info';}
    else if(age>=7){reason='Stale';tone='warn';}
    return {task:t,score,reason,tone,age};
  }).filter(Boolean).sort((a,b)=>b.score-a.score||new Date(b.task.date)-new Date(a.task.date));
  if(ranked.length>=5)return ranked.slice(0,5);
  const used=new Set(ranked.map(x=>x.task.id));
  const fillers=open
    .filter(t=>!used.has(t.id))
    .sort((a,b)=>new Date(b.lastReplyAt||b.date||0)-new Date(a.lastReplyAt||a.date||0))
    .slice(0,5-ranked.length)
    .map(t=>({task:t,score:0,reason:'Open',tone:'info',age:daysSince(t.date)}));
  return [...ranked,...fillers].slice(0,5);
}

function renderExecutiveAttention(){
  const el=document.getElementById('exec-attention-list');
  if(!el)return;
  const items=executiveAttentionItems();
  if(!items.length){
    el.innerHTML=`<div style="padding:28px 10px;text-align:center;color:var(--muted);font-size:12px">No executive attention items right now.</div>`;
    return;
  }
  el.innerHTML=items.map((x,i)=>{
    const t=x.task;
    const meta=[
      escapeHtml(t.dept||'No department'),
      escapeHtml(t.person||'Unassigned'),
      nstt(t.status),
      String(t.priority||'Normal').toLowerCase()==='high'?'High priority':'Normal'
    ].filter(Boolean);
    return `<div class="exec-item" onclick="openDetail(${t.id})">
      <div class="exec-top">
        <div class="exec-rank">${i+1}</div>
        <div class="exec-title">${escapeHtml(emailSubject(t))}</div>
        <div class="exec-reason ${x.tone}">${escapeHtml(x.reason)}</div>
      </div>
      <div class="exec-meta">${meta.join(' <span style="color:var(--border)">•</span> ')}</div>
    </div>`;
  }).join('');
}

function renderCharts(){
  const th=chartTheme();
  // Dept chart — grouped bars: overdue / pending / done
  const dp={};
  tasks.forEach(t=>{
    if(nstt(t.status)==="Cancelled")return;
    if(!dp[t.dept])dp[t.dept]={tot:0,dn:0,ov:0};
    dp[t.dept].tot++;
    if(nstt(t.status)==="Done")dp[t.dept].dn++;
    else if(isOverdueTask(t))dp[t.dept].ov++;
  });
  allDepartments().forEach(d=>{if(!dp[d])dp[d]={tot:0,dn:0,ov:0};});
  const dk=allDepartments().filter(d=>dp[d].tot>0);
  const deptHost=document.getElementById('ch-dept')?.parentElement;
  if(deptHost)deptHost.style.height=dk.length>1?'300px':'250px';
  const deptLabelAngle=dk.length>1?42:0;
  mkChart("ch-dept","bar",dk,[
    {label:"Overdue",data:dk.map(d=>dp[d].ov),backgroundColor:"#ef9a9a",hoverBackgroundColor:"#e57373",borderRadius:5,borderSkipped:false,maxBarThickness:18,barPercentage:.8,categoryPercentage:.72},
    {label:"Pending",data:dk.map(d=>dp[d].tot-dp[d].dn-dp[d].ov),backgroundColor:th.dark?"#64748b":"#d7dee8",hoverBackgroundColor:th.dark?"#94a3b8":"#c3ccd8",borderRadius:5,borderSkipped:false,maxBarThickness:18,barPercentage:.8,categoryPercentage:.72},
    {label:"Done",data:dk.map(d=>dp[d].dn),backgroundColor:"#8ecaa0",hoverBackgroundColor:"#6fb783",borderRadius:5,borderSkipped:false,maxBarThickness:18,barPercentage:.8,categoryPercentage:.72},
  ],{legend:true,interaction:{mode:"index",axis:"x",intersect:false},tooltipCallbacks:{title:items=>items.length?[dk[items[0].dataIndex]]:[],label:item=>`${item.dataset.label}: ${item.parsed.y}`},extra:{layout:{padding:{bottom:4}},scales:{x:{grid:{display:false},ticks:{font:{family:"Inter",size:10,weight:"600"},color:th.text,minRotation:deptLabelAngle,maxRotation:deptLabelAngle,autoSkip:false,padding:7},border:{display:false}},y:{grid:{color:th.grid},ticks:{font:{family:"Inter",size:11,weight:"600"},color:th.muted,precision:0},beginAtZero:true,border:{display:false}}}},onClick:(evt,elements)=>{if(elements.length){const d=dk[elements[0].index];nav("people","departments");selectDept(d);}}});

  // Open workload — normalise by staffKey so the same person doesn't split across entries
  const pp={};
  tasks.forEach(t=>{
    if(nstt(t.status)==="Cancelled")return;
    const key=staffKey(t.email,t.person);
    if(!pp[key])pp[key]={name:staffConfig[key]?.name||t.person,tot:0,dn:0,open:0,ov:0};
    pp[key].tot++;
    if(nstt(t.status)==="Done")pp[key].dn++;
    else{pp[key].open++;if(isOverdueTask(t))pp[key].ov++;}
  });
  const pk=Object.keys(pp).filter(k=>pp[k].open>0).sort((a,b)=>pp[b].open-pp[a].open||pp[b].tot-pp[a].tot);
  if(pk.length){
    const wrapName=n=>{
      const s=String(n||'Unknown').trim();
      const pts=s.split(/\s+/);
      if(s.length<=16)return s;
      if(pts.length===1)return s.slice(0,16);
      return [pts[0],pts.slice(1).join(' ')];
    };
    const pLabels=pk.map(k=>wrapName(pp[k].name||k));
    const host=document.getElementById('ch-person')?.parentElement;
    if(host)host.style.height=Math.max(190,pk.length*30+76)+'px';
    mkChart("ch-person","bar",pLabels,[
      {label:"Overdue",stack:"workload",data:pk.map(k=>pp[k].ov),backgroundColor:"#ef9a9a",hoverBackgroundColor:"#e57373",borderRadius:5,borderSkipped:false,maxBarThickness:14,order:2},
      {label:"Open",stack:"workload",data:pk.map(k=>pp[k].open-pp[k].ov),backgroundColor:"#f2c17c",hoverBackgroundColor:"#e9a85f",borderRadius:5,borderSkipped:false,maxBarThickness:14,order:2},
      {type:"line",label:"Total open",stack:"line",data:pk.map(k=>pp[k].open),borderColor:"#2E7D3F",backgroundColor:"#2E7D3F",pointBackgroundColor:th.surface,pointBorderColor:"#2E7D3F",pointRadius:3.5,pointHoverRadius:5,borderWidth:2,tension:.35,fill:false,order:1},
    ],{legend:true,interaction:{mode:'index',axis:'y',intersect:false},tooltipCallbacks:{title:items=>items.length?[pp[pk[items[0].dataIndex]].name||'Unknown']:[],label:item=>`${item.dataset.label}: ${item.parsed.x}`},onClick:(evt,elements)=>{if(elements.length){const person=pp[pk[elements[0].dataIndex]].name;nav('people');setTimeout(()=>{curPplFilter=person.toLowerCase();filterPpl(person);selectPpl(person);},150);}},extra:{indexAxis:"y",scales:{x:{stacked:true,grid:{color:th.grid},ticks:{font:{family:"Inter",size:11,weight:"600"},color:th.muted,precision:0},beginAtZero:true,border:{display:false}},y:{stacked:true,grid:{display:false},ticks:{font:{family:"Inter",size:10,weight:"600"},color:th.text,autoSkip:false},border:{display:false}}}}});
  }else{
    const host=document.getElementById('ch-person')?.parentElement;
    if(host)host.style.height='190px';
  }

  // Status doughnut
  const sc={"Pending":0,"Done":0,"Overdue":0};
  tasks.forEach(t=>{
    if(nstt(t.status)==="Cancelled")return;
    const bucket=nstt(t.status)==="Done"?"Done":(isOverdueTask(t)?"Overdue":"Pending");
    sc[bucket]++;
  });
  mkChart("ch-status","doughnut",Object.keys(sc),[{data:Object.values(sc),backgroundColor:["#f2c17c","#8ecaa0","#ef9a9a"],hoverBackgroundColor:["#e9a85f","#6fb783","#e57373"],borderWidth:3,borderColor:th.surface,hoverOffset:8}],{legend:true,extra:{cutout:"58%"},onClick:(evt,elements)=>{if(elements.length){const s=Object.keys(sc)[elements[0].index];document.getElementById("sf-status").value=s;nav("master");}}});

  renderExecutiveAttention();
}
function syncPulse(){
  const now=new Date();
  const wStart=new Date(now);wStart.setDate(now.getDate()-now.getDay());wStart.setHours(0,0,0,0);
  const wEnd=new Date(wStart);wEnd.setDate(wStart.getDate()+7);
  // "Added" uses task.id (a Date.now()-based creation timestamp), not
  // task.date (the due date) — otherwise this counted tasks due this week,
  // not tasks actually added this week.
  const added=tasks.filter(t=>{const d=new Date(t.id);return d>=wStart&&d<wEnd;}).length;
  // "Completed" prefers completedAt (when it was actually marked done);
  // falls back to the due date only for tasks completed before that field existed.
  const closed=tasks.filter(t=>{
    if(nstt(t.status)!=="Done")return false;
    const d=new Date(t.completedAt||t.date);
    return d>=wStart&&d<wEnd;
  }).length;
  const overdue=tasks.filter(isOverdueTask).length;
  const el1=document.getElementById("p-added");if(el1)el1.textContent=added;
  const el2=document.getElementById("p-closed");if(el2)el2.textContent=closed;
  const el3=document.getElementById("p-overdue");if(el3)el3.textContent=overdue;
}
