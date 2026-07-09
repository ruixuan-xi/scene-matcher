// ── Supabase 配置 ──
// 注意：这里使用 anon key，只读权限由 Supabase RLS 控制
// 在 Supabase SQL Editor 中执行以下 SQL 开启只读权限：
//   create policy "Public read scenes" on scenes for select using (true);
//   alter table scenes enable row level security;
const SUPABASE_URL = "https://hbiiwjqewgbeoalpwwrv.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhiaWl3anFld2diZW9hbHB3d3J2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI2MDI3NTAsImV4cCI6MjA5ODE3ODc1MH0.8QvTA58rqnQKl-2jNqaOL7r_t1UuCwrGx6rxcUKPrgk";

let SCENES = [];
let currentDept = "__all__";
let currentQuery = "";

const DEPT_COLORS = {
  "商务部":["#0ea5e9","#06b6d4","#e0f2fe"],"财务":["#8b5cf6","#a78bfa","#ede9fe"],
  "售后":["#f97316","#fb923c","#ffedd5"],"推广":["#ec4899","#f472b6","#fce7f3"],
  "运营":["#10b981","#34d399","#d1fae5"],"采购":["#f59e0b","#fbbf24","#fef3c7"],
  "人事":["#6366f1","#818cf8","#e0e7ff"],"市场":["#ef4444","#f87171","#fee2e2"],
};
function gc(dept,i){const c=DEPT_COLORS[dept]||["#0ea5e9","#06b6d4","#e0f2fe"];return c[i]}
function esc(s){return s?s.replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m])):""}

function showToast(msg){
  const t=document.getElementById("toast");
  t.textContent=msg;t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"),2200);
}

async function loadData(){
  document.getElementById("loading").style.display="block";
  document.getElementById("error").style.display="none";
  document.getElementById("grid").style.display="none";
  document.getElementById("empty").style.display="none";
  document.getElementById("export-btn").disabled=true;
  document.getElementById("search").disabled=true;
  try{
    const url=`${SUPABASE_URL}/rest/v1/scenes?select=id,scene_id,scene_title,process_steps,department,owner,platform,source,pain&order=department.asc,id.asc`;
    const resp=await fetch(url,{headers:{"apikey":SUPABASE_ANON_KEY,"Authorization":`Bearer ${SUPABASE_ANON_KEY}`}});
    if(!resp.ok) throw new Error(`HTTP ${resp.status}`);
    SCENES=await resp.json();
    if(!SCENES.length) throw new Error("场景库为空");
    initUI();
  }catch(e){
    if(window.__FALLBACK_DATA__ && window.__FALLBACK_DATA__.length){
      SCENES=window.__FALLBACK_DATA__;
      initUI();
      showToast("在线加载失败，使用缓存数据");
      return;
    }
    document.getElementById("loading").style.display="none";
    document.getElementById("error").style.display="block";
    document.getElementById("error-msg").textContent=`加载失败：${e.message}`;
  }
}

function initUI(){
  document.getElementById("loading").style.display="none";
  document.getElementById("grid").style.display="grid";
  document.getElementById("export-btn").disabled=false;
  document.getElementById("search").disabled=false;
  const deptMap={};
  SCENES.forEach(s=>{deptMap[s.department]=(deptMap[s.department]||0)+1});
  document.getElementById("total-count").textContent=SCENES.length;
  document.getElementById("dept-count").textContent=Object.keys(deptMap).length;
  const chipsEl=document.getElementById("chips");chipsEl.innerHTML="";
  function mkChip(label,count,val){
    const c=document.createElement("span");
    c.className="chip"+(val==="__all__"?" active":"");
    c.dataset.dept=val;
    c.innerHTML=label+(count!=null?` <span class="count">${count}</span>`:"");
    c.onclick=()=>selectDept(val);
    return c;
  }
  chipsEl.appendChild(mkChip("全部",SCENES.length,"__all__"));
  Object.keys(deptMap).sort((a,b)=>deptMap[b]-deptMap[a]).forEach(d=>chipsEl.appendChild(mkChip(d,deptMap[d],d)));
  render();
}

function selectDept(d){currentDept=d;document.querySelectorAll(".chip").forEach(c=>c.classList.toggle("active",c.dataset.dept===d));render()}
document.getElementById("search").addEventListener("input",e=>{currentQuery=e.target.value.trim().toLowerCase();render()});

function filterScenes(){
  return SCENES.filter(s=>{
    if(currentDept!=="__all__"&&s.department!==currentDept)return false;
    if(currentQuery){
      const hay=(s.scene_title+" "+s.process_steps+" "+(s.pain||"")+" "+(s.platform||"")+" "+(s.owner||"")+" "+(s.source||"")).toLowerCase();
      if(!hay.includes(currentQuery))return false;
    }
    return true;
  });
}

function toggleSteps(btn){
  const steps=btn.previousElementSibling;
  const expanded=steps.classList.toggle("expanded");
  steps.classList.toggle("collapsed",!expanded);
  btn.textContent=expanded?"收起 ▴":"展开全部 ▾";
}
window.toggleSteps=toggleSteps;

function render(){
  const items=filterScenes();
  document.getElementById("current-filter").innerHTML=`<strong>${items.length}</strong>个场景 · ${currentDept==="__all__"?"全部部门":currentDept}`;
  if(!items.length){document.getElementById("grid").innerHTML="";document.getElementById("empty").style.display="block";return}
  document.getElementById("empty").style.display="none";
  document.getElementById("grid").innerHTML=items.map(s=>{
    const color=gc(s.department,0),color2=gc(s.department,1),bg=gc(s.department,2);
    const needsToggle=(s.process_steps||"").length>80;
    const meta=[];
    if(s.owner)meta.push(`<span>👤 ${esc(s.owner)}</span>`);
    if(s.platform)meta.push(`<span>💻 ${esc(s.platform)}</span>`);
    if(s.source)meta.push(`<span>🏢 ${esc(s.source)}</span>`);
    return `<div class="card" style="--dept-color:${color};--dept-color2:${color2};--dept-bg:${bg}">
      <span class="dept-badge">${esc(s.department)}</span>
      <div class="card-title">${esc(s.scene_title)}</div>
      <div class="card-meta">${meta.join("")}</div>
      <div class="divider"></div>
      <div class="section-label">业务流程</div>
      <div class="steps ${needsToggle?'collapsed':''}">${esc(s.process_steps||"")}</div>
      ${needsToggle?'<button class="expand-btn" onclick="toggleSteps(this)">展开全部 ▾</button>':''}
      ${s.pain?`<div class="pain"><strong>痛点：</strong>${esc(s.pain)}</div>`:''}
    </div>`;
  }).join("");
}

// ── 导出 Excel ──
function exportExcel(){
  const items=filterScenes();
  if(!items.length){showToast("没有可导出的数据");return}
  const rows=[["ID","部门","场景名称","业务流程","提交人","涉及系统","企业名称","业务痛点"]];
  items.forEach(s=>{
    rows.push([s.id,s.department,s.scene_title,s.process_steps||"",s.owner||"",s.platform||"",s.source||"",s.pain||""]);
  });
  const ws=XLSX.utils.aoa_to_sheet(rows);
  // 列宽
  ws["!cols"]=[{wch:6},{wch:10},{wch:24},{wch:60},{wch:10},{wch:14},{wch:14},{wch:30}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"场景库");
  const deptLabel=currentDept==="__all__"?"全部":currentDept;
  const ts=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`场景库_${deptLabel}_${ts}.xlsx`);
  showToast(`已导出 ${items.length} 条场景`);
}
window.exportExcel=exportExcel;

// 内嵌 fallback 数据（构建时注入）
window.__FALLBACK_DATA__ = [];

loadData();
