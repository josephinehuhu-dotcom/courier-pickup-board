const storageKey='courier-pickup-orders-v1';
let orders;
try{orders=JSON.parse(localStorage.getItem(storageKey)||'[]');}catch{orders=[];}
if(!Array.isArray(orders))orders=[];
let tab='pending', station='全部驿站', previewOrder=null;
const app=document.querySelector('#app'), fileInput=document.querySelector('#file-input');
const persist=()=>localStorage.setItem(storageKey,JSON.stringify(orders));
const escape=s=>String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const last4=s=>String(s||'').replace(/\D/g,'').slice(-4)||'未填';
const initials=s=>String(s||'?').slice(0,1);
const orderStatus=o=>o.status||'pending';
function normalizeHeader(s){return String(s||'').trim().replace(/[（(].*?[）)]/g,'').replace(/\s/g,'');}
function screenshotUrl(value){
  const raw=String(value||'').trim();
  if(!raw)return '';
  try{const url=new URL(raw);return /^https?:$/.test(url.protocol)?url.href:'';}catch{return '';}
}
function parseCSV(text){const rows=[];let row=[],cell='',quote=false;for(let i=0;i<text.length;i++){const ch=text[i],next=text[i+1];if(ch==='"'&&quote&&next==='"'){cell+='"';i++;}else if(ch==='"'){quote=!quote;}else if(ch===','&&!quote){row.push(cell.trim());cell='';}else if((ch==='\n'||ch==='\r')&&!quote){if(ch==='\r'&&next==='\n')i++;row.push(cell.trim());if(row.some(Boolean))rows.push(row);row=[];cell='';}else cell+=ch;}row.push(cell.trim());if(row.some(Boolean))rows.push(row);return rows;}
function rowsToOrders(rows){
  const headerIndex=rows.findIndex(r=>r.some(v=>normalizeHeader(v).includes('收件人姓名'))&&r.some(v=>normalizeHeader(v).includes('取件码')));
  if(headerIndex<0)return [];
  const heads=rows[headerIndex].map(normalizeHeader), data=rows.slice(headerIndex+1);
  const get=(r,names)=>{const i=heads.findIndex(h=>names.some(n=>h===n||h.includes(n)));return i<0?'':r[i];};
  const screenshotIndex=heads.findIndex(h=>h==='订单截图');
  return data.map((r,i)=>({id:`import-${Date.now()}-${i}`,name:String(get(r,['收件人姓名','姓名'])||'').trim(),phone:String(get(r,['收件人电话','电话'])||'').trim(),station:String(get(r,['驿站选择','驿站'])||'').trim(),code:String(get(r,['取件码'])||'').trim(),imageUrl:screenshotUrl(screenshotIndex<0?'':r[screenshotIndex]),submittedAt:String(get(r,['填写时间','提交时间'])||'').trim(),status:'pending'})).filter(o=>o.name||o.code);
}
async function importSheet(file){
  try{
    const buffer=await file.arrayBuffer(), bytes=new Uint8Array(buffer), isZip=bytes[0]===0x50&&bytes[1]===0x4b;
    let rows;
    if(isZip||/\.(xlsx|xls)$/i.test(file.name)){
      if(!window.XLSX)throw new Error('Excel 解析组件未加载，请联网后刷新页面重试');
      const workbook=XLSX.read(buffer,{type:'array',cellDates:false});
      const sheet=workbook.Sheets[workbook.SheetNames[0]];
      rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:''});
    }else{
      let text;
      try{text=new TextDecoder('utf-8',{fatal:true}).decode(buffer);}catch{text=new TextDecoder('gb18030').decode(buffer);}
      rows=parseCSV(text.replace(/^\uFEFF/,''));
    }
    const parsed=rowsToOrders(rows);
    if(!parsed.length)throw new Error('没有找到订单，请确认表格包含“收件人姓名、收件人电话、驿站选择、取件码”');
    orders=parsed;persist();station='全部驿站';tab='pending';showToast(`已导入 ${orders.length} 单，并按收件人自动归组`);render();
  }catch(error){showToast(error.message||'文件解析失败，请重新导出后再试');}
}
function groups(){const list=orders.filter(o=>(tab==='pending'?orderStatus(o)!=='done':orderStatus(o)==='done')&&(station==='全部驿站'||o.station===station));const map=new Map();list.forEach(o=>{const k=`${o.name}|${last4(o.phone)}`;if(!map.has(k))map.set(k,{key:k,name:o.name,phone:o.phone,items:[]});map.get(k).items.push(o);});return [...map.values()].sort((a,b)=>a.name.localeCompare(b.name,'zh-CN'));}
function render(){const pending=orders.filter(o=>orderStatus(o)!=='done').length,done=orders.length-pending;const sts=['全部驿站',...new Set(orders.map(o=>o.station).filter(Boolean))];const content=groups();app.innerHTML=`
  <header class="top"><div><h1>快递代取工作台</h1><div class="subtitle">今日截止 14:00 · 目标 19:00 前取回超市</div></div><button class="import" data-action="import">导入 Excel / CSV</button></header>
  <div class="notice">导入腾讯文档 Excel / CSV 后，会原样读取「驿站选择」并自动生成分类；数据只保存在当前设备，不会回写原表。</div>
  <section class="dashboard"><div class="metric"><b>${pending}</b><span>待取件</span></div><div class="metric"><b>${done}</b><span>已取件</span></div><div class="metric"><b>${new Set(orders.map(o=>o.name+'|'+last4(o.phone))).size}</b><span>收件人</span></div></section>
  <nav class="tabs"><button class="tab ${tab==='pending'?'active':''}" data-tab="pending">待取件 ${pending}</button><button class="tab ${tab==='done'?'active':''}" data-tab="done">已取件 ${done}</button></nav>
  <div class="filters">${sts.map(s=>`<button class="chip ${station===s?'active':''}" data-station="${escape(s)}">${escape(s)}</button>`).join('')}</div>
  <section>${content.length?content.map(groupHTML).join(''):`<div class="empty">📦<strong>${tab==='pending'?(orders.length?'所有快递都已取完':'暂无待取订单'):'还没有已取件记录'}</strong>${tab==='pending'?(orders.length?'辛苦啦，可以切换到已取件查看记录。':'请先导入 Excel / CSV 表格。'):'导入表格后，完成取件的订单会在这里出现。'}</div>`}</section>
  <footer class="bottom"><div class="bottom-in"><small>按“姓名 + 手机尾号”自动归到同一收纳框 · 已取件状态锁定</small></div></footer>
  ${previewOrder?previewHTML(previewOrder):''}`;
}
function groupHTML(g){const total=g.items.length, done=g.items.filter(o=>orderStatus(o)==='done').length;return `<article class="group"><div class="group-head"><div class="person"><span class="avatar">${escape(initials(g.name))}</span><div><b>${escape(g.name||'未填姓名')}</b><small>手机尾号 ${escape(last4(g.phone))} · ${done}/${total} 件已取</small></div></div><span class="count">同框 ${total} 件</span></div>${g.items.map(parcelHTML).join('')}</article>`;}
function parcelHTML(o){const done=orderStatus(o)==='done',imageUrl=screenshotUrl(o.imageUrl);const imageAction=imageUrl?`<button class="secondary" data-action="preview" data-id="${o.id}">打开订单截图</button>`:`<span class="secondary disabled">无订单截图</span>`;return `<div class="parcel ${done?'done':''}"><div class="parcel-top"><span class="station">${escape(o.station||'未填驿站')}</span></div><div class="code">${escape(o.code||'未填写取件码')}</div><div class="actions">${imageAction}<button class="secondary" data-action="copy" data-id="${o.id}">复制码</button>${done?'<span class="complete locked">已取到 ✓</span>':`<button class="complete" data-action="complete" data-id="${o.id}">标记已取到</button>`}</div></div>`;}
function previewHTML(o){return `<div class="image-modal" role="dialog" aria-modal="true" aria-label="订单截图预览"><div class="image-modal-bar"><strong>${escape(o.name||'订单截图')} · ${escape(o.station||'未填驿站')}</strong><button class="image-close" data-action="close-preview" aria-label="关闭图片预览">×</button></div><div class="image-stage"><img class="order-image" src="${escape(screenshotUrl(o.imageUrl))}" alt="${escape(o.name||'收件人')}的订单截图"><p class="image-error" hidden>图片加载失败，请检查截图链接或访问权限。</p></div></div>`;}
function showToast(message){document.querySelector('.toast')?.remove();const el=document.createElement('div');el.className='toast';el.textContent=message;document.body.append(el);setTimeout(()=>el.remove(),2200);}
app.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;if(b.dataset.tab){tab=b.dataset.tab;render();return;}if(b.dataset.station){station=b.dataset.station;render();return;}const action=b.dataset.action,id=b.dataset.id,o=orders.find(x=>x.id===id);if(action==='import')fileInput.click();if(action==='preview'&&o&&screenshotUrl(o.imageUrl)){previewOrder=o;render();}if(action==='close-preview'){previewOrder=null;render();}if(action==='copy'&&o){navigator.clipboard?.writeText(o.code);showToast('取件码已复制');}if(action==='complete'&&o&&orderStatus(o)!=='done'){o.status='done';persist();showToast('已标记为取到，状态已锁定');render();}});
app.addEventListener('error',e=>{if(!e.target.matches?.('.order-image'))return;e.target.hidden=true;const message=e.target.nextElementSibling;if(message)message.hidden=false;},true);
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&previewOrder){previewOrder=null;render();}});
fileInput.addEventListener('change',e=>{if(e.target.files[0])importSheet(e.target.files[0]);e.target.value='';});
render();
