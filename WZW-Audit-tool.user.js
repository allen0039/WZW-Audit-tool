// ==UserScript==
// @name         WZW-Audit-tool
// @namespace    http://tampermonkey.net/
// @version      8.1
// @description  审计实训万能填表助手 v8.1 - 在线题库下载，支持 脚本猫/油猴
// @author       Allen
// @match        http://10.18.0.178:9350/*
// @grant        none
// @supportURL   https://github.com/allen0039/WZW-Audit-tool
// @homepageURL  https://github.com/allen0039/WZW-Audit-tool
// @license      MIT
// ==/UserScript==

// ============================================================
// ★ 免责声明：
//   1. 本项目（WZW-Audit-tool）仅用于学习交流及辅助教学，请勿依赖。
//   2. 题库数据无法保证 100% 准确无误，填写结果仅供参考，必须经人工审核！
//   3. 本项目完全开源免费，【严禁用于任何商业用途】！
// ============================================================

(function(){'use strict';

// ============================================================
// ★ 配置区：可用的下载源
// ============================================================
const TIKU_SOURCES = {
  'jsDelivr CDN（国内推荐）': 'https://cdn.jsdelivr.net/gh/allen0039/WZW-Audit-tool@main/tiku.json',
  'GitHub Raw（原始源）': 'https://raw.githubusercontent.com/allen0039/WZW-Audit-tool/main/tiku.json'
};
const SOURCE_KEY = 'shenjitools_source';
const STORAGE_KEY = 'shenjitools_tiku';

function getTiku() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

async function downloadTiku() {
  const btn = document.getElementById('dl-tiku-btn');
  const sel = document.getElementById('tiku-source-select');
  if(!btn) return;
  const label = sel ? sel.value : Object.keys(TIKU_SOURCES)[0];
  const url = TIKU_SOURCES[label];
  if (!url) { showToast('❌ 未找到所选下载源', 'error'); return; }
  if(sel) localStorage.setItem(SOURCE_KEY, label);
  btn.innerHTML = '⏳ 下载中...';
  btn.disabled = true;
  try {
    const resp = await fetch(url);
    if (!resp.ok) throw new Error('HTTP ' + resp.status + ' ' + resp.statusText);
    const data = await resp.json();
    const count = Object.keys(data).length;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    btn.innerHTML = '✅ 已下载';
    updateStatus();
    showToast('✅ 题库下载完成，共 ' + count + ' 个数据项！', 'success');
  } catch (e) {
    showToast('❌ 下载失败（' + label + '）：' + e.message, 'error');
  } finally {
    setTimeout(() => {
      btn.innerHTML = '📥 下载题库';
      btn.disabled = false;
    }, 2000);
  }
}

function clearTiku() {
  localStorage.removeItem(STORAGE_KEY);
  updateStatus();
  showToast('✅ 本地题库已清理', 'success');
}

function updateStatus() {
  const statusEl = document.getElementById('wzw-status');
  if (!statusEl) return;
  const tiku = getTiku();
  if (tiku) {
    statusEl.innerHTML = `🟢 题库就绪（已缓存 ${Object.keys(tiku).length} 项）`;
    statusEl.className = 'wzw-status success';
  } else {
    statusEl.innerHTML = `🟡 题库未检测到，请先下载`;
    statusEl.className = 'wzw-status warning';
  }
}

function getSpread(){
  if(!window.GC||!window.GC.Spread||!window.GC.Spread.Sheets)return null;
  const{findControl}=GC.Spread.Sheets;
  for(const el of document.querySelectorAll('*')){
   const ctrl=findControl(el);
   if(ctrl&&ctrl instanceof GC.Spread.Sheets.Workbook)return ctrl;
  }
  return null;
}

function detectAllIndices(sheet){
  const mr=Math.min(sheet.getRowCount(),120), mc=Math.min(sheet.getColumnCount(),20);
  const pat=/^[A-Z]{1,3}\d*(?:-\d+)*$/;
  const idxMap={};
  for(let r=0;r<mr;r++)
   for(let c=0;c<mc;c++){
    const v=sheet.getValue(r,c);
    if(v){const s=v.toString().trim(); if(s.match(pat))idxMap[s]={idx:s,row:r,col:c};}
   }
  const result=Object.values(idxMap);
  if(result.length===0){
   const sname=sheet.name();
    if(sname){const m=sname.trim().match(/^([A-Z]{1,4}\d*(?:-\d+)*)/); if(m&&m[1].length>=2)result.push({idx:m[1],row:0,col:0});}
  }
  return result;
}

function calcOffset(sheet, idx, tikuCells){
  let sheetR=-1,sheetC=-1;
  const mr=Math.min(sheet.getRowCount(),120), mc=Math.min(sheet.getColumnCount(),20);
  for(let r=0;r<mr;r++) for(let c=0;c<mc;c++){
   const v=sheet.getValue(r,c);
   if(v&&v.toString().trim()===idx){sheetR=r;sheetC=c;break;}
   if(sheetR>=0)break;
  }
  let tikuR=-1,tikuC=-1;
  for(const[relR,relC,val]of tikuCells){
   if(val&&val.toString().trim()===idx){tikuR=relR;tikuC=relC;break;}
  }
  if(sheetR>=0&&tikuR>=0){
   return {rowOffset:sheetR-tikuR, colOffset:sheetC-tikuC};
  }
  return {rowOffset:2,colOffset:1};
}

function getBillText(doc, questionData){
  const parts=[];
  if(doc?.body?.innerText)parts.push(doc.body.innerText);
  if(questionData){
   for(const key of ['questionTitle','title','name','indexNo']){
    if(questionData[key])parts.push(String(questionData[key]));
   }
   const detail=questionData.stuQuestionDetail;
   if(detail){
    for(const key of ['questionName','title','fieldA']){
     if(detail[key])parts.push(typeof detail[key]==='string'?detail[key]:JSON.stringify(detail[key]));
    }
   }
  }
  return parts.join('\n');
}

function resolveBillFormAnswer(entry,doc,questionData={},fallbackFields={}){
  if(!entry||entry.type!=='billForm')return {fields:entry||{},rows:{}};
  const byCustomer=entry.byCustomer||{};
  const text=getBillText(doc,questionData);
  let customer=Object.keys(byCustomer).find(name=>text.includes(name));
  if(!customer&&fallbackFields){
   const fallbackText=Object.values(fallbackFields).map(v=>String(v||'')).join('\n');
   customer=Object.keys(byCustomer).find(name=>fallbackText.includes(name));
  }
  if(!customer){
   const match=text.match(/致[:：]\s*([^\n\r\t ]+?(?:有限公司|公司))/);
   if(match)customer=Object.keys(byCustomer).find(name=>name.includes(match[1])||match[1].includes(name));
  }
  const keys=Object.keys(byCustomer);
  const selected=customer?byCustomer[customer]:(keys.length===1?byCustomer[keys[0]]:null);
  if(!selected)return {fields:{},rows:{}};
  return {
   customer,
   fields:Object.assign({},selected.fields||{}),
   rows:selected.rows||{}
  };
}

function setInputValue(input,value,force=false){
  if(!input||input.readOnly)return 'readonly';
  const next=value==null?'':String(value);
  if(!force&&(input.value||'')===next)return 'skipped';
  input.value=next;
  for(const type of ['input','change','blur']){
   try{input.dispatchEvent(new Event(type,{bubbles:true}));}catch(e){}
  }
  return 'filled';
}

function fillInputsByName(doc,name,value,index,force=false){
  const inputs=Array.from(doc.getElementsByName(name)||[]);
  if(!inputs.length)return {filled:0,skipped:0};
  const targets=index==null?inputs:[inputs[index]].filter(Boolean);
  let filled=0,skipped=0;
  for(const input of targets){
   const status=setInputValue(input,value,force);
   if(status==='filled')filled++;
   else if(status==='skipped'||status==='readonly')skipped++;
  }
  return {filled,skipped};
}

function applyBillFormAnswer(doc,answer){
  const rowGroups={
   salesReceivable:{start:0,names:['b1','b2','b3','b4','b5']}
  };
  let filled=0,skipped=0;
  for(const[name,value]of Object.entries(answer.fields||{})){
   const stat=fillInputsByName(doc,name,value,null,true);
   filled+=stat.filled;skipped+=stat.skipped;
  }
  for(const[groupName,rows]of Object.entries(answer.rows||{})){
   const group=rowGroups[groupName];
   if(!group||!Array.isArray(rows))continue;
   rows.forEach((row,rowIndex)=>{
    const targetIndex=group.start+rowIndex;
    for(const name of group.names){
     if(!(name in row))continue;
     const stat=fillInputsByName(doc,name,row[name],targetIndex,true);
     filled+=stat.filled;skipped+=stat.skipped;
    }
   });
  }
  return {filled,skipped};
}

function autoFillBillForm(){
  const tiku = getTiku();
  if (!tiku) return false;
  const isInIframe=window!==window.top;
  let doc;
  if(isInIframe) doc=document; else {
   const ifr=document.querySelector('iframe[src*="billForm.html"]');
   if(!ifr) return false;
   try{doc=ifr.contentDocument;}catch(e){return false;}
   if(!doc) return false;
  }
  let store;
  try{store=JSON.parse(localStorage.getItem('auditVuex')||'{}');}catch(e){store={};}
  const questionData=store?.question?.questionData;
  const detail=questionData?.stuQuestionDetail;

  const answers={};
  if(detail?.fieldA){
   try{
    const fa=typeof detail.fieldA==='string'?JSON.parse(detail.fieldA):detail.fieldA;
    Object.assign(answers,fa);
   }catch(e){}
  }

  const sxBill=questionData?.sxBill;
  let billId='';
  if(sxBill?.billId) billId=sxBill.billId + (sxBill.version ? '@' + sxBill.version : '');

  const indexNo=questionData?.indexNo;
  let tikuFields=tiku[billId]||tiku[indexNo];
  let billAnswer={fields:answers,rows:{}};
  if(tikuFields&&!Array.isArray(tikuFields)){
   if(tikuFields.type==='billForm'){
    const resolved=resolveBillFormAnswer(tikuFields,doc,questionData,answers);
    billAnswer={fields:Object.assign({},answers,resolved.fields),rows:resolved.rows};
   }else{
    Object.assign(answers,tikuFields);
    billAnswer={fields:answers,rows:{}};
   }
  }
  if(Object.keys(billAnswer.fields||{}).length===0&&Object.keys(billAnswer.rows||{}).length===0)return false;
  const {filled,skipped}=applyBillFormAnswer(doc,billAnswer);
  if(filled>0||skipped>0) showToast('✅表单完成(填'+filled+'项,跳过'+skipped+')','success');
  return true;
}

async function autoFill(){
  const tiku = getTiku();
  if (!tiku) {
    showToast('⚠️ 题库未下载，请先点击「下载题库」', 'warning');
    return;
  }
  const btn=document.getElementById('auto-fill-btn');
  if(!btn) return;
  const originalHtml = btn.innerHTML;
  btn.innerHTML = '⏳ 自动识别中...';
  btn.disabled = true;
   try{
    if(autoFillBillForm()){
      showToast('✅ 表单自动填写完成', 'success');
      return;
    }
    if(!window.GC?.Spread?.Sheets){
     await new Promise((res,rej)=>{
      let n=0;
      const iv=setInterval(()=>{if(window.GC?.Spread?.Sheets){clearInterval(iv);res()}else if(n++>50){clearInterval(iv);rej()}},200);
     });
    }
    const spread=getSpread();
    if(!spread){showToast('⚠️ 未检测到适用页面');return;}
    
   const sheet=spread.getActiveSheet();
   const maxR=sheet.getRowCount();
   const maxC=sheet.getColumnCount();
   
    const indices=detectAllIndices(sheet);
    if(indices.length===0){showToast('⚠️ 未检测到索引号，该题目暂无题库答案');return;}
    
    const matched=indices.filter(i=>tiku[i.idx]);
    if(matched.length===0){showToast('⚠️ 题库无['+indices.map(i=>i.idx).join(',')+']数据，还未录入');return;}
   
    let totalFilled=0,totalSkipped=0,totalErrors=0;
    const msgs=[];
    matched.sort((a,b)=>a.row-b.row);
    const mainIdx=matched[0].idx;
    for(const {idx} of matched){
     const answer=tiku[idx];
     const[title,cells]=answer;
     const {rowOffset,colOffset}=calcOffset(sheet,idx,cells);
     msgs.push('['+idx+']'+cells.length+'项(偏移行'+rowOffset+'列'+colOffset+')');
     if(idx!==mainIdx)continue;
     
     let filled=0,skipped=0,errors=0;
    for(const[relR,relC,answerVal]of cells){
     const targetR=relR+rowOffset;
     const targetC=relC+colOffset;
     if(targetR>=maxR||targetC>=maxC)continue;
     const cur=sheet.getValue(targetR,targetC);
     const curStr=cur?cur.toString().trim().replace(/\u00a0/g,' '):'';
     const ansStr=answerVal.toString().trim().replace(/\u00a0/g,' ');
     if(curStr===ansStr){skipped++;continue;}
     try{
      sheet.setValue(targetR,targetC,answerVal);
      filled++;
     }catch(e){errors++;}
    }
    totalFilled+=filled;totalSkipped+=skipped;totalErrors+=errors;
   }
   
   showToast('📋'+msgs.join(' | '),'info');
   await new Promise(r=>setTimeout(r,1000));
   showToast('✅填写完成(填'+totalFilled+'项,跳过'+totalSkipped+',错误'+totalErrors+')','success');
  }catch(e){
   console.error(e);
   showToast('❌'+e.message,'error');
  } finally {
   btn.innerHTML = originalHtml;
   btn.disabled = false;
  }
}

function showToast(msg,type='info'){
  const old=document.getElementById('auto-fill-toast');
  if(old)old.remove();
  const c={success:'#10b981',error:'#ef4444',info:'#3b82f6',warning:'#f59e0b'};
  const t=document.createElement('div');
  t.id='auto-fill-toast';t.textContent=msg;
  t.style.cssText='position:fixed;top:20px;right:20px;background:'+c[type]+';color:white;padding:12px 20px;border-radius:8px;font-size:13px;z-index:9999999;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1);font-family:system-ui,sans-serif;pointer-events:none;';
  document.body.appendChild(t);
  setTimeout(()=>t.remove(),3000);
}

function init() {
  const s = document.createElement('style');
  s.textContent = `
    #wzw-audit-panel {
      position: fixed;
      bottom: 80px;
      right: 20px;
      width: 280px;
      background: rgba(255, 255, 255, 0.98);
      border: 1px solid rgba(0, 0, 0, 0.08);
      border-radius: 16px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
      backdrop-filter: blur(10px);
      z-index: 999999;
      font-family: system-ui, -apple-system, sans-serif;
      transition: width 0.2s, height 0.2s, border-radius 0.2s, box-shadow 0.2s;
      overflow: hidden;
    }
    #wzw-audit-panel.minimized {
      width: 50px;
      height: 50px;
      border-radius: 50%;
      cursor: pointer;
      background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 6px 20px rgba(79, 70, 229, 0.4);
    }
    #wzw-audit-panel.minimized:hover {
      transform: scale(1.05);
    }
    .wzw-header {
      padding: 12px 16px;
      background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
      color: #f8fafc;
      display: flex;
      align-items: center;
      justify-content: space-between;
      cursor: move;
      user-select: none;
    }
    .wzw-title-group {
      pointer-events: none;
    }
    .wzw-title {
      font-size: 13px;
      font-weight: 700;
    }
    .wzw-subtitle {
      font-size: 10px;
      opacity: 0.6;
    }
    .wzw-toggle-btn {
      background: none;
      border: none;
      color: white;
      cursor: pointer;
      font-size: 14px;
      padding: 2px 6px;
      border-radius: 4px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .wzw-toggle-btn:hover {
      background: rgba(255, 255, 255, 0.15);
    }
    .wzw-body {
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .wzw-status {
      font-size: 11px;
      padding: 8px 12px;
      border-radius: 8px;
      border-left: 3px solid #64748b;
      font-weight: 500;
    }
    .wzw-status.success {
      color: #0f5132;
      background: #d1e7dd;
      border-left-color: #10b981;
    }
    .wzw-status.warning {
      color: #664d03;
      background: #fff3cd;
      border-left-color: #f59e0b;
    }
    .wzw-btn {
      border: none;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 600;
      padding: 10px 16px;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
    }
    .wzw-btn-primary {
      background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
      color: white;
      box-shadow: 0 4px 6px rgba(79, 70, 229, 0.2);
      width: 100%;
    }
    .wzw-btn-primary:hover:not(:disabled) {
      box-shadow: 0 6px 12px rgba(79, 70, 229, 0.3);
      transform: translateY(-1px);
    }
    .wzw-btn-secondary {
      background: #f1f5f9;
      color: #334155;
      border: 1px solid #e2e8f0;
      flex: 1;
    }
    .wzw-btn-secondary:hover:not(:disabled) {
      background: #e2e8f0;
    }
    .wzw-btn-danger {
      background: #fff5f5;
      color: #e53e3e;
      border: 1px solid #fed7d7;
      flex: 1;
    }
    .wzw-btn-danger:hover:not(:disabled) {
      background: #e53e3e;
      color: white;
    }
    .wzw-btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    .wzw-row {
      display: flex;
      gap: 8px;
    }
    .wzw-select {
      width: 100%;
      padding: 7px 8px;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      font-size: 11px;
      background: #f8fafc;
      color: #334155;
      cursor: pointer;
      outline: none;
      box-sizing: border-box;
    }
    .wzw-select:hover {
      border-color: #cbd5e1;
    }
    .wzw-divider {
      height: 1px;
      background: #e2e8f0;
    }
    .wzw-bubble-icon {
      font-size: 22px;
      user-select: none;
      display: none;
    }
    #wzw-audit-panel.minimized .wzw-bubble-icon {
      display: block;
    }
    #wzw-audit-panel.minimized .wzw-header,
    #wzw-audit-panel.minimized .wzw-body {
      display: none;
    }
    .wzw-disclaimer {
      font-size: 9px;
      color: #94a3b8;
      line-height: 1.4;
      margin-top: 4px;
      text-align: center;
      border-top: 1px dashed #e2e8f0;
      padding-top: 8px;
      user-select: none;
    }
  `;
  document.head.appendChild(s);

  // 创建面板
  const panel = document.createElement('div');
  panel.id = 'wzw-audit-panel';

  // 创建最小化图标
  const bubbleIcon = document.createElement('div');
  bubbleIcon.className = 'wzw-bubble-icon';
  bubbleIcon.innerHTML = '🤖';
  panel.appendChild(bubbleIcon);

  // 创建头部
  const header = document.createElement('div');
  header.className = 'wzw-header';
  
  const titleGroup = document.createElement('div');
  titleGroup.className = 'wzw-title-group';
  
  const title = document.createElement('div');
  title.className = 'wzw-title';
  title.innerHTML = '🤖 WZW-Audit-tool';
  titleGroup.appendChild(title);
  
  const subtitle = document.createElement('div');
  subtitle.className = 'wzw-subtitle';
  subtitle.innerHTML = 'v8.1 • by Allen';
  titleGroup.appendChild(subtitle);
  
  header.appendChild(titleGroup);

  const toggleBtn = document.createElement('button');
  toggleBtn.className = 'wzw-toggle-btn';
  toggleBtn.innerHTML = '➖';
  header.appendChild(toggleBtn);
  
  panel.appendChild(header);

  // 创建主体内容
  const body = document.createElement('div');
  body.className = 'wzw-body';

  // 状态指示
  const statusEl = document.createElement('div');
  statusEl.id = 'wzw-status';
  body.appendChild(statusEl);

  // 主按钮：智能填表
  const fillBtn = document.createElement('button');
  fillBtn.id = 'auto-fill-btn';
  fillBtn.className = 'wzw-btn wzw-btn-primary';
  fillBtn.innerHTML = '🤖 自动填充表格/单据';
  fillBtn.addEventListener('click', autoFill);
  body.appendChild(fillBtn);

  // 分割线
  const divider = document.createElement('div');
  divider.className = 'wzw-divider';
  body.appendChild(divider);

  // 下载源选择
  const sel = document.createElement('select');
  sel.id = 'tiku-source-select';
  sel.className = 'wzw-select';
  const savedLabel = localStorage.getItem(SOURCE_KEY);
  for (const label of Object.keys(TIKU_SOURCES)) {
    const opt = document.createElement('option');
    opt.value = label;
    opt.textContent = label;
    if (label === savedLabel) opt.selected = true;
    sel.appendChild(opt);
  }
  body.appendChild(sel);

  // 辅助按键行 (下载 + 清理)
  const row = document.createElement('div');
  row.className = 'wzw-row';

  const dlBtn = document.createElement('button');
  dlBtn.id = 'dl-tiku-btn';
  dlBtn.className = 'wzw-btn wzw-btn-secondary';
  dlBtn.innerHTML = '📥 下载题库';
  dlBtn.addEventListener('click', downloadTiku);
  row.appendChild(dlBtn);

  const clearBtn = document.createElement('button');
  clearBtn.id = 'clear-tiku-btn';
  clearBtn.className = 'wzw-btn wzw-btn-danger';
  clearBtn.innerHTML = '🗑 清理';
  clearBtn.addEventListener('click', clearTiku);
  row.appendChild(clearBtn);

  body.appendChild(row);

  // 创建免责声明
  const disclaimerEl = document.createElement('div');
  disclaimerEl.className = 'wzw-disclaimer';
  disclaimerEl.innerHTML = '⚠️ 免责声明：本插件仅供学习辅助使用，不保证数据100%正确，所有自动填写的结果均需人工二次审核，禁止商业用途！';
  body.appendChild(disclaimerEl);

  panel.appendChild(body);

  document.body.appendChild(panel);

  // 初始化状态显示
  updateStatus();

  // 读取上次存储的面板状态
  if (localStorage.getItem('wzw_panel_minimized') === 'true') {
    panel.classList.add('minimized');
  }

  // 拖拽与点击的状态变量
  let isDragging = false;
  let dragMoved = false;
  let startX, startY;
  let initialX, initialY;

  // 展开面板事件
  panel.addEventListener('click', (e) => {
    if (dragMoved) return; // 如果刚才进行了拖拽，不触发点击展开
    if (panel.classList.contains('minimized')) {
      panel.classList.remove('minimized');
      localStorage.setItem('wzw_panel_minimized', 'false');
    }
  });

  // 收起面板事件
  toggleBtn.addEventListener('click', (e) => {
    e.stopPropagation(); // 阻止冒泡，避免触发 panel.click 再次展开
    panel.classList.add('minimized');
    localStorage.setItem('wzw_panel_minimized', 'true');
  });

  // 绑定拖拽事件到整个面板
  panel.addEventListener('mousedown', dragStart);

  function dragStart(e) {
    // 展开状态下只允许通过头部拖拽；收起状态下允许拖拽整个圆形气泡
    if (!panel.classList.contains('minimized') && !e.target.closest('.wzw-header')) {
      return;
    }
    // 如果点击的是收起按钮或功能按钮，不触发拖拽
    if (e.target.closest('.wzw-toggle-btn') || e.target.closest('.wzw-btn')) {
      return;
    }
    isDragging = true;
    dragMoved = false;
    startX = e.clientX;
    startY = e.clientY;
    
    const rect = panel.getBoundingClientRect();
    initialX = rect.left;
    initialY = rect.top;
    
    // 清除 right/bottom 布局，转为以 left/top 绝对定位，保证拖拽流畅
    panel.style.right = 'auto';
    panel.style.bottom = 'auto';
    panel.style.left = initialX + 'px';
    panel.style.top = initialY + 'px';
    
    document.addEventListener('mousemove', dragMove);
    document.addEventListener('mouseup', dragEnd);
    e.preventDefault();
  }

  function dragMove(e) {
    if (!isDragging) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    
    // 鼠标移动距离超过 5 像素判定为拖拽操作，防止误触点击
    if (Math.abs(dx) > 5 || Math.abs(dy) > 5) {
      dragMoved = true;
    }
    
    // 限制拖拽边界，不允许拖出浏览器可视区域
    const newX = Math.max(10, Math.min(window.innerWidth - panel.offsetWidth - 10, initialX + dx));
    const newY = Math.max(10, Math.min(window.innerHeight - panel.offsetHeight - 10, initialY + dy));
    
    panel.style.left = newX + 'px';
    panel.style.top = newY + 'px';
  }

  function dragEnd() {
    isDragging = false;
    document.removeEventListener('mousemove', dragMove);
    document.removeEventListener('mouseup', dragEnd);
    // 延时重置 dragMoved 状态，确保在 click 事件触发之后才重置
    setTimeout(() => { dragMoved = false; }, 50);
  }
}

if(window.__WZW_AUDIT_DISABLE_INIT__){
 window.__WZW_AUDIT_TESTS__={resolveBillFormAnswer,applyBillFormAnswer,fillInputsByName,setInputValue};
}else{
 document.readyState==='loading'?document.addEventListener('DOMContentLoaded',init):init();
}

})();
