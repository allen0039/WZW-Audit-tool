// ==UserScript==
// @name         shenjitools
// @namespace    http://tampermonkey.net/
// @version      8.0
// @description  审计实训万能填表助手 v8.0 - 在线题库下载/清理，脚本仅 3KB
// @author       AI
// @match        http://10.18.0.178:9350/*
// @grant        none
// ==/UserScript==

(function(){'use strict';

// ============================================================
// ★ 配置区：将 tiku.json 发布到 GitHub 后修改此 URL
//    支持 GitHub Raw、GitHub Pages、任何静态托管
// ============================================================
const TIKU_URL = 'https://raw.githubusercontent.com/allen0039/WZW-Audit-tool/main/tiku.json';
// ============================================================

const STORAGE_KEY = 'shenjitools_tiku';

function getTiku() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

async function downloadTiku() {
  const btn = document.getElementById('dl-tiku-btn');
  btn.textContent = '⏳下载中...';
  btn.style.opacity = '0.7';
  try {
    const resp = await fetch(TIKU_URL);
    if (!resp.ok) throw new Error('HTTP ' + resp.status + ' ' + resp.statusText);
    const data = await resp.json();
    const count = Object.keys(data).length;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    btn.innerHTML = '✅已下载';
    showToast('✅题库下载完成，共 ' + count + ' 个索引（含表单+单据答案），请刷新后使用', 'success');
    setTimeout(() => { btn.innerHTML = '📥下载题库'; btn.style.opacity = '1'; }, 3000);
  } catch (e) {
    showToast('❌下载失败：' + e.message, 'error');
    btn.innerHTML = '📥下载题库';
    btn.style.opacity = '1';
  }
}

function clearTiku() {
  localStorage.removeItem(STORAGE_KEY);
  const btn = document.getElementById('clear-tiku-btn');
  if (btn) {
    btn.innerHTML = '✅已清理';
    setTimeout(() => { btn.innerHTML = '🗑清理本地题库'; }, 2000);
  }
  showToast('✅本地题库已清理，再次使用需重新下载', 'success');
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

  const tikuFields=tiku[billId];
  if(tikuFields) Object.assign(answers,tikuFields);

  const entries=Object.entries(answers);
  if(entries.length===0) return false;
  let filled=0,skipped=0;
  for(const[name,value]of entries){
   const inputs=doc.getElementsByName(name);
   if(!inputs.length) continue;
   for(const input of inputs){
    if(input.readOnly) continue;
    if((input.value||'')===value){skipped++;continue;}
    input.value=value;filled++;
   }
  }
  if(filled>0||skipped>0) showToast('✅表单完成(填'+filled+'项,跳过'+skipped+')','success');
  return true;
}

async function autoFill(){
  const tiku = getTiku();
  if (!tiku) {
    showToast('⚠️题库未下载，请先点击「下载题库」按钮', 'warning');
    return;
  }
  const btn=document.getElementById('auto-fill-btn');
  btn.textContent='⏳识别中...';btn.style.opacity='0.7';
   try{
    if(autoFillBillForm()){btn.innerHTML='✅已填';setTimeout(()=>{btn.innerHTML='🤖万能填表';btn.style.opacity='1';},3000);return;}
    if(!window.GC?.Spread?.Sheets){
     await new Promise((res,rej)=>{
      let n=0;
      const iv=setInterval(()=>{if(window.GC?.Spread?.Sheets){clearInterval(iv);res();}else if(n++>50){clearInterval(iv);rej();}},200);
     });
    }
    const spread=getSpread();
    if(!spread){showToast('⚠️未检测到适用页面');btn.innerHTML='🤖万能填表';btn.style.opacity='1';return;}
    
   const sheet=spread.getActiveSheet();
   const maxR=sheet.getRowCount();
   const maxC=sheet.getColumnCount();
   
    const indices=detectAllIndices(sheet);
    if(indices.length===0){showToast('⚠️未检测到索引号，该题目暂无题库答案');btn.innerHTML='🤖万能填表';btn.style.opacity='1';return;}
    
    const matched=indices.filter(i=>tiku[i.idx]);
    if(matched.length===0){showToast('⚠️题库无['+indices.map(i=>i.idx).join(',')+']数据，还未录入');btn.innerHTML='🤖万能填表';btn.style.opacity='1';return;}
   
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
   await new Promise(r=>setTimeout(r,1500));
   showToast('✅填写完成(填'+totalFilled+'项,跳过'+totalSkipped+',错误'+totalErrors+')','success');
   btn.innerHTML='✅已填';
   setTimeout(()=>{btn.innerHTML='🤖万能填表';btn.style.opacity='1';},3000);
  }catch(e){
   console.error(e);
   showToast('❌'+e.message,'error');
   btn.innerHTML='🤖万能填表';btn.style.opacity='1';
  }
}

function showToast(msg,type='info'){
  const old=document.getElementById('auto-fill-toast');
  if(old)old.remove();
  const c={success:'#4CAF50',error:'#f44336',info:'#2196F3',warning:'#FF9800'};
  const t=document.createElement('div');
  t.id='auto-fill-toast';t.textContent=msg;
  t.style.cssText='position:fixed;top:20px;right:20px;background:'+c[type]+';color:white;padding:15px 25px;border-radius:8px;font-size:14px;z-index:9999999;box-shadow:0 4px 12px rgba(0,0,0,0.3);';
  document.body.appendChild(t);
  setTimeout(()=>t.remove(),3000);
}

function init(){
  const s=document.createElement('style');
  s.textContent='#auto-fill-btn{position:fixed;bottom:100px;right:20px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:12px 20px;border-radius:25px;cursor:pointer;font-size:14px;font-weight:bold;box-shadow:0 4px 15px rgba(0,0,0,0.3);z-index:999999;user-select:none}#dl-tiku-btn{position:fixed;bottom:150px;right:20px;background:linear-gradient(135deg,#43e97b 0%,#38f9d7 100%);color:white;padding:8px 16px;border-radius:20px;cursor:pointer;font-size:12px;font-weight:bold;box-shadow:0 4px 12px rgba(0,0,0,0.2);z-index:999999;user-select:none}#clear-tiku-btn{position:fixed;bottom:190px;right:20px;background:linear-gradient(135deg,#fa709a 0%,#fee140 100%);color:white;padding:8px 16px;border-radius:20px;cursor:pointer;font-size:12px;font-weight:bold;box-shadow:0 4px 12px rgba(0,0,0,0.2);z-index:999999;user-select:none}';
  document.head.appendChild(s);

  const dlBtn=document.createElement('div');
  dlBtn.id='dl-tiku-btn';dlBtn.innerHTML='📥下载题库';
  dlBtn.addEventListener('click',downloadTiku);
  document.body.appendChild(dlBtn);

  const clearBtn=document.createElement('div');
  clearBtn.id='clear-tiku-btn';clearBtn.innerHTML='🗑清理本地题库';
  clearBtn.addEventListener('click',clearTiku);
  document.body.appendChild(clearBtn);

  const btn=document.createElement('div');
  btn.id='auto-fill-btn';btn.innerHTML='🤖万能填表';
  btn.addEventListener('click',autoFill);
  document.body.appendChild(btn);

  const tiku=getTiku();
  if (tiku) {
    showToast('✅万能填表就绪（本地题库 '+Object.keys(tiku).length+' 项）','info');
  } else {
    showToast('📥首次使用请点击「下载题库」按钮','warning');
  }
}

document.readyState==='loading'?document.addEventListener('DOMContentLoaded',init):init();

})();
