// মেসার্স সাদ এন্টারপ্রাইজ — গুগল শিট সিঙ্ক (Google Apps Script)
const TOKEN = 'এখানে-নিজের-গোপন-কোড-লিখুন'; // অবশ্যই বদলান (ইংরেজি অক্ষর ও সংখ্যা ব্যবহার করা ভালো)
const COLS = {
  p: ['id','name','cat','brand','unit','price','stock','min'],
  s: ['id','t','date','cust','ph','pid','pname','qty','price','total','paid'],
  pay: ['id','t','date','cust','ph','amount','note']
};
const NAMES = {p:'পণ্য', s:'বিক্রয়', pay:'বাকি জমা'};
const TEXT = ['id','name','cat','brand','unit','date','cust','ph','pid','pname','note'];
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const ver = () => Number(PropertiesService.getScriptProperties().getProperty('v')) || 0;
function sheet(k){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(NAMES[k]);
  if(!sh){ sh = ss.insertSheet(NAMES[k]); sh.appendRow(COLS[k]); sh.setFrozenRows(1); sh.getRange(1,1,1,COLS[k].length).setFontWeight('bold').setBackground('#14243B').setFontColor('#ffffff'); }
  return sh;
}
function doGet(e){
  if(e.parameter.token !== TOKEN) return out({ok:false,error:'auth'});
  if(e.parameter.action === 'ver') return out({ok:true,v:ver()});
  const d = {};
  for(const k in COLS){
    const r = sheet(k).getDataRange().getValues(); r.shift();
    d[k] = r.filter(x => x[0] !== '').map(x => { const o = {}; COLS[k].forEach((c,i) => o[c] = x[i]); return o; });
  }
  return out({ok:true,v:ver(),d});
}
function doPost(e){
  const b = JSON.parse(e.postData.contents);
  if(b.token !== TOKEN) return out({ok:false,error:'auth'});
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    if(!b.force && b.base !== ver()) return out({ok:false,conflict:true,v:ver()});
    for(const k in COLS){
      const sh = sheet(k), n = COLS[k].length, rows = (b.d[k] || []).map(o => COLS[k].map(c => o[c] === undefined ? '' : o[c]));
      sh.clearContents();
      COLS[k].forEach((c,i) => { if(TEXT.indexOf(c) >= 0) sh.getRange(1,i+1,sh.getMaxRows(),1).setNumberFormat('@'); });
      sh.getRange(1,1,1,n).setValues([COLS[k]]);
      if(rows.length) sh.getRange(2,1,rows.length,n).setValues(rows);
    }
    const v = Date.now(); PropertiesService.getScriptProperties().setProperty('v', String(v));
    return out({ok:true,v});
  } finally { lock.releaseLock(); }
}
