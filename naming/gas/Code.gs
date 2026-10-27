// 公館運動館 命名總表：免登入儲存後端
// 部署：script.google.com 新專案 → 貼上本檔 → 部署 → 新增部署作業 → 網頁應用程式
//      執行身分：我；誰可以存取：任何人 → 複製 /exec 網址，貼到 index.html 的 GAS_URL
const SHEET_ID = ""; // 留空＝使用綁定試算表；也可貼指定試算表 ID

function _ss(){ return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function _sheet(name){
  var ss=_ss(), sh=ss.getSheetByName(name);
  if(!sh){ sh=ss.insertSheet(name); sh.appendRow(["id","final","pick","updated"]); }
  return sh;
}
function _json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){
  var out={names:{},areas:{}};
  ["names","areas"].forEach(function(n){
    var v=_sheet(n).getDataRange().getValues();
    for(var i=1;i<v.length;i++){ if(!v[i][0])continue;
      out[n][String(v[i][0])]={final:String(v[i][1]||""),pick:String(v[i][2]||"")}; }
  });
  return _json(out);
}

function doPost(e){
  var lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    var d=JSON.parse(e.postData.contents);
    var col = d.col==="areas" ? "areas" : "names";
    var id = String(d.id||"").slice(0,100);
    if(!id) return _json({ok:false,error:"no id"});
    var fin = String(d.final||"").slice(0,200);
    var pick = String(d.pick||"").slice(0,10);
    var sh=_sheet(col), v=sh.getDataRange().getValues(), row=-1;
    for(var i=1;i<v.length;i++){ if(String(v[i][0])===id){ row=i+1; break; } }
    var rec=[id,fin,pick,new Date()];
    if(row>0) sh.getRange(row,1,1,4).setValues([rec]); else sh.appendRow(rec);
    return _json({ok:true});
  }catch(err){ return _json({ok:false,error:String(err)});
  }finally{ lock.releaseLock(); }
}
