const SYSTEM_SHEET = '晚餐系統';

function doGet(e) {
  try {
    const action = String((e.parameter && e.parameter.action) || '').trim();
    const date = String((e.parameter && e.parameter.date) || '').trim();
    if (action === 'list') return json_(buildDay_(date));
    if (action === 'names') return json_({ok:true, date:date, names:getRosterNames_(date)});
    return json_({ok:false, error:'unsupported_action'});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (data.action === 'order') return json_(upsertOrder_(data));
    if (data.action === 'status') return json_(updateStatus_(data));
    return json_({ok:false, error:'unsupported_action'});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}

function getSystemSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SYSTEM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SYSTEM_SHEET);
    sh.getRange(1,1,1,8).setValues([['日期','姓名','份數','單價','應收','已付款','已取餐','更新時間']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function buildDay_(date) {
  if (!date) return {ok:false, error:'missing_date'};
  const monthly = getMonthlyDinnerOrders_(date);
  const system = getSystemRows_(date);
  const map = {};

  monthly.forEach(p => {
    map[p.name] = {date:date, name:p.name, qty:p.qty, paid:false, picked:false, source:'月表'};
  });

  system.forEach(p => {
    if (map[p.name]) {
      map[p.name].qty += Math.max(0, p.qty);
      map[p.name].paid = p.paid;
      map[p.name].picked = p.picked;
      if (p.qty > 0) map[p.name].source = '月表＋網頁';
    } else if (p.qty > 0) {
      map[p.name] = {date:date, name:p.name, qty:p.qty, paid:p.paid, picked:p.picked, source:'網頁'};
    }
  });

  const people = Object.keys(map).map(k => map[k]).filter(p => p.name && p.qty > 0);
  const names = getRosterNames_(date);
  return {ok:true, date:date, people:people, names:names};
}

function getMonthlyDinnerOrders_(date) {
  const info = parseDate_(date);
  if (!info) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(info.sheetName);
  if (!sh) return [];

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 3) return [];
  const row1 = values[0] || [];
  const row2 = values[1] || [];
  const datePrefix = info.month + '/' + info.day;
  let dateCol = -1;

  for (let c = 0; c < row1.length; c++) {
    const s = String(row1[c] || '').trim();
    if (s === datePrefix || s.indexOf(datePrefix + '(') === 0 || s.indexOf(datePrefix + '（') === 0) {
      dateCol = c; break;
    }
  }
  if (dateCol < 0) return [];

  let dinnerCol = -1;
  for (let c = dateCol; c <= Math.min(dateCol + 2, row2.length - 1); c++) {
    if (String(row2[c] || '').trim() === '晚餐') { dinnerCol = c; break; }
  }
  if (dinnerCol < 0) return [];

  const out = [];
  for (let r = 2; r < values.length; r++) {
    const name = String((values[r] && values[r][1]) || '').trim();
    if (!name) continue;
    const raw = String((values[r] && values[r][dinnerCol]) || '').trim();
    const qty = Number(raw);
    if (isFinite(qty) && qty > 0) out.push({name:name, qty:qty});
  }
  return out;
}

function getRosterNames_(date) {
  const info = parseDate_(date);
  if (!info) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(info.sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return [];
  const names = sh.getRange(3,2,lastRow-2,1).getDisplayValues()
    .map(r => String(r[0] || '').trim())
    .filter(Boolean);
  return Array.from(new Set(names));
}

function getSystemRows_(date) {
  const sh = getSystemSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    if (normDate_(values[i][0]) !== date) continue;
    const name = String(values[i][1] || '').trim();
    if (!name) continue;
    out.push({
      row:i+1,
      name:name,
      qty:Number(values[i][2]) || 0,
      paid:bool_(values[i][5]),
      picked:bool_(values[i][6])
    });
  }
  return out;
}

function upsertOrder_(data) {
  const sh = getSystemSheet_();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  const qty = Math.max(1, Number(data.qty) || 1);
  const price = Math.max(0, Number(data.price) || 110);
  if (!date || !name) return {ok:false, error:'missing_date_or_name'};

  const rows = getSystemRows_(date);
  const old = rows.find(r => r.name === name);
  if (old) {
    const nextQty = Math.max(0, old.qty) + qty;
    sh.getRange(old.row,3,1,6).setValues([[nextQty,price,nextQty*price,old.paid,old.picked,new Date()]]);
    return {ok:true, updated:true, date:date, name:name, qty:nextQty};
  }
  sh.appendRow([date,name,qty,price,qty*price,false,false,new Date()]);
  return {ok:true, created:true, date:date, name:name, qty:qty};
}

function updateStatus_(data) {
  const sh = getSystemSheet_();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  if (!date || !name) return {ok:false, error:'missing_date_or_name'};
  const paid = Boolean(data.paid);
  const picked = Boolean(data.picked);

  const rows = getSystemRows_(date);
  const old = rows.find(r => r.name === name);
  if (old) {
    sh.getRange(old.row,6,1,3).setValues([[paid,picked,new Date()]]);
    return {ok:true, updated:true, date:date, name:name};
  }

  sh.appendRow([date,name,0,110,0,paid,picked,new Date()]);
  return {ok:true, createdStatus:true, date:date, name:name};
}

function parseDate_(date) {
  const m = String(date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return {year:Number(m[1]), month:Number(m[2]), day:Number(m[3]), sheetName:m[1] + m[2]};
}

function normDate_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return String(v || '').trim();
}

function bool_(v) {
  return v === true || ['true','1','yes','y','是','已付款','已取餐'].includes(String(v).trim().toLowerCase());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
