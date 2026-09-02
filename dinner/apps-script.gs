const SYSTEM_SHEET = '晚餐系統';

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  if (action !== 'list') return json_({ ok: false, error: 'unsupported_action' });
  const date = (e.parameter.date || '').trim();
  const sh = getSystemSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  const people = values
    .filter(r => String(r[idx.date] || '').trim() === date)
    .map(r => ({
      name: String(r[idx.name] || '').trim(),
      qty: Number(r[idx.qty]) || 1,
      paid: bool_(r[idx.paid]),
      picked: bool_(r[idx.picked])
    }))
    .filter(p => p.name);
  return json_({ ok: true, date, people });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const action = data.action;
    if (action === 'order') return json_(upsertOrder_(data));
    if (action === 'status') return json_(updateStatus_(data));
    return json_({ ok: false, error: 'unsupported_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getSystemSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SYSTEM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SYSTEM_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['日期','姓名','份數','單價','應收','已付款','已取餐','更新時間']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function upsertOrder_(data) {
  const sh = getSystemSheet_();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  const qty = Math.max(1, Number(data.qty) || 1);
  const price = Math.max(0, Number(data.price) || 110);
  if (!date || !name) return { ok: false, error: 'missing_date_or_name' };

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === date && String(values[i][1]).trim() === name) {
      const nextQty = (Number(values[i][2]) || 0) + qty;
      sh.getRange(i + 1, 3, 1, 6).setValues([[nextQty, price, nextQty * price, bool_(values[i][5]), bool_(values[i][6]), new Date()]]);
      return { ok: true, updated: true, date, name, qty: nextQty };
    }
  }
  sh.appendRow([date, name, qty, price, qty * price, false, false, new Date()]);
  return { ok: true, created: true, date, name, qty };
}

function updateStatus_(data) {
  const sh = getSystemSheet_();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  if (!date || !name) return { ok: false, error: 'missing_date_or_name' };

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === date && String(values[i][1]).trim() === name) {
      sh.getRange(i + 1, 6, 1, 3).setValues([[Boolean(data.paid), Boolean(data.picked), new Date()]]);
      return { ok: true, updated: true, date, name };
    }
  }
  return { ok: false, error: 'record_not_found' };
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return {
    date: map['日期'], name: map['姓名'], qty: map['份數'],
    paid: map['已付款'], picked: map['已取餐']
  };
}

function bool_(v) {
  return v === true || ['true','1','yes','y','是','已付款','已取餐'].includes(String(v).trim().toLowerCase());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
