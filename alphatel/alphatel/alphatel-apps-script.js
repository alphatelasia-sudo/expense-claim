// ═══════════════════════════════════════════════════════════
//  Alphatel Asia — Expense Claim Apps Script
// ═══════════════════════════════════════════════════════════

const LARK_MANAGER_WEBHOOK = 'https://open.larksuite.com/open-apis/bot/v2/hook/f624f835-cd10-4a13-91e7-d501173530b6';
const LARK_ACCOUNTING_WEBHOOK = 'https://open.larksuite.com/open-apis/bot/v2/hook/791df9ef-9c2d-4118-afa0-9d01d9c83888';
const APPROVAL_BASE_URL = 'https://alphatelasia-sudo.github.io/expense-claim/alphatel/approval.html';
const COMPANY_NAME = 'Alphatel Asia';

function doGet(e) {
  const callback = e.parameter.callback;
  const action = e.parameter.action;
  let result = {};

  try {
    if (action === 'getNextDocNo') {
      result = { success: true, docNo: getNextDocNo() };
    } else if (action === 'saveClaim') {
      const data = JSON.parse(e.parameter.data);
      saveClaim(data);
      result = { success: true };
    } else if (action === 'notifyLark') {
      const data = JSON.parse(e.parameter.data);
      notifyLark(data);
      result = { success: true };
    } else if (action === 'approvalDecision') {
      const data = JSON.parse(e.parameter.data);
      approvalDecision(data);
      result = { success: true };
    } else if (action === 'getClaim') {
      const claim = getClaim(e.parameter.docNo);
      result = claim ? { success: true, data: claim } : { success: false, error: 'ไม่พบข้อมูล' };
    } else {
      result = { success: false, error: 'Unknown action' };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }

  const output = callback
    ? ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ── Doc No. ───────────────────────────────────────────────────────
function getNextDocNo() {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const now = new Date();
  const yy = now.getFullYear().toString().slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const currentMonth = parseInt(yy + mm);

  const lastNum   = parseInt(config.getRange('B1').getValue()) || 0;
  const lastMonth = parseInt(config.getRange('B2').getValue()) || 0;

  const nextNum = (currentMonth !== lastMonth) ? 1 : lastNum + 1;

  config.getRange('B1').setValue(nextNum);
  config.getRange('B2').setValue(currentMonth);

  const seq = String(nextNum).padStart(3, '0');
  return 'AC-' + yy + mm + seq;
}

// ── Save Claim ────────────────────────────────────────────────────
function saveClaim(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Doc No.','Request By','Date','To','CC','Objective',
      'Detail','Project','Item1','Amt1','Item2','Amt2','Item3','Amt3',
      'Item4','Amt4','Item5','Amt5','Item6','Amt6','Item7','Amt7',
      'Item8','Amt8','Item9','Amt9','Total','Saved At','Status','Approved By','Comment']);
  }
  const itemCols = [];
  (data.items || []).forEach(function(it) {
    itemCols.push(it.desc || '', it.amt || '');
  });
  sheet.appendRow([
    data.docNo, data.requestBy, data.requestDate,
    data.toField, data.ccField, data.objective,
    data.otherObjective, data.project
  ].concat(itemCols).concat([
    data.total,
    new Date().toLocaleString('th-TH'),
    'รอการอนุมัติ', '', ''
  ]));
}

// ── Get Claim ─────────────────────────────────────────────────────
function getClaim(docNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docNo) {
      const row = data[i];
      const items = [];
      for (let j = 8; j < 26; j += 2) {
        if (row[j]) items.push({ desc: row[j], amt: row[j+1] });
      }
      return {
        docNo: row[0], requestBy: row[1], requestDate: row[2],
        toField: row[3], ccField: row[4], objective: row[5],
        otherObjective: row[6], project: row[7],
        total: row[26], items: items
      };
    }
  }
  return null;
}

// ── Notify Lark (หัวหน้า) ─────────────────────────────────────────
function notifyLark(data) {
  const approvalUrl = APPROVAL_BASE_URL + '?docNo=' + encodeURIComponent(data.docNo);
  const total = parseFloat(data.total || 0).toLocaleString('th-TH', {minimumFractionDigits: 2});

  const payload = {
    msg_type: 'interactive',
    card: {
      header: {
        title: { tag: 'plain_text', content: '📋 [' + COMPANY_NAME + '] ใบเบิกใหม่รอการอนุมัติ' },
        template: 'blue'
      },
      elements: [
        {
          tag: 'div',
          fields: [
            { is_short: true, text: { tag: 'lark_md', content: '**เลขที่:** ' + data.docNo } },
            { is_short: true, text: { tag: 'lark_md', content: '**ผู้เบิก:** ' + data.requestBy } },
            { is_short: true, text: { tag: 'lark_md', content: '**วัตถุประสงค์:** ' + (data.objective || '-') } },
            { is_short: true, text: { tag: 'lark_md', content: '**ยอดรวม:** ฿' + total } }
          ]
        },
        {
          tag: 'action',
          actions: [
            {
              tag: 'button',
              text: { tag: 'plain_text', content: '✅ เปิดอนุมัติ' },
              type: 'primary',
              url: approvalUrl
            }
          ]
        }
      ]
    }
  };

  UrlFetchApp.fetch(LARK_MANAGER_WEBHOOK, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

// ── Approval Decision ─────────────────────────────────────────────
function approvalDecision(payload) {
  const docNo = payload.docNo;
  const decision = payload.decision;
  const approverName = payload.approverName;
  const comment = payload.comment || '';
  const claimData = payload.claimData;
  const total = parseFloat(claimData.total || 0).toLocaleString('th-TH', {minimumFractionDigits: 2});

  // อัปเดตสถานะใน Google Sheets
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docNo) {
      sheet.getRange(i + 1, 29).setValue(decision === 'approved' ? 'อนุมัติ' : 'ไม่อนุมัติ');
      sheet.getRange(i + 1, 30).setValue(approverName);
      sheet.getRange(i + 1, 31).setValue(comment);
      break;
    }
  }

  if (decision === 'approved') {
    const msg = {
      msg_type: 'interactive',
      card: {
        header: {
          title: { tag: 'plain_text', content: '✅ [' + COMPANY_NAME + '] ใบเบิกผ่านการอนุมัติแล้ว' },
          template: 'green'
        },
        elements: [{
          tag: 'div',
          fields: [
            { is_short: true, text: { tag: 'lark_md', content: '**เลขที่:** ' + docNo } },
            { is_short: true, text: { tag: 'lark_md', content: '**ผู้เบิก:** ' + claimData.requestBy } },
            { is_short: true, text: { tag: 'lark_md', content: '**อนุมัติโดย:** ' + approverName } },
            { is_short: true, text: { tag: 'lark_md', content: '**ยอดรวม:** ฿' + total } }
          ]
        }]
      }
    };
    UrlFetchApp.fetch(LARK_ACCOUNTING_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(msg)
    });
  } else {
    const msg = {
      msg_type: 'text',
      content: { text: '❌ [' + COMPANY_NAME + '] ใบเบิก ' + docNo + ' ไม่ได้รับการอนุมัติ\nโดย: ' + approverName + (comment ? '\nเหตุผล: ' + comment : '') }
    };
    UrlFetchApp.fetch(LARK_MANAGER_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(msg)
    });
  }
}

// ── Test function ─────────────────────────────────────────────────
function testNotify() {
  notifyLark({
    docNo: 'AC-TEST001',
    requestBy: 'ทดสอบระบบ',
    objective: 'Petty Cash',
    total: 100
  });
}
