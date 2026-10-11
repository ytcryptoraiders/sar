// ====== সার চেকার – Google Sheet রিসিভার ======
// ১) এই কোড Google Sheet-এর Extensions > Apps Script-এ পেস্ট করুন।
// ২) নিচের TOKEN নিজের একটি গোপন শব্দ দিয়ে বদলান (অ্যাপেও একই শব্দ দেবেন)।
// ৩) Deploy > New deployment > Web app
//      Execute as: Me
//      Who has access: Anyone
//    তারপর যে URL পাবেন সেটি অ্যাপে বসান।
const TOKEN = 'এখানে-নিজের-গোপন-শব্দ-লিখুন';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out({ ok: false, error: 'টোকেন মেলেনি' });

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sh.getLastRow() === 0) {
      sh.appendRow(['NID', 'নাম', 'কেজি', 'সময়']);
      sh.getRange('A:A').setNumberFormat('@'); // NID টেক্সট হিসেবে থাকবে
    }
    const have = new Set();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => have.add(String(r[0])));
    }
    const fresh = [];
    (d.rows || []).forEach(r => {
      const nid = String(r.nid);
      if (!have.has(nid)) { have.add(nid); fresh.push([nid, r.name || '', r.kg || 25, r.time || '']); }
    });
    if (fresh.length) {
      sh.getRange(sh.getLastRow() + 1, 1, fresh.length, 4).setValues(fresh);
    }
    return out({ ok: true, added: fresh.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
