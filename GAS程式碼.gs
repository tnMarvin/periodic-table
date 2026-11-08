// 化學式組合遊戲：排行榜（Google Apps Script）
// 貼到「Google 試算表 → 擴充功能 → Apps Script」，部署為網頁應用程式。
// 成績會存在試算表的「成績」工作表；要刪除不當暱稱或清空榜單，直接在試算表刪除列即可。

const SHEET = '成績';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['時間', '暱稱', '類型(q=化學式,g=認識元素)', '毫秒']);
  }
  return sh;
}

function clean_(s) {
  return String(s || '').replace(/[\u0000-\u001f]/g, '').trim().replace(/^[=+\-@]+/, '').slice(0, 12);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 每位玩家：化學式累積毫秒/次數、認識元素累積毫秒/次數
function players_() {
  const v = sheet_().getDataRange().getValues(), m = {};
  for (let i = 1; i < v.length; i++) {
    const nick = String(v[i][1] || ''), kind = String(v[i][2]), ms = Number(v[i][3]);
    if (!nick || !(ms > 0)) continue;
    const p = m[nick] || (m[nick] = { nick: nick, qT: 0, qN: 0, gT: 0, gN: 0 });
    if (kind === 'g') { p.gT += ms; p.gN++; } else { p.qT += ms; p.qN++; }
  }
  return Object.keys(m).map(function (k) { return m[k]; });
}

const avgAll_ = function (p) { return (p.qT + p.gT) / (p.qN + p.gN); };
const avgQ_ = function (p) { return p.qT / p.qN; };
const avgG_ = function (p) { return p.gT / p.gN; };

function boards_() {
  const P = players_();
  return {
    all: P.slice().sort(function (a, b) { return avgAll_(a) - avgAll_(b); }),
    q: P.filter(function (p) { return p.qN > 0; }).sort(function (a, b) { return avgQ_(a) - avgQ_(b); }),
    g: P.filter(function (p) { return p.gN > 0; }).sort(function (a, b) { return avgG_(a) - avgG_(b); })
  };
}

function rankOf_(list, nick) {
  for (let i = 0; i < list.length; i++) if (list[i].nick === nick) return i + 1;
  return 0;
}

// 讀取排行榜：…/exec?action=board
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'board') {
      const b = boards_();
      return out_({ ok: true, all: b.all.slice(0, 30), q: b.q.slice(0, 30), g: b.g.slice(0, 30) });
    }
    return out_({ ok: true, msg: '化學式組合遊戲排行榜運作中' });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  }
}

// 送出成績：POST {"action":"submit","nick":"小明","kind":"q 或 g","ms":123456}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (d.action !== 'submit') return out_({ ok: false, err: '未知的動作' });
    const nick = clean_(d.nick), kind = d.kind === 'g' ? 'g' : 'q', ms = Number(d.ms);
    if (!nick) return out_({ ok: false, err: '暱稱無效' });
    if (!(ms >= 3000 && ms <= 3600000)) return out_({ ok: false, err: '成績無效' });
    sheet_().appendRow([new Date(), nick, kind, Math.round(ms)]);
    SpreadsheetApp.flush();
    const b = boards_(), me = b.all.filter(function (p) { return p.nick === nick; })[0];
    return out_({
      ok: true,
      rank: { all: rankOf_(b.all, nick), q: rankOf_(b.q, nick), g: rankOf_(b.g, nick) },
      n: { all: b.all.length, q: b.q.length, g: b.g.length },
      avg: avgAll_(me),
      count: me.qN + me.gN
    });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally {
    lock.releaseLock();
  }
}
