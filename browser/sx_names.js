// Saudi Exchange names check for Mapping.xlsx (monthly, Task 0; added 29 Oct 2026).
// Run in the built-in browser on any www.saudiexchange.sa page (Chrome with the Saudi VPN if blocked).
// Paste this file, then call, in the same javascript_tool step:
//   await SX.diff("<digest printed by refresh_mapping.py --digest>")
// It returns only what differs from the sheet, so the full directory never enters the chat:
//   new     listed now, not in the sheet: [code, board, ISIN, short EN, long EN, short AR, long AR]
//   gone    in the sheet, not listed now (suspended, delisted or old Nomu code)
//   market  board differs from the sheet: [code, sheet, exchange]
//   changed names or ISIN differ: [code, ISIN, short EN, long EN, short AR, long AR]
// Sectors: on the main market watch page (ourmarkets/main-market-watch?locale=en) and the parallel
// one (ourmarkets/nomuc-market-watch?locale=en) call  SX.sectors(["4145", ...])  after the table has
// loaded; it returns {code: sector} for the codes asked (only new or moved companies need it).
var SX = (function () {
  function norm(s) { return String(s || '').replace(/ی/g, 'ي').replace(/ک/g, 'ك')
                       .replace(/\s*\*\s*$/, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
  function fnv(s) {               // FNV-1a 32-bit over UTF-8, same as refresh_mapping.py
    var b = new TextEncoder().encode(s), h = 0x811c9dc5;
    for (var i = 0; i < b.length; i++) { h ^= b[i]; h = Math.imul(h, 0x01000193) >>> 0; }
    return ('0000000' + h.toString(16)).slice(-8).slice(0, 6);
  }
  function key(x) { return [x.companyNameEN, x.tradingNameEn, x.tradingNameAr, x.isin].map(norm).join('|'); }
  async function diff(digest) {
    var want = {};
    digest.split(',').forEach(function (t) { var p = t.split(':'); if (p[0]) want[p[0]] = {m: p[1], h: p[2]}; });
    var dir = await (await fetch('/tadawul.eportal.theme.helper/ThemeSearchUtilityServlet')).json();
    var tick = await (await fetch('/tadawul.eportal.theme.helper/TickerServlet')).json();
    var listed = {};
    (tick.stockData || []).forEach(function (s) { listed[s.pk_rf_company] = 1; });
    var out = {checked: 0, new: [], gone: [], market: [], changed: []};
    var seen = {};
    dir.forEach(function (x) {
      if (!/^\d{4}$/.test(x.symbol) || (x.market_type !== 'M' && x.market_type !== 'S') || !listed[x.symbol]) return;
      seen[x.symbol] = 1; out.checked++;
      var board = x.market_type === 'M' ? 'Main' : 'Nomu', w = want[x.symbol];
      var row = [x.symbol, board, x.isin, (x.tradingNameEn || '').trim(), (x.companyNameEN || '').trim(),
                 norm(x.tradingNameAr), norm(x.companyNameAR)];
      if (!w) { out.new.push(row); return; }
      if (w.m !== board[0]) out.market.push([x.symbol, w.m, board]);
      if (w.h !== fnv(key(x))) out.changed.push([x.symbol].concat(row.slice(2)));
    });
    Object.keys(want).forEach(function (c) { if (!seen[c]) out.gone.push(c); });
    return out;
  }
  function sectors(codes) {
    var cur = null, m = {};
    document.querySelectorAll('table tbody tr').forEach(function (tr) {
      if (tr.classList.contains('dtrg-group') || tr.classList.contains('dtrg-start')) { cur = tr.innerText.trim(); return; }
      var x = tr.innerText.match(/\b(\d{4})\b/);
      if (x && cur && (!codes || codes.indexOf(x[1]) >= 0)) m[x[1]] = cur;
    });
    return m;
  }
  return {diff: diff, sectors: sectors, fnv: fnv};
})();
