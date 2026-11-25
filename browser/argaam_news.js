// Argaam news, compact (open item 2, 29 Oct 2026).
// Run with the built-in browser's javascript_tool on any www.argaam.com tab: it fetches the feeds
// and list pages itself (same origin), so no navigation and no page text enter the chat.
// Order agreed with the user: RSS feeds first, list pages only for items the feeds lack; a helper
// agent only if Argaam changes its layout and this script returns nothing.
// First call in a tab: set CFG and paste the whole file as the javascript_tool text. It also
// defines window.argaamNews, so later calls in the same tab (no navigation in between) are one
// line: await argaamNews({sinceId: 1948269, text: ["1948301"]})
const CFG = {
  sinceId: 0,          // only article IDs above this (IDs rise with time). Use the maxId the last
                       // run logged in the source log. Feed times are unreliable (see argaam.md).
  sinceDate: null,     // optional, "dd/mm/yyyy": list items dated on or after this day
  listPages: 1,        // main-news list pages to scan (about 50 items a page, 1.5 days)
  disclosures: true,   // also scan disclosures list page 1
  text: [],            // article IDs whose body text to return (feed text when the feed has it)
  textMax: 1500,       // characters of body text per article
};
window.argaamNews = async (cfg) => {
  const CFG = {sinceId: 0, sinceDate: null, listPages: 1, disclosures: true, text: [], textMax: 1500, ...cfg};
  const clean = s => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
  const idOf = u => (u.match(/articledetail\/id\/(\d+)/) || [])[1] || null;
  const get = async p => new DOMParser().parseFromString(await (await fetch(p, {credentials: 'include'})).text(),
                                                          p.includes('/rss/') ? 'text/xml' : 'text/html');
  const feeds = {M: '/en/rss/ho-main-news?sectionid=1524', D: '/en/rss/ho-company-disclosures?sectionid=244'};
  const feed = [], feedText = {}, notes = [];
  for (const [f, p] of Object.entries(feeds)) {
    const items = [...(await get(p)).querySelectorAll('item')];
    if (!items.length) notes.push(`feed ${f} returned no items`);
    for (const it of items) {
      const u = it.querySelector('link')?.textContent || '', id = idOf(u);
      const desc = it.querySelector('description')?.textContent || '';
      const row = {id: id || u.replace(/^https?:\/\/[^/]+/, ''), f, h: clean(it.querySelector('title')?.textContent),
                   t: clean(it.querySelector('pubDate')?.textContent).replace(/ GMT$/, ''), n: desc.length};
      if (id && +id <= CFG.sinceId) continue;
      if (!feed.some(r => r.id === row.id)) feed.push(row);
      if (id && desc.length > 50) feedText[id] = clean(new DOMParser().parseFromString(desc, 'text/html').body.textContent);
    }
  }
  const inFeed = new Set(feed.map(r => r.id));
  const day = s => { const [d, m, y] = s.split('/'); return `${y}${m.padStart(2, '0')}${d.padStart(2, '0')}`; };
  const lists = [];
  for (let n = 1; n <= CFG.listPages; n++) lists.push(['L', `/en/article/sectionarticles/sectionid/138/marketid/3/pageno/${n}/main-news-saudi`]);
  if (CFG.disclosures) lists.push(['LD', '/en/article/sectionarticles/sectionid/244/marketid/3/pageno/1/disclosures']);
  const listOnly = [], seen = new Set(), flags = {}, onList = new Set();
  let scanned = 0;
  for (const [f, p] of lists) {
    const rows = [...(await get(p)).querySelectorAll('div.article-tbl')];
    if (!rows.length) notes.push(`list ${p} returned no rows: layout changed?`);
    for (const r of rows) {
      const a = r.querySelector('.title a[href*="articledetail/id/"]'); if (!a) continue;
      const id = idOf(a.getAttribute('href')); scanned++; if (id) onList.add(id);
      const d = clean(r.querySelector('.date')?.textContent);
      const lock = r.classList.contains('locked-article') || !!r.querySelector('.lock, .fa-lock, [class*="lock"]');
      const chart = !!r.querySelector('.fr-chart-link, [class*="chart"]');
      if (id && (lock || chart)) flags[id] = {lock, chart};
      if (!id || seen.has(id) || inFeed.has(id) || +id <= CFG.sinceId) continue;
      if (CFG.sinceDate && day(d) < day(CFG.sinceDate)) continue;
      seen.add(id);
      listOnly.push({id, f, d, h: clean(a.textContent), src: clean(r.querySelector('.source')?.textContent),
                     lock, chart});
    }
  }
  // Skip lock (paywalled, "Argaam Special") and bar-chart (financial report) rows everywhere,
  // including feed items: the RSS feeds carry no icon, so flags come from the list pages
  // (user, 7 Nov 2026). A feed item not seen on any scanned list page is marked unchecked:
  // true (raise listPages, or check its icon by hand).
  const skipped = [];
  const keep = r => { const fl = flags[r.id]; if (fl) { skipped.push({id: r.id, h: r.h, ...fl}); return false; } return true; };
  const feedK = feed.filter(keep).map(r => onList.has(r.id) ? r : {...r, unchecked: true}), listK = listOnly.filter(keep);
  const text = {};
  for (const id of CFG.text.map(String)) {
    if (feedText[id]) { text[id] = {via: 'feed', body: feedText[id].slice(0, CFG.textMax)}; continue; }
    const html = await (await fetch(`/en/article/articledetail/id/${id}`, {credentials: 'include'})).text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const pub = (html.match(/(\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M)/) || [])[1] || null;
    text[id] = {via: 'page', h: clean(doc.querySelector('h1')?.textContent), pub,
                body: clean(doc.querySelector('div.ck-editor.article')?.textContent).slice(0, CFG.textMax)};
    if (!text[id].body) notes.push(`article ${id}: no body found (layout changed or locked)`);
  }
  const ids = [...feed, ...listOnly].map(r => +r.id).filter(Number.isFinite);
  return {ran: new Date().toISOString(), maxId: ids.length ? Math.max(...ids) : CFG.sinceId,
          counts: {feed: feedK.length, listScanned: scanned, listOnly: listK.length, skipped: skipped.length},
          feed: feedK, listOnly: listK, skipped, text, notes};
};
await argaamNews(CFG)
