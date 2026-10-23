// Device layout audit (v2.13): runs every page of the site at PC, laptop and iPad sizes and lists what breaks.
// How to run: serve public/ on localhost (python -m http.server 8765), sign in on http://localhost:8765, then open
// http://localhost:8765/markup-test.html, paste this file into the DevTools console and wait (about 15 s per page).
// The site loads in a same-origin frame at each size, so media queries match the device and the sign-in carries over.
// Results: console.table, and window.deviceAudit for a closer look. Checks per page:
//   overflow  the page scrolls sideways (should be 0 or less)
//   loose     elements past the right edge that are not inside a scrolling box (should be empty)
//   squeezed  labels or cells narrower than 40 px (a select or input squeezing its label)
//   error     the page showed an error instead of its content
(async () => {
  const SIZES = [
    ['PC 1920', 1920, 1080], ['Laptop 1440', 1440, 900], ['Laptop 1366', 1366, 768], ['Laptop 1280', 1280, 800],
    ['iPad Air 4/5, Air 11″', 820, 1180], ['iPad Pro 11″ M4', 834, 1210], ['iPad Air 13″, Pro 12.9″', 1024, 1366], ['iPad Pro 13″ M4', 1032, 1376],
  ];
  const WAIT = 5000; // per page; raise it if pages are still loading when measured
  const frame = document.createElement('iframe');
  frame.style.cssText = 'border:0;display:block;background:#fff';
  document.body.replaceChildren(frame);
  frame.src = '/#/';
  const sleep = (ms) => new Promise((r) => { setTimeout(r, ms); });
  await sleep(WAIT * 2);
  // An organization with projects, for the three organization tabs.
  const org = [...frame.contentDocument.querySelectorAll('tr[data-ref]')].find((r) => /\d/.test(r.cells[3]?.textContent || ''))?.dataset.ref;
  const routes = ['#/dashboard', '#/', '#/mine', '#/admin', '#/control', ...(org ? ['prepitch', 'pitching', 'resubmit'].map((t) => `#/org/${org}/${t}`) : [])];
  const tag = (e) => `${e.tagName.toLowerCase()}${e.id ? `#${e.id}` : ''}${typeof e.className === 'string' && e.className.trim() ? `.${e.className.trim().split(/\s+/)[0]}` : ''}`;
  const out = [];
  for (const [device, w, h] of SIZES) {
    const turns = w < h ? [[w, h, 'portrait'], [h, w, 'landscape']] : [[w, h, '']];
    for (const [W, H, ori] of turns) {
      frame.width = W; frame.height = H;
      for (const route of routes) {
        frame.contentWindow.location.hash = route;
        await sleep(WAIT);
        const d = frame.contentDocument; const vw = frame.contentWindow.innerWidth;
        const scrolls = (e) => { for (let p = e.parentElement; p && p !== d.body; p = p.parentElement) if (/(auto|scroll|hidden|clip)/.test(getComputedStyle(p).overflowX)) return true; return false; };
        const loose = [...d.querySelectorAll('body *')].filter((e) => { const b = e.getBoundingClientRect(); return b.width && b.right > vw + 1 && !scrolls(e); });
        const squeezed = [...d.querySelectorAll('.row > .grow, .row > span, label > span, th, td')]
          .filter((e) => { const b = e.getBoundingClientRect(); return e.textContent.trim().length > 3 && b.width > 0 && b.width < 40; });
        out.push({
          device, orientation: ori, route, overflow: d.documentElement.scrollWidth - vw,
          loose: [...new Set(loose.map(tag))].slice(0, 5).join(' '), squeezed: squeezed.slice(0, 3).map((e) => e.textContent.trim().slice(0, 20)).join(' | '),
          error: d.querySelector('.page .empty p')?.textContent || '',
        });
      }
    }
  }
  window.deviceAudit = out;
  const bad = out.filter((x) => x.overflow > 0 || x.loose || x.squeezed || x.error);
  console.log(bad.length ? `${bad.length} problems:` : `No problems on ${out.length} pages.`);
  if (bad.length) console.table(bad);
})();
