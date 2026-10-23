# Pitch feedback#2: working rules

## Every new or changed feature: check PC, laptop and iPad

Before calling a feature done, check it on all three. Say in the report which checks ran and which could not.

**Sizes** (CSS px; iPads in portrait and landscape):

| Device | Size |
| --- | --- |
| PC | 1920×1080 |
| Laptop | 1440×900, 1366×768, 1280×800 |
| iPad Air 4/5, iPad Air 11″ | 820×1180 |
| iPad Pro 11″ (M4) | 834×1210 |
| iPad Air 13″, iPad Pro 12.9″ | 1024×1366 |
| iPad Pro 13″ (M4) | 1032×1376 |

**Layout:** run `tools/device-audit.js` (how-to at its top). Every page must have no sideways scroll, nothing past the
right edge outside a scrolling box, and no squeezed labels. Watch the global `select` / `input` width of 100%
(`public/css/app.css`): in a flex `.row` it squeezes the label beside it; `.row > select` already sizes to its options.

**Browsers:** the oldest supported is Safari 15.4 (iPadOS 15.4). `index.html` shows "Please update your browser" below it.
- Any web API or CSS newer than Safari 15.4 needs a fallback.
  - Check caniuse.com for the Safari version.
  - Examples: `Promise.withResolvers` (17.4), `color-mix` (16.2), `Array.prototype.toSorted` (16), `@container` (16).
- Libraries loaded from a CDN count too. pdf.js 4.x needs Safari 17.4, so older Safari gets its legacy build (`pdf-markup.js`).

**Touch** (iPad has no mouse, no hover, often no keyboard):
- Nothing only on hover or in a `title` tooltip. Give it a tap too (see `bindTimeline`, dashboard `hover()`).
- Nothing only on `mouseup` / `mousemove` / `dblclick`. Use pointer events, or a button.
- Every keyboard shortcut needs a button as well.
- A canvas or editor area that takes drags sets `touch-action: none` only while a drawing tool is on, so a finger can still scroll.
- File pickers, downloads and new tabs must start from a tap (Safari blocks them otherwise).

## Other notes

- Bump `CODE_VERSION` in `apps-script/Code.gs` whenever an answer's shape changes (cached answers are shared by all deployments).
- Local test runs: on localhost, `localStorage.scriptUrl` points the page at a test deployment (`public/js/config.js`).
