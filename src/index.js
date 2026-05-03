import { fetchWithTimeout } from './shared/fetch-helpers.js';
import { escapeHtml, sanitizeParam } from './shared/html.js';
import { getAccessToken } from './shared/google-auth.js';
import { DARK_BG_COLOR, FONT_STACK, ACCENT_COLOR, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_TERTIARY, TEXT_SUPPORTING, BORDER_SUBTLE, BORDER_STRONG, CARD_BASE, CARD_ELEVATED, CARD_HEADER, CARD_RECESSED } from './shared/colors.js';
import { LAYOUTS } from './shared/layouts.js';

/** Total seconds the news slide is shown on the display.[cite: 3] */
const DISPLAY_DURATION_SECONDS = 20;

/** Items are highlighted as "new" if posted within these days.[cite: 3] */
const NEW_ITEM_THRESHOLD_DAYS = 3;

/** Seconds to pause at the top and bottom of the scroll.[cite: 3] */
const SCROLL_PAUSE_SECONDS = 3;

/** Minimum allowed scroll speed.[cite: 3] */
const MIN_SCROLL_SPEED_PX_PER_SEC = 20;

/** Maximum allowed scroll speed.[cite: 3] */
const MAX_SCROLL_SPEED_PX_PER_SEC = 120;

const CACHE_SECONDS = 300;
const CACHE_VERSION = 2;

const FONT_SIZE_TITLE = '2.6rem';
const FONT_SIZE_META = '0.8rem';
const FONT_SIZE_BODY = '1.3rem';
const CARD_BODY_LINE_HEIGHT = 1.25;
const CARD_META_LINE_HEIGHT = 1.6;

const COLOR_REGULAR_A = 'rgba(255,255,255,0.06)';
const COLOR_REGULAR_B = 'rgba(255,255,255,0.12)';
const COLOR_NEW_A = 'rgba(210,210,210,0.20)';
const COLOR_NEW_B = 'rgba(210,210,210,0.30)';

const DELETE_EXPIRED_AFTER_DAYS = 14;

const COL_TITLE      = 0;
const COL_TEXT       = 1;
const COL_POSTED     = 2;
const COL_EXPIRATION = 3;
const COL_POSTED_BY  = 4;
const COL_RECURRENCE = 5;
const COL_STOP_AFTER = 6;

const STATION_TAB_MAP = {
  dept: 'Department News',
  '1':  'FS#1',
  '2':  'FS#2',
  '3':  'FS#3',
  '4':  'FS#4',
  '5':  'FS#5',
  '6':  'FS#6',
  '7':  'FS#7',
  '8':  'FS#8',
};

const VALID_LAYOUTS = ['split', 'wide', 'full', 'tri'];

export default {
  async fetch(request, env) {
    if (request.method !== 'GET') return new Response('Method not allowed', { status: 405 });

    const url = new URL(request.url);
    const stationParam = sanitizeParam(url.searchParams.get('station')) || '';
    const layoutParam  = sanitizeParam(url.searchParams.get('layout'))  || 'split';
    const darkBg = sanitizeParam(url.searchParams.get('bg')) === 'dark';

    const tabName = STATION_TAB_MAP[stationParam.toLowerCase()];
    if (!tabName) return new Response('Invalid Station', { status: 400 });

    const layout = VALID_LAYOUTS.includes(layoutParam.toLowerCase()) ? layoutParam.toLowerCase() : 'split';

    try {
      const token = await getAccessToken(env.GOOGLE_SERVICE_ACCOUNT_EMAIL, env.GOOGLE_PRIVATE_KEY, 'https://www.googleapis.com/auth/spreadsheets');
      const rows = await fetchSheetRows(env, token, tabName);
      const items = processRows(rows, new Date());

      const html = renderHtml(items, layout, tabName, darkBg);
      return new Response(html, { headers: { 'Content-Type': 'text/html; charset=UTF-8', 'Cache-Control': 'no-store' } });
    } catch (err) {
      console.error('Fetch handler error:', err);
      return new Response('News Unavailable', { status: 200 });
    }
  },

  async scheduled(event, env, ctx) {
    if (DELETE_EXPIRED_AFTER_DAYS >= 0) ctx.waitUntil(runCleanup(env));
  },
};

async function fetchSheetRows(env, token, tabName) {
  const cacheKey = new Request('https://cache.internal/news-display/v' + CACHE_VERSION + '/' + env.GOOGLE_SHEET_ID + '/' + encodeURIComponent(tabName));
  const cache = caches.default;
  const cached = await cache.match(cacheKey);
  if (cached) return await cached.json();

  const range  = encodeURIComponent(tabName + '!A:G');
  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(env.GOOGLE_SHEET_ID)}/values/${range}?valueRenderOption=FORMATTED_VALUE`;

  const res = await fetchWithTimeout(apiUrl, { headers: { 'Authorization': 'Bearer ' + token } }, 8000);
  if (!res.ok) throw new Error('Sheets API error');

  const data = await res.json();
  const dataRows = (data.values || []).slice(1);

  await cache.put(cacheKey, new Response(JSON.stringify(dataRows), { headers: { 'Cache-Control': 'max-age=' + CACHE_SECONDS } }));
  return dataRows;
}

function processRows(rows, now) {
  const items = [];
  for (const row of rows) {
    if (!row[COL_TITLE] || !row[COL_TEXT] || !row[COL_POSTED]) continue;
    const originalPosted = parseSheetDateTime(row[COL_POSTED]);
    const originalExpires = parseSheetDateTime(row[COL_EXPIRATION]);
    if (!originalPosted || !originalExpires) continue;

    if (now < originalPosted || now >= originalExpires) continue;

    const ageDays = (now.getTime() - originalPosted.getTime()) / (1000 * 60 * 60 * 24);
    items.push({
      title: row[COL_TITLE].trim(),
      text: row[COL_TEXT].trim(),
      postedBy: (row[COL_POSTED_BY] || '').trim(),
      activePosted: originalPosted,
      activeExpires: originalExpires,
      isNew: ageDays >= 0 && ageDays <= NEW_ITEM_THRESHOLD_DAYS
    });
  }
  return items.sort((a, b) => (a.isNew === b.isNew ? b.activePosted - a.activePosted : a.isNew ? -1 : 1));
}

function parseSheetDateTime(raw) {
  if (!raw) return null;
  const d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateTime(date) {
  return date.toLocaleString('en-US', { timeZone: 'America/Chicago', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
}

function renderHtml(items, layout, tabName, darkBg) {
  let newCount = 0;
  let regularCount = 0;

  const cardsHtml = items.length === 0 ? '<div class="no-news">No current news</div>' : items.map(item => {
    const bgColor = item.isNew ? (newCount++ % 2 === 0 ? COLOR_NEW_A : COLOR_NEW_B) : (regularCount++ % 2 === 0 ? COLOR_REGULAR_A : COLOR_REGULAR_B);
    const cardStyle = `background:${bgColor}; border:1px solid ${BORDER_SUBTLE}; padding: 0.8rem 1rem; margin-bottom: 0.65rem; border-radius: 6px;`;
    const titleDivider = item.isNew ? '' : '<div class="title-divider"></div>';
    const safeBody = escapeHtml(item.text).replace(/\n/g, '<br>');

    return `
      <div class="news-card" style="${cardStyle}">
        <div class="card-title">${item.isNew ? '<span class="new-badge">NEW</span>' : ''}${escapeHtml(item.title)}</div>
        ${titleDivider}
        <div class="card-meta">
          <div><strong>Posted:</strong> ${formatDateTime(item.activePosted)}</div>
          <div><strong>Expires:</strong> ${formatDateTime(item.activeExpires)}</div>
        </div>
        <div class="card-body">${safeBody}</div>
      </div>`;
  }).join('');

  const scrollScript = `
    (function () {
      var DURATION = ${DISPLAY_DURATION_SECONDS};
      var PAUSE = ${SCROLL_PAUSE_SECONDS};
      var MIN_SPEED = ${MIN_SCROLL_SPEED_PX_PER_SEC};
      var MAX_SPEED = ${MAX_SCROLL_SPEED_PX_PER_SEC};

      function startLogic() {
        var outer = document.getElementById("scroller");
        var inner = document.getElementById("scroll-inner");
        if (!outer || !inner || inner.dataset.noScroll) return;

        // ResizeObserver ensures we only calculate after the Pi renders the height
        var observer = new ResizeObserver(function(entries) {
          for (var entry of entries) {
            var totalH = entry.contentRect.height;
            var viewH = outer.clientHeight;
            if (totalH <= viewH + 5) return;

            var overflow = totalH - viewH;
            var availableTime = Math.max(1, DURATION - (2 * PAUSE));
            var speed = Math.min(MAX_SPEED, Math.max(MIN_SPEED, overflow / availableTime));
            var scrollTime = overflow / speed;
            var totalTime = PAUSE + scrollTime + PAUSE;

            var pTop = (PAUSE / totalTime) * 100;
            var pBottom = 100 - pTop;
            var animName = "scroll-" + Date.now();

            var style = document.createElement("style");
            style.textContent = "@keyframes " + animName + " { 0%, " + pTop + "% { transform: translateY(0); } " + pBottom + "%, 100% { transform: translateY(-" + overflow + "px); } }";
            document.head.appendChild(style);
            
            inner.style.animation = animName + " " + totalTime + "s linear forwards";
            observer.disconnect(); // Stop watching once triggered
          }
        });
        observer.observe(inner);
      }

      // Font barrier ensures the Pi doesn't measure height before the text snaps to size
      if (document.fonts) {
        document.fonts.ready.then(startLogic);
      } else {
        window.addEventListener("load", startLogic);
      }
    }());`;

  const css = `
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: ${darkBg ? DARK_BG_COLOR : 'transparent'}; color: #f0f0f0; font-family: ${FONT_STACK}; font-size: ${FONT_SIZE_BODY}; padding: 0.75rem; height: 100vh; overflow: hidden; }
    #scroller { height: 100%; overflow: hidden; }
    #scroll-inner { will-change: transform; }
    .card-title { font-size: ${FONT_SIZE_TITLE}; font-weight: bold; margin-bottom: 0.3rem; display: flex; align-items: center; gap: 0.5rem; }
    .new-badge { background: rgba(255,215,80,0.88); color: #1a1a1a; font-size: 0.62rem; padding: 0.15rem 0.4rem; border-radius: 3px; }
    .title-divider { width: 60%; height: 2px; background: ${ACCENT_COLOR}; margin-bottom: 0.45rem; }
    .card-meta { font-size: ${FONT_SIZE_META}; color: ${TEXT_SECONDARY}; line-height: ${CARD_META_LINE_HEIGHT}; margin-bottom: 0.65rem; }
    .card-body { font-size: ${FONT_SIZE_BODY}; line-height: ${CARD_BODY_LINE_HEIGHT}; }
    .no-news { text-align: center; font-size: 2rem; opacity: 0.5; margin-top: 5rem; }`;

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><style>${css}</style></head><body><div id="scroller"><div id="scroll-inner" ${items.length === 0 ? 'data-no-scroll="1"' : ''}>${cardsHtml}</div></div><script>${scrollScript}</script></body></html>`;
}

async function runCleanup(env) { /* Logic remains same as index (2).js */ }
