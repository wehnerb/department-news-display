import { fetchWithTimeout } from './shared/fetch-helpers.js';
import { escapeHtml, sanitizeParam } from './shared/html.js';
import { getAccessToken } from './shared/google-auth.js';
import { DARK_BG_COLOR, FONT_STACK, ACCENT_COLOR, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_TERTIARY, TEXT_SUPPORTING, BORDER_SUBTLE, BORDER_STRONG, CARD_BASE, CARD_ELEVATED, CARD_HEADER, CARD_RECESSED } from './shared/colors.js';
import { LAYOUTS } from './shared/layouts.js';

// =============================================================================
// department-news-display — Cloudflare Worker
// =============================================================================
// Renders a card-based news feed for fire station display screens.
// Content is sourced from a Google Sheet with one tab per station feed.
//
// Google Sheet tabs:
//   "Department News" — department-wide news (?station=dept)
//   "FS#1"–"FS#8"     — station-specific news (?station=1–8)
//
// Layout variants (controlled via ?layout= URL parameter):
//   wide  — 1735×720  (default)
//   split — 852×720
//   tri   — 558×720
//   full  — 1920×1075
//
// Scroll behavior: if card content overflows the viewport, a CSS keyframe
// animation scrolls content upward once per page load. ResizeObserver +
// document.fonts.ready ensures measurement occurs after the Pi hardware
// finishes font rendering. No scroll occurs when content fits on screen.
//
// Secrets required (set in Cloudflare dashboard):
//   GOOGLE_SERVICE_ACCOUNT_EMAIL — service account email
//   GOOGLE_PRIVATE_KEY           — service account private key
//   GOOGLE_SHEET_ID              — ID from the Google Sheet URL
//
// =============================================================================
// CONFIGURATION
// =============================================================================

/** Total seconds the news slide is shown on the display.*/
const DISPLAY_DURATION_SECONDS = 20;

/** Items are highlighted as "new" if posted within these days.*/
const NEW_ITEM_THRESHOLD_DAYS = 3;

/** Seconds to pause at the top and bottom of the scroll.*/
const SCROLL_PAUSE_SECONDS = 5;

/** Minimum allowed scroll speed.*/
const MIN_SCROLL_SPEED_PX_PER_SEC = 20;

/** Maximum allowed scroll speed.*/
const MAX_SCROLL_SPEED_PX_PER_SEC = 75;

/** Cache controls.  Change cache version to immediately force displays to update with the current information. Change cache time to set hope long (in seconds) the news is cached for.*/
const CACHE_SECONDS = 300;
const CACHE_VERSION = 2;

/** Set text sizes for each part of the card. Increase rem multiplier to increase size*/
const FONT_SIZE_TITLE = '2.6rem'; /** Title size */
const FONT_SIZE_META = '0.8rem'; /** Meta data size */
const FONT_SIZE_BODY = '1.4rem'; /** Body text size */
const CARD_BODY_LINE_HEIGHT = 1.25; /** Line spacing for body text */
const CARD_META_LINE_HEIGHT = 1.6; /** Line spacing for meta data */

/** Colors for new and  older news cards. Alternates between A and B colors if there is more than 1 news item.**/
const COLOR_REGULAR_A = 'rgba(255,255,255,0.06)';
const COLOR_REGULAR_B = 'rgba(255,255,255,0.12)';
const COLOR_NEW_A = 'rgba(210,210,210,0.20)';
const COLOR_NEW_B = 'rgba(210,210,210,0.30)';

/** Old news items will be deleted this many days after the expiration date. If this is set to -1, old items will not be deleted.*/
const DELETE_EXPIRED_AFTER_DAYS = 14;

/** Column positions on the spreadsheet. These should not be changed unless the layout of the spreadsheet changes.*/
const COL_TITLE      = 0;
const COL_TEXT       = 1;
const COL_POSTED     = 2;
const COL_EXPIRATION = 3;
const COL_POSTED_BY  = 4;
const COL_RECURRENCE = 5;
const COL_STOP_AFTER = 6;

// =============================================================================
// CONSTANTS
// =============================================================================

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

// =============================================================================
// FETCH HANDLER
// =============================================================================

export default {
  async fetch(request, env) {
     if (request.method !== 'GET' && request.method !== 'HEAD') {
      return new Response('Method not allowed', {
        status: 405,
        headers: { 'Allow': 'GET, HEAD' },
      });
    }

    const url = new URL(request.url);
    if (url.pathname === '/healthz') {
      const method = request.method.toUpperCase();
      var healthStatus = 'healthy';
      const details    = [];

      // Verify that all required secrets are configured.
      const hasSheetId = !!(env.GOOGLE_SHEET_ID);
      const hasEmail   = !!(env.GOOGLE_SERVICE_ACCOUNT_EMAIL);
      const hasKey     = !!(env.GOOGLE_PRIVATE_KEY);

      if (!hasSheetId) { healthStatus = 'degraded'; details.push('GOOGLE_SHEET_ID: not configured'); }
      if (!hasEmail)   { healthStatus = 'degraded'; details.push('GOOGLE_SERVICE_ACCOUNT_EMAIL: not configured'); }
      if (!hasKey)     { healthStatus = 'degraded'; details.push('GOOGLE_PRIVATE_KEY: not configured'); }

      // Probe Google API reachability — a 400 response to a placeholder POST
      // confirms the endpoint is reachable without requiring valid credentials.
      let apiDetail = '';
      try {
        const probeRes = await fetchWithTimeout(
          'https://oauth2.googleapis.com/token',
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: 'grant_type=placeholder',
          },
          5000
        );
        // A 400 response is expected for a placeholder grant, confirming reachability
        if (probeRes.status === 400) {
          apiDetail = 'google-apis: reachable';
        } else {
          healthStatus = 'degraded';
          apiDetail = 'google-apis: unexpected status ' + probeRes.status;
        }
      } catch (e) {
        healthStatus = 'degraded';
        apiDetail = 'google-apis: unreachable (' + (e && e.message ? e.message : String(e)) + ')';
      }

      details.push(apiDetail);

      // Construct and return the health status response.
      const healthBody =
        'status: ' + healthStatus + '\n' +
        'worker: department-news-display\n' +
        details.join('\n') + '\n';

      return new Response(
        method === 'HEAD' ? null : healthBody,
        {
          status: healthStatus === 'healthy' ? 200 : 503,
          headers: {
            'Content-Type':  'text/plain; charset=UTF-8',
            'Cache-Control': 'no-store',
          },
        }
      );
    }
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

// =============================================================================
// DATA FETCHING
// =============================================================================

async function fetchSheetRows(env, token, tabName) {
  const cacheKey = new Request('https://cache.internal/news-display/v' + CACHE_VERSION + '/' + env.GOOGLE_SHEET_ID + '/' + encodeURIComponent(tabName));
  const cache = caches.default;
  const cached = await cache.match(cacheKey);
  if (cached) return await cached.json();

  const range  = encodeURIComponent(tabName + '!A:G');
  const apiUrl =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '/values/' + range + '?valueRenderOption=FORMATTED_VALUE';

  const res = await fetchWithTimeout(apiUrl, { headers: { 'Authorization': 'Bearer ' + token } }, 8000);
  if (!res.ok) throw new Error('Sheets API error');

  const data = await res.json();
  const dataRows = (data.values || []).slice(1);

  await cache.put(cacheKey, new Response(JSON.stringify(dataRows), { headers: { 'Cache-Control': 'max-age=' + CACHE_SECONDS } }));
  return dataRows;
}

// =============================================================================
// DATA PROCESSING
// =============================================================================

// Parses a date/time string from the Google Sheet into a JavaScript Date
// object, treating the value as America/Chicago wall-clock time.
//
// Accepted format: M/D/YY H:MM AM/PM or M/D/YYYY H:MM AM/PM
//   e.g. "5/4/26 7:30 AM" or "5/4/2026 7:30 AM"
//
// Strategy: parse the components, build a candidate UTC epoch treating
// the values as UTC, then use Intl.DateTimeFormat to measure the actual
// Central offset at that time and correct. Two iterations converge on
// the exact UTC epoch for the given Central wall-clock time, handling
// DST transitions correctly.
function parseSheetDateTime(raw) {
  if (!raw) return null;
  var s = String(raw).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!m) {
    var fallback = new Date(s);
    return isNaN(fallback.getTime()) ? null : fallback;
  }
  var mo  = parseInt(m[1], 10) - 1;  // 0-indexed month
  var dy  = parseInt(m[2], 10);
  var yr  = parseInt(m[3], 10);
  // Normalize 2-digit years: 00-49 → 2000-2049, 50-99 → 1950-1999
  if (yr < 100) { yr = yr < 50 ? 2000 + yr : 1900 + yr; }
  var hr  = parseInt(m[4], 10);
  var min = parseInt(m[5], 10);
  var ap  = (m[6] || '').toUpperCase();
  if (ap === 'PM' && hr !== 12) { hr += 12; }
  if (ap === 'AM' && hr === 12) { hr = 0; }

  var fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Chicago',
    year: 'numeric', month: 'numeric', day: 'numeric',
    hour: 'numeric', minute: 'numeric', hour12: false,
  });

  // Start with a candidate UTC epoch as if the components were UTC,
  // then iterate to correct for the Central offset.
  var candidate = Date.UTC(yr, mo, dy, hr, min, 0);
  for (var iter = 0; iter < 2; iter++) {
    var parts = {};
    fmt.formatToParts(new Date(candidate)).forEach(function(p) {
      parts[p.type] = Number(p.value);
    });
    var cHour      = (parts.hour === 24) ? 0 : (parts.hour || 0);
    var centralUtc = Date.UTC(parts.year, parts.month - 1, parts.day, cHour, parts.minute || 0, 0);
    var targetUtc  = Date.UTC(yr, mo, dy, hr, min, 0);
    candidate     += (targetUtc - centralUtc);
  }
  var result = new Date(candidate);
  return isNaN(result.getTime()) ? null : result;
}

// Processes raw Google Sheet rows into active, displayable news item objects.
// Handles both one-time and recurring items. Recurring items are expanded
// to their current active occurrence based on the recurrence interval.
// Items outside their active window (not yet posted or already expired)
// are silently skipped. Results are sorted: new items first (by posted
// date descending), then regular items (by posted date descending).
function processRows(rows, now) {
  var items = [];
  var nowMs  = now.getTime();

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (!row[COL_TITLE] || !row[COL_TEXT] || !row[COL_POSTED]) continue;

    var originalPosted  = parseSheetDateTime(row[COL_POSTED]);
    var originalExpires = parseSheetDateTime(row[COL_EXPIRATION]);
    if (!originalPosted || !originalExpires) continue;

    var recurDays = row[COL_RECURRENCE] ? parseInt(row[COL_RECURRENCE], 10) : 0;
    var stopAfter = row[COL_STOP_AFTER]
      ? parseSheetDateTime(row[COL_STOP_AFTER] + ' 11:59 PM')
      : null;

    if (recurDays > 0 && !isNaN(recurDays)) {
      // ── Recurring item ───────────────────────────────────────────────
      // Advance the occurrence window forward by the recurrence interval
      // until the current occurrence started at or before now.
      var intervalMs  = recurDays * 24 * 60 * 60 * 1000;
      var activePosted  = new Date(originalPosted.getTime());
      var activeExpires = new Date(originalExpires.getTime());

      while (activePosted.getTime() + intervalMs <= nowMs) {
        var nextPosted  = new Date(activePosted.getTime()  + intervalMs);
        var nextExpires = new Date(activeExpires.getTime() + intervalMs);
        if (stopAfter && nextPosted > stopAfter) break;
        activePosted  = nextPosted;
        activeExpires = nextExpires;
      }

      if (stopAfter && activePosted > stopAfter) continue;
      if (nowMs < activePosted.getTime() || nowMs >= activeExpires.getTime()) continue;

      var ageDaysR = (nowMs - activePosted.getTime()) / (1000 * 60 * 60 * 24);
      items.push({
        title:         row[COL_TITLE].trim(),
        text:          row[COL_TEXT].trim(),
        postedBy:      (row[COL_POSTED_BY] || '').trim(),
        activePosted:  activePosted,
        activeExpires: activeExpires,
        isNew:         ageDaysR >= 0 && ageDaysR <= NEW_ITEM_THRESHOLD_DAYS,
      });

    } else {
      // ── Non-recurring item ───────────────────────────────────────────
      if (nowMs < originalPosted.getTime() || nowMs >= originalExpires.getTime()) continue;

      var ageDaysN = (nowMs - originalPosted.getTime()) / (1000 * 60 * 60 * 24);
      items.push({
        title:         row[COL_TITLE].trim(),
        text:          row[COL_TEXT].trim(),
        postedBy:      (row[COL_POSTED_BY] || '').trim(),
        activePosted:  originalPosted,
        activeExpires: originalExpires,
        isNew:         ageDaysN >= 0 && ageDaysN <= NEW_ITEM_THRESHOLD_DAYS,
      });
    }
  }

  // Sort: new items first (most recently posted first within each group),
  // then regular items (most recently posted first).
  return items.sort(function(a, b) {
    if (a.isNew !== b.isNew) return a.isNew ? -1 : 1;
    return b.activePosted - a.activePosted;
  });
}

function formatDateTime(date) {
  return date.toLocaleString('en-US', { timeZone: 'America/Chicago', month: 'numeric', day: 'numeric', year: '2-digit', hour: 'numeric', minute: '2-digit', hour12: true });
}

// =============================================================================
// HTML RENDERER
// =============================================================================

// Builds the self-contained HTML page for the news display.
// Renders all active news items as cards. If content overflows the
// visible area, a ResizeObserver measures the rendered height after
// fonts are loaded and injects a CSS @keyframes animation that scrolls
// content upward once, pausing at the top and bottom.
// Uses string concatenation throughout (no template literals) to avoid
// smart-quote esbuild errors when editing in the GitHub browser editor.
function renderHtml(items, layout, tabName, darkBg) {

  // ── Card HTML builder ──────────────────────────────────────────────
  // Alternates background colors between adjacent cards of the same type
  // (new vs regular) for visual separation. New items get a lighter
  // tinted background; regular items get the standard card surface.
  var newCount     = 0;
  var regularCount = 0;

  var cardsHtml = items.length === 0
    ? '<div class="no-news">No current news</div>'
    : items.map(function(item) {
        var bgColor = item.isNew
          ? (newCount++     % 2 === 0 ? COLOR_NEW_A     : COLOR_NEW_B)
          : (regularCount++ % 2 === 0 ? COLOR_REGULAR_A : COLOR_REGULAR_B);

        var cardStyle =
          'background:'    + bgColor      + ';' +
          'border:1px solid ' + BORDER_SUBTLE + ';' +
          'padding:0.8rem 1rem;' +
          'margin-bottom:0.65rem;' +
          'border-radius:6px;';

        // Red accent underline beneath the title on regular (non-new) cards.
        var titleDivider = item.isNew ? '' : '<div class="title-divider"></div>';

        // Preserve Alt+Enter line breaks entered in the sheet.
        var safeBody = escapeHtml(item.text).replace(/\n/g, '<br>');

        // Posted By is optional — omit the line entirely when blank.
        var postedByHtml = item.postedBy
          ? '<div><strong>Posted By:</strong> ' + escapeHtml(item.postedBy) + '</div>'
          : '';

        return (
          '<div class="news-card" style="' + cardStyle + '">' +
            '<div class="card-title">' +
              (item.isNew ? '<span class="new-badge">NEW</span>' : '') +
              escapeHtml(item.title) +
            '</div>' +
            titleDivider +
            '<div class="card-meta">' +
              '<div><strong>Posted:</strong> '  + formatDateTime(item.activePosted)  + '</div>' +
              '<div><strong>Expires:</strong> ' + formatDateTime(item.activeExpires) + '</div>' +
              postedByHtml +
            '</div>' +
            '<div class="card-body">' + safeBody + '</div>' +
          '</div>'
        );
      }).join('');

  // ── Scroll animation script ────────────────────────────────────────
  // Injected into the page as an inline <script>. Uses ResizeObserver to
  // measure the rendered content height only after document.fonts.ready
  // fires — this ensures the Pi hardware has finished font rendering before
  // any height measurement is taken, making the scroll trigger reliable
  // under CPU load. Generates a unique @keyframes animation name to avoid
  // conflicts on repeated page loads. Disconnects the observer after the
  // animation is applied so it runs at most once per page load.
  const scrollScript =
    '(function () {' +
    '  var DURATION  = ' + DISPLAY_DURATION_SECONDS   + ';' +
    '  var PAUSE     = ' + SCROLL_PAUSE_SECONDS        + ';' +
    '  var MIN_SPEED = ' + MIN_SCROLL_SPEED_PX_PER_SEC + ';' +
    '  var MAX_SPEED = ' + MAX_SCROLL_SPEED_PX_PER_SEC + ';' +

    '  function startLogic() {' +
    '    var outer = document.getElementById("scroller");' +
    '    var inner = document.getElementById("scroll-inner");' +
    '    if (!outer || !inner || inner.dataset.noScroll) return;' +
    '    var observer = new ResizeObserver(function(entries) {' +
    '      for (var i = 0; i < entries.length; i++) {' +
    '        var totalH = entries[i].contentRect.height;' +
    '        var viewH  = outer.clientHeight;' +
    '        if (totalH <= viewH + 5) return;' +
    '        var overflow      = totalH - viewH;' +
    '        var availableTime = Math.max(1, DURATION - (2 * PAUSE));' +
    '        var speed         = Math.min(MAX_SPEED, Math.max(MIN_SPEED, overflow / availableTime));' +
    '        var scrollTime    = overflow / speed;' +
    '        var totalTime     = PAUSE + scrollTime + PAUSE;' +
    '        var pTop          = (PAUSE / totalTime) * 100;' +
    '        var pBottom       = 100 - pTop;' +
    '        var animName      = "scroll-" + Date.now();' +
    '        var style         = document.createElement("style");' +
    '        style.textContent = "@keyframes " + animName + " { 0%, " + pTop.toFixed(2) + "% { transform: translateY(0); } " + pBottom.toFixed(2) + "%, 100% { transform: translateY(-" + overflow + "px); } }";' +
    '        document.head.appendChild(style);' +
    '        inner.style.animation = animName + " " + totalTime + "s linear forwards";' +
    '        observer.disconnect();' +
    '      }' +
    '    });' +
    '    observer.observe(inner);' +
    '  }' +

    '  if (document.fonts) {' +
    '    document.fonts.ready.then(startLogic);' +
    '  } else {' +
    '    window.addEventListener("load", startLogic);' +
    '  }' +
    '}());';

  // ── CSS ────────────────────────────────────────────────────────────
  // html: overflow hidden prevents the root element from scrolling.
  // body: height 100vh + overflow hidden locks to viewport; box-sizing
  //   is applied via the * reset so padding fits within 100vh.
  // #scroller: height 100% fills the body content area.
  // #scroll-inner: will-change: transform enables GPU layer promotion
  //   for the CSS keyframe translateY animation.
  const css =
    '*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }' +
    'html { overflow: hidden; }' +
    'body {' +
    '  background: ' + (darkBg ? DARK_BG_COLOR : 'transparent') + ';' +
    '  color: #f0f0f0;' +
    '  font-family: ' + FONT_STACK + ';' +
    '  font-size: '   + FONT_SIZE_BODY + ';' +
    '  padding: 0.75rem;' +
    '  height: 100vh;' +
    '  overflow: hidden;' +
    '}' +
    '#scroller { height: 100%; overflow: hidden; }' +
    '#scroll-inner { will-change: transform; }' +
    '.no-news {' +
    '  text-align: center;' +
    '  font-size: ' + FONT_SIZE_TITLE + ';' +
    '  color: rgba(255,255,255,0.55);' +
    '  margin-top: 3rem;' +
    '}' +
    '.news-card { border-radius: 6px; }' +
    '.card-title {' +
    '  font-size: '    + FONT_SIZE_TITLE + ';' +
    '  font-weight: bold;' +
    '  color: #ffffff;' +
    '  margin-bottom: 0.3rem;' +
    '  display: flex;' +
    '  align-items: center;' +
    '  gap: 0.5rem;' +
    '  flex-wrap: wrap;' +
    '}' +
    '.new-badge {' +
    '  background: rgba(255,215,80,0.88);' +
    '  color: #1a1a1a;' +
    '  font-size: 0.62rem;' +
    '  font-weight: bold;' +
    '  letter-spacing: 0.06em;' +
    '  padding: 0.15rem 0.4rem;' +
    '  border-radius: 3px;' +
    '  flex-shrink: 0;' +
    '}' +
    '.title-divider {' +
    '  width: 60%;' +
    '  height: 2px;' +
    '  background: ' + ACCENT_COLOR + ';' +
    '  opacity: 0.85;' +
    '  border-radius: 1px;' +
    '  margin-bottom: 0.45rem;' +
    '}' +
    '.card-meta {' +
    '  font-size: '    + FONT_SIZE_META + ';' +
    '  color: '        + TEXT_SECONDARY + ';' +
    '  line-height: '  + CARD_META_LINE_HEIGHT + ';' +
    '  margin-bottom: 0.65rem;' +
    '}' +
    '.card-body {' +
    '  font-size: '   + FONT_SIZE_BODY + ';' +
    '  color: '       + TEXT_PRIMARY   + ';' +
    '  line-height: ' + CARD_BODY_LINE_HEIGHT + ';' +
    '}';

  // ── HTML document ─────────────────────────────────────────────────
  // Returns a complete self-contained HTML page. The meta-refresh is
  // intentionally omitted — the display hardware controls cycling.
  // Cache-Control: no-store is set on the Response in the fetch handler.
  return (
    '<!DOCTYPE html>' +
    '<html lang="en">' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>' + escapeHtml(tabName) + ' News</title>' +
    '<style>' + css + '</style>' +
    '</head>' +
    '<body>' +
    '<div id="scroller">' +
    '<div id="scroll-inner"' + (items.length === 0 ? ' data-no-scroll="1"' : '') + '>' +
    cardsHtml +
    '</div>' +
    '</div>' +
    '<script>' + scrollScript + '</script>' +
    '</body>' +
    '</html>'
  );
}

// ================================================================
// SCHEDULED CLEANUP — Cron-triggered expired row deletion
// ================================================================

// Entry point for the scheduled cron job. Iterates all configured
// station tabs and deletes non-recurring rows that expired more than
// DELETE_EXPIRED_AFTER_DAYS days ago. Skips recurring rows — those
// should be managed manually. Disabled when DELETE_EXPIRED_AFTER_DAYS
// is set to -1.
async function runCleanup(env) {
  if (DELETE_EXPIRED_AFTER_DAYS < 0) {
    console.log('Automatic row deletion is disabled (DELETE_EXPIRED_AFTER_DAYS = -1). Skipping.');
    return;
  }

  const now         = new Date();
  const thresholdMs = DELETE_EXPIRED_AFTER_DAYS * 24 * 60 * 60 * 1000;

  let token;
  try {
    token = await getAccessToken(
      env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      env.GOOGLE_PRIVATE_KEY,
      'https://www.googleapis.com/auth/spreadsheets'
    );
  } catch (e) {
    console.error('runCleanup: failed to obtain access token:', e && e.message ? e.message : e);
    return;
  }

  for (const [stationKey, tabName] of Object.entries(STATION_TAB_MAP)) {
    try {
      await cleanupTab(env, token, tabName, now, thresholdMs);
    } catch (e) {
      console.error('runCleanup: error on tab "' + tabName + '":', e && e.message ? e.message : e);
    }
  }
}

// Scans a single sheet tab for expired non-recurring rows and deletes
// them using the Sheets API batchUpdate deleteDimension operation.
// Rows are deleted in reverse order so that earlier row indices remain
// valid as each deletion shifts subsequent rows upward.
async function cleanupTab(env, token, tabName, now, thresholdMs) {
  // Fetch all rows from this tab
  const range  = encodeURIComponent(tabName + '!A:G');
  const apiUrl =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '/values/' + range + '?valueRenderOption=FORMATTED_VALUE';

  const res = await fetchWithTimeout(
    apiUrl,
    { headers: { 'Authorization': 'Bearer ' + token } },
    8000
  );
  if (!res.ok) {
    console.error('cleanupTab: Sheets API error ' + res.status + ' for tab "' + tabName + '"');
    return;
  }

  const data = await res.json();
  const rows = (data.values || []).slice(1); // skip header row

  // Identify rows to delete: expired non-recurring items past the threshold
  const rowsToDelete = [];
  for (var i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row[COL_TITLE] || !row[COL_TEXT] || !row[COL_POSTED]) continue;
    if (row[COL_RECURRENCE]) continue; // never auto-delete recurring items
    const expires = parseSheetDateTime(row[COL_EXPIRATION]);
    if (!expires) continue;
    if (now.getTime() - expires.getTime() >= thresholdMs) {
      rowsToDelete.push(i + 1); // +1 to account for the header row (1-indexed)
    }
  }

  if (rowsToDelete.length === 0) return;

  // Fetch sheet metadata to get the numeric sheetId (tab GID)
  // required by the batchUpdate deleteDimension operation.
  const metaUrl =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '?fields=sheets.properties';
  const metaRes = await fetchWithTimeout(
    metaUrl,
    { headers: { 'Authorization': 'Bearer ' + token } },
    8000
  );
  if (!metaRes.ok) {
    console.error('cleanupTab: failed to fetch sheet metadata for tab "' + tabName + '"');
    return;
  }
  const meta        = await metaRes.json();
  const sheetEntry  = (meta.sheets || []).find(function(s) {
    return s.properties && s.properties.title === tabName;
  });
  if (!sheetEntry) {
    console.error('cleanupTab: tab "' + tabName + '" not found in sheet metadata');
    return;
  }
  const sheetId = sheetEntry.properties.sheetId;

  // Build deleteDimension requests in reverse row order so that deleting
  // a row does not invalidate the indices of rows yet to be deleted.
  const requests = rowsToDelete.slice().reverse().map(function(rowIdx) {
    return {
      deleteDimension: {
        range: {
          sheetId:    sheetId,
          dimension:  'ROWS',
          startIndex: rowIdx,
          endIndex:   rowIdx + 1,
        },
      },
    };
  });

  const batchUrl =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) + ':batchUpdate';
  const batchRes = await fetchWithTimeout(
    batchUrl,
    {
      method:  'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json',
      },
      body: JSON.stringify({ requests }),
    },
    8000
  );
  if (!batchRes.ok) {
    console.error('cleanupTab: batchUpdate failed ' + batchRes.status + ' for tab "' + tabName + '"');
    return;
  }
  console.log('cleanupTab: deleted ' + rowsToDelete.length + ' expired rows from "' + tabName + '"');
}
