import { fetchWithTimeout } from './shared/fetch-helpers.js';
import { escapeHtml, sanitizeParam } from './shared/html.js';
import { getAccessToken } from './shared/google-auth.js';
import { DARK_BG_COLOR, FONT_STACK, ACCENT_COLOR, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_TERTIARY, TEXT_SUPPORTING, BORDER_SUBTLE, BORDER_STRONG, CARD_BASE, CARD_ELEVATED, CARD_HEADER, CARD_RECESSED } from './shared/colors.js';
import { LAYOUTS } from './shared/layouts.js';

// ================================================================
// department-news-display — Cloudflare Worker
// Renders active station or department news from Google Sheets.
// Supports recurrence, auto-scroll, and scheduled row cleanup.
//
// Security:
//   - All credentials stored as Cloudflare Worker secrets — never
//     in source code
//   - URL parameters validated against allowlists before use
//   - All sheet content HTML-escaped before injection into pages
//   - No X-Frame-Options header — this Worker is loaded as a
//     full-screen iframe; SAMEORIGIN causes immediate white screens
//   - Non-GET requests rejected to reduce attack surface
// ================================================================


// ================================================================
// CONFIGURATION — adjust these constants without touching logic
// ================================================================

/** Total seconds the news slide is shown on the display.
 *  Scroll speed is calculated from this value. */
const DISPLAY_DURATION_SECONDS = 20;

/** Items are highlighted as "new" if their posted date (or the
 *  start of the current recurrence cycle) is within this many days. */
const NEW_ITEM_THRESHOLD_DAYS = 3;

/** Seconds to pause before auto-scroll begins, giving viewers
 *  time to start reading before the content starts moving. */
const SCROLL_PAUSE_SECONDS = 3;

/** Slowest allowed scroll speed in pixels per second.
 *  Prevents content from scrolling too slowly to finish in time. */
const MIN_SCROLL_SPEED_PX_PER_SEC = 20;

/** Fastest allowed scroll speed in pixels per second.
 *  Prevents content from scrolling too fast to be readable. */
const MAX_SCROLL_SPEED_PX_PER_SEC = 120;

/** How long to cache Google Sheets API responses, in seconds.
 *  Reduces API calls; increase if the sheet changes infrequently. */
const CACHE_SECONDS = 300;

/** Increment this value to immediately bust the sheet data cache,
 *  such as after configuration changes or when stale data persists. */
const CACHE_VERSION = 1;

// --- Font configuration ---

/** Font size for news item titles. */
const FONT_SIZE_TITLE = '2.6rem';

/** Font size for metadata lines (Posted, Expires, Posted By). */
const FONT_SIZE_META = '0.8rem';

/** Font size for the main news body text. */
const FONT_SIZE_BODY = '1.4rem';

// Line height for the card body text (news item content).
// Increase for more breathing room between lines,
// decrease to fit more content per card. Applies to all body text including
// items with line breaks entered in the sheet.
const CARD_BODY_LINE_HEIGHT = 1.25;

// Line height for the card metadata block (Posted, Expires, Posted By).
// Adjust alongside CARD_BODY_LINE_HEIGHT if needed.
const CARD_META_LINE_HEIGHT = 1.6;

// --- Card color configuration ---
// Regular (non-new) items alternate between these two backgrounds.
// Both are subtle dark tints to visually separate cards against
// the charcoal display background.

/** Regular card color, even positions. */
const COLOR_REGULAR_A = 'rgba(255,255,255,0.06)';

/** Regular card color, odd positions. */
const COLOR_REGULAR_B = 'rgba(255,255,255,0.12)';

// "New" items alternate between these two backgrounds.
// Both are noticeably brighter than regular cards to signal recency.

/** New item card color, even positions. */
const COLOR_NEW_A = 'rgba(210,210,210,0.20)';

/** New item card color, odd positions. */
const COLOR_NEW_B = 'rgba(210,210,210,0.30)';

// --- Cleanup configuration ---

/** Non-recurring rows expired more than this many days ago are
 *  deleted by the daily cron job.
 *  Set to -1 to disable automatic deletion entirely. */
const DELETE_EXPIRED_AFTER_DAYS = 14;


// ================================================================
// SHEET COLUMN INDICES (0-based)
// Update these if columns are ever reordered in the sheet.
// ================================================================
const COL_TITLE      = 0; // News title
const COL_TEXT       = 1; // News body text
const COL_POSTED     = 2; // Posted date/time (visibility start)
const COL_EXPIRATION = 3; // Expiration date/time (visibility end)
const COL_POSTED_BY  = 4; // Name of person who posted
const COL_RECURRENCE = 5; // Recurrence interval in days (blank = none)
const COL_STOP_AFTER = 6; // Stop recurring after this date (blank = no end)


// ================================================================
// STATION → SHEET TAB NAME MAPPING
// Keys are valid ?station= parameter values (lowercase).
// Values must exactly match the tab names in the Google Sheet.
// Add entries here if stations are added in the future.
// ================================================================
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


// ================================================================
// VALID LAYOUT VALUES
// ================================================================
const VALID_LAYOUTS = ['split', 'wide', 'full', 'tri'];


// ================================================================
// MAIN EXPORT — fetch handler + scheduled cron handler
// ================================================================
export default {

  /**
   * Handles HTTP requests.
   * Validates parameters, fetches sheet data, and renders the news page.
   */
  async fetch(request, env) {

    // Allow GET and HEAD (HEAD is used by UptimeRobot health monitoring).
    // All other methods are rejected to reduce attack surface.
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      return new Response('Method not allowed', {
        status: 405,
        headers: { 'Allow': 'GET, HEAD' },
      });
    }

    const url          = new URL(request.url);

    if (url.pathname === '/healthz') {
      var healthStatus = 'healthy';
      var healthDetail = '';

      try {
        var probeRes = await fetchWithTimeout(
          'https://oauth2.googleapis.com/token',
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: 'grant_type=placeholder',
          },
          5000
        );
        if (probeRes.status === 400) {
          healthDetail = 'google-apis: reachable';
        } else {
          healthStatus = 'degraded';
          healthDetail = 'google-apis: unexpected status ' + probeRes.status;
        }
      } catch (e) {
        healthStatus = 'degraded';
        healthDetail = 'google-apis: unreachable (' + (e && e.message ? e.message : String(e)) + ')';
      }

      const healthBody =
        'status: ' + healthStatus + '\n' +
        'worker: department-news-display\n' +
        healthDetail + '\n';
      return new Response(
        request.method === 'HEAD' ? null : healthBody,
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

    // Parse ?bg=dark here — before any error responses — so every page
    // rendered by this worker can respect the dark background parameter,
    // including the invalid station error page below.
    const darkBg = sanitizeParam(url.searchParams.get('bg')) === 'dark';

    // Validate ?station= parameter against known tab names.
    const tabName = STATION_TAB_MAP[stationParam.toLowerCase()];
    if (!tabName) {
      return new Response(
        '<!DOCTYPE html>' +
        '<html lang="en">' +
        '<head><meta charset="UTF-8"><title>FFD News</title>' +
        '<style>' +
        '*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }' +
        'html, body { width: 100vw; height: 100vh; overflow: hidden;' +
        '  background: ' + (darkBg ? DARK_BG_COLOR : 'transparent') + ';' +
        '  font-family: ' + FONT_STACK + ';' +
        '  display: flex; align-items: center; justify-content: center; }' +
        '.err-wrap { display: flex; flex-direction: column; align-items: center; gap: 8px; text-align: center; padding: 0 5vw; }' +
        '.err-title { font-size: 1.8rem; font-weight: 700; color: ' + ACCENT_COLOR + '; letter-spacing: 0.06em; }' +
        '.err-sub   { font-size: 1.1rem; color: ' + TEXT_PRIMARY + '; }' +
        '</style></head>' +
        '<body>' +
        '<div class="err-wrap">' +
        '<div class="err-title">INVALID STATION</div>' +
        '<div class="err-sub">Check URL configuration</div>' +
        '</div>' +
        '</body></html>',
        { status: 400, headers: { 'Content-Type': 'text/html; charset=UTF-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' } }
      );
    }

    // Validate ?layout= parameter; fall back to 'split' if unrecognized.
    const layout = VALID_LAYOUTS.includes(layoutParam.toLowerCase())
      ? layoutParam.toLowerCase()
      : 'split';

    var REQUIRED_SECRETS = [
      'GOOGLE_SERVICE_ACCOUNT_EMAIL',
      'GOOGLE_PRIVATE_KEY',
      'GOOGLE_SHEET_ID'
    ];
    for (var i = 0; i < REQUIRED_SECRETS.length; i++) {
      var key = REQUIRED_SECRETS[i];
      if (!env[key]) {
        console.error('[department-news-display] Missing required secret: ' + key);
        return new Response(
          '<!DOCTYPE html>' +
          '<html lang="en"><head><meta charset="UTF-8">' +
          '<meta http-equiv="refresh" content="60">' +
          '<style>' +
          '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}' +
          'html,body{' +
          '  width:100vw;height:100vh;overflow:hidden;' +
          '  background:' + (darkBg ? DARK_BG_COLOR : 'transparent') + ';' +
          '  font-family:' + FONT_STACK + ';' +
          '  display:flex;align-items:center;justify-content:center;}' +
          '.err-wrap{display:flex;flex-direction:column;align-items:center;' +
          '  gap:8px;text-align:center;padding:0 5vw;}' +
          '.err-title{font-size:1.8rem;font-weight:700;color:' + ACCENT_COLOR + ';' +
          '  letter-spacing:0.06em;}' +
          '.err-sub{font-size:1.1rem;color:' + TEXT_PRIMARY + ';}' +
          '</style></head><body>' +
          '<div class="err-wrap">' +
          '<div class="err-title">CONFIGURATION ERROR</div>' +
          '<div class="err-sub">Missing secret: ' + key + '</div>' +
          '</div></body></html>',
          {
            status: 500,
            headers: {
              'Content-Type':           'text/html; charset=UTF-8',
              'Cache-Control':          'no-store',
              'X-Content-Type-Options': 'nosniff',
            },
          }
        );
      }
    }

    try {
      // Authenticate with Google and fetch sheet data.
      const token = await getAccessToken(
        env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        env.GOOGLE_PRIVATE_KEY,
        'https://www.googleapis.com/auth/spreadsheets'
      );
      const rows = await fetchSheetRows(env, token, tabName);

      // Process rows into active, displayable items.
      const now   = new Date();
      const items = processRows(rows, now);

      // Render and return the HTML page.
      const html = renderHtml(items, layout, tabName, darkBg);
      return new Response(html, {
        headers: {
          'Content-Type':           'text/html; charset=UTF-8',
          'Cache-Control':          'no-store',
          'X-Content-Type-Options': 'nosniff',
        },
      });

    } catch (err) {
      // Log full error server-side; return only a generic message to the
      // client to avoid leaking implementation details.
      console.error('Fetch handler error:', err);
      const errHtml =
        '<!DOCTYPE html>' +
        '<html lang="en">' +
        '<head>' +
        '<meta charset="UTF-8">' +
        '<meta http-equiv="refresh" content="60">' +
        '<title>FFD News</title>' +
        '<style>' +
        '*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }' +
        'html, body {' +
        '  width: 100vw; height: 100vh; overflow: hidden;' +
        '  background: ' + (darkBg ? DARK_BG_COLOR : 'transparent') + ';' +
        '  font-family: ' + FONT_STACK + ';' +
        '  display: flex; align-items: center; justify-content: center;' +
        '}' +
        '.err-wrap { display: flex; flex-direction: column; align-items: center; gap: 8px; text-align: center; padding: 0 5vw; }' +
        '.err-title { font-size: 1.8rem; font-weight: 700; color: ' + ACCENT_COLOR + '; letter-spacing: 0.06em; }' +
        '.err-sub   { font-size: 1.1rem; color: ' + TEXT_PRIMARY + '; }' +
        '</style>' +
        '</head>' +
        '<body>' +
        '<div class="err-wrap">' +
        '<div class="err-title">NEWS UNAVAILABLE</div>' +
        '<div class="err-sub">Retrying shortly</div>' +
        '</div>' +
        '</body>' +
        '</html>';
      return new Response(errHtml, {
        status: 200,
        headers: {
          'Content-Type':           'text/html; charset=UTF-8',
          'Cache-Control':          'no-store',
          'X-Content-Type-Options': 'nosniff',
        },
      });
    }
  },

  /**
   * Handles scheduled cron events.
   * Deletes non-recurring rows expired beyond the configured threshold.
   */
  async scheduled(event, env, ctx) {
    if (DELETE_EXPIRED_AFTER_DAYS < 0) {
      console.log('Automatic row deletion is disabled (DELETE_EXPIRED_AFTER_DAYS = -1). Skipping.');
      return;
    }
    ctx.waitUntil(runCleanup(env));
  },
};


// ================================================================
// SHEET DATA FETCHING
// ================================================================

/**
 * Fetches all data rows (excluding the header) from the specified tab.
 * Uses the Workers Cache API with a synthetic cache key so that the
 * rotating Google auth token does not prevent effective caching.
 *
 * @param {object} env
 * @param {string} token   - OAuth2 access token
 * @param {string} tabName - Sheet tab name to read
 * @returns {Promise<string[][]>} Array of row arrays (string values)
 */
async function fetchSheetRows(env, token, tabName) {

  // Synthetic cache key — stable across token rotations, versioned for
  // easy cache busting by incrementing CACHE_VERSION.
  const cacheKey = new Request(
    'https://cache.internal/news-display/v' + CACHE_VERSION +
    '/' + env.GOOGLE_SHEET_ID +
    '/' + encodeURIComponent(tabName)
  );
  const cache = caches.default;

  // Return cached rows if still fresh.
  const cached = await cache.match(cacheKey);
  if (cached) {
    return await cached.json();
  }

  // Fetch columns A–G from the sheet.
  // FORMATTED_VALUE returns dates as human-readable strings matching the
  // cell's display format, which parseSheetDateTime can reliably parse.
  const range  = encodeURIComponent(tabName + '!A:G');
  const apiUrl =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '/values/' + range +
    '?valueRenderOption=FORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING';

  const res = await fetchWithTimeout(apiUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
  }, 8000);

  if (!res.ok) {
    throw new Error('Sheets API error ' + res.status + ' fetching tab "' + tabName + '"');
  }

  const data    = await res.json();
  const allRows = data.values || [];

  // Row 0 is the header; data begins at row 1.
  const dataRows = allRows.slice(1);

  // Store processed rows in cache for CACHE_SECONDS.
  await cache.put(
    cacheKey,
    new Response(JSON.stringify(dataRows), {
      headers: { 'Cache-Control': 'max-age=' + CACHE_SECONDS },
    })
  );

  return dataRows;
}


// ================================================================
// ROW PROCESSING & RECURRENCE LOGIC
// ================================================================

/**
 * Processes raw sheet rows into active, displayable news item objects.
 * Applies recurrence math, visibility window filtering, and "new" flagging.
 * Results are sorted: new items first (newest posted date first within
 * that group), then regular items (newest posted date first).
 *
 * @param {string[][]} rows - Raw data rows from the sheet
 * @param {Date}       now  - Current time
 * @returns {object[]} Sorted array of active news item objects
 */
function processRows(rows, now) {
  const items = [];

  for (const row of rows) {
    // Skip rows missing any required field.
    if (!row[COL_TITLE] || !row[COL_TEXT] || !row[COL_POSTED] || !row[COL_EXPIRATION]) {
      continue;
    }

    const title    = row[COL_TITLE].trim();
    const text     = row[COL_TEXT].trim();
    const postedBy = (row[COL_POSTED_BY] || '').trim();

    const originalPosted  = parseSheetDateTime(row[COL_POSTED]);
    const originalExpires = parseSheetDateTime(row[COL_EXPIRATION]);
    if (!originalPosted || !originalExpires) continue;

    // Parse optional recurrence fields.
    const rawRecur     = (row[COL_RECURRENCE] || '').trim();
    const intervalDays = rawRecur !== '' ? parseFloat(rawRecur) : NaN;
    const isRecurring  = !isNaN(intervalDays) && intervalDays > 0;
    const stopAfter    = row[COL_STOP_AFTER]
      ? parseSheetDateTime(row[COL_STOP_AFTER])
      : null;

    let activePosted  = originalPosted;
    let activeExpires = originalExpires;

    if (isRecurring) {
      // --- Recurrence calculation ---
      // Each cycle shifts both posted and expires forward by intervalDays.
      // The duration within each cycle (posted → expires) stays fixed.
      const intervalMs      = intervalDays * 24 * 60 * 60 * 1000;
      const cycleDurationMs = originalExpires.getTime() - originalPosted.getTime();

      // How many complete intervals have elapsed since the original posted date?
      const elapsed      = now.getTime() - originalPosted.getTime();
      const cyclesPassed = Math.floor(elapsed / intervalMs);

      // Check the current cycle and the one before it to handle boundary
      // cases where 'now' sits right on a cycle edge.
      let foundActive = false;
      for (let n = (0, cyclesPassed); n >= (0, cyclesPassed - 1); n--) {
        const cyclePosted  = new Date(originalPosted.getTime() + n * intervalMs);
        const cycleExpires = new Date(cyclePosted.getTime() + cycleDurationMs);

        // Respect the stop-recurring-after date if one is set.
        if (stopAfter && cyclePosted > stopAfter) continue;

        // Is 'now' within this cycle's active window?
        if (now >= cyclePosted && now < cycleExpires) {
          activePosted  = cyclePosted;
          activeExpires = cycleExpires;
          foundActive   = true;
          break;
        }
      }

      // Not currently within any active cycle window — skip this row.
      if (!foundActive) continue;

    } else {
      // Non-recurring: straightforward date window check.
      if (now < originalPosted || now >= originalExpires) continue;
    }

    // Flag as "new" if the current cycle's posted date is within the threshold.
    const ageDays = (now.getTime() - activePosted.getTime()) / (1000 * 60 * 60 * 24);
    const isNew   = ageDays >= 0 && ageDays <= NEW_ITEM_THRESHOLD_DAYS;

    items.push({ title, text, postedBy, activePosted, activeExpires, isNew });
  }

  // Sort: new items first; within each group, newest posted date first.
  items.sort((a, b) => {
    if (a.isNew !== b.isNew) return a.isNew ? -1 : 1;
    return b.activePosted.getTime() - a.activePosted.getTime();
  });

  return items;
}

/**
 * Parses a date/time string returned by the Google Sheets API
 * (FORMATTED_VALUE mode), e.g. "4/8/2025 8:00 AM".
 * Returns null if the string is missing or cannot be parsed.
 *
 * @param {string} raw
 * @returns {Date|null}
 */
function parseSheetDateTime(raw) {
  if (!raw || typeof raw !== 'string') return null;
  var s = raw.trim();
  if (!s) return null;

  // Primary format: M/D/YYYY H:MM AM/PM (e.g. "4/18/2026 3:00 PM")
  // This is the format used in the department-news Google Sheet.
  // Parsed as America/Chicago wall-clock time using the offset-correction
  // technique from calendar-display's parseLocalDateTimeInZone — avoids the
  // bug where new Date("4/18/2026 3:00 PM") is treated as UTC by the
  // Cloudflare Workers runtime (which runs on UTC hosts).
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (m) {
    var mo   = parseInt(m[1], 10);
    var dy   = parseInt(m[2], 10);
    var yr   = parseInt(m[3], 10);
    // Normalize 2-digit years: 00-49 → 2000-2049, 50-99 → 1950-1999.
    // This matches the convention used by most spreadsheet applications.
    if (yr < 100) { yr = yr < 50 ? 2000 + yr : 1900 + yr; }
    var hr   = parseInt(m[4], 10);
    var mn   = parseInt(m[5], 10);
    var ampm = m[6] ? m[6].toUpperCase() : null;

    // Convert 12-hour clock to 24-hour
    if (ampm === 'PM' && hr !== 12) { hr += 12; }
    if (ampm === 'AM' && hr === 12) { hr = 0;   }

    // Step 1: treat the components as UTC to get an initial Date
    var utcApprox = new Date(Date.UTC(yr, mo - 1, dy, hr, mn, 0));

    // Step 2: find what wall-clock time that UTC instant shows in Chicago
    var parts = {};
    var formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/Chicago',
      year:     'numeric',
      month:    '2-digit',
      day:      '2-digit',
      hour:     '2-digit',
      minute:   '2-digit',
      second:   '2-digit',
      hour12:   false,
    });
    for (var i = 0; i < formatter.formatToParts(utcApprox).length; i++) {
      var part = formatter.formatToParts(utcApprox)[i];
      if (part.type !== 'literal') { parts[part.type] = part.value; }
    }

    // Step 3: compute the offset and correct the UTC instant so that
    // the wall-clock time in Chicago equals what was entered in the sheet
    var displayedMs = Date.UTC(
      parseInt(parts.year,   10),
      parseInt(parts.month,  10) - 1,
      parseInt(parts.day,    10),
      parseInt(parts.hour,   10) % 24,
      parseInt(parts.minute, 10),
      parseInt(parts.second, 10)
    );
    var intendedMs = Date.UTC(yr, mo - 1, dy, hr, mn, 0);

    return new Date(utcApprox.getTime() - (displayedMs - intendedMs));
  }

  // Fallback: native Date parsing for any format not matching the primary pattern.
  // Note: new Date() on an ambiguous string may still parse as UTC on this runtime.
  // If the sheet format ever changes, update the primary regex above rather than
  // relying on this fallback.
  var fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
}


// ================================================================
// HTML RENDERING
// ================================================================

/**
 * Formats a Date as a human-readable Central time string.
 * @param {Date} date
 * @returns {string}
 */
function formatDateTime(date) {
  return date.toLocaleString('en-US', {
    timeZone: 'America/Chicago',
    month:    'numeric',
    day:      'numeric',
    year:     'numeric',
    hour:     'numeric',
    minute:   '2-digit',
    hour12:   true,
  });
}

/**
 * Renders the complete HTML page for the news display.
 * Scroll constants are injected as inline JS variables so the
 * client-side scroll logic uses the server-side configuration values.
 *
 * @param {object[]} items   - Active, sorted news items
 * @param {string}   layout  - Validated layout parameter
 * @param {string}   tabName - Sheet tab name (used in <title> only)
 * @param {boolean}  darkBg  - True when ?bg=dark is set; renders solid dark
 *                             background for browser-based testing
 * @returns {string} Full HTML document string
 */
function renderHtml(items, layout, tabName, darkBg) {

  // Build card HTML. Track new/regular counts separately so the
  // alternating color resets independently between the two groups.
  let newCount     = 0;
  let regularCount = 0;

  const cardsHtml = items.length === 0
    ? '<div class="no-news">No current news</div>'
    : items.map(function (item) {

        let bgColor;
        let cardStyle;
        let titleDivider;

        if (item.isNew) {
          // New items: alternating brighter backgrounds + red left border stripe.
          bgColor   = (newCount % 2 === 0) ? COLOR_NEW_A : COLOR_NEW_B;
          // Left border is 4px red; remaining three sides use the standard
          // subtle white border. Left padding is reduced by 3px to compensate
          // for the wider border so body text stays visually aligned.
          cardStyle = 'background:' + bgColor + ';' +
                      'border-top:1px solid ' + BORDER_SUBTLE + ';' +
                      'border-right:1px solid ' + BORDER_SUBTLE + ';' +
                      'border-bottom:1px solid ' + BORDER_SUBTLE + ';' +
                      'border-left:1px solid ' + BORDER_SUBTLE + ';' +
                      'padding-left:calc(1rem - 3px);';
          titleDivider = '';
          newCount++;
        } else {
          // Regular items: alternating subtle backgrounds + red title underline.
          bgColor      = (regularCount % 2 === 0) ? COLOR_REGULAR_A : COLOR_REGULAR_B;
          cardStyle    = 'background:' + bgColor + ';' +
                         'border:1px solid ' + BORDER_SUBTLE + ';';
          // Thin red line beneath the title, matching the probationary display
          // divider style. Width is 60% to feel proportional without spanning
          // the full card.
          titleDivider = '<div class="title-divider"></div>';
          regularCount++;
        }

        const newBadge = item.isNew
          ? '<span class="new-badge">NEW</span>'
          : '';

        // Escape body text first, then convert literal newlines to <br>
        // so line breaks entered in the sheet are preserved on screen.
        const safeBody = escapeHtml(item.text).replace(/\n/g, '<br>');

        return (
          '<div class="news-card" style="' + cardStyle + '">' +
            '<div class="card-title">' + newBadge + escapeHtml(item.title) + '</div>' +
            titleDivider +
            '<div class="card-meta">' +
              '<div><strong>Posted:</strong> '    + formatDateTime(item.activePosted)  + '</div>' +
              '<div><strong>Expires:</strong> '   + formatDateTime(item.activeExpires) + '</div>' +
              '<div><strong>Posted By:</strong> ' + (escapeHtml(item.postedBy) || '\u2014') + '</div>' +
            '</div>' +
            '<div class="card-body">' + safeBody + '</div>' +
          '</div>'
        );
      }).join('');

  // Inline scroll script — configuration constants injected from Worker.
  const scrollScript =
    '(function () {' +
    '  var DISPLAY_DURATION_SECONDS = ' + DISPLAY_DURATION_SECONDS + ';' +
    '  var SCROLL_PAUSE_SECONDS     = ' + SCROLL_PAUSE_SECONDS     + ';' +
    '  var MIN_SPEED                = ' + MIN_SCROLL_SPEED_PX_PER_SEC + ';' +
    '  var MAX_SPEED                = ' + MAX_SCROLL_SPEED_PX_PER_SEC + ';' +
    '  function initScroll(attempt) {' +
    '    attempt = attempt || 0;' +
    '    var outer = document.getElementById("scroller");' +
    '    var inner = document.getElementById("scroll-inner");' +
    '    if (!outer || !inner) return;' +
    '    if (inner.dataset.noScroll) return;' +
    '    var viewH  = outer.clientHeight || window.innerHeight;' +
    '    var totalH = inner.offsetHeight;' +
    '    if ((viewH === 0 || totalH === 0) && attempt < 20) {' +
    '      setTimeout(function() { initScroll(attempt + 1); }, 100);' +
    '      return;' +
    '    }' +
    '    if (totalH <= viewH + 2) return;' +
    '    var overflow      = totalH - viewH;' +
    '    var availableTime = Math.max(1, DISPLAY_DURATION_SECONDS - (2 * SCROLL_PAUSE_SECONDS));' +
    '    var rawSpeed      = overflow / availableTime;' +
    '    var speed         = Math.min(MAX_SPEED, Math.max(MIN_SPEED, rawSpeed));' +
    '    var actualScrollTime = overflow / speed;' +
    '    var totalDuration    = SCROLL_PAUSE_SECONDS + actualScrollTime + SCROLL_PAUSE_SECONDS;' +
    '    var pauseTopPct      = (SCROLL_PAUSE_SECONDS / totalDuration) * 100;' +
    '    var pauseBottomPct   = 100 - ((SCROLL_PAUSE_SECONDS / totalDuration) * 100);' +
    '    var animationName    = "ffd-scroll-" + Date.now();' +
    '    var keyframes =' +
    '      "@keyframes " + animationName + " {" +' +
    '        "0% { transform: translateY(0); }" +' +
    '        pauseTopPct.toFixed(2) + "% { transform: translateY(0); }" +' +
    '        pauseBottomPct.toFixed(2) + "% { transform: translateY(-" + overflow + "px); }" +' +
    '        "100% { transform: translateY(-" + overflow + "px); }" +' +
    '      "}";' +
    '    var styleEl = document.createElement("style");' +
    '    styleEl.textContent = keyframes;' +
    '    document.head.appendChild(styleEl);' +
    '    inner.style.animation = animationName + " " + totalDuration + "s linear infinite";' +
    '  }' +
    '  if (document.readyState === "complete") {' +
    '    setTimeout(initScroll, 500);' +
    '  } else {' +
    '    window.addEventListener("load", function () {' +
    '      setTimeout(initScroll, 500);' +
    '    });' +
    '  }' +
    '}());';

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

    '#scroller {' +
    '  height: 100%;' +
    '  overflow: hidden;' +
    '}' +

    '#scroll-inner {' +
    '  will-change: transform;' +
    '}' +

    '.no-news {' +
    '  text-align: center;' +
    '  font-size: ' + FONT_SIZE_TITLE + ';' +
    '  color: rgba(255,255,255,0.55);' +
    '  margin-top: 3rem;' +
    '}' +

    // Card base — border and background are set inline per card so that
    // new and regular items can use different border styles.
    '.news-card {' +
    '  border-radius: 6px;' +
    '  padding: 0.8rem 1rem;' +
    '  margin-bottom: 0.65rem;' +
    '}' +

    // Title row — flex so the NEW badge and title sit on the same line.
    '.card-title {' +
    '  font-size: '   + FONT_SIZE_TITLE + ';' +
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

    // Red accent underline beneath the title on regular (non-new) cards.
    // Matches the divider style used in the probationary firefighter display.
    '.title-divider {' +
    '  width: 60%;' +
    '  height: 2px;' +
    '  background: ' + ACCENT_COLOR + ';' +
    '  opacity: 0.85;' +
    '  border-radius: 1px;' +
    '  margin-bottom: 0.45rem;' +
    '}' +

    // Metadata block — Posted, Expires, Posted By each on their own line.
    '.card-meta {' +
    '  font-size: ' + FONT_SIZE_META + ';' +
    '  color: ' + TEXT_SECONDARY + ';' +
    '  line-height: ' + CARD_META_LINE_HEIGHT + ';' +
    '  margin-bottom: 0.65rem;' +
    '}' +

    // Body text — preserves line breaks entered in the sheet.
    '.card-body {' +
    '  font-size: ' + FONT_SIZE_BODY + ';' +
    '  color: ' + TEXT_PRIMARY + ';' +
    '  line-height: ' + CARD_BODY_LINE_HEIGHT + ';' +
    '}';

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

/**
 * Scans all configured sheet tabs and deletes non-recurring rows
 * that expired more than DELETE_EXPIRED_AFTER_DAYS days ago.
 * Deletions are processed bottom-up within each tab so that earlier
 * row indices remain valid as rows are removed during the batch.
 *
 * @param {object} env
 */
async function runCleanup(env) {
  var REQUIRED_CLEANUP_SECRETS = [
    'GOOGLE_SERVICE_ACCOUNT_EMAIL',
    'GOOGLE_PRIVATE_KEY',
    'GOOGLE_SHEET_ID'
  ];
  for (var ci = 0; ci < REQUIRED_CLEANUP_SECRETS.length; ci++) {
    var ckey = REQUIRED_CLEANUP_SECRETS[ci];
    if (!env[ckey]) {
      console.error('[department-news-display] runCleanup: missing secret: ' + ckey);
      return;
    }
  }

  const token       = await getAccessToken(
    env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    env.GOOGLE_PRIVATE_KEY,
    'https://www.googleapis.com/auth/spreadsheets'
  );
  const now         = new Date();
  const thresholdMs = DELETE_EXPIRED_AFTER_DAYS * 24 * 60 * 60 * 1000;

  // Fetch spreadsheet metadata to resolve tab names → internal numeric sheet IDs.
  const metaRes = await fetchWithTimeout(
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '?fields=sheets(properties(title,sheetId))',
    { headers: { 'Authorization': 'Bearer ' + token } },
    8000
  );
  if (!metaRes.ok) {
    console.error('Cleanup: failed to fetch spreadsheet metadata (' + metaRes.status + ')');
    return;
  }
  const meta = await metaRes.json();

  // Build a title → numeric sheetId lookup map.
  const sheetIdByTitle = {};
  for (const sheet of (meta.sheets || [])) {
    sheetIdByTitle[sheet.properties.title] = sheet.properties.sheetId;
  }

  // Process each configured tab independently.
  for (const tabName of Object.values(STATION_TAB_MAP)) {
    const sheetId = sheetIdByTitle[tabName];
    if (sheetId === undefined) {
      console.warn('Cleanup: tab "' + tabName + '" not found in spreadsheet — skipping.');
      continue;
    }

    // Fetch all rows for this tab fresh — bypass the Worker cache to ensure
    // deletion decisions are based on current data, not cached data.
    const range  = encodeURIComponent(tabName + '!A:G');
    const apiUrl =
      'https://sheets.googleapis.com/v4/spreadsheets/' +
      encodeURIComponent(env.GOOGLE_SHEET_ID) +
      '/values/' + range +
      '?valueRenderOption=FORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING';

    const rowsRes = await fetchWithTimeout(apiUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
    }, 8000);
    if (!rowsRes.ok) {
      console.error('Cleanup: failed to fetch rows for "' + tabName + '" (' + rowsRes.status + ')');
      continue;
    }

    const rowsData = await rowsRes.json();
    const allRows  = rowsData.values || [];
    // allRows[0] = header row. Data rows begin at index 1.
    // Array index == 0-based sheet row index for batchUpdate deleteDimension.

    const deleteRequests = [];

    // Traverse data rows bottom-up so earlier indices stay valid as rows
    // are removed during the subsequent batch delete.
    for (let i = allRows.length - 1; i >= 1; i--) {
      const row = allRows[i];
      if (!row) continue;

      const rawExpires = (row[COL_EXPIRATION] || '').trim();
      const rawRecur   = (row[COL_RECURRENCE] || '').trim();
      if (!rawExpires) continue;

      // Never delete recurring rows — they cycle forward indefinitely.
      const intervalDays = rawRecur !== '' ? parseFloat(rawRecur) : NaN;
      if (!isNaN(intervalDays) && intervalDays > 0) continue;

      const expires = parseSheetDateTime(rawExpires);
      if (!expires) continue;

      const expiredAgoMs = now.getTime() - expires.getTime();
      if (expiredAgoMs >= thresholdMs) {
        deleteRequests.push({
          deleteDimension: {
            range: {
              sheetId,
              dimension:  'ROWS',
              startIndex: i,
              endIndex:   i + 1,
            },
          },
        });
      }
    }

    if (deleteRequests.length === 0) {
      console.log('Cleanup: no expired rows to delete in "' + tabName + '".');
      continue;
    }

    // Execute all deletions for this tab in a single batch request.
    const batchRes = await fetchWithTimeout(
      'https://sheets.googleapis.com/v4/spreadsheets/' +
      encodeURIComponent(env.GOOGLE_SHEET_ID) + ':batchUpdate',
      {
        method:  'POST',
        headers: {
          'Authorization':  'Bearer ' + token,
          'Content-Type':   'application/json',
        },
        body: JSON.stringify({ requests: deleteRequests }),
      },
      8000
    );

    if (!batchRes.ok) {
      console.error('Cleanup: batch delete failed for "' + tabName + '" (' + batchRes.status + ')');
    } else {
      console.log('Cleanup: deleted ' + deleteRequests.length + ' expired row(s) from "' + tabName + '".');
    }
  }
}
