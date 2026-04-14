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
const DISPLAY_DURATION_SECONDS = 60;

/** Items are highlighted as "new" if their posted date (or the
 *  start of the current recurrence cycle) is within this many days. */
const NEW_ITEM_THRESHOLD_DAYS = 3;

/** Seconds to pause before auto-scroll begins, giving viewers
 *  time to start reading before the content starts moving. */
const SCROLL_PAUSE_SECONDS = 5;

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

/** Font family applied to all text on the display. */
const FONT_FAMILY = "'Segoe UI', Arial, sans-serif";

/** Font size for news item titles. */
const FONT_SIZE_TITLE = '1.4rem';

/** Font size for metadata lines (Posted, Expires, Posted By). */
const FONT_SIZE_META = '0.85rem';

/** Font size for the main news body text. */
const FONT_SIZE_BODY = '1rem';

// --- Accent color ---

/** FFD brand red — used as a left border stripe on new items and
 *  as a title underline on regular items. */
const ACCENT_COLOR = '#C8102E';

/** Background color used when ?bg=dark is set.
 *  Matches the probationary-firefighter-display dark testing background. */
const DARK_BG_COLOR = '#111111';

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
const DELETE_EXPIRED_AFTER_DAYS = 7;


// ================================================================
// GOOGLE AUTH SCOPE
// spreadsheets scope is required for both read (display) and
// write (row deletion) access to the sheet.
// ================================================================
const GOOGLE_AUTH_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';


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

    // Reject non-GET requests with a generic status to reduce attack surface.
    if (request.method !== 'GET') {
      return new Response('Method not allowed', { status: 405 });
    }

    const url          = new URL(request.url);
    const stationParam = sanitizeParam(url.searchParams.get('station')) || '';
    const layoutParam  = sanitizeParam(url.searchParams.get('layout'))  || 'split';

    // Validate ?station= parameter against known tab names.
    const tabName = STATION_TAB_MAP[stationParam.toLowerCase()];
    if (!tabName) {
      return new Response(
        'Invalid or missing ?station= parameter. Valid values: dept, 1\u20138.',
        { status: 400, headers: { 'Content-Type': 'text/plain; charset=UTF-8' } }
      );
    }

    // Validate ?layout= parameter; fall back to 'split' if unrecognized.
    const layout = VALID_LAYOUTS.includes(layoutParam.toLowerCase())
      ? layoutParam.toLowerCase()
      : 'split';

    // ?bg=dark renders with a solid dark background for browser-based testing.
    // Matches the probationary-firefighter-display ?bg=dark parameter behaviour.
    const darkBg = sanitizeParam(url.searchParams.get('bg')) === 'dark';

    try {
      // Authenticate with Google and fetch sheet data.
      const token = await getAccessToken(
        env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        env.GOOGLE_PRIVATE_KEY
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
        '  font-family: "Segoe UI", Arial, Helvetica, sans-serif;' +
        '  display: flex; align-items: center; justify-content: center;' +
        '}' +
        '.err-wrap { display: flex; flex-direction: column; align-items: center; gap: 8px; text-align: center; padding: 0 5vw; }' +
        '.err-title { font-size: 1.8rem; font-weight: 700; color: rgba(255,255,255,0.92); letter-spacing: 0.06em; }' +
        '.err-sub   { font-size: 1.1rem; color: rgba(255,255,255,0.55); }' +
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
// GOOGLE SERVICE ACCOUNT AUTHENTICATION
// ================================================================
// Generates a short-lived Google OAuth2 access token from service
// account credentials stored as Worker secrets. Uses RSA-SHA256
// JWT signing via the Web Crypto API — no external dependencies.
//
// Required secrets:
//   GOOGLE_SERVICE_ACCOUNT_EMAIL — service account email address
//   GOOGLE_PRIVATE_KEY           — RSA private key from Google
//                                  Cloud JSON key file

/**
 * Builds a signed JWT and exchanges it for a Google OAuth2 access token.
 *
 * @param {string} email         - Service account email address
 * @param {string} rawPrivateKey - PEM private key (with literal \n sequences)
 * @returns {Promise<string>} A valid OAuth2 access token
 */
async function getAccessToken(email, rawPrivateKey) {

  // Step 1 — Build the JWT header and payload.
  const now     = Math.floor(Date.now() / 1000);
  const header  = base64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = base64url(JSON.stringify({
    iss:   email,
    scope: GOOGLE_AUTH_SCOPE,
    aud:   'https://oauth2.googleapis.com/token',
    iat:   now,
    exp:   now + 3600,
  }));

  const signingInput = header + '.' + payload;

  // Step 2 — Import the RSA private key via the Web Crypto API.
  // The key arrives from the secret with literal \n sequences;
  // convert them to real newlines before stripping the PEM envelope.
  // Both PKCS#8 and traditional RSA key headers are handled.
  const pemString = rawPrivateKey.replace(/\\n/g, '\n');
  const pemBody   = pemString
    .replace('-----BEGIN PRIVATE KEY-----',     '')
    .replace('-----END PRIVATE KEY-----',       '')
    .replace('-----BEGIN RSA PRIVATE KEY-----', '')
    .replace('-----END RSA PRIVATE KEY-----',   '')
    .replace(/\n/g, '')
    .trim();

  const binaryKey = Uint8Array.from(atob(pemBody), c => c.charCodeAt(0));

  const cryptoKey = await crypto.subtle.importKey(
    'pkcs8',
    binaryKey,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['sign']
  );

  // Step 3 — Sign the JWT.
  const signatureBuf = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    cryptoKey,
    new TextEncoder().encode(signingInput)
  );

  const jwt = signingInput + '.' + arrayBufferToBase64url(signatureBuf);

  // Step 4 — Exchange the signed JWT for a short-lived access token.
  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body:    'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt,
  });

  if (!tokenRes.ok) {
    const errText = await tokenRes.text();
    throw new Error('Token exchange failed (' + tokenRes.status + '): ' + errText);
  }

  const tokenData = await tokenRes.json();
  return tokenData.access_token;
}

/**
 * Encodes a UTF-8 string to base64url format (used in JWT construction).
 * @param {string} str
 * @returns {string}
 */
function base64url(str) {
  return btoa(str)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=/g,  '');
}

/**
 * Converts an ArrayBuffer to base64url using a safe byte-by-byte loop.
 * The spread operator can throw a RangeError on large buffers (such as
 * RSA signatures) — this approach avoids that risk entirely.
 * @param {ArrayBuffer} buffer
 * @returns {string}
 */
function arrayBufferToBase64url(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary  = '';
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=/g,  '');
}

/**
 * Strips leading/trailing whitespace from a URL parameter value and
 * removes any characters that are not alphanumeric, hyphens, or underscores.
 * Returns null if the input is null.
 * @param {string|null} val
 * @returns {string|null}
 */
function sanitizeParam(val) {
  if (val === null) return null;
  return val.trim().replace(/[^a-zA-Z0-9\-_#]/g, '');
}


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

  const res = await fetch(apiUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
  });

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
      for (let n = Math.max(0, cyclesPassed); n >= Math.max(0, cyclesPassed - 1); n--) {
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
  const d = new Date(raw.trim());
  return isNaN(d.getTime()) ? null : d;
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
 * Escapes characters with special meaning in HTML to prevent XSS.
 * Must be called on every user-supplied string before HTML injection.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
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
                      'border-top:1px solid rgba(255,255,255,0.10);' +
                      'border-right:1px solid rgba(255,255,255,0.10);' +
                      'border-bottom:1px solid rgba(255,255,255,0.10);' +
                      'border-left:4px solid ' + ACCENT_COLOR + ';' +
                      'padding-left:calc(1rem - 3px);';
          titleDivider = '';
          newCount++;
        } else {
          // Regular items: alternating subtle backgrounds + red title underline.
          bgColor      = (regularCount % 2 === 0) ? COLOR_REGULAR_A : COLOR_REGULAR_B;
          cardStyle    = 'background:' + bgColor + ';' +
                         'border:1px solid rgba(255,255,255,0.10);';
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
    '  window.addEventListener("load", function () {' +
    '    var totalH = document.documentElement.scrollHeight;' +
    '    var viewH  = window.innerHeight;' +
    '    if (totalH <= viewH) return;' +
    '    var overflow      = totalH - viewH;' +
    '    var availableTime = Math.max(1, DISPLAY_DURATION_SECONDS - SCROLL_PAUSE_SECONDS);' +
    '    var rawSpeed      = overflow / availableTime;' +
    '    var speed         = Math.min(MAX_SPEED, Math.max(MIN_SPEED, rawSpeed));' +
    '    var pauseElapsed  = 0;' +
    '    var scrollStarted = false;' +
    '    var lastTimestamp = null;' +
    '    function step(timestamp) {' +
    '      if (lastTimestamp === null) { lastTimestamp = timestamp; }' +
    '      var delta = (timestamp - lastTimestamp) / 1000;' +
    '      lastTimestamp = timestamp;' +
    '      if (!scrollStarted) {' +
    '        pauseElapsed += delta;' +
    '        if (pauseElapsed >= SCROLL_PAUSE_SECONDS) { scrollStarted = true; }' +
    '      } else {' +
    '        window.scrollBy(0, speed * delta);' +
    '        if (window.scrollY + viewH >= totalH) { return; }' +
    '      }' +
    '      requestAnimationFrame(step);' +
    '    }' +
    '    requestAnimationFrame(step);' +
    '  });' +
    '}());';

  const css =
    '*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }' +

    'body {' +
    '  background: ' + (darkBg ? DARK_BG_COLOR : 'transparent') + ';' +
    '  color: #f0f0f0;' +
    '  font-family: ' + FONT_FAMILY + ';' +
    '  font-size: '   + FONT_SIZE_BODY + ';' +
    '  padding: 0.75rem;' +
    '  overflow-x: hidden;' +
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
    '  color: rgba(255,255,255,0.68);' +
    '  line-height: 1.6;' +
    '  margin-bottom: 0.65rem;' +
    '}' +

    // Body text — preserves line breaks entered in the sheet.
    '.card-body {' +
    '  font-size: ' + FONT_SIZE_BODY + ';' +
    '  color: rgba(255,255,255,0.92);' +
    '  line-height: 1.55;' +
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
    cardsHtml +
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
  const token       = await getAccessToken(
    env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    env.GOOGLE_PRIVATE_KEY
  );
  const now         = new Date();
  const thresholdMs = DELETE_EXPIRED_AFTER_DAYS * 24 * 60 * 60 * 1000;

  // Fetch spreadsheet metadata to resolve tab names → internal numeric sheet IDs.
  const metaRes = await fetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(env.GOOGLE_SHEET_ID) +
    '?fields=sheets(properties(title,sheetId))',
    { headers: { 'Authorization': 'Bearer ' + token } }
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

    const rowsRes = await fetch(apiUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
    });
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
    const batchRes = await fetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' +
      encodeURIComponent(env.GOOGLE_SHEET_ID) + ':batchUpdate',
      {
        method:  'POST',
        headers: {
          'Authorization':  'Bearer ' + token,
          'Content-Type':   'application/json',
        },
        body: JSON.stringify({ requests: deleteRequests }),
      }
    );

    if (!batchRes.ok) {
      console.error('Cleanup: batch delete failed for "' + tabName + '" (' + batchRes.status + ')');
    } else {
      console.log('Cleanup: deleted ' + deleteRequests.length + ' expired row(s) from "' + tabName + '".');
    }
  }
}