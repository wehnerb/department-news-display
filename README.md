# Department News Display

A Cloudflare Worker that fetches active news items from Google Sheets and renders them as a styled HTML page for fire station display screens. Supports per-station and department-wide feeds, recurring news items, auto-scrolling for long content, and automated cleanup of expired rows.

## 📄 System Documentation
Full documentation (architecture, setup, account transfer, IT reference): https://github.com/wehnerb/ffd-display-system-documentation

---

## Live URLs

| Environment | URL |
|---|---|
| Production | `https://department-news-display.bwehner.workers.dev/` |
| Staging | `https://department-news-display-staging.bwehner.workers.dev/` |

---

## URL Parameters

| Parameter | Required | Options |
|---|---|---|
| `?station=` | Yes | `dept`, `1`–`8` |
| `?layout=` | No | `split` (default), `wide`, `full`, `tri` |

**Examples:**
- `?station=dept` — Department-wide news feed
- `?station=1` — Station 1 news feed
- `?station=1&layout=wide` — Station 1 news, wide layout

---

## Layout Parameter

| Layout | Width | Height |
|---|---|---|
| `split` | 852px | 720px |
| `wide` | 1735px | 720px |
| `full` | 1920px | 1075px |
| `tri` | 558px | 720px |

---

## Configuration (`src/index.js`)

All routine configuration is at the top of `src/index.js`.

| Constant | Default | Description |
|---|---|---|
| `DISPLAY_DURATION_SECONDS` | `60` | Seconds the slide is shown; controls scroll speed |
| `NEW_ITEM_THRESHOLD_DAYS` | `3` | Days after posted date an item is highlighted as new |
| `SCROLL_PAUSE_SECONDS` | `5` | Seconds to pause before auto-scroll begins |
| `MIN_SCROLL_SPEED_PX_PER_SEC` | `20` | Slowest allowed scroll speed (px/sec) |
| `MAX_SCROLL_SPEED_PX_PER_SEC` | `120` | Fastest allowed scroll speed (px/sec) |
| `CACHE_SECONDS` | `300` | How long sheet data is cached (seconds) |
| `CACHE_VERSION` | *(current)* | Increment to immediately invalidate all cached data |
| `FONT_FAMILY` | `'Segoe UI', Arial, sans-serif` | Font applied to all display text |
| `FONT_SIZE_TITLE` | `1.4rem` | News item title font size |
| `FONT_SIZE_META` | `0.85rem` | Metadata line font size (Posted, Expires, Posted By) |
| `FONT_SIZE_BODY` | `1rem` | News body text font size |
| `COLOR_REGULAR_A` | See code | Regular card background, even positions |
| `COLOR_REGULAR_B` | See code | Regular card background, odd positions |
| `COLOR_NEW_A` | See code | New item card background, even positions |
| `COLOR_NEW_B` | See code | New item card background, odd positions |
| `DELETE_EXPIRED_AFTER_DAYS` | `60` | Days after expiration before a non-recurring row is deleted. Set to `-1` to disable deletion. |

---

## Google Sheet Structure

Single Google Sheets file with one tab per station plus a department tab.

| Tab Name | Feed |
|---|---|
| `Department News` | `?station=dept` |
| `FS#1` | `?station=1` |
| `FS#2`–`FS#8` | `?station=2`–`?station=8` |

**Column layout (all tabs):**

| Column | Header | Notes |
|---|---|---|
| A | Title | Required |
| B | Text | Required. Alt+Enter line breaks are preserved on display. |
| C | Posted Date/Time | Required. Controls visibility start. |
| D | Expiration Date/Time | Required. Controls visibility end. |
| E | Posted By | Displayed on screen. |
| F | Recurrence Interval (days) | Optional. Blank = one-time item. Example: `14` = every 2 weeks. |
| G | Stop Recurring After | Optional. Date only. Blank = recurs indefinitely. |

The service account must have **Editor** access to the file (required for the cleanup cron to delete rows).

---

## Secrets

| Secret | Where Set | Description |
|---|---|---|
| `CLOUDFLARE_API_TOKEN` | GitHub Actions | Cloudflare API token — Workers edit permissions |
| `CLOUDFLARE_ACCOUNT_ID` | GitHub Actions | Cloudflare account ID |
| `GOOGLE_SERVICE_ACCOUNT_EMAIL` | Cloudflare Dashboard | Service account email address |
| `GOOGLE_PRIVATE_KEY` | Cloudflare Dashboard | RSA private key from Google Cloud JSON key file |
| `GOOGLE_SHEET_ID` | Cloudflare Dashboard | Google Sheets file ID |

---

## Scheduled Cleanup

A cron trigger runs daily at **3:00 AM CST / 4:00 AM CDT** (9:00 UTC). It scans all tabs and deletes non-recurring rows that expired more than `DELETE_EXPIRED_AFTER_DAYS` days ago. Recurring rows are never deleted by the cleanup job. Set `DELETE_EXPIRED_AFTER_DAYS = -1` to disable deletion entirely.

---

## Deployment

| Branch | Deploys To | Purpose |
|---|---|---|
| `staging` | `department-news-display-staging.bwehner.workers.dev` | Testing |
| `main` | `department-news-display.bwehner.workers.dev` | Production |

Push to either branch — GitHub Actions deploys automatically (~30–45 sec).  
**Always stage and test before merging to main.**  
To roll back: use the Cloudflare dashboard **Deployments** tab, then revert the commit on `main`.
