# Outstanding Dashboard — setup

Two files were added to the Vendor Collection Apps Script project (already saved into
`D:\Shefali - Projects\Vendor collection` on your computer):

- **code.gs** — your original file, unchanged except for three additions at the bottom
  (`showDashboard`, `doGet`, `getDashboardData` and two small helpers) and one new menu
  item. Nothing in `setupDealerMaster`, `updateAndProtectAgeing`, `getData`, or
  `sendWhatsAppFromSidebar` was touched.
- **Dashboard.html** — the new dashboard page.

## To deploy

1. Open the Google Sheet this Apps Script project is bound to, then **Extensions → Apps
   Script**.
2. In the Apps Script editor, open `Code.gs` (or whatever your file is named there) and
   replace its contents with the new `code.gs`. If you've made other edits in Apps
   Script since this was written, merge the **Outstanding Dashboard** section at the
   bottom in by hand instead of overwriting — everything above the
   `OUTSTANDING DASHBOARD` comment block is your original code.
3. Add a new HTML file named exactly `Dashboard` (Apps Script drops the `.html`) and
   paste in the contents of `Dashboard.html`.
4. Save, reload the Sheet. The menu **📱 WhatsApp Tool** now has a **📊 Open Outstanding
   Dashboard** item — it opens the dashboard in a modal dialog reading live from
   `AGEING_ANALYSIS`.
5. Optional — for a version you can open on a phone without opening the Sheet:
   **Deploy → New deployment → Web app**, execute as you, accessible to whoever you
   want (e.g. yourself or the executives). That URL runs the same page via `doGet()`.

## What it shows

- KPI tiles: total outstanding, dealers with dues, 90+-day overdue amount, and dealers
  with dues who haven't had a WhatsApp sent yet.
- Ageing breakdown (0–30 / 31–60 / 61–90 / 90+ days) computed from each month column
  against today's date — no manual bucket column needed.
- Outstanding trend across the month columns already in your sheet.
- Executive-wise and city-wise breakdown.
- Top 15 overdue dealers, and a searchable/filterable/sortable table of every dealer
  with a per-row ageing split and their WhatsApp status.
- Export CSV button on the filtered table.

## Notes

- Ageing buckets are derived from calendar months versus today, the same "Mon YY"
  format `sortMonthHeaders` already parses — no changes to how `AGEING_ANALYSIS` is
  structured are needed.
- "City" and "WhatsApp status" come straight from the `City` and `Status` columns
  already in `AGEING_ANALYSIS`; dealers without a city show as "Unknown".
- A shareable **preview with sample data** (so you can see the design without touching
  Apps Script yet) was published — the CSV export button doesn't work inside that
  preview sandbox, only in the deployed version.
