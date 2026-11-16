# NIFTY EOD Excel automation

`NIFTY_EOD_Automation.xlsm` is built from the supplied empty template. It keeps the existing `Money` and `Top 10` model and adds:

- `Settings`: editable credentials, expiry, as-of date, strike step, lookback, and download folder.
- `REFRESH EOD DATA`: logs in to iCharts, downloads one all-strikes file per trading day, then copies each file into `Tue1` through `Mon3`.
- `RESET WORKBOOK`: empties the day sheets, the `Money` grid, the Top 10 AVWAP cells, `Expiry Summary` and the downloaded day files, ready for the next expiry. Settings, logins and the holiday list are kept.
- `Expiry List`: every weekly expiry with the data range it will pull. Double-click a row to make that expiry active.
- Calendar picker: double-click the Expiry Date or As Of Date cell. Expiry Date also has a drop-down of the listed expiries.
- `Data Range From` / `Data Range To` on `Settings`: calculated cells showing the exact window the next refresh will fetch.
- `Holidays`: seeded with the official NSE equity-derivatives trading holidays, refreshable with the green `IMPORT NSE HOLIDAYS` button. Dates listed here are never fetched, and days with no iCharts data are appended automatically.
- `Notes`: every explanation that used to clutter the `Settings` sheet, grouped by topic.
- `Trade Plan`: which strike to sell and which to buy as the hedge on each side, with credit, capped loss, break-evens and a verdict.
- Automatic resizing of the `Money` strike grid to the full imported strike range.
- Automatic Top 10 AVWAP using volume-weighted iCharts daily VWAP.
- `Expiry Summary`: strongest support/resistance plus a capital-weighted writer range:
  - strict all-Top-10 writer overlap (or an explicit `NO OVERLAP` warning),
  - configurable defended range (80% of CE and PE money by default),
  - actual two-sided protection confidence,
  - money-weighted minimum-pain pin and a strike-rounded closing zone.

## Use

1. Open `NIFTY_EOD_Automation.xlsm` in desktop Excel and enable macros.
2. On `Settings`, enter the iCharts ID and password.
3. Set the expiry date from the drop-down, the calendar, or by double-clicking a row on `Expiry List`. Leave it blank for automatic Tuesday expiry selection, and leave As Of Date blank to use the latest iCharts EOD.
4. Check `Data Range From` / `Data Range To` on `Settings` — that is the window that will be downloaded.
5. Click `REFRESH EOD DATA`.

Switching to a different expiry needs no manual cleanup: the refresh clears the download folder and the day sheets itself. `RESET WORKBOOK` is there for when you want the workbook left empty.

Every refresh deletes the existing `SYMBOL_*.csv` files in the download folder before downloading, so each run is a fresh copy and nothing from a previous run or expiry can be reused. Files that do not match that pattern, such as your own spreadsheets or another symbol's downloads, are left untouched. The count of deleted files is reported in the completion message and in the `Status` cell.

Downloaded day files are written to `Download Folder` (default `iCharts Downloads` next to the workbook) as one CSV per trading day, named `SYMBOL_EXPIRY_YYYY-MM-DD.csv`. These hold the same rows as the website's Download CSV button, which is generated in the browser and therefore cannot be fetched as a URL; the macro reads the same option-chain feed the page uses and writes the files itself.

`RefreshEODSilent` and `ResetWorkbookSilent` run the same refresh and reset without message boxes, for scheduled or scripted runs.

Passwords are not stored in source code or git. They are read from Settings **B2** (ID) and **B3** (password) at run time. B3 looks blank because of a hidden number format; the value is still there. If refresh fails, read **B17** — that cell keeps the real error text.

TradingView credentials are present as editable optional fields because they were requested, but the macro does not submit them. TradingView has no supported Excel/VBA data API for this calculation. AVWAP is calculated locally from the same imported option-chain EOD VWAP and volume, avoiding fragile browser screen scraping.

The `Trade Plan` sheet turns the defended range into legs. On each side it sells the strike carrying the most writer money within six steps beyond the boundary — the wall large writers are defending — and buys a hedge `Hedge Width (strike steps)` further out, so the loss is capped at width minus credit. It then reports the total credit, both break-evens and a verdict that flags a missing premium, a leg that prices as a debit, or a break-even sitting inside the defended range. Premiums are the latest EOD LTP in the loaded window, so live quotes will differ; sizing, margin and stops are not modelled.

`Big Money Coverage Target` on `Settings` controls the defended range. At 80%, the lower boundary is the capital-weighted PE BEP that protects at least 80% of selected PE money; the upper boundary is the capital-weighted CE BEP that protects at least 80% of selected CE money. The suggested pin minimizes money-weighted break-even violations across the Top 10, and the closing zone is rounded to available strike steps. This is a model estimate—not a guaranteed expiry settlement or trading recommendation.

`IMPORT NSE HOLIDAYS` reads `nseindia.com/api/holiday-master?type=trading` and takes the `FO` (equity derivatives) segment. NSE serves that feed only to browser-like clients, so the macro loads the holidays page first to collect cookies and sends Chrome-style headers. Rows whose note begins `NSE:` are replaced on every import; rows you typed yourself and auto-detected ones are kept. If NSE blocks the request, the sheet keeps the seeded list and the reason is written to `Holidays E1`.

Expiry day itself is never loaded. The window runs from the Tuesday two weeks before expiry to the Monday before it, so it spans the previous Tuesday and last Tuesday but stops one day short of expiry, whose VWAP is distorted by settlement. That is why there are ten day sheets, `Tue1` through `Mon3`, and why the As Of Date is capped at expiry minus one day.

Normal mode loads the fixed expiry-relative dates from expiry minus 14 calendar days through expiry/as-of date. On Wednesday after the selected expiry has completed, it uses the configurable 7-day lookback. Weekends and holidays are skipped, and their weekday sheet stays blank rather than shifting data into the wrong sheet.

## Rebuild

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_workbook.ps1
```

If Excel blocks module injection, temporarily enable:

`File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model`

Then rerun the builder. The standalone `modOptionAutomation.bas` is also included for review or manual import.
