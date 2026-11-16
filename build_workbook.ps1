param(
    [string]$Template = "C:\Users\RRNandyal\Downloads\empty _templeate.xlsx",
    [string]$Output = (Join-Path $PSScriptRoot "NIFTY_EOD_Automation.xlsm")
)

$ErrorActionPreference = "Stop"
$excel = $null
$workbook = $null
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("nifty-eod-" + [guid]::NewGuid().ToString("N"))
$expanded = Join-Path $tempRoot "expanded"
$workingTemplate = Join-Path $tempRoot "template-unprotected.xlsx"

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Directory]::CreateDirectory($expanded) | Out-Null
    [IO.Compression.ZipFile]::ExtractToDirectory($Template, $expanded)
    $workbookXml = Join-Path $expanded "xl\workbook.xml"
    $xmlText = [IO.File]::ReadAllText($workbookXml)
    $xmlText = [regex]::Replace($xmlText, "<workbookProtection\b[^>]*/>", "")
    [IO.File]::WriteAllText($workbookXml, $xmlText, [Text.UTF8Encoding]::new($false))
    foreach ($sheetXml in Get-ChildItem (Join-Path $expanded "xl\worksheets") -Filter "*.xml") {
        $sheetText = [IO.File]::ReadAllText($sheetXml.FullName)
        $sheetText = [regex]::Replace($sheetText, "<sheetProtection\b[^>]*/>", "")
        [IO.File]::WriteAllText($sheetXml.FullName, $sheetText, [Text.UTF8Encoding]::new($false))
    }
    [IO.Compression.ZipFile]::CreateFromDirectory($expanded, $workingTemplate)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($workingTemplate)

    foreach ($sheetName in @("Settings", "Expiry Summary", "Holidays", "Calendar", "Expiry List", "Notes", "Trade Plan")) {
        try {
            $workbook.Worksheets.Item($sheetName).Delete()
        } catch {
            # Sheet does not exist in a clean template.
        }
    }

    # Expiry-day VWAP is distorted by settlement, so the last Tuesday is dropped
    # from the model: clear its Money block first, then delete the sheet itself.
    $money = $workbook.Worksheets.Item("Money")
    $null = $money.Range("AP1:AS1200").ClearContents()
    $money.Range("AU3:AU201").Formula = '=C3+G3+K3+O3+S3+W3+AA3+AE3+AI3+AM3'
    $money.Range("AV3:AV201").Formula = '=D3+H3+L3+P3+T3+X3+AB3+AF3+AJ3+AN3'
    try {
        $workbook.Worksheets.Item("Tue3").Delete()
    } catch {
        # Already absent when rebuilding from an updated template.
    }

    $settings = $workbook.Worksheets.Add()
    $settings.Name = "Settings"
    $settings.Move($workbook.Worksheets.Item(1))
    $settings.Range("A1:D1").Merge()
    $settings.Range("A1").Value2 = "NIFTY EOD OPTION AUTOMATION"
    $settings.Range("A1").Font.Bold = $true
    $settings.Range("A1").Font.Size = 16
    $settings.Range("A1").Interior.Color = 15773696

    $labels = @(
        "iCharts ID",
        "iCharts Password",
        "TradingView ID (optional)",
        "TradingView Password (optional)",
        "Symbol",
        "Expiry Date (blank = auto Tuesday)",
        "Data Range From (auto)",
        "Data Range To (auto)",
        "As Of Date (blank = latest iCharts EOD)",
        "Strike Step",
        "Normal Lookback Days",
        "Wednesday Post-Expiry Lookback Days",
        "AVWAP Source",
        "Download Folder (blank = workbook folder)",
        "Last Successful Update",
        "Status",
        "Big Money Coverage Target",
        "Hedge Width (strike steps)"
    )
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $row = $i + 2
        $settings.Cells.Item($row, 1).Value2 = $labels[$i]
        $settings.Cells.Item($row, 1).Font.Bold = $true
    }

    $settings.Range("B6").Value2 = "NIFTY"
    $settings.Range("B7").Value2 = [datetime]"2026-09-08"
    $settings.Range("B8").Formula = '=IF($B$7="","",$B$7-IF(AND(WEEKDAY(TODAY(),2)=3,TODAY()>$B$7),$B$13,$B$12))'
    $settings.Range("B9").Formula = '=IF($B$7="","",IF($B$10="",$B$7,MIN($B$10,$B$7)))'
    $settings.Range("B7:B10").NumberFormat = "dd-mmm-yyyy"
    $settings.Range("B8:B9").Font.Italic = $true
    $settings.Range("B8:B9").Interior.Color = 14737632
    $settings.Range("B11").Value2 = 50
    $settings.Range("B12").Value2 = 14
    $settings.Range("B13").Value2 = 7
    $settings.Range("B14").Value2 = "iCharts EOD VWAP weighted by daily volume"
    $settings.Range("B16").NumberFormat = "dd-mmm-yyyy hh:mm"
    $settings.Range("B18").Value2 = 0.8
    $settings.Range("B18").NumberFormat = "0%"
    $settings.Range("B19").Value2 = 4
    $settings.Range("B3").NumberFormat = ";;;"
    $settings.Range("B5").NumberFormat = ";;;"
    $settings.Range("D2").Value2 = "All explanations now live on the Notes sheet."
    $settings.Range("D2").Font.Italic = $true
    $settings.Range("A2:B19").Borders.LineStyle = 1
    $settings.Columns.Item("A").ColumnWidth = 39
    $settings.Columns.Item("B").ColumnWidth = 35
    $settings.Columns.Item("D").ColumnWidth = 46
    $settings.Columns.Item("F").ColumnWidth = 31
    $settings.Rows.Item("1:19").RowHeight = 22

    # Forms buttons cannot be coloured, so the two actions are shapes instead.
    # They are parked over column F so they never cover the settings or notes.
    $green = 3381504
    $red = 192
    $refresh = $settings.Shapes.AddShape(5, 330, 40, 210, 48)
    $refresh.Name = "btnRefreshEOD"
    $refresh.Fill.ForeColor.RGB = $green
    $refresh.Line.ForeColor.RGB = 2308878
    $refresh.TextFrame2.TextRange.Text = "REFRESH EOD DATA"
    $refresh.TextFrame2.TextRange.Font.Size = 12
    $refresh.TextFrame2.TextRange.Font.Bold = -1
    $refresh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $refresh.OnAction = "RefreshEOD"

    $reset = $settings.Shapes.AddShape(5, 330, 96, 210, 40)
    $reset.Name = "btnResetWorkbook"
    $reset.Fill.ForeColor.RGB = $red
    $reset.Line.ForeColor.RGB = 128
    $reset.TextFrame2.TextRange.Text = "RESET WORKBOOK"
    $reset.TextFrame2.TextRange.Font.Size = 12
    $reset.TextFrame2.TextRange.Font.Bold = -1
    $reset.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $reset.OnAction = "ResetWorkbook"

    $refresh.Left = $settings.Range("F2").Left
    $refresh.Top = $settings.Range("F2").Top
    $reset.Left = $settings.Range("F5").Left
    $reset.Top = $settings.Range("F5").Top

    $holidays = $workbook.Worksheets.Add()
    $holidays.Name = "Holidays"
    $holidays.Range("A1").Value2 = "Holiday Date"
    $holidays.Range("B1").Value2 = "Note"
    $holidays.Range("A1:B1").Font.Bold = $true
    $holidays.Range("A1:B1").Interior.Color = 13434828
    $holidays.Columns.Item("A").ColumnWidth = 16
    $holidays.Columns.Item("B").ColumnWidth = 52

    # Official NSE equity-derivatives (FO) trading holidays, read from
    # nseindia.com/api/holiday-master?type=trading. The IMPORT button refreshes them.
    $nseHolidays = @(
        @("2026-01-15", "Municipal Corporation Election - Maharashtra"),
        @("2026-01-26", "Republic Day"),
        @("2026-02-15", "Mahashivratri"),
        @("2026-03-03", "Holi"),
        @("2026-03-21", "Id-Ul-Fitr (Ramadan Eid)"),
        @("2026-03-26", "Shri Ram Navami"),
        @("2026-03-31", "Shri Mahavir Jayanti"),
        @("2026-04-03", "Good Friday"),
        @("2026-04-14", "Dr. Baba Saheb Ambedkar Jayanti"),
        @("2026-05-01", "Maharashtra Day"),
        @("2026-05-28", "Bakri Id"),
        @("2026-06-26", "Muharram"),
        @("2026-08-15", "Independence Day"),
        @("2026-09-14", "Ganesh Chaturthi"),
        @("2026-10-02", "Mahatma Gandhi Jayanti"),
        @("2026-10-20", "Dussehra"),
        @("2026-11-08", "Diwali Laxmi Pujan (Muhurat trading)"),
        @("2026-11-10", "Diwali-Balipratipada"),
        @("2026-11-24", "Prakash Gurpurb Sri Guru Nanak Dev"),
        @("2026-12-25", "Christmas")
    )
    for ($i = 0; $i -lt $nseHolidays.Count; $i++) {
        $row = $i + 2
        $d = [datetime]$nseHolidays[$i][0]
        $holidays.Cells.Item($row, 1).Formula = "=DATE({0},{1},{2})" -f $d.Year, $d.Month, $d.Day
        $holidays.Cells.Item($row, 2).Value2 = "NSE: " + $nseHolidays[$i][1]
        $holidays.Cells.Item($row, 3).Formula = "=TEXT(A$row,""dddd"")"
    }
    $holidayLastRow = $nseHolidays.Count + 1
    $holidays.Range("C1").Value2 = "Weekday"
    $holidays.Range("C1").Font.Bold = $true
    $holidays.Range("C1").Interior.Color = 13434828
    $holidays.Range("A2:A$holidayLastRow").NumberFormat = "dd-mmm-yyyy"
    $holidays.Range("A1:C$holidayLastRow").Borders.LineStyle = 1
    $holidays.Columns.Item("C").ColumnWidth = 12
    $holidays.Rows.Item("1:1").RowHeight = 22

    $importButton = $holidays.Shapes.AddShape(5, 700, 60, 220, 40)
    $importButton.Name = "btnImportNseHolidays"
    $importButton.Fill.ForeColor.RGB = $green
    $importButton.Line.ForeColor.RGB = 2308878
    $importButton.TextFrame2.TextRange.Text = "IMPORT NSE HOLIDAYS"
    $importButton.TextFrame2.TextRange.Font.Size = 11
    $importButton.TextFrame2.TextRange.Font.Bold = -1
    $importButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $importButton.OnAction = "ImportNseHolidays"

    $holidays.Range("D1").Value2 = "Last import"
    $holidays.Range("D1").Font.Bold = $true
    $holidays.Range("E1").Value2 = "Seeded from NSE equity-derivatives trading holidays for 2026."
    $holidays.Columns.Item("D").ColumnWidth = 12
    $holidays.Columns.Item("E").ColumnWidth = 58
    $importButton.Left = $holidays.Range("G3").Left
    $importButton.Top = $holidays.Range("G3").Top

    $notes = $workbook.Worksheets.Add()
    $notes.Name = "Notes"
    $notes.Range("A1:B1").Merge()
    $notes.Range("A1").Value2 = "HOW THIS WORKBOOK WORKS"
    $notes.Range("A1").Font.Bold = $true
    $notes.Range("A1").Font.Size = 15
    $notes.Range("A1").Interior.Color = 15773696
    $noteRows = @(
        @("HEAD", "Buttons"),
        @("Settings: green REFRESH EOD DATA", "Logs in to iCharts, deletes the old day files, downloads every trading day of the window, fills the day sheets, rebuilds the Money grid and recalculates the range model."),
        @("Settings: red RESET WORKBOOK", "Empties the 10 day sheets, the Money grid, the Top 10 AVWAP cells, the Expiry Summary and the downloaded files. Settings, logins and holidays are kept."),
        @("Holidays: green IMPORT NSE HOLIDAYS", "Reads the official NSE equity-derivatives trading holidays and rewrites every row whose note starts with 'NSE:'. Your own rows are untouched."),
        @("HEAD", "Logins"),
        @("iCharts ID / Password", "Required. Read at run time only and never saved in the code or in git. The password cell looks blank because of a hidden number format; the value is still there."),
        @("TradingView ID / Password", "Optional and unused. AVWAP is calculated from the imported EOD VWAP and volume, so no TradingView login is needed."),
        @("HEAD", "Dates and expiry"),
        @("Expiry Date", "Pick from the drop-down, double-click for a calendar, or double-click a row on the Expiry List sheet. Leave blank for the automatic Tuesday expiry."),
        @("Data Range From / To", "Calculated, not editable input. They show the exact window the next refresh will download."),
        @("As Of Date", "Blank means the latest EOD date iCharts has. Double-click for a calendar."),
        @("Normal Lookback Days", "Normal mode loads expiry minus 14 calendar days through the As Of Date."),
        @("Wednesday Post-Expiry Lookback Days", "On a Wednesday after the selected expiry has completed, the shorter 7-day window is used instead."),
        @("HEAD", "Data and files"),
        @("Download Folder", "Blank uses 'iCharts Downloads' next to the workbook. Type any folder you can write to, in this machine's path style: C:\Users\you\Downloads on Windows, /Users/you/Downloads on macOS. Every refresh first deletes the SYMBOL_*.csv files there, then downloads fresh copies. Other files are left alone."),
        @("HEAD", "Windows and macOS"),
        @("Where the refresh runs", "REFRESH EOD DATA needs Windows Excel. The iCharts login and download use WinHttp and Scripting.Dictionary, and neither exists in Mac Excel, so on a Mac the button stops with a message instead of a cryptic error."),
        @("Using a Mac", "Either run this workbook in Windows Excel (a VM or a Windows PC), or have the day CSVs downloaded by a separate script and let the Mac workbook only import them. Ask for the script route if you want it."),
        @("Folder paths", "Paths are assembled with the separator of whichever machine is running, so the same workbook works on both once the Download Folder is set correctly."),
        @("Day files", "One CSV per trading day, named SYMBOL_EXPIRY_YYYY-MM-DD.csv, copied into Tue1 ... Mon3."),
        @("Expiry day is excluded", "The window runs from the Tuesday two weeks before expiry to the Monday before it, so it covers the previous Tuesday and last Tuesday but never expiry day itself. Settlement distorts the expiry-day VWAP, so AVWAP and the range ignore it."),
        @("Holidays", "Dates on the Holidays sheet are never fetched. A day that returns no iCharts data is appended automatically, so the list improves itself."),
        @("HEAD", "Trade Plan sheet"),
        @("Sell strikes", "The strike holding the most writer money within six steps beyond each defended boundary. Big money defends those strikes, so they are the ones to sell into."),
        @("Hedge strikes", "Hedge Width (strike steps) away from the sold strike. Four steps at a 50-point step means a 200-point wing, capping the loss at width minus credit."),
        @("Verdict", "Confirms both break-evens sit outside the defended range. WEAK or WATCH means the credit is too small for the risk being taken."),
        @("Premiums", "Latest EOD LTP found in the loaded window, so live quotes will differ. Sizing, margin and stops are your decision."),
        @("HEAD", "Range model on Expiry Summary"),
        @("Big Money Coverage Target", "Share of Top-10 writer money that must stay protected on BOTH the call and put side. 80% is the recommended default."),
        @("Strict floor and ceiling", "The highest Top-10 PE break-even and the lowest Top-10 CE break-even. If the floor is above the ceiling there is no price where every large writer avoids loss, and the sheet says so plainly."),
        @("Defended range", "The capital-weighted range that protects the target share of writer money on each side, with the protection actually achieved shown next to it."),
        @("Suggested pin and closing zone", "The strike that minimises money-weighted break-even damage, and that pin rounded to tradable strikes. A model estimate, not a guaranteed settlement.")
    )
    $noteRow = 3
    foreach ($entry in $noteRows) {
        if ($entry[0] -eq "HEAD") {
            $notes.Cells.Item($noteRow, 1).Value2 = $entry[1]
            $notes.Range("A$noteRow" + ":B$noteRow").Merge()
            $notes.Cells.Item($noteRow, 1).Font.Bold = $true
            $notes.Cells.Item($noteRow, 1).Font.Size = 12
            $notes.Cells.Item($noteRow, 1).Interior.Color = 13434828
        } else {
            $notes.Cells.Item($noteRow, 1).Value2 = $entry[0]
            $notes.Cells.Item($noteRow, 1).Font.Bold = $true
            $notes.Cells.Item($noteRow, 2).Value2 = $entry[1]
            $notes.Cells.Item($noteRow, 2).WrapText = $true
        }
        $noteRow++
    }
    $notes.Columns.Item("A").ColumnWidth = 38
    $notes.Columns.Item("B").ColumnWidth = 96
    $notes.Range("A3:B" + ($noteRow - 1)).Borders.LineStyle = 1
    $notes.Range("A1:B" + ($noteRow - 1)).VerticalAlignment = -4160
    $null = $notes.Rows.Item("3:" + ($noteRow - 1)).AutoFit()

    $calendar = $workbook.Worksheets.Add()
    $calendar.Name = "Calendar"
    $calendar.Range("A1").Value2 = "<<"
    $calendar.Range("G1").Value2 = ">>"
    $calendar.Range("B1:F1").Merge()
    $calendar.Range("A1:G1").Font.Bold = $true
    $calendar.Range("A1:G1").Interior.Color = 15773696
    $calendar.Range("A2:G2").Merge()
    $calendar.Range("A2").Value2 = "Click a day to use it. A1 = previous month, G1 = next month."
    $dayNames = @("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    for ($i = 0; $i -lt $dayNames.Count; $i++) {
        $calendar.Cells.Item(3, $i + 1).Value2 = $dayNames[$i]
    }
    $calendar.Range("A3:G3").Font.Bold = $true
    $calendar.Range("A3:G3").Interior.Color = 13434828
    $calendar.Range("A1:G9").HorizontalAlignment = -4108
    $calendar.Range("A1:G9").Borders.LineStyle = 1
    $calendar.Columns.Item("A").ColumnWidth = 6
    $calendar.Range("A1:G9").ColumnWidth = 6
    $calendar.Rows.Item("1:9").RowHeight = 22

    $expiryList = $workbook.Worksheets.Add()
    $expiryList.Name = "Expiry List"
    $expiryList.Range("A1").Value2 = "Expiry Date"
    $expiryList.Range("B1").Value2 = "Data Range From"
    $expiryList.Range("C1").Value2 = "Data Range To"
    $expiryList.Range("D1").Value2 = "Range"
    $expiryList.Range("E1").Value2 = "Status"
    $expiryList.Range("A1:E1").Font.Bold = $true
    $expiryList.Range("A1:E1").Interior.Color = 13434828
    $expiryList.Range("G1").Value2 = "Double-click any row to make that expiry active on Settings."
    $expiryList.Range("G2").Value2 = "Weekly Tuesday expiries. For an older Thursday/Wednesday expiry, type the date on Settings instead."

    $firstExpiry = [datetime]"2026-01-06"
    $lastExpiry = [datetime]"2027-12-31"
    $weeks = [int][math]::Floor(($lastExpiry - $firstExpiry).TotalDays / 7) + 1
    $lastListRow = $weeks + 1
    $expiryList.Range("A2").Formula = "=DATE({0},{1},{2})" -f $firstExpiry.Year, $firstExpiry.Month, $firstExpiry.Day
    if ($lastListRow -ge 3) {
        $expiryList.Range("A3:A$lastListRow").Formula = '=$A2+7'
    }
    $expiryList.Range("B2:B$lastListRow").Formula = '=IF($A2="","",$A2-Settings!$B$12)'
    $expiryList.Range("C2:C$lastListRow").Formula = '=IF($A2="","",$A2)'
    $expiryList.Range("D2:D$lastListRow").Formula = '=TEXT($B2,"dd-mmm-yyyy")&"  to  "&TEXT($C2,"dd-mmm-yyyy")'
    $expiryList.Range("E2:E$lastListRow").Formula = '=IF($A2<TODAY(),"past",IF($A2=TODAY(),"today","upcoming"))'
    $expiryList.Range("A2:C$lastListRow").NumberFormat = "dd-mmm-yyyy"
    $expiryList.Range("A1:E$lastListRow").Borders.LineStyle = 1
    $expiryList.Range("A1:E$lastListRow").HorizontalAlignment = -4108
    $expiryList.Columns.Item("A").ColumnWidth = 15
    $expiryList.Columns.Item("B").ColumnWidth = 17
    $expiryList.Columns.Item("C").ColumnWidth = 17
    $expiryList.Columns.Item("D").ColumnWidth = 34
    $expiryList.Columns.Item("E").ColumnWidth = 11
    $expiryList.Columns.Item("G").ColumnWidth = 70
    $expiryList.Rows.Item("1:1").RowHeight = 22

    $validation = $settings.Range("B7").Validation
    $validation.Delete()
    $validation.Add(3, 1, 1, "='Expiry List'!`$A`$2:`$A`$$lastListRow")
    $validation.IgnoreBlank = $true
    $validation.InCellDropdown = $true
    $validation.ShowError = $false
    $validation.InputTitle = "Expiry"
    $validation.InputMessage = "Pick a weekly expiry, or double-click for a calendar. Blank = next Tuesday."

    $summary = $workbook.Worksheets.Add()
    $summary.Name = "Expiry Summary"
    $summary.Range("A1:B1").Merge()
    $summary.Range("A1").Value2 = "SUPPORT / RESISTANCE AND EXPIRY RANGE"
    $summary.Range("A1").Font.Bold = $true
    $summary.Range("A1").Font.Size = 15
    $summary.Range("A1").Interior.Color = 13434828
    $summaryLabels = @{
        2 = "Expiry Date"
        3 = "Data Loaded From"
        4 = "Data Loaded Through"
        5 = "Trading Days Loaded"
        7 = "Strongest Put Strike (Support)"
        8 = "Put AVWAP"
        9 = "Put BEP / Lower Range"
        11 = "Strongest Call Strike (Resistance)"
        12 = "Call AVWAP"
        13 = "Call BEP / Upper Range"
        15 = "Strict PE Floor (max Top-10 PE BEP)"
        16 = "Strict CE Ceiling (min Top-10 CE BEP)"
        17 = "Can Every Top-10 Writer Avoid Loss?"
        19 = "Capital Protection Target"
        20 = "Defended Range Low"
        21 = "Defended Range High"
        22 = "PE Money Protected at Low"
        23 = "CE Money Protected at High"
        24 = "Two-Sided Confidence"
        26 = "Suggested Money-Weighted Pin"
        27 = "Suggested Closing Zone Low"
        28 = "Suggested Closing Zone High"
        29 = "Suggested Closing Zone"
        30 = "Important"
    }
    foreach ($entry in $summaryLabels.GetEnumerator()) {
        $summary.Cells.Item([int]$entry.Key, 1).Value2 = $entry.Value
        $summary.Cells.Item([int]$entry.Key, 1).Font.Bold = $true
    }
    $summary.Columns.Item("A").ColumnWidth = 42
    $summary.Columns.Item("B").ColumnWidth = 48
    $summary.Range("A2:B30").Borders.LineStyle = 1
    $summary.Range("A15:B17").Interior.Color = 13421823
    $summary.Range("A19:B24").Interior.Color = 13561798
    $summary.Range("A26:B29").Interior.Color = 13434828
    $summary.Range("A30:B30").Interior.Color = 10092543
    $summary.Range("B30").WrapText = $true
    $summary.Rows.Item(30).RowHeight = 34

    $trade = $workbook.Worksheets.Add()
    $trade.Name = "Trade Plan"
    $trade.Range("A1:B1").Merge()
    $trade.Range("A1").Value2 = "STRIKES TO SELL AND HEDGE"
    $trade.Range("A1").Font.Bold = $true
    $trade.Range("A1").Font.Size = 15
    $trade.Range("A1").Interior.Color = 13434828
    $tradeLabels = @{
        2 = "CALL SIDE (bear call spread)"
        3 = "Sell CE Strike"
        4 = "Writer Money at Sell Strike"
        5 = "Sell CE Premium (latest EOD LTP)"
        6 = "Buy CE Strike (hedge)"
        7 = "Buy CE Premium"
        8 = "Net Credit per Lot (points)"
        9 = "Max Loss per Lot (points)"
        11 = "PUT SIDE (bull put spread)"
        12 = "Sell PE Strike"
        13 = "Writer Money at Sell Strike"
        14 = "Sell PE Premium (latest EOD LTP)"
        15 = "Buy PE Strike (hedge)"
        16 = "Buy PE Premium"
        17 = "Net Credit per Lot (points)"
        18 = "Max Loss per Lot (points)"
        20 = "COMBINED (iron condor)"
        21 = "Total Credit (points)"
        22 = "Lower Break-Even"
        23 = "Upper Break-Even"
        24 = "Money-Weighted Pin"
        25 = "Verdict"
        26 = "Order Summary"
        27 = "Important"
    }
    foreach ($entry in $tradeLabels.GetEnumerator()) {
        $trade.Cells.Item([int]$entry.Key, 1).Value2 = $entry.Value
        $trade.Cells.Item([int]$entry.Key, 1).Font.Bold = $true
    }
    $trade.Range("A2:B2,A11:B11,A20:B20").Interior.Color = 15773696
    $trade.Range("A2:B2,A11:B11,A20:B20").Font.Size = 12
    $trade.Range("A3:B9").Interior.Color = 13561798
    $trade.Range("A12:B18").Interior.Color = 13561798
    $trade.Range("A21:B24").Interior.Color = 13434828
    $trade.Range("A25:B27").Interior.Color = 10092543
    $trade.Range("B25:B27").WrapText = $true
    $trade.Columns.Item("A").ColumnWidth = 36
    $trade.Columns.Item("B").ColumnWidth = 62
    $trade.Range("A2:B27").Borders.LineStyle = 1
    $trade.Range("D3").Value2 = "How the strikes are chosen"
    $trade.Range("D3").Font.Bold = $true
    $trade.Range("D4").Value2 = "Sell strikes: the strike carrying the most writer money within 6 steps beyond each defended boundary on Expiry Summary - the wall big players defend."
    $trade.Range("D5").Value2 = "Hedge strikes: Hedge Width (strike steps) on Settings away from the sold strike, so the loss on each side is capped at width minus credit."
    $trade.Range("D6").Value2 = "Verdict checks whether both break-evens sit outside the defended range. If they sit inside, the spread is being paid too little for the risk."
    $trade.Range("D4:D6").WrapText = $true
    $trade.Columns.Item("D").ColumnWidth = 62
    $trade.Rows.Item("4:6").RowHeight = 46
    $trade.Rows.Item("25:27").RowHeight = 32

    try {
        foreach ($component in @($workbook.VBProject.VBComponents)) {
            if ($component.Name -in @("modOptionAutomation", "modCalendar")) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        foreach ($moduleName in @("modOptionAutomation.bas", "modCalendar.bas")) {
            $null = $workbook.VBProject.VBComponents.Import((Join-Path $PSScriptRoot $moduleName))
        }

        $settingsCode = @'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim addr As String
    addr = Target.Address(False, False)
    If addr = "B7" Or addr = "B10" Then
        Cancel = True
        ShowCalendarFor Target
    End If
End Sub
'@
        $calendarCode = @'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    CalendarClicked Target
End Sub
'@
        $expiryListCode = @'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Row >= 2 And Target.Column <= 5 Then
        Cancel = True
        UseExpiryFromList Target
    End If
End Sub
'@
        $workbook.VBProject.VBComponents.Item($settings.CodeName).CodeModule.AddFromString($settingsCode)
        $workbook.VBProject.VBComponents.Item($calendar.CodeName).CodeModule.AddFromString($calendarCode)
        $workbook.VBProject.VBComponents.Item($expiryList.CodeName).CodeModule.AddFromString($expiryListCode)
        $calendar.Visible = 0
    } catch {
        throw "Excel blocked VBA module import. In Excel, enable File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', then rerun this builder. Original error: $($_.Exception.Message)"
    }

    # Sheets added later land in front of Settings, so fix the tab order last.
    $settings.Move($workbook.Worksheets.Item(1))
    $notes.Move($workbook.Worksheets.Item(2))
    $summary.Move($workbook.Worksheets.Item(3))
    $trade.Move($workbook.Worksheets.Item(4))
    $null = $settings.Activate()
    $null = $settings.Range("A1").Select()

    $workbook.SaveAs($Output, 52)
    $workbook.Close($true)
    $workbook = $null
    Write-Output "Created $Output"
}
finally {
    if ($workbook -ne $null) {
        try { $workbook.Close($false) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tempRoot) {
        Remove-Item $tempRoot -Recurse -Force
    }
}
