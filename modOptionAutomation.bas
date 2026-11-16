Attribute VB_Name = "modOptionAutomation"
Option Explicit

Private Const LOGIN_URL As String = "https://www.icharts.in/opt/login.php"
Private Const CHAIN_URL As String = "https://www.icharts.in/opt/OptionChain_Old.php"
Private Const TABLE_URL As String = "https://www.icharts.in/opt/OptionChainTable_v16.php"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const SUMMARY_SHEET As String = "Expiry Summary"
Private Const TRADE_SHEET As String = "Trade Plan"
Private Const DEFAULT_MIN_STRIKE As Double = 21000
Private Const DEFAULT_STRIKE_ROWS As Long = 199
Private Const NSE_HOME_URL As String = "https://www.nseindia.com/resources/exchange-communication-holidays"
Private Const NSE_HOLIDAY_URL As String = "https://www.nseindia.com/api/holiday-master?type=trading"

Private gCookies As Object
Private gNseCookies As Object
Private gLastHtml As String
Private gSessionId As String
Private gUserName As String
Private gSilent As Boolean

' Runs the same refresh without pop-ups, for scripted or scheduled use.
Public Sub RefreshEODSilent()
    gSilent = True
    On Error Resume Next
    RefreshEOD
    On Error GoTo 0
    gSilent = False
End Sub

' Empties every sheet the refresh fills, so the workbook is ready for another expiry.
' Settings, credentials and the holiday list are deliberately left alone.
Public Sub ResetWorkbook()
    If Not gSilent Then
        If MsgBox("Clear the 10 day sheets, the Money grid, the Top 10 AVWAP values, the Expiry Summary" & _
                  " and the downloaded day files?" & vbCrLf & vbCrLf & _
                  "Settings, logins and the holiday list are kept.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Reset workbook") <> vbYes Then Exit Sub
    End If

    ResetAll

    If Not gSilent Then
        MsgBox "Workbook cleared." & vbCrLf & _
               "Pick the expiry date on Settings, then press REFRESH EOD DATA.", _
               vbInformation, "NIFTY Option Automation"
    End If
End Sub

' Same reset without pop-ups, for scripted use.
Public Sub ResetWorkbookSilent()
    gSilent = True
    On Error Resume Next
    ResetAll
    On Error GoTo 0
    gSilent = False
End Sub

Private Sub ResetAll()
    Dim wsCfg As Worksheet, strikeStep As Long, symbol As String, removedFiles As Long
    Dim oldCalc As XlCalculation, errMsg As String

    oldCalc = Application.Calculation
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsCfg = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    strikeStep = CLng(Val(wsCfg.Range("B11").Value2))
    If strikeStep <= 0 Then strikeStep = 50
    symbol = UCase$(Trim$(CStr(wsCfg.Range("B6").Value2)))
    If symbol = "" Then symbol = "NIFTY"

    removedFiles = CleanDownloadFolder(ResolveDownloadFolder(wsCfg), symbol)
    ClearRawSheets wipeFormats:=True
    WriteMoneyGrid DEFAULT_MIN_STRIKE, strikeStep, DEFAULT_STRIKE_ROWS
    ThisWorkbook.Worksheets(SUMMARY_SHEET).Range("B2:B30").ClearContents
    On Error Resume Next
    ThisWorkbook.Worksheets(TRADE_SHEET).Range("B3:B27").ClearContents
    On Error GoTo Failed
    wsCfg.Range("B16").ClearContents
    wsCfg.Range("B17").Value = "Reset on " & Format$(Now, "dd-mmm-yyyy hh:mm") & ": " & _
                               removedFiles & " day file(s) deleted. Ready for a new refresh."

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

Failed:
    errMsg = Trim$(Err.Description)
    On Error Resume Next
    If Not wsCfg Is Nothing Then wsCfg.Range("B17").Value = Left$("RESET ERROR: " & errMsg, 255)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If oldCalc <> 0 Then Application.Calculation = oldCalc
    If Not gSilent Then MsgBox errMsg, vbCritical, "Reset failed"
End Sub

Public Sub RefreshEOD()
    Dim wsCfg As Worksheet
    Dim symbol As String, userId As String, password As String
    Dim expiryDate As Date, asOfDate As Date
    Dim normalLookback As Long, wedLookback As Long, lookbackDays As Long
    Dim strikeStep As Long, expiryCode As String
    Dim offsets As Variant, sheetNames As Variant
    Dim i As Long, tradeDate As Date, importedRows As Long, importedDays As Long
    Dim skippedDays As Long, downloadFolder As String, filePath As String
    Dim removedFiles As Long
    Dim firstLoaded As Date
    Dim latestDateText As String
    Dim oldCalc As XlCalculation
    Dim errMsg As String

    oldCalc = xlCalculationAutomatic
    On Error GoTo Failed

    Set wsCfg = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    userId = Trim$(CStr(wsCfg.Range("B2").Value2))
    password = SettingsPassword(wsCfg.Range("B3"))
    symbol = UCase$(Trim$(CStr(wsCfg.Range("B6").Value2)))
    strikeStep = CLng(Val(wsCfg.Range("B11").Value2))
    normalLookback = CLng(Val(wsCfg.Range("B12").Value2))
    wedLookback = CLng(Val(wsCfg.Range("B13").Value2))
    downloadFolder = ResolveDownloadFolder(wsCfg)

    If userId = "" Or password = "" Then Err.Raise vbObjectError + 100, , "Enter the iCharts ID in Settings B2 and the password in B3."
    If symbol = "" Then symbol = "NIFTY"
    If strikeStep <= 0 Then strikeStep = 50
    If normalLookback <= 0 Then normalLookback = 14
    If wedLookback <= 0 Then wedLookback = 7
    If IsDate(wsCfg.Range("B7").Value) Then
        expiryDate = DateValue(wsCfg.Range("B7").Value)
    ElseIf Trim$(CStr(wsCfg.Range("B7").Value2)) = "" Then
        expiryDate = AutomaticNiftyExpiry(Date)
    Else
        Err.Raise vbObjectError + 101, , "Enter a valid expiry date on Settings, or leave it blank for automatic Tuesday expiry."
    End If
    expiryCode = IChartsExpiryCode(expiryDate)

    EnsureFolder downloadFolder

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    wsCfg.Range("B17").Value = "Logging in to iCharts..."

    Set gCookies = CreateObject("Scripting.Dictionary")
    gCookies.CompareMode = vbTextCompare
    gLastHtml = HttpRequest("GET", LOGIN_URL, "")
    gLastHtml = HttpRequest("POST", LOGIN_URL, _
        "username=" & UrlEncode(userId) & "&password=" & UrlEncode(password) & _
        "&signnedWithGoogle=")

    If Not LoginSucceeded(gLastHtml) Then
        Err.Raise vbObjectError + 102, , "iCharts login failed. Check B2/B3, enable macros, and confirm the subscription is active." & vbCrLf & LoginHint(gLastHtml)
    End If

    If InStr(1, gLastHtml, "sessionID", vbTextCompare) = 0 Or InStr(1, gLastHtml, "latestDate", vbTextCompare) = 0 Then
        gLastHtml = HttpRequest("GET", CHAIN_URL, "")
    End If
    gSessionId = HtmlInputValue(gLastHtml, "sessionID")
    gUserName = HtmlInputValue(gLastHtml, "userName")
    If gSessionId = "" Or gUserName = "" Then
        Err.Raise vbObjectError + 105, , "Logged in, but iCharts did not return a session id. Try again after opening the option-chain page in a browser once."
    End If

    If IsDate(wsCfg.Range("B10").Value) Then
        asOfDate = DateValue(wsCfg.Range("B10").Value)
    Else
        latestDateText = HtmlInputValue(gLastHtml, "latestDate")
        If IsDate(latestDateText) Then
            asOfDate = DateValue(latestDateText)
        Else
            asOfDate = PreviousWeekday(Date)
        End If
    End If
    ' Never reach into expiry day: the model ends on the Monday before it.
    If asOfDate >= expiryDate Then asOfDate = DateAdd("d", -1, expiryDate)

    lookbackDays = normalLookback
    If Weekday(Date, vbMonday) = 3 And Date > expiryDate Then lookbackDays = wedLookback

    ' Start from an empty folder so nothing from an earlier run or expiry is reused.
    removedFiles = CleanDownloadFolder(downloadFolder, symbol)
    ClearRawSheets
    offsets = DayOffsets()
    sheetNames = DaySheetNames()

    For i = LBound(offsets) To UBound(offsets)
        tradeDate = DateAdd("d", CLng(offsets(i)), expiryDate)
        If tradeDate >= DateAdd("d", -lookbackDays, expiryDate) _
           And tradeDate <= asOfDate _
           And Weekday(tradeDate, vbMonday) <= 5 Then
            If IsHoliday(tradeDate) Then
                skippedDays = skippedDays + 1
            Else
                wsCfg.Range("B17").Value = "Downloading " & Format$(tradeDate, "dd-mmm-yyyy") & "..."
                DoEvents
                filePath = DownloadDayFile(symbol, expiryCode, tradeDate, downloadFolder)
                If filePath = "" Then
                    ' No EOD chain for this date: market holiday. Remember it and move on.
                    RecordHoliday tradeDate
                    skippedDays = skippedDays + 1
                Else
                    importedRows = ImportFromFile(filePath, ThisWorkbook.Worksheets(CStr(sheetNames(i))))
                    If importedRows > 0 Then
                        importedDays = importedDays + 1
                        If firstLoaded = 0 Then firstLoaded = tradeDate
                    End If
                End If
            End If
        End If
    Next i

    If importedDays = 0 Then
        Err.Raise vbObjectError + 103, , "No historical option rows were returned. Check expiry and as-of dates."
    End If

    RebuildMoneyGrid strikeStep
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    FillTop10Avwap
    If firstLoaded = 0 Then firstLoaded = DateAdd("d", -lookbackDays, expiryDate)
    UpdateExpirySummary expiryDate, firstLoaded, asOfDate, importedDays
    Application.CalculateFull
    UpdateBigMoneyAnalysis

    wsCfg.Range("B16").Value = Now
    wsCfg.Range("B17").Value = "Completed: " & importedDays & " EOD sheet(s) loaded, " & _
                               skippedDays & " holiday(s) skipped, " & _
                               removedFiles & " old file(s) deleted."
    If Not gSilent Then
        MsgBox "EOD refresh complete." & vbCrLf & _
               importedDays & " trading day(s) loaded through " & Format$(asOfDate, "dd-mmm-yyyy") & "." & vbCrLf & _
               skippedDays & " holiday(s) skipped." & vbCrLf & _
               removedFiles & " file(s) from earlier runs deleted before downloading." & vbCrLf & vbCrLf & _
               "Downloaded files: " & downloadFolder, _
               vbInformation, "NIFTY Option Automation"
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If oldCalc <> 0 Then Application.Calculation = oldCalc
    Exit Sub

Failed:
    errMsg = Trim$(Err.Description)
    If errMsg = "" Then errMsg = "Excel error " & CStr(Err.Number) & " (no description)."
    On Error Resume Next
    If Not wsCfg Is Nothing Then wsCfg.Range("B17").Value = Left$("ERROR: " & errMsg, 255)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If oldCalc <> 0 Then Application.Calculation = oldCalc
    If Not gSilent Then MsgBox errMsg, vbCritical, "EOD refresh failed"
End Sub

Private Function SettingsPassword(ByVal rng As Range) As String
    Dim raw As Variant
    raw = rng.Value2
    If IsError(raw) Or IsEmpty(raw) Then
        SettingsPassword = ""
    Else
        SettingsPassword = CStr(raw)
    End If
End Function

Private Function LoginSucceeded(ByVal html As String) As Boolean
    If InStr(1, html, "Logout", vbTextCompare) > 0 Then
        LoginSucceeded = True
    ElseIf InStr(1, html, "name=""userName""", vbTextCompare) > 0 And InStr(1, html, "Enter your Login ID", vbTextCompare) = 0 Then
        LoginSucceeded = True
    ElseIf InStr(1, html, "OptionChain_Old.php", vbTextCompare) > 0 And InStr(1, html, "rdDataType", vbTextCompare) > 0 Then
        LoginSucceeded = True
    End If
End Function

Private Function LoginHint(ByVal html As String) As String
    Dim snippet As String
    snippet = CleanText(Left$(Replace(Replace(html, "<", " "), ">", " "), 180))
    If InStr(1, html, "Enter your Login ID", vbTextCompare) > 0 Then
        LoginHint = "The site returned the login page, so the ID/password were rejected or the session cookie was not kept."
    ElseIf InStr(1, html, "subscription", vbTextCompare) > 0 Then
        LoginHint = "iCharts mentioned subscription. The account may not have access to Option Chain."
    ElseIf snippet <> "" Then
        LoginHint = "Server said: " & snippet
    End If
End Function

Private Function GetHistoricalChain(ByVal symbol As String, ByVal expiryCode As String, _
                                    ByVal tradeDate As Date) As String
    Dim url As String, payload As String
    url = TABLE_URL & "?txtDate=" & FormatIsoDate(tradeDate) & _
          "&optSymbol=" & UrlEncode(symbol) & _
          "&optExpDate=" & UrlEncode(expiryCode) & _
          "&dType=hist&striketype=allstrikes&SymbType=FNO" & _
          "&sessionID=" & UrlEncode(gSessionId) & _
          "&userName=" & UrlEncode(gUserName)
    payload = HttpRequest("GET", url, "", 0, True)
    If InStr(1, payload, "Enter your Login ID", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 104, , "The iCharts session expired during download."
    End If
    If InStr(1, payload, "Invalid subscription", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 106, , "iCharts returned Invalid subscription for the option-chain table."
    End If
    If InStr(1, payload, """aaData""", vbTextCompare) = 0 Then
        ' A market holiday returns a page instead of chain JSON, so treat it as "no data".
        GetHistoricalChain = ""
        Exit Function
    End If
    GetHistoricalChain = payload
End Function

Private Function HttpRequest(ByVal method As String, ByVal url As String, ByVal body As String, _
                             Optional ByVal redirectCount As Long = 0, _
                             Optional ByVal asJson As Boolean = False) As String
    Dim req As Object, statusCode As Long, location As String
    If redirectCount > 5 Then Err.Raise vbObjectError + 110, , "Too many redirects from iCharts."

#If Mac Then
    Err.Raise vbObjectError + 141, , "The iCharts download needs Windows Excel." & vbCrLf & vbCrLf & _
        "Mac Excel has no WinHttp or Scripting.Dictionary, so the login and JSON download cannot run here. " & _
        "See the Notes sheet for the two supported ways to use this on a Mac."
#End If

    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Option(6) = False
    On Error Resume Next
    req.Option(9) = 2688
    On Error GoTo 0
    req.SetTimeouts 15000, 15000, 30000, 120000
    req.Open method, url, False
    req.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Excel-EOD-Automation"
    req.SetRequestHeader "Referer", CHAIN_URL
    If asJson Then
        req.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
        req.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    Else
        req.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    End If
    If Not gCookies Is Nothing Then
        If gCookies.Count > 0 Then req.SetRequestHeader "Cookie", CookieHeader()
    End If
    If UCase$(method) = "POST" Then
        req.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        req.SetRequestHeader "Origin", "https://www.icharts.in"
        req.Send body
    Else
        req.Send ""
    End If

    CaptureCookies req.GetAllResponseHeaders
    statusCode = req.Status
    If statusCode >= 300 And statusCode < 400 Then
        location = req.GetResponseHeader("Location")
        If location = "" Then Err.Raise vbObjectError + 111, , "iCharts returned a redirect without a location."
        If Left$(location, 1) = "/" Then location = "https://www.icharts.in" & location
        HttpRequest = HttpRequest("GET", location, "", redirectCount + 1, asJson)
    ElseIf statusCode >= 200 And statusCode < 300 Then
        HttpRequest = req.ResponseText
    Else
        Err.Raise vbObjectError + 112, , "iCharts HTTP error " & statusCode & ": " & req.StatusText
    End If
End Function

Private Sub CaptureCookies(ByVal headers As String)
    Dim lines As Variant, line As Variant, cookiePair As String
    Dim p As Long, eq As Long, cookieName As String, cookieValue As String
    If gCookies Is Nothing Then Set gCookies = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(headers, vbCr, ""), vbLf)
    For Each line In lines
        If LCase$(Left$(Trim$(CStr(line)), 11)) = "set-cookie:" Then
            cookiePair = Trim$(Mid$(Trim$(CStr(line)), 12))
            p = InStr(cookiePair, ";")
            If p > 0 Then cookiePair = Left$(cookiePair, p - 1)
            eq = InStr(cookiePair, "=")
            If eq > 1 Then
                cookieName = Trim$(Left$(cookiePair, eq - 1))
                cookieValue = Mid$(cookiePair, eq + 1)
                gCookies(cookieName) = cookieValue
            End If
        End If
    Next line
End Sub

Private Function CookieHeader() As String
    Dim key As Variant, result As String
    For Each key In gCookies.Keys
        If result <> "" Then result = result & "; "
        result = result & CStr(key) & "=" & CStr(gCookies(key))
    Next key
    CookieHeader = result
End Function

' Saves one trading day as its own option-chain file, the same content the
' website's Download CSV button produces. Returns "" when the day has no data.
Private Function DownloadDayFile(ByVal symbol As String, ByVal expiryCode As String, _
                                 ByVal tradeDate As Date, ByVal folder As String) As String
    Dim payload As String, rows As Collection, path As String

    payload = GetHistoricalChain(symbol, expiryCode, tradeDate)
    If payload = "" Then Exit Function

    Set rows = ParseAaData(payload)
    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then Exit Function

    path = PathJoin(folder, symbol & "_" & expiryCode & "_" & FormatIsoDate(tradeDate) & ".csv")
    If Not WriteChainCsv(rows, path) Then Exit Function
    DownloadDayFile = path
End Function

Private Function WriteChainCsv(ByVal rows As Collection, ByVal path As String) As Boolean
    Dim handle As Integer, i As Long, c As Long
    Dim cells As Variant, line As String, value As String, written As Long

    handle = FreeFile
    Open path For Output As #handle
    Print #handle, CsvHeaderLine()
    For i = 1 To rows.Count
        cells = rows(i)
        If IsArray(cells) Then
            If UBound(cells) >= 17 Then
                If IsNumeric(Replace(CleanText(StripTags(CStr(cells(17)))), ",", "")) Then
                    line = ""
                    For c = 0 To 36
                        value = ""
                        If c <= UBound(cells) Then
                            value = CleanText(Replace(StripTags(CStr(cells(c))), "_", " "))
                        End If
                        If c > 0 Then line = line & ","
                        line = line & CsvField(value)
                    Next c
                    Print #handle, line
                    written = written + 1
                End If
            End If
        End If
    Next i
    Close #handle
    WriteChainCsv = (written > 0)
End Function

Private Function CsvHeaderLine() As String
    Dim headers As Variant, i As Long, line As String
    headers = Split("Buildup|Trend|Time|Vega|Theta|Gamma|Delta|IV Chg|IV|OI Chg|OI|Volume|LTP Chg|LTP|VWAP|Bid|Ask|Strike Price|Bid|Ask|VWAP|LTP|LTP Chg|Volume|OI|OI Chg|IV|IV Chg|Delta|Gamma|Theta|Vega|PCR-OI|PCR-Vol|Time|Trend|Buildup", "|")
    For i = LBound(headers) To UBound(headers)
        If i > LBound(headers) Then line = line & ","
        line = line & """" & CStr(headers(i)) & """"
    Next i
    CsvHeaderLine = line
End Function

Private Function CsvField(ByVal value As String) As String
    Dim plain As String
    If value = "" Then Exit Function
    plain = Replace(Replace(value, ",", ""), ChrW(8722), "-")
    If IsNumeric(plain) Then
        CsvField = plain
    Else
        CsvField = """" & Replace(value, """", """""") & """"
    End If
End Function

' Opens the downloaded day file and copies it into the matching workbook sheet.
Private Function ImportFromFile(ByVal filePath As String, ByVal ws As Worksheet) As Long
    Dim wb As Workbook, src As Worksheet, lastRow As Long, copied As Long

    If filePath = "" Then Exit Function

#If Mac Then
    ' Mac Excel is sandboxed: a file outside its container needs explicit consent.
    If Not GrantAccessToMultipleFiles(Array(filePath)) Then
        Err.Raise vbObjectError + 142, , "Access to the day file was not granted:" & vbCrLf & filePath
    End If
#End If

    If Dir$(filePath) = "" Then Exit Function

    Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    Set src = wb.Worksheets(1)
    lastRow = src.Cells(src.Rows.Count, "R").End(xlUp).Row
    If lastRow > 1 Then
        ws.Cells.ClearContents
        ws.Range("A1").Resize(lastRow, 37).Value = src.Range("A1").Resize(lastRow, 37).Value
        copied = lastRow - 1
    End If
    wb.Close SaveChanges:=False

    If copied > 0 Then ApplyDayFormats ws, copied + 1
    ImportFromFile = copied
End Function

Private Sub ApplyDayFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    ws.Range("K2:L" & lastRow).NumberFormat = "#,##0"
    ws.Range("Y2:Z" & lastRow).NumberFormat = "#,##0"
    ws.Range("M2:P" & lastRow).NumberFormat = "0.00"
    ws.Range("S2:V" & lastRow).NumberFormat = "0.00"
    ws.Range("C2:C" & lastRow).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Range("AI2:AI" & lastRow).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Columns("A:AK").AutoFit
End Sub

' Windows uses "\" and macOS uses "/", so paths are always assembled here.
Private Function PathJoin(ByVal folder As String, ByVal leaf As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(folder, 1) = sep Then
        PathJoin = folder & leaf
    Else
        PathJoin = folder & sep & leaf
    End If
End Function

Private Sub EnsureFolder(ByVal folder As String)
    If folder = "" Then Exit Sub
    If Dir$(folder, vbDirectory) <> "" Then Exit Sub

    On Error GoTo CannotCreate
    MkDir folder
    Exit Sub

CannotCreate:
    Err.Raise vbObjectError + 140, , "Cannot create or open the download folder:" & vbCrLf & folder & vbCrLf & vbCrLf & _
        "Type a folder you can write to in Settings B15, using this machine's path style " & _
        "(" & Application.PathSeparator & " as the separator)."
End Sub

Private Function ResolveDownloadFolder(ByVal wsCfg As Worksheet) As String
    Dim folder As String
    folder = Trim$(CStr(wsCfg.Range("B15").Value2))
    If folder = "" Then folder = PathJoin(ThisWorkbook.Path, "iCharts Downloads")
    ResolveDownloadFolder = folder
End Function

' Deletes the day files this macro created, so every run downloads fresh copies.
' Only files matching SYMBOL_*.csv are touched; anything else in the folder is left alone.
Private Function CleanDownloadFolder(ByVal folder As String, ByVal symbol As String) As Long
    Dim names As Collection, fileName As String, i As Long, removed As Long

    If folder = "" Or symbol = "" Then Exit Function
    If Dir$(folder, vbDirectory) = "" Then Exit Function

    ' Collect first: deleting while Dir$ is still enumerating skips files.
    Set names = New Collection
    fileName = Dir$(PathJoin(folder, symbol & "_*.csv"))
    Do While fileName <> ""
        names.Add fileName
        fileName = Dir$
    Loop

    For i = 1 To names.Count
        On Error Resume Next
        Kill PathJoin(folder, CStr(names(i)))
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
        On Error GoTo 0
    Next i
    CleanDownloadFolder = removed
End Function

' Reads the official NSE equity-derivatives trading holidays and refreshes the
' Holidays sheet. NSE serves this list only to browser-like clients, so the call
' first loads the site to collect cookies, exactly as a browser would.
Public Sub ImportNseHolidays()
    Dim payload As String, added As Long, replaced As Long, errMsg As String

    On Error GoTo Failed
    Set gNseCookies = CreateObject("Scripting.Dictionary")
    gNseCookies.CompareMode = vbTextCompare

    payload = NseRequest(NSE_HOME_URL)
    payload = NseRequest(NSE_HOLIDAY_URL)
    added = ApplyNseHolidays(payload, replaced)

    SetHolidayStatus "Imported from NSE on " & Format$(Now, "dd-mmm-yyyy hh:mm") & _
                     ": " & added & " date(s) listed."
    If Not gSilent Then
        MsgBox "NSE trading holidays imported." & vbCrLf & _
               replaced & " previous NSE row(s) replaced." & vbCrLf & _
               added & " holiday date(s) now listed.", _
               vbInformation, "NSE Holidays"
    End If
    Exit Sub

Failed:
    errMsg = Trim$(Err.Description)
    SetHolidayStatus "NSE import failed on " & Format$(Now, "dd-mmm-yyyy hh:mm") & ": " & errMsg
    If Not gSilent Then
        MsgBox "Could not read the NSE holiday list." & vbCrLf & vbCrLf & errMsg & vbCrLf & vbCrLf & _
               "NSE sometimes blocks non-browser requests. The seeded list on this sheet stays usable," & _
               " and you can always type dates in column A by hand.", _
               vbExclamation, "NSE Holidays"
    End If
End Sub

' Same import without pop-ups, for scripted use.
Public Sub ImportNseHolidaysSilent()
    gSilent = True
    On Error Resume Next
    ImportNseHolidays
    On Error GoTo 0
    gSilent = False
End Sub

Private Sub SetHolidayStatus(ByVal message As String)
    Dim ws As Worksheet
    Set ws = HolidaySheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Range("E1").Value = Left$(message, 255)
    On Error GoTo 0
End Sub

Private Function NseRequest(ByVal url As String) As String
    Dim req As Object, statusCode As Long

    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Option(6) = True
    On Error Resume Next
    req.Option(9) = 2688
    On Error GoTo 0
    req.SetTimeouts 15000, 15000, 30000, 60000
    req.Open "GET", url, False
    req.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    req.SetRequestHeader "Accept", "*/*"
    req.SetRequestHeader "Accept-Language", "en-US,en;q=0.9"
    req.SetRequestHeader "Referer", "https://www.nseindia.com/resources/exchange-communication-holidays"
    If Not gNseCookies Is Nothing Then
        If gNseCookies.Count > 0 Then req.SetRequestHeader "Cookie", NseCookieHeader()
    End If
    req.Send ""

    CaptureNseCookies req.GetAllResponseHeaders
    statusCode = req.Status
    If statusCode < 200 Or statusCode >= 300 Then
        Err.Raise vbObjectError + 130, , "NSE returned HTTP " & statusCode & " for " & url
    End If
    NseRequest = req.ResponseText
End Function

Private Sub CaptureNseCookies(ByVal headers As String)
    Dim lines As Variant, line As Variant, pair As String, p As Long, eq As Long
    If gNseCookies Is Nothing Then Set gNseCookies = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(headers, vbCr, ""), vbLf)
    For Each line In lines
        If LCase$(Left$(Trim$(CStr(line)), 11)) = "set-cookie:" Then
            pair = Trim$(Mid$(Trim$(CStr(line)), 12))
            p = InStr(pair, ";")
            If p > 0 Then pair = Left$(pair, p - 1)
            eq = InStr(pair, "=")
            If eq > 1 Then gNseCookies(Trim$(Left$(pair, eq - 1))) = Mid$(pair, eq + 1)
        End If
    Next line
End Sub

Private Function NseCookieHeader() As String
    Dim key As Variant, result As String
    For Each key In gNseCookies.Keys
        If result <> "" Then result = result & "; "
        result = result & CStr(key) & "=" & CStr(gNseCookies(key))
    Next key
    NseCookieHeader = result
End Function

' Replaces previously imported NSE rows and keeps anything you added yourself.
Private Function ApplyNseHolidays(ByVal payload As String, ByRef replaced As Long) As Long
    Dim ws As Worksheet, segment As String, chunk As String
    Dim startPos As Long, endPos As Long, objEnd As Long
    Dim tradeDate As Date, description As String, rawDate As String
    Dim lastRow As Long, r As Long, nextRow As Long, listed As Long

    Set ws = HolidaySheet()
    If ws Is Nothing Then Err.Raise vbObjectError + 131, , "The Holidays sheet is missing."

    startPos = InStr(1, payload, """FO"":[", vbTextCompare)
    If startPos = 0 Then Err.Raise vbObjectError + 132, , "The response did not contain the equity-derivatives (FO) segment."
    startPos = InStr(startPos, payload, "[")
    endPos = InStr(startPos, payload, "]")
    If endPos <= startPos Then Err.Raise vbObjectError + 133, , "The equity-derivatives holiday list could not be read."
    segment = Mid$(payload, startPos + 1, endPos - startPos - 1)

    ' Drop rows from an earlier import so re-running does not duplicate them.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = lastRow To 2 Step -1
        If Left$(CStr(ws.Cells(r, "B").Value2), 4) = "NSE:" Then
            ws.Rows(r).Delete
            replaced = replaced + 1
        End If
    Next r

    startPos = 1
    Do
        startPos = InStr(startPos, segment, "{")
        If startPos = 0 Then Exit Do
        objEnd = InStr(startPos, segment, "}")
        If objEnd = 0 Then Exit Do
        chunk = Mid$(segment, startPos, objEnd - startPos + 1)
        rawDate = JsonStringValue(chunk, "tradingDate")
        description = JsonStringValue(chunk, "description")
        tradeDate = ParseNseDate(rawDate)
        If tradeDate <> 0 Then
            If Not IsHoliday(tradeDate) Then
                nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
                If nextRow < 2 Then nextRow = 2
                ws.Cells(nextRow, "A").Value = tradeDate
                ws.Cells(nextRow, "A").NumberFormat = "dd-mmm-yyyy"
                ws.Cells(nextRow, "B").Value = "NSE: " & description
            End If
        End If
        startPos = objEnd + 1
    Loop

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then listed = lastRow - 1
    ApplyNseHolidays = listed
End Function

Private Function JsonStringValue(ByVal chunk As String, ByVal key As String) As String
    Dim p As Long, valueEnd As Long
    p = InStr(1, chunk, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, chunk, """")
    If p = 0 Then Exit Function
    valueEnd = InStr(p + 1, chunk, """")
    If valueEnd <= p Then Exit Function
    JsonStringValue = Mid$(chunk, p + 1, valueEnd - p - 1)
End Function

' NSE sends dd-Mmm-yyyy. Parsed by hand so the result never depends on the
' machine's regional date settings.
Private Function ParseNseDate(ByVal text As String) As Date
    Dim parts As Variant, monthNames As Variant, i As Long, monthNo As Long
    If Trim$(text) = "" Then Exit Function
    parts = Split(Trim$(text), "-")
    If UBound(parts) <> 2 Then Exit Function
    monthNames = Array("jan", "feb", "mar", "apr", "may", "jun", _
                       "jul", "aug", "sep", "oct", "nov", "dec")
    For i = LBound(monthNames) To UBound(monthNames)
        If LCase$(Left$(CStr(parts(1)), 3)) = monthNames(i) Then monthNo = i + 1
    Next i
    If monthNo = 0 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(2)) Then Exit Function
    ParseNseDate = DateSerial(CLng(parts(2)), monthNo, CLng(parts(0)))
End Function

Private Function HolidaySheet() As Worksheet
    On Error Resume Next
    Set HolidaySheet = ThisWorkbook.Worksheets("Holidays")
End Function

Private Function IsHoliday(ByVal checkDate As Date) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, listed As Variant
    Set ws = HolidaySheet()
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        listed = ws.Cells(r, "A").Value
        If IsDate(listed) Then
            If DateValue(listed) = checkDate Then
                IsHoliday = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub RecordHoliday(ByVal holidayDate As Date)
    Dim ws As Worksheet, nextRow As Long
    Set ws = HolidaySheet()
    If ws Is Nothing Then Exit Sub
    If IsHoliday(holidayDate) Then Exit Sub
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    ws.Cells(nextRow, "A").Value = holidayDate
    ws.Cells(nextRow, "A").NumberFormat = "dd-mmm-yyyy"
    ws.Cells(nextRow, "B").Value = "Auto-detected: no EOD data on iCharts"
End Sub

' The 10 day sheets, oldest first. Expiry day itself is deliberately absent:
' its VWAP is distorted by settlement, so the model stops at the Monday before.
Private Function DaySheetNames() As Variant
    DaySheetNames = Array("Tue1", "Wed1", "Thu1", "Fri1", "Mon2", _
                          "Tue2", "Wed2", "Thu2", "Fri2", "Mon3")
End Function

' Offsets in calendar days from expiry, matching DaySheetNames position for position.
Private Function DayOffsets() As Variant
    DayOffsets = Array(-14, -13, -12, -11, -8, -7, -6, -5, -4, -1)
End Function

Private Sub ClearRawSheets(Optional ByVal wipeFormats As Boolean = False)
    Dim names As Variant, n As Variant, ws As Worksheet
    names = DaySheetNames()
    For Each n In names
        Set ws = ThisWorkbook.Worksheets(CStr(n))
        If wipeFormats Then
            ws.Cells.Clear
            ws.Cells.ColumnWidth = 8.43
        Else
            ws.Cells.ClearContents
        End If
    Next n
End Sub

Private Sub RebuildMoneyGrid(ByVal strikeStep As Long)
    Dim rawNames As Variant, n As Variant
    Dim s As Worksheet, lastRow As Long, r As Long, strikeValue As Variant
    Dim minStrike As Double, maxStrike As Double, countRows As Long

    minStrike = 1E+30
    maxStrike = -1E+30
    rawNames = DaySheetNames()
    For Each n In rawNames
        Set s = ThisWorkbook.Worksheets(CStr(n))
        lastRow = s.Cells(s.Rows.Count, "R").End(xlUp).Row
        For r = 2 To lastRow
            strikeValue = s.Cells(r, "R").Value2
            If IsNumeric(strikeValue) Then
                If CDbl(strikeValue) < minStrike Then minStrike = CDbl(strikeValue)
                If CDbl(strikeValue) > maxStrike Then maxStrike = CDbl(strikeValue)
            End If
        Next r
    Next n
    If minStrike > maxStrike Then Err.Raise vbObjectError + 120, , "No strikes were imported."

    minStrike = Int(minStrike / strikeStep) * strikeStep
    maxStrike = Int((maxStrike + strikeStep - 1) / strikeStep) * strikeStep
    countRows = CLng((maxStrike - minStrike) / strikeStep) + 1
    If countRows > 1000 Then Err.Raise vbObjectError + 121, , "Unexpected strike range returned by iCharts."

    WriteMoneyGrid minStrike, strikeStep, countRows
End Sub

' Rewrites the Money sheet for the given strike ladder, reusing the formulas that
' already sit on row 3, and repoints the Top 10 lookups at the new last row.
Private Sub WriteMoneyGrid(ByVal minStrike As Double, ByVal strikeStep As Long, ByVal countRows As Long)
    Dim ws As Worksheet, r As Long, outLast As Long, totalRow As Long
    Dim baseFormulas As Variant

    outLast = countRows + 2
    Set ws = ThisWorkbook.Worksheets("Money")
    baseFormulas = ws.Range("B3:AW3").FormulaR1C1
    ws.Range("A3:AW1200").ClearContents
    For r = 3 To outLast
        ws.Cells(r, "A").Value = minStrike + (r - 3) * strikeStep
        ws.Range("B" & r & ":AW" & r).FormulaR1C1 = baseFormulas
    Next r
    totalRow = outLast + 2
    ws.Cells(totalRow, "AU").Formula = "=SUM(AU3:AU" & outLast & ")"
    ws.Cells(totalRow, "AV").Formula = "=SUM(AV3:AV" & outLast & ")"

    UpdateTop10Ranges outLast, totalRow
End Sub

Private Sub UpdateTop10Ranges(ByVal moneyLastRow As Long, ByVal totalRow As Long)
    Dim ws As Worksheet, r As Long, rankNo As Long
    Set ws = ThisWorkbook.Worksheets("Top 10")
    For r = 3 To 12
        rankNo = r - 2
        ws.Cells(r, "A").Formula = "=VLOOKUP(LARGE(Money!$AU$3:$AU$" & moneyLastRow & "," & rankNo & "),Money!$AU$3:$AW$" & moneyLastRow & ",3,FALSE)"
        ws.Cells(r, "B").Formula = "=INDEX(Money!$AU$3:$AU$" & moneyLastRow & ",MATCH($A" & r & ",Money!$AW$3:$AW$" & moneyLastRow & ",0))/100"
        ws.Cells(r, "D").Formula = "=A" & r & "+C" & r
        ws.Cells(r, "E").Formula = "=A" & r
        ws.Cells(r, "G").Formula = "=VLOOKUP(LARGE(Money!$AV$3:$AV$" & moneyLastRow & "," & rankNo & "),Money!$AV$3:$AW$" & moneyLastRow & ",2,FALSE)"
        ws.Cells(r, "H").Formula = "=INDEX(Money!$AV$3:$AV$" & moneyLastRow & ",MATCH($G" & r & ",Money!$AW$3:$AW$" & moneyLastRow & ",0))/100"
        ws.Cells(r, "J").Formula = "=G" & r & "-I" & r
        ws.Cells(r, "K").Formula = "=G" & r
    Next r
    ws.Range("C3:C12,I3:I12").ClearContents
    ws.Range("A13").Formula = "=Money!AU" & totalRow
    ws.Range("B13").Formula = "=SUM(B3:B12)"
    ws.Range("G13").Formula = "=Money!AV" & totalRow
    ws.Range("H13").Formula = "=SUM(H3:H12)"
End Sub

Private Sub FillTop10Avwap()
    Dim ws As Worksheet, r As Long, strikeValue As Variant
    Set ws = ThisWorkbook.Worksheets("Top 10")
    For r = 3 To 12
        strikeValue = ws.Cells(r, "A").Value2
        If IsNumeric(strikeValue) Then ws.Cells(r, "C").Value = WeightedAvwap(CDbl(strikeValue), True)
        strikeValue = ws.Cells(r, "G").Value2
        If IsNumeric(strikeValue) Then ws.Cells(r, "I").Value = WeightedAvwap(CDbl(strikeValue), False)
    Next r
    ws.Range("C3:C12,I3:I12").NumberFormat = "0.00"
End Sub

Private Function WeightedAvwap(ByVal strikeValue As Double, ByVal isCall As Boolean) As Variant
    Dim names As Variant, n As Variant, ws As Worksheet, matchCell As Range
    Dim vwapCol As String, volumeCol As String, vw As Variant, vol As Variant
    Dim weightedSum As Double, totalVolume As Double

    names = DaySheetNames()
    If isCall Then
        vwapCol = "O": volumeCol = "L"
    Else
        vwapCol = "U": volumeCol = "X"
    End If
    For Each n In names
        Set ws = ThisWorkbook.Worksheets(CStr(n))
        Set matchCell = ws.Columns("R").Find(What:=strikeValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not matchCell Is Nothing Then
            vw = ws.Cells(matchCell.Row, vwapCol).Value2
            vol = ws.Cells(matchCell.Row, volumeCol).Value2
            If IsNumeric(vw) And IsNumeric(vol) Then
                If CDbl(vol) > 0 Then
                    weightedSum = weightedSum + CDbl(vw) * CDbl(vol)
                    totalVolume = totalVolume + CDbl(vol)
                End If
            End If
        End If
    Next n
    If totalVolume > 0 Then
        WeightedAvwap = Round(weightedSum / totalVolume, 2)
    Else
        WeightedAvwap = Empty
    End If
End Function

Private Sub UpdateExpirySummary(ByVal expiryDate As Date, ByVal dataFrom As Date, _
                                ByVal asOfDate As Date, ByVal importedDays As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    ws.Range("B2").Value = expiryDate
    ws.Range("B3").Value = dataFrom
    ws.Range("B4").Value = asOfDate
    ws.Range("B5").Value = importedDays
    ws.Range("B7").Formula = "='Top 10'!G3"
    ws.Range("B8").Formula = "='Top 10'!I3"
    ws.Range("B9").Formula = "='Top 10'!J3"
    ws.Range("B11").Formula = "='Top 10'!A3"
    ws.Range("B12").Formula = "='Top 10'!C3"
    ws.Range("B13").Formula = "='Top 10'!D3"
    ws.Range("B2:B4").NumberFormat = "dd-mmm-yyyy"
    ws.Range("B5").NumberFormat = "0"
    ws.Range("B7:B13").NumberFormat = "0.00"
End Sub

' Recalculates only the range model from the current Top 10 table.
' Useful after manually adjusting an AVWAP or coverage target.
Public Sub RecalculateRangeAnalysis()
    Application.CalculateFull
    UpdateBigMoneyAnalysis
End Sub

' Writer-side interpretation:
'   CE writer is protected when settlement <= CE BEP.
'   PE writer is protected when settlement >= PE BEP.
' The strict interval uses every Top-10 row. The defended interval instead
' protects the requested percentage of CE and PE money independently.
Private Sub UpdateBigMoneyAnalysis()
    Dim wsTop As Worksheet, wsOut As Worksheet, wsCfg As Worksheet
    Dim ceMoney(1 To 10) As Double, ceBep(1 To 10) As Double
    Dim peMoney(1 To 10) As Double, peBep(1 To 10) As Double
    Dim ceCount As Long, peCount As Long, r As Long
    Dim strictLow As Double, strictHigh As Double
    Dim defendedLow As Double, defendedHigh As Double
    Dim peCoverage As Double, ceCoverage As Double, target As Double
    Dim pin As Double, zoneLow As Double, zoneHigh As Double
    Dim strikeStep As Long

    Set wsTop = ThisWorkbook.Worksheets("Top 10")
    Set wsOut = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    Set wsCfg = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    For r = 3 To 12
        If IsNumeric(wsTop.Cells(r, "B").Value2) And IsNumeric(wsTop.Cells(r, "D").Value2) Then
            If CDbl(wsTop.Cells(r, "B").Value2) > 0 Then
                ceCount = ceCount + 1
                ceMoney(ceCount) = CDbl(wsTop.Cells(r, "B").Value2)
                ceBep(ceCount) = CDbl(wsTop.Cells(r, "D").Value2)
            End If
        End If
        If IsNumeric(wsTop.Cells(r, "H").Value2) And IsNumeric(wsTop.Cells(r, "J").Value2) Then
            If CDbl(wsTop.Cells(r, "H").Value2) > 0 Then
                peCount = peCount + 1
                peMoney(peCount) = CDbl(wsTop.Cells(r, "H").Value2)
                peBep(peCount) = CDbl(wsTop.Cells(r, "J").Value2)
            End If
        End If
    Next r

    If ceCount = 0 Or peCount = 0 Then
        wsOut.Range("B15:B30").ClearContents
        wsOut.Range("B17").Value = "Insufficient Top 10 data"
        On Error Resume Next
        ThisWorkbook.Worksheets(TRADE_SHEET).Range("B3:B27").ClearContents
        On Error GoTo 0
        Exit Sub
    End If

    target = 0.8
    If IsNumeric(wsCfg.Range("B18").Value2) Then target = CDbl(wsCfg.Range("B18").Value2)
    If target > 1 Then target = target / 100
    If target < 0.5 Or target > 0.95 Then target = 0.8

    strikeStep = CLng(Val(wsCfg.Range("B11").Value2))
    If strikeStep <= 0 Then strikeStep = 50

    strictLow = ArrayMaximum(peBep, peCount)
    strictHigh = ArrayMinimum(ceBep, ceCount)
    defendedLow = WeightedQuantile(peBep, peMoney, peCount, target)
    defendedHigh = WeightedQuantile(ceBep, ceMoney, ceCount, 1 - target)
    peCoverage = PutProtectedFraction(peBep, peMoney, peCount, defendedLow)
    ceCoverage = CallProtectedFraction(ceBep, ceMoney, ceCount, defendedHigh)

    pin = MinimumPainPin(ceBep, ceMoney, ceCount, peBep, peMoney, peCount, _
                         defendedLow, defendedHigh, strikeStep)
    zoneLow = CeilToStep(defendedLow, strikeStep)
    If zoneLow < pin - strikeStep Then zoneLow = pin - strikeStep
    zoneHigh = FloorToStep(defendedHigh, strikeStep)
    If zoneHigh > pin + strikeStep Then zoneHigh = pin + strikeStep
    If zoneLow > zoneHigh Then
        zoneLow = pin
        zoneHigh = pin
    End If

    wsOut.Range("B15").Value = strictLow
    wsOut.Range("B16").Value = strictHigh
    If strictLow <= strictHigh Then
        wsOut.Range("B17").Value = "YES: " & Format$(strictLow, "0") & " to " & Format$(strictHigh, "0")
    Else
        wsOut.Range("B17").Value = "NO OVERLAP: PE floor exceeds CE ceiling by " & _
                                   Format$(strictLow - strictHigh, "0.00")
    End If
    wsOut.Range("B19").Value = target
    wsOut.Range("B20").Value = defendedLow
    wsOut.Range("B21").Value = defendedHigh
    wsOut.Range("B22").Value = peCoverage
    wsOut.Range("B23").Value = ceCoverage
    wsOut.Range("B24").Value = Application.WorksheetFunction.Min(peCoverage, ceCoverage)
    wsOut.Range("B26").Value = pin
    wsOut.Range("B27").Value = zoneLow
    wsOut.Range("B28").Value = zoneHigh
    wsOut.Range("B29").Value = Format$(zoneLow, "0") & " to " & Format$(zoneHigh, "0")
    wsOut.Range("B30").Value = "Model estimate, not a guaranteed settlement range."

    wsOut.Range("B15:B16,B20:B21,B26:B28").NumberFormat = "0.00"
    wsOut.Range("B19,B22:B24").NumberFormat = "0.0%"

    UpdateTradePlan defendedLow, defendedHigh, pin, strikeStep, wsCfg
End Sub

' Turns the defended range into concrete legs: sell into the writer wall just
' outside each boundary, and buy a further-out strike as the hedge.
Private Sub UpdateTradePlan(ByVal defendedLow As Double, ByVal defendedHigh As Double, _
                            ByVal pin As Double, ByVal strikeStep As Long, ByVal wsCfg As Worksheet)
    Dim ws As Worksheet, hedgeSteps As Long, hedgeWidth As Double
    Dim sellCall As Double, sellPut As Double, hedgeCall As Double, hedgePut As Double
    Dim sellCallPrem As Double, sellPutPrem As Double
    Dim hedgeCallPrem As Double, hedgePutPrem As Double
    Dim callCredit As Double, putCredit As Double, totalCredit As Double
    Dim lowerBe As Double, upperBe As Double, verdict As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRADE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    hedgeSteps = 4
    If IsNumeric(wsCfg.Range("B19").Value2) Then hedgeSteps = CLng(Val(wsCfg.Range("B19").Value2))
    If hedgeSteps < 1 Or hedgeSteps > 20 Then hedgeSteps = 4
    hedgeWidth = hedgeSteps * strikeStep

    sellCall = ChooseSellStrike(True, CeilToStep(defendedHigh, strikeStep), strikeStep)
    sellPut = ChooseSellStrike(False, FloorToStep(defendedLow, strikeStep), strikeStep)
    hedgeCall = sellCall + hedgeWidth
    hedgePut = sellPut - hedgeWidth

    sellCallPrem = LatestPremium(sellCall, True)
    hedgeCallPrem = LatestPremium(hedgeCall, True)
    sellPutPrem = LatestPremium(sellPut, False)
    hedgePutPrem = LatestPremium(hedgePut, False)
    callCredit = sellCallPrem - hedgeCallPrem
    putCredit = sellPutPrem - hedgePutPrem
    totalCredit = callCredit + putCredit
    lowerBe = sellPut - totalCredit
    upperBe = sellCall + totalCredit

    If sellCallPrem = 0 Or hedgeCallPrem = 0 Or sellPutPrem = 0 Or hedgePutPrem = 0 Then
        verdict = "INCOMPLETE: no EOD premium found for at least one leg, so the credit figures are unreliable."
    ElseIf callCredit <= 0 Or putCredit <= 0 Then
        verdict = "CHECK DATA: one side prices as a debit, which should not happen for a further-out hedge."
    ElseIf upperBe >= defendedHigh And lowerBe <= defendedLow Then
        verdict = "OK: both break-evens sit outside the defended range."
    ElseIf upperBe < defendedHigh And lowerBe > defendedLow Then
        verdict = "WEAK: both break-evens fall inside the defended range. Widen the strikes or skip."
    ElseIf upperBe < defendedHigh Then
        verdict = "WATCH: upper break-even is inside the defended range."
    Else
        verdict = "WATCH: lower break-even is inside the defended range."
    End If

    ws.Range("B3").Value = sellCall
    ws.Range("B4").Value = MoneyAtStrike(sellCall, True)
    ws.Range("B5").Value = sellCallPrem
    ws.Range("B6").Value = hedgeCall
    ws.Range("B7").Value = hedgeCallPrem
    ws.Range("B8").Value = callCredit
    ws.Range("B9").Value = hedgeWidth - callCredit

    ws.Range("B12").Value = sellPut
    ws.Range("B13").Value = MoneyAtStrike(sellPut, False)
    ws.Range("B14").Value = sellPutPrem
    ws.Range("B15").Value = hedgePut
    ws.Range("B16").Value = hedgePutPrem
    ws.Range("B17").Value = putCredit
    ws.Range("B18").Value = hedgeWidth - putCredit

    ws.Range("B21").Value = totalCredit
    ws.Range("B22").Value = lowerBe
    ws.Range("B23").Value = upperBe
    ws.Range("B24").Value = pin
    ws.Range("B25").Value = verdict
    ws.Range("B26").Value = "Sell " & Format$(sellCall, "0") & " CE / buy " & Format$(hedgeCall, "0") & " CE" & _
                            "  +  sell " & Format$(sellPut, "0") & " PE / buy " & Format$(hedgePut, "0") & " PE"
    ws.Range("B27").Value = "Premiums are the latest EOD LTP in the loaded window, so live quotes will differ." & _
                            " Position sizing, margin and stop rules are yours to set."

    ws.Range("B3:B9,B12:B18,B21:B24").NumberFormat = "0.00"
End Sub

' Picks the strike carrying the most writer money within a few steps beyond the
' boundary: that concentration is the wall the big players defend.
Private Function ChooseSellStrike(ByVal isCall As Boolean, ByVal boundary As Double, _
                                  ByVal strikeStep As Long) As Double
    Dim i As Long, candidate As Double, money As Double
    Dim bestStrike As Double, bestMoney As Double

    bestStrike = boundary
    For i = 0 To 6
        If isCall Then
            candidate = boundary + i * strikeStep
        Else
            candidate = boundary - i * strikeStep
        End If
        money = MoneyAtStrike(candidate, isCall)
        If money > bestMoney Then
            bestMoney = money
            bestStrike = candidate
        End If
    Next i
    ChooseSellStrike = bestStrike
End Function

Private Function MoneyAtStrike(ByVal strikeValue As Double, ByVal isCall As Boolean) As Double
    Dim ws As Worksheet, lastRow As Long, matchCell As Range, col As String

    Set ws = ThisWorkbook.Worksheets("Money")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then Exit Function
    Set matchCell = ws.Range("A3:A" & lastRow).Find(What:=strikeValue, LookIn:=xlValues, LookAt:=xlWhole)
    If matchCell Is Nothing Then Exit Function
    If isCall Then col = "AU" Else col = "AV"
    If IsNumeric(ws.Cells(matchCell.Row, col).Value2) Then
        MoneyAtStrike = CDbl(ws.Cells(matchCell.Row, col).Value2)
    End If
End Function

' Most recent traded price for a strike, scanning the day sheets newest first.
Private Function LatestPremium(ByVal strikeValue As Double, ByVal isCall As Boolean) As Double
    Dim names As Variant, i As Long, ws As Worksheet, matchCell As Range
    Dim ltpCol As String, value As Variant

    names = DaySheetNames()
    If isCall Then ltpCol = "N" Else ltpCol = "V"

    For i = UBound(names) To LBound(names) Step -1
        Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
        Set matchCell = ws.Columns("R").Find(What:=strikeValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not matchCell Is Nothing Then
            value = ws.Cells(matchCell.Row, ltpCol).Value2
            If IsNumeric(value) Then
                If CDbl(value) > 0 Then
                    LatestPremium = CDbl(value)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function WeightedQuantile(ByRef values() As Double, ByRef weights() As Double, _
                                  ByVal itemCount As Long, ByVal probability As Double) As Double
    Dim sortedValues() As Double, sortedWeights() As Double
    Dim i As Long, j As Long, total As Double, running As Double
    Dim swapValue As Double

    ReDim sortedValues(1 To itemCount)
    ReDim sortedWeights(1 To itemCount)
    For i = 1 To itemCount
        sortedValues(i) = values(i)
        sortedWeights(i) = weights(i)
        total = total + weights(i)
    Next i
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If sortedValues(j) < sortedValues(i) Then
                swapValue = sortedValues(i): sortedValues(i) = sortedValues(j): sortedValues(j) = swapValue
                swapValue = sortedWeights(i): sortedWeights(i) = sortedWeights(j): sortedWeights(j) = swapValue
            End If
        Next j
    Next i
    For i = 1 To itemCount
        running = running + sortedWeights(i)
        If running >= probability * total Then
            WeightedQuantile = sortedValues(i)
            Exit Function
        End If
    Next i
    WeightedQuantile = sortedValues(itemCount)
End Function

Private Function PutProtectedFraction(ByRef beps() As Double, ByRef weights() As Double, _
                                      ByVal itemCount As Long, ByVal settlement As Double) As Double
    Dim i As Long, total As Double, protected As Double
    For i = 1 To itemCount
        total = total + weights(i)
        If beps(i) <= settlement Then protected = protected + weights(i)
    Next i
    If total > 0 Then PutProtectedFraction = protected / total
End Function

Private Function CallProtectedFraction(ByRef beps() As Double, ByRef weights() As Double, _
                                       ByVal itemCount As Long, ByVal settlement As Double) As Double
    Dim i As Long, total As Double, protected As Double
    For i = 1 To itemCount
        total = total + weights(i)
        If beps(i) >= settlement Then protected = protected + weights(i)
    Next i
    If total > 0 Then CallProtectedFraction = protected / total
End Function

Private Function MinimumPainPin(ByRef ceBep() As Double, ByRef ceMoney() As Double, ByVal ceCount As Long, _
                                ByRef peBep() As Double, ByRef peMoney() As Double, ByVal peCount As Long, _
                                ByVal rangeLow As Double, ByVal rangeHigh As Double, _
                                ByVal strikeStep As Long) As Double
    Dim candidate As Double, firstCandidate As Double, lastCandidate As Double
    Dim pain As Double, bestPain As Double, i As Long

    firstCandidate = CeilToStep(rangeLow, strikeStep)
    lastCandidate = FloorToStep(rangeHigh, strikeStep)
    If firstCandidate > lastCandidate Then
        MinimumPainPin = RoundToStep((rangeLow + rangeHigh) / 2, strikeStep)
        Exit Function
    End If

    bestPain = 1E+308
    candidate = firstCandidate
    Do While candidate <= lastCandidate
        pain = 0
        For i = 1 To ceCount
            If candidate > ceBep(i) Then pain = pain + ceMoney(i) * (candidate - ceBep(i))
        Next i
        For i = 1 To peCount
            If candidate < peBep(i) Then pain = pain + peMoney(i) * (peBep(i) - candidate)
        Next i
        If pain < bestPain Then
            bestPain = pain
            MinimumPainPin = candidate
        End If
        candidate = candidate + strikeStep
    Loop
End Function

Private Function ArrayMinimum(ByRef values() As Double, ByVal itemCount As Long) As Double
    Dim i As Long
    ArrayMinimum = values(1)
    For i = 2 To itemCount
        If values(i) < ArrayMinimum Then ArrayMinimum = values(i)
    Next i
End Function

Private Function ArrayMaximum(ByRef values() As Double, ByVal itemCount As Long) As Double
    Dim i As Long
    ArrayMaximum = values(1)
    For i = 2 To itemCount
        If values(i) > ArrayMaximum Then ArrayMaximum = values(i)
    Next i
End Function

Private Function RoundToStep(ByVal value As Double, ByVal strikeStep As Long) As Double
    RoundToStep = Int(value / strikeStep + 0.5) * strikeStep
End Function

Private Function CeilToStep(ByVal value As Double, ByVal strikeStep As Long) As Double
    CeilToStep = -Int(-value / strikeStep) * strikeStep
End Function

Private Function FloorToStep(ByVal value As Double, ByVal strikeStep As Long) As Double
    FloorToStep = Int(value / strikeStep) * strikeStep
End Function

Private Function HtmlInputValue(ByVal html As String, ByVal inputId As String) As String
    Dim p As Long, chunk As String, vPos As Long, q As String, vEnd As Long
    p = InStr(1, html, "id=""" & inputId & """", vbTextCompare)
    If p = 0 Then p = InStr(1, html, "id='" & inputId & "'", vbTextCompare)
    If p = 0 Then p = InStr(1, html, "name=""" & inputId & """", vbTextCompare)
    If p = 0 Then p = InStr(1, html, "name='" & inputId & "'", vbTextCompare)
    If p = 0 Then Exit Function
    If p > 160 Then
        chunk = Mid$(html, p - 160, 480)
    Else
        chunk = Mid$(html, 1, p + 320)
    End If
    vPos = InStr(1, chunk, "value=""", vbTextCompare)
    If vPos = 0 Then
        vPos = InStr(1, chunk, "value='", vbTextCompare)
        If vPos = 0 Then Exit Function
        q = "'"
    Else
        q = """"
    End If
    vPos = vPos + 7
    vEnd = InStr(vPos, chunk, q)
    If vEnd > vPos Then HtmlInputValue = Mid$(chunk, vPos, vEnd - vPos)
End Function

Private Function ParseAaData(ByVal json As String) As Collection
    Dim startPos As Long, i As Long, ch As String, rows As Collection
    Dim rowValues As Collection, token As String, inString As Boolean, escaped As Boolean
    Dim depth As Long

    startPos = InStr(1, json, """aaData""", vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, json, "[")
    If startPos = 0 Then Exit Function

    Set rows = New Collection
    Set rowValues = Nothing
    inString = False
    escaped = False
    depth = 0
    token = ""

    For i = startPos To Len(json)
        ch = Mid$(json, i, 1)
        If inString Then
            If escaped Then
                token = token & ch
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            Else
                token = token & ch
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                    token = ""
                Case "["
                    depth = depth + 1
                    If depth = 2 Then
                        Set rowValues = New Collection
                        token = ""
                    End If
                Case ",", "]"
                    If depth = 2 Then
                        If Len(token) > 0 Or ch = "," Then
                            If Len(token) > 0 Then rowValues.Add JsonToken(token)
                            token = ""
                        End If
                    End If
                    If ch = "]" Then
                        If depth = 2 Then
                            rows.Add CollectionToArray(rowValues)
                            Set rowValues = Nothing
                        End If
                        depth = depth - 1
                        If depth = 0 Then Exit For
                    End If
                Case " ", vbTab, vbCr, vbLf
                Case Else
                    token = token & ch
            End Select
        End If
    Next i
    Set ParseAaData = rows
End Function

Private Function JsonToken(ByVal token As String) As String
    If token = "null" Then
        JsonToken = ""
    Else
        JsonToken = token
    End If
End Function

Private Function CollectionToArray(ByVal items As Collection) As Variant
    Dim arr() As String, i As Long
    If items Is Nothing Then
        CollectionToArray = Array()
        Exit Function
    End If
    If items.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    ReDim arr(0 To items.Count - 1)
    For i = 1 To items.Count
        arr(i - 1) = CStr(items(i))
    Next i
    CollectionToArray = arr
End Function

Private Function StripTags(ByVal html As String) As String
    Dim result As String, inTag As Boolean, i As Long, ch As String
    result = Replace(Replace(Replace(html, "<br>", " "), "<br/>", " "), "<br />", " ")
    result = Replace(result, "&nbsp;", " ")
    result = Replace(result, "&#45;", "-")
    For i = 1 To Len(result)
        ch = Mid$(result, i, 1)
        If ch = "<" Then
            inTag = True
        ElseIf ch = ">" Then
            inTag = False
        ElseIf Not inTag Then
            StripTags = StripTags & ch
        End If
    Next i
    StripTags = CleanText(StripTags)
End Function

Private Function ToCellValue(ByVal text As String) As Variant
    Dim candidate As String
    If text = "" Then
        ToCellValue = Empty
    ElseIf text = "-" Then
        ToCellValue = "-"
    Else
        candidate = Replace(text, ",", "")
        candidate = Replace(candidate, ChrW(8722), "-")
        If IsNumeric(candidate) Then
            ToCellValue = CDbl(candidate)
        Else
            ToCellValue = text
        End If
    End If
End Function

Private Function CleanText(ByVal value As String) As String
    Dim result As String
    result = Replace(value, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    result = Replace(result, ChrW(160), " ")
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    CleanText = Trim$(result)
End Function

Private Function UrlEncode(ByVal value As String) As String
    Dim i As Long, ch As Long, result As String, c As String
    For i = 1 To Len(value)
        c = Mid$(value, i, 1)
        ch = AscW(c)
        Select Case ch
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & c
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Right$("0" & Hex$(ch And &HFF), 2)
        End Select
    Next i
    UrlEncode = result
End Function

Private Function FormatIsoDate(ByVal value As Date) As String
    FormatIsoDate = Format$(Year(value), "0000") & "-" & Format$(Month(value), "00") & "-" & Format$(Day(value), "00")
End Function

Private Function IChartsExpiryCode(ByVal value As Date) As String
    Dim months As Variant
    months = Array("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", _
                   "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    IChartsExpiryCode = Format$(Day(value), "00") & months(Month(value)) & Right$(CStr(Year(value)), 2)
End Function

Private Function PreviousWeekday(ByVal value As Date) As Date
    Dim d As Date
    d = DateAdd("d", -1, value)
    Do While Weekday(d, vbMonday) > 5
        d = DateAdd("d", -1, d)
    Loop
    PreviousWeekday = d
End Function

Private Function AutomaticNiftyExpiry(ByVal runDate As Date) As Date
    Dim weekdayNo As Long, daysToTuesday As Long
    weekdayNo = Weekday(runDate, vbMonday)
    If weekdayNo = 3 Then
        AutomaticNiftyExpiry = DateAdd("d", -1, runDate)
    Else
        daysToTuesday = (2 - weekdayNo + 7) Mod 7
        AutomaticNiftyExpiry = DateAdd("d", daysToTuesday, runDate)
    End If
End Function
