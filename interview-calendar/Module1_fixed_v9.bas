Attribute VB_Name = "Module1"
Option Explicit

' =====================================================================
'  面談カレンダー マクロ（v9）
'
'  v8 からの変更点
'   1. 月別シートの描画元に「履歴ブロック（W列以降 1回目～10回目）」を追加。
'      O列が翌月の予定に置き換わっても、履歴に残っている当月の面談は
'      カレンダーに表示され続ける（過去日=グレー／未来日=緑）。
'      → 8月がまだ終わっていなくても、実施済みの人のO・P列を9月の予定へ
'         上書きしてよい（先に「履歴へ転記」しておくこと）。
'   2. 履歴からの描画は年も照合する（シート先頭の曜日から年を推定）。
'      同じ行にO列と同じ日付の履歴があっても二重表示しない。
'   3. 「シートへ反映」は、O列がそのシートの月でない行（＝履歴から描かれた
'      表示）を動かしてもO列を書き換えない。誤って過去の実績を壊さないため。
' =====================================================================

Private Const COLOR_CONFIRMED As Long = 13434828
Private Const COLOR_PENDING As Long = 10092543
Private Const COLOR_ALERT As Long = 13551615
Private Const COLOR_DOUBLE As Long = 10079487
Private Const COLOR_WEEKEND As Long = 14737632
Private Const COLOR_DONE As Long = 14277081
Private Const COLOR_PROPOSED As Long = 16772300   ' 水色 RGB(204,236,255) メール打診済み

' 履歴ブロック（フォローアップシート）の位置
Private Const HIST_FIRST_COL As Long = 23    ' W列 = 1回目の先頭
Private Const HIST_BLOCK_SIZE As Long = 4    ' 日程・時刻・面談形態・約定
Private Const HIST_BLOCK_COUNT As Long = 10  ' 1回目～10回目

Public Sub UpdateCalendar()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsCal As Worksheet
    Set wsCal = ActiveSheet
    If Not IsMonthSheet(wsCal.Name) Then
        MsgBox "Open a monthly calendar sheet first.", vbExclamation
        GoTo FinallyExit
    End If

    ClearAutoEntries wsCal
    ReflectCalendar wsCal
    InstallCalendarButtons

    MsgBox wsCal.Name & " calendar update is complete.", vbInformation

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Update error: " & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

Public Sub CreateMonthSheet()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim inputText As String
    inputText = InputBox("Enter year/month. Example: 2026/7 or 7", "Create month sheet")
    If Len(Trim$(inputText)) = 0 Then GoTo FinallyExit

    Dim targetYear As Long, targetMonth As Long
    ParseYearMonth inputText, targetYear, targetMonth
    If targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "Month must be 1 to 12.", vbExclamation
        GoTo FinallyExit
    End If

    Dim sheetName As String
    sheetName = CStr(targetMonth) & J(26376)
    If WorksheetExists(sheetName) Then
        MsgBox sheetName & " already exists.", vbInformation
        GoTo FinallyExit
    End If
    If Not WorksheetExists("6" & J(26376)) Then
        MsgBox "Template sheet 6 month was not found.", vbExclamation
        GoTo FinallyExit
    End If

    ThisWorkbook.Worksheets("6" & J(26376)).Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Dim wsNew As Worksheet
    Set wsNew = ActiveSheet
    wsNew.Name = sheetName

    SetupMonthLayout wsNew, targetYear, targetMonth
    ClearAutoEntries wsNew
    ClearCalendarContents wsNew
    InstallCalendarButtons

    wsNew.Activate
    MsgBox sheetName & " was created.", vbInformation

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Create month sheet error: " & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

Public Sub InstallCalendarButtons()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then
            AddOrResetButton ws, "btnCalendarUpdate", "Calendar Update", "UpdateCalendar", ws.Range("Q1").Left, 8, 105, 28
            AddOrResetButton ws, "btnCreateMonthSheet", "Create Month", "CreateMonthSheet", ws.Range("Q1").Left + 115, 8, 115, 28
            AddOrResetButton ws, "btnSyncBack", "シートへ反映", "SyncCalendarToSheet", ws.Range("Q1").Left + 360, 8, 110, 28
        End If
    Next ws
End Sub

Private Sub ReflectCalendar(ByVal wsCal As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = GetSourceSheet()
    If wsSrc Is Nothing Then
        MsgBox "Source sheet was not found. Check follow-up sheet date/time columns.", vbExclamation
        Exit Sub
    End If

    Dim monthNo As Long
    monthNo = MonthFromSheetName(wsCal.Name)

    Dim colEmp As Long, colName As Long, colDate As Long, colTime As Long, colFixed As Long
    Dim colType As Long, colDone As Long, colReminder As Long, colMail As Long
    colEmp = 3
    colName = 4
    colDate = 15     ' O列 面談日
    colTime = 16     ' P列 時間
    colFixed = 17    ' Q列 日時確定
    colType = 18     ' R列 形式
    colDone = 19     ' S列 約定
    colReminder = 21
    colMail = 14     ' N列 打診メール○

    Dim lastRow As Long
    lastRow = LastUsedRow(wsSrc, colDate, colTime, colEmp, colName, colReminder)

    Dim entries As Object, colors As Object, counts As Object
    Set entries = CreateObject("Scripting.Dictionary")
    Set colors = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")

    Dim adjustingCount As Long, confirmedCount As Long, mailedCount As Long, doneCount As Long
    Dim histDoneCount As Long
    Dim unplacedNames As String
    Dim r As Long

    ' このシートが何年の月かを推定（履歴の年違いを拾わないため）
    Dim sheetYear As Long
    sheetYear = GuessSheetYear(wsCal, monthNo)

    For r = 3 To lastRow
        Dim interviewDate As Variant, interviewTime As Variant
        interviewDate = CoerceDate(wsSrc.Cells(r, colDate).Value)
        interviewTime = wsSrc.Cells(r, colTime).Value

        ' この行でO列から描画した日付（履歴との二重表示を防ぐため覚えておく）
        Dim currentDate As Variant
        currentDate = Empty

        ' --- (1) O列・P列の予定 --------------------------------------------
        If IsDate(interviewDate) And Len(Trim$(CStr(interviewTime))) > 0 Then
            If Month(CDate(interviewDate)) = monthNo Then
                currentDate = DateValue(CDate(interviewDate))

                Dim fixedText As String, doneText As String, reminderText As String, mailedText As String
                fixedText = Trim$(CStr(wsSrc.Cells(r, colFixed).Value))
                doneText = Trim$(CStr(wsSrc.Cells(r, colDone).Value))
                reminderText = Trim$(CStr(wsSrc.Cells(r, colReminder).Value))
                mailedText = Trim$(CStr(wsSrc.Cells(r, colMail).Value))

                Dim isFixed As Boolean, isDone As Boolean, isMailed As Boolean
                isFixed = (fixedText = J(9675) Or fixedText = J(12295))
                isDone = (doneText = J(9675) Or doneText = J(12295) Or doneText = J(9651) Or doneText = J(9650) Or doneText = J(215))
                isMailed = (mailedText = J(9675) Or mailedText = J(12295))

                Dim fillColor As Long
                If isDone Then
                    doneCount = doneCount + 1
                    fillColor = COLOR_DONE
                ElseIf isFixed Then
                    confirmedCount = confirmedCount + 1
                    fillColor = COLOR_CONFIRMED
                ElseIf isMailed Then
                    mailedCount = mailedCount + 1
                    fillColor = COLOR_PROPOSED
                Else
                    adjustingCount = adjustingCount + 1
                    fillColor = COLOR_PENDING
                End If

                PlaceEntry wsCal, entries, counts, colors, _
                           CStr(wsSrc.Cells(r, colName).Value), _
                           Trim$(CStr(wsSrc.Cells(r, colEmp).Value)), _
                           CStr(wsSrc.Cells(r, colType).Value), _
                           CDate(interviewDate), interviewTime, fillColor, unplacedNames
            End If
        End If

        ' --- (2) 履歴ブロック（W列以降）にある当月の面談 ---------------------
        Dim b As Long, hc As Long
        Dim histDate As Variant, histTime As Variant, histColor As Long
        For b = 1 To HIST_BLOCK_COUNT
            hc = HIST_FIRST_COL + (b - 1) * HIST_BLOCK_SIZE
            histDate = CoerceDate(wsSrc.Cells(r, hc).Value)
            histTime = wsSrc.Cells(r, hc + 1).Value

            If IsDate(histDate) And Len(Trim$(CStr(histTime))) > 0 Then
                If Month(CDate(histDate)) = monthNo And Year(CDate(histDate)) = sheetYear Then
                    Dim skipHist As Boolean
                    skipHist = False
                    If IsDate(currentDate) Then
                        If DateValue(CDate(histDate)) = CDate(currentDate) Then skipHist = True
                    End If

                    If Not skipHist Then
                        If DateValue(CDate(histDate)) <= Date Then
                            histColor = COLOR_DONE
                            doneCount = doneCount + 1
                            histDoneCount = histDoneCount + 1
                        Else
                            histColor = COLOR_CONFIRMED
                            confirmedCount = confirmedCount + 1
                        End If

                        PlaceEntry wsCal, entries, counts, colors, _
                                   CStr(wsSrc.Cells(r, colName).Value), _
                                   Trim$(CStr(wsSrc.Cells(r, colEmp).Value)), _
                                   CStr(wsSrc.Cells(r, hc + 2).Value), _
                                   CDate(histDate), histTime, histColor, unplacedNames
                    End If
                End If
            End If
        Next b
    Next r

    Dim addr As Variant
    For Each addr In entries.Keys
        Dim target As Range
        Set target = wsCal.Range(CStr(addr))

        Dim originalPattern As Long, originalColor As Long
        originalPattern = target.Interior.Pattern
        originalColor = target.Interior.Color

        If Len(CStr(target.Value)) > 0 Then
            target.Value = CStr(target.Value) & vbLf & CStr(entries(addr))
        Else
            target.Value = CStr(entries(addr))
        End If
        target.WrapText = True

        If CLng(counts(addr)) >= 2 Then
            target.Interior.Color = COLOR_DOUBLE
        Else
            target.Interior.Color = CLng(colors(addr))
        End If

        AddLogEntry wsCal.Name, CStr(addr), CStr(entries(addr)), originalPattern, originalColor
    Next addr

    wsCal.Range("A1").Value = "仮日程(未打診) " & adjustingCount
    wsCal.Range("A2").Value = "打診済み " & mailedCount
    wsCal.Range("A3").Value = "日時確定 " & confirmedCount
    wsCal.Range("A4").Value = "面談済み " & doneCount & IIf(histDoneCount > 0, "（うち履歴 " & histDoneCount & "）", "")
    wsCal.Range("A1:A4").Font.Bold = True

    If Len(unplacedNames) > 0 Then
        MsgBox "カレンダーの枠（8:30～20:00）に該当せず配置できなかった予定があります。" & _
               "P列の時間を確認してください。" & unplacedNames, vbExclamation
    End If

    If entries.Count = 0 Then
        MsgBox "No target data was found.", vbInformation
    End If
End Sub

' カレンダーの該当枠に1件配置する（上段=氏名、下段=社員番号）
Private Sub PlaceEntry(ByVal wsCal As Worksheet, ByVal entries As Object, ByVal counts As Object, ByVal colors As Object, _
                       ByVal personName As String, ByVal empText As String, ByVal typeText As String, _
                       ByVal interviewDate As Date, ByVal interviewTime As Variant, ByVal fillColor As Long, _
                       ByRef unplacedNames As String)
    Dim dayCol As Long, timeRow As Long
    dayCol = FindDayColumn(wsCal, Day(interviewDate))
    timeRow = FindTimeRow(wsCal, interviewTime)

    If dayCol > 0 And timeRow > 0 Then
        Dim key As String, empKey As String
        key = wsCal.Cells(timeRow, dayCol).Address(False, False)
        empKey = wsCal.Cells(timeRow + 1, dayCol).Address(False, False)

        AddEntry entries, counts, key, BuildDisplayText(personName, typeText)
        AddEntry entries, counts, empKey, empText

        AddColor colors, key, fillColor
        AddColor colors, empKey, fillColor
    Else
        unplacedNames = unplacedNames & vbCrLf & "・" & Trim$(personName) & _
            "（" & Format(interviewDate, "m/d") & " " & Format(TimeAsDouble(interviewTime), "h:mm") & "）"
    End If
End Sub

' 月別シートが何年の月かを推定する
' （シート名に年がないため、行4の曜日と行5の「1日」から照合する）
Private Function GuessSheetYear(ByVal wsCal As Worksheet, ByVal monthNo As Long) As Long
    Dim baseYear As Long, diff As Long
    baseYear = Year(Date)
    diff = monthNo - Month(Date)
    If diff <= -6 Then baseYear = baseYear + 1
    If diff >= 7 Then baseYear = baseYear - 1

    GuessSheetYear = baseYear
    If monthNo < 1 Or monthNo > 12 Then Exit Function

    Dim firstCol As Long
    firstCol = FindDayColumn(wsCal, 1)
    If firstCol = 0 Then Exit Function

    Dim wd As String
    wd = Trim$(CStr(wsCal.Cells(4, firstCol).Value))
    If Len(wd) = 0 Then Exit Function

    ' 当年 → 前年 → 翌年 の順に、1日の曜日が一致する年を採用
    Dim k As Long, y As Long
    For k = 1 To 3
        y = baseYear + Choose(k, 0, -1, 1)
        If WeekdayNameJp(DateSerial(y, monthNo, 1)) = wd Then
            GuessSheetYear = y
            Exit Function
        End If
    Next k
End Function

Private Sub AddEntry(ByVal entries As Object, ByVal counts As Object, ByVal key As String, ByVal textValue As String)
    If entries.Exists(key) Then
        entries(key) = entries(key) & vbLf & textValue
        counts(key) = CLng(counts(key)) + 1
    Else
        entries.Add key, textValue
        counts.Add key, 1
    End If
End Sub

Private Sub AddColor(ByVal colors As Object, ByVal key As String, ByVal fillColor As Long)
    If colors.Exists(key) Then
        If fillColor = COLOR_DONE Then colors(key) = fillColor
    Else
        colors.Add key, fillColor
    End If
End Sub

Private Function GetSourceSheet() As Worksheet
    Dim ws As Worksheet, bestWs As Worksheet
    Dim bestCount As Long, cnt As Long
    For Each ws In ThisWorkbook.Worksheets
        If Not IsMonthSheet(ws.Name) Then
            cnt = CountDateTimeRows(ws, 15, 16, 3)
            If cnt > bestCount Then
                bestCount = cnt
                Set bestWs = ws
            End If
        End If
    Next ws
    Set GetSourceSheet = bestWs
End Function

Private Function CountDateTimeRows(ByVal ws As Worksheet, ByVal dateCol As Long, ByVal timeCol As Long, ByVal firstRow As Long) As Long
    Dim r As Long, lastRow As Long
    lastRow = LastUsedRow(ws, dateCol, timeCol)
    For r = firstRow To lastRow
        If IsDate(CoerceDate(ws.Cells(r, dateCol).Value)) And Len(Trim$(CStr(ws.Cells(r, timeCol).Value))) > 0 Then
            CountDateTimeRows = CountDateTimeRows + 1
        End If
    Next r
End Function

Private Function BuildDisplayText(ByVal personName As String, ByVal interviewType As String) As String
    personName = Trim$(personName)
    interviewType = Trim$(interviewType)
    If interviewType = "WEB" & J(38754, 35527) Then
        BuildDisplayText = personName & J(12373, 12435) & "WEB"
    ElseIf interviewType = J(38651, 35441) Then
        BuildDisplayText = personName & J(12373, 12435) & "TEL"
    ElseIf Len(interviewType) > 0 Then
        BuildDisplayText = personName & J(12373, 12435) & interviewType
    Else
        BuildDisplayText = personName & J(12373, 12435)
    End If
End Function

Private Sub ClearAutoEntries(ByVal wsCal As Worksheet)
    Dim wsLog As Worksheet
    Set wsLog = GetLogSheet()

    Dim lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastLog < 2 Then Exit Sub

    Dim rowsToDelete As Collection
    Set rowsToDelete = New Collection

    Dim i As Long
    For i = 2 To lastLog
        If CStr(wsLog.Cells(i, 1).Value) = wsCal.Name Then
            Dim addr As String
            addr = CStr(wsLog.Cells(i, 2).Value)
            If Len(addr) > 0 Then
                Dim target As Range
                Set target = wsCal.Range(addr)
                target.Value = RemoveAutoLines(CStr(target.Value), CStr(wsLog.Cells(i, 3).Value))
                target.Interior.Pattern = CLng(Val(wsLog.Cells(i, 4).Value))
                If target.Interior.Pattern <> xlNone Then
                    target.Interior.Color = CLng(Val(wsLog.Cells(i, 5).Value))
                End If
            End If
            rowsToDelete.Add i
        End If
    Next i

    For i = rowsToDelete.Count To 1 Step -1
        wsLog.Rows(CLng(rowsToDelete(i))).Delete
    Next i
End Sub

Private Function RemoveAutoLines(ByVal currentText As String, ByVal autoText As String) As String
    If Len(currentText) = 0 Or Len(autoText) = 0 Then
        RemoveAutoLines = currentText
        Exit Function
    End If

    Dim autoDict As Object
    Set autoDict = CreateObject("Scripting.Dictionary")

    Dim autoLines() As String, i As Long
    autoLines = Split(autoText, vbLf)
    For i = LBound(autoLines) To UBound(autoLines)
        autoDict(Trim$(autoLines(i))) = True
    Next i

    Dim result As String, line As Variant
    For Each line In Split(currentText, vbLf)
        If Not autoDict.Exists(Trim$(CStr(line))) Then
            If Len(result) > 0 Then result = result & vbLf
            result = result & CStr(line)
        End If
    Next line
    RemoveAutoLines = result
End Function

Private Sub AddLogEntry(ByVal sheetName As String, ByVal addr As String, ByVal autoText As String, ByVal patternValue As Long, ByVal colorValue As Long)
    Dim wsLog As Worksheet
    Set wsLog = GetLogSheet()

    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(nextRow, 1).Value = sheetName
    wsLog.Cells(nextRow, 2).Value = addr
    wsLog.Cells(nextRow, 3).Value = autoText
    wsLog.Cells(nextRow, 4).Value = patternValue
    wsLog.Cells(nextRow, 5).Value = colorValue
End Sub

Private Function GetLogSheet() As Worksheet
    Dim logName As String
    logName = "__CalendarAutoLog"
    If Not WorksheetExists(logName) Then
        Set GetLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetLogSheet.Name = logName
        GetLogSheet.Range("A1:E1").Value = Array("Sheet", "Address", "AutoText", "OriginalPattern", "OriginalColor")
        GetLogSheet.Visible = xlSheetVeryHidden
    Else
        Set GetLogSheet = ThisWorkbook.Worksheets(logName)
    End If
End Function

Private Sub SetupMonthLayout(ByVal ws As Worksheet, ByVal targetYear As Long, ByVal targetMonth As Long)
    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

    Dim c As Long, d As Long
    For c = 2 To 32
        d = c - 1
        If d <= daysInMonth Then
            ws.Cells(4, c).Value = WeekdayNameJp(DateSerial(targetYear, targetMonth, d))
            ws.Cells(5, c).Value = CStr(d) & J(26085)
            If ws.Cells(4, c).Value = J(22303) Or ws.Cells(4, c).Value = J(26085) Then
                ws.Range(ws.Cells(4, c), ws.Cells(54, c)).Interior.Color = COLOR_WEEKEND
            Else
                ws.Range(ws.Cells(4, c), ws.Cells(54, c)).Interior.Pattern = xlNone
            End If
        Else
            ws.Range(ws.Cells(4, c), ws.Cells(54, c)).ClearContents
            ws.Range(ws.Cells(4, c), ws.Cells(54, c)).Interior.Pattern = xlNone
        End If
    Next c

    ws.Range("A1").Value = "仮日程(未打診) 0"
    ws.Range("A2").Value = "打診済み 0"
    ws.Range("A3").Value = "日時確定 0"
    ws.Range("A4").Value = "面談済み 0"
End Sub

Private Sub ClearCalendarContents(ByVal ws As Worksheet)
    Dim lastRowTime As Long
    lastRowTime = GetLastTimeRow(ws)
    If lastRowTime < 6 Then lastRowTime = 54
    ws.Range(ws.Cells(6, 2), ws.Cells(lastRowTime, 32)).ClearContents
End Sub

Private Function WeekdayNameJp(ByVal d As Date) As String
    WeekdayNameJp = Choose(Weekday(d, vbSunday), J(26085), J(26376), J(28779), J(27700), J(26408), J(37329), J(22303))
End Function

Private Function LastUsedRow(ByVal ws As Worksheet, ParamArray cols() As Variant) As Long
    Dim i As Long, lr As Long, maxRow As Long
    For i = LBound(cols) To UBound(cols)
        lr = ws.Cells(ws.Rows.Count, CLng(cols(i))).End(xlUp).Row
        If lr > maxRow Then maxRow = lr
    Next i
    If maxRow < 2 Then maxRow = 2
    LastUsedRow = maxRow
End Function

Private Function FindDayColumn(ByVal ws As Worksheet, ByVal targetDay As Long) As Long
    Dim c As Long, cellDay As Long
    For c = 2 To 32
        cellDay = ExtractNumber(CStr(ws.Cells(5, c).Value))
        If cellDay = targetDay Then
            FindDayColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function FindTimeRow(ByVal ws As Worksheet, ByVal targetTime As Variant) As Long
    Dim targetValue As Double
    targetValue = TimeAsDouble(targetTime)
    If targetValue < 0 Then Exit Function

    Dim r As Long, cellValue As Double
    For r = 1 To GetLastTimeRow(ws)
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            cellValue = TimeAsDouble(ws.Cells(r, 1).Value)
            If cellValue >= 0 Then
                ' 枠の時刻から30分未満のずれは同じ枠に配置（例：11:40→11:30の枠）
                If targetValue >= cellValue - (1# / 1440#) And targetValue < cellValue + (29# / 1440#) Then
                    FindTimeRow = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function GetLastTimeRow(ByVal ws As Worksheet) As Long
    GetLastTimeRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If GetLastTimeRow < 6 Then GetLastTimeRow = 54
End Function

Private Function TimeAsDouble(ByVal v As Variant) As Double
    On Error GoTo BadValue
    If IsDate(v) Then
        TimeAsDouble = CDbl(TimeValue(CDate(v)))
    ElseIf IsNumeric(v) Then
        TimeAsDouble = CDbl(v) - Fix(CDbl(v))
    Else
        TimeAsDouble = CDbl(TimeValue(CStr(v)))
    End If
    Exit Function
BadValue:
    TimeAsDouble = -1
End Function

Private Function ExtractNumber(ByVal textValue As String) As Long
    Dim i As Long, ch As String, result As String
    textValue = ToHalfWidthNumber(textValue)
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then result = result & ch
    Next i
    If Len(result) > 0 Then ExtractNumber = CLng(result)
End Function

Private Function ToHalfWidthNumber(ByVal s As String) As String
    Dim i As Long
    For i = 0 To 9
        s = Replace(s, ChrW(65296 + i), CStr(i))
    Next i
    ToHalfWidthNumber = s
End Function

Private Function MonthFromSheetName(ByVal sheetName As String) As Long
    MonthFromSheetName = ExtractNumber(sheetName)
End Function

Private Function IsMonthSheet(ByVal sheetName As String) As Boolean
    Dim m As Long
    m = MonthFromSheetName(sheetName)
    IsMonthSheet = (m >= 1 And m <= 12 And InStr(sheetName, J(26376)) > 0)
End Function

Private Sub ParseYearMonth(ByVal inputText As String, ByRef targetYear As Long, ByRef targetMonth As Long)
    inputText = Replace(ToHalfWidthNumber(Trim$(inputText)), J(24180), "/")
    inputText = Replace(inputText, J(26376), "")
    inputText = Replace(inputText, "-", "/")
    If InStr(inputText, "/") > 0 Then
        targetYear = CLng(Split(inputText, "/")(0))
        targetMonth = CLng(Split(inputText, "/")(1))
    Else
        targetYear = Year(Date)
        targetMonth = CLng(inputText)
    End If
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub AddOrResetButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal captionText As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPos, topPos, widthVal, heightVal)
    btn.Name = shapeName
    btn.Caption = captionText
    btn.OnAction = macroName
End Sub

Private Function J(ParamArray codes() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(codes) To UBound(codes)
        s = s & ChrW(CLng(codes(i)))
    Next i
    J = s
End Function


' 日付書式が外れてシリアル値（例：46213）になっているセルも日付として扱う
Private Function CoerceDate(ByVal v As Variant) As Variant
    If IsDate(v) Then
        CoerceDate = CDate(v)
    ElseIf IsNumeric(v) Then
        If Val(v) >= 40000 And Val(v) <= 80000 Then
            CoerceDate = CDate(CLng(Int(Val(v))))
        Else
            CoerceDate = Empty
        End If
    Else
        CoerceDate = Empty
    End If
End Function

' =====================================================================
'  月締め：開いている月別シートを「YYYY年M月実績」として保存し、
'  フォローアップシートのO～S列のうち「締める月以前の日付の行だけ」を
'  クリアする（来月以降の予定が入っている行はそのまま残す）
' =====================================================================
Public Sub CloseMonth()
    On Error GoTo ErrHandler

    Dim wsCal As Worksheet
    Set wsCal = ActiveSheet
    If Not IsMonthSheet(wsCal.Name) Then
        MsgBox "月別シート（例：7月）を開いた状態で実行してください。", vbExclamation
        Exit Sub
    End If

    Dim monthNo As Long
    monthNo = MonthFromSheetName(wsCal.Name)

    Dim yearText As String
    yearText = InputBox("実績シートに付ける年を入力してください。", "月締め", CStr(Year(Date)))
    If Len(Trim$(yearText)) = 0 Then Exit Sub
    If Not IsNumeric(yearText) Then
        MsgBox "年は数値で入力してください。", vbExclamation
        Exit Sub
    End If

    Dim resultName As String
    resultName = CStr(CLng(yearText)) & "年" & CStr(monthNo) & "月実績"
    If WorksheetExists(resultName) Then
        MsgBox "「" & resultName & "」は既に存在します。", vbExclamation
        Exit Sub
    End If

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("フォローアップシート")

    Dim monthEnd As Date
    monthEnd = DateSerial(CLng(yearText), monthNo + 1, 1) - 1

    If MsgBox("「" & wsCal.Name & "」を締めます。" & vbCrLf & vbCrLf & _
              "1. カレンダーを最新の状態に更新" & vbCrLf & _
              "2. 実績シート「" & resultName & "」を作成" & vbCrLf & _
              "3. O～S列のうち面談日が " & Format(monthEnd, "yyyy/m/d") & " 以前の行だけをクリア" & vbCrLf & _
              "   （来月以降の予定が入っている行は残ります）" & vbCrLf & vbCrLf & _
              "よろしいですか？", vbYesNo + vbQuestion, "月締め") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. カレンダーを最新化
    ClearAutoEntries wsCal
    ReflectCalendar wsCal

    ' 2. 実績シートとしてコピー（ボタンなどの図形は削除）
    wsCal.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Dim wsNew As Worksheet
    Set wsNew = ActiveSheet
    wsNew.Name = resultName
    Dim shp As Shape
    For Each shp In wsNew.Shapes
        shp.Delete
    Next shp

    ' 3. 締める月以前の行だけクリア（来月以降の予定は残す）
    Dim lastRow As Long, r As Long, dv As Variant
    Dim clearedCount As Long, keptCount As Long
    lastRow = LastUsedRow(wsSrc, 3, 4, 15, 16)
    For r = 3 To lastRow
        dv = CoerceDate(wsSrc.Cells(r, 15).Value)
        If IsDate(dv) Then
            If CDate(dv) <= monthEnd Then
                wsSrc.Range(wsSrc.Cells(r, 15), wsSrc.Cells(r, 19)).ClearContents
                clearedCount = clearedCount + 1
            Else
                keptCount = keptCount + 1
            End If
        End If
    Next r

    wsCal.Activate
    MsgBox "月締めが完了しました。" & vbCrLf & _
           "実績シート「" & resultName & "」を作成しました。" & vbCrLf & _
           "クリアした行：" & clearedCount & " 行 ／ 来月以降の予定として残した行：" & keptCount & " 行", _
           vbInformation, "月締め"

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "月締めエラー: " & Err.Description, vbExclamation
    Resume FinallyExit
End Sub



' =====================================================================
'  カレンダー→フォローアップシートの逆反映
'  カレンダー上で氏名＋社員番号のセルを手動で別の枠に移動したあと、
'  このマクロ（「シートへ反映」ボタン）を実行すると、
'  社員番号の位置から日付・時間を読み取ってO・P列を更新し、
'  カレンダーを最新の状態に描き直します。
'  ※移動するときは氏名セルと下段の社員番号セルを2つセットで動かすこと
' =====================================================================
Public Sub SyncCalendarToSheet()
    On Error GoTo ErrHandler

    Dim wsCal As Worksheet
    Set wsCal = ActiveSheet
    If Not IsMonthSheet(wsCal.Name) Then
        MsgBox "月別シート（例：8月）を開いた状態で実行してください。", vbExclamation
        Exit Sub
    End If
    Dim wsSrc As Worksheet
    Set wsSrc = GetSourceSheet()
    If wsSrc Is Nothing Then
        MsgBox "フォローアップシートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim monthNo As Long, yearNo As Long
    monthNo = MonthFromSheetName(wsCal.Name)
    yearNo = GuessSheetYear(wsCal, monthNo)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 社員番号 → フォローアップシート行
    Dim rowByEmp As Object
    Set rowByEmp = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, k As String
    lastRow = LastUsedRow(wsSrc, 3, 4)
    Dim dupEmp As String, histLocked As String
    For r = 3 To lastRow
        k = Trim$(CStr(wsSrc.Cells(r, 3).Value))
        If Len(k) > 0 Then
            If Not rowByEmp.Exists(k) Then
                rowByEmp.Add k, r
            Else
                ' 同じ社員番号が複数行 → 面談日(O列)が入っている行を優先して更新対象にする
                If Not IsDate(CoerceDate(wsSrc.Cells(rowByEmp(k), 15).Value)) Then
                    If IsDate(CoerceDate(wsSrc.Cells(r, 15).Value)) Then rowByEmp(k) = r
                End If
                If InStr(dupEmp, k) = 0 Then
                    dupEmp = dupEmp & vbCrLf & "・" & Trim$(CStr(wsSrc.Cells(r, 4).Value)) & "（" & k & "）"
                End If
            End If
        End If
    Next r

    Dim changedCount As Long, changeList As String, notFound As String
    Dim lastT As Long, tr As Long, c As Long, i As Long
    Dim tv As Double, dayNo As Long, srcRow As Long
    Dim empLines As Variant, newD As Date
    Dim curD As Variant, curT As Double, needUpd As Boolean
    Dim oldTxt As String, fixedMark As String

    lastT = GetLastTimeRow(wsCal)
    For tr = 6 To lastT
        If Len(Trim$(CStr(wsCal.Cells(tr, 1).Value))) > 0 Then
            tv = TimeAsDouble(wsCal.Cells(tr, 1).Value)
            If tv >= 0 Then
                For c = 2 To 32
                    dayNo = ExtractNumber(CStr(wsCal.Cells(5, c).Value))
                    If dayNo > 0 And Len(Trim$(CStr(wsCal.Cells(tr + 1, c).Value))) > 0 Then
                        empLines = Split(CStr(wsCal.Cells(tr + 1, c).Value), vbLf)
                        For i = LBound(empLines) To UBound(empLines)
                            k = DigitsOf(CStr(empLines(i)))
                            If Len(k) >= 5 Then
                                If rowByEmp.Exists(k) Then
                                    srcRow = rowByEmp(k)
                                    newD = DateSerial(yearNo, monthNo, dayNo)
                                    curD = CoerceDate(wsSrc.Cells(srcRow, 15).Value)
                                    curT = TimeAsDouble(wsSrc.Cells(srcRow, 16).Value)

                                    ' O列がこのシートの月でない行（＝履歴ブロックから描かれた実績）は
                                    ' 動かしてもO列に書き戻さない。過去の実績や翌月の予定を壊さないため。
                                    Dim isPlanRow As Boolean
                                    isPlanRow = False
                                    If IsDate(curD) Then
                                        If Month(CDate(curD)) = monthNo Then isPlanRow = True
                                    End If

                                    needUpd = isPlanRow
                                    If Not isPlanRow Then
                                        If InStr(histLocked, k) = 0 Then
                                            histLocked = histLocked & vbCrLf & "・" & Trim$(CStr(wsSrc.Cells(srcRow, 4).Value)) & "（" & k & "）"
                                        End If
                                    ElseIf IsDate(curD) Then
                                        If DateValue(CDate(curD)) = newD And curT >= 0 Then
                                            ' 同じ30分枠内なら変更なし扱い（11:40入力が11:30枠に表示されるケースを保護）
                                            If Int(curT * 48 + 0.0001) = Int(tv * 48 + 0.0001) Then needUpd = False
                                        End If
                                    End If
                                    If needUpd Then
                                        If IsDate(curD) Then
                                            oldTxt = Format(CDate(curD), "m/d") & " " & Format(curT, "h:mm")
                                        Else
                                            oldTxt = "未設定"
                                        End If
                                        fixedMark = ""
                                        If Len(Trim$(CStr(wsSrc.Cells(srcRow, 17).Value))) > 0 Then fixedMark = "　※日時確定済み・要再連絡"
                                        wsSrc.Cells(srcRow, 15).Value = newD
                                        wsSrc.Cells(srcRow, 16).Value = tv
                                        changedCount = changedCount + 1
                                        changeList = changeList & vbCrLf & "・" & Trim$(CStr(wsSrc.Cells(srcRow, 4).Value)) & _
                                                     "：" & oldTxt & " → " & Format(newD, "m/d") & " " & Format(tv, "h:mm") & fixedMark
                                    End If
                                Else
                                    notFound = notFound & vbCrLf & "・" & CStr(empLines(i)) & "（社員番号がシートに見つかりません）"
                                End If
                            End If
                        Next i
                    End If
                Next c
            End If
        End If
    Next tr

    ' カレンダーを最新状態に描き直す（手動移動の跡・色ズレを正常化）
    ClearAutoEntries wsCal
    ClearEntryCells wsCal
    ClearLogForSheet wsCal.Name
    ReflectCalendar wsCal
    InstallCalendarButtons

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim msg As String
    If changedCount > 0 Then
        msg = "フォローアップシートに " & changedCount & " 件反映し、カレンダーを描き直しました。" & changeList
    Else
        msg = "変更はありませんでした（カレンダーとシートは一致しています）。"
    End If
    If notFound <> "" Then msg = msg & vbCrLf & vbCrLf & "【反映できなかったもの】" & notFound
    If histLocked <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【履歴の実績なので動かしませんでした】" & vbCrLf & _
              "履歴ブロックから表示している面談です。日程を変えるにはO・P列に入力し直してください。" & histLocked
    End If
    If dupEmp <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【注意】同じ社員番号の行がフォローアップシートに複数あります。" & vbCrLf & _
              "使わない方の行のO・P列を空にしてください（放置するとカレンダーに二重表示されます）。" & dupEmp
    End If
    MsgBox msg, vbInformation, "シートへ反映"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "反映エラー: " & Err.Description, vbExclamation
End Sub

' 文字列中の数字だけを取り出す（全角数字も対応）
Private Function DigitsOf(ByVal s As String) As String
    Dim i As Long, ch As String, resultTxt As String
    s = ToHalfWidthNumber(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then resultTxt = resultTxt & ch
    Next i
    DigitsOf = resultTxt
End Function

' カレンダー上の面談エントリらしきセル（氏名＋社員番号）をすべて消す
Private Sub ClearEntryCells(ByVal wsCal As Worksheet)
    Dim lastT As Long, r As Long, c As Long
    lastT = GetLastTimeRow(wsCal)
    For r = 6 To lastT
        If Len(Trim$(CStr(wsCal.Cells(r, 1).Value))) > 0 Then
            If TimeAsDouble(wsCal.Cells(r, 1).Value) >= 0 Then
                For c = 2 To 32
                    If LooksLikeEntry(wsCal, r, c) Then
                        wsCal.Cells(r, c).ClearContents
                        wsCal.Cells(r + 1, c).ClearContents
                        ResetCellFill wsCal, r, c
                        ResetCellFill wsCal, r + 1, c
                    End If
                Next c
            End If
        End If
    Next r
End Sub

Private Function LooksLikeEntry(ByVal wsCal As Worksheet, ByVal r As Long, ByVal c As Long) As Boolean
    If InStr(CStr(wsCal.Cells(r, c).Value), "さん") > 0 Then
        LooksLikeEntry = True
    ElseIf Len(DigitsOf(CStr(wsCal.Cells(r + 1, c).Value))) >= 5 Then
        LooksLikeEntry = True
    End If
End Function

' セルの塗りを既定に戻す（土日はグレー・平日は塗りなし）
Private Sub ResetCellFill(ByVal wsCal As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim wd As String
    wd = CStr(wsCal.Cells(4, c).Value)
    If wd = J(22303) Or wd = J(26085) Then
        wsCal.Cells(r, c).Interior.Color = COLOR_WEEKEND
    Else
        wsCal.Cells(r, c).Interior.Pattern = xlNone
    End If
End Sub

' 転記ログから指定シート分を削除
Private Sub ClearLogForSheet(ByVal sheetName As String)
    Dim wsLog As Worksheet
    Set wsLog = GetLogSheet()
    Dim i As Long, lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    For i = lastLog To 2 Step -1
        If CStr(wsLog.Cells(i, 1).Value) = sheetName Then wsLog.Rows(i).Delete
    Next i
End Sub
