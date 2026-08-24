Attribute VB_Name = "HistoryTransferModule"
Option Explicit

' =====================================================================
'  面談実績 → 履歴ブロック 自動転記マクロ
'
'  フォローアップシートの O～S列（当月の面談日・時間・形式・約定）を、
'  W列以降の履歴ブロック（1回目～10回目 ／ 1ブロック4列 =
'  日程・時刻・面談形態・約定）の空いている回次へまとめて転記します。
'
'  ■毎月の運用手順
'    1. 月別シート（例：8月）で「Calendar Update」→ カレンダーを最新化
'    2. 「履歴へ転記」（このマクロ）→ 当月分を履歴ブロックへ記録
'    3. 「月締め」→ 実績シートを作成し、O～S列をクリア
'    ※ 月締めを先に実行するとO～S列が消えて転記できなくなります。
'      必ず「履歴へ転記 → 月締め」の順で実行してください。
'
'  ■安全機構
'    ・同じ日付が既に履歴ブロックに入っている行は二重転記しません
'    ・書き込んだ内容は __HistoryTransferLog（非表示シート）に記録され、
'      UndoLastHistoryTransfer で直前の転記をまるごと取り消せます
'    ・書き込む前に「何件転記するか」を確認ダイアログで表示します
'
'  ■導入
'    Alt+F11 →「ファイル」→「ファイルのインポート」で本ファイルを取り込み、
'    Alt+F8 → SetupHistoryTransfer を実行（月別シートにボタンを設置）
' =====================================================================

Private Const SHEET_FOLLOW As String = "フォローアップシート"
Private Const LOG_SHEET As String = "__HistoryTransferLog"
Private Const DATA_START_ROW As Long = 3

Private Const COL_EMPNO As Long = 3      ' C列 社員番号
Private Const COL_NAME As Long = 4       ' D列 氏名
Private Const COL_COUNT As Long = 12     ' L列 面談回数
Private Const COL_DATE As Long = 15      ' O列 面談日
Private Const COL_TIME As Long = 16      ' P列 時間
Private Const COL_TYPE As Long = 18      ' R列 形式
Private Const COL_DONE As Long = 19      ' S列 約定

Private Const HIST_FIRST_COL As Long = 23   ' W列 = 1回目の先頭
Private Const HIST_BLOCK_SIZE As Long = 4   ' 日程・時刻・面談形態・約定
Private Const HIST_BLOCK_COUNT As Long = 10 ' 1回目～10回目

' True にすると、転記後に L列（面談回数）を「n回目」へ自動更新する
Private Const UPDATE_INTERVIEW_COUNT As Boolean = False
' True にすると、転記先の日付／時刻セルに表示形式（yyyy/m/d, h:mm）を設定する
Private Const FIX_DATE_FORMAT As Boolean = True
' 確認ダイアログに名前を並べる最大件数
Private Const MAX_LIST As Long = 8


' =====================================================================
'  初回だけ実行：月別シートに「履歴へ転記」ボタンを設置
' =====================================================================
Public Sub SetupHistoryTransfer()
    Dim ws As Worksheet
    Dim installed As Long

    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheetName(ws.Name) Then
            AddOrResetTransferButton ws, "btnHistoryTransfer", "履歴へ転記", _
                "TransferMonthToHistory", ws.Range("Q1").Left + 480, 8, 120, 28
            installed = installed + 1
        End If
    Next ws

    MsgBox "月別シート " & installed & " 枚に「履歴へ転記」ボタンを設置しました。" & vbCrLf & vbCrLf & _
           "毎月の順番：" & vbCrLf & _
           "  (1) Calendar Update　(2) 履歴へ転記　(3) 月締め", vbInformation, "設置完了"
End Sub


' =====================================================================
'  本体：指定した月の面談実績を履歴ブロックへ転記
' =====================================================================
Public Sub TransferMonthToHistory()
    On Error GoTo ErrHandler

    If Not SheetExists(SHEET_FOLLOW) Then
        MsgBox "「" & SHEET_FOLLOW & "」シートが見つかりません。", vbExclamation, "履歴へ転記"
        Exit Sub
    End If

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(SHEET_FOLLOW)

    ' --- 対象月の入力 ---------------------------------------------------
    Dim inputText As String
    inputText = InputBox("履歴に転記する月を入力してください。" & vbCrLf & _
                         "例：2026/8　または　8" & vbCrLf & vbCrLf & _
                         "※その月末以前の面談日（O列）が入っている行が対象です。" & vbCrLf & _
                         "※翌月以降の予定が入っている行は対象外です。", _
                         "面談実績を履歴へ転記", DefaultMonthText())
    If Len(Trim$(inputText)) = 0 Then Exit Sub

    Dim targetYear As Long, targetMonth As Long
    ParseMonthText inputText, targetYear, targetMonth
    If targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "月は1～12で入力してください。", vbExclamation, "履歴へ転記"
        Exit Sub
    End If
    If targetYear < 2000 Or targetYear > 2100 Then
        MsgBox "年の指定が正しくありません。", vbExclamation, "履歴へ転記"
        Exit Sub
    End If

    Dim monthEnd As Date
    monthEnd = DateSerial(targetYear, targetMonth + 1, 0)

    ' --- 転記範囲の選択 -------------------------------------------------
    Dim ans As VbMsgBoxResult
    ans = MsgBox("転記する行を選んでください。" & vbCrLf & vbCrLf & _
                 "［はい］　　実施済みの行だけ（S列にマークがある行）" & vbCrLf & _
                 "［いいえ］　" & Format$(monthEnd, "yyyy/m/d") & " 以前の面談日がある行すべて" & vbCrLf & _
                 "［キャンセル］中止" & vbCrLf & vbCrLf & _
                 "※月締めでクリアされるのは［いいえ］と同じ範囲です。" & vbCrLf & _
                 "　記録を残さず消えてしまう行をなくしたい場合は［いいえ］。", _
                 vbYesNoCancel + vbQuestion, "転記対象")
    If ans = vbCancel Then Exit Sub

    Dim doneOnly As Boolean
    doneOnly = (ans = vbYes)

    ' --- 1回目：下見（書き込みはしない） --------------------------------
    Dim lastRow As Long
    lastRow = LastDataRow(wsSrc)
    If lastRow < DATA_START_ROW Then
        MsgBox "データ行が見つかりませんでした。", vbInformation, "履歴へ転記"
        Exit Sub
    End If

    Dim planRows() As Long, planBlocks() As Long
    ReDim planRows(1 To lastRow)
    ReDim planBlocks(1 To lastRow)

    Dim planCount As Long, skipDup As Long, skipNotDone As Long
    Dim skipFuture As Long, skipFull As Long
    Dim fullNames As String, backNames As String, backCount As Long
    Dim r As Long

    For r = DATA_START_ROW To lastRow
        Dim dv As Variant
        dv = CoerceDateValue(wsSrc.Cells(r, COL_DATE).Value)

        If IsDate(dv) Then
            If CDate(dv) > monthEnd Then
                skipFuture = skipFuture + 1
            ElseIf doneOnly And Len(Trim$(CStr(wsSrc.Cells(r, COL_DONE).Value))) = 0 Then
                skipNotDone = skipNotDone + 1
            ElseIf AlreadyInHistory(wsSrc, r, CDate(dv)) Then
                skipDup = skipDup + 1
            Else
                Dim blk As Long
                blk = NextBlockIndex(wsSrc, r)
                If blk > HIST_BLOCK_COUNT Then
                    skipFull = skipFull + 1
                    If skipFull <= MAX_LIST Then
                        fullNames = fullNames & vbCrLf & "  ・" & RowLabel(wsSrc, r)
                    End If
                Else
                    Dim lastHist As Variant
                    lastHist = LastHistoryDate(wsSrc, r)
                    If IsDate(lastHist) Then
                        If CDate(dv) < CDate(lastHist) Then
                            backCount = backCount + 1
                            If backCount <= MAX_LIST Then
                                backNames = backNames & vbCrLf & "  ・" & RowLabel(wsSrc, r) & _
                                    "（履歴の最終 " & Format$(CDate(lastHist), "m/d") & _
                                    " → 今回 " & Format$(CDate(dv), "m/d") & "）"
                            End If
                        End If
                    End If

                    planCount = planCount + 1
                    planRows(planCount) = r
                    planBlocks(planCount) = blk
                End If
            End If
        End If
    Next r

    If planCount = 0 Then
        MsgBox "転記対象の行がありませんでした。" & vbCrLf & vbCrLf & _
               "・既に履歴へ登録済み：" & skipDup & " 行" & vbCrLf & _
               "・実施済みでないため除外：" & skipNotDone & " 行" & vbCrLf & _
               "・翌月以降の予定のため対象外：" & skipFuture & " 行", _
               vbInformation, "履歴へ転記"
        Exit Sub
    End If

    ' --- 確認 -----------------------------------------------------------
    Dim msg As String
    msg = Format$(monthEnd, "yyyy年m月") & "分（" & Format$(monthEnd, "yyyy/m/d") & " 以前）を履歴ブロックへ転記します。" & vbCrLf & vbCrLf & _
          "  転記する行　　　　　　：" & planCount & " 行" & vbCrLf & _
          "  既に履歴にある（重複）：" & skipDup & " 行 …スキップ" & vbCrLf & _
          "  実施済みでない　　　　：" & skipNotDone & " 行 …スキップ" & vbCrLf & _
          "  翌月以降の予定　　　　：" & skipFuture & " 行 …対象外" & vbCrLf & _
          "  履歴が10回目まで満杯　：" & skipFull & " 行 …スキップ" & fullNames
    If skipFull > MAX_LIST Then msg = msg & vbCrLf & "  　（ほか " & (skipFull - MAX_LIST) & " 名）"

    If backCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "※履歴の最終日より前の日付を追記する行が " & backCount & " 行あります。" & backNames
        If backCount > MAX_LIST Then msg = msg & vbCrLf & "  　（ほか " & (backCount - MAX_LIST) & " 名）"
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "書き込む内容：O列→日程 ／ P列→時刻 ／ R列→面談形態 ／ S列→約定" & vbCrLf & _
          "（転記後に取り消したい場合は UndoLastHistoryTransfer を実行）" & vbCrLf & vbCrLf & _
          "実行してよろしいですか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "履歴へ転記") <> vbYes Then Exit Sub

    ' --- 2回目：書き込み ------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsLog As Worksheet
    Set wsLog = GetTransferLog()

    Dim batchId As Long
    batchId = NextBatchId(wsLog) + 1

    Dim stamp As String
    stamp = Format$(Now, "yyyy/mm/dd hh:nn:ss")

    Dim scopeText As String
    scopeText = Format$(monthEnd, "yyyy年m月") & IIf(doneOnly, "（実施済みのみ）", "（月末以前すべて）")

    Dim i As Long, baseCol As Long, writtenRows As Long
    For i = 1 To planCount
        r = planRows(i)
        baseCol = HIST_FIRST_COL + (planBlocks(i) - 1) * HIST_BLOCK_SIZE

        WriteAndLog wsSrc, wsLog, batchId, stamp, scopeText, r, baseCol, _
                    CDate(CoerceDateValue(wsSrc.Cells(r, COL_DATE).Value)), "yyyy/m/d"
        WriteAndLog wsSrc, wsLog, batchId, stamp, scopeText, r, baseCol + 1, _
                    wsSrc.Cells(r, COL_TIME).Value, "h:mm"
        WriteAndLog wsSrc, wsLog, batchId, stamp, scopeText, r, baseCol + 2, _
                    wsSrc.Cells(r, COL_TYPE).Value, ""
        WriteAndLog wsSrc, wsLog, batchId, stamp, scopeText, r, baseCol + 3, _
                    wsSrc.Cells(r, COL_DONE).Value, ""

        If UPDATE_INTERVIEW_COUNT Then
            WriteAndLog wsSrc, wsLog, batchId, stamp, scopeText, r, COL_COUNT, _
                        CStr(planBlocks(i) + 1) & "回目", ""
        End If

        writtenRows = writtenRows + 1
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "転記が完了しました。" & vbCrLf & vbCrLf & _
           "  転記した行：" & writtenRows & " 行" & vbCrLf & _
           "  スキップ　：重複 " & skipDup & " 行 ／ 未実施 " & skipNotDone & " 行 ／ 満杯 " & skipFull & " 行" & vbCrLf & vbCrLf & _
           "内容を確認したら、月別シートで「月締め」を実行してください。" & vbCrLf & _
           "（取り消す場合は Alt+F8 → UndoLastHistoryTransfer）", _
           vbInformation, "履歴へ転記"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "転記エラー: " & Err.Description, vbExclamation, "履歴へ転記"
End Sub


' =====================================================================
'  直前の転記をまるごと取り消す
' =====================================================================
Public Sub UndoLastHistoryTransfer()
    On Error GoTo ErrHandler

    If Not SheetExists(LOG_SHEET) Then
        MsgBox "転記の記録がありません。", vbInformation, "転記の取り消し"
        Exit Sub
    End If

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET)

    Dim lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastLog < 2 Then
        MsgBox "転記の記録がありません。", vbInformation, "転記の取り消し"
        Exit Sub
    End If

    Dim batchId As Long
    batchId = NextBatchId(wsLog)

    Dim cellCount As Long, i As Long
    For i = 2 To lastLog
        If CLng(Val(wsLog.Cells(i, 1).Value)) = batchId Then cellCount = cellCount + 1
    Next i

    Dim stamp As String, scopeText As String
    For i = 2 To lastLog
        If CLng(Val(wsLog.Cells(i, 1).Value)) = batchId Then
            stamp = CStr(wsLog.Cells(i, 6).Value)
            scopeText = CStr(wsLog.Cells(i, 7).Value)
            Exit For
        End If
    Next i

    If MsgBox("直前の転記を取り消します。" & vbCrLf & vbCrLf & _
              "  実行日時：" & stamp & vbCrLf & _
              "  対象　　：" & scopeText & vbCrLf & _
              "  戻すセル：" & cellCount & " セル" & vbCrLf & vbCrLf & _
              "よろしいですか？", vbYesNo + vbExclamation, "転記の取り消し") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(SHEET_FOLLOW)

    Dim restored As Long
    For i = lastLog To 2 Step -1
        If CLng(Val(wsLog.Cells(i, 1).Value)) = batchId Then
            Dim r As Long, c As Long
            r = CLng(Val(wsLog.Cells(i, 2).Value))
            c = CLng(Val(wsLog.Cells(i, 3).Value))

            If Len(CStr(wsLog.Cells(i, 4).Value)) = 0 Then
                wsSrc.Cells(r, c).ClearContents
            Else
                wsSrc.Cells(r, c).Value = wsLog.Cells(i, 4).Value
            End If
            If Len(CStr(wsLog.Cells(i, 8).Value)) > 0 Then
                wsSrc.Cells(r, c).NumberFormatLocal = CStr(wsLog.Cells(i, 8).Value)
            End If

            wsLog.Rows(i).Delete
            restored = restored + 1
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox restored & " セルを元に戻しました。", vbInformation, "転記の取り消し"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "取り消しエラー: " & Err.Description, vbExclamation, "転記の取り消し"
End Sub


' =====================================================================
'  おまけ：O列と履歴ブロックの日付列にまとめて日付書式を設定する
'  （シリアル値のまま「46244」と表示されている列を読めるようにする）
' =====================================================================
Public Sub FixHistoryDateFormat()
    If Not SheetExists(SHEET_FOLLOW) Then Exit Sub

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(SHEET_FOLLOW)

    Dim lastRow As Long
    lastRow = LastDataRow(wsSrc)
    If lastRow < DATA_START_ROW Then Exit Sub

    If MsgBox("O列（面談日）と履歴ブロックの日程列に" & vbCrLf & _
              "日付表示形式（yyyy/m/d）を設定します。" & vbCrLf & vbCrLf & _
              "値は変更されません。よろしいですか？", _
              vbYesNo + vbQuestion, "日付書式の設定") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    wsSrc.Range(wsSrc.Cells(DATA_START_ROW, COL_DATE), wsSrc.Cells(lastRow, COL_DATE)).NumberFormatLocal = "yyyy/m/d"

    Dim b As Long, c As Long
    For b = 1 To HIST_BLOCK_COUNT
        c = HIST_FIRST_COL + (b - 1) * HIST_BLOCK_SIZE
        wsSrc.Range(wsSrc.Cells(DATA_START_ROW, c), wsSrc.Cells(lastRow, c)).NumberFormatLocal = "yyyy/m/d"
    Next b

    Application.ScreenUpdating = True
    MsgBox "日付書式を設定しました。", vbInformation, "日付書式の設定"
End Sub


' =====================================================================
'  内部処理
' =====================================================================

' 1セル書き込み＋取り消し用ログ記録
Private Sub WriteAndLog(ByVal wsSrc As Worksheet, ByVal wsLog As Worksheet, _
                        ByVal batchId As Long, ByVal stamp As String, ByVal scopeText As String, _
                        ByVal r As Long, ByVal c As Long, ByVal newValue As Variant, _
                        ByVal numFormat As String)
    Dim target As Range
    Set target = wsSrc.Cells(r, c)

    ' 空欄を書き込もうとした場合は何もしない（ログも残さない）
    If Not IsDate(newValue) Then
        If Len(Trim$(CStr(newValue))) = 0 Then Exit Sub
    End If

    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = batchId
    wsLog.Cells(nextRow, 2).Value = r
    wsLog.Cells(nextRow, 3).Value = c
    wsLog.Cells(nextRow, 4).Value = target.Value
    wsLog.Cells(nextRow, 5).Value = CStr(newValue)
    wsLog.Cells(nextRow, 6).Value = stamp
    wsLog.Cells(nextRow, 7).Value = scopeText
    wsLog.Cells(nextRow, 8).Value = target.NumberFormatLocal

    target.Value = newValue
    If FIX_DATE_FORMAT And Len(numFormat) > 0 And IsDate(newValue) Then
        target.NumberFormatLocal = numFormat
    End If
End Sub

' 同じ日付が既にどこかの履歴ブロックに入っているか
Private Function AlreadyInHistory(ByVal wsSrc As Worksheet, ByVal r As Long, ByVal d As Date) As Boolean
    Dim b As Long, c As Long, hv As Variant

    For b = 1 To HIST_BLOCK_COUNT
        c = HIST_FIRST_COL + (b - 1) * HIST_BLOCK_SIZE
        hv = CoerceDateValue(wsSrc.Cells(r, c).Value)
        If IsDate(hv) Then
            If CDate(hv) = d Then
                AlreadyInHistory = True
                Exit Function
            End If
        End If
    Next b
End Function

' 書き込み先の回次（最後に使われているブロックの次）
Private Function NextBlockIndex(ByVal wsSrc As Worksheet, ByVal r As Long) As Long
    Dim b As Long, c As Long, lastUsed As Long

    For b = 1 To HIST_BLOCK_COUNT
        c = HIST_FIRST_COL + (b - 1) * HIST_BLOCK_SIZE
        If Len(Trim$(CStr(wsSrc.Cells(r, c).Value))) > 0 Then lastUsed = b
    Next b

    NextBlockIndex = lastUsed + 1
End Function

' 履歴ブロックに入っている最後の日付
Private Function LastHistoryDate(ByVal wsSrc As Worksheet, ByVal r As Long) As Variant
    Dim b As Long, c As Long, hv As Variant

    LastHistoryDate = Empty
    For b = 1 To HIST_BLOCK_COUNT
        c = HIST_FIRST_COL + (b - 1) * HIST_BLOCK_SIZE
        hv = CoerceDateValue(wsSrc.Cells(r, c).Value)
        If IsDate(hv) Then LastHistoryDate = CDate(hv)
    Next b
End Function

Private Function RowLabel(ByVal wsSrc As Worksheet, ByVal r As Long) As String
    RowLabel = "行" & r & " " & Trim$(CStr(wsSrc.Cells(r, COL_NAME).Value)) & _
               "（" & Trim$(CStr(wsSrc.Cells(r, COL_EMPNO).Value)) & "）"
End Function

Private Function LastDataRow(ByVal wsSrc As Worksheet) As Long
    Dim a As Long, b As Long
    a = wsSrc.Cells(wsSrc.Rows.Count, COL_EMPNO).End(xlUp).Row
    b = wsSrc.Cells(wsSrc.Rows.Count, COL_NAME).End(xlUp).Row
    LastDataRow = IIf(a > b, a, b)
End Function

Private Function GetTransferLog() As Worksheet
    If SheetExists(LOG_SHEET) Then
        Set GetTransferLog = ThisWorkbook.Worksheets(LOG_SHEET)
        Exit Function
    End If

    Dim prevSheet As Object
    Set prevSheet = ActiveSheet

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = LOG_SHEET
    ws.Range("A1:H1").Value = Array("BatchId", "Row", "Col", "OldValue", "NewValue", "DateTime", "Scope", "OldFormat")

    ' 追加直後は新しいシートがアクティブになるため、元のシートへ戻してから隠す
    On Error Resume Next
    prevSheet.Activate
    On Error GoTo 0
    ws.Visible = xlSheetVeryHidden

    Set GetTransferLog = ws
End Function

Private Function NextBatchId(ByVal wsLog As Worksheet) As Long
    Dim lastLog As Long, i As Long, maxId As Long

    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastLog
        If CLng(Val(wsLog.Cells(i, 1).Value)) > maxId Then maxId = CLng(Val(wsLog.Cells(i, 1).Value))
    Next i

    NextBatchId = maxId
End Function

' 日付書式が外れてシリアル値（例：46244）になっているセルも日付として扱う
Private Function CoerceDateValue(ByVal v As Variant) As Variant
    If IsDate(v) Then
        CoerceDateValue = CDate(v)
    ElseIf IsNumeric(v) Then
        If Val(v) >= 40000 And Val(v) <= 80000 Then
            CoerceDateValue = CDate(CLng(Int(Val(v))))
        Else
            CoerceDateValue = Empty
        End If
    Else
        CoerceDateValue = Empty
    End If
End Function

Private Sub ParseMonthText(ByVal inputText As String, ByRef targetYear As Long, ByRef targetMonth As Long)
    Dim s As String
    s = Trim$(inputText)
    s = ToHalfWidth(s)
    s = Replace(s, "年", "/")
    s = Replace(s, "月", "")
    s = Replace(s, "-", "/")
    s = Replace(s, ".", "/")

    On Error GoTo Bad
    If InStr(s, "/") > 0 Then
        targetYear = CLng(Split(s, "/")(0))
        targetMonth = CLng(Split(s, "/")(1))
    Else
        targetYear = Year(Date)
        targetMonth = CLng(s)
    End If
    Exit Sub
Bad:
    targetYear = 0
    targetMonth = 0
End Sub

' 全角の数字とスラッシュを半角に直す
Private Function ToHalfWidth(ByVal s As String) As String
    Dim i As Long

    For i = 0 To 9
        s = Replace(s, ChrW(65296 + i), CStr(i))    ' ０～９
    Next i
    s = Replace(s, ChrW(65295), "/")                ' ／

    ToHalfWidth = s
End Function

' 実行時に開いているシートが月別シートなら、その月を初期値にする
Private Function DefaultMonthText() As String
    Dim m As Long
    m = MonthOfSheetName(ActiveSheet.Name)
    If m >= 1 And m <= 12 Then
        DefaultMonthText = CStr(Year(Date)) & "/" & CStr(m)
    Else
        DefaultMonthText = CStr(Year(Date)) & "/" & CStr(Month(Date))
    End If
End Function

Private Function IsMonthSheetName(ByVal sheetName As String) As Boolean
    Dim m As Long
    m = MonthOfSheetName(sheetName)
    IsMonthSheetName = (m >= 1 And m <= 12 And InStr(sheetName, "月") > 0 And InStr(sheetName, "年") = 0)
End Function

' 「8月」→8 ／「2026年8月実績」は年が入るので対象外
Private Function MonthOfSheetName(ByVal sheetName As String) As Long
    Dim i As Long, ch As String, digits As String

    For i = 1 To Len(sheetName)
        ch = Mid$(sheetName, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then
        MonthOfSheetName = 0
    Else
        MonthOfSheetName = CLng(digits)
    End If
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub AddOrResetTransferButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                                     ByVal captionText As String, ByVal macroName As String, _
                                     ByVal leftPos As Double, ByVal topPos As Double, _
                                     ByVal widthVal As Double, ByVal heightVal As Double)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPos, topPos, widthVal, heightVal)
    btn.Name = shapeName
    btn.Caption = captionText
    btn.OnAction = macroName
End Sub
