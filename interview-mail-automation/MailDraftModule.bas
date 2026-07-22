Attribute VB_Name = "MailDraftModule"
Option Explicit

' =====================================================================
'  面談打診メール自動化マクロ（Outlook下書き作成）
'
'  ■初回だけ：SetupInterviewMail を実行
'     → 「メールテンプレ」シート作成・チェック列追加・ボタン設置
'  ■ふだんの運用：
'     1. フォローアップシートのT列「次回予定」に次回面談予定日を入力
'     2. 「面談打診メール作成」ボタンを押す
'     3. 対象日数以内・未送信の人のOutlook下書きが一括作成される
'        （宛先はエンジニアマスタのメールアドレスを自動参照）
'     4. Outlookの「下書き」フォルダで内容を確認して送信する
'  ■可視化：
'     BL列「打診メール」…下書き作成済み＝○（緑）
'     オレンジ＝要対応（アドレス不明、または予定日超過で未送信）
' =====================================================================

Private Const SHEET_FOLLOW As String = "フォローアップシート"
Private Const SHEET_MASTER As String = "エンジニアマスタ"
Private Const SHEET_TEMPLATE As String = "メールテンプレ"

Private Const COL_EMPNO As Long = 3       ' C列 社員番号
Private Const COL_NAME As Long = 4        ' D列 氏名
Private Const COL_COMPANY As Long = 6     ' F列 企業名
Private Const COL_SALES As Long = 8       ' H列 担当営業
Private Const COL_METHOD As Long = 11     ' K列 面談方法
Private Const COL_NEXTDATE As Long = 20   ' T列 次回予定（面談予定日を入力する列）
Private Const COL_CHECK As Long = 64      ' BL列 打診メール（○＝下書き作成済）
Private Const COL_DRAFTDATE As Long = 65  ' BM列 下書き作成日

Private Const DATA_START_ROW As Long = 3

Private Const TPL_SUBJECT_ROW As Long = 3
Private Const TPL_CC_ROW As Long = 5
Private Const TPL_LEAD_ROW As Long = 6
Private Const TPL_BODY_START As Long = 9

' ---------------------------------------------------------------------
' 初回セットアップ（1回だけ実行すればOK。何度実行しても壊れません）
' ---------------------------------------------------------------------
Public Sub SetupInterviewMail()
    EnsureTemplateSheet
    EnsureCheckColumns
    InstallMailButton
    MsgBox "セットアップが完了しました。" & vbCrLf & vbCrLf & _
           "1. 「" & SHEET_TEMPLATE & "」シートで件名・本文を編集してください。" & vbCrLf & _
           "2. " & SHEET_FOLLOW & "のT列「次回予定」に面談予定日を入力してください。" & vbCrLf & _
           "3. 「面談打診メール作成」ボタンでOutlookの下書きが作成されます。", _
           vbInformation, "面談打診メール"
End Sub

' ---------------------------------------------------------------------
' メイン：対象者のOutlook下書きを一括作成
' ---------------------------------------------------------------------
Public Sub CreateInterviewMailDrafts()
    Dim wsF As Worksheet, wsM As Worksheet, wsT As Worksheet

    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets(SHEET_FOLLOW)
    Set wsM = ThisWorkbook.Worksheets(SHEET_MASTER)
    On Error GoTo 0
    If wsF Is Nothing Or wsM Is Nothing Then
        MsgBox "「" & SHEET_FOLLOW & "」または「" & SHEET_MASTER & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    EnsureTemplateSheet
    EnsureCheckColumns
    Set wsT = ThisWorkbook.Worksheets(SHEET_TEMPLATE)

    ' --- テンプレート読み込み ---
    Dim subjectTpl As String, bodyTpl As String, ccAddr As String
    Dim leadDays As Long
    subjectTpl = Trim(CStr(wsT.Cells(TPL_SUBJECT_ROW, 2).Value & ""))
    ccAddr = Trim(CStr(wsT.Cells(TPL_CC_ROW, 2).Value & ""))
    leadDays = Val(wsT.Cells(TPL_LEAD_ROW, 2).Value & "")
    If leadDays <= 0 Then leadDays = 14

    Dim lastBodyRow As Long, i As Long
    lastBodyRow = wsT.Cells(wsT.Rows.Count, 2).End(xlUp).Row
    If subjectTpl = "" Or lastBodyRow < TPL_BODY_START Then
        MsgBox "「" & SHEET_TEMPLATE & "」シートの件名または本文が空です。先にテンプレートを入力してください。", vbExclamation
        Exit Sub
    End If
    For i = TPL_BODY_START To lastBodyRow
        If i > TPL_BODY_START Then bodyTpl = bodyTpl & vbCrLf
        bodyTpl = bodyTpl & CStr(wsT.Cells(i, 2).Value & "")
    Next i

    ' --- エンジニアマスタから宛先辞書を作成（社員番号→アドレス、氏名→アドレス） ---
    Dim emailByNo As Object, emailByName As Object
    Set emailByNo = CreateObject("Scripting.Dictionary")
    Set emailByName = CreateObject("Scripting.Dictionary")
    Dim mLast As Long, keyNo As String, keyName As String, addr As String
    mLast = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    For i = 2 To mLast
        keyNo = Trim(CStr(wsM.Cells(i, 1).Value & ""))
        keyName = NormalizeName(CStr(wsM.Cells(i, 2).Value & ""))
        addr = Trim(CStr(wsM.Cells(i, 7).Value & ""))
        If addr <> "" Then
            If keyNo <> "" And Not emailByNo.Exists(keyNo) Then emailByNo.Add keyNo, addr
            If keyName <> "" And Not emailByName.Exists(keyName) Then emailByName.Add keyName, addr
        End If
    Next i

    ' --- Outlook起動 ---
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlookを起動できませんでした。デスクトップ版Outlookがインストールされているか確認してください。", vbExclamation
        Exit Sub
    End If

    ' --- 対象行を走査して下書き作成 ---
    Dim lastRow As Long, r As Long
    Dim nm As String, dtVal As Variant, d As Date
    Dim madeCount As Long, madeNames As String
    Dim missNames As String, lateNames As String
    Dim mail As Object

    lastRow = wsF.Cells(wsF.Rows.Count, COL_NAME).End(xlUp).Row
    Application.ScreenUpdating = False

    For r = DATA_START_ROW To lastRow
        nm = Trim(CStr(wsF.Cells(r, COL_NAME).Value & ""))
        If nm <> "" Then
            dtVal = wsF.Cells(r, COL_NEXTDATE).Value
            d = 0
            If IsDate(dtVal) Then
                d = Int(CDate(dtVal))
            ElseIf IsNumeric(dtVal) Then
                ' 日付書式が外れてシリアル値（例：46213）になっているセルにも対応
                If Val(dtVal) >= 40000 And Val(dtVal) <= 80000 Then d = Int(Val(dtVal))
            End If
            If d > 0 Then
                ' 送信済み（BL列に何か入っている）はスキップ → 二重作成防止
                If Trim(CStr(wsF.Cells(r, COL_CHECK).Value & "")) = "" Then
                    If d < Date Then
                        ' 予定日を過ぎたのに未送信 → オレンジで警告表示
                        wsF.Cells(r, COL_CHECK).Interior.Color = RGB(255, 199, 206)
                        lateNames = lateNames & vbCrLf & "・" & nm & "（" & Format(d, "m月d日") & "）"
                    ElseIf d <= Date + leadDays Then
                        ' 対象：アドレスを引いて下書き作成
                        addr = ""
                        keyNo = Trim(CStr(wsF.Cells(r, COL_EMPNO).Value & ""))
                        If emailByNo.Exists(keyNo) Then addr = emailByNo(keyNo)
                        If addr = "" Then
                            keyName = NormalizeName(nm)
                            If emailByName.Exists(keyName) Then addr = emailByName(keyName)
                        End If

                        If addr = "" Then
                            wsF.Cells(r, COL_CHECK).Interior.Color = RGB(255, 199, 206)
                            missNames = missNames & vbCrLf & "・" & nm
                        Else
                            Set mail = olApp.CreateItem(0)
                            mail.To = addr
                            If ccAddr <> "" Then mail.CC = ccAddr
                            mail.Subject = ReplaceTags(subjectTpl, wsF, r, d)
                            mail.Body = ReplaceTags(bodyTpl, wsF, r, d)
                            mail.Save   ' Outlookの「下書き」フォルダに保存（送信はしない）

                            With wsF.Cells(r, COL_CHECK)
                                .Value = "○"
                                .HorizontalAlignment = xlCenter
                                .Interior.Color = RGB(198, 239, 206)
                            End With
                            With wsF.Cells(r, COL_DRAFTDATE)
                                .Value = Date
                                .NumberFormatLocal = "m/d"
                            End With
                            madeCount = madeCount + 1
                            madeNames = madeNames & vbCrLf & "・" & nm & "（" & Format(d, "m月d日") & " 面談）"
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Application.ScreenUpdating = True

    ' --- 結果表示 ---
    Dim msg As String
    If madeCount > 0 Then
        msg = "Outlookの下書きを " & madeCount & " 件作成しました。" & madeNames & vbCrLf & vbCrLf & _
              "Outlookの「下書き」フォルダで内容を確認してから送信してください。"
    Else
        msg = "今回、下書きの作成対象はありませんでした。" & vbCrLf & _
              "（対象：T列の次回予定日が今日から " & leadDays & " 日以内で、BL列が未チェックの人）"
    End If
    If missNames <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【アドレス不明（エンジニアマスタを確認）】" & missNames
    End If
    If lateNames <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【予定日を過ぎて未送信】" & lateNames
    End If
    MsgBox msg, vbInformation, "面談打診メール"
End Sub

' ---------------------------------------------------------------------
' 差し込みタグの置換
' ---------------------------------------------------------------------
Private Function ReplaceTags(ByVal tpl As String, ByVal wsF As Worksheet, _
                             ByVal r As Long, ByVal d As Date) As String
    Dim s As String
    s = tpl
    s = Replace(s, "{氏名}", Trim(CStr(wsF.Cells(r, COL_NAME).Value & "")))
    s = Replace(s, "{面談予定日}", Format(d, "yyyy年m月d日(aaa)"))
    s = Replace(s, "{企業名}", Trim(CStr(wsF.Cells(r, COL_COMPANY).Value & "")))
    s = Replace(s, "{担当営業}", Trim(CStr(wsF.Cells(r, COL_SALES).Value & "")))
    s = Replace(s, "{面談方法}", Trim(CStr(wsF.Cells(r, COL_METHOD).Value & "")))
    s = Replace(s, "{社員番号}", Trim(CStr(wsF.Cells(r, COL_EMPNO).Value & "")))
    ReplaceTags = s
End Function

' 氏名照合用：全角・半角スペースを除去して比較
Private Function NormalizeName(ByVal s As String) As String
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    NormalizeName = Trim(s)
End Function

' ---------------------------------------------------------------------
' 「メールテンプレ」シートが無ければ既定内容で作成
' ---------------------------------------------------------------------
Private Sub EnsureTemplateSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_TEMPLATE)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SHEET_MASTER))
    ws.Name = SHEET_TEMPLATE

    With ws
        .Range("A1").Value = "面談打診メール テンプレート（黄色いセルを編集してください）"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12

        .Range("A3").Value = "件名"
        .Range("B3").Value = "【面談日程のご相談】次回面談のお願い（{氏名}様）"

        .Range("A5").Value = "CC（任意）"
        .Range("A6").Value = "対象日数"
        .Range("B6").Value = 14
        .Range("C6").Value = "← T列の次回予定日が今日から○日以内（当日含む）の未送信者を下書き作成の対象にします"

        .Range("A8").Value = "本文（9行目以降のB列を上から順に連結します。空白セルは空行になります）"
        .Range("A8").Font.Bold = True

        Dim bodyLines As Variant, i As Long
        bodyLines = Array( _
            "{氏名}様", _
            "", _
            "お疲れ様です。担当の〇〇です。", _
            "", _
            "次回の定期面談を下記の日程で予定しております。", _
            "", _
            "■面談予定日：{面談予定日}", _
            "■面談方法　：{面談方法}", _
            "", _
            "上記の日程でご都合はいかがでしょうか。", _
            "難しい場合は、ご都合の良い日時をいくつかお知らせください。", _
            "", _
            "よろしくお願いいたします。")
        For i = 0 To UBound(bodyLines)
            .Cells(TPL_BODY_START + i, 2).Value = bodyLines(i)
        Next i

        .Range("E3").Value = "使える差し込みタグ"
        .Range("E3").Font.Bold = True
        .Range("E4").Value = "{氏名}　{面談予定日}　{企業名}　{担当営業}　{面談方法}　{社員番号}"
        .Range("E5").Value = "※{面談予定日}は「2026年7月30日(木)」の形式で入ります"

        .Range("B3").Interior.Color = RGB(255, 255, 153)
        .Range("B5").Interior.Color = RGB(255, 255, 153)
        .Range("B6").Interior.Color = RGB(255, 255, 153)
        .Range(.Cells(TPL_BODY_START, 2), .Cells(TPL_BODY_START + 30, 2)).Interior.Color = RGB(255, 255, 153)

        .Columns(1).ColumnWidth = 12
        .Columns(2).ColumnWidth = 60
        .Columns(5).ColumnWidth = 60
    End With
End Sub

' ---------------------------------------------------------------------
' フォローアップシートにチェック列の見出しを追加
' ---------------------------------------------------------------------
Private Sub EnsureCheckColumns()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets(SHEET_FOLLOW)
    With wsF
        If Trim(CStr(.Cells(2, COL_CHECK).Value & "")) = "" Then .Cells(2, COL_CHECK).Value = "打診メール"
        If Trim(CStr(.Cells(2, COL_DRAFTDATE).Value & "")) = "" Then .Cells(2, COL_DRAFTDATE).Value = "下書き作成日"
        .Cells(2, COL_CHECK).Font.Bold = True
        .Cells(2, COL_DRAFTDATE).Font.Bold = True
        .Columns(COL_CHECK).ColumnWidth = 10
        .Columns(COL_DRAFTDATE).ColumnWidth = 12
    End With
End Sub

' ---------------------------------------------------------------------
' フォローアップシートに実行ボタンを設置（既存マクロと同じ方式）
' ---------------------------------------------------------------------
Private Sub InstallMailButton()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets(SHEET_FOLLOW)

    On Error Resume Next
    wsF.Shapes("btnMailDraft").Delete
    On Error GoTo 0

    Dim btn As Button
    Set btn = wsF.Buttons.Add(wsF.Range("N1").Left, 2, 150, 26)
    btn.Name = "btnMailDraft"
    btn.Caption = "面談打診メール作成"
    btn.OnAction = "CreateInterviewMailDrafts"
End Sub
