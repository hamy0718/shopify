Attribute VB_Name = "MailDraftModule"
Option Explicit

' =====================================================================
'  面談打診メール一括下書き作成マクロ（N列運用版）
'
'  ■運用ルール
'    O列=面談日(仮)・P列=時間(仮) … 仮日程を入力しておく
'    N列=打診メール ○ … 下書き作成時に自動で○が付く（未送信の人は空欄）
'    Q列=日時確定 ○ … 相手と確定したら手動で○
'
'  ■初回だけ：SetupInterviewMail を実行
'     → 「メールテンプレ」シート作成＋「面談打診メール作成」ボタン設置
'  ■ふだん：ボタンを押すと、
'     「O列の面談日が今日から対象日数以内・N列が空欄・Q列が空欄」の人の
'     Outlook下書きを一括作成し、N列に○を付ける
'     （メールは送信されません。Outlookの下書きフォルダで確認して送信）
' =====================================================================

Private Const SHEET_FOLLOW As String = "フォローアップシート"
Private Const SHEET_MASTER As String = "エンジニアマスタ"
Private Const SHEET_TEMPLATE As String = "メールテンプレ"

Private Const COL_EMPNO As Long = 3      ' C列 社員番号
Private Const COL_NAME As Long = 4       ' D列 氏名
Private Const COL_COMPANY As Long = 6    ' F列 企業名
Private Const COL_SALES As Long = 8      ' H列 担当営業
Private Const COL_METHOD As Long = 11    ' K列 面談方法
Private Const COL_MAIL As Long = 14      ' N列 打診メール○
Private Const COL_DATE As Long = 15      ' O列 面談日（仮）
Private Const COL_TIME As Long = 16      ' P列 時間（仮）
Private Const COL_FIXED As Long = 17     ' Q列 日時確定○

Private Const DATA_START_ROW As Long = 3

Private Const TPL_SUBJECT_ROW As Long = 3
Private Const TPL_CC_ROW As Long = 5
Private Const TPL_LEAD_ROW As Long = 6
Private Const TPL_BODY_START As Long = 9

' ---------------------------------------------------------------------
' 初回セットアップ（何度実行しても壊れません）
' ---------------------------------------------------------------------
Public Sub SetupInterviewMail()
    EnsureTemplateSheet
    InstallMailButton
    MsgBox "セットアップ完了。" & vbCrLf & _
           "「" & SHEET_TEMPLATE & "」シートで件名・本文・対象日数を確認してから、" & vbCrLf & _
           "フォローアップシートの「面談打診メール作成」ボタンを押してください。", vbInformation
End Sub

' ---------------------------------------------------------------------
' メイン：打診メールの下書きを一括作成
' ---------------------------------------------------------------------
Public Sub CreateInterviewMailDrafts()
    Dim wsF As Worksheet, wsM As Worksheet, wsT As Worksheet
    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets(SHEET_FOLLOW)
    Set wsM = ThisWorkbook.Worksheets(SHEET_MASTER)
    On Error GoTo 0
    If wsF Is Nothing Or wsM Is Nothing Then
        MsgBox "「" & SHEET_FOLLOW & "」または「" & SHEET_MASTER & "」が見つかりません。", vbExclamation
        Exit Sub
    End If
    EnsureTemplateSheet
    Set wsT = ThisWorkbook.Worksheets(SHEET_TEMPLATE)

    ' テンプレート読み込み
    Dim subjectTpl As String, bodyTpl As String, ccAddr As String, leadDays As Long
    subjectTpl = Trim(CStr(wsT.Cells(TPL_SUBJECT_ROW, 2).Value & ""))
    ccAddr = Trim(CStr(wsT.Cells(TPL_CC_ROW, 2).Value & ""))
    leadDays = Val(wsT.Cells(TPL_LEAD_ROW, 2).Value & "")
    If leadDays <= 0 Then leadDays = 14
    Dim lastBody As Long, i As Long
    lastBody = wsT.Cells(wsT.Rows.Count, 2).End(xlUp).Row
    If subjectTpl = "" Or lastBody < TPL_BODY_START Then
        MsgBox "「" & SHEET_TEMPLATE & "」の件名または本文が空です。", vbExclamation
        Exit Sub
    End If
    For i = TPL_BODY_START To lastBody
        If i > TPL_BODY_START Then bodyTpl = bodyTpl & vbCrLf
        bodyTpl = bodyTpl & CStr(wsT.Cells(i, 2).Value & "")
    Next i

    ' エンジニアマスタ → 宛先辞書（社員番号／氏名）
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

    ' Outlook起動
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlookを起動できませんでした。デスクトップ版Outlookが必要です。", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long, r As Long
    Dim nm As String, dv As Variant, d As Date, tv As Variant, tTxt As String
    Dim madeCount As Long, madeNames As String, missNames As String
    Dim mail As Object

    lastRow = wsF.Cells(wsF.Rows.Count, COL_NAME).End(xlUp).Row
    Application.ScreenUpdating = False

    For r = DATA_START_ROW To lastRow
        nm = Trim(CStr(wsF.Cells(r, COL_NAME).Value & ""))
        If nm <> "" Then
            ' N列に○あり／Q列で確定済みはスキップ
            If Trim(CStr(wsF.Cells(r, COL_MAIL).Value & "")) = "" _
               And Trim(CStr(wsF.Cells(r, COL_FIXED).Value & "")) = "" Then
                dv = CoerceDateM(wsF.Cells(r, COL_DATE).Value)
                If IsDate(dv) Then
                    d = CDate(dv)
                    tv = wsF.Cells(r, COL_TIME).Value
                    tTxt = TimeText(tv)
                    If d >= Date And d <= Date + leadDays And tTxt <> "" Then
                        ' 宛先を検索
                        addr = ""
                        keyNo = Trim(CStr(wsF.Cells(r, COL_EMPNO).Value & ""))
                        If emailByNo.Exists(keyNo) Then addr = emailByNo(keyNo)
                        If addr = "" Then
                            keyName = NormalizeName(nm)
                            If emailByName.Exists(keyName) Then addr = emailByName(keyName)
                        End If

                        If addr = "" Then
                            wsF.Cells(r, COL_MAIL).Interior.Color = RGB(255, 199, 206)
                            missNames = missNames & vbCrLf & "・" & nm
                        Else
                            Set mail = olApp.CreateItem(0)
                            mail.To = addr
                            If ccAddr <> "" Then mail.CC = ccAddr
                            mail.Subject = ReplaceTags(subjectTpl, wsF, r, d, tTxt)
                            mail.Body = ReplaceTags(bodyTpl, wsF, r, d, tTxt)
                            mail.Save   ' 下書きフォルダへ保存（送信しない）

                            With wsF.Cells(r, COL_MAIL)
                                .Value = "○"
                                .HorizontalAlignment = xlCenter
                                .Interior.Pattern = xlNone
                            End With
                            madeCount = madeCount + 1
                            madeNames = madeNames & vbCrLf & "・" & nm & "（" & Format(d, "m月d日(aaa)") & " " & tTxt & "）"
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Application.ScreenUpdating = True

    Dim msg As String
    If madeCount > 0 Then
        msg = "Outlook下書きを " & madeCount & " 件作成し、N列に○を付けました。" & madeNames & vbCrLf & vbCrLf & _
              "Outlookの「下書き」フォルダで内容を確認して送信してください。"
    Else
        msg = "作成対象がありませんでした。" & vbCrLf & _
              "（対象：O列の面談日が今日から " & leadDays & " 日以内・N列空欄・Q列空欄・P列に時間あり）"
    End If
    If missNames <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【アドレス不明（N列をオレンジ表示。エンジニアマスタを確認）】" & missNames
    End If
    MsgBox msg, vbInformation, "面談打診メール"
End Sub

' ---------------------------------------------------------------------
Private Function ReplaceTags(ByVal tpl As String, ByVal wsF As Worksheet, _
        ByVal r As Long, ByVal d As Date, ByVal tTxt As String) As String
    Dim s As String
    s = tpl
    s = Replace(s, "{氏名}", Trim(CStr(wsF.Cells(r, COL_NAME).Value & "")))
    s = Replace(s, "{面談予定日}", Format(d, "yyyy年m月d日(aaa)"))
    s = Replace(s, "{面談時間}", tTxt)
    s = Replace(s, "{企業名}", Trim(CStr(wsF.Cells(r, COL_COMPANY).Value & "")))
    s = Replace(s, "{担当営業}", Trim(CStr(wsF.Cells(r, COL_SALES).Value & "")))
    s = Replace(s, "{面談方法}", Trim(CStr(wsF.Cells(r, COL_METHOD).Value & "")))
    s = Replace(s, "{社員番号}", Trim(CStr(wsF.Cells(r, COL_EMPNO).Value & "")))
    ReplaceTags = s
End Function

Private Function NormalizeName(ByVal s As String) As String
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    NormalizeName = Trim(s)
End Function

' 日付書式の外れたシリアル値にも対応
Private Function CoerceDateM(ByVal v As Variant) As Variant
    If IsDate(v) Then
        CoerceDateM = CDate(v)
    ElseIf IsNumeric(v) Then
        If Val(v) >= 40000 And Val(v) <= 80000 Then
            CoerceDateM = CDate(CLng(Int(Val(v))))
        Else
            CoerceDateM = Empty
        End If
    Else
        CoerceDateM = Empty
    End If
End Function

' 時刻セル（時刻値／小数／文字列）を "h:mm" 文字列へ
Private Function TimeText(ByVal v As Variant) As String
    On Error GoTo BadValue
    If IsDate(v) Then
        TimeText = Format(CDate(v), "h:mm")
    ElseIf IsNumeric(v) Then
        If Val(v) > 0 And Val(v) < 1 Then TimeText = Format(CDbl(v), "h:mm")
    ElseIf Len(Trim(CStr(v))) > 0 Then
        TimeText = Format(TimeValue(Trim(CStr(v))), "h:mm")
    End If
    Exit Function
BadValue:
    TimeText = ""
End Function

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
        .Range("A1").Value = "面談打診メール テンプレート（黄色いセルを編集）"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        .Range("A3").Value = "件名"
        .Range("B3").Value = "【面談日程のご相談】次回面談のお願い（{氏名}様）"
        .Range("A5").Value = "CC（任意）"
        .Range("A6").Value = "対象日数"
        .Range("B6").Value = 14
        .Range("C6").Value = "← O列の面談日が今日から○日以内（当日含む）の人を対象にします"
        .Range("A8").Value = "本文（9行目以降のB列を上から連結。空白セルは空行）"
        .Range("A8").Font.Bold = True
        Dim bodyLines As Variant
        Dim i As Long
        bodyLines = Array( _
            "{氏名}様", _
            "", _
            "お疲れ様です。担当の〇〇です。", _
            "", _
            "次回の定期面談を下記の日程で予定しております。", _
            "", _
            "■日時：{面談予定日}　{面談時間}～（30分程度）", _
            "■方法：{面談方法}", _
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
        .Range("E4").Value = "{氏名}　{面談予定日}　{面談時間}　{企業名}　{担当営業}　{面談方法}　{社員番号}"
        .Range("E5").Value = "※{面談予定日}は「2026年8月12日(水)」、{面談時間}は「10:30」の形式"
        .Range("B3").Interior.Color = RGB(255, 255, 153)
        .Range("B5").Interior.Color = RGB(255, 255, 153)
        .Range("B6").Interior.Color = RGB(255, 255, 153)
        .Range(.Cells(TPL_BODY_START, 2), .Cells(TPL_BODY_START + 30, 2)).Interior.Color = RGB(255, 255, 153)
        .Columns(1).ColumnWidth = 12
        .Columns(2).ColumnWidth = 60
        .Columns(5).ColumnWidth = 60
    End With
End Sub

Private Sub InstallMailButton()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets(SHEET_FOLLOW)
    On Error Resume Next
    wsF.Shapes("btnMailDraft").Delete
    On Error GoTo 0
    Dim btn As Button
    Set btn = wsF.Buttons.Add(wsF.Range("F1").Left, 2, 150, 24)
    btn.Name = "btnMailDraft"
    btn.Caption = "面談打診メール作成"
    btn.OnAction = "CreateInterviewMailDrafts"
End Sub
