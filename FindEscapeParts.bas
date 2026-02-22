Attribute VB_Name = "Module1"
Option Explicit

'============================================================
' xlsmツール（別ファイル）から、選択した xlsx を開いて加工するマクロ
' - "A1-1-1" シートのみを対象に処理する
' - B列の 5行目以降を走査する
' - B列の "sqlX(...)"などエスケープ 部分だけを赤字＋太字（複数ヒット対応）で装飾
' - ヒットした行の C列に "SQLインジェクション対策済み" を赤字で書く
'
' 前提:
'  - この xlsm に "Config" シートがあり、A2以降に prefix（例: sqlS, sqlN）を列挙していること
'============================================================
Public Sub FindEscapeParts()
    Dim targetPath As String
    targetPath = PickExcelFilePath()
    If targetPath = "" Then Exit Sub ' キャンセル時は終了

    '------------------------------------------------------------
    ' Config シートから prefix 一覧を読み込む
    ' 例: sqlS / sqlN / sqlX など
    '------------------------------------------------------------
    Dim prefixes As Collection
    Set prefixes = LoadPrefixesFromConfig()

    If prefixes.Count = 0 Then
        MsgBox "Config シートの A2 以降に prefix（例: sqlS, sqlN）を1つ以上入力してください。", vbExclamation
        Exit Sub
    End If

    Dim app As Application
    Set app = Application

    '------------------------------------------------------------
    ' Excelの状態を退避（高速化設定のため）
    ' ※処理後に必ず元に戻す
    '------------------------------------------------------------
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalc As XlCalculation

    prevScreenUpdating = app.ScreenUpdating
    prevEnableEvents = app.EnableEvents
    prevDisplayAlerts = app.DisplayAlerts
    prevCalc = app.Calculation

    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    app.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    '------------------------------------------------------------
    ' 選択されたブックを開く
    '------------------------------------------------------------
    Dim wb As Workbook
    Set wb = app.Workbooks.Open(Filename:=targetPath, ReadOnly:=False)

    '------------------------------------------------------------
    ' ★ "A1-1-1" シートのみ処理する
    '------------------------------------------------------------
    Dim ws As Worksheet
    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets("A1-1-1")
    On Error GoTo 0

    If ws Is Nothing Then
        wb.Close SaveChanges:=False
        MsgBox "対象ファイルに「A1-1-1」シートが存在しませんでした。", vbExclamation
        GoTo CleanExit
    End If

    ' 1シート分の処理を実行
    ProcessOneSheet ws, prefixes

    ' 保存して閉じる
    wb.Save
    wb.Close SaveChanges:=False

    MsgBox "完了しました。" & vbCrLf & _
           "「A1-1-1」シートを更新しました:" & vbCrLf & targetPath, vbInformation

CleanExit:
    '------------------------------------------------------------
    ' Excelの状態を元に戻す
    '------------------------------------------------------------
    app.ScreenUpdating = prevScreenUpdating
    app.EnableEvents = prevEnableEvents
    app.DisplayAlerts = prevDisplayAlerts
    app.Calculation = prevCalc
    Exit Sub

CleanFail:
    '------------------------------------------------------------
    ' エラー時もブックを閉じて、Excel状態を戻す
    '------------------------------------------------------------
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'============================================================
' 1シート分処理:
' - B列（5行目以降）を走査して sqlX(...) を装飾
' - ヒット行のC列にメッセージ＆赤字
'============================================================
Private Sub ProcessOneSheet(ByVal ws As Worksheet, ByVal prefixes As Collection)
    Const START_ROW As Long = 5   ' B5 から開始

    '------------------------------------------------------------
    ' B列の最終行を取得
    ' （B列が空の場合は最終行が1になりやすいので、START_ROW未満なら終了）
    '------------------------------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If lastRow < START_ROW Then Exit Sub

    Dim r As Long
    For r = START_ROW To lastRow
        Dim bCell As Range
        Set bCell = ws.Cells(r, "B")

        ' セルに文字がある場合のみ判定
        If Len(CStr(bCell.Value2)) > 0 Then
            Dim hit As Boolean
            hit = MarkSqlPartsInCell(bCell, prefixes)

            If hit Then
                '----------------------------------------------------
                ' 同じ行のC列にメッセージを赤文字で記入
                '----------------------------------------------------
                Dim cCell As Range
                Set cCell = ws.Cells(r, "C")

                cCell.Value2 = "SQLインジェクション対策済み"
                cCell.Font.Color = vbRed
                ' 太字指定は要件にないため未設定（必要なら .Bold = True を追加）
            End If
        End If
    Next r
End Sub

'============================================================
' セル内の複数パターンをすべて装飾する
' - prefix + "(" の開始位置を探す
' - そこから次の ")" までを赤字＋太字
' - 同一セル内に複数存在してもすべて処理
'
' 戻り値:
'   True  = 1つ以上ヒットして装飾した
'   False = ヒットなし
'============================================================
Private Function MarkSqlPartsInCell(ByVal cell As Range, ByVal prefixes As Collection) As Boolean
    Dim text As String
    text = CStr(cell.Value2)

    Dim anyHit As Boolean
    anyHit = False

    Dim i As Long
    For i = 1 To prefixes.Count
        Dim prefix As String
        prefix = CStr(prefixes(i))

        ' 各prefixについて、セル内の該当箇所をすべて装飾
        anyHit = MarkAllOccurrencesForOnePrefix(cell, text, prefix) Or anyHit
    Next i

    MarkSqlPartsInCell = anyHit
End Function

'============================================================
' 1つの prefix について、セル内の全出現箇所を装飾する
' - 例: prefix="sqlS" なら "sqlS(" をすべて探す
' - 見つけたら、直後の ")" を探して、その範囲を装飾する
' - 同じセル内に複数あっても全部処理する
'============================================================
Private Function MarkAllOccurrencesForOnePrefix(ByVal cell As Range, ByVal text As String, ByVal prefix As String) As Boolean
    Dim pattern As String
    pattern = prefix & "("

    Dim startPos As Long
    startPos = 1

    Dim hit As Boolean
    hit = False

    Do
        Dim openPos As Long
        ' vbTextCompare で大文字小文字を区別しない検索
        openPos = InStr(startPos, text, pattern, vbTextCompare)
        If openPos = 0 Then Exit Do ' もう見つからない

        Dim closePos As Long
        ' "sqlX(" の後ろから最初の ")" を探す
        closePos = InStr(openPos + Len(pattern), text, ")", vbTextCompare)

        If closePos > 0 Then
            Dim lengthToFormat As Long
            ' openPos から closePos まで（")"を含む）
            lengthToFormat = (closePos - openPos) + 1

            '--------------------------------------------------------
            ' 見つかった部分だけ赤字＋太字にする
            '--------------------------------------------------------
            With cell.Characters(openPos, lengthToFormat).Font
                .Color = vbRed
                .Bold = True
            End With

            hit = True

            ' 次回検索は、今回の ")" の次から開始
            startPos = closePos + 1
        Else
            ' ")" が見つからなかった場合は無限ループ防止のため1文字進める
            startPos = openPos + 1
        End If
    Loop

    MarkAllOccurrencesForOnePrefix = hit
End Function

'============================================================
' Config シート A2:A(最終行) から prefix を読み込む
' - 空欄は無視
'============================================================
Private Function LoadPrefixesFromConfig() As Collection
    Dim prefixes As New Collection

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    If ws Is Nothing Then
        Set LoadPrefixesFromConfig = prefixes
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim v As String
        v = Trim$(CStr(ws.Cells(r, "A").Value2))

        If v <> "" Then
            prefixes.Add v
        End If
    Next r

    Set LoadPrefixesFromConfig = prefixes
End Function

'============================================================
' ファイル選択ダイアログ（Excelファイル用）
'============================================================
Private Function PickExcelFilePath() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker

    With fd
        .Title = "加工対象の Excel ファイルを選択してください（xlsx推奨）"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        If .Show <> -1 Then
            PickExcelFilePath = ""
            Exit Function
        End If

        PickExcelFilePath = .SelectedItems(1)
    End With
End Function

