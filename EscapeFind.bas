Attribute VB_Name = "Module1"
Option Explicit

'============================================================
' xlsmツール（別ファイル）から、選択した Excel ファイルを開いて加工するマクロ
'
' 【現在の仕様】
' - 対象シート: "A1-1-1" のみ
' - 対象範囲  : B列の 5行目以降
' - 装飾対象  : B列セル内の "sqlX(...)" 部分だけを赤字＋太字
'               （X は Config シートに定義した prefix）
' - 出力先    : ヒットした行の C列に "SQLインジェクション対策済み" を赤字で記入
'               ※C列に既存文字がある場合は " ・ " をつけて追記
'
' 【前提】
' - この xlsm に "Config" シートがあること
' - Config!A2 以降に prefix を列挙していること（例: sqlS / sqlN）
'   ※ "(" は付けずに prefix 名だけを記入してください
'============================================================
Public Sub ApplySqlMarks_ToSelectedWorkbook()
    Dim targetPath As String
    targetPath = PickExcelFilePath()

    ' ユーザーがキャンセルした場合は終了
    If targetPath = "" Then Exit Sub

    '------------------------------------------------------------
    ' Config シートから prefix 一覧を読み込む
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
    ' Excel の状態を退避（処理後に必ず戻す）
    '------------------------------------------------------------
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalc As XlCalculation

    prevScreenUpdating = app.ScreenUpdating
    prevEnableEvents = app.EnableEvents
    prevDisplayAlerts = app.DisplayAlerts
    prevCalc = app.Calculation

    ' 高速化＆誤動作防止のため設定を一時変更
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    app.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    '------------------------------------------------------------
    ' 対象ブックを開く
    '------------------------------------------------------------
    Dim wb As Workbook
    Set wb = app.Workbooks.Open(Filename:=targetPath, ReadOnly:=False)

    '------------------------------------------------------------
    ' 対象シート "A1-1-1" を取得
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

    '------------------------------------------------------------
    ' 対象シートを処理
    '------------------------------------------------------------
    ProcessOneSheet ws, prefixes

    ' 保存して閉じる
    wb.Save
    wb.Close SaveChanges:=False

    MsgBox "完了しました。" & vbCrLf & _
           "「A1-1-1」シートを更新しました:" & vbCrLf & targetPath, vbInformation

CleanExit:
    '------------------------------------------------------------
    ' Excel の設定を元に戻す
    '------------------------------------------------------------
    app.ScreenUpdating = prevScreenUpdating
    app.EnableEvents = prevEnableEvents
    app.DisplayAlerts = prevDisplayAlerts
    app.Calculation = prevCalc
    Exit Sub

CleanFail:
    '------------------------------------------------------------
    ' 例外時もブックを閉じて、Excel 設定を必ず戻す
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
' 1シート分処理
'
' - B列の 5行目以降を走査
' - "sqlX(...)" を見つけたら、その部分だけ赤字＋太字に装飾
' - ヒットした行の C列に "SQLインジェクション対策済み" を赤字で出力
'   ※既存文字がある場合は " ・ " をつけて追記
'============================================================
Private Sub ProcessOneSheet(ByVal ws As Worksheet, ByVal prefixes As Collection)
    Const START_ROW As Long = 5  ' B5 から開始する

    '------------------------------------------------------------
    ' B列の最終行を取得
    ' （B列にデータがほぼ無い場合の安全チェックも行う）
    '------------------------------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If lastRow < START_ROW Then
        ' B5 以降にデータが無い場合は何もしない
        Exit Sub
    End If

    Dim r As Long
    For r = START_ROW To lastRow
        Dim bCell As Range
        Set bCell = ws.Cells(r, "B")

        ' 空セルはスキップ
        If Len(CStr(bCell.Value2)) > 0 Then
            Dim hit As Boolean

            ' B列セル内の sqlX(...) を装飾
            hit = MarkSqlPartsInCell(bCell, prefixes)

            If hit Then
                '----------------------------------------------------
                ' ヒット行の C列へステータスを書き込む（追記対応）
                '----------------------------------------------------
                Dim cCell As Range
                Set cCell = ws.Cells(r, "C")

                AppendStatusToCell cCell, "SQLインジェクション対策済み"
            End If
        End If
    Next r
End Sub
'============================================================
' C列セルにステータス文字列を書き込む（追記対応＋重複防止）
'
' 仕様:
' - 空セルならそのまま設定
' - 既存文字があるなら " ・ " をつけて追記
' - ただし、同じ文言がすでに含まれている場合は追記しない
' - 文字色は赤
'
' 例:
'   既存: "要確認"
'   追記: "SQLインジェクション対策済み"
'   結果: "要確認 ・ SQLインジェクション対策済み"
'
'   既存: "要確認 ・ SQLインジェクション対策済み"
'   追記: "SQLインジェクション対策済み"
'   結果: 変更なし（重複追記しない）
'============================================================
Private Sub AppendStatusToCell(ByVal targetCell As Range, ByVal statusText As String)
    Dim existingText As String
    existingText = CStr(targetCell.Value2)

    If Len(Trim$(existingText)) = 0 Then
        ' 何も入っていない場合はそのまま設定
        targetCell.Value2 = statusText
    Else
        ' 既に同じ文言が含まれているなら追記しない
        If Not HasStatusToken(existingText, statusText) Then
            targetCell.Value2 = existingText & " ・ " & statusText
        End If
    End If

    ' C列は赤文字にする（セル全体）
    targetCell.Font.Color = vbRed
End Sub
'============================================================
' 既存セル文字列の中に、指定したステータス文言が
' 「トークンとして」含まれているかを判定する
'
' 区切り文字は " ・ " を想定して分割して比較する
' （部分一致ではなく、各要素の完全一致で判定）
'
' 例:
'   existingText = "要確認 ・ SQLインジェクション対策済み"
'   statusText   = "SQLインジェクション対策済み"
'   → True
'============================================================
Private Function HasStatusToken(ByVal existingText As String, ByVal statusText As String) As Boolean
    Dim parts() As String
    Dim i As Long

    ' " ・ " 区切りで分割
    parts = Split(existingText, " ・ ")

    For i = LBound(parts) To UBound(parts)
        ' 前後の空白を除去して比較（大文字小文字は区別しない）
        If StrComp(Trim$(parts(i)), Trim$(statusText), vbTextCompare) = 0 Then
            HasStatusToken = True
            Exit Function
        End If
    Next i

    HasStatusToken = False
End Function

'============================================================
' セル内の複数パターンをすべて装飾する
'
' - Config で定義された prefix を順番に確認する
' - 例: sqlS / sqlN など
' - 同一セル内に複数ヒットしてもすべて処理する
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

        ' 1つの prefix について、セル内の全出現箇所を装飾
        anyHit = MarkAllOccurrencesForOnePrefix(cell, text, prefix) Or anyHit
    Next i

    MarkSqlPartsInCell = anyHit
End Function

'============================================================
' 1つの prefix について、セル内の全出現箇所を装飾する（ネスト括弧対応）
'
' 例:
'   prefix = "sqlS" の場合、"sqlS(" を探して対応する ")" までを装飾する
'
' 対応:
' - sqlS(AAA)
' - sqlS(isset(...) ? XXX : YYY)   ← ネスト括弧対応
' - 同一セル内に複数ヒット
'
' 非対応（現仕様）:
' - 文字列リテラル中の括弧を厳密に無視する高度解析
'   （必要になれば後で追加可能）
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
        ' "sqlS(" の先頭（= prefix の先頭位置）を探す
        openPos = InStr(startPos, text, pattern, vbTextCompare)
        If openPos = 0 Then Exit Do

        Dim openParenPos As Long
        ' "(" の位置 = prefix の開始位置 + prefixの文字数
        openParenPos = openPos + Len(prefix)

        Dim closePos As Long
        ' 対応する ")" を括弧の深さで探す
        closePos = FindMatchingCloseParen(text, openParenPos)

        If closePos > 0 Then
            Dim lengthToFormat As Long
            ' prefix の先頭から対応する ")" まで（")" を含む）
            lengthToFormat = (closePos - openPos) + 1

            '--------------------------------------------------------
            ' 該当部分だけ赤字＋太字にする
            '--------------------------------------------------------
            With cell.Characters(openPos, lengthToFormat).Font
                .Color = vbRed
                .Bold = True
            End With

            hit = True

            ' 次の検索は、今回見つけた ")" の次から
            startPos = closePos + 1
        Else
            ' 対応する ")" が見つからない場合（括弧不整合など）
            ' 無限ループ防止のため1文字進める
            startPos = openPos + 1
        End If
    Loop

    MarkAllOccurrencesForOnePrefix = hit
End Function

'============================================================
' 指定した "(" に対応する ")" の位置を返す（ネスト括弧対応）
'
' 引数:
'   text        : 対象文字列
'   openParenPos: "(" の位置（1始まり）
'
' 戻り値:
'   対応する ")" の位置（1始まり）
'   見つからない場合は 0
'============================================================
Private Function FindMatchingCloseParen(ByVal text As String, ByVal openParenPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String

    ' 引数チェック（範囲外なら失敗）
    If openParenPos < 1 Or openParenPos > Len(text) Then
        FindMatchingCloseParen = 0
        Exit Function
    End If

    ' 指定位置が本当に "(" か確認
    If Mid$(text, openParenPos, 1) <> "(" Then
        FindMatchingCloseParen = 0
        Exit Function
    End If

    ' 最初の "(" を1つ開いた状態から開始
    depth = 1

    ' "(" の次の文字から順に走査して、括弧の深さを数える
    For i = openParenPos + 1 To Len(text)
        ch = Mid$(text, i, 1)

        If ch = "(" Then
            depth = depth + 1
        ElseIf ch = ")" Then
            depth = depth - 1

            ' 深さ0になったら、最初の "(" に対応する ")" が見つかった
            If depth = 0 Then
                FindMatchingCloseParen = i
                Exit Function
            End If
        End If
    Next i

    ' 最後まで見ても閉じ括弧が見つからない場合
    FindMatchingCloseParen = 0
End Function

'============================================================
' Config シート A2:A(最終行) から prefix を読み込む
'
' - 空欄は無視
' - Config シートが無い場合は空の Collection を返す
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
'
' - ユーザーに加工対象のExcelファイルを選ばせる
' - キャンセル時は "" を返す
'============================================================
Private Function PickExcelFilePath() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker

    With fd
        .Title = "加工対象の Excel ファイルを選択してください（xlsx推奨）"
        .AllowMultiSelect = False

        ' ファイルフィルタを設定
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        ' ユーザーがキャンセルした場合
        If .Show <> -1 Then
            PickExcelFilePath = ""
            Exit Function
        End If

        ' 選択された1件目のパスを返す
        PickExcelFilePath = .SelectedItems(1)
    End With
End Function

