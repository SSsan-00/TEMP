Attribute VB_Name = "modSourceSyntaxCheckWriter"
Option Explicit

' ==========================================
' Excel用: 現行ソースシートのマーキング + 個別シート出力
' 仕様に合わせて、文字列ベース（部分一致）で構文を判定する初版実装
' ==========================================

Private Const SOURCE_TEXT_COL As Long = 3   ' C列
Private Const MARK_COL As Long = 2          ' B列
Private Const SECTION_HEADER_START_ROW As Long = 9
Private Const BLOCK_STEP_NORMAL As Long = 5
Private Const MARK_STRING As String = "*****"
Private Const SYMBOL_FILLED As String = "■"
Private Const SYMBOL_EMPTY As String = "□"
Private Const SHEET_KEY_CURRENT_SOURCE As String = "現行ソース"
Private Const SHEET_KEY_INDIVIDUAL_PREFIX As String = "【個別】"

Public Sub RunMain()
    On Error GoTo ErrorHandler

    Dim featureName As String
    Dim workbookPath As String
    Dim targetWorkbook As Workbook
    Dim currentSourceSheet As Worksheet
    Dim individualSheet As Worksheet
    Dim syntaxEvents As Collection
    Dim markedCount As Long
    Dim resultMessage As String

    ' 1) 機能名を入力
    featureName = PromptFeatureName()
    If Len(featureName) = 0 Then
        MsgBox "処理をキャンセルしました（機能名が未入力です）。", vbInformation
        Exit Sub
    End If

    ' 2) 対象ファイルを選択
    workbookPath = SelectTargetWorkbookPath()
    If Len(workbookPath) = 0 Then
        MsgBox "処理をキャンセルしました（対象ファイルが未選択です）。", vbInformation
        Exit Sub
    End If

    ' 3) 対象ブックを開く
    Set targetWorkbook = OpenTargetWorkbook(workbookPath)
    If targetWorkbook Is Nothing Then
        MsgBox "対象ブックを開けませんでした。", vbExclamation
        Exit Sub
    End If

    ' 4) 対象シートを判定
    Set currentSourceSheet = FindCurrentSourceSheet(targetWorkbook, featureName)
    If currentSourceSheet Is Nothing Then
        MsgBox "現行ソースシートが見つかりませんでした。" & vbCrLf & _
               "条件: シート名に「" & SHEET_KEY_CURRENT_SOURCE & "」を含む", vbExclamation
        Exit Sub
    End If

    Set individualSheet = FindIndividualSheet(targetWorkbook, featureName)

    ' 5) 現行ソースシートに対してマーキング
    MarkCurrentSourceSheet currentSourceSheet, markedCount

    ' 6) 個別シートがある場合は解析結果を書き込む
    If Not individualSheet Is Nothing Then
        Set syntaxEvents = CollectSyntaxEvents(currentSourceSheet)
        WriteIndividualSheet individualSheet, syntaxEvents
        resultMessage = "個別シート出力: 実施（" & individualSheet.Name & "）"
    Else
        resultMessage = "個別シート出力: スキップ（対象シートなし）"
    End If

    ' 対象ブックへ書き込みした内容を保存
    targetWorkbook.Save

    ' 7) 完了メッセージ
    MsgBox "処理が完了しました。" & vbCrLf & _
           "対象ブック: " & targetWorkbook.Name & vbCrLf & _
           "現行ソース: " & currentSourceSheet.Name & vbCrLf & _
           "マーキング件数: " & CStr(markedCount) & vbCrLf & _
           resultMessage, vbInformation

    Exit Sub

ErrorHandler:
    ' 8) エラー時はMsgBox表示のみ（ログ不要）
    MsgBox "エラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation
End Sub

Private Function PromptFeatureName() As String
    ' 機能名の入力を受け取り、前後空白を除去して返す
    Dim inputValue As String
    inputValue = InputBox("機能名を入力してください。", "機能名入力")
    PromptFeatureName = Trim$(inputValue)
End Function

Private Function SelectTargetWorkbookPath() As String
    ' .xlsx を選択させる簡易ダイアログ
    Dim selectedPath As Variant

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excel ブック (*.xlsx),*.xlsx", _
        Title:="対象のExcelファイル（.xlsx）を選択してください")

    If VarType(selectedPath) = vbBoolean Then
        SelectTargetWorkbookPath = vbNullString
    Else
        SelectTargetWorkbookPath = CStr(selectedPath)
    End If
End Function

Private Function OpenTargetWorkbook(ByVal workbookPath As String) As Workbook
    ' 既に開いている場合は既存のWorkbookを返し、未オープンなら開く
    Dim wb As Workbook

    If Len(Trim$(workbookPath)) = 0 Then
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            Set OpenTargetWorkbook = wb
            Exit Function
        End If
    Next wb

    Set OpenTargetWorkbook = Application.Workbooks.Open( _
        Filename:=workbookPath, _
        UpdateLinks:=0, _
        ReadOnly:=False)
End Function

Private Function FindCurrentSourceSheet(ByVal targetWorkbook As Workbook, ByVal featureName As String) As Worksheet
    ' 現行ソースシートの判定ルール:
    ' - 名前に「現行ソース」を含むシートを候補
    ' - 候補が1枚ならそれを採用
    ' - 候補が複数なら、機能名を含むものを優先
    Dim ws As Worksheet
    Dim candidates As Collection
    Dim item As Variant

    Set candidates = New Collection

    For Each ws In targetWorkbook.Worksheets
        If ContainsText(ws.Name, SHEET_KEY_CURRENT_SOURCE) Then
            candidates.Add ws
        End If
    Next ws

    If candidates.Count = 0 Then
        Exit Function
    End If

    If candidates.Count = 1 Then
        Set FindCurrentSourceSheet = candidates.Item(1)
        Exit Function
    End If

    If Len(featureName) > 0 Then
        For Each item In candidates
            Set ws = item
            If ContainsText(ws.Name, featureName) Then
                Set FindCurrentSourceSheet = ws
                Exit Function
            End If
        Next item
    End If

    ' 複数候補があり、機能名一致がない場合は未確定としてNothingを返す
End Function

Private Function FindIndividualSheet(ByVal targetWorkbook As Workbook, ByVal featureName As String) As Worksheet
    ' 個別シート名: 「【個別】」 + 機能名（完全一致）
    Dim targetSheetName As String
    Dim ws As Worksheet

    targetSheetName = SHEET_KEY_INDIVIDUAL_PREFIX & featureName

    For Each ws In targetWorkbook.Worksheets
        If StrComp(ws.Name, targetSheetName, vbTextCompare) = 0 Then
            Set FindIndividualSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub MarkCurrentSourceSheet(ByVal sourceSheet As Worksheet, ByRef markedCount As Long)
    ' 現行ソースシートのC列を走査し、対象構文に該当する行のB列へ*****を設定
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim lineText As String

    markedCount = 0
    lastRow = GetLastRow(sourceSheet, SOURCE_TEXT_COL)

    For rowIndex = 1 To lastRow
        lineText = GetCellText(sourceSheet.Cells(rowIndex, SOURCE_TEXT_COL))

        If IsMarkTargetLine(lineText) Then
            sourceSheet.Cells(rowIndex, MARK_COL).Value = MARK_STRING
            markedCount = markedCount + 1
        End If
    Next rowIndex
End Sub

Private Function CollectSyntaxEvents(ByVal sourceSheet As Worksheet) As Collection
    ' 個別シート出力用に、現行ソースシートの構文イベントを上から順に収集する
    ' 初版は文字列ベース判定（部分一致）
    Dim events As Collection
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim lineText As String
    Dim eventItem As Collection
    Dim switchEndRow As Long

    Set events = New Collection
    lastRow = GetLastRow(sourceSheet, SOURCE_TEXT_COL)
    rowIndex = 1

    Do While rowIndex <= lastRow
        lineText = GetCellText(sourceSheet.Cells(rowIndex, SOURCE_TEXT_COL))

        If Len(Trim$(lineText)) = 0 Then
            rowIndex = rowIndex + 1
            GoTo ContinueLoop
        End If

        ' function を最優先で判定（新しい処理セクション開始のため）
        If IsFunctionLine(lineText) Then
            Set eventItem = CreateFunctionEvent(ParseFunctionName(lineText))
            events.Add eventItem

        ' switch は後続行の case/default を収集するので、まとめてイベント化する
        ElseIf IsSwitchLine(lineText) Then
            Set eventItem = CollectSwitchEvent(sourceSheet, rowIndex, lastRow, switchEndRow)
            events.Add eventItem
            If switchEndRow > rowIndex Then
                rowIndex = switchEndRow
            End If

        ' else / else if / default / case の単体行は、初版では個別シート出力をしない
        ElseIf IsElseIfLine(lineText) Then
            ' 個別シート出力なし
        ElseIf IsElseLine(lineText) Then
            ' 個別シート出力なし
        ElseIf IsCaseLine(lineText) Then
            ' switch収集中に扱う想定
        ElseIf IsDefaultLine(lineText) Then
            ' switch収集中に扱う想定

        ' 通常ブロックの判定（順序に注意: foreach を for より先に判定）
        ElseIf IsForeachLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("FOREACH")
            events.Add eventItem
        ElseIf IsForLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("FOR")
            events.Add eventItem
        ElseIf IsWhileLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("WHILE")
            events.Add eventItem
        ElseIf IsTernaryLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("TERNARY")
            events.Add eventItem
        ElseIf IsIfLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("IF")
            events.Add eventItem
        End If

        rowIndex = rowIndex + 1

ContinueLoop:
    Loop

    Set CollectSyntaxEvents = events
End Function

Private Sub WriteIndividualSheet(ByVal individualSheet As Worksheet, ByVal syntaxEvents As Collection)
    ' 個別シートへ、仕様の書式で処理セクション/確認ブロックを書き込む
    Dim sectionIndex As Long
    Dim nextBlockStartRow As Long
    Dim i As Long
    Dim eventItem As Collection
    Dim eventKind As String
    Dim functionName As String

    ' 初期値（固定）
    sectionIndex = 1
    WriteSectionHeader individualSheet, SECTION_HEADER_START_ROW, sectionIndex, "MAIN"
    nextBlockStartRow = SECTION_HEADER_START_ROW + 1

    For i = 1 To syntaxEvents.Count
        Set eventItem = syntaxEvents.Item(i)
        eventKind = UCase$(EventText(eventItem, "Kind"))

        Select Case eventKind
            Case "FUNCTION"
                ' functionを検出したら、新しい処理セクション見出しを開始
                sectionIndex = sectionIndex + 1
                functionName = EventText(eventItem, "FunctionName", "UNKNOWN")
                WriteSectionHeader individualSheet, nextBlockStartRow, sectionIndex, functionName
                nextBlockStartRow = nextBlockStartRow + 1

            Case "SWITCH"
                nextBlockStartRow = WriteSwitchBlock(individualSheet, nextBlockStartRow, eventItem)

            Case "IF", "TERNARY", "FOR", "FOREACH", "WHILE"
                WriteNormalBlock individualSheet, nextBlockStartRow, eventItem
                nextBlockStartRow = nextBlockStartRow + BLOCK_STEP_NORMAL
        End Select
    Next i
End Sub

Private Sub WriteSectionHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal sectionIndex As Long, ByVal sectionName As String)
    ' 処理セクション見出し行の出力
    ws.Range("A" & CStr(headerRow)).Value = "B" & CStr(sectionIndex)
    ws.Range("K" & CStr(headerRow)).Value = SYMBOL_FILLED
    ws.Range("M" & CStr(headerRow)).Value = "処理（" & sectionName & "）"
End Sub

Private Sub WriteNormalBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal eventItem As Collection)
    ' if / 三項演算子 / for / foreach / while の共通形式
    ws.Range("E" & CStr(startRow)).Value = "XXXXX"
    ws.Range("L" & CStr(startRow)).Value = SYMBOL_EMPTY
    ws.Range("N" & CStr(startRow)).Value = EventText(eventItem, "Title")

    ws.Range("H" & CStr(startRow + 1)).Value = 1
    ws.Range("M" & CStr(startRow + 1)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(startRow + 1)).Value = EventText(eventItem, "Cond1")
    ws.Range("AX" & CStr(startRow + 1)).Value = EventText(eventItem, "Result1")
    ws.Range("CF" & CStr(startRow + 1)).Value = "1,4"

    ws.Range("H" & CStr(startRow + 3)).Value = 2
    ws.Range("M" & CStr(startRow + 3)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(startRow + 3)).Value = EventText(eventItem, "Cond2")
    ws.Range("AX" & CStr(startRow + 3)).Value = EventText(eventItem, "Result2")
    ws.Range("CF" & CStr(startRow + 3)).Value = "1,4"
End Sub

Private Function WriteSwitchBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal eventItem As Collection) As Long
    ' switchブロック:
    ' - ヘッダ1行
    ' - 分岐行を 2行おき（r+1, r+3, ...）に配置
    ' - 最後に「上記のいずれでもない場合」を追加
    ' 戻り値は次の確認ブロック開始行
    Dim switchArg As String
    Dim titleText As String
    Dim caseValues As Collection
    Dim hasDefault As Boolean
    Dim branchRow As Long
    Dim seqNo As Long
    Dim i As Long
    Dim caseValue As String
    Dim lastUsedBranchRow As Long

    switchArg = EventText(eventItem, "SwitchArg", "UNKNOWN")
    titleText = "条件分岐（SWITCH文（" & switchArg & "の値））の確認"
    hasDefault = EventFlag(eventItem, "HasDefault")
    Set caseValues = EventCollection(eventItem, "CaseValues")

    ws.Range("E" & CStr(startRow)).Value = "XXXXX"
    ws.Range("L" & CStr(startRow)).Value = SYMBOL_EMPTY
    ws.Range("N" & CStr(startRow)).Value = titleText

    branchRow = startRow + 1
    seqNo = 1
    lastUsedBranchRow = startRow

    If Not caseValues Is Nothing Then
        For i = 1 To caseValues.Count
            caseValue = CStr(caseValues.Item(i))
            WriteSwitchBranchRow ws, branchRow, seqNo, caseValue & "の場合", "CASE内の処理が行われること"
            lastUsedBranchRow = branchRow
            seqNo = seqNo + 1
            branchRow = branchRow + 2
        Next i
    End If

    ' 追加ケース（いずれでもない場合）
    If hasDefault Then
        WriteSwitchBranchRow ws, branchRow, seqNo, "上記のいずれでもない場合", "DEFAULT内の処理が行われること"
    Else
        WriteSwitchBranchRow ws, branchRow, seqNo, "上記のいずれでもない場合", "CASE内の処理が行われないこと"
    End If
    lastUsedBranchRow = branchRow

    ' 次の開始行は、最後に使った分岐行の2行後（1行空け）
    WriteSwitchBlock = lastUsedBranchRow + 2
End Function

Private Sub WriteSwitchBranchRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal seqNo As Long, ByVal conditionText As String, ByVal expectedText As String)
    ' switchの分岐行（case/default相当）の共通出力
    ws.Range("H" & CStr(rowIndex)).Value = seqNo
    ws.Range("M" & CStr(rowIndex)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(rowIndex)).Value = conditionText
    ws.Range("AX" & CStr(rowIndex)).Value = expectedText
    ws.Range("CF" & CStr(rowIndex)).Value = "1,4"
End Sub

Private Function CreateNormalSyntaxEvent(ByVal syntaxKind As String) As Collection
    ' 通常ブロック（if / ternary / for / foreach / while）の文言をまとめたイベント
    Dim ev As Collection
    Dim kindUpper As String

    kindUpper = UCase$(Trim$(syntaxKind))
    Set ev = NewEvent(kindUpper)

    Select Case kindUpper
        Case "IF"
            ev.Add "条件分岐（IF文）の確認", "Title"
            ev.Add "条件が成立する場合", "Cond1"
            ev.Add "IF内の処理が行われること", "Result1"
            ev.Add "条件が成立しない場合", "Cond2"
            ev.Add "IF内の処理が行われないこと", "Result2"

        Case "TERNARY"
            ev.Add "条件分岐（三項演算子）の確認", "Title"
            ev.Add "条件が成立する場合", "Cond1"
            ev.Add "真である場合の処理が行われること", "Result1"
            ev.Add "条件が成立しない場合", "Cond2"
            ev.Add "偽である場合の処理が行われること", "Result2"

        Case "FOR"
            ev.Add "ループ（FOR文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case "FOREACH"
            ev.Add "ループ（FOREACH文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case "WHILE"
            ev.Add "ループ（WHILE文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case Else
            ' 想定外の種類が来ても最低限の形で返す
            ev.Add "確認", "Title"
            ev.Add "条件1", "Cond1"
            ev.Add "期待結果1", "Result1"
            ev.Add "条件2", "Cond2"
            ev.Add "期待結果2", "Result2"
    End Select

    Set CreateNormalSyntaxEvent = ev
End Function

Private Function CreateFunctionEvent(ByVal functionName As String) As Collection
    ' function検出イベント（個別シートでは新しい処理セクション開始に使用）
    Dim ev As Collection
    Set ev = NewEvent("FUNCTION")
    ev.Add functionName, "FunctionName"
    Set CreateFunctionEvent = ev
End Function

Private Function CollectSwitchEvent(ByVal sourceSheet As Worksheet, ByVal switchRow As Long, ByVal lastRow As Long, ByRef endRow As Long) As Collection
    ' switch行を起点に、後続のcase/defaultを簡易的に収集して1イベントにまとめる
    ' 終端判定は厳密にせず、以下のような簡易条件で打ち切る:
    ' - 次のfunctionが来た
    ' - 次のswitchが来た
    ' - case/defaultを拾った後に空行が2行続いた
    Dim ev As Collection
    Dim caseValues As Collection
    Dim switchArg As String
    Dim hasDefault As Boolean
    Dim r As Long
    Dim lineText As String
    Dim trimmedText As String
    Dim blankStreak As Long
    Dim foundBranch As Boolean
    Dim parsedCase As String

    Set ev = NewEvent("SWITCH")
    Set caseValues = New Collection
    switchArg = ParseSwitchArgument(GetCellText(sourceSheet.Cells(switchRow, SOURCE_TEXT_COL)))

    For r = switchRow + 1 To lastRow
        lineText = GetCellText(sourceSheet.Cells(r, SOURCE_TEXT_COL))
        trimmedText = Trim$(lineText)

        ' 次のfunction / switch は次の構文として扱いたいので、ここで打ち切る
        If IsFunctionLine(lineText) Then
            Exit For
        End If
        If IsSwitchLine(lineText) Then
            Exit For
        End If

        If Len(trimmedText) = 0 Then
            blankStreak = blankStreak + 1
            If foundBranch And blankStreak >= 2 Then
                Exit For
            End If
        Else
            blankStreak = 0

            If IsCaseLine(lineText) Then
                parsedCase = ParseCaseValue(lineText)
                caseValues.Add parsedCase
                foundBranch = True
            ElseIf IsDefaultLine(lineText) Then
                hasDefault = True
                foundBranch = True
            End If
        End If
    Next r

    ev.Add switchArg, "SwitchArg"
    ev.Add hasDefault, "HasDefault"
    ev.Add caseValues, "CaseValues"

    ' endRow は、外側ループで再判定したくない範囲の最後の行
    If r > lastRow Then
        endRow = lastRow
    Else
        endRow = r - 1
        If endRow < switchRow Then
            endRow = switchRow
        End If
    End If

    Set CollectSwitchEvent = ev
End Function

Private Function ParseFunctionName(ByVal lineText As String) As String
    ' `function` の後ろの識別子を簡易抽出
    ' 例: function foo(XXX){  -> foo
    Dim posFunction As Long
    Dim restText As String
    Dim i As Long
    Dim ch As String
    Dim nameBuffer As String

    posFunction = InStr(1, lineText, "function", vbTextCompare)
    If posFunction = 0 Then
        ParseFunctionName = "UNKNOWN"
        Exit Function
    End If

    restText = Mid$(lineText, posFunction + Len("function"))
    restText = Trim$(restText)

    If Len(restText) = 0 Then
        ParseFunctionName = "UNKNOWN"
        Exit Function
    End If

    For i = 1 To Len(restText)
        ch = Mid$(restText, i, 1)

        If ch = "(" Or ch = " " Or ch = vbTab Or ch = "{" Then
            If Len(nameBuffer) > 0 Then
                Exit For
            End If
        Else
            nameBuffer = nameBuffer & ch
        End If
    Next i

    If Len(nameBuffer) = 0 Then
        ParseFunctionName = "UNKNOWN"
    Else
        ParseFunctionName = nameBuffer
    End If
End Function

Private Function ParseSwitchArgument(ByVal lineText As String) As String
    ' `switch(YYY)` の括弧内を簡易抽出
    Dim posSwitch As Long
    Dim posOpen As Long
    Dim posClose As Long
    Dim argText As String

    posSwitch = InStr(1, lineText, "switch", vbTextCompare)
    If posSwitch = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    posOpen = InStr(posSwitch, lineText, "(", vbBinaryCompare)
    If posOpen = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    posClose = InStr(posOpen + 1, lineText, ")", vbBinaryCompare)
    If posClose = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    argText = Mid$(lineText, posOpen + 1, posClose - posOpen - 1)
    argText = Trim$(argText)

    If Len(argText) = 0 Then
        ParseSwitchArgument = "UNKNOWN"
    Else
        ParseSwitchArgument = argText
    End If
End Function

Private Function ParseCaseValue(ByVal lineText As String) As String
    ' `case XXX:` の XXX 部分を簡易抽出
    Dim posCase As Long
    Dim restText As String
    Dim posColon As Long

    posCase = InStr(1, lineText, "case", vbTextCompare)
    If posCase = 0 Then
        ParseCaseValue = "UNKNOWN"
        Exit Function
    End If

    restText = Mid$(lineText, posCase + Len("case"))
    posColon = InStr(1, restText, ":", vbBinaryCompare)
    If posColon > 0 Then
        restText = Left$(restText, posColon - 1)
    End If

    restText = Trim$(restText)
    If Len(restText) = 0 Then
        ParseCaseValue = "UNKNOWN"
    Else
        ParseCaseValue = restText
    End If
End Function

Private Function GetLastRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    ' 指定列の最終行を返す（列が空でも最低1を返す）
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 1
    End If

    GetLastRow = lastRow
End Function

Private Function GetCellText(ByVal targetCell As Range) As String
    ' エラー値セルを安全に文字列化するためのヘルパー
    On Error GoTo SafeExit

    If IsError(targetCell.Value) Then
        GetCellText = vbNullString
    ElseIf IsEmpty(targetCell.Value) Then
        GetCellText = vbNullString
    Else
        GetCellText = CStr(targetCell.Value)
    End If
    Exit Function

SafeExit:
    GetCellText = vbNullString
End Function

Private Function IsMarkTargetLine(ByVal lineText As String) As Boolean
    ' 現行ソースシートのB列マーキング対象
    ' ※ 仕様どおり、文字列ベースの部分一致判定を採用
    If Len(Trim$(lineText)) = 0 Then
        Exit Function
    End If

    If IsElseIfLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsElseLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsIfLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsTernaryLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsForeachLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsForLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsWhileLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsSwitchLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsCaseLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsDefaultLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsFunctionLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If
End Function

Private Function IsIfLine(ByVal lineText As String) As Boolean
    ' else if は別扱いなので除外
    If IsElseIfLine(lineText) Then Exit Function
    IsIfLine = ContainsText(lineText, "if")
End Function

Private Function IsElseIfLine(ByVal lineText As String) As Boolean
    IsElseIfLine = ContainsText(lineText, "else if")
End Function

Private Function IsElseLine(ByVal lineText As String) As Boolean
    IsElseLine = ContainsText(lineText, "else")
End Function

Private Function IsTernaryLine(ByVal lineText As String) As Boolean
    IsTernaryLine = (InStr(1, lineText, "?", vbBinaryCompare) > 0 And _
                     InStr(1, lineText, ":", vbBinaryCompare) > 0)
End Function

Private Function IsForeachLine(ByVal lineText As String) As Boolean
    IsForeachLine = ContainsText(lineText, "foreach")
End Function

Private Function IsForLine(ByVal lineText As String) As Boolean
    ' foreach とは区別する
    If IsForeachLine(lineText) Then Exit Function
    IsForLine = ContainsText(lineText, "for")
End Function

Private Function IsWhileLine(ByVal lineText As String) As Boolean
    IsWhileLine = ContainsText(lineText, "while")
End Function

Private Function IsSwitchLine(ByVal lineText As String) As Boolean
    IsSwitchLine = ContainsText(lineText, "switch")
End Function

Private Function IsCaseLine(ByVal lineText As String) As Boolean
    ' 初版は「case」と「:」を含むで判定
    IsCaseLine = ContainsText(lineText, "case") And (InStr(1, lineText, ":", vbBinaryCompare) > 0)
End Function

Private Function IsDefaultLine(ByVal lineText As String) As Boolean
    IsDefaultLine = ContainsText(lineText, "default:")
End Function

Private Function IsFunctionLine(ByVal lineText As String) As Boolean
    IsFunctionLine = ContainsText(lineText, "function")
End Function

Private Function ContainsText(ByVal sourceText As String, ByVal findText As String) As Boolean
    ' 大文字小文字を無視した部分一致
    If Len(findText) = 0 Then
        ContainsText = False
    Else
        ContainsText = (InStr(1, sourceText, findText, vbTextCompare) > 0)
    End If
End Function

Private Function NewEvent(ByVal eventKind As String) As Collection
    ' 疑似イベントオブジェクト（Collection + Key）を生成
    Dim ev As Collection
    Set ev = New Collection
    ev.Add eventKind, "Kind"
    Set NewEvent = ev
End Function

Private Function EventText(ByVal ev As Collection, ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    ' Collectionのキー取得（文字列）
    On Error GoTo UseDefault
    EventText = CStr(ev.Item(keyName))
    Exit Function

UseDefault:
    EventText = defaultValue
End Function

Private Function EventFlag(ByVal ev As Collection, ByVal keyName As String) As Boolean
    ' Collectionのキー取得（Boolean）
    On Error GoTo UseFalse
    EventFlag = CBool(ev.Item(keyName))
    Exit Function

UseFalse:
    EventFlag = False
End Function

Private Function EventCollection(ByVal ev As Collection, ByVal keyName As String) As Collection
    ' Collectionのキー取得（Collectionオブジェクト）
    On Error GoTo NoCollection
    Set EventCollection = ev.Item(keyName)
    Exit Function

NoCollection:
    Set EventCollection = Nothing
End Function

