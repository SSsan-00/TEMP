Attribute VB_Name = "modEvidenceSheetGenerator"
Option Explicit

' ============================================================
' 単体テストエビデンス シート生成マクロ
' ------------------------------------------------------------
' このモジュールは、マクロブック（ThisWorkbook）にある雛形シートを使って、
' ユーザーが選択したターゲットブック（xlsx）へエビデンスシートを生成する
'
' 主な流れ
' 1. 対象xlsxファイルを選択する
' 2. 入力ファイル名（例: foo.php）を入力する
' 3. マクロブックのREFERシートを参照して referValue を取得する
' 4. ターゲットブックの【共通】/【個別】参照元シートを走査する
' 5. 雛形A1を複製してエビデンスシートを作成する
' 6. 共通モードのみ A1-1-1 を先頭に貼り付け、〇〇〇 を置換する
' 7. B/C列の値をスロット規則に従って書き込む
' ============================================================

' ===== マクロブック内の固定シート名 =====
Private Const TEMPLATE_HEADER_SHEET_NAME As String = "A1-1-1" ' 共通モードの先頭上書きテンプレ
Private Const TEMPLATE_BODY_SHEET_NAME As String = "A1"       ' エビデンスシート本体テンプレ
Private Const REFER_SHEET_NAME As String = "REFER"            ' 入力ファイル名 -> referValue 参照用

' ===== REFERシートの列定義（列記号 -> 列番号へ変換して使用） =====
' 注意:
' Cells(row, col) の第2引数に "E" のような列記号文字列を直接渡すと
' 実行時エラーになる環境があるため、必ず列番号へ変換してから使う
Private Const REFER_KEY_COL_LETTER As String = "E"   ' 入力ファイル名のキー列（完全一致）
Private Const REFER_VALUE_COL_LETTER As String = "F" ' referValue を取得する列

' ===== 参照元シートの走査条件 =====
Private Const SOURCE_START_ROW As Long = 8          ' 仕様にある開始行
Private Const SOURCE_COL_A As Long = 1              ' A列: 作成するエビデンスシート名
Private Const SOURCE_COL_B As Long = 2              ' B列: pendingB 用
Private Const SOURCE_COL_C As Long = 3              ' C列: 確定トリガ
Private Const EMPTY_STREAK_STOP_COUNT As Long = 20  ' A/B/C空行が連続したら走査終了

' ===== エビデンスシートへの書き込み（スロット） =====
Private Const FIRST_DEST_ROW As Long = 3 ' slot0 の書き込み開始行
Private Const SLOT_HEIGHT As Long = 30   ' 30行刻み
Private Const DEST_COL_A As Long = 1     ' 書き込み先 A列
Private Const DEST_COL_B As Long = 2     ' 書き込み先 B列
Private Const DEST_COL_C As Long = 3     ' 書き込み先 C列

' ===== 共通モードヘッダ置換 =====
Private Const HEADER_PLACEHOLDER As String = "〇〇〇"

' ===== Office定数を数値で扱う（参照設定に依存しにくくするため） =====
Private Const FILE_DIALOG_PICKER As Long = 3 ' msoFileDialogFilePicker

' ============================================================
' エントリポイント
' ============================================================

Public Sub RunMain()
    ' 既存の呼び出し名からでも実行できるようにするための入口
    GenerateEvidenceSheets
End Sub

Public Sub GenerateEvidenceSheets()
    On Error GoTo ErrorHandler

    Dim macroWb As Workbook
    Dim targetWb As Workbook
    Dim referWs As Worksheet
    Dim templateBodyWs As Worksheet
    Dim templateHeaderWs As Worksheet

    Dim targetPath As String
    Dim inputFileName As String
    Dim baseName As String
    Dim referValue As String

    Dim commonSourceSheetName As String
    Dim commonSourceWs As Worksheet
    Dim individualSourceWs As Worksheet

    Dim commonSummary As String
    Dim individualSummary As String
    Dim finalMessage As String
    Dim processedAnyMode As Boolean

    ' Application状態は、エラー時でも必ず元に戻す
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim appStateCaptured As Boolean

    Set macroWb = ThisWorkbook

    ' まず必要なテンプレ/REFERシートが存在するか確認して、以降の処理を分かりやすく失敗させる
    Set templateBodyWs = GetWorksheetOrRaise(macroWb, TEMPLATE_BODY_SHEET_NAME, "雛形シート（本体）")
    Set templateHeaderWs = GetWorksheetOrRaise(macroWb, TEMPLATE_HEADER_SHEET_NAME, "雛形シート（ヘッダー）")
    Set referWs = GetWorksheetOrRaise(macroWb, REFER_SHEET_NAME, "REFERシート")

    ' 処理対象のターゲットブック（xlsx）を選択する
    targetPath = SelectTargetWorkbookPath()
    If Len(targetPath) = 0 Then
        MsgBox "処理をキャンセルしました（ターゲットブックが未選択です）。", vbInformation
        Exit Sub
    End If

    ' REFER検索キーになる入力ファイル名を受け取る（例: foo.php）。
    inputFileName = PromptInputFileName()
    If Len(inputFileName) = 0 Then
        MsgBox "処理をキャンセルしました（入力ファイル名が未入力です）。", vbInformation
        Exit Sub
    End If

    ' 後続処理で共通/個別シート名や置換に使うため、拡張子なし名を作成する
    baseName = RemoveExtension(inputFileName)
    If Len(baseName) = 0 Then
        Err.Raise vbObjectError + 2001, "GenerateEvidenceSheets", _
                  "入力ファイル名から拡張子なしの名前を取得できませんでした。"
    End If

    ' ターゲットブックを開く（既に開いていればそのインスタンスを再利用）。
    Set targetWb = OpenTargetWorkbook(targetPath)
    If targetWb Is Nothing Then
        Err.Raise vbObjectError + 2002, "GenerateEvidenceSheets", _
                  "ターゲットブックを開けませんでした。"
    End If

    If targetWb.ReadOnly Then
        Err.Raise vbObjectError + 2003, "GenerateEvidenceSheets", _
                  "ターゲットブックが読み取り専用で開かれているため、更新できません。"
    End If

    ' REFERシートから referValue を取得する（キーは入力ファイル名）。
    referValue = GetReferValueFromReferSheet(referWs, inputFileName)

    ' 速度改善のため、画面更新や再計算を一時的に止める
    ' 大量のシートコピーやセル書き込みで体感速度が大きく変わりる
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' -------------------------
    ' 共通モード（【共通】）
    ' -------------------------
    commonSourceSheetName = "【共通】" & referValue
    Set commonSourceWs = FindWorksheetExact(targetWb, commonSourceSheetName)

    If commonSourceWs Is Nothing Then
        commonSummary = "共通モード: スキップ（参照元シートなし: " & commonSourceSheetName & "）"
    Else
        commonSummary = ProcessReferenceSheet( _
            sourceWs:=commonSourceWs, _
            targetWb:=targetWb, _
            templateBodyWs:=templateBodyWs, _
            templateHeaderWs:=templateHeaderWs, _
            baseName:=baseName, _
            applyHeaderOverlay:=True, _
            modeLabel:="共通")
        processedAnyMode = True
    End If

    ' -------------------------
    ' 個別モード（【個別】）
    ' 優先順:
    ' 1) 【個別】baseName
    ' 2) baseName
    ' -------------------------
    Set individualSourceWs = FindIndividualSourceSheet(targetWb, baseName)
    If individualSourceWs Is Nothing Then
        individualSummary = "個別モード: スキップ（参照元シートなし: 【個別】" & baseName & " / " & baseName & "）"
    Else
        individualSummary = ProcessReferenceSheet( _
            sourceWs:=individualSourceWs, _
            targetWb:=targetWb, _
            templateBodyWs:=templateBodyWs, _
            templateHeaderWs:=templateHeaderWs, _
            baseName:=baseName, _
            applyHeaderOverlay:=False, _
            modeLabel:="個別")
        processedAnyMode = True
    End If

    ' 何か1つでもモードを処理したら保存します。
    If processedAnyMode Then
        targetWb.Save
    End If

    finalMessage = "処理が完了しました。" & vbCrLf & _
                   "ターゲットブック: " & targetWb.Name & vbCrLf & _
                   "入力ファイル名: " & inputFileName & vbCrLf & _
                   "baseName: " & baseName & vbCrLf & _
                   "REFER(F): " & referValue & vbCrLf & vbCrLf & _
                   commonSummary & vbCrLf & _
                   individualSummary

    GoTo SafeExit

ErrorHandler:
    finalMessage = "エラーが発生しました。" & vbCrLf & _
                   Err.Number & " : " & Err.Description

SafeExit:
    ' エラー時でも必ずアプリ状態を戻す
    On Error Resume Next
    Application.CutCopyMode = False
    If appStateCaptured Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Application.EnableEvents = prevEnableEvents
        Application.Calculation = prevCalculation
    End If
    On Error GoTo 0

    If Len(finalMessage) > 0 Then
        ' エラー/正常終了のどちらでも、ユーザーが次に何を見ればいいか分かるようにメッセージを返す
        If Left$(finalMessage, 6) = "エラーが発生" Then
            MsgBox finalMessage, vbExclamation
        Else
            MsgBox finalMessage, vbInformation
        End If
    End If
End Sub

' ============================================================
' 入力・ブック取得
' ============================================================

Private Function SelectTargetWorkbookPath() As String
    ' FileDialog を使って、更新対象の xlsx をユーザーに選ばせる
    ' 参照設定依存を避けるため、FileDialog型ではなく Object で扱う
    Dim fd As Object

    On Error GoTo Fallback

    Set fd = Application.FileDialog(FILE_DIALOG_PICKER)
    With fd
        .Title = "対象のxlsxファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック (*.xlsx)", "*.xlsx"

        If .Show <> -1 Then
            SelectTargetWorkbookPath = vbNullString
            Exit Function
        End If

        SelectTargetWorkbookPath = CStr(.SelectedItems(1))
    End With
    Exit Function

Fallback:
    ' 環境差で FileDialog が使えない場合に備え、GetOpenFilename にフォールバックする
    Dim selectedPath As Variant

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excel ブック (*.xlsx),*.xlsx", _
        Title:="対象のxlsxファイルを選択してください")

    If VarType(selectedPath) = vbBoolean Then
        SelectTargetWorkbookPath = vbNullString
    Else
        SelectTargetWorkbookPath = CStr(selectedPath)
    End If
End Function

Private Function PromptInputFileName() As String
    ' REFER検索キーになる入力ファイル名を受け取る
    ' 前後の空白は誤入力になりやすいため Trim する
    Dim s As String

    s = InputBox("入力ファイル名を入力してください（例: foo.php）", "入力ファイル名")
    PromptInputFileName = Trim$(s)
End Function

Private Function OpenTargetWorkbook(ByVal workbookPath As String) As Workbook
    ' 既に同じファイルが開いている場合は再利用し、未オープンなら開く
    Dim wb As Workbook

    If Len(Trim$(workbookPath)) = 0 Then Exit Function

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

' ============================================================
' REFER参照
' ============================================================

Private Function GetReferValueFromReferSheet( _
    ByVal referWs As Worksheet, _
    ByVal inputFileName As String) As String

    ' REFERシートから、入力ファイル名をキーに該当行を探し、F列の値を返す
    ' 仕様上、完全一致を前提にする
    Dim matchedRow As Long
    Dim matchCount As Long
    Dim valueColIndex As Long
    Dim referValueRaw As Variant

    matchedRow = FindRowByExactMatch(referWs, REFER_KEY_COL_LETTER, inputFileName, matchCount)

    If matchCount = 0 Then
        Err.Raise vbObjectError + 2101, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_KEY_COL_LETTER & "列に完全一致する値が見つかりませんでした。" & vbCrLf & _
                  "入力値: " & inputFileName
    End If

    If matchCount > 1 Then
        Err.Raise vbObjectError + 2102, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_KEY_COL_LETTER & "列に完全一致する値が複数あります。" & vbCrLf & _
                  "入力値: " & inputFileName & vbCrLf & _
                  "件数: " & CStr(matchCount)
    End If

    valueColIndex = ColumnLetterToIndex(REFER_VALUE_COL_LETTER, "REFER値列")
    referValueRaw = referWs.Cells(matchedRow, valueColIndex).Value

    If IsError(referValueRaw) Then
        Err.Raise vbObjectError + 2103, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_VALUE_COL_LETTER & "列にエラー値が入っています。"
    End If

    GetReferValueFromReferSheet = Trim$(CStr(referValueRaw))
    If Len(GetReferValueFromReferSheet) = 0 Then
        Err.Raise vbObjectError + 2104, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_VALUE_COL_LETTER & "列の値が空です。"
    End If
End Function

Private Function FindRowByExactMatch( _
    ByVal ws As Worksheet, _
    ByVal targetColLetter As String, _
    ByVal searchValue As String, _
    ByRef matchCount As Long) As Long

    ' 文字列の完全一致（vbBinaryCompare）で検索
    ' 大文字/小文字や全角/半角の違いも区別
    Dim targetColIndex As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As Variant
    Dim cellText As String

    matchCount = 0
    targetColIndex = ColumnLetterToIndex(targetColLetter, "検索列")
    lastRow = ws.Cells(ws.Rows.Count, targetColIndex).End(xlUp).Row

    If lastRow < 1 Then Exit Function

    For r = 1 To lastRow
        cellValue = ws.Cells(r, targetColIndex).Value

        If IsError(cellValue) Then
            Err.Raise vbObjectError + 2111, "FindRowByExactMatch", _
                      "REFERシートの検索列にエラー値が含まれています（行: " & CStr(r) & "）。"
        End If

        cellText = CStr(cellValue)
        If StrComp(cellText, searchValue, vbBinaryCompare) = 0 Then
            matchCount = matchCount + 1
            If FindRowByExactMatch = 0 Then
                FindRowByExactMatch = r
            End If
        End If
    Next r
End Function

Private Function ColumnLetterToIndex( _
    ByVal columnLetter As String, _
    Optional ByVal labelForError As String = vbNullString) As Long

    ' "A" -> 1, "F" -> 6, "AA" -> 27 のように列記号を列番号へ変換
    ' Cells(row, "F") のような文字列渡しを避けるための関数
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim chCode As Long
    Dim prefix As String

    normalized = UCase$(Trim$(columnLetter))
    If Len(normalized) = 0 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 2121, "ColumnLetterToIndex", prefix & "列指定が空です。"
    End If

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        chCode = Asc(ch)

        If chCode < 65 Or chCode > 90 Then
            prefix = BuildErrorLabelPrefix(labelForError)
            Err.Raise vbObjectError + 2122, "ColumnLetterToIndex", _
                      prefix & "列指定が不正です: " & columnLetter
        End If

        ColumnLetterToIndex = (ColumnLetterToIndex * 26) + (chCode - 64)
    Next i

    If ColumnLetterToIndex < 1 Or ColumnLetterToIndex > 16384 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 2123, "ColumnLetterToIndex", _
                  prefix & "列番号がExcelの範囲外です: " & columnLetter
    End If
End Function

Private Function BuildErrorLabelPrefix(ByVal labelText As String) As String
    If Len(Trim$(labelText)) = 0 Then
        BuildErrorLabelPrefix = vbNullString
    Else
        BuildErrorLabelPrefix = Trim$(labelText) & " "
    End If
End Function

' ============================================================
' 参照元シート探索
' ============================================================

Private Function FindWorksheetExact(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' シート名完全一致で取得します。見つからない場合は Nothing を返す
    On Error Resume Next
    Set FindWorksheetExact = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindIndividualSourceSheet(ByVal targetWb As Workbook, ByVal baseName As String) As Worksheet
    ' 個別モードの優先順:
    ' 1) 【個別】baseName
    ' 2) baseName
    Dim candidateName As String

    candidateName = "【個別】" & baseName
    Set FindIndividualSourceSheet = FindWorksheetExact(targetWb, candidateName)
    If Not FindIndividualSourceSheet Is Nothing Then Exit Function

    Set FindIndividualSourceSheet = FindWorksheetExact(targetWb, baseName)
End Function

' ============================================================
' 参照元シート -> エビデンスシート生成
' ============================================================

Private Function ProcessReferenceSheet( _
    ByVal sourceWs As Worksheet, _
    ByVal targetWb As Workbook, _
    ByVal templateBodyWs As Worksheet, _
    ByVal templateHeaderWs As Worksheet, _
    ByVal baseName As String, _
    ByVal applyHeaderOverlay As Boolean, _
    ByVal modeLabel As String) As String

    ' 参照元シート（共通または個別）を走査し、A/B/Cのルールに従って
    ' エビデンスシートを作成・更新する
    Dim r As Long
    Dim emptyStreak As Long

    Dim currentEvidenceWs As Worksheet
    Dim currentEvidenceSheetName As String

    Dim slotIndex As Long
    Dim hasPendingB As Boolean
    Dim pendingB As Variant

    Dim rawA As Variant
    Dim rawB As Variant
    Dim rawC As Variant

    Dim hasA As Boolean
    Dim hasB As Boolean
    Dim hasC As Boolean
    Dim aSheetName As String

    Dim createdSheetCount As Long
    Dim slotWriteCount As Long
    Dim ignoredDataBeforeSheetCount As Long

    r = SOURCE_START_ROW
    emptyStreak = 0
    Set currentEvidenceWs = Nothing
    currentEvidenceSheetName = vbNullString
    slotIndex = 0
    hasPendingB = False

    Do While emptyStreak < EMPTY_STREAK_STOP_COUNT
        rawA = sourceWs.Cells(r, SOURCE_COL_A).Value
        rawB = sourceWs.Cells(r, SOURCE_COL_B).Value
        rawC = sourceWs.Cells(r, SOURCE_COL_C).Value

        ' エラー値が紛れていると原因が分かりにくくなるため、行番号付きで即時中断する
        EnsureNotErrorValue rawA, sourceWs.Name, r, "A"
        EnsureNotErrorValue rawB, sourceWs.Name, r, "B"
        EnsureNotErrorValue rawC, sourceWs.Name, r, "C"

        hasA = HasValueForSourceCell(rawA)
        hasB = HasValueForSourceCell(rawB)
        hasC = HasValueForSourceCell(rawC)

        If (Not hasA) And (Not hasB) And (Not hasC) Then
            emptyStreak = emptyStreak + 1
        Else
            emptyStreak = 0
        End If

        ' A列に値が来たら、現在シートを切り替える
        ' その前に pendingB が残っていれば、前シートに B単体として確定させる
        If hasA Then
            If Not currentEvidenceWs Is Nothing Then
                FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
            End If

            aSheetName = NormalizeEvidenceSheetName(rawA, sourceWs.Name, r)

            Set currentEvidenceWs = RecreateEvidenceSheetFromTemplate( _
                targetWb:=targetWb, _
                templateBodyWs:=templateBodyWs, _
                newSheetName:=aSheetName, _
                currentSourceSheetName:=sourceWs.Name)

            createdSheetCount = createdSheetCount + 1
            currentEvidenceSheetName = aSheetName

            ' シートが変わったら、スロットと pendingB を新しいシート用に初期化する
            slotIndex = 0
            hasPendingB = False

            ' 共通モードのみ、A1-1-1 テンプレを先頭へ貼り付け、〇〇〇 を baseName に置換する
            If applyHeaderOverlay Then
                ApplyHeaderOverlayAndReplace templateHeaderWs, currentEvidenceWs, baseName
            End If
        End If

        ' B/C は「現在のエビデンスシート」が決まっている場合にのみ処理を行う
        ' Aがまだ一度も出ていない場合は、仕様に必要な書き込み先が未確定なのでスキップする
        If hasB Or hasC Then
            If currentEvidenceWs Is Nothing Then
                ignoredDataBeforeSheetCount = ignoredDataBeforeSheetCount + 1
            Else
                ' 先にBを pending として保持（同一行に C がある場合、直後の C でペア確定させるため）
                If hasB Then
                    ' Bが連続で来た場合は、前のpendingBを単体として確定してから新しいBを保持する
                    If hasPendingB Then
                        FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
                    End If

                    pendingB = rawB
                    hasPendingB = True
                End If

                If hasC Then
                    If hasPendingB Then
                        WritePairSlot currentEvidenceWs, slotIndex, pendingB, rawC
                        slotWriteCount = slotWriteCount + 1
                        slotIndex = slotIndex + 1
                        hasPendingB = False
                    Else
                        WriteCOnlySlot currentEvidenceWs, slotIndex, rawC
                        slotWriteCount = slotWriteCount + 1
                        slotIndex = slotIndex + 1
                    End If
                End If
            End If
        End If

        r = r + 1
    Loop

    ' 走査終了時にも pendingB が残っていれば、最後の1件を取りこぼさないよう確定させる
    If Not currentEvidenceWs Is Nothing Then
        FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
    End If

    ProcessReferenceSheet = modeLabel & "モード: 完了（参照元=" & sourceWs.Name & _
                           ", 作成シート数=" & CStr(createdSheetCount) & _
                           ", スロット書込数=" & CStr(slotWriteCount) & _
                           IIf(ignoredDataBeforeSheetCount > 0, _
                               ", 先行B/Cスキップ行=" & CStr(ignoredDataBeforeSheetCount), _
                               vbNullString) & ")"
End Function

' ============================================================
' エビデンスシート作成・テンプレ適用
' ============================================================

Private Function RecreateEvidenceSheetFromTemplate( _
    ByVal targetWb As Workbook, _
    ByVal templateBodyWs As Worksheet, _
    ByVal newSheetName As String, _
    ByVal currentSourceSheetName As String) As Worksheet

    ' 同名シートが既にある場合は削除して作り直す
    ' ただし、現在走査中の参照元シートは削除してはいけないので保護する
    ValidateWorksheetName newSheetName

    If StrComp(newSheetName, currentSourceSheetName, vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 2201, "RecreateEvidenceSheetFromTemplate", _
                  "参照元シート名と同じ名前のエビデンスシートは作成できません: " & newSheetName
    End If

    DeleteWorksheetIfExists targetWb, newSheetName, currentSourceSheetName

    ' 雛形シートA1（マクロブック側）をターゲットブックへコピーして、新しいエビデンスシートを作る
    templateBodyWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)
    Set RecreateEvidenceSheetFromTemplate = targetWb.Worksheets(targetWb.Worksheets.Count)

    On Error GoTo RenameError
    RecreateEvidenceSheetFromTemplate.Name = newSheetName
    On Error GoTo 0
    Exit Function

RenameError:
    Err.Raise vbObjectError + 2202, "RecreateEvidenceSheetFromTemplate", _
              "エビデンスシート名を設定できませんでした: " & newSheetName & vbCrLf & _
              "（シート名の文字数・使用禁止文字・重複を確認してください）"
End Function

Private Sub DeleteWorksheetIfExists( _
    ByVal wb As Workbook, _
    ByVal targetSheetName As String, _
    Optional ByVal protectedSheetName As String = vbNullString)

    ' 既存シート削除用ヘルパー。
    ' DisplayAlerts は上位で OFF にしている前提だが、ここでは警告表示の制御は行わない
    Dim ws As Worksheet

    Set ws = FindWorksheetExact(wb, targetSheetName)
    If ws Is Nothing Then Exit Sub

    If Len(protectedSheetName) > 0 Then
        If StrComp(ws.Name, protectedSheetName, vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 2211, "DeleteWorksheetIfExists", _
                      "保護対象のシートを削除しようとしました: " & ws.Name
        End If
    End If

    ' マクロブックの雛形シートは絶対に削除しないよう、念のためガードします。
    If wb Is ThisWorkbook Then
        If StrComp(ws.Name, TEMPLATE_BODY_SHEET_NAME, vbBinaryCompare) = 0 Or _
           StrComp(ws.Name, TEMPLATE_HEADER_SHEET_NAME, vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 2212, "DeleteWorksheetIfExists", _
                      "マクロブックの雛形シートは削除できません: " & ws.Name
        End If
    End If

    ws.Delete
End Sub

Private Sub ApplyHeaderOverlayAndReplace( _
    ByVal headerTemplateWs As Worksheet, _
    ByVal destEvidenceWs As Worksheet, _
    ByVal baseName As String)

    ' 共通モード専用処理:
    ' A1-1-1 テンプレを A1 起点で貼り付けて、貼り付け範囲の "〇〇〇" を baseName に置換する
    Dim srcRange As Range
    Dim pastedRange As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set srcRange = headerTemplateWs.UsedRange
    If srcRange Is Nothing Then Exit Sub

    rowCount = srcRange.Rows.Count
    colCount = srcRange.Columns.Count
    If rowCount <= 0 Or colCount <= 0 Then Exit Sub

    srcRange.Copy Destination:=destEvidenceWs.Range("A1")
    Set pastedRange = destEvidenceWs.Range("A1").Resize(rowCount, colCount)

    pastedRange.Replace What:=HEADER_PLACEHOLDER, _
                        Replacement:=baseName, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        MatchCase:=False
End Sub

' ============================================================
' 参照元 A/B/C の読み取り補助
' ============================================================

Private Sub EnsureNotErrorValue( _
    ByVal cellValue As Variant, _
    ByVal sheetName As String, _
    ByVal rowNumber As Long, _
    ByVal colLetter As String)

    If IsError(cellValue) Then
        Err.Raise vbObjectError + 2301, "EnsureNotErrorValue", _
                  "参照元シートにエラー値が含まれています。" & vbCrLf & _
                  "シート: " & sheetName & " / セル: " & colLetter & CStr(rowNumber)
    End If
End Sub

Private Function HasValueForSourceCell(ByVal cellValue As Variant) As Boolean
    ' A/B/C列の「値あり判定」。
    ' 文字列は Trim 後に空なら空扱い、数値は 0 でも値あり扱いにする
    If IsEmpty(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function

    If VarType(cellValue) = vbString Then
        HasValueForSourceCell = (Len(Trim$(CStr(cellValue))) > 0)
    Else
        HasValueForSourceCell = (Len(CStr(cellValue)) > 0)
    End If
End Function

Private Function NormalizeEvidenceSheetName( _
    ByVal rawValue As Variant, _
    ByVal sourceSheetName As String, _
    ByVal rowNumber As Long) As String

    ' A列の値をシート名として使うため、文字列化＋前後空白除去を行う
    ' 空になってしまう場合は呼び出し元のロジックと矛盾するため、明示的にエラーにする
    NormalizeEvidenceSheetName = Trim$(CStr(rawValue))

    If Len(NormalizeEvidenceSheetName) = 0 Then
        Err.Raise vbObjectError + 2311, "NormalizeEvidenceSheetName", _
                  "A列のシート名が空です（シート: " & sourceSheetName & ", 行: " & CStr(rowNumber) & "）。"
    End If
End Function

Private Sub ValidateWorksheetName(ByVal sheetNameText As String)
    ' Excelシート名として明らかに不正な値は、コピー/リネーム前に弾いて原因を明確にする
    Dim invalidChars As Variant
    Dim i As Long

    If Len(sheetNameText) = 0 Then
        Err.Raise vbObjectError + 2321, "ValidateWorksheetName", "シート名が空です。"
    End If

    If Len(sheetNameText) > 31 Then
        Err.Raise vbObjectError + 2322, "ValidateWorksheetName", _
                  "シート名は31文字以内である必要があります: " & sheetNameText
    End If

    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(1, sheetNameText, CStr(invalidChars(i)), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2323, "ValidateWorksheetName", _
                      "シート名に使用できない文字が含まれています: " & CStr(invalidChars(i))
        End If
    Next i
End Sub

' ============================================================
' スロット書き込み（B/C -> エビデンスシート）
' ============================================================

Private Sub FlushPendingBIfNeeded( _
    ByVal destWs As Worksheet, _
    ByRef slotIndex As Long, _
    ByRef hasPendingB As Boolean, _
    ByRef pendingB As Variant, _
    ByRef slotWriteCount As Long)

    ' pendingB が残っている場合、仕様どおり B単体 として1スロット書き込む
    If Not hasPendingB Then Exit Sub

    WriteBOnlySlot destWs, slotIndex, pendingB
    slotWriteCount = slotWriteCount + 1
    slotIndex = slotIndex + 1
    hasPendingB = False
End Sub

Private Sub WritePairSlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal pendingB As Variant, _
    ByVal cValue As Variant)

    ' ペア書き込みルール:
    ' - slot0   : B -> A列, C -> B列
    ' - slot1以降: B -> B列, C -> C列
    Dim destRow As Long
    Dim bDestCol As Long
    Dim cDestCol As Long

    destRow = GetDestRowForSlot(slotIndex)

    If slotIndex = 0 Then
        bDestCol = DEST_COL_A
        cDestCol = DEST_COL_B
    Else
        bDestCol = DEST_COL_B
        cDestCol = DEST_COL_C
    End If

    destWs.Cells(destRow, bDestCol).Value = pendingB
    destWs.Cells(destRow, cDestCol).Value = cValue
End Sub

Private Sub WriteCOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal cValue As Variant)

    ' C単体は、どのスロットでも B列に書き込む
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)
    destWs.Cells(destRow, DEST_COL_B).Value = cValue
End Sub

Private Sub WriteBOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal bValue As Variant)

    ' B単体は:
    ' - slot0   : A列
    ' - slot1以降: B列
    Dim destRow As Long
    Dim destCol As Long

    destRow = GetDestRowForSlot(slotIndex)
    If slotIndex = 0 Then
        destCol = DEST_COL_A
    Else
        destCol = DEST_COL_B
    End If

    destWs.Cells(destRow, destCol).Value = bValue
End Sub

Private Function GetDestRowForSlot(ByVal slotIndex As Long) As Long
    If slotIndex < 0 Then
        Err.Raise vbObjectError + 2401, "GetDestRowForSlot", "slotIndex が負数です。"
    End If

    GetDestRowForSlot = FIRST_DEST_ROW + (slotIndex * SLOT_HEIGHT)
End Function

' ============================================================
' 共通ユーティリティ
' ============================================================

Private Function GetWorksheetOrRaise( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByVal labelForMessage As String) As Worksheet

    Set GetWorksheetOrRaise = FindWorksheetExact(wb, sheetName)
    If GetWorksheetOrRaise Is Nothing Then
        Err.Raise vbObjectError + 2501, "GetWorksheetOrRaise", _
                  labelForMessage & " が見つかりません: " & sheetName
    End If
End Function

Private Function RemoveExtension(ByVal fileNameText As String) As String
    ' "foo.php" -> "foo"
    ' "foo.bar.php" -> "foo.bar"
    ' "foo" -> "foo"
    ' パスが混ざっていても最後の区切り以降だけを対象にする
    Dim lastDotPos As Long
    Dim lastSlashPos As Long
    Dim lastBackslashPos As Long
    Dim lastSeparatorPos As Long

    fileNameText = Trim$(fileNameText)
    If Len(fileNameText) = 0 Then Exit Function

    lastSlashPos = InStrRev(fileNameText, "/")
    lastBackslashPos = InStrRev(fileNameText, "\")
    If lastSlashPos > lastBackslashPos Then
        lastSeparatorPos = lastSlashPos
    Else
        lastSeparatorPos = lastBackslashPos
    End If

    lastDotPos = InStrRev(fileNameText, ".")
    If lastDotPos > (lastSeparatorPos + 1) Then
        RemoveExtension = Left$(fileNameText, lastDotPos - 1)
    Else
        RemoveExtension = fileNameText
    End If
End Function
