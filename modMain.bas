Attribute VB_Name = "modMain"
Option Explicit

' =============================================================================
' modMain
'   操作用 .xlsm から外部 .xlsx を編集するエントリポイント。
' =============================================================================

' 走査対象設定:
'   ここを変更すれば、対象シート・列・開始行をまとめて変更できる。
Private Const TARGET_SHEET_NAME As String = "ソース"
Private Const LEFT_COLUMN As String = "C"
Private Const RIGHT_COLUMN As String = "D"
Private Const START_ROW As Long = 5
Private Const OPERATION_SHEET_NAME As String = "操作"

Private Type AppState
  ScreenUpdating As Boolean
  EnableEvents As Boolean
  Calculation As XlCalculation
End Type

Public Sub Run()
  ' メイン実行: ファイル選択 → バックアップ → C/D 列整形 → 保存。
  Dim filePath As String

  On Error GoTo ErrorHandler

  filePath = PickTargetXlsxPath()
  If Len(filePath) = 0 Then
    Exit Sub
  End If

  RunAlignmentForFile filePath
  Exit Sub

ErrorHandler:
  MsgBox "整形処理でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

Public Sub SetupOperationSheet()
  ' 操作用シートが無ければ作成し、「整形実行」ボタンを配置する。
  Dim ws As Worksheet
  Dim btn As Button

  Set ws = GetOrCreateSheet(ThisWorkbook, OPERATION_SHEET_NAME)

  ws.Range("A1").Value = "外部 .xlsx の " & LEFT_COLUMN & CStr(START_ROW) & ":" & RIGHT_COLUMN & " を整形します。"
  ws.Range("A2").Value = "ボタンを押すとファイル選択ダイアログが開きます。"

  On Error Resume Next
  Set btn = ws.Buttons("btnRunAlign")
  On Error GoTo 0

  If btn Is Nothing Then
    Set btn = ws.Buttons.Add(ws.Range("A4").Left, ws.Range("A4").Top, 140, 32)
    btn.Name = "btnRunAlign"
  End If

  btn.Characters.Text = "整形実行"
  btn.OnAction = "'" & ThisWorkbook.Name & "'!Run"
End Sub

Public Sub RunAlignmentForFile(ByVal targetPath As String)
  Dim state As AppState
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim alreadyOpen As Boolean
  Dim openedByMacro As Boolean
  Dim backupPath As String

  On Error GoTo ErrorHandler

  EnterFastMode state

  Set wb = FindOpenWorkbook(targetPath)
  alreadyOpen = Not wb Is Nothing

  If wb Is Nothing Then
    Set wb = Workbooks.Open(Filename:=targetPath, ReadOnly:=False, UpdateLinks:=0)
    openedByMacro = True
  End If

  backupPath = BuildBackupPath(targetPath)
  wb.SaveCopyAs backupPath

  Set ws = Nothing
  On Error Resume Next
  Set ws = wb.Worksheets(TARGET_SHEET_NAME)
  On Error GoTo ErrorHandler

  If ws Is Nothing Then
    Err.Raise vbObjectError + 2000, "RunAlignmentForFile", "対象シート「" & TARGET_SHEET_NAME & "」が見つかりません。"
  End If

  AlignSourceSheet ws

  wb.Save

  If openedByMacro Then
    wb.Close SaveChanges:=False
  End If

  LeaveFastMode state

  MsgBox "整形が完了しました。" & vbCrLf & _
         "保存先: " & targetPath & vbCrLf & _
         "バックアップ: " & backupPath, vbInformation
  Exit Sub

ErrorHandler:
  On Error Resume Next
  If Not wb Is Nothing Then
    If openedByMacro Then
      wb.Close SaveChanges:=False
    ElseIf Not alreadyOpen Then
      wb.Close SaveChanges:=False
    End If
  End If
  LeaveFastMode state
  On Error GoTo 0

  MsgBox "整形処理を中断しました: " & Err.Description, vbExclamation
End Sub

Private Sub AlignSourceSheet(ByVal ws As Worksheet)
  Dim lastRowC As Long
  Dim lastRowD As Long
  Dim lastRow As Long
  Dim sourceValues As Variant
  Dim rowCount As Long
  Dim i As Long
  Dim leftOriginal() As String
  Dim rightOriginal() As String
  Dim alignedLeft() As String
  Dim alignedRight() As String
  Dim alignedCount As Long
  Dim writeValues As Variant
  Dim writeLastRow As Long
  Dim clearLastRow As Long

  lastRowC = LastUsedRowInColumn(ws, LEFT_COLUMN)
  lastRowD = LastUsedRowInColumn(ws, RIGHT_COLUMN)
  lastRow = MaxLong(lastRowC, lastRowD)

  If lastRow < START_ROW Then
    MsgBox LEFT_COLUMN & CStr(START_ROW) & ":" & RIGHT_COLUMN & " に整形対象データがありません。", vbInformation
    Exit Sub
  End If

  sourceValues = ws.Range(SourceRangeAddress(lastRow)).Value2
  rowCount = UBound(sourceValues, 1)

  ReDim leftOriginal(0 To rowCount - 1)
  ReDim rightOriginal(0 To rowCount - 1)

  For i = 1 To rowCount
    leftOriginal(i - 1) = VariantToString(sourceValues(i, 1))
    rightOriginal(i - 1) = VariantToString(sourceValues(i, 2))
  Next i

  AlignColumnsByDiff leftOriginal, rightOriginal, alignedLeft, alignedRight
  alignedCount = GetStringArrayCount(alignedLeft)

  If alignedCount = 0 Then
    ws.Range(SourceRangeAddress(lastRow)).ClearContents
    Exit Sub
  End If

  ReDim writeValues(1 To alignedCount, 1 To 2)
  For i = 1 To alignedCount
    writeValues(i, 1) = alignedLeft(i - 1)
    writeValues(i, 2) = alignedRight(i - 1)
  Next i

  writeLastRow = (START_ROW - 1) + alignedCount
  clearLastRow = MaxLong(lastRow, writeLastRow)

  ws.Range(SourceRangeAddress(clearLastRow)).ClearContents
  ws.Range(SourceRangeAddress(writeLastRow)).Value2 = writeValues
End Sub

Private Function PickTargetXlsxPath() As String
  Dim fd As FileDialog

  Set fd = Application.FileDialog(msoFileDialogFilePicker)
  With fd
    .Title = "整形対象の .xlsx を選択してください"
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "Excel Workbook", "*.xlsx"

    If .Show <> -1 Then
      PickTargetXlsxPath = ""
      Exit Function
    End If

    PickTargetXlsxPath = .SelectedItems(1)
  End With
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
  Dim wb As Workbook

  For Each wb In Application.Workbooks
    If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
      Set FindOpenWorkbook = wb
      Exit Function
    End If
  Next wb

  Set FindOpenWorkbook = Nothing
End Function

Private Function BuildBackupPath(ByVal targetPath As String) As String
  Dim separator As String
  Dim folderPath As String
  Dim fileName As String
  Dim dotPos As Long
  Dim baseName As String
  Dim stamp As String

  separator = Application.PathSeparator
  folderPath = Left$(targetPath, InStrRev(targetPath, separator))
  fileName = Mid$(targetPath, InStrRev(targetPath, separator) + 1)

  dotPos = InStrRev(fileName, ".")
  If dotPos > 1 Then
    baseName = Left$(fileName, dotPos - 1)
  Else
    baseName = fileName
  End If

  stamp = Format$(Now, "yyyymmdd_Hhnnss")
  BuildBackupPath = folderPath & baseName & "_" & stamp & ".bak.xlsx"
End Function

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal colLetter As String) As Long
  LastUsedRowInColumn = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
End Function

Private Function SourceRangeAddress(ByVal endRow As Long) As String
  SourceRangeAddress = LEFT_COLUMN & CStr(START_ROW) & ":" & RIGHT_COLUMN & CStr(endRow)
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
  If a > b Then
    MaxLong = a
  Else
    MaxLong = b
  End If
End Function

Private Function VariantToString(ByVal value As Variant) As String
  If IsError(value) Then
    VariantToString = ""
  ElseIf IsNull(value) Then
    VariantToString = ""
  ElseIf IsEmpty(value) Then
    VariantToString = ""
  Else
    VariantToString = CStr(value)
  End If
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
  Dim ws As Worksheet

  On Error Resume Next
  Set ws = wb.Worksheets(sheetName)
  On Error GoTo 0

  If ws Is Nothing Then
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
  End If

  Set GetOrCreateSheet = ws
End Function

Private Sub EnterFastMode(ByRef state As AppState)
  state.ScreenUpdating = Application.ScreenUpdating
  state.EnableEvents = Application.EnableEvents
  state.Calculation = Application.Calculation

  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
End Sub

Private Sub LeaveFastMode(ByRef state As AppState)
  On Error Resume Next
  Application.ScreenUpdating = state.ScreenUpdating
  Application.EnableEvents = state.EnableEvents
  Application.Calculation = state.Calculation
  On Error GoTo 0
End Sub
