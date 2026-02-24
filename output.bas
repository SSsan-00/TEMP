Attribute VB_Name = "modEvidenceFileGenerator"
Option Explicit

Private Const REFER_SHEET_NAME As String = "REFER"
Private Const MATCH_COL As String = "E"
Private Const COL_BETA As String = "D"
Private Const COL_GAMMA As String = "F"
Private Const COL_ALPHA As String = "J"

Public Sub RunMain()
    GenerateEvidenceFiles
End Sub

Public Sub GenerateEvidenceFiles()
    On Error GoTo ErrorHandler

    Dim inputFileName As String
    Dim referSheet As Worksheet
    Dim matchedRow As Long
    Dim matchCount As Long
    Dim alpha As String
    Dim beta As String
    Dim gamma As String
    Dim commonFileName As String
    Dim individualFileName As String
    Dim outputFolder As String
    Dim commonPath As String
    Dim individualPath As String

    inputFileName = PromptLookupFileName()
    If Len(inputFileName) = 0 Then
        MsgBox "処理をキャンセルしました（ファイル名が未入力です）。", vbInformation
        Exit Sub
    End If

    Set referSheet = GetReferSheet(ThisWorkbook)
    If referSheet Is Nothing Then
        MsgBox "REFERシートが見つかりません。", vbExclamation
        Exit Sub
    End If

    matchedRow = FindRowByExactMatch(referSheet, MATCH_COL, inputFileName, matchCount)
    If matchCount = 0 Then
        MsgBox "REFERシートのE列に完全一致する値が見つかりませんでした。" & vbCrLf & _
               "入力値: " & inputFileName, vbExclamation
        Exit Sub
    End If

    If matchCount > 1 Then
        MsgBox "REFERシートのE列に完全一致する値が複数あります。" & vbCrLf & _
               "入力値: " & inputFileName & vbCrLf & _
               "件数: " & CStr(matchCount), vbExclamation
        Exit Sub
    End If

    alpha = RemoveExtension(GetCellString(referSheet.Cells(matchedRow, COL_ALPHA).Value))
    If Len(alpha) = 0 Then
        MsgBox "J列の値が空のため、ファイル名を生成できません。", vbExclamation
        Exit Sub
    End If

    beta = ToTwoDigitString(referSheet.Cells(matchedRow, COL_BETA).Value)
    If Len(beta) = 0 Then
        MsgBox "D列の値を""00""形式に変換できません。数値（例: 1, 2, 10）を設定してください。", vbExclamation
        Exit Sub
    End If

    gamma = GetCellString(referSheet.Cells(matchedRow, COL_GAMMA).Value)
    If Len(gamma) = 0 Then
        MsgBox "F列の値が空のため、ファイル名を生成できません。", vbExclamation
        Exit Sub
    End If

    commonFileName = alpha & "_【共通】" & beta & gamma & "_単体テストエビデンス_初期開発.xlsx"
    individualFileName = alpha & "_【個別】" & beta & gamma & "_単体テストエビデンス_初期開発.xlsx"

    ValidateWindowsFileName commonFileName
    ValidateWindowsFileName individualFileName

    outputFolder = ThisWorkbook.Path
    If Len(outputFolder) = 0 Then
        MsgBox "このxlsmファイルが未保存のため、出力先フォルダを特定できません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    commonPath = BuildPath(outputFolder, commonFileName)
    individualPath = BuildPath(outputFolder, individualFileName)

    CreateEmptyXlsxFile commonPath
    CreateEmptyXlsxFile individualPath

    MsgBox "ファイルを生成しました。" & vbCrLf & _
           "共通: " & commonFileName & vbCrLf & _
           "個別: " & individualFileName, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation
End Sub

Private Function PromptLookupFileName() As String
    Dim inputValue As String
    inputValue = InputBox("REFERシートのE列に完全一致するファイル名を入力してください。", "ファイル名入力")
    PromptLookupFileName = Trim$(inputValue)
End Function

Private Function GetReferSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetReferSheet = wb.Worksheets(REFER_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function FindRowByExactMatch( _
    ByVal ws As Worksheet, _
    ByVal targetCol As String, _
    ByVal searchValue As String, _
    ByRef matchCount As Long) As Long

    Dim lastRow As Long
    Dim r As Long
    Dim cellText As String

    matchCount = 0
    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row

    If lastRow < 1 Then Exit Function

    For r = 1 To lastRow
        If IsError(ws.Cells(r, targetCol).Value) Then
            Err.Raise vbObjectError + 1002, "FindRowByExactMatch", "REFERシートの検索列にエラー値が含まれています。"
        End If
        cellText = CStr(ws.Cells(r, targetCol).Value)
        If StrComp(cellText, searchValue, vbBinaryCompare) = 0 Then
            matchCount = matchCount + 1
            If FindRowByExactMatch = 0 Then
                FindRowByExactMatch = r
            End If
        End If
    Next r
End Function

Private Function GetCellString(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        Err.Raise vbObjectError + 1001, "GetCellString", "セルにエラー値が含まれています。"
    End If
    GetCellString = Trim$(CStr(cellValue))
End Function

Private Function RemoveExtension(ByVal fileNameText As String) As String
    Dim lastDotPos As Long
    Dim lastSlashPos As Long
    Dim lastBackslashPos As Long
    Dim lastSepPos As Long

    lastSlashPos = InStrRev(fileNameText, "/")
    lastBackslashPos = InStrRev(fileNameText, "\")
    If lastSlashPos > lastBackslashPos Then
        lastSepPos = lastSlashPos
    Else
        lastSepPos = lastBackslashPos
    End If

    lastDotPos = InStrRev(fileNameText, ".")
    If lastDotPos > lastSepPos + 1 Then
        RemoveExtension = Left$(fileNameText, lastDotPos - 1)
    Else
        RemoveExtension = fileNameText
    End If
End Function

Private Function ToTwoDigitString(ByVal valueD As Variant) As String
    Dim numericValue As Double

    If IsError(valueD) Or IsEmpty(valueD) Then Exit Function
    If Not IsNumeric(valueD) Then Exit Function

    numericValue = CDbl(valueD)
    If numericValue <> Fix(numericValue) Then Exit Function

    ToTwoDigitString = Format$(CLng(numericValue), "00")
End Function

Private Sub ValidateWindowsFileName(ByVal fileNameText As String)
    Dim invalidChars As Variant
    Dim i As Long

    If Len(fileNameText) = 0 Then
        Err.Raise vbObjectError + 1101, "ValidateWindowsFileName", "ファイル名が空です。"
    End If

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(1, fileNameText, CStr(invalidChars(i)), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 1102, "ValidateWindowsFileName", _
                      "ファイル名に使用できない文字が含まれています: " & CStr(invalidChars(i))
        End If
    Next i

    If Right$(fileNameText, 1) = "." Or Right$(fileNameText, 1) = " " Then
        Err.Raise vbObjectError + 1103, "ValidateWindowsFileName", _
                  "ファイル名の末尾にピリオドまたは空白は使用できません。"
    End If
End Sub

Private Function BuildPath(ByVal folderPath As String, ByVal fileNameText As String) As String
    If Right$(folderPath, 1) = "\" Then
        BuildPath = folderPath & fileNameText
    Else
        BuildPath = folderPath & "\" & fileNameText
    End If
End Function

Private Sub CreateEmptyXlsxFile(ByVal fullPath As String)
    Dim newWb As Workbook
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean

    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo CleanFail

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    If Len(Dir$(fullPath)) > 0 Then
        Kill fullPath
    End If

    Set newWb = Application.Workbooks.Add(xlWBATWorksheet)
    newWb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    Set newWb = Nothing

    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not newWb Is Nothing Then
        newWb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
