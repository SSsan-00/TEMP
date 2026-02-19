Attribute VB_Name = "modTests"
Option Explicit

' =============================================================================
' modTests
'   即時ウィンドウから実行する簡易テスト。
' =============================================================================

Public Sub RunAllTests()
  Test_AppendLineAlignment
  Test_InsertDeleteAlignment
  Debug.Print "All tests passed."
End Sub

Public Sub Test_AppendLineAlignment()
  Dim leftLines(0 To 2) As String
  Dim rightLines(0 To 2) As String
  Dim alignedLeft() As String
  Dim alignedRight() As String
  Dim ops() As AlignOp
  Dim i As Long
  Dim foundHeadPair As Boolean

  leftLines(0) = "<head>"
  leftLines(1) = "  <meta charset=utf-8 />"
  leftLines(2) = "</head>"

  rightLines(0) = "sb.AppendLine(""<head>"");"
  rightLines(1) = "sb.AppendLine(""  <meta charset=utf-8 />"");"
  rightLines(2) = "sb.AppendLine(""</head>"");"

  AlignColumnsByDiffWithOps leftLines, rightLines, alignedLeft, alignedRight, ops

  For i = 0 To GetStringArrayCount(alignedLeft) - 1
    Debug.Print CStr(i + 1) & " | L: " & alignedLeft(i) & " | R: " & alignedRight(i)
    If InStr(1, alignedLeft(i), "<head>", vbTextCompare) > 0 And InStr(1, alignedRight(i), "AppendLine(""<head>"")", vbTextCompare) > 0 Then
      foundHeadPair = True
    End If
  Next i

  AssertTrue foundHeadPair, "AppendLine(""<head>"") と <head> が同じ行に揃っていません。"
End Sub

Public Sub Test_InsertDeleteAlignment()
  Dim leftLines(0 To 2) As String
  Dim rightLines(0 To 3) As String
  Dim alignedLeft() As String
  Dim alignedRight() As String
  Dim i As Long
  Dim hasRightOnlyLine As Boolean

  leftLines(0) = "line-a"
  leftLines(1) = "line-b"
  leftLines(2) = "line-c"

  rightLines(0) = "line-a"
  rightLines(1) = "line-b"
  rightLines(2) = "line-new"
  rightLines(3) = "line-c"

  AlignColumnsByDiff leftLines, rightLines, alignedLeft, alignedRight

  For i = 0 To GetStringArrayCount(alignedLeft) - 1
    If alignedLeft(i) = "" And alignedRight(i) = "line-new" Then
      hasRightOnlyLine = True
      Exit For
    End If
  Next i

  AssertTrue hasRightOnlyLine, "insert 行に対して左側空セルが挿入されていません。"
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
  If Not condition Then
    Err.Raise vbObjectError + 3000, "modTests", message
  End If
End Sub
