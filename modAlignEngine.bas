Attribute VB_Name = "modAlignEngine"
Option Explicit

' =============================================================================
' modAlignEngine
'   diff-viewer の行差分ロジック(diffLines + pairReplace)を VBA 向けに移植。
'   目的は C列(PHP) / D列(C#) を「横並び」へ整形すること。
' =============================================================================

Public Enum DiffOpType
  DiffOpEqual = 0
  DiffOpInsert = 1
  DiffOpDelete = 2
  DiffOpReplace = 3
End Enum

Public Type AlignOp
  OpType As DiffOpType
  LeftLine As String
  RightLine As String
  LeftLineNo As Long
  RightLineNo As Long
  HasLeftLineNo As Boolean
  HasRightLineNo As Boolean
End Type

Private Type LineOpVector
  Items() As AlignOp
  Count As Long
  Capacity As Long
End Type

Private Type UniquePair
  LeftIndex As Long
  RightIndex As Long
End Type

Private Type UniquePairVector
  Items() As UniquePair
  Count As Long
  Capacity As Long
End Type

Private Type PairCandidate
  DeleteIndex As Long
  InsertIndex As Long
  IndentDiff As Long
  Score As Double
  Distance As Long
End Type

Private Type PairCandidateVector
  Items() As PairCandidate
  Count As Long
  Capacity As Long
End Type

Private Const WINDOW_SIZE As Long = 40
Private Const SCORE_THRESHOLD As Double = 4#

Public Sub AlignColumnsByDiff(ByRef leftOriginal() As String, ByRef rightOriginal() As String, ByRef alignedLeft() As String, ByRef alignedRight() As String)
  Dim ops() As AlignOp
  AlignColumnsByDiffWithOps leftOriginal, rightOriginal, alignedLeft, alignedRight, ops
End Sub

Public Sub AlignColumnsByDiffWithOps(ByRef leftOriginal() As String, ByRef rightOriginal() As String, ByRef alignedLeft() As String, ByRef alignedRight() As String, ByRef opsOut() As AlignOp)
  Dim lineOps As LineOpVector
  Dim pairedOps As LineOpVector

  lineOps = DiffLinesFromLines(leftOriginal, rightOriginal)
  pairedOps = PairReplace(lineOps)

  ConvertOpsToAlignedArrays pairedOps, alignedLeft, alignedRight
  ExportOps pairedOps, opsOut
End Sub

Public Function GetStringArrayCount(ByRef values() As String) As Long
  On Error GoTo EmptyArray
  GetStringArrayCount = UBound(values) - LBound(values) + 1
  Exit Function
EmptyArray:
  GetStringArrayCount = 0
End Function

' -----------------------------------------------------------------------------
' diffLines.ts 移植本体
' -----------------------------------------------------------------------------

Private Function DiffLinesFromLines(ByRef leftLines() As String, ByRef rightLines() As String) As LineOpVector
  Dim leftCompare() As String
  Dim rightCompare() As String

  leftCompare = BuildCompareLines(leftLines)
  rightCompare = BuildCompareLines(rightLines)

  DiffLinesFromLines = DiffLinesPatience(leftLines, rightLines, leftCompare, rightCompare, 0, 0)
End Function

Private Function BuildCompareLines(ByRef lines() As String) As String()
  Dim count As Long
  Dim i As Long
  Dim result() As String

  count = GetStringArrayCount(lines)
  If count = 0 Then
    Erase result
    BuildCompareLines = result
    Exit Function
  End If

  ReDim result(0 To count - 1)
  For i = 0 To count - 1
    result(i) = NormalizeForMatch(lines(i))
  Next i

  BuildCompareLines = result
End Function

Private Function DiffLinesPatience(ByRef leftLines() As String, ByRef rightLines() As String, ByRef leftCompare() As String, ByRef rightCompare() As String, ByVal leftOffset As Long, ByVal rightOffset As Long) As LineOpVector
  Dim leftCount As Long
  Dim rightCount As Long
  Dim anchors As UniquePairVector
  Dim result As LineOpVector
  Dim leftStart As Long
  Dim rightStart As Long
  Dim i As Long
  Dim anchor As UniquePair
  Dim leftSegment() As String
  Dim rightSegment() As String
  Dim leftCompareSegment() As String
  Dim rightCompareSegment() As String
  Dim segOps As LineOpVector
  Dim leftLine As String
  Dim rightLine As String
  Dim leftKey As String
  Dim rightKey As String
  Dim oneLeft(0 To 0) As String
  Dim oneRight(0 To 0) As String
  Dim oneLeftKey(0 To 0) As String
  Dim oneRightKey(0 To 0) As String
  Dim tailLeft() As String
  Dim tailRight() As String
  Dim tailLeftCompare() As String
  Dim tailRightCompare() As String

  leftCount = GetStringArrayCount(leftLines)
  rightCount = GetStringArrayCount(rightLines)

  If leftCount = 0 And rightCount = 0 Then
    DiffLinesPatience = result
    Exit Function
  End If

  anchors = LongestIncreasingPairs(BuildUniquePairs(leftLines, rightLines))
  If anchors.Count = 0 Then
    DiffLinesPatience = DiffLinesMyers(leftLines, rightLines, leftCompare, rightCompare, leftOffset, rightOffset)
    Exit Function
  End If

  leftStart = 0
  rightStart = 0

  For i = 0 To anchors.Count - 1
    anchor = anchors.Items(i)

    leftSegment = SliceStringArray(leftLines, leftStart, anchor.LeftIndex)
    rightSegment = SliceStringArray(rightLines, rightStart, anchor.RightIndex)
    leftCompareSegment = SliceStringArray(leftCompare, leftStart, anchor.LeftIndex)
    rightCompareSegment = SliceStringArray(rightCompare, rightStart, anchor.RightIndex)

    segOps = DiffLinesPatience(leftSegment, rightSegment, leftCompareSegment, rightCompareSegment, leftOffset + leftStart, rightOffset + rightStart)
    AppendOpVector result, segOps

    leftLine = NzString(leftLines, anchor.LeftIndex)
    rightLine = NzString(rightLines, anchor.RightIndex)
    leftKey = NzString(leftCompare, anchor.LeftIndex)
    rightKey = NzString(rightCompare, anchor.RightIndex)

    If leftLine = rightLine Then
      PushOp result, MakeEqualOp(leftLine, rightLine, leftOffset + anchor.LeftIndex, rightOffset + anchor.RightIndex)
    ElseIf leftKey = rightKey And IsBlankLine(leftLine) And IsBlankLine(rightLine) Then
      PushOp result, MakeEqualOp(leftLine, rightLine, leftOffset + anchor.LeftIndex, rightOffset + anchor.RightIndex)
    ElseIf leftKey = rightKey Then
      PushOp result, MakeDeleteOp(leftLine, leftOffset + anchor.LeftIndex)
      PushOp result, MakeInsertOp(rightLine, rightOffset + anchor.RightIndex)
    Else
      oneLeft(0) = leftLine
      oneRight(0) = rightLine
      oneLeftKey(0) = leftKey
      oneRightKey(0) = rightKey
      segOps = DiffLinesMyers(oneLeft, oneRight, oneLeftKey, oneRightKey, leftOffset + anchor.LeftIndex, rightOffset + anchor.RightIndex)
      AppendOpVector result, segOps
    End If

    leftStart = anchor.LeftIndex + 1
    rightStart = anchor.RightIndex + 1
  Next i

  tailLeft = SliceStringArray(leftLines, leftStart, leftCount)
  tailRight = SliceStringArray(rightLines, rightStart, rightCount)
  tailLeftCompare = SliceStringArray(leftCompare, leftStart, leftCount)
  tailRightCompare = SliceStringArray(rightCompare, rightStart, rightCount)

  segOps = DiffLinesPatience(tailLeft, tailRight, tailLeftCompare, tailRightCompare, leftOffset + leftStart, rightOffset + rightStart)
  AppendOpVector result, segOps

  DiffLinesPatience = result
End Function

Private Function DiffLinesMyers(ByRef leftLines() As String, ByRef rightLines() As String, ByRef leftCompare() As String, ByRef rightCompare() As String, ByVal leftOffset As Long, ByVal rightOffset As Long) As LineOpVector
  Dim leftCount As Long
  Dim rightCount As Long
  Dim trace As Variant
  Dim lastD As Long
  Dim offset As Long
  Dim ops As LineOpVector

  leftCount = GetStringArrayCount(leftLines)
  rightCount = GetStringArrayCount(rightLines)

  If leftCount = 0 And rightCount = 0 Then
    DiffLinesMyers = ops
    Exit Function
  End If

  trace = BuildMyersTrace(leftCompare, rightCompare, lastD, offset)
  ops = BacktrackOps(leftLines, rightLines, leftCompare, rightCompare, trace, lastD, offset)
  OffsetOps ops, leftOffset, rightOffset

  DiffLinesMyers = ops
End Function

Private Function BuildMyersTrace(ByRef leftCompare() As String, ByRef rightCompare() As String, ByRef lastD As Long, ByRef offset As Long) As Variant
  Dim n As Long
  Dim m As Long
  Dim maxD As Long
  Dim v() As Long
  Dim trace() As Long
  Dim d As Long
  Dim k As Long
  Dim kIndex As Long
  Dim x As Long
  Dim y As Long
  Dim idx As Long
  Dim found As Boolean
  Dim moveDown As Boolean

  n = GetStringArrayCount(leftCompare)
  m = GetStringArrayCount(rightCompare)
  maxD = n + m
  offset = maxD

  If maxD = 0 Then
    ReDim trace(0 To 0, 0 To 0)
    trace(0, 0) = 0
    lastD = 0
    BuildMyersTrace = trace
    Exit Function
  End If

  ReDim v(0 To 2 * maxD)
  ReDim trace(0 To maxD, 0 To 2 * maxD)

  found = False
  For d = 0 To maxD
    For k = -d To d Step 2
      kIndex = k + offset

      If k = -d Then
        moveDown = True
      ElseIf k = d Then
        moveDown = False
      Else
        moveDown = (v(kIndex - 1) < v(kIndex + 1))
      End If

      If moveDown Then
        x = v(kIndex + 1)
      Else
        x = v(kIndex - 1) + 1
      End If

      y = x - k
      Do While x < n And y < m And leftCompare(x) = rightCompare(y)
        x = x + 1
        y = y + 1
      Loop

      v(kIndex) = x
      If x >= n And y >= m Then
        found = True
        Exit For
      End If
    Next k

    For idx = 0 To 2 * maxD
      trace(d, idx) = v(idx)
    Next idx

    If found Then
      Exit For
    End If
  Next d

  lastD = d
  BuildMyersTrace = trace
End Function

Private Function BacktrackOps(ByRef leftLines() As String, ByRef rightLines() As String, ByRef leftCompare() As String, ByRef rightCompare() As String, ByRef trace As Variant, ByVal lastD As Long, ByVal offset As Long) As LineOpVector
  Dim n As Long
  Dim m As Long
  Dim ops As LineOpVector
  Dim x As Long
  Dim y As Long
  Dim d As Long
  Dim k As Long
  Dim kIndex As Long
  Dim prevK As Long
  Dim prevX As Long
  Dim prevY As Long
  Dim leftLine As String
  Dim rightLine As String
  Dim leftKey As String
  Dim rightKey As String
  Dim moveDown As Boolean

  n = GetStringArrayCount(leftLines)
  m = GetStringArrayCount(rightLines)
  x = n
  y = m

  For d = lastD To 0 Step -1
    k = x - y
    kIndex = k + offset

    If k = -d Then
      moveDown = True
    ElseIf k = d Then
      moveDown = False
    Else
      moveDown = (trace(d, kIndex - 1) < trace(d, kIndex + 1))
    End If

    If moveDown Then
      prevK = k + 1
    Else
      prevK = k - 1
    End If

    prevX = trace(d, prevK + offset)
    prevY = prevX - prevK

    Do While x > prevX And y > prevY
      leftLine = leftLines(x - 1)
      rightLine = rightLines(y - 1)
      leftKey = leftCompare(x - 1)
      rightKey = rightCompare(y - 1)

      If leftLine = rightLine Then
        PushOp ops, MakeEqualOp(leftLine, rightLine, x - 1, y - 1)
      ElseIf leftKey = rightKey And IsBlankLine(leftLine) And IsBlankLine(rightLine) Then
        PushOp ops, MakeEqualOp(leftLine, rightLine, x - 1, y - 1)
      Else
        PushOp ops, MakeInsertOp(rightLine, y - 1)
        PushOp ops, MakeDeleteOp(leftLine, x - 1)
      End If

      x = x - 1
      y = y - 1
    Loop

    If d = 0 Then
      Exit For
    End If

    If x = prevX Then
      PushOp ops, MakeInsertOp(rightLines(y - 1), y - 1)
      y = y - 1
    Else
      PushOp ops, MakeDeleteOp(leftLines(x - 1), x - 1)
      x = x - 1
    End If
  Next d

  BacktrackOps = ReverseOps(ops)
End Function

Private Sub OffsetOps(ByRef ops As LineOpVector, ByVal leftOffset As Long, ByVal rightOffset As Long)
  Dim i As Long

  For i = 0 To ops.Count - 1
    If ops.Items(i).HasLeftLineNo Then
      ops.Items(i).LeftLineNo = ops.Items(i).LeftLineNo + leftOffset
    End If
    If ops.Items(i).HasRightLineNo Then
      ops.Items(i).RightLineNo = ops.Items(i).RightLineNo + rightOffset
    End If
  Next i
End Sub

Private Function BuildUniquePairs(ByRef leftLines() As String, ByRef rightLines() As String) As UniquePairVector
  Dim leftMap As Object
  Dim rightMap As Object
  Dim key As Variant
  Dim leftEntry As Variant
  Dim rightEntry As Variant
  Dim result As UniquePairVector
  Dim p As UniquePair

  Set leftMap = BuildKeyMap(leftLines)
  Set rightMap = BuildKeyMap(rightLines)

  For Each key In leftMap.Keys
    leftEntry = leftMap(key)
    If CLng(leftEntry(1)) <> 1 Then
      GoTo ContinueKey
    End If

    If Not rightMap.Exists(key) Then
      GoTo ContinueKey
    End If

    rightEntry = rightMap(key)
    If CLng(rightEntry(1)) <> 1 Then
      GoTo ContinueKey
    End If

    p.LeftIndex = CLng(leftEntry(0))
    p.RightIndex = CLng(rightEntry(0))
    PushPair result, p
ContinueKey:
  Next key

  If result.Count > 1 Then
    SortPairsByLeft result
  End If

  BuildUniquePairs = result
End Function

Private Function BuildKeyMap(ByRef lines() As String) As Object
  Dim dict As Object
  Dim i As Long
  Dim count As Long
  Dim compareLine As String
  Dim rawKey As String
  Dim key As String
  Dim entry As Variant

  Set dict = CreateObject("Scripting.Dictionary")

  count = GetStringArrayCount(lines)
  For i = 0 To count - 1
    compareLine = ToAppendLiteralOrLine(lines(i))
    rawKey = ExtractLineKey(compareLine)
    If Len(rawKey) > 0 Then
      key = rawKey
    Else
      key = LTrimWhitespace(compareLine)
    End If

    If dict.Exists(key) Then
      entry = dict(key)
      entry(1) = CLng(entry(1)) + 1
      dict(key) = entry
    Else
      dict.Add key, Array(i, 1)
    End If
  Next i

  Set BuildKeyMap = dict
End Function

Private Function LongestIncreasingPairs(ByRef pairs As UniquePairVector) As UniquePairVector
  Dim result As UniquePairVector
  Dim tailValues() As Long
  Dim tailIndices() As Long
  Dim prevIndices() As Long
  Dim i As Long
  Dim pos As Long
  Dim tailLen As Long
  Dim k As Long
  Dim seq As UniquePairVector

  If pairs.Count = 0 Then
    LongestIncreasingPairs = result
    Exit Function
  End If

  ReDim tailValues(0 To pairs.Count - 1)
  ReDim tailIndices(0 To pairs.Count - 1)
  ReDim prevIndices(0 To pairs.Count - 1)

  For i = 0 To pairs.Count - 1
    prevIndices(i) = -1
  Next i

  tailLen = 0
  For i = 0 To pairs.Count - 1
    pos = LowerBoundLong(tailValues, tailLen, pairs.Items(i).RightIndex)

    tailValues(pos) = pairs.Items(i).RightIndex
    tailIndices(pos) = i

    If pos > 0 Then
      prevIndices(i) = tailIndices(pos - 1)
    End If

    If pos = tailLen Then
      tailLen = tailLen + 1
    End If
  Next i

  k = tailIndices(tailLen - 1)
  Do While k >= 0
    PushPair seq, pairs.Items(k)
    k = prevIndices(k)
  Loop

  result = ReversePairs(seq)
  LongestIncreasingPairs = result
End Function

Private Function LowerBoundLong(ByRef values() As Long, ByVal length As Long, ByVal value As Long) As Long
  Dim low As Long
  Dim high As Long
  Dim mid As Long

  low = 0
  high = length

  Do While low < high
    mid = (low + high) \ 2
    If values(mid) < value Then
      low = mid + 1
    Else
      high = mid
    End If
  Loop

  LowerBoundLong = low
End Function

' -----------------------------------------------------------------------------
' pairReplace.ts 移植（Replace へ束ねる段）
' -----------------------------------------------------------------------------

Private Function PairReplace(ByRef ops As LineOpVector) As LineOpVector
  Dim result As LineOpVector
  Dim i As Long
  Dim op As AlignOp
  Dim deletes As LineOpVector
  Dim inserts As LineOpVector
  Dim j As Long
  Dim pairedBlock As LineOpVector

  i = 0
  Do While i < ops.Count
    op = ops.Items(i)

    If op.OpType <> DiffOpDelete Then
      PushOp result, ToPairedOp(op)
      i = i + 1
      GoTo ContinueLoop
    End If

    ClearOpVector deletes
    Do While i < ops.Count And ops.Items(i).OpType = DiffOpDelete
      PushOp deletes, ops.Items(i)
      i = i + 1
    Loop

    ClearOpVector inserts
    j = i
    Do While j < ops.Count And ops.Items(j).OpType = DiffOpInsert
      PushOp inserts, ops.Items(j)
      j = j + 1
    Loop

    If inserts.Count = 0 Then
      AppendOpVector result, deletes
      GoTo ContinueLoop
    End If

    pairedBlock = PairBlock(deletes, inserts)
    AppendOpVector result, pairedBlock
    i = j
ContinueLoop:
  Loop

  PairReplace = AlignBracePairs(result)
End Function

Private Function PairBlock(ByRef deletes As LineOpVector, ByRef inserts As LineOpVector) As LineOpVector
  Dim result As LineOpVector
  Dim candidates As PairCandidateVector
  Dim matches() As Long
  Dim usedInserts() As Boolean
  Dim emittedInserts() As Boolean
  Dim deleteCount As Long
  Dim insertCount As Long
  Dim i As Long
  Dim insertIndex As Long
  Dim insertCursor As Long
  Dim stopRightLineNo As Long
  Dim insertLineNo As Long
  Dim leftOp As AlignOp
  Dim rightOp As AlignOp

  deleteCount = deletes.Count
  insertCount = inserts.Count

  If deleteCount = 0 And insertCount = 0 Then
    PairBlock = result
    Exit Function
  End If

  If deleteCount > 0 Then
    ReDim matches(0 To deleteCount - 1)
  Else
    ReDim matches(0 To 0)
  End If
  If insertCount > 0 Then
    ReDim usedInserts(0 To insertCount - 1)
    ReDim emittedInserts(0 To insertCount - 1)
  Else
    ReDim usedInserts(0 To 0)
    ReDim emittedInserts(0 To 0)
  End If

  For i = 0 To deleteCount - 1
    matches(i) = -1
  Next i

  candidates = BuildCandidates(deletes, inserts)
  SortCandidates candidates

  For i = 0 To candidates.Count - 1
    If usedInserts(candidates.Items(i).InsertIndex) Then
      GoTo ContinueCandidate
    End If
    If matches(candidates.Items(i).DeleteIndex) <> -1 Then
      GoTo ContinueCandidate
    End If

    matches(candidates.Items(i).DeleteIndex) = candidates.Items(i).InsertIndex
    usedInserts(candidates.Items(i).InsertIndex) = True
ContinueCandidate:
  Next i

  insertCursor = 0

  For i = 0 To deleteCount - 1
    insertIndex = matches(i)

    If insertIndex <> -1 Then
      rightOp = inserts.Items(insertIndex)
      If rightOp.HasRightLineNo Then
        stopRightLineNo = rightOp.RightLineNo
      Else
        stopRightLineNo = 2147483647
      End If

      Do While insertCursor < insertCount
        If inserts.Items(insertCursor).HasRightLineNo Then
          insertLineNo = inserts.Items(insertCursor).RightLineNo
        Else
          insertLineNo = 2147483647
        End If
        If insertLineNo >= stopRightLineNo Then
          Exit Do
        End If
        If (Not usedInserts(insertCursor)) And (Not emittedInserts(insertCursor)) Then
          PushOp result, ToPairedOp(inserts.Items(insertCursor))
          emittedInserts(insertCursor) = True
        End If
        insertCursor = insertCursor + 1
      Loop

      leftOp = deletes.Items(i)
      rightOp = inserts.Items(insertIndex)
      PushOp result, MakeReplaceOp(leftOp.LeftLine, rightOp.RightLine, leftOp.LeftLineNo, rightOp.RightLineNo, leftOp.HasLeftLineNo, rightOp.HasRightLineNo)

      Do While insertCursor <= insertIndex And insertCursor < insertCount
        insertCursor = insertCursor + 1
      Loop
    Else
      PushOp result, ToPairedOp(deletes.Items(i))
    End If
  Next i

  Do While insertCursor < insertCount
    If (Not usedInserts(insertCursor)) And (Not emittedInserts(insertCursor)) Then
      PushOp result, ToPairedOp(inserts.Items(insertCursor))
      emittedInserts(insertCursor) = True
    End If
    insertCursor = insertCursor + 1
  Loop

  For i = 0 To insertCount - 1
    If (Not usedInserts(i)) And (Not emittedInserts(i)) Then
      PushOp result, ToPairedOp(inserts.Items(i))
      emittedInserts(i) = True
    End If
  Next i

  PairBlock = result
End Function

Private Function BuildCandidates(ByRef deletes As LineOpVector, ByRef inserts As LineOpVector) As PairCandidateVector
  Dim result As PairCandidateVector
  Dim deleteCount As Long
  Dim insertCount As Long
  Dim deleteCore() As String
  Dim insertCore() As String
  Dim deleteLineKey() As String
  Dim insertLineKey() As String
  Dim deleteTrimmed() As String
  Dim insertTrimmed() As String
  Dim deleteIndent() As Long
  Dim insertIndent() As Long
  Dim deleteTokens() As Object
  Dim insertTokens() As Object
  Dim insertIndexMap As Object
  Dim d As Long
  Dim i As Long
  Dim c As PairCandidate
  Dim score As Double
  Dim candidateSet As Object
  Dim idx As Variant

  deleteCount = deletes.Count
  insertCount = inserts.Count

  If deleteCount = 0 Or insertCount = 0 Then
    BuildCandidates = result
    Exit Function
  End If

  ReDim deleteCore(0 To deleteCount - 1)
  ReDim insertCore(0 To insertCount - 1)
  ReDim deleteLineKey(0 To deleteCount - 1)
  ReDim insertLineKey(0 To insertCount - 1)
  ReDim deleteTrimmed(0 To deleteCount - 1)
  ReDim insertTrimmed(0 To insertCount - 1)
  ReDim deleteIndent(0 To deleteCount - 1)
  ReDim insertIndent(0 To insertCount - 1)
  ReDim deleteTokens(0 To deleteCount - 1)
  ReDim insertTokens(0 To insertCount - 1)

  For d = 0 To deleteCount - 1
    deleteCore(d) = BuildSimpleSimilarityKey(deletes.Items(d).LeftLine)
    deleteLineKey(d) = ExtractLineKey(ToAppendLiteralOrLine(deletes.Items(d).LeftLine))
    deleteTrimmed(d) = LTrimWhitespace(deletes.Items(d).LeftLine)
    deleteIndent(d) = CountIndent(deletes.Items(d).LeftLine)
    Set deleteTokens(d) = BuildIndexTokenSet(deletes.Items(d).LeftLine, deleteCore(d), deleteLineKey(d))
  Next d

  Set insertIndexMap = CreateObject("Scripting.Dictionary")
  insertIndexMap.CompareMode = 1 ' TextCompare

  For i = 0 To insertCount - 1
    insertCore(i) = BuildSimpleSimilarityKey(inserts.Items(i).RightLine)
    insertLineKey(i) = ExtractLineKey(ToAppendLiteralOrLine(inserts.Items(i).RightLine))
    insertTrimmed(i) = LTrimWhitespace(inserts.Items(i).RightLine)
    insertIndent(i) = CountIndent(inserts.Items(i).RightLine)
    Set insertTokens(i) = BuildIndexTokenSet(inserts.Items(i).RightLine, insertCore(i), insertLineKey(i))
    AddTokensToIndexMap insertTokens(i), i, insertIndexMap
  Next i

  For d = 0 To deleteCount - 1
    Set candidateSet = BuildCandidateIndices(d, insertCount, deleteTokens(d), insertIndexMap)

    For Each idx In candidateSet.Keys
      i = CLng(idx)

      If deleteTrimmed(d) = insertTrimmed(i) And Len(deleteTrimmed(d)) > 0 And deletes.Items(d).LeftLine <> inserts.Items(i).RightLine Then
        c.DeleteIndex = d
        c.InsertIndex = i
        c.IndentDiff = Abs(deleteIndent(d) - insertIndent(i))
        c.Score = SCORE_THRESHOLD + 5#
        c.Distance = Abs(d - i)
        PushCandidate result, c
        GoTo ContinueInsert
      End If

      score = ScoreLinePairSimple(deleteCore(d), insertCore(i), deleteLineKey(d), insertLineKey(i), deleteTokens(d), insertTokens(i))
      If score < SCORE_THRESHOLD Then
        GoTo ContinueInsert
      End If

      c.DeleteIndex = d
      c.InsertIndex = i
      c.IndentDiff = Abs(deleteIndent(d) - insertIndent(i))
      c.Score = score
      c.Distance = Abs(d - i)
      PushCandidate result, c
ContinueInsert:
    Next idx
  Next d

  BuildCandidates = result
End Function

Private Function BuildCandidateIndices(ByVal deleteIndex As Long, ByVal insertCount As Long, ByVal tokenSet As Object, ByVal indexMap As Object) As Object
  Dim indices As Object
  Dim startIndex As Long
  Dim endIndex As Long
  Dim i As Long
  Dim token As Variant
  Dim bucket As Collection
  Dim entry As Variant

  Set indices = CreateObject("Scripting.Dictionary")

  startIndex = deleteIndex - WINDOW_SIZE
  If startIndex < 0 Then startIndex = 0
  endIndex = deleteIndex + WINDOW_SIZE
  If endIndex > insertCount - 1 Then endIndex = insertCount - 1

  For i = startIndex To endIndex
    indices(CStr(i)) = True
  Next i

  For Each token In tokenSet.Keys
    If indexMap.Exists(token) Then
      Set bucket = indexMap(token)
      For Each entry In bucket
        indices(CStr(CLng(entry))) = True
      Next entry
    End If
  Next token

  Set BuildCandidateIndices = indices
End Function

Private Function ScoreLinePairSimple(ByVal leftCore As String, ByVal rightCore As String, ByVal leftLineKey As String, ByVal rightLineKey As String, ByVal leftTokens As Object, ByVal rightTokens As Object) As Double
  Dim score As Double
  Dim overlap As Long

  overlap = TokenSetOverlap(leftTokens, rightTokens)

  If Len(leftCore) > 0 And Len(rightCore) > 0 And leftCore = rightCore Then
    score = score + 8#
  End If

  If Len(leftLineKey) > 0 And Len(rightLineKey) > 0 Then
    If leftLineKey = rightLineKey Then
      score = score + 6#
    ElseIf overlap = 0 And leftCore <> rightCore Then
      ' diff-viewer lineSimilarity の「primaryId 不一致で根拠なしなら棄却」に寄せる。
      ScoreLinePairSimple = -1#
      Exit Function
    End If
  End If

  score = score + CDbl(overlap)

  If score < SCORE_THRESHOLD Then
    If Len(leftCore) > 0 And Len(rightCore) > 0 Then
      If InStr(1, leftCore, rightCore, vbTextCompare) > 0 Or InStr(1, rightCore, leftCore, vbTextCompare) > 0 Then
        score = score + 4#
      End If
    End If
  End If

  ScoreLinePairSimple = score
End Function

Private Function BuildIndexTokenSet(ByVal line As String, ByVal coreKey As String, ByVal lineKey As String) As Object
  Dim setDict As Object
  Dim idTokens As Collection
  Dim litTokens As Collection
  Dim numTokens As Collection
  Dim item As Variant

  Set setDict = CreateObject("Scripting.Dictionary")
  setDict.CompareMode = 1 ' TextCompare

  If Len(coreKey) > 0 Then setDict("core:" & coreKey) = True
  If Len(lineKey) > 0 Then setDict("line:" & lineKey) = True

  Set idTokens = ExtractIdentifierTokens(ToAppendLiteralOrLine(line))
  For Each item In idTokens
    setDict("id:" & CStr(item)) = True
  Next item

  Set litTokens = ExtractLiteralTokens(ToAppendLiteralOrLine(line))
  For Each item In litTokens
    setDict("lit:" & CStr(item)) = True
  Next item

  Set numTokens = ExtractNumberTokens(ToAppendLiteralOrLine(line))
  For Each item In numTokens
    setDict("num:" & CStr(item)) = True
  Next item

  Set BuildIndexTokenSet = setDict
End Function

Private Sub AddTokensToIndexMap(ByVal tokenSet As Object, ByVal index As Long, ByVal indexMap As Object)
  Dim token As Variant
  Dim bucket As Collection

  For Each token In tokenSet.Keys
    If indexMap.Exists(token) Then
      Set bucket = indexMap(token)
    Else
      Set bucket = New Collection
      indexMap.Add token, bucket
    End If
    bucket.Add index
  Next token
End Sub

Private Function TokenSetOverlap(ByVal leftSet As Object, ByVal rightSet As Object) As Long
  Dim key As Variant
  Dim count As Long

  count = 0
  For Each key In leftSet.Keys
    If rightSet.Exists(key) Then
      count = count + 1
    End If
  Next key

  TokenSetOverlap = count
End Function

Private Function ExtractLiteralTokens(ByVal line As String) As Collection
  Dim matches As Object
  Dim re As Object
  Dim result As New Collection
  Dim i As Long
  Dim value As String
  Dim normalized As String

  Set re = CreateObject("VBScript.RegExp")
  re.Global = True
  re.IgnoreCase = True
  re.Pattern = "'([^'\\]|\\.)*'|""([^""\\]|\\.)*"""

  Set matches = re.Execute(line)
  For i = 0 To matches.Count - 1
    value = matches(i).Value
    If Len(value) >= 2 Then
      value = Mid$(value, 2, Len(value) - 2)
    End If
    value = UnescapeLiteral(value)
    normalized = LCase$(NormalizeComparisonWhitespace(value))
    If Len(normalized) > 0 Then
      On Error Resume Next
      result.Add normalized, normalized
      On Error GoTo 0
    End If
  Next i

  Set ExtractLiteralTokens = result
End Function

Private Function UnescapeLiteral(ByVal value As String) As String
  Dim s As String
  s = value
  s = Replace(s, "\\", "\")
  s = Replace(s, "\" & Chr$(34), Chr$(34))
  s = Replace(s, "\'", "'")
  s = Replace(s, "¥", "\")
  s = Replace(s, "￥", "\")
  UnescapeLiteral = s
End Function

Private Function ExtractNumberTokens(ByVal line As String) As Collection
  Dim re As Object
  Dim matches As Object
  Dim result As New Collection
  Dim i As Long
  Dim token As String

  Set re = CreateObject("VBScript.RegExp")
  re.Global = True
  re.IgnoreCase = True
  re.Pattern = "\b\d+(\.\d+)?\b"

  Set matches = re.Execute(line)
  For i = 0 To matches.Count - 1
    token = CStr(matches(i).Value)
    On Error Resume Next
    result.Add token, token
    On Error GoTo 0
  Next i

  Set ExtractNumberTokens = result
End Function

Private Function AlignBracePairs(ByRef ops As LineOpVector) As LineOpVector
  Dim result As LineOpVector
  Dim i As Long
  Dim current As AlignOp
  Dim nextOp As AlignOp
  Dim prev As AlignOp
  Dim hasPrev As Boolean
  Dim prevPaired As Boolean

  For i = 0 To ops.Count - 1
    current = ops.Items(i)

    If result.Count > 0 Then
      prev = result.Items(result.Count - 1)
      hasPrev = True
      prevPaired = (prev.OpType = DiffOpEqual Or prev.OpType = DiffOpReplace)
    Else
      hasPrev = False
      prevPaired = False
    End If

    If i < ops.Count - 1 Then
      nextOp = ops.Items(i + 1)
    End If

    If hasPrev And prevPaired Then
      If current.OpType = DiffOpDelete And i < ops.Count - 1 Then
        If nextOp.OpType = DiffOpInsert And IsBraceLine(current.LeftLine) And IsBraceLine(nextOp.RightLine) And current.HasLeftLineNo And nextOp.HasRightLineNo Then
          PushOp result, MakeReplaceOp(current.LeftLine, nextOp.RightLine, current.LeftLineNo, nextOp.RightLineNo, True, True)
          i = i + 1
          GoTo ContinueLoop
        End If
      End If

      If current.OpType = DiffOpInsert And i < ops.Count - 1 Then
        If nextOp.OpType = DiffOpDelete And IsBraceLine(nextOp.LeftLine) And IsBraceLine(current.RightLine) And nextOp.HasLeftLineNo And current.HasRightLineNo Then
          PushOp result, MakeReplaceOp(nextOp.LeftLine, current.RightLine, nextOp.LeftLineNo, current.RightLineNo, True, True)
          i = i + 1
          GoTo ContinueLoop
        End If
      End If
    End If

    PushOp result, current
ContinueLoop:
  Next i

  AlignBracePairs = result
End Function

Private Function IsBraceLine(ByVal line As String) As Boolean
  Dim trimmed As String
  trimmed = Trim$(line)
  IsBraceLine = RegexTest(trimmed, "^}\s*;?\s*$", True)
End Function

Private Function ToPairedOp(ByRef op As AlignOp) As AlignOp
  ToPairedOp = op
End Function

' -----------------------------------------------------------------------------
' 出力整形
' -----------------------------------------------------------------------------

Private Sub ConvertOpsToAlignedArrays(ByRef ops As LineOpVector, ByRef alignedLeft() As String, ByRef alignedRight() As String)
  Dim i As Long
  Dim count As Long

  count = ops.Count
  If count = 0 Then
    Erase alignedLeft
    Erase alignedRight
    Exit Sub
  End If

  ReDim alignedLeft(0 To count - 1)
  ReDim alignedRight(0 To count - 1)

  For i = 0 To count - 1
    Select Case ops.Items(i).OpType
      Case DiffOpEqual, DiffOpReplace
        alignedLeft(i) = ops.Items(i).LeftLine
        alignedRight(i) = ops.Items(i).RightLine
      Case DiffOpDelete
        alignedLeft(i) = ops.Items(i).LeftLine
        alignedRight(i) = ""
      Case DiffOpInsert
        alignedLeft(i) = ""
        alignedRight(i) = ops.Items(i).RightLine
    End Select
  Next i
End Sub

Private Sub ExportOps(ByRef vec As LineOpVector, ByRef output() As AlignOp)
  Dim i As Long

  If vec.Count = 0 Then
    Erase output
    Exit Sub
  End If

  ReDim output(0 To vec.Count - 1)
  For i = 0 To vec.Count - 1
    output(i) = vec.Items(i)
  Next i
End Sub

' -----------------------------------------------------------------------------
' ベクタ / 配列ヘルパ
' -----------------------------------------------------------------------------

Private Function SliceStringArray(ByRef source() As String, ByVal startIndex As Long, ByVal endExclusive As Long) As String()
  Dim count As Long
  Dim result() As String
  Dim i As Long

  count = endExclusive - startIndex
  If count <= 0 Then
    Erase result
    SliceStringArray = result
    Exit Function
  End If

  ReDim result(0 To count - 1)
  For i = 0 To count - 1
    result(i) = source(startIndex + i)
  Next i

  SliceStringArray = result
End Function

Private Function ReverseOps(ByRef inputVec As LineOpVector) As LineOpVector
  Dim result As LineOpVector
  Dim i As Long

  For i = inputVec.Count - 1 To 0 Step -1
    PushOp result, inputVec.Items(i)
  Next i

  ReverseOps = result
End Function

Private Function ReversePairs(ByRef inputVec As UniquePairVector) As UniquePairVector
  Dim result As UniquePairVector
  Dim i As Long

  For i = inputVec.Count - 1 To 0 Step -1
    PushPair result, inputVec.Items(i)
  Next i

  ReversePairs = result
End Function

Private Sub PushOp(ByRef vec As LineOpVector, ByRef op As AlignOp)
  EnsureOpCapacity vec, vec.Count + 1
  vec.Items(vec.Count) = op
  vec.Count = vec.Count + 1
End Sub

Private Sub AppendOpVector(ByRef target As LineOpVector, ByRef source As LineOpVector)
  Dim i As Long
  If source.Count = 0 Then Exit Sub
  EnsureOpCapacity target, target.Count + source.Count
  For i = 0 To source.Count - 1
    target.Items(target.Count) = source.Items(i)
    target.Count = target.Count + 1
  Next i
End Sub

Private Sub EnsureOpCapacity(ByRef vec As LineOpVector, ByVal minCapacity As Long)
  Dim newCapacity As Long

  If vec.Capacity >= minCapacity Then Exit Sub

  If vec.Capacity = 0 Then
    newCapacity = 16
  Else
    newCapacity = vec.Capacity
  End If

  Do While newCapacity < minCapacity
    newCapacity = newCapacity * 2
  Loop

  ReDim Preserve vec.Items(0 To newCapacity - 1)
  vec.Capacity = newCapacity
End Sub

Private Sub ClearOpVector(ByRef vec As LineOpVector)
  vec.Count = 0
End Sub

Private Sub PushPair(ByRef vec As UniquePairVector, ByRef pair As UniquePair)
  EnsurePairCapacity vec, vec.Count + 1
  vec.Items(vec.Count) = pair
  vec.Count = vec.Count + 1
End Sub

Private Sub EnsurePairCapacity(ByRef vec As UniquePairVector, ByVal minCapacity As Long)
  Dim newCapacity As Long

  If vec.Capacity >= minCapacity Then Exit Sub

  If vec.Capacity = 0 Then
    newCapacity = 8
  Else
    newCapacity = vec.Capacity
  End If

  Do While newCapacity < minCapacity
    newCapacity = newCapacity * 2
  Loop

  ReDim Preserve vec.Items(0 To newCapacity - 1)
  vec.Capacity = newCapacity
End Sub

Private Sub SortPairsByLeft(ByRef vec As UniquePairVector)
  Dim i As Long
  Dim j As Long
  Dim tmp As UniquePair

  For i = 1 To vec.Count - 1
    tmp = vec.Items(i)
    j = i - 1
    Do While j >= 0 And vec.Items(j).LeftIndex > tmp.LeftIndex
      vec.Items(j + 1) = vec.Items(j)
      j = j - 1
    Loop
    vec.Items(j + 1) = tmp
  Next i
End Sub

Private Sub PushCandidate(ByRef vec As PairCandidateVector, ByRef candidate As PairCandidate)
  EnsureCandidateCapacity vec, vec.Count + 1
  vec.Items(vec.Count) = candidate
  vec.Count = vec.Count + 1
End Sub

Private Sub EnsureCandidateCapacity(ByRef vec As PairCandidateVector, ByVal minCapacity As Long)
  Dim newCapacity As Long

  If vec.Capacity >= minCapacity Then Exit Sub

  If vec.Capacity = 0 Then
    newCapacity = 32
  Else
    newCapacity = vec.Capacity
  End If

  Do While newCapacity < minCapacity
    newCapacity = newCapacity * 2
  Loop

  ReDim Preserve vec.Items(0 To newCapacity - 1)
  vec.Capacity = newCapacity
End Sub

Private Sub SortCandidates(ByRef vec As PairCandidateVector)
  If vec.Count <= 1 Then Exit Sub
  QuickSortCandidates vec, 0, vec.Count - 1
End Sub

Private Sub QuickSortCandidates(ByRef vec As PairCandidateVector, ByVal lo As Long, ByVal hi As Long)
  Dim i As Long
  Dim j As Long
  Dim pivot As PairCandidate
  Dim tmp As PairCandidate

  i = lo
  j = hi
  pivot = vec.Items((lo + hi) \ 2)

  Do While i <= j
    Do While CompareCandidates(vec.Items(i), pivot) < 0
      i = i + 1
    Loop

    Do While CompareCandidates(vec.Items(j), pivot) > 0
      j = j - 1
    Loop

    If i <= j Then
      tmp = vec.Items(i)
      vec.Items(i) = vec.Items(j)
      vec.Items(j) = tmp
      i = i + 1
      j = j - 1
    End If
  Loop

  If lo < j Then QuickSortCandidates vec, lo, j
  If i < hi Then QuickSortCandidates vec, i, hi
End Sub

Private Function CompareCandidates(ByRef a As PairCandidate, ByRef b As PairCandidate) As Long
  If a.Score <> b.Score Then
    CompareCandidates = Sgn(b.Score - a.Score)
    Exit Function
  End If
  If a.IndentDiff <> b.IndentDiff Then
    CompareCandidates = Sgn(a.IndentDiff - b.IndentDiff)
    Exit Function
  End If
  If a.Distance <> b.Distance Then
    CompareCandidates = Sgn(a.Distance - b.Distance)
    Exit Function
  End If
  If a.DeleteIndex <> b.DeleteIndex Then
    CompareCandidates = Sgn(a.DeleteIndex - b.DeleteIndex)
    Exit Function
  End If
  CompareCandidates = Sgn(a.InsertIndex - b.InsertIndex)
End Function

Private Function MakeEqualOp(ByVal leftLine As String, ByVal rightLine As String, ByVal leftNo As Long, ByVal rightNo As Long) As AlignOp
  Dim op As AlignOp
  op.OpType = DiffOpEqual
  op.LeftLine = leftLine
  op.RightLine = rightLine
  op.LeftLineNo = leftNo
  op.RightLineNo = rightNo
  op.HasLeftLineNo = True
  op.HasRightLineNo = True
  MakeEqualOp = op
End Function

Private Function MakeInsertOp(ByVal rightLine As String, ByVal rightNo As Long) As AlignOp
  Dim op As AlignOp
  op.OpType = DiffOpInsert
  op.RightLine = rightLine
  op.RightLineNo = rightNo
  op.HasRightLineNo = True
  MakeInsertOp = op
End Function

Private Function MakeDeleteOp(ByVal leftLine As String, ByVal leftNo As Long) As AlignOp
  Dim op As AlignOp
  op.OpType = DiffOpDelete
  op.LeftLine = leftLine
  op.LeftLineNo = leftNo
  op.HasLeftLineNo = True
  MakeDeleteOp = op
End Function

Private Function MakeReplaceOp(ByVal leftLine As String, ByVal rightLine As String, ByVal leftNo As Long, ByVal rightNo As Long, ByVal hasLeftNo As Boolean, ByVal hasRightNo As Boolean) As AlignOp
  Dim op As AlignOp
  op.OpType = DiffOpReplace
  op.LeftLine = leftLine
  op.RightLine = rightLine
  op.LeftLineNo = leftNo
  op.RightLineNo = rightNo
  op.HasLeftLineNo = hasLeftNo
  op.HasRightLineNo = hasRightNo
  MakeReplaceOp = op
End Function

Private Function CountIndent(ByVal line As String) As Long
  Dim i As Long
  Dim c As String

  For i = 1 To Len(line)
    c = Mid$(line, i, 1)
    If c = " " Or c = vbTab Then
      CountIndent = CountIndent + 1
    Else
      Exit For
    End If
  Next i
End Function

Private Function NzString(ByRef source() As String, ByVal index As Long) As String
  On Error GoTo EmptyValue
  NzString = source(index)
  Exit Function
EmptyValue:
  NzString = ""
End Function

Private Function RegexTest(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean) As Boolean
  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = False
  RegexTest = re.Test(text)
End Function
