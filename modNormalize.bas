Attribute VB_Name = "modNormalize"
Option Explicit

' =============================================================================
' modNormalize
'   diff-viewer/src/diffEngine の比較前処理を VBA に移植したモジュール。
'   行対応の精度に関わるため、ここで比較キーを作る。
' =============================================================================

Private mLineSignatureKeywords As Object

Public Function NormalizeText(ByVal text As String) As String
  ' TypeScript: normalize.ts#normalizeText
  NormalizeText = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

Public Function NormalizeComparisonWhitespace(ByVal text As String) As String
  ' 追加仕様: 比較キーでは末尾空白を無視し、タブはスペース化、連続スペースを圧縮。
  Dim s As String
  s = Replace(text, vbTab, " ")
  s = RTrim$(s)
  s = RegexReplace(s, " {2,}", " ", True, True)
  NormalizeComparisonWhitespace = s
End Function

Public Function StripRazorLinePrefix(ByVal line As String) As String
  ' TypeScript: diffLines.ts / lineSignature.ts の stripRazorLinePrefix を移植。
  Dim i As Long
  Dim j As Long
  i = 1

  Do While i <= Len(line)
    Select Case Mid$(line, i, 1)
      Case " ", vbTab
        i = i + 1
      Case Else
        Exit Do
    End Select
  Loop

  If Mid$(line, i, 2) <> "@:" Then
    StripRazorLinePrefix = line
    Exit Function
  End If

  j = i + 2
  Do While j <= Len(line)
    Select Case Mid$(line, j, 1)
      Case " ", vbTab
        j = j + 1
      Case Else
        Exit Do
    End Select
  Loop

  StripRazorLinePrefix = Left$(line, i - 1) & Mid$(line, j)
End Function

Public Function IsBlankLine(ByVal line As String) As Boolean
  IsBlankLine = (Len(Trim$(Replace(line, vbTab, " "))) = 0)
End Function

Public Function NormalizeForMatch(ByVal line As String) As String
  ' TypeScript: diffLines.ts#normalizeForMatch を移植。
  Dim trimmed As String
  Dim initVar As String
  Dim appendLiteral As String
  Dim literal As String

  trimmed = StripRazorLinePrefix(line)
  trimmed = LTrimWhitespace(trimmed)

  If IsBlankLine(trimmed) Then
    NormalizeForMatch = ""
    Exit Function
  End If

  initVar = ExtractInitVariable(trimmed)
  If Len(initVar) > 0 Then
    NormalizeForMatch = "init:" & LCase$(initVar)
    Exit Function
  End If

  If IsAppendLike(trimmed) Then
    appendLiteral = ExtractAppendLiteral(trimmed)
    If Len(appendLiteral) > 0 Then
      NormalizeForMatch = "append:" & LCase$(NormalizeComparisonWhitespace(appendLiteral))
      Exit Function
    End If
  End If

  literal = ExtractFirstLiteral(trimmed)
  If Len(literal) > 0 And IsAppendLike(trimmed) Then
    NormalizeForMatch = "append:" & LCase$(NormalizeComparisonWhitespace(literal))
    Exit Function
  End If

  NormalizeForMatch = NormalizeComparisonWhitespace(trimmed)
End Function

Public Function ToAppendLiteralOrLine(ByVal line As String) As String
  ' TypeScript: appendLiteral.ts#toAppendLiteralOrLine
  Dim payload As String
  payload = ExtractAppendLiteral(line)
  If Len(payload) > 0 Then
    ToAppendLiteralOrLine = payload
  Else
    ToAppendLiteralOrLine = line
  End If
End Function

Public Function ExtractAppendLiteral(ByVal line As String) As String
  ' TypeScript: appendLiteral.ts#extractAppendLiteral
  ExtractAppendLiteral = ParseAppendLiteral(line, False)
End Function

Public Function ExtractLineKey(ByVal line As String) As String
  ' TypeScript: lineSignature.ts#extractLineKey
  Dim normalizedLine As String
  Dim braceToken As String
  Dim normalized As String
  Dim candidate As String
  Dim tokens As Collection
  Dim i As Long

  normalizedLine = StripRazorLinePrefix(line)
  braceToken = ExtractBraceToken(normalizedLine)
  If Len(braceToken) > 0 Then
    ExtractLineKey = braceToken
    Exit Function
  End If

  normalized = StripDollarIdentifiers(LTrimWhitespace(normalizedLine))

  candidate = RegexFirstGroup(normalized, "([A-Za-z_][A-Za-z0-9_]*)\s*\(", True)
  If Len(candidate) > 0 Then
    candidate = LCase$(candidate)
    If Not IsLineSignatureKeyword(candidate) Then
      ExtractLineKey = candidate
      Exit Function
    End If
  End If

  Set tokens = RegexMatches(normalized, "[A-Za-z_][A-Za-z0-9_]*", True)
  For i = 1 To tokens.Count
    candidate = LCase$(tokens(i))
    If Not IsLineSignatureKeyword(candidate) Then
      ExtractLineKey = candidate
      Exit Function
    End If
  Next i

  ExtractLineKey = ""
End Function

Public Function BuildSimpleSimilarityKey(ByVal line As String) As String
  ' Pairing 用の軽量キー。出力文字列は元の行を保持し、比較キーだけを正規化する。
  Dim core As String
  core = ToAppendLiteralOrLine(StripRazorLinePrefix(line))
  core = NormalizeComparisonWhitespace(LTrimWhitespace(core))
  BuildSimpleSimilarityKey = LCase$(core)
End Function

Public Function ExtractIdentifierTokens(ByVal line As String) As Collection
  Dim normalized As String
  Dim matches As Collection
  Dim result As New Collection
  Dim i As Long
  Dim token As String

  normalized = StripDollarIdentifiers(LTrimWhitespace(line))
  Set matches = RegexMatches(normalized, "[A-Za-z_][A-Za-z0-9_]*", True)

  For i = 1 To matches.Count
    token = LCase$(matches(i))
    If Not IsLineSignatureKeyword(token) Then
      On Error Resume Next
      result.Add token, token
      On Error GoTo 0
    End If
  Next i

  Set ExtractIdentifierTokens = result
End Function

Private Function ParseAppendLiteral(ByVal line As String, ByVal preserveEscapes As Boolean) As String
  ' TypeScript: appendLiteral.ts#parseAppendLiteral を移植。
  Dim callStart As Long
  Dim callMatch As String
  Dim startFrom As Long
  Dim quotePos As Long
  Dim prefix As String
  Dim isInterpolated As Boolean
  Dim isVerbatim As Boolean
  Dim i As Long
  Dim n As Long
  Dim ch As String
  Dim nextCh As String
  Dim resultText As String

  callMatch = RegexFirstMatch(line, "\.(append|appendline|appendformat)\s*\(", True)
  If Len(callMatch) = 0 Then
    ParseAppendLiteral = ""
    Exit Function
  End If

  callStart = RegexFirstIndex(line, "\.(append|appendline|appendformat)\s*\(", True)
  If callStart < 0 Then
    ParseAppendLiteral = ""
    Exit Function
  End If

  startFrom = callStart + Len(callMatch) + 1
  quotePos = InStr(startFrom, line, """")
  If quotePos = 0 Then
    ParseAppendLiteral = ""
    Exit Function
  End If

  prefix = Mid$(line, startFrom, quotePos - startFrom)
  isInterpolated = (InStr(1, prefix, "$") > 0)
  isVerbatim = (InStr(1, prefix, "@") > 0)

  i = quotePos + 1
  n = Len(line)

  Do While i <= n
    ch = Mid$(line, i, 1)

    If (Not isVerbatim) And ch = "\" And i < n Then
      nextCh = Mid$(line, i + 1, 1)
      If preserveEscapes Then
        resultText = resultText & "\" & nextCh
      Else
        Select Case nextCh
          Case "t"
            resultText = resultText & vbTab
          Case "n"
            resultText = resultText & vbLf
          Case "r"
            resultText = resultText & vbCr
          Case Else
            resultText = resultText & nextCh
        End Select
      End If
      i = i + 2
      GoTo ContinueLoop
    End If

    If isVerbatim And ch = """" And i < n And Mid$(line, i + 1, 1) = """" Then
      resultText = resultText & """"
      i = i + 2
      GoTo ContinueLoop
    End If

    If ch = """" Then
      Exit Do
    End If

    If isInterpolated And ch = "{" Then
      If i < n And Mid$(line, i + 1, 1) = "{" Then
        resultText = resultText & "{"
        i = i + 2
        GoTo ContinueLoop
      End If

      resultText = resultText & "{expr}"
      i = i + 1
      SkipInterpolatedBlock line, i, isVerbatim
      GoTo ContinueLoop
    End If

    If isInterpolated And ch = "}" And i < n And Mid$(line, i + 1, 1) = "}" Then
      resultText = resultText & "}"
      i = i + 2
      GoTo ContinueLoop
    End If

    resultText = resultText & ch
    i = i + 1
ContinueLoop:
  Loop

  If Len(Trim$(resultText)) = 0 Then
    ParseAppendLiteral = ""
  Else
    ParseAppendLiteral = resultText
  End If
End Function

Private Sub SkipInterpolatedBlock(ByVal source As String, ByRef indexPos As Long, ByVal isVerbatim As Boolean)
  ' { ... } のネストを壊さないための最小パーサ。
  Dim depth As Long
  Dim n As Long
  Dim ch As String

  depth = 1
  n = Len(source)

  Do While indexPos <= n And depth > 0
    ch = Mid$(source, indexPos, 1)

    If ch = "{" Then
      If indexPos = n Or Mid$(source, indexPos + 1, 1) <> "{" Then
        depth = depth + 1
      Else
        indexPos = indexPos + 1
      End If
    ElseIf ch = "}" Then
      If indexPos = n Or Mid$(source, indexPos + 1, 1) <> "}" Then
        depth = depth - 1
      Else
        indexPos = indexPos + 1
      End If
    ElseIf ch = """" And (Not isVerbatim) Then
      indexPos = indexPos + 1
      SkipQuotedString source, indexPos
      GoTo ContinueLoop
    End If

    indexPos = indexPos + 1
ContinueLoop:
  Loop
End Sub

Private Sub SkipQuotedString(ByVal source As String, ByRef indexPos As Long)
  Dim n As Long
  Dim ch As String

  n = Len(source)
  Do While indexPos <= n
    ch = Mid$(source, indexPos, 1)
    If ch = "\" And indexPos < n Then
      indexPos = indexPos + 2
      GoTo ContinueLoop
    End If
    If ch = """" Then
      indexPos = indexPos + 1
      Exit Do
    End If
    indexPos = indexPos + 1
ContinueLoop:
  Loop
End Sub

Private Function ExtractInitVariable(ByVal line As String) As String
  Dim csharpVar As String
  Dim phpVar As String

  csharpVar = RegexFirstGroup(line, "\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*new\b", True)
  If Len(csharpVar) > 0 Then
    ExtractInitVariable = LCase$(csharpVar)
    Exit Function
  End If

  phpVar = RegexFirstGroup(line, "\$([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(['""])\s*\2", True)
  If Len(phpVar) > 0 Then
    ExtractInitVariable = LCase$(phpVar)
    Exit Function
  End If

  ExtractInitVariable = ""
End Function

Private Function ExtractFirstLiteral(ByVal line As String) As String
  Dim m As String
  m = RegexFirstGroup(line, "'(([^'\\]|\\.)*)'|""(([^""\\]|\\.)*)""", True, 0)
  If Len(m) = 0 Then
    ExtractFirstLiteral = ""
    Exit Function
  End If

  ' RegexFirstGroup(, groupIndex:=0) なので、引用符付きで返る。
  If Len(m) >= 2 Then
    m = Mid$(m, 2, Len(m) - 2)
  End If
  m = RegexReplace(m, "\s+", " ", True, True)
  ExtractFirstLiteral = LCase$(m)
End Function

Private Function IsAppendLike(ByVal line As String) As Boolean
  IsAppendLike = _
    RegexTest(line, "\.(append|appendline|appendformat)\s*\(", True) Or _
    RegexTest(line, "\.\=", True) Or _
    (RegexTest(line, "\+=", True) And HasStringLiteral(line))
End Function

Private Function HasStringLiteral(ByVal line As String) As Boolean
  HasStringLiteral = RegexTest(line, "'([^'\\]|\\.)*'|""([^""\\]|\\.)*""", True)
End Function

Private Function StripDollarIdentifiers(ByVal line As String) As String
  StripDollarIdentifiers = RegexReplace(line, "\$([A-Za-z_][A-Za-z0-9_]*)", "$1", True, True)
End Function

Private Function ExtractBraceToken(ByVal line As String) As String
  Dim trimmed As String
  Dim body As String

  trimmed = Trim$(line)
  If trimmed = "{" Then
    ExtractBraceToken = "brace_open"
    Exit Function
  End If
  If trimmed = "}" Then
    ExtractBraceToken = "brace_close"
    Exit Function
  End If

  body = TryExtractPhpBraceBody(trimmed)
  If body = "{" Then
    ExtractBraceToken = "brace_open"
    Exit Function
  End If
  If body = "}" Then
    ExtractBraceToken = "brace_close"
    Exit Function
  End If
  ExtractBraceToken = ""
End Function

Private Function TryExtractPhpBraceBody(ByVal trimmed As String) As String
  Dim s As String

  s = LCase$(trimmed)
  If Left$(s, 2) <> "<?" Then
    TryExtractPhpBraceBody = ""
    Exit Function
  End If

  s = Mid$(s, 3)
  s = Trim$(s)

  If Left$(s, 3) = "php" Then
    s = Trim$(Mid$(s, 4))
  End If

  If Right$(s, 2) = "?>" Then
    s = Left$(s, Len(s) - 2)
  End If
  If Right$(s, 1) = ">" Then
    s = Left$(s, Len(s) - 1)
  End If

  s = Trim$(s)
  If s = "{" Or s = "}" Then
    TryExtractPhpBraceBody = s
  Else
    TryExtractPhpBraceBody = ""
  End If
End Function

Private Function IsLineSignatureKeyword(ByVal tokenLower As String) As Boolean
  If mLineSignatureKeywords Is Nothing Then
    InitializeLineSignatureKeywords
  End If
  IsLineSignatureKeyword = mLineSignatureKeywords.Exists(tokenLower)
End Function

Private Sub InitializeLineSignatureKeywords()
  Dim values As Variant
  Dim i As Long

  values = Array( _
    "var", "let", "const", "function", "fn", "string", "int", "float", "double", _
    "bool", "boolean", "void", "public", "private", "protected", "static", "final", _
    "abstract", "async", "await", "return", "class", "interface", "trait", "extends", _
    "implements", "new", "namespace", "using", "if", "else", "elseif", "for", "foreach", _
    "while", "do", "switch", "case", "break", "continue", "try", "catch", "finally", _
    "throw", "include", "require", "global", "internal", "readonly", "override", "virtual", _
    "define" _
  )

  Set mLineSignatureKeywords = CreateObject("Scripting.Dictionary")
  mLineSignatureKeywords.CompareMode = 1 ' TextCompare

  For i = LBound(values) To UBound(values)
    mLineSignatureKeywords(values(i)) = True
  Next i
End Sub

Public Function LTrimWhitespace(ByVal text As String) As String
  Dim i As Long
  i = 1
  Do While i <= Len(text)
    Select Case Mid$(text, i, 1)
      Case " ", vbTab
        i = i + 1
      Case Else
        Exit Do
    End Select
  Loop
  LTrimWhitespace = Mid$(text, i)
End Function

Private Function RegexTest(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean) As Boolean
  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = False
  RegexTest = re.Test(text)
End Function

Private Function RegexReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, ByVal ignoreCase As Boolean, ByVal globalMatch As Boolean) As String
  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = globalMatch
  RegexReplace = re.Replace(text, replacement)
End Function

Private Function RegexFirstGroup(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean, Optional ByVal groupIndex As Long = 1) As String
  Dim re As Object
  Dim matches As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = False

  Set matches = re.Execute(text)
  If matches.Count = 0 Then
    RegexFirstGroup = ""
    Exit Function
  End If

  If groupIndex = 0 Then
    RegexFirstGroup = matches(0).Value
    Exit Function
  End If

  If matches(0).SubMatches.Count < groupIndex Then
    RegexFirstGroup = ""
    Exit Function
  End If

  RegexFirstGroup = CStr(matches(0).SubMatches(groupIndex - 1))
End Function

Private Function RegexFirstIndex(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean) As Long
  Dim re As Object
  Dim matches As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = False

  Set matches = re.Execute(text)
  If matches.Count = 0 Then
    RegexFirstIndex = -1
  Else
    RegexFirstIndex = CLng(matches(0).FirstIndex)
  End If
End Function

Private Function RegexFirstMatch(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean) As String
  Dim re As Object
  Dim matches As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = False

  Set matches = re.Execute(text)
  If matches.Count = 0 Then
    RegexFirstMatch = ""
  Else
    RegexFirstMatch = CStr(matches(0).Value)
  End If
End Function

Private Function RegexMatches(ByVal text As String, ByVal pattern As String, ByVal ignoreCase As Boolean) As Collection
  Dim re As Object
  Dim matches As Object
  Dim result As New Collection
  Dim i As Long

  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = ignoreCase
  re.Global = True

  Set matches = re.Execute(text)
  For i = 0 To matches.Count - 1
    result.Add CStr(matches(i).Value)
  Next i

  Set RegexMatches = result
End Function
