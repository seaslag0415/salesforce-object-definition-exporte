Attribute VB_Name = "modJsonParser"
Option Explicit

Private m_Text As String
Private m_Pos As Long

Public Function ParseJson(ByVal jsonText As String) As Object
    m_Text = jsonText
    m_Pos = 1
    SkipWhitespace

    If PeekChar() <> "{" Then
        Err.Raise vbObjectError + 200, , "JSON root is not an object."
    End If

    Set ParseJson = ParseObject()
End Function

Private Function ParseObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Consume "{"
    SkipWhitespace

    If PeekChar() = "}" Then
        Consume "}"
        Set ParseObject = dict
        Exit Function
    End If

    Do
        Dim key As String
        key = ParseString()

        Consume ":"
        AddParsedValueToDictionary dict, key

        SkipWhitespace
        If PeekChar() = "}" Then
            Consume "}"
            Exit Do
        End If
        Consume ","
    Loop

    Set ParseObject = dict
End Function

Private Function ParseArray() As Collection
    Dim col As Collection
    Set col = New Collection

    Consume "["
    SkipWhitespace

    If PeekChar() = "]" Then
        Consume "]"
        Set ParseArray = col
        Exit Function
    End If

    Do
        AddParsedValueToCollection col

        SkipWhitespace
        If PeekChar() = "]" Then
            Consume "]"
            Exit Do
        End If
        Consume ","
    Loop

    Set ParseArray = col
End Function

Private Sub AddParsedValueToDictionary(ByVal dict As Object, ByVal key As String)
    SkipWhitespace

    Select Case PeekChar()
        Case "{"
            Set dict(key) = ParseObject()
        Case "["
            Set dict(key) = ParseArray()
        Case """"
            dict(key) = ParseString()
        Case "t"
            ExpectLiteral "true"
            dict(key) = True
        Case "f"
            ExpectLiteral "false"
            dict(key) = False
        Case "n"
            ExpectLiteral "null"
            dict(key) = Null
        Case Else
            dict(key) = ParseNumber()
    End Select
End Sub

Private Sub AddParsedValueToCollection(ByVal col As Collection)
    SkipWhitespace

    Select Case PeekChar()
        Case "{"
            col.Add ParseObject()
        Case "["
            col.Add ParseArray()
        Case """"
            col.Add ParseString()
        Case "t"
            ExpectLiteral "true"
            col.Add True
        Case "f"
            ExpectLiteral "false"
            col.Add False
        Case "n"
            ExpectLiteral "null"
            col.Add Null
        Case Else
            col.Add ParseNumber()
    End Select
End Sub

Private Function ParseString() As String
    Consume """"

    Dim result As String
    result = ""

    Do While m_Pos <= Len(m_Text)
        Dim ch As String
        ch = Mid$(m_Text, m_Pos, 1)
        m_Pos = m_Pos + 1

        If ch = """" Then
            ParseString = result
            Exit Function
        ElseIf ch = "\" Then
            If m_Pos > Len(m_Text) Then Err.Raise vbObjectError + 203, , "Invalid JSON escape."

            Dim esc As String
            esc = Mid$(m_Text, m_Pos, 1)
            m_Pos = m_Pos + 1

            Select Case esc
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "/": result = result & "/"
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u": result = result & ParseUnicodeEscape()
                Case Else
                    Err.Raise vbObjectError + 204, , "Unsupported JSON escape: \" & esc
            End Select
        Else
            result = result & ch
        End If
    Loop

    Err.Raise vbObjectError + 205, , "Unterminated JSON string."
End Function

Private Function ParseUnicodeEscape() As String
    If m_Pos + 3 > Len(m_Text) Then Err.Raise vbObjectError + 206, , "Invalid unicode escape."

    Dim hexValue As String
    hexValue = Mid$(m_Text, m_Pos, 4)
    m_Pos = m_Pos + 4

    ParseUnicodeEscape = ChrW$(CLng("&H" & hexValue))
End Function

Private Function ParseNumber() As Variant
    Dim startPos As Long
    startPos = m_Pos

    Do While m_Pos <= Len(m_Text)
        Dim ch As String
        ch = Mid$(m_Text, m_Pos, 1)
        If InStr(1, "0123456789+-.eE", ch, vbBinaryCompare) = 0 Then Exit Do
        m_Pos = m_Pos + 1
    Loop

    Dim token As String
    token = Mid$(m_Text, startPos, m_Pos - startPos)
    If token = "" Then Err.Raise vbObjectError + 202, , "Invalid JSON token near position " & m_Pos & "."

    If InStr(1, token, ".", vbBinaryCompare) > 0 Or InStr(1, token, "e", vbTextCompare) > 0 Then
        ParseNumber = CDbl(token)
    Else
        ParseNumber = CDbl(token)
    End If
End Function

Private Sub ExpectLiteral(ByVal literal As String)
    If Mid$(m_Text, m_Pos, Len(literal)) <> literal Then
        Err.Raise vbObjectError + 207, , "Expected JSON literal: " & literal
    End If
    m_Pos = m_Pos + Len(literal)
End Sub

Private Sub Consume(ByVal expected As String)
    SkipWhitespace

    If PeekChar() <> expected Then
        Err.Raise vbObjectError + 208, , "Expected '" & expected & "' near position " & m_Pos & "."
    End If

    m_Pos = m_Pos + 1
End Sub

Private Function PeekChar() As String
    If m_Pos > Len(m_Text) Then
        PeekChar = ""
    Else
        PeekChar = Mid$(m_Text, m_Pos, 1)
    End If
End Function

Private Sub SkipWhitespace()
    Do While m_Pos <= Len(m_Text)
        Dim ch As String
        ch = Mid$(m_Text, m_Pos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        m_Pos = m_Pos + 1
    Loop
End Sub
