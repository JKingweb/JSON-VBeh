VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Enum ObjExp
    ExpInit = 0
    ExpColon = 1
    ExpValue = 2
    ExpComma = 3
    ExpKey = 4
End Enum

Public Enum jsonDuplicateKeyStrategy
    jsonKeyLast = 0
    jsonKeyFirst = 1
    jsonKeyHalt = 2
End Enum

Public Enum jsonFailMode
    jsonFailStop = 0
    jsonFailNull = 1
    jsonFailConv = 10
End Enum

Public DuplicateKeys As jsonDuplicateKeyStrategy
Public SerializeFailMode As jsonFailMode
Public DateFormat As String

Private EOD As Long
Private Pos As Long

Private Sub Class_Initialize()
    DateFormat = "yyyy""-""mm""-""dd hh"":""mm"":""ss"
End Sub

Public Function Serialize(Value) As String
    Dim a As Long
    Dim Conv As Boolean
    Conv = (SerializeFailMode \ jsonFailConv) > 0
    Select Case VarType(Value)
        Case vbNull, vbEmpty
            Serialize = "null"
        Case vbBoolean
            If Value Then
                Serialize = "true"
            Else
                Serialize = "false"
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Serialize = CStr(Value)
        Case vbString
            Serialize = MakeStr(Value)
        Case vbDate
            If Conv Then
                Serialize = """" & Format(Value, DateFormat) & """"
            Else
                Serialize = GenSerErr(100)
            End If
        Case vbError
            If Conv Then
                Serialize = CLng(Value)
            Else
                Serialize = GenSerErr(101)
            End If
        Case vbUserDefinedType
            Serialize = GenSerErr(102)
        Case vbDataObject
            ' I am unclear what this type is for
            Serialize = GenSerErr(103)
        Case vbObject
            Select Case TypeName(Value)
                Case "Collection"
                    Serialize = "["
                    If Value.Count > 0 Then
                        Serialize = Serialize & Serialize(Value(1))
                        For a = 2 To Value.Count
                            Serialize = Serialize & "," & Serialize(Value(1))
                        Next
                    End If
                    Serialize = Serialize & "]"
                Case "Dictionary"
                    Dim Keys, Key
                    Keys = Value.Keys
                    Serialize = "{"
                    If Value.Count > 0 Then
                        Serialize = Serialize & MakeStr(Keys(LBound(Keys))) & ":" & Serialize(Value(Keys(LBound(Keys))))
                        For a = LBound(Keys) + 1 To UBound(Keys)
                            Serialize = Serialize & "," & MakeStr(Keys(a)) & ":" & Serialize(Value(Keys(a)))
                        Next
                    End If
                    Serialize = Serialize & "}"
                Case Else
                    If Conv Then
                        On Error GoTo ObjectNotConvertibleError
                        Serialize = CStr(Value)
                        On Error GoTo 0
                    Else
                        Serialize = GenSerErr(104, TypeName(Value))
                    End If
            End Select
        Case Is > vbArray
            Dim Sizes As Variant, Pointers As Variant, Member As Variant
            Dim U As Long, L As Long
            Dim RepBefore As Integer, RepAfter As Integer
            Dim First As Boolean, FirstMember As Boolean
            FirstMember = True
            On Error GoTo DimensionSearchError
            Do
                a = a + 1
                U = UBound(Value, a)
                L = LBound(Value, a)
                If L > -1 And U > -1 Then
                    Push Sizes, Array(L, U)
                    Push Pointers, -100
                    GoTo ContinueDimensionSearch
                End If
StopDimensionSearch:
                Exit Do
ContinueDimensionSearch:
            Loop
            On Error GoTo 0
            For Each Member In Value
                RepBefore = 0
                RepAfter = 0
                For a = UBound(Pointers) To LBound(Pointers) Step -1
                    If Pointers(a) < Sizes(a)(1) Then
                        RepBefore = RepBefore + 1
                        Pointers(a) = Sizes(a)(1)
                        If a > LBound(Pointers) Then Pointers(a - 1) = Pointers(a - 1) + 1
                    ElseIf Pointers(a) > Sizes(a)(2) Then
                        RepBefore = RepBefore + 1
                        RepAfter = RepAfter + 1
                        Pointers(a) = Sizes(a)(1)
                        If a > LBound(Pointers) Then Pointers(a - 1) = Pointers(a - 1) + 1
                    Else
                        Exit For
                    End If
                Next
                First = RepBefore > 0
                Pointers(UBound(Pointers)) = Pointers(UBound(Pointers)) + 1
                Serialize = Serialize & String(RepAfter, "]") & IIf(Not FirstMember And RepBefore > 0, ",", "") & String(RepBefore, "[") & IIf(Not First, ",", "") & Serialize(Member)
                FirstMember = False
            Next
            If Serialize <> "" Then
                Serialize = Serialize & String(UBound(Pointers), "]")
            Else
                Serialize = "[]"
            End If
    End Select
    Exit Function
DimensionSearchError:
    Resume StopDimensionSearch
ObjectNotConvertibleError:
    Serialize = GenSerErr(104, TypeName(Value))
    Resume Next
End Function

Public Function Parse(Data) As Variant
    Pos = 1
    EOD = Len(Data)
    ' Skip byte order mark
    If Mid(Data, Pos, 1) = ChrW(65279) Then Incr
    ' skip any leading whitespace
    ConsumeWhitespace Data
    ' check for legal characters
    Select Case Mid(Data, Pos, 1)
        Case ""
            ' no data present; this is an error
            GenParseErr 0, Pos
        Case "{"
            ' object
            Incr
            Set Parse = ConsumeObject(Data)
        Case "["
            ' array
            Incr
            Parse = ConsumeArray(Data)
        Case """"
            ' string
            Incr
            Parse = ConsumeString(Data)
        Case 0 To 9, "-"
            ' possible number
            Parse = ConsumeNumber(Data)
        Case "t", "f", "n"
            ' possible boolean or null
            Parse = ConsumeLiteral(Data)
        Case Else
            ' anything else is illegal
            GenParseErr 1
    End Select
    ' skip whitespace at the end
    ConsumeWhitespace Data
    ' any remaining characters are garbage; this is an error
    If Mid(Data, Pos, 1) <> "" Then GenParseErr 2
End Function

Public Function IsObject(Data) As Boolean
    '' checks to see if the return value will be a
    '' JSON object (implemented as a Dictionary),
    '' which will need to be Set unlike other types
    Pos = 1
    EOD = Len(Data) + 1
    ' Skip byte order mark
    If Mid(Data, Pos, 1) = ChrW(65279) Then Incr
    ' skip any leading whitespace
    ConsumeWhitespace Data
    ' check for curly brace
    If Mid(Data, Pos, 1) = "{" Then
        IsObject = True
    Else
        IsObject = False
    End If
End Function

Private Function ConsumeNumber(Data) As Variant
    Dim Buffer As String, Char As String
    Dim Exponent As Long
    Dim FirstSeen As Boolean, Zero As Boolean, Signed As Boolean, Fractional As Boolean, Raised As Boolean
    Do
        Char = Mid(Data, Pos, 1)
        Select Case Char
            Case "0"
                Buffer = Buffer & Char
                If Not FirstSeen Then Zero = True
            Case "-"
                If FirstSeen Or Signed Then Exit Do
                Signed = True
                Buffer = Buffer & Char
            Case 1 To 9
                If Zero Then Incr -1: GenParseErr 3
                If Not FirstSeen Then FirstSeen = True
                Buffer = Buffer & Char
            Case "."
                If Not (FirstSeen Or Zero) Then Exit Do
                Fractional = True
                Buffer = Buffer & Char
            Case "e", "E"
                If Not (FirstSeen Or Zero) Then Exit Do
                Raised = True
                Incr
                Exponent = ConsumeExponent(Data)
                Exit Do
            Case Else
                Exit Do
        End Select
        Incr
    Loop Until Pos > EOD
    If Not (FirstSeen Or Zero) Then GenParseErr 4
    If Fractional And Raised Then
        ConsumeNumber = CDbl(Buffer) ^ CLng(Exponent)
    ElseIf Raised Then
        ConsumeNumber = CLng(Buffer) ^ CLng(Exponent)
    ElseIf Fractional Then
        ConsumeNumber = CDbl(Buffer)
    Else
        ConsumeNumber = CLng(Buffer)
    End If
End Function

Private Function ConsumeExponent(Data) As Long
    Dim Buffer As String, Char As String
    Dim FirstSeen As Boolean, Signed As Boolean
    EOD = Len(Data)
    Do
        Char = Mid(Data, Pos, 1)
        Select Case Char
            Case "-"
                If FirstSeen Or Signed Then Exit Do
                Signed = True
                Buffer = Buffer & Char
            Case "+"
                If FirstSeen Or Signed Then Exit Do
                Signed = True
            Case 0 To 9
                If Not FirstSeen Then FirstSeen = True
                Buffer = Buffer & Char
            Case Else
                Exit Do
        End Select
        Incr
    Loop Until Pos > EOD
    If Not FirstSeen Then GenParseErr 5
    ConsumeExponent = CLng(Buffer)
End Function

Private Function ConsumeLiteral(Data) As Variant
    If Mid(Data, Pos, 4) = "true" Then
        Incr 4
        ConsumeLiteral = True
    ElseIf Mid(Data, Pos, 5) = "false" Then
        Incr 5
        ConsumeLiteral = False
    ElseIf Mid(Data, Pos, 4) = "null" Then
        Incr 4
        ConsumeLiteral = Null
    Else
        GenParseErr 6
    End If
End Function

Private Function ConsumeString(Data) As String
    Dim Char As String
    Do
        Char = Mid(Data, Pos, 1)
        Select Case Char
            Case """", ""
                Incr
                Exit Do
            Case "\"
                Incr
                Char = Mid(Data, Pos, 1)
                Select Case Char
                    Case """", "\", "/"
                        ConsumeString = ConsumeString & Char
                    Case "b"
                        ConsumeString = ConsumeString & vbBack
                    Case "f"
                        ConsumeString = ConsumeString & vbFormFeed
                    Case "n"
                        ConsumeString = ConsumeString & vbLf
                    Case "r"
                        ConsumeString = ConsumeString & vbCr
                    Case "t"
                        ConsumeString = ConsumeString & vbTab
                    Case "u"
                        Incr
                        Char = Mid(Data, Pos, 4)
                        On Error GoTo InvalidHex
                        ConsumeString = ConsumeString & ChrW(CLng("&H" & Char))
                        On Error GoTo 0
                        Incr 3
                    Case Else
                        GenParseErr 7
                End Select
            Case ""
                GenParseErr 99
            Case Else
                ConsumeString = ConsumeString & Char
        End Select
        Incr
    Loop Until Pos > EOD
    Exit Function
InvalidHex:
    GenParseErr 8
End Function

Private Function ConsumeArray(Data) As Variant
    Dim Char As String
    Dim ValueNext As Boolean
    Dim Arr As Variant
    Do
        ConsumeWhitespace Data
        Char = Mid(Data, Pos, 1)
        If Char = "" Then GenParseErr 99
        If Not ValueNext Then
            Select Case Char
                Case "]"
                    Incr
                    Exit Do
                Case ","
                    If IsEmpty(Arr) Then GenParseErr 9
                    ValueNext = True
                    Incr
                Case Else
                    If Not IsEmpty(Arr) Then GenParseErr 10
                    ValueNext = True
            End Select
        Else
            Select Case Char
                Case "]"
                    GenParseErr 9
                Case "["
                    Incr
                    Push Arr, ConsumeArray(Data)
                Case "{"
                    Incr
                    Push Arr, ConsumeObject(Data)
                Case """"
                    Incr
                    Push Arr, ConsumeString(Data)
                Case 0 To 9, "-"
                    Push Arr, ConsumeNumber(Data)
                Case "t", "f", "n"
                    Push Arr, ConsumeLiteral(Data)
                Case ","
                    GenParseErr 9
                Case Else
                    GenParseErr 11
            End Select
            ValueNext = False
        End If
    Loop Until Pos > EOD
    If IsEmpty(Arr) Then Arr = Array()
    ConsumeArray = Arr
End Function

Private Sub Push(Arr, Value)
    '' implements the push half of a stack for VBA arrays
    '' with automatic incremental resizing
    Dim Index As Long
    If IsEmpty(Arr) Then
        Index = 1
        ReDim Arr(1 To Index)
    Else
        Index = UBound(Arr) + 1
        ReDim Preserve Arr(1 To Index)
    End If
    If VBA.IsObject(Value) Then
        Set Arr(Index) = Value
    Else
        Arr(Index) = Value
    End If
End Sub

Private Function ConsumeObject(Data) As Scripting.Dictionary
    Dim Char As String
    Dim Expect As ObjExp
    Dim Key As String, Value As Variant
    Set ConsumeObject = New Scripting.Dictionary
    Do
        ConsumeWhitespace Data
        Char = Mid(Data, Pos, 1)
        If Char = "" Then GenParseErr 99
        Select Case Expect
            Case ExpInit
                Select Case Char
                    Case """"
                        Incr
                        Key = ConsumeString(Data)
                        Expect = ExpColon
                    Case "}"
                        Incr
                        Exit Do
                    Case ","
                        GenParseErr 12
                    Case Else
                        GenParseErr 14
                End Select
            Case ExpColon
                If Char <> ":" Then GenParseErr 13
                Incr
                Expect = ExpValue
            Case ExpValue
                Select Case Char
                    Case "}"
                        GenParseErr 12
                    Case "{"
                        Incr
                        Set Value = ConsumeObject(Data)
                    Case "["
                        Incr
                        Value = ConsumeArray(Data)
                    Case """"
                        Incr
                        Value = ConsumeString(Data)
                    Case 0 To 9, "-"
                        Value = ConsumeNumber(Data)
                    Case "t", "f", "n"
                        Value = ConsumeLiteral(Data)
                    Case ","
                        GenParseErr 12
                    Case Else
                        GenParseErr 14
                End Select
                Select Case DuplicateKeys
                    ' how to handle duplicate keys in an object is implementation-dependent;
                    ' RFC 7159 suggests using first, using last, or erroring out, but makes
                    ' no recommendation.  This implementation's default is to use the last
                    ' value for any given key, but behaviour is configurable
                    Case jsonKeyFirst
                        ' do nothing
                    Case jsonKeyLast
                        If VBA.IsObject(Value) Then
                            Set ConsumeObject.Item(Key) = Value
                        Else
                            ConsumeObject.Item(Key) = Value
                        End If
                    Case jsonKeyHalt
                        GenParseErr 15
                End Select
                Expect = ExpComma
            Case ExpComma
                Select Case Char
                    Case ","
                        Incr
                        Expect = ExpKey
                    Case "}"
                        Incr
                        Exit Do
                    Case Else
                        GenParseErr 14
                End Select
            Case ExpKey
                Select Case Char
                    Case """"
                        Incr
                        Key = ConsumeString(Data)
                        Expect = ExpColon
                    Case "}"
                        Stop
                        GenParseErr 12
                    Case Else
                        GenParseErr 14
                End Select
        End Select
    Loop Until Pos > EOD
End Function

Private Sub ConsumeWhitespace(Data)
    Do
        Select Case Mid(Data, Pos, 1)
            Case vbCrLf, vbCr, vbLf, vbTab, " "
                Pos = Pos + 1
            Case ""
                Data = "1221"
                Exit Do
            Case Else
                Exit Do
        End Select
    Loop Until Pos > EOD
End Sub

Private Sub Incr(Optional HowMuch As Long = 1)
    Pos = Pos + HowMuch
End Sub

Private Function MakeStr(Str) As String
    MakeStr = Str
    MakeStr = Replace(MakeStr, "\", "\\")
    MakeStr = Replace(MakeStr, """", "\""")
    MakeStr = Replace(MakeStr, vbFormFeed, "\f")
    MakeStr = Replace(MakeStr, vbNullChar, "\u0000")
    MakeStr = Replace(MakeStr, vbBack, "\b")
    MakeStr = Replace(MakeStr, vbTab, "\t")
    MakeStr = Replace(MakeStr, vbVerticalTab, "\v")
    MakeStr = Replace(MakeStr, vbCrLf, "\r\n")
    MakeStr = Replace(MakeStr, vbCr, "\r")
    MakeStr = Replace(MakeStr, vbLf, "\n")
    MakeStr = """" & MakeStr & """"
End Function

Private Sub GenParseErr(Code, Optional ErrPos)
    Dim T As String
    Select Case Code
        Case 0:  T = "Expecting value"
        Case 1:  T = "Illegal character"
        Case 2:  T = "Extraneous data after value"
        Case 3:  T = "Forbidden leading zero in number"
        Case 4:  T = "Sign without number"
        Case 5:  T = "Missing exponent value in number"
        Case 6:  T = "Illegal literal value"
        Case 7:  T = "Illegal escape sequence in string"
        Case 8:  T = "Malformed Unicode escape sequence in string"
        Case 9:  T = "Expecting value in array"
        Case 10: T = "Expecting separator in array"
        Case 11: T = "Illegal character in array"
        Case 12: T = "Expecting value in object"
        Case 13: T = "Expecting colon in object"
        Case 14: T = "Illegal character in object"
        Case 15: T = "Duplicate key in object"
        Case 99: T = "Unexpected end of input"
    End Select
    If IsMissing(ErrPos) Then ErrPos = Pos
    Err.Raise vbObjectError + Code, "JSONParser", T & " at offset " & ErrPos
End Sub

Private Function GenSerErr(Code As Long, Optional ObjType As String) As String
    Dim T As String
    Dim FailMode As jsonFailMode
    FailMode = SerializeFailMode Mod jsonFailConv
    Select Case Code
        Case 100: T = "Unserializable data type: Date"
        Case 101: T = "Unserializable data type: Error"
        Case 102: T = "Unserializable data type: User Defined Type"
        Case 103: T = "Unserializable data type: Data Access Object"
        Case 104: T = "Unserializable object class: " & ObjType
    End Select
    Select Case FailMode
        Case jsonFailStop
            Err.Raise vbObjectError + Code, "JSON", T
        Case jsonFailNull
            GenSerErr = "null"
    End Select
End Function
