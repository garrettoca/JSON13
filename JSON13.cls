VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************
'** This is a single class represents JSON objects and implements a JSON parser
'** Author: Garrett O'Carroll, Docland Software www.docland.com
'** Version: 1 April 2017
'** (c) Docland Software 2017
'********************************************************************************
'********************************************************************************
'parser properties
Dim parser_src As String
Dim parser_index As Integer
Dim parser_token As String

' data item properties
' every leaf value (string, number), array and object is wrapped in an instance of JSON
' every element array or object property is also a JSON object
Public value As Variant
Public name As String
Public keys As Variant ' will contain a collection of strings 'keys' for properties
Public JTYPE As String

' helper method to safely assign the data value to a variant
Sub assign(target As Variant)
    Select Case VarType(value)
    Case vbObject: Set target = value
    Case Else: target = value
    End Select
End Sub
 
' get a named value form collection in the .value property
Public Function Item(key As String) As Variant
    Set Item = value.Item(key)
End Function

' same as .Item
Public Function getObj(key As Variant) As Variant
    Set getObj = value.Item(key)
End Function

' get a named string value form collection in the .value property
Public Function getString(key As Variant) As String
    getString = value.Item(key).value
End Function

' get a named variant value form collection in the .value property
Public Function getVar(key As Variant) As Variant
    getVar = value.Item(key).value
End Function

' render object as a string
Public Function toString()
    If (Me.isObject()) Then
        toString = Me.obj2String
    ElseIf (isArray()) Then
        toString = Me.arr2String
    ElseIf VarType(Me.value) = vbString Then
        toString = """" & encode(Me.value) & """ "
    ElseIf VarType(Me.value) = vbDouble Then
        toString = Me.value
    End If
End Function

' render object as a string
Public Function obj2String()
    Dim str As String
    str = " {"
    Dim jObj As JSON13
    Dim k As Variant
    Dim ix As Integer
    ix = 0
    
    For Each k In Me.keys
       Set jObj = Me.value.Item(k)
       If ix > 0 Then str = str + ", "
       str = str & " """ & Me.name & """:" & Me.value.Item(k).toString()
       ix = ix + 1
    Next
    obj2String = str & "} "
End Function

' render array as a string
Public Function arr2String()
    Dim str As String
    str = " ["
   
    Dim arr As Variant
    Set arr = Me.value
    Dim ix As Integer
    For ix = 1 To arr.count
      'each k will also be a JSON13 object
      If ix > 1 Then str = str + ", "
      str = str & arr(ix).toString()
    Next ix
    arr2String = str & "] "
End Function

' test if the JSON value is an object
Public Function isObject()
    isObject = Not IsEmpty(Me.keys) And (VarType(Me.value) = vbObject)
End Function

' test if the JSON value is an array
Public Function isArray()
    isArray = IsEmpty(Me.keys) And (VarType(Me.value) = vbObject)
End Function

' dump the object to the debug console
Sub dump()
    If (Me.isObject()) Then
        Call Me.dumpObject
    ElseIf (isArray()) Then
        Call Me.dumpArray
    Else
        Debug.Print "JSON13.dump()=" & Me.value
    End If
End Sub

' dump the object to the debug console
Sub dumpObject()
    Dim jObj As JSON13
    Dim k As Variant
    For Each k In Me.keys
       Debug.Print "JSON13.dumpObj(): " & k & ".."
       Set jObj = Me.value.Item(k)
       Call jObj.dump
    Next
End Sub

' dump array to the debug console
Sub dumpArray()
    Dim arr As Variant
    Set arr = Me.value
    Dim ix As Integer
    For ix = 1 To arr.count
       'each k will also be a JSON13 object
       Debug.Print "JSON13.dumpArray(): " & ix & ".."
       Call arr(ix).dump
    Next ix
End Sub

'*************************************************************
'* entry point for the parser
'*************************************************************
Public Function parse(ByRef src As String) As JSON13
    parser_src = src
    parser_index = 1
    Set parse = parseJSONObject()
    parse.name = "root"
End Function

Private Function getToken(Optional esc As Boolean = False) As String
    Dim tNext As String
    parser_token = Mid(parser_src, parser_index, 1)
    
    If (esc) And (parser_token = "\") Then
        tNext = Mid(parser_src, parser_index + 1, 1)
        ' handle escape codes
        If (InStr("\""bfnrt", tNext) > 0) Then
            parser_token = Mid(parser_src, parser_index, 2)
        ElseIf tNext = "u" And hexDigits(parser_index + 2, 4) Then
            parser_token = Mid(parser_src, parser_index, 6)
        End If
        getToken = parser_token
    Else
        getToken = parser_token
    End If
    
End Function

Private Function hexDigits(start, count) As Boolean
    Dim i As Integer
    For i = 1 To 4
       If (InStr("0123456789ABCDEFabcdef", Mid(parser_src, start + i, 1)) = 0) Then
           hexDigits = False
           Exit Function
       End If
    Next i
    hexDigits = True
End Function

Private Function decode(tok As String) As String
    ' map the esc sequence to a output
    Select Case tok:
    Case "\\":
        decode = "\"
    Case "\""":
        decode = """"
    Case "\b":
        decode = Chr(8)
    Case "\f":
        decode = Chr(10)
    Case "\n":
        decode = Chr(10)
    Case "\r":
        decode = Chr(13)
    Case "\t":
        decode = Chr(9)
    Case Else
        ' if (tok starts with \u its a unicode character)
        If (Len(tok) = 1) Then
            decode = tok
        ElseIf (Len(tok) = 6) Then
            ' convert  the hex  string to a number & use chrW
            decode = ChrW(CInt("&H" & Right(tok, 4)))
        Else
            ' error
             MsgBox ("JSON Syntax Error: unrecognised character escape sequence " & tok & " @" & parser_index)
             End
        End If
    End Select
End Function

' encode converts tabs, line feed and UNICODE characters to valid
Function encode(str As Variant) As Variant
    encode = Replace(str, """", "\""")
End Function

' move current location forward
Private Sub advance()
    Dim l As Integer
    l = Len(parser_token)
    If (l = 0) Then l = 1
    parser_index = parser_index + l
End Sub

' parse a property in the form of "key" : "value" or "key":9999
Private Function parseJSONProp() As JSON13
    Dim prop  As JSON13
    Dim propName
    propName = parseJSONKey()
    
    Call skipToPropValue
    Dim cStart  As String
    
    ' read a string, object, array or a number
    cStart = getToken()
    
    If (isJSONObjectValue(cStart)) Then
        Set parseJSONProp = parseJSONObject()
        
    ElseIf (isJSONArrayValue(cStart)) Then
        Set prop = parseJSONArray()
        Set parseJSONProp = prop
    Else
        Set prop = New JSON13
        prop.JTYPE = "VALUE"
        prop.value = parseJSONVariantValue(cStart)
        Set parseJSONProp = prop
    End If
    parseJSONProp.name = propName
End Function

' parse a property name
Private Function parseJSONKey()
   ' skip the opening quote
    Dim cStart, pKey As String
    cStart = getToken()
            
    ' read the step forward
    Call advance
    
    pKey = ""
    While (getToken() <> cStart)
       pKey = pKey + parser_token
       Call advance
    Wend
    Call advance
    parseJSONKey = pKey
End Function

' skip colon & spaces
Private Sub skipToPropValue()
    ' skip spaces
    Call skipWhiteSpace
    
    ' skip ':' and more spaces
    If (getToken() <> ":") Then
       ' this is an error
       MsgBox ("JSON Syntax Error: Colon expected at " & parser_index & " (found " & parser_token & " )")
       End
    End If
    
    Call advance
    Call skipWhiteSpace
End Sub

' skip white space
Private Sub skipWhiteSpace()
    While (getToken() = " ")
       Call advance
    Wend
End Sub

' test next value type object {}
Private Function isJSONObjectValue(cStart) As Boolean
  isJSONObjectValue = (cStart = "{")
End Function

' test next value type array []
Private Function isJSONArrayValue(cStart) As Boolean
  isJSONArrayValue = (cStart = "[")
End Function

' test next value is a string - "" delimited
Private Function isJSONStringValue(cStart) As Boolean
  isJSONStringValue = (cStart = """") Or (cStart = "'")
End Function

' test next value is a number
Private Function isJSONNumericValue(cStart) As Boolean
  isJSONNumericValue = (InStr("0123456790", cStart) > 0)
End Function

' parse a variant value (string or number
Private Function parseJSONVariantValue(cStart) As Variant
    Dim pVal As String
    Call skipWhiteSpace
    If (cStart = """") Or (cStart = "'") Then
        ' value is a string
        Call advance
        While (getToken(True) <> cStart)
           pVal = pVal + decode(parser_token)
           Call advance
        Wend
        parseJSONVariantValue = pVal
        Call advance
    ElseIf (InStr("0123456790", cStart) > 0) Then
        ' no delimiter, no comma triad separator accepted
        While (InStr("0123456790.", getToken()) > 0)
           pVal = pVal + parser_token
           Call advance
        Wend
        parseJSONVariantValue = CDbl(Val(pVal))
    Else
        MsgBox ("parseJSONVariantValue: Syntax Error: Value expected at " & parser_index & " (found '" & cStart & "')")
        End
    End If
End Function

' parse an array - [] delimited
Private Function parseJSONArray() As JSON13
    
    Dim coll As collection
    Set coll = New collection
    
    Dim cStart As String
    
    ' parser_index points to '['
    Call advance
    
    Call skipWhiteSpace
    cStart = getToken()
    
    Dim count As Integer
    count = 0
    
    Dim staging As Variant
    
    ' parse a list of values separated by commas, accept a trailing comma
    While cStart <> "]"
        
        If (isJSONObjectValue(cStart)) Then
            Set staging = parseJSONObject()
            coll.add staging
        ElseIf (isJSONArrayValue(cStart)) Then
            Set staging = parseJSONArray()
            coll.add staging
        Else
            Dim jObj As JSON13
            Set jObj = New JSON13
            
            Dim iVal As Variant
            iVal = parseJSONVariantValue(cStart)
            jObj.value = iVal
            coll.add jObj
        End If
        
        Call skipWhiteSpace
        
        If (getToken() = ",") Then
            Call advance
            Call skipWhiteSpace
        End If
        cStart = getToken()
    Wend
    
    ' char should be a ']'
    If (cStart <> "]") Then
       ' this is an error
       MsgBox ("parseJSONArray: JSON Syntax Error: Array closing brackey ] expected at " & parser_index)
       End
    End If
    
    Call advance
    Dim rObj As JSON13
    Set rObj = New JSON13
    Set rObj.value = coll
    rObj.JTYPE = "ARRAY"
    Set parseJSONArray = rObj
    
End Function

' parse an object - {} delimited
Private Function parseJSONObject() As JSON13
    Dim obj As JSON13
    Set obj = New JSON13
    
    Dim dict As collection
    Set dict = New collection
    
    Dim prop As JSON13
    Dim allKeys As collection
    Set allKeys = New collection
    
    Set obj.keys = allKeys
  
    ' parser_index points to '{'
    Call advance
    Call skipWhiteSpace
    
    ' parse a list of values separated by commas, accept a trailing comma
    While getToken() <> "}"
        Set prop = parseJSONProp()
        
        ' Add the named object to the collection.
        dict.add prop, prop.name
        allKeys.add (prop.name)
        
        Call skipWhiteSpace
        If (getToken() = ",") Then
            Call advance
            Call skipWhiteSpace
        End If
    Wend
    
    ' char should be a '}'
    If (getToken() <> "}") Then
       ' this is an error
       MsgBox ("JSON Syntax Error: Object closing brace } expected at " & parser_index)
       End
    End If
    
    ' consume the closing token
    Call advance
    Set obj.value = dict
    obj.JTYPE = "OBJECT"
    Set parseJSONObject = obj
End Function
