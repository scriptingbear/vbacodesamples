'---------------------------------------------------------------------------------------
' Module    : mdlStringFunctions
' Author    : Adiv Abramson
' Date      : 
' Purpose   : Contains various convenience and wrapper functions for working with strings
'           :
'           :
'           :
'           :
'           :
' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Option Explicit
Option Private Module


Private Const MODULE_NAME = "mdlStringFunctions"


Public Function CleanString(ByVal InputText As String) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : CleanString
      ' Author    : Adiv Abramson
      ' Date      : 
      ' Purpose   : Remove blanks and all non-alphanumeric characters from input string.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************

      '*********************************

      'Numerics:
      '*********************************

      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objRegx As RegExp
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    CleanString = ""

      '========================================================
      'Validate input
      '========================================================
30    If Len(InputText) = 0 Then Exit Function

40    Set objRegx = New RegExp
50    With objRegx
60       .Global = True
70       .IgnoreCase = True
80       .Pattern = "\W"
90       CleanString = .Replace(InputText, "")
100   End With



110   Exit Function

ErrProc:
120      CleanString = ""
130      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure CleanString of Module " & MODULE_NAME
End Function

Public Function SplitX(ByVal InputText As String, ByVal DELIMITER As String) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : SplitX
      ' Author    : Adiv Abramson
      ' Date      : 
      ' Purpose   : Split a string using Regex.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strChar As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intSize As Integer
      Dim intMatch As Integer
      Dim intMatches As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************
      Dim arOutput As Variant
      Dim arSpecialChars As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      Dim objMatch As Match
      Dim objMatches As MatchCollection
      Dim objSubmatches As SubMatches
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************
      Dim blnSplitAll As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    SplitX = Null

30    If Trim(InputText) = "" Then Exit Function
40    If InStr(1, InputText, DELIMITER) = 0 Then
         'Return input as a single element array,
         'indicating that it could not be split on
         'the provided delimiter
50       SplitX = Array(InputText)
60       Exit Function
70    End If 'InStr(1, InputText, Delimiter) = 0

      'A ZLS delimiter means split input into array of individual characters
80    If DELIMITER = "" Then blnSplitAll = True

      'Must escape special characters used as delimiters
90    arSpecialChars = Array("\", ".", "-", "?", "|")
100   intSize = UBound(arSpecialChars)
110   For i = 0 To intSize
120      If DELIMITER = arSpecialChars(i) Then
130         DELIMITER = "\" & DELIMITER
140         Exit For
150      End If
160   Next i

170   Set objRegex = New RegExp
180   With objRegex
190      .Global = True
200      .IgnoreCase = True
210      If Not blnSplitAll Then
220         .Pattern = "[^" & DELIMITER & "]+"
230      Else
240         .Pattern = "."
250      End If 'Not blnSplitAll
         
260      If Not .Test(InputText) Then Exit Function
270      Set objMatches = .Execute(InputText)
280   End With


290   For Each objMatch In objMatches
300      AddToArray arOutput, objMatch.Value, False
310   Next objMatch

320   SplitX = arOutput

330   Exit Function

ErrProc:
340      SplitX = Null
350      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SplitX of Module mdlStringFunctions"
End Function



Public Function EmbedQuotes(ByVal text As String) As String
'---------------------------------------------------------------------------------------
' Procedure : EmbedQuotes
' Author    : Adiv Abramson
' Date      : 
' Purpose   : Replace backtick (`) in string with double quote (")
'           : Useful for string building in which embedded double quotes
'           : may occur
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'  Versions : 1. - 08/31/2019 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------



On Error GoTo ErrProc

EmbedQuotes = ""

If Len(text) = 0 Then Exit Function

If CharCount(text, BACKTICK) = 0 Then EmbedQuotes = text: Exit Function
  

EmbedQuotes = Replace(text, BACKTICK, QUOTE)



Exit Function

ErrProc:
    EmbedQuotes = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure EmbedQuotes of Module mdlStringFunctions"
End Function
Public Function GetChars(ByVal strText As String, ByVal intType As GetCharsType) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GetChars
      ' Author    : Adiv
      ' Date      : 
      ' Purpose   : Return alpha or numerics from strText
      '           : Changed Pattern for Numerics so that it returns floating point values,
      '           : not just integers; also allow for negative numbers
      '           : Corrected fatal flaw: when no match is found, there is nothing to replace with ""
      '           : So return original string, don't return a zero length string!
      '           : Make function exit immediately if strText is ""
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      Dim strResult As String
      Dim objRegex As New RegExp



10    On Error GoTo ErrProc


20    If Trim(strText) = "" Then
30       GetChars = ""
40       Exit Function
50    End If

60    With objRegex
70       .Global = True
80       .IgnoreCase = True
90       If intType = Alpha Then
100         .Pattern = "[^a-z]"
110      Else
120         .Pattern = "[^0-9.-]"
130      End If
140      If .Test(strText) Then
150         strResult = .Replace(strText, "")
160         GetChars = strResult
170      Else
180         GetChars = strText
190      End If
200   End With


210   Exit Function

ErrProc:

220       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure GetChars of Module mdlStringFunctions"
End Function

Public Function LTRUNC(ByVal strText As String, ByVal lngNumChars As Long) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : LTRUNC
      ' Author    : Adiv
      ' Date      : 
      ' Purpose   : Removes specified # of characters from left end of string
      '           :
      '---------------------------------------------------------------------------------------

10    On Error GoTo ErrProc

20    If lngNumChars <= 0 Then
30       LTRUNC = ""
40       Exit Function
50    End If

60    If Len(strText) = 0 Then
70       LTRUNC = ""
80       Exit Function
90    End If

100   If lngNumChars >= Len(strText) Then
110       LTRUNC = ""
120       Exit Function
130   End If

140   LTRUNC = Mid(strText, lngNumChars + 1)

150   Exit Function


ErrProc:

160       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure LTRUNC of Module mdlStringFunctions"

End Function

Public Function RTRUNC(ByVal strText As String, ByVal lngNumChars As Long) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : RTRUNC
      ' Author    : Adiv
      ' Date      : 
      ' Purpose   : Removes specified # of characters from right end of string
      '           :
      '           :
      '---------------------------------------------------------------------------------------


10    On Error GoTo ErrProc


20    If lngNumChars <= 0 Then
30       RTRUNC = ""
40       Exit Function
50    End If

60    If Len(strText) = 0 Then
70       RTRUNC = ""
80       Exit Function
90    End If

100   If lngNumChars >= Len(strText) Then
110       RTRUNC = ""
120       Exit Function
130   End If

140   RTRUNC = Mid(strText, 1, Len(strText) - lngNumChars)

150   Exit Function


ErrProc:

160       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure RTRUNC of Module mdlStringFunctions"

End Function

Public Function CharCount(ByVal text As String, ByVal Char As String, Optional IgnoreCase As Boolean = True) As Long
'---------------------------------------------------------------------------------------
' Procedure : CharCount
' Author    : Adiv Abramson
' Date      : 
' Purpose   : Counts the number of occurrences of Char in Text
'           : (1.1) Now supporting search string of more than one character
'           :
'           :
'           :
'           :
'           :
'Versions   : 1.0 - 00/00/0000 - Adiv Abramson
'           : 1.0 - 00/00/0000 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

'Strings:
'*********************************

'*********************************

'Numerics:
'*********************************
Dim lngTextLength As Long
Dim lngReplacedLength As Long
Dim lngSearchTextLength As Long
Dim lngCompareOption As Long
'*********************************

'Worksheets:
'*********************************

'*********************************

'Workbooks:
'*********************************

'*********************************

'Ranges:
'*********************************

'*********************************

'Arrays:
'*********************************

'*********************************

'Objects:
'*********************************

'*********************************

'Variants
'*********************************

'*********************************

'Booleans
'*********************************

'*********************************

'Constants
'*********************************

'*********************************


10    On Error GoTo ErrProc
      
20    CharCount = 0
      
30    lngTextLength = Len(text)
40    lngSearchTextLength = Len(Char)
      
      '==============================================
      'Validate inputs
      '==============================================
80    If lngTextLength = 0 Or lngSearchTextLength = 0 Then Exit Function
90    If lngSearchTextLength > lngTextLength Then Exit Function
      
100   lngCompareOption = IIf(IgnoreCase, VbCompareMethod.vbTextCompare, VbCompareMethod.vbBinaryCompare)
      
110   lngReplacedLength = Len(Replace(Expression:=text, Replace:="", Find:=Char, Compare:=lngCompareOption))
      
120   CharCount = (lngTextLength - lngReplacedLength) / lngSearchTextLength



   
Exit Function

ErrProc:
    CharCount = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure CharCount of Module mdlStringFunctions"


End Function

Public Function PrintFormat(ByVal text As String, ParamArray Values() As Variant) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : PrintFormat
      ' Author    : Adiv Abramson
      ' Date      : 
      ' Purpose   : Emulate string interpolation. Replace place markers "{$n} with corresponding
      '           : items in Values array, where n is the index
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strPlaceHolder As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intSize As Integer
      Dim intIndex As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim vntValue As Variant
      Dim vntPlaceHolder As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

      '========================================================
      'Return original string if validations fail.
      '========================================================
20    PrintFormat = text

      '========================================================
      'Validate inputs
      '========================================================
30    If Trim(text) = "" Then Exit Function
40    If Not IsArray(Values) Then Exit Function

      '========================================================
      'Iterate over Values and make substitutions in input text.
      '========================================================
50    intSize = UBound(Values)
60    For i = 0 To intSize
70       vntValue = Values(i)
80       If Not IsNull(vntValue) Then
90          vntValue = CStr(vntValue)
100         strPlaceHolder = "{$" & i & "}"
110         text = Replace(Expression:=text, Find:=strPlaceHolder, Replace:=vntValue)
120      End If
130   Next i

140   PrintFormat = text

150   Exit Function

ErrProc:
160      PrintFormat = ""
170      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of  procedure PrintFormat of Module " & MODULE_NAME
End Function

Public Function SmartCapitalize(ByVal text As String) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : SmartCapitalize
      ' Author    : Adiv Abramson
      ' Date      : 
      ' Purpose   : Convert Text to Proper Case but skip over abbreviations in ALL CAPS.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strNew As String
      '*********************************

      'Numerics:
      '*********************************

      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      Dim objMatches As MatchCollection
      Dim objMatch As Match
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    SmartCapitalize = ""

30    strNew = StrConv(text, vbProperCase)

40    Set objRegex = New RegExp
50    With objRegex
60       .Global = True
70       .IgnoreCase = False
80       .Pattern = "[A-Z]{2,}"
90       If .Test(text) Then
100         Set objMatches = .Execute(text)
110         For Each objMatch In objMatches
120            Mid(strNew, objMatch.FirstIndex + 1) = UCase(objMatch)
130         Next objMatch
140      End If '.Test(Text)
150   End With

160   SmartCapitalize = strNew

170   Exit Function

ErrProc:
180      SmartCapitalize = ""
190      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure SmartCapitalize of Module " & MODULE_NAME
End Function

Public Function PrintFormat2(ByVal text As String, ParamArray Values() As Variant) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : PrintFormat2
      ' Author    : Adiv Abramson
      ' Date      : 
      ' Purpose   : Alternative to PrintFormat() that uses text labels before
      '           : each value of the ParamArray for substitions.
      '           : (1.1) Values array must have an even number of elements: one tag
      '           : for each element.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strPlaceHolder As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intSize As Integer
      Dim intIndex As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim vntValue As Variant
      Dim vntPlaceHolder As Variant
      Dim vntTextLabel As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

      '========================================================
      'Return original string if validations fail.
      '========================================================
20    PrintFormat2 = text

      '========================================================
      'Validate inputs
      '========================================================
30    If Trim(text) = "" Then Exit Function
40    If Not IsArray(Values) Then Exit Function
50    If UBound(Values) Mod 2 = 0 Then Exit Function

      '========================================================
      'Iterate over Values and make substitutions in input text.
      '========================================================
60    intSize = UBound(Values)
70    For i = 0 To intSize Step 2
80       vntTextLabel = Values(i)
90       vntValue = Values(i + 1)
100      If Not IsNull(vntValue) Then
110         vntValue = CStr(vntValue)
120         strPlaceHolder = "{$" & vntTextLabel & "}"
130         text = Replace(Expression:=text, Find:=strPlaceHolder, Replace:=vntValue)
140      End If
150   Next i

160   PrintFormat2 = text


170   Exit Function

ErrProc:
180      PrintFormat2 = ""
190      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure PrintFormat2 of Module " & MODULE_NAME
End Function
