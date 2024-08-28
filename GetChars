Public Function GetChars(ByVal Text As String, ByVal CharType As GetCharsType) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GetChars
      ' Author    : Adiv Abramson
      ' Date      : 04/02/2018
      ' Purpose   : Return alpha or numerics from Text
      '           : Changed Pattern for Numerics so that it returns floating point values,
      '           : not just integers; also allow for negative numbers
      '           : Corrected fatal flaw: when no match is found, there is nothing to replace with ""
      '           : So return original string, don't return a zero length string.
      '           : Make function exit immediately if Text is ""
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 04/02/2018 - Adiv Abramson
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


20    If Trim(Text) = "" Then
30       GetChars = ""
40       Exit Function
50    End If

60    With objRegex
70       .Global = True
80       .IgnoreCase = True
90       If CharType = Alpha Then
100         .PATTERN = "[^a-z]"
110      Else
120         .PATTERN = "[^0-9.-]"
130      End If
140      If .Test(Text) Then
150         strResult = .Replace(Text, "")
160         GetChars = strResult
170      Else
180         GetChars = Text
190      End If
200   End With


210   Exit Function

ErrProc:

220       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                 & CStr(Erl) & " in procedure GetChars of Module mdlStringFunctions"
End Function

==========================
'Place in module Declarations section
'of a separate module
Public Enum GetCharsType
  Alpha = 0
  Numeric = 1
End Enum
