Public Function XLNz(ByVal InputValue As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : XLNz
      ' Author    : Adiv Abramson
      ' Date      : 09/19/2024
      ' Purpose   : Mimics the Nz() function in MS Access. If the argument is Null,
      '           : return 0 by default or ValueIfNull, if specified
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
      ' Versions  : 1.0 - 09/19/2024 - Adiv Abramson
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

20    XLNz = 0

30    If Not IsNull(InputValue) Then
40       XLNz = InputValue
50    Else
60       XLNz = ValueIfNull
70    End If


80    Exit Function

ErrProc:
90       XLNz = 0
100      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                & Erl & " in procedure XLNz of Module " & MODULE_NAME
End Function
