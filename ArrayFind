Public Function ArrayFind(ByVal SearchValue As Variant, ByVal DataArray As Variant) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : ArrayFind
      ' Author    : Adiv Abramson
      ' Date      : 01/17/2022
      ' Purpose   : Returns the index of SearchValue in DataArray, if it exists.
      '           : Both Application.WorksheetFunction.Match() and Application.Match()
      '           : are 1-based. Therefore, subtract 1 from the index before returning, since
      '           : all arrays in the application are 0-based.
      '           :
      '           : Return -1 if SearchValue not found.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 01/17/2022 - Adiv Abramson
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
      Dim vntIndex As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    ArrayFind = -1

30    If Not IsArray(DataArray) Then Exit Function
40    If IsError(SearchValue) Then Exit Function
50    If IsArray(SearchValue) Then Exit Function

60    vntIndex = Application.Match(SearchValue, DataArray, 0)
70    If Not IsError(vntIndex) Then ArrayFind = vntIndex - 1



80    Exit Function

ErrProc:
90       ArrayFind = -1
100      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure ArrayFind of Module " & MODULE_NAME
End Function

