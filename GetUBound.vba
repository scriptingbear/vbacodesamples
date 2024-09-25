Public Function GetUBound(ByVal DataArray As Variant) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : GetUBound
      ' Author    : Adiv Abramson
      ' Date      : 09/23/2024
      ' Purpose   : Use if there's a chance that the array will not have any elements.
      '           : Normally, calling UBound() on an empty array throws an error. This
      '           : wrapper function returns a -1 if the array has no elements.
      '           : Invoking Ubound() on an empty fixed size array returns -1 without errors.
      '           : E.g. Ubound(Array()) => -1 and GetUbound(Array()) => -1.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 09/23/2024 - Adiv Abramson
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
      Dim lngUbound As Long
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

20    GetUBound = -1

      '========================================================
      'Validate input
      '========================================================
30    If Not IsArray(DataArray) Then Exit Function

40    On Error Resume Next
50    lngUbound = UBound(DataArray)
60    If Err <> 0 Then
70       Err.Clear
80       Exit Function
90    End If

100   On Error GoTo ErrProc

110   GetUBound = lngUbound

120   Exit Function

ErrProc:
130      GetUBound = -1
140      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                & Erl & " in procedure GetUBound of Module " & MODULE_NAME
End Function

