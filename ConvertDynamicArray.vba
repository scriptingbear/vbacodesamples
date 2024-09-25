Public Function ConvertDynamicArray(ByVal DynamicArray As Range) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : ConvertDynamicArray
      ' Author    : Adiv Abramson
      ' Date      : 5/7/2024
      ' Purpose   : Treating a spilled range as a regular range for populating
      '           : arrays doesn't seem to be working. This function iterates
      '           : through the cells of the spilled range and adds each value
      '           : to a new 1D array that VBA can work with.
      '           : For example Join(<dynamic array>, ",") doesn't work.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 5/7/2024 - Adiv Abramson
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
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arValues As Variant
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

20    ConvertDynamicArray = Null

30    If DynamicArray Is Nothing Then Exit Function

40    For Each rngCell In DynamicArray.Cells
50       AddToArray DataArray:=arValues, InputValue:=rngCell.Value, KeepUnique:=False
60    Next rngCell

70    ConvertDynamicArray = arValues


80    Exit Function

ErrProc:
90       ConvertDynamicArray = Null
100      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure ConvertDynamicArray of Module " & MODULE_NAME
End Function
