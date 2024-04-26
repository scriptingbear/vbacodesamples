Public Function IsTable(ByRef DataSheet As Worksheet, ByVal TableName As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsTable
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Determine if specified table (list object) exists on specified worksheet.
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
      Dim objList As ListObject
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

10    IsTable = False
20    On Error Resume Next

30    Set objList = DataSheet.ListObjects(TableName)

40    If Err <> 0 Then
50       Err.Clear
60       Exit Function
70    End If

80    On Error GoTo ErrProc

90    IsTable = True


100   Exit Function

ErrProc:
110      IsTable = False
120      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure IsTable of Module " & MODULE_NAME
End Function
