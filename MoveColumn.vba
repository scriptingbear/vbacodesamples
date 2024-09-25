Public Sub MoveColumn(ByRef DataWorksheet As Worksheet, ByVal SourceHeader As String, _
     ByVal TargetHeader As String, Optional MoveAfter As Boolean = False)
      '---------------------------------------------------------------------------------------
      ' Procedure : MoveColumn
      ' Author    : Adiv
      ' Date      : 08/27/2018
      ' Purpose   : Find column with SourceHeader in row 1. Find column with TargetHeader in row 1.
      '           : Insert blank column just before TargetHeader column. Cut SourceHeader column
      '           : and paste into blank column. Delete SourceHeader column
      '           : Assume worksheet is unlocked and autofilter has been removed
      '           : If optional MoveAfter parameter is True, the blank column should be
      '           : inserted AFTER the target column, on the right, instead of on the left
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
      ' Versions  : 1.0 - 08/27/2018 - Adiv Abramson
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
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strSource_Col As String
      Dim strTarget_Col As String
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
      Dim rngTargetCol As Range
      Dim rngSourceCol As Range
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

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

      'Validate inputs
20    If Trim(SourceHeader) = "" Or Trim(TargetHeader) = "" Then
30       Exit Sub
40    End If

50    strSource_Col = GetHeaderInfo(DataWorksheet, SourceHeader, HeaderInfoColumn)
60    strTarget_Col = GetHeaderInfo(DataWorksheet, TargetHeader, HeaderInfoColumn)

70    If strSource_Col = "" Or strTarget_Col = "" Then
80       Exit Sub
90    ElseIf StringCompare(strSource_Col, strTarget_Col) Then
100      Exit Sub
110   End If

      'Ensure that source column isn't already just before (or after) target column
120   If Not MoveAfter Then
130      With DataWorksheet
140         If .Range(strSource_Col & 1).Column = .Range(strTarget_Col & 1).Column - 1 Then
150            msgAttention "Column " & SourceHeader & " is already just before column " & TargetHeader & "."
160            Exit Sub
170         End If
180      End With
190   Else
200      With DataWorksheet
210         If .Range(strSource_Col & 1).Column = .Range(strTarget_Col & 1).Column + 1 Then
220            msgAttention "Column " & SourceHeader & " is already just after column " & TargetHeader & "."
230            Exit Sub
240         End If
250      End With
260   End If 'Not MoveAfter

      'Insert blank column to the left of the target column
270   If Not MoveAfter Then
         'Insert immediately to the left of target column
280      DataWorksheet.Range(strTarget_Col & 1).EntireColumn.Insert
290   Else
         'Insert immediately to the right of the target column
         'Useful when moving a column to the end of the existing set of columns
300      DataWorksheet.Range(strTarget_Col & 1).Offset(ColumnOffset:=1).EntireColumn.Insert
310   End If

      'Must reevaluate source column letter since new blank column has been introduced
320   strSource_Col = GetHeaderInfo(DataWorksheet, SourceHeader, HeaderInfoColumn)
330   Set rngSourceCol = DataWorksheet.Range(strSource_Col & 1).EntireColumn

      'Actual target column is the new blank column, which is to the left of the specified target column
      'by default, unless MoveAfter is True
      'Must also reevaluate target column letter
340   strTarget_Col = GetHeaderInfo(DataWorksheet, TargetHeader, HeaderInfoColumn)
350   If Not MoveAfter Then
360      Set rngTargetCol = DataWorksheet.Range(strTarget_Col & 1).Offset(ColumnOffset:=-1)
370   Else
380      Set rngTargetCol = DataWorksheet.Range(strTarget_Col & 1).Offset(ColumnOffset:=1)
390   End If 'Not MoveAfter

      'Now copy data from source to target
400   rngSourceCol.Copy Destination:=rngTargetCol
410   Application.CutCopyMode = False

      'Delete source column
420   rngSourceCol.EntireColumn.Delete



430   Exit Sub

ErrProc:
        
440      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure MoveColumn of Module mdlUtil"
End Sub



