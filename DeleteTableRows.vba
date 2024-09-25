Public Sub DeleteTableRows(ByVal TableName As String, ParamArray ColumnsToClear() As Variant)
'---------------------------------------------------------------------------------------
' Procedure : DeleteTableRows
' Author    : Adiv Abramson
' Date      : 4/22/2024
' Purpose   : Clear all rows and columns from the table unless otherwise specified.
'           : Optional parameter is a ParamArray of <col_name1>, <keep_first_cell>,
'           : <col_name2>, <keep_first_cell>,...
'           : This allows for granular control of which columns are completely cleared
'           : of data and which ones retain the contents of the first cell.
'           :
'           : The resulting table will always consist of the header and the first row
'           : (1.1) Renamed and moved from mdlTest::TestDeleteTableRows()
'           : to mdlUtilTableFunctions::DeleteTableRows
'           :
'           :
'           :
'           :
' Versions  : 1.0 - 4/22/2024 - Adiv Abramson
'           : 1.1 - 04/24/2024 - Adiv Abramson
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
Dim lngDataRows As Long
Dim i As Integer
Dim intColumns As Integer
'*********************************

'Worksheets:
'*********************************

'*********************************

'Workbooks:
'*********************************

'*********************************

'Ranges:
'*********************************
Dim rngFirstListRow As Range
Dim rngHeaders As Range
Dim rngResizedTable As Range
Dim rngColumnToClear As Range
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
Dim vntColumnName As Variant
Dim vntKeepFirstCell As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************

10    On Error GoTo ErrProc
      
      '========================================================
      'Validate inputs.
      '========================================================
20    Set objList = GetTable(TableName:=TableName, DataWorksheet:=WorksheetNames.CodeHelper)
30    If objList Is Nothing Then Exit Sub
40    If objList.DataBodyRange Is Nothing Then Exit Sub
      
50    If Not IsMissing(ColumnsToClear) Then
60       intColumns = UBound(ColumnsToClear)
         '========================================================
         'Every column must be followed by a Boolean value. Since
         'an array index starts at 0, if the upper bound is evenly
         'divisible by 2, it has an odd number of items and thus
         'one or more columns doesn't have a paired Boolean value.
         '========================================================
130      If UBound(ColumnsToClear) Mod 2 = 0 Then Exit Sub
         
         '========================================================
         'Ensure each pair of values in the ParamArray is valid,
         'i.e. the first item of each pair must be a column in the
         'table and the second item of each pair must be a Boolean
         '========================================================
200      For i = 0 To intColumns - 1 Step 2
210         vntColumnName = ColumnsToClear(i)
220         vntKeepFirstCell = ColumnsToClear(i + 1)
230         If Not IsTableField(FieldName:=vntColumnName, _
                                DataTableName:=TableName, _
                                DataWorksheetName:=WorksheetNames.CodeHelper) Then Exit Sub
           
270         If TypeName(vntKeepFirstCell) <> "Boolean" Then Exit Sub
280      Next i
290   End If 'Not IsMissing(ColumnsToClear)
      
300   With objList
        
320      lngDataRows = .DataBodyRange.Rows.Count
         
340      If IsMissing(ColumnsToClear) Then
350         .DataBodyRange.ClearContents
360      Else
            '========================================================
            'Get a reference to each of the specified columns
            'and clear contents of all cells unless otherwise
            'specified.
            '========================================================
420         For i = 0 To intColumns - 1 Step 2
430            vntColumnName = ColumnsToClear(i)
440            vntKeepFirstCell = ColumnsToClear(i + 1)
               
460            Set rngColumnToClear = .ListColumns(vntColumnName).DataBodyRange
470            If vntKeepFirstCell Then
480               rngColumnToClear.Offset(RowOffset:=1).Resize(RowSize:=lngDataRows - 1).ClearContents
490            Else
500               rngColumnToClear.ClearContents
510            End If 'vntKeepFirstCell
520         Next i
530      End If 'IsMissing(ColumnsToClear)
      
         '========================================================
         'Resize table to include the header row and the first
         'row of the DataBodyRange.
         '========================================================
580      Set rngFirstListRow = .ListRows(1).Range
590      Set rngHeaders = .HeaderRowRange
600      Set rngResizedTable = Application.Union(rngHeaders, rngFirstListRow)
       
620   End With
      
630   objList.Resize rngResizedTable



Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure DeleteTableRows of Module " & MODULE_NAME
End Sub

