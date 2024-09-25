'---------------------------------------------------------------------------------------
' Module    : mdlUtil
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Contains general purpose utility functions and procedures
'           : (1.1) Added code found online to clear Windows clipboard
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
'           :
'  Versions : 1.0 - 00/00/0000 - Adiv Abramson
'           : 
'           : 
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Option Explicit
Option Private Module


Private Const MODULE_NAME = "mdlUtil"


Public Function GetNameValue(ByVal WorkbookName As String) As Variant
'---------------------------------------------------------------------------------------
' Procedure : GetNameValue
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Return the value referenced in a workbook name, not as a formula, which is
'           : what the Value property returns, but the actual value when the formula is
'           : evaluated.
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
Dim objName As Name
'*********************************

'Variants:
'*********************************
Dim vntValue As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************


On Error GoTo ErrProc

GetNameValue = Null

'========================================================
'Validate input.
'========================================================
If Trim(WorkbookName) = "" Then Exit Function

'========================================================
'Name must be valid.
'========================================================
On Error Resume Next

Set objName = ThisWorkbook.Names(WorkbookName)
If Err <> 0 Then
   Err.Clear
   Exit Function
End If


'========================================================
'If parameter isn't a simple data type like a string
'or numeric value, Evaluate() should fail.
'========================================================
vntValue = Evaluate(WorkbookName)
If Err <> 0 Then
   Err.Clear
   Exit Function
Else
   GetNameValue = vntValue
End If




Exit Function

ErrProc:
   GetNameValue = Null
   MsgBox "Error" & Err.Number & " (" & Err.Description & ") at line " & Erl & " in Function GetNameValue of Module " & MODULE_NAME
End Function

Public Function IsNamedRange(ByVal RangeName As String, Optional ByRef DataWorkbook As Workbook) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsNamedRange
      ' Author    : Adiv Abramson
      ' Date      : 3/20/2024
      ' Purpose   : Determine if specified RangeName is a named range in specified DataWorkbook.
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
      ' Versions  : 1.0 - 3/20/2024 - Adiv Abramson
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
      Dim objName As Name
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

20    IsNamedRange = False

30    If ActiveWorkbook Is Nothing Then
40       Exit Function
50    ElseIf UCase(ActiveWorkbook.Name) = "PERSONAL.XLSB" Then
60       Exit Function
70    ElseIf DataWorkbook Is Nothing Then
80       Set DataWorkbook = ActiveWorkbook
90    End If


100   For Each objName In DataWorkbook.Names
110      If UCase(objName.Name) = UCase(RangeName) Then
120         IsNamedRange = True
130         Exit Function
140      End If
150   Next objName


160   Exit Function

ErrProc:
170      IsNamedRange = False
180      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure IsNamedRange of Module " & MODULE_NAME
End Function

Public Function GetColumnIndex(ByVal TableName As String, ByVal ColumnName As String, Optional DataWorksheet As Variant = "Lists") As Integer
      '---------------------------------------------------------------------------------------
      ' Procedure : GetColumnIndex
      ' Author    : Adiv Abramson
      ' Date      : 10/1/2023
      ' Purpose   : Return the ordinal position of the specified column from the specified table
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
      ' Versions  : 1.0 - 10/1/2023 - Adiv Abramson
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
      Dim arColumns As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objDataTable As ListObject
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

20    GetColumnIndex = -1

30    Set objDataTable = GetTable(TableName:=TableName, _
                                  DataWorksheet:=IIf(IsMissing(DataWorksheet), WorksheetNames.Lists, DataWorksheet))
40    arColumns = GetUniqueRangeValues(DataRange:=objDataTable.HeaderRowRange)
      '========================================================
      'ArrayFind() works with 0-based arrays. So must add 1 to correspond to
      'table column position. If specified column does not exist, ArrayFind() returns
      '-1, so GetColumnIndex() will return a 0, which calling code can watch for.
      '========================================================
50    GetColumnIndex = ArrayFind(SearchValue:=ColumnName, DataArray:=arColumns) + 1



60    Exit Function

ErrProc:
70       GetColumnIndex = -1
80       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetColumnIndex of Module " & MODULE_NAME
End Function





Public Function CreateWorkbookQueryFromTable(ByRef DataWorkbook As Workbook, _
                                             ByVal TableName As String, _
                                             Optional ByVal QueryName As String = "", _
                                             Optional ByVal Description As String = "") As Boolean
'---------------------------------------------------------------------------------------
' Procedure : CreateWorkbookQueryFromTable
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Create a Power Query query from specified table in this workbook
'           : and load it to as a connection only, i.e. don't download the same
'           : data to a new or existing worksheet.
'           : Returns False if unable to create query, such as when a query
'           : with the same name (not necessarily the same data source,
'           : which can be loaded multiple times as long as a different
'           : name is used each time).
'           : Changed from sub to function so calling code can stop processing
'           : or modify count of loaded tables depending on value returned.
'           : (1.1) Now supports loading tables from an external workbook, instead
'           : of always loading from current workbook, i.e. the workbook in
'           : which this code is running. Had to add new required parameter
'           : specified the workbook from which to load tables. Also had to
'           : update M code to refer to an external workbook when necessary.
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
Dim strQueryFormula As String
Dim strQueryName As String
Dim strFormat As String
Dim strWorksheetName As String
'*********************************

'Numerics:
'*********************************

'*********************************

'Worksheets:
'*********************************
Dim wsQueryData As Worksheet
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
Dim objQuery As WorkbookQuery
Dim objQueries As Queries
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
      
20    CreateWorkbookQueryFromTable = False
      
      '========================================================
      'Validate inputs.
      '========================================================
30    If Trim(TableName) = "" Then Exit Function
40    If Len(QueryName) <> 0 And Trim(QueryName) = "" Then Exit Function
50    If Len(Description) <> 0 And Trim(Description) = "" Then Exit Function
      
      '========================================================
      'A non blank value must be assigned to QueryName in order
      'to create the Power Query query. So if no value is supplied
      'set QueryName equal to TableName
      '========================================================
60    If QueryName = "" Then QueryName = TableName
      
      '========================================================
      'Construct the value of the Formula property using
      'PrintFormat2().
      '========================================================
      '========================================================
      'NEW CODE: 00/00/0000 Now supporting loading tables from
      'an external workbook.
      '========================================================
70    If DataWorkbook.Name = ThisWorkbook.Name Then
80       strFormat = "let {$LF}"
90       strFormat = strFormat & "{$TAB}Source = Excel.CurrentWorkbook(){[Name=""{$tablename}""]}[Content]"
100      strFormat = strFormat & "{$LF}in{$LF}"
110      strFormat = strFormat & "{$TAB}Source"
120   Else
130      strFormat = "let {$LF}"
140      strFormat = strFormat & "{$TAB}Source = Excel.Workbook(File.Contents(""{$external_wb_name}""), null, true),"
150      strFormat = strFormat & "{$LF}{$tablename} = Source{[Item=""{$tablename}"",Kind=""Table""]}[Data]"
160      strFormat = strFormat & "{$LF}in{$LF}"
170      strFormat = strFormat & "{$TAB}Source"
180   End If 'DataWorkbook.Name = ThisWorkbook.Name
      
      
190   strQueryFormula = PrintFormat2(strFormat, "LF", vbCrLf, "TAB", vbTab, _
                                                "tablename", TableName, _
                                                "external_wb_name", DataWorkbook.FullName)
      'Debug.Print strQueryFormula
      
      '========================================================
      'Check if query with specified name already exists.
      '========================================================
220   On Error Resume Next
230   Set objQuery = DataWorkbook.Queries(QueryName)
240   If Not objQuery Is Nothing Then
250      msgAttention PrintFormat2("A query named ""{$queryname}"" already exists.", "queryname", QueryName)
260      Exit Function
270   End If
      
280   On Error GoTo ErrProc
290   Set objQuery = DataWorkbook.Queries.Add(Name:=QueryName, _
                                              Formula:=strQueryFormula, _
                                              Description:=Description)
      
320   CreateWorkbookQueryFromTable = True

Exit Function

ErrProc:
   CreateWorkbookQueryFromTable = False
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure CreateWorkbookQueryFromTable of Module " & MODULE_NAME
End Function


Public Function QueryList(ByVal TableName As String, ByVal ReturnFields As Variant, _
                          Optional ByVal QueryParams As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : QueryList
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of values from specified field of specified listobject
      '           : where optional query conditions are satisfied
      '           : (1.1) Enhance functionality with the ability to return any number of fields.
      '           : (1.2) Now that the filtered range can be discontiguous, the range for the result
      '           : may contain multiple areas
      '           : (1.3) Fixed error where code was validating criteria fields even when none were
      '           : specified.
      '           : (1.4) Correct bug in which function tries to query an empty table. If table is
      '           : empty, exit early and return Null, since an empty table cannot logically be
      '           : queried.
      '           : *Replaced ParamArray parameter with regular but still optional Array(),
      '           : for greater flexibility, since Arrays can be built on the fly whereas
      '           : ParamArrays have to be hard coded.
      '           : * Use GetSpecialCells() to reference visible cells in rngColumnData before passing
      '           : it to GetAllRangeValues(), since if the referenced range has only one cell, VBA's
      '           : built in SpecialCells() method blows up.
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
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           : 1.3 - 00/00/0000 - Adiv Abramson
      '           : 1.4 - 00/00/0000 - Adiv Abramson
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
      Dim strField As String
      Dim strReturnField As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intFields As Integer
      Dim intFieldIndex As Integer
      Dim intMatchingRows As Integer
      Dim intReturnField As Integer
      Dim intReturnFields As Integer
      Dim intRow As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngColumnData As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arFields As Variant
      Dim arCriteria As Variant
      Dim arValues() As Variant
      Dim arRangeValues As Variant
      Dim arTemp As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objList As ListObject
      Dim objListColumn As ListColumn
      '*********************************

      'Variants:
      '*********************************
      Dim vntCriteria As Variant
      Dim vntReturnField As Variant
      Dim vntRangeValue As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnHasCriteria As Boolean
      Dim blnIsField As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    QueryList = Null

30    If m_wsLists Is Nothing Then Exit Function

40    ReleaseListTabFilters

      '================================================================
      'Validate inputs
      '================================================================
50    If Not IsMissing(QueryParams) Then
60       intFields = UBound(QueryParams)
         'Array must have even number of elements.
         'Array is 0 based
70       If intFields Mod 2 = 0 Then
80          msgAttention "Invalid QueryParams list: Number of fields must match number of criteria."
90          Exit Function
100      End If 'intFields Mod 2 = 0
110      blnHasCriteria = True
120   Else
130      blnHasCriteria = False
140   End If 'Not IsMissing(QueryParams)

      '========================================================
      'NEW CODE: 00/00/0000 Populate both arrays in one pass.
      'No longer restricting criteria to unique values since
      'it's possible for two or more fields in a table to be
      'of the same data type and we want to return results
      'when all criteria are met, e.g. Gross Weight > 50,
      'Net Weight > 50.
      '========================================================
150   If blnHasCriteria Then
160      For i = 0 To intFields - 1 Step 2
170         AddToArray arFields, QueryParams(i), True
180         AddToArray arCriteria, QueryParams(i + 1), False
190      Next i
200   End If 'blnHasCriteria

210   Set objList = GetTable(TableName)
      '========================================================
      'NEW CODE: 00/00/0000 If table is empty, exit early
      '========================================================
220   If objList.DataBodyRange Is Nothing Then Exit Function

      '================================================================
      'Ensure ReturnField and any QueryFields actually exist in the
      'specified table
      '================================================================
      '========================================================
      'NEW CODE: 00/00/0000 Now supporting multiple return
      'fields.
      '========================================================
230   If Not IsArray(ReturnFields) Then
240      msgAttention "Invalid Return Fields. This parameter must be an array."
250      Exit Function
260   Else
270      intReturnFields = UBound(ReturnFields)
280   End If
         
      '========================================================
      'NEW CODE: 00/00/0000 Verify that each specificed return
      'field actually exists in the table
      '========================================================
290   For Each vntReturnField In ReturnFields
300      blnIsField = False
310      For Each objListColumn In objList.ListColumns
320         blnIsField = StringCompare(String1:=vntReturnField, String2:=objListColumn.Name)
330         If blnIsField Then Exit For
340      Next objListColumn
350      If Not blnIsField Then
360         msgAttention PrintFormat("{$0} does not exist in {$1} table.", vntReturnField, objList.Name)
370         Exit Function
380      End If
390   Next vntReturnField


      '========================================================
      'NEW CODE: 00/00/0000 Verify that the criteria fields
      'exist in the table.
      '========================================================
400   If blnHasCriteria Then
410      intFields = UBound(arFields)
420      For i = 0 To intFields
430         strField = arFields(i)
440         blnIsField = False
450         For Each objListColumn In objList.ListColumns
460            blnIsField = StringCompare(String1:=strField, String2:=objListColumn.Name)
470            If blnIsField Then Exit For
480         Next objListColumn
490         If Not blnIsField Then
500            msgAttention PrintFormat("{$0} does not exist in {$1} table.", strField, objList.Name)
510            Exit Function
520         End If
530      Next i
540   End If 'blnHasCriteria

      '================================================================
      'Apply filters, if any
      '================================================================
550   If blnHasCriteria Then
560      For i = 0 To intFields
570         strField = arFields(i)
580         vntCriteria = arCriteria(i)
590         Set objListColumn = objList.ListColumns(strField)
600         intFieldIndex = objListColumn.Index
610         objList.DataBodyRange.AutoFilter Field:=intFieldIndex, Criteria1:=vntCriteria
620      Next i
630   End If

      '========================================================
      'NEW CODE: 00/00/0000 Since filtered range may be discon-
      'tiguous, it may contain multiple areas. So Rows.Count
      'cannot be used since it only references the first area.
      'But it can be used to detect if there are any results
      'at all in the filtered table.
      '========================================================
      '========================================================
      'NEW CODE: 00/00/0000 Use GetSpecialCells() since SpecialCells()
      'will blow up if referenced range has only one cell.
      '========================================================
640   Set rngColumnData = GetSpecialCells(DataRange:=objList.DataBodyRange.Columns(1), CellType:=xlCellTypeVisible)
      '========================================================
      'NEW CODE: 00/00/0000 If no records matched the AutoFilter,
      'GetSpecialCells() now returns Nothing. In this case,
      'clean up and exit early.
      '========================================================
650   If rngColumnData Is Nothing Then
660      ReleaseListTabFilters
670      Exit Function
680   End If 'rngColumnData Is Nothing

690   intMatchingRows = rngColumnData.Rows.Count
700   arTemp = GetAllRangeValues(DataRange:=rngColumnData)
710   intMatchingRows = UBound(arTemp) + 1

720   On Error GoTo ErrProc

      '========================================================
      'NEW CODE: 00/00/0000 Now that it's possible to return
      'multiple fields, the output array must be sized before
      'being populated
      '========================================================
730   ReDim arValues(intMatchingRows - 1, intReturnFields)

740   For intReturnField = 0 To intReturnFields
750      strReturnField = ReturnFields(intReturnField)
         '========================================================
         'NEW CODE: 00/00/0000 Use GetSpecialCells() since SpecialCells()
         'will blow up if referenced range has only one cell.
         '========================================================
760      Set rngColumnData = GetSpecialCells(DataRange:=objList.ListColumns(strReturnField).DataBodyRange, CellType:=xlCellTypeVisible)
         '========================================================
         'NEW CODE: 00/00/0000 Range referenced may contain more
         'than one area.
         '========================================================
770      intRow = 0
780      arRangeValues = GetAllRangeValues(DataRange:=rngColumnData)
790      For Each vntRangeValue In arRangeValues
800         arValues(intRow, intReturnField) = vntRangeValue
810         Incr intRow, 1
820      Next vntRangeValue
830   Next intReturnField


840   ReleaseListTabFilters

850   QueryList = arValues

860   Exit Function

ErrProc:
870    QueryList = Null
880    ReleaseListTabFilters
890    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure QueryList of Module mdlUtil"

End Function


Public Function CompareAll(ByVal ComparisonOp As ComparisonOperator, _
                           ByVal CompareValue As Variant, _
                           ByVal CompareValues As Variant, _
                           Optional MatchingValues As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CompareAll
      ' Author    : Adiv Abramson
      ' Date      : 9/21/2016
      ' Purpose   : Allow multiple values to be compared against a single value
      '           : (1.1) Renamed vntValues to CompareValues. Also changed to
      '           : regular parameter from ParamArray so that new optional
      '           : MatchingValues array can be populated with values from
      '           : CompareValues array that satisfy criterion.
      '           : *Now counting the number of matching values to determine
      '           : the return value of the function. This simplifies the code
      '           : for adding to the MatchingValues array
      '           : (1.2) Now supporting IsNotEqualTo comparisons.
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
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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
      Dim intSize As Integer
      Dim i As Integer
      Dim intMatchingValues As Integer
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
      Dim blnReturnMatchingValues As Boolean
      Dim blnIsMatch As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

      'validate input
20    If Not IsArray(CompareValues) Then
30       CompareAll = False
40       Exit Function
50    ElseIf UBound(CompareValues) = -1 Then
         'this test will pass if function is called without any comparison values
60       CompareAll = False
70       Exit Function
80    End If

90    intSize = UBound(CompareValues)

      '========================================================
      'NEW CODE: 00/00/0000 If the option MatchingValues
      'parameter is supplied, populate with values that
      'satisfy matching criterion.
      '========================================================
100   blnReturnMatchingValues = Not IsMissing(MatchingValues)

      '========================================================
      'Initialize variable to track if all comparisons are true
      'if any comparison fails, this will be set to false and
      'the function will exit
      '========================================================
      '========================================================
      'NEW CODE: 00/00/0000 Function returns True if number of
      'matching values equals number of items in CompareValues
      'array
      '========================================================

110   intMatchingValues = -1

120   For i = 0 To intSize
130      blnIsMatch = False
140      Select Case ComparisonOp
            Case ComparisonOperator.IsEqualTo
150            blnIsMatch = (CompareValues(i) = CompareValue)
160         Case ComparisonOperator.IsGreaterThan
170            blnIsMatch = (CompareValues(i) > CompareValue)
180         Case ComparisonOperator.IsGreaterThanOrEqualTo
190            blnIsMatch = (CompareValues(i) >= CompareValue)
200         Case ComparisonOperator.IsLessThan
210            blnIsMatch = (CompareValues(i) < CompareValue)
220         Case ComparisonOperator.IsLessThanOrEqualTo
230            blnIsMatch = (CompareValues(i) <= CompareValue)
240         Case ComparisonOperator.IsNotEqualTo
250            blnIsMatch = (CompareValues(i) <> CompareValue)
260      End Select
         
270      If blnIsMatch Then
280         Incr intMatchingValues, 1
290         If blnReturnMatchingValues Then
300            AddToArray DataArray:=MatchingValues, InputValue:=CompareValues(i), KeepUnique:=False
310         End If 'blnReturnMatchingValues
320      End If 'blnIsMatch
330   Next i

340   CompareAll = (UBound(CompareValues) = intMatchingValues)

350   Exit Function

ErrProc:

360       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure CompareAll of Module " & MODULE_NAME
End Function





Public Function IsBetween(ByVal vntTestValue As Variant, _
                          ByVal vntValue1 As Variant, _
                          ByVal vntValue2 As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsBetween
      ' Author    : Adiv Abramson
      ' Date      : 8/17/2016
      ' Purpose   : Indicate if vntTestValue is between vntValue1 and vntValue2
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 8/17/2016 - Adiv Abramson
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------



10    On Error GoTo ErrProc

20    If IsNull(vntTestValue) Or IsNull(vntValue1) Or IsNull(vntValue2) Then
30       IsBetween = False
40       Exit Function
50    End If

60    If vntTestValue >= vntValue1 And vntTestValue <= vntValue2 Then
70       IsBetween = True
80    Else
90       IsBetween = False
100   End If


110   Exit Function

ErrProc:

120       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsBetween of Module " & MODULE_NAME
End Function



Public Function GetDataLink() As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetDataLink
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Display Data Link Properties dialog so user can select a data source. Return the connection string
      '           : (1.1) Handle case where user cancels dialog.
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
      Dim objDL As Object
      '*********************************

      'Variants:
      '*********************************
      Dim vntConnection As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************


10    On Error Resume Next

20    GetDataLink = Null
30    Set objDL = CreateObject("DataLinks")
40    vntConnection = objDL.PromptNew()
50    If Err <> 0 Then
60       Err.Clear
70       Exit Function
80    End If

90    On Error GoTo ErrProc

100   GetDataLink = vntConnection


110   Exit Function

ErrProc:
120      GetDataLink = Null
130      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl _
         & " of procedure GetDataLink of Module " & MODULE_NAME
End Function

Public Sub Navigate(ByVal Direction As TabNavigation)
      '---------------------------------------------------------------------------------------
      ' Procedure : Navigate
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   :
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
      Dim lngCurrentRow As Long
      Dim intCurrentCol As Integer
      Dim lngFirstRow As Long
      Dim lngFirstCol As Integer
      Dim lngLastRow As Long
      Dim lngLastCol As Long
      Dim lngArea As Long
      Dim lngAreas As Long
      '*********************************

      'Worksheets:
      '*********************************
      Dim wsData As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngHeaders As Range
      Dim rngData As Range
      Dim rngTarget As Range
      Dim rngArea As Range
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

20    Set wsData = ActiveSheet
30    If Not HasData(wsData) Then Exit Sub
      'If IsInList(wsData.Name, WorksheetNames.ChangeReport, WorksheetNames.ExecutiveSummary, _
                  WorksheetNames.MonthlyChangePivotReport) Then Exit Sub

40    Set rngData = wsData.UsedRange.SpecialCells(xlCellTypeVisible)
50    lngAreas = rngData.Areas.Count

60    With ActiveCell
70       lngCurrentRow = .Row
80       intCurrentCol = .Column
90    End With

100   Select Case Direction
         Case MoveToFirstColumn
110         Set rngArea = rngData.Areas(1)
120         Set rngTarget = wsData.Cells(lngCurrentRow, rngArea.Cells(1, 1).Column)
130      Case MoveToFirstRow
140         Set rngArea = rngData.Areas(1)
150         Set rngTarget = wsData.Cells(rngArea.Cells(1, 1).Row, intCurrentCol)
160      Case MoveToLastColumn
170         Set rngArea = rngData.Areas(lngAreas)
180         Set rngTarget = wsData.Cells(lngCurrentRow, rngArea.Cells(1, rngArea.Columns.Count).Column)
190      Case MoveToLastRow
200         Set rngArea = rngData.Areas(lngAreas)
210         Set rngTarget = wsData.Cells(rngArea.Cells(rngArea.Rows.Count, 1).Row, intCurrentCol)

220   End Select

230   rngTarget.Select
240   With ActiveWindow
250      .ScrollRow = rngTarget.Row
260      .ScrollColumn = rngTarget.Column
270   End With


280   Exit Sub

ErrProc:

290      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Navigate of Module " & MODULE_NAME

End Sub


Public Function GetAllRangeValues(ByRef DataRange As Range) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetAllRangeValues
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of the all values in specified 1 column or 1 row range.
      '           :
      '           : Calling SpecialCells(xlCellTypeVisible) seems to return invalid results if the range to
      '           : which it is applied is only a single cell. It will return reference to entire worksheet
      '           : (1.1) Now iterating only through visible cells.
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
      Dim lngArea As Long
      Dim lngAreas As Long
      Dim lngSize As Long
      Dim lngCellCount As Long
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
      Dim rngArea As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arAllValues() As Variant
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

20    GetAllRangeValues = Null

      '========================================================
      'Initialize UBound of output array to dummy value
      '========================================================
30    lngSize = -1

40    If DataRange Is Nothing Then
50       GetAllRangeValues = Null
60       Exit Function
70    ElseIf DataRange.Cells.Count = 1 Then
80       GetAllRangeValues = Array(DataRange.Cells(1, 1).Value)
90       Exit Function
100   Else
110      Set DataRange = DataRange.SpecialCells(xlCellTypeVisible)
120   End If

      '========================================================
      'Iterate through all the cells of all the areas of the
      'specified range and add the values to the output array.
      '========================================================

      '<NEW CODE:00/00/0000> When range has been filtered or cells have otherwise
      'been hidden, we will have multiple areas to cycle through
130   lngAreas = DataRange.Areas.Count
140   For lngArea = 1 To lngAreas
150      Set rngArea = DataRange.Areas(lngArea)
160      For Each rngCell In rngArea.Cells
170         Incr lngSize, 1
180         ReDim Preserve arAllValues(lngSize)
190         arAllValues(lngSize) = rngCell.Value
200      Next rngCell
210   Next lngArea

      '========================================================
      'Like GetUniqueRangeValues, the returned
      'array is 0 based in this function
      '========================================================

220   GetAllRangeValues = arAllValues


230   Exit Function

ErrProc:
240      GetAllRangeValues = Null
250      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetAllRangeValues of Module mdlUtil"
End Function




Public Function GetLastCell(ByRef DataRange As Range) As Range
      '---------------------------------------------------------------------------------------
      ' Procedure : GetLastCell
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Find the last cell in the specified range. Account for multiple areas.
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
      Dim intAreas As Integer
      Dim intColumns As Integer
      Dim lngRows As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngLastArea As Range
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

20    Set GetLastCell = Nothing

30    If DataRange Is Nothing Then Exit Function
40    intAreas = DataRange.Areas.Count
50    Set rngLastArea = DataRange.Areas(intAreas)
60    With rngLastArea
70       lngRows = .Rows.Count
80       intColumns = .Columns.Count
90       Set GetLastCell = .Cells(lngRows, intColumns)
100   End With

110   Exit Function

ErrProc:
120      Set GetLastCell = Nothing
130      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetLastCell of Module mdlUtil"

End Function

Public Function IsReallyNumeric(ByVal vntValue As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsReallyNumeric
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Since IsNumeric() returns true on what appears to be an empty cell
      '           : use this custom function instead, which will return true only if
      '           : the cell contains a value whose length is more than zero and that
      '           : can be cast as a number
      '           : (1.1) VBA thinks a string like "7,8,9" is numeric. We can't have that!
      '           : So use regex to detect any non numeric chars and if found return False
      '           : (1.2) Now checking for Nulls
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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
      Dim intLen As Integer
      Dim dblTestValue As Double
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
      Dim objRegex As New RegExp
      Dim objMatches As MatchCollection
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

20    IsReallyNumeric = False

30    If IsNull(vntValue) Then Exit Function
40    If IsEmpty(vntValue) Then Exit Function
50    If Len(vntValue) = 0 Then Exit Function

      'Set up regex to find non numerics, if any
60    With objRegex
70       .Global = True
80       .Pattern = "^\-?(\d|\.)+$"
90       If .Test(vntValue) Then
            'vntValue might have more than one decimal point, which would make it not a number
100         On Error Resume Next
110         dblTestValue = CDbl(vntValue)
120         If Err = 0 Then
130            IsReallyNumeric = True
140         Else
150            Err.Clear
160            IsReallyNumeric = False
170         End If 'Err = 0
180      Else
190         IsReallyNumeric = False
200      End If
210   End With

220   Exit Function

ErrProc:
230       IsReallyNumeric = False
240       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsReallyNumeric of Module mdlUtil"
End Function






Public Sub ReleaseListTabAutoFilters()
      '---------------------------------------------------------------------------------------
      ' Procedure : ReleaseListTabAutoFilters
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Clear any filtering on the tables on the Lists tab
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
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 00/00/0000 - Adiv Abramson
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
      Dim wsLists As Worksheet
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
      Dim objTable As ListObject
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

      '=========================================
      'Release any AutoFilters on Lists tab
      '=========================================
20    Set wsLists = ThisWorkbook.Worksheets(WorksheetNames.Lists)
30    For Each objTable In wsLists.ListObjects
40       With objTable
50          If Not .AutoFilter Is Nothing Then
60             If .AutoFilter.FilterMode Then
70                .AutoFilter.ShowAllData
80             End If '.AutoFilter.FilterMode
90          End If 'Not .AutoFilter Is Nothing
100      End With
110   Next objTable

        
120   Exit Sub

ErrProc:

130       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure ReleaseListTabAutoFilters of Module mdlFFP"
End Sub


Public Sub CleanList(ByVal TableName As String)
'---------------------------------------------------------------------------------------
' Procedure : CleanList
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Delete all databodyrange rows in specified table
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
Dim objListRow As ListRow
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

10    On Error GoTo ErrProc
      
20    ReleaseListTabAutoFilters
      
30    Set objList = GetTable(TableName:=TableName)
40    With objList
50       If Not .DataBodyRange Is Nothing Then
60          While .ListRows.Count > 0
70             Set objListRow = .ListRows(1)
80             objListRow.Delete
90          Wend
100      End If
110   End With




Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure CleanList of Module mdlUtil"

End Sub

Public Function GetSpecialCells(ByRef DataRange As Range, CellType As XlCellType) As Range
      '---------------------------------------------------------------------------------------
      ' Procedure : GetSpecialCells
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Wrapper function to return a single cell matching cell type
      '           : when DataRange consists of only 1 cell. Otherwise Excel will
      '           : return something unexpected, like multiple ranges of cells
      '           : (1.) Corrected bug where code failed to detect a single cell
      '           : range before trying to use SpecialCells() on it. When a single
      '           : cell range is passed to the function, ALWAYS return it as is.
      '           : (1.2) If no cells match specifed type, exit early and return
      '           : Nothing.
      '           : (1.3) Check if single cell range reference is hidden, as when an AutoFilter
      '           : is applied to a range but no cells match the criteria. In such case, exit early
      '           : and return Nothing.
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
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           : 1.3 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************

      '*********************************

      'Numerics:
      '*********************************
      Dim lngCells As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngResult As Range
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

10    On Error Resume Next

20    Set GetSpecialCells = Nothing

30    If DataRange.Cells.Count = 1 Then
         '========================================================
         'NEW CODE: 00/00/0000 If single cell range reference is
         'hidden, must exit early and return Nothing. It's because
         'no cells satisfied AutoFilter criteria.
         '========================================================
40       If DataRange.EntireRow.Hidden Then Exit Function
50       Set GetSpecialCells = DataRange
60       Exit Function
70    Else
80       On Error Resume Next
90       Set rngResult = DataRange.SpecialCells(CellType)
100      If rngResult Is Nothing Then Exit Function
         
110      lngCells = rngResult.Cells.Count
120      If Err <> 0 Then
130         Err.Clear
140         Exit Function
150      End If
160      Set GetSpecialCells = rngResult
170   End If
       
180   Exit Function

ErrProc:
190       Set GetSpecialCells = Nothing
200       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetSpecialCells of Module mdlUtil"
End Function


Public Function HasValue(ByRef ctrl As control) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : HasValue
      ' Author    : Adiv Abramson on 00/00/0000 at 15:54
      ' Date      : 00/00/0000
      ' Purpose   : Determine if control has a value other than Null or zero length string
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------



10    On Error GoTo ErrProc

20    If IsNull(ctrl.Value) Or ctrl.Value = "" Then
30       HasValue = False
40    Else
50      HasValue = True
60    End If



70    Exit Function

ErrProc:

80        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure HasValue of Module mdlUtil"

End Function
Public Sub CleanWorksheet(ByVal SheetName As String)
      '---------------------------------------------------------------------------------------
      ' Procedure : CleanWorksheet
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Copies corresponding template of specified worksheet, replacing existing
      '           : specified worksheet or, if it doesn't exist, creating a new instance
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
      '           :
      '           :
      '           :
      'Change Log :
      '           :---------------------------------
      '           : 1.0 - 00/00/0000 - Adiv Abramson
      '           :---------------------------------
      '           : *Initial coding
      '           :---------------------------------
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           :---------------------------------
      '           : *Ensure that that there are at least
      '           : 2 visible worksheets otherwise trying to
      '           : delete the only visible worksheet will fail
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
      '           :---------------------------------
      '           : 1.X - NN/NN/2019 - Adiv Abramson
      '           :---------------------------------
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
      Dim strTimeStamp As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intVisibleWorksheets As Integer
      '*********************************

      'Worksheets:
      '*********************************
      Dim wsCleanSheet As Worksheet
      Dim wsTemplate As Worksheet
      Dim wsTemp As Worksheet
      Dim ws As Worksheet
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

      'Constants:
      '*********************************
      Const TEMPLATE_PREFIX = "_"
      '*********************************

10    On Error GoTo ErrProc


      '========================================================
      '00/00/0000>NEW CODE: If there is only one visible
      'worksheet must add a temp worksheet otherwise the delete
      'operation will fail
      '========================================================
20    intVisibleWorksheets = 0
30    For Each ws In ThisWorkbook.Worksheets
40       If ws.Visible = xlSheetVisible Then
50          Incr intVisibleWorksheets, 1
60       End If
70    Next ws

80    If intVisibleWorksheets = 1 Then
90       Set wsTemp = ThisWorkbook.Worksheets.Add
100      strTimeStamp = Format(Now, "hhnnss")
110      wsTemp.Name = strTimeStamp
120   End If

130   If Not IsWorksheet(SheetName:=TEMPLATE_PREFIX & SheetName) Then
140      msgAttention "Template for worksheet " & SheetName & " does not exist."
150      Exit Sub
160   Else
170      Set wsTemplate = ThisWorkbook.Worksheets(TEMPLATE_PREFIX & SheetName)
180   End If 'Not IsWorksheet(SheetName:=TEMPLATE_PREFIX & SheetName)

      'Delete and replace an existing worksheet
190   If IsWorksheet(SheetName:=SheetName) Then
200      Application.DisplayAlerts = False
210      ThisWorkbook.Worksheets(SheetName).Delete
220      Application.DisplayAlerts = True
230   End If 'IsWorksheet(SheetName:=SheetName)

240   With wsTemplate
250      .Visible = xlSheetVisible
260      .Copy Before:=ThisWorkbook.Worksheets(1)
270      .Visible = xlSheetHidden
280   End With

290   ActiveSheet.Name = SheetName

      '========================================================
      '00/00/0000>NEW CODE: Remove temp worksheet, if it exists
      '========================================================
300   If Not wsTemp Is Nothing Then
310      Application.DisplayAlerts = False
320      wsTemp.Delete
330      Application.DisplayAlerts = True
340   End If


350   Exit Sub

ErrProc:
360      Application.DisplayAlerts = True
370      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure CleanWorksheet of Module mdlUtil"
End Sub


Public Function GetWorksheet_ColInfo(ByRef DataWorksheet As Worksheet, _
    Optional ByVal HeaderRow As Integer = 1) As Dictionary
      '---------------------------------------------------------------------------------------
      ' Procedure : GetWorksheet_ColInfo
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Automatically create dictionary based on specified worksheet's headers.
      '           : Keys are headers and items are Array(<column letter>, <index>, <range_excluding_header>)
      '           : Data must begin in column A
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
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
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
      Dim strCol As String
      Dim strAddress As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intIndex As Integer
      Dim lngLastDataRow As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngHeaders As Range
      Dim rngCell As Range
      Dim rngColumnData As Range
      Dim rngDataWithoutHeaders As Range
      Dim rngDataWithHeaders As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arColInfo As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim dictColInfo As Dictionary
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

20    Set GetWorksheet_ColInfo = Nothing

30    If Application.WorksheetFunction.CountA(DataWorksheet.Rows(HeaderRow)) = 0 Then
40       Exit Function
50    End If

      '===========================================
      'Get last row of data on specified worksheet
      'to compose range references
      '===========================================
60    lngLastDataRow = GetLastDataRow(DataWorksheet:=DataWorksheet)

      '===========================================
      'Inspect headers to get dictionary keys
      '===========================================
70    Set rngHeaders = DataWorksheet.Rows(HeaderRow).SpecialCells(xlCellTypeConstants)
80    Set dictColInfo = New Dictionary

90    For Each rngCell In rngHeaders.Cells
100      arColInfo = GetHeaderInfo(DataWorksheet:=DataWorksheet, Pattern:=rngCell.Value, ReturnType:=HeaderInfoColumnAndIndex, WholeString:=True)
110      strCol = arColInfo(COLUMN_HEADING)
120      intIndex = arColInfo(COLUMN_INDEX)
130      If lngLastDataRow > 1 Then
140         strAddress = strCol & HeaderRow + 1 & ":" & strCol & lngLastDataRow
150         Set rngColumnData = DataWorksheet.Range(strAddress)
160      Else
170         Set rngColumnData = Nothing
180      End If 'lngLastDataRow > 1
190      dictColInfo.Item(UCase(rngCell.Value)) = Array(strCol, intIndex, rngColumnData)
200   Next rngCell

210   With DataWorksheet
220      Set rngDataWithHeaders = .Range(.Cells(HeaderRow, 1), .Cells(lngLastDataRow, _
            rngHeaders.Cells(1, rngHeaders.Columns.Count).Column))
         
230      Set rngDataWithoutHeaders = .Range(.Cells(HeaderRow + 1, 1), .Cells(lngLastDataRow, _
            rngHeaders.Cells(1, rngHeaders.Columns.Count).Column))
240   End With

250   Set dictColInfo.Item(DATA_RANGE) = Array(rngDataWithHeaders, rngDataWithoutHeaders)

260   Set GetWorksheet_ColInfo = dictColInfo

270   Exit Function

ErrProc:
280      Set GetWorksheet_ColInfo = Nothing
290      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetWorksheet_ColInfo of Module mdlUtil"
End Function


Public Function GetLastDataRow(ByRef DataWorksheet As Worksheet, _
                               Optional ByVal HeaderRow As Integer = 1) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : GetLastDataRow
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Determine last row of data in used range of specified worksheet
      '           : (1.1) Must return the same last row whether or not some rows have been hidden.
      '           : Last cell in used range column may appear to be empty but End() may
      '           : pass over it as if it were populated. So must determine where the final
      '           : non blank cell is in each column.
      '           : (1.2) Fixed bug in which all cells in a column of the UsedRange were not being
      '           : properly referenced.
      '           : *UsedRange contains blank cells. They should not be included when determining
      '           : last row of data.
      '           : *Now supporting optional HeaderRow parameter.
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
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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

      '*********************************

      'Numerics:
      '*********************************
      Dim lngLastRow As Long
      Dim lngLastRow_Col As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngData As Range
      Dim rngCell As Range
      Dim rngHeaders As Range
      Dim rngColumnData As Range
      Dim rngLastCell As Range
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

20    GetLastDataRow = 0


30    Set rngData = DataWorksheet.UsedRange
40    If Application.WorksheetFunction.CountA(rngData) = 0 Then
50       Exit Function
60    End If

70    Set rngHeaders = rngData.Rows(HeaderRow)

80    lngLastRow = 0
90    For Each rngCell In rngHeaders.Cells
100      Set rngColumnData = DataWorksheet.UsedRange.Columns((rngCell.Column))
110      If Application.WorksheetFunction.CountA(rngColumnData) <= 1 Then
            '=================================
            'Column is either completely blank
            'or only has a header
            '=================================
120         lngLastRow_Col = 0
130      ElseIf Application.WorksheetFunction.CountBlank(rngColumnData) = rngColumnData.Cells.Count Then
140         lngLastRow_Col = 0
150      Else
160         With rngColumnData
170            Set rngLastCell = .Cells(.Rows.Count, 1)
180         End With
            '=============================================
            'Last cell may appear to be blank but End()
            'will pass over it as if it were populated.
            'So find last cell in column that is truly
            'not a ZLS.
            'Watch out for cells with errors.
            '=============================================
190         While rngLastCell.Row > HeaderRow
200            If Not IsError(rngLastCell) Then
210               If rngLastCell.Value = "" Or IsEmpty(rngLastCell) Then
220                  Set rngLastCell = rngLastCell.Offset(RowOffset:=-1)
230               Else
240                  lngLastRow_Col = rngLastCell.Row
250                  GoTo NextCell
260               End If 'rngLastCell.Value = "" Or IsEmpty(rngLastCell)
270            Else
280               Set rngLastCell = rngLastCell.Offset(RowOffset:=-1)
290               lngLastRow_Col = rngLastCell.Row
300            End If
310         Wend
320      End If 'Application.WorksheetFunction.CountA(rngColumnData) <= 1
NextCell:
330      lngLastRow = Application.WorksheetFunction.Max(lngLastRow_Col, lngLastRow)
340   Next rngCell


350   GetLastDataRow = lngLastRow

360   Exit Function

ErrProc:
370      GetLastDataRow = 0
380      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetLastDataRow of Module mdlUtil"
End Function


Public Function GetHeaderInfo(ByRef DataWorksheet As Worksheet, _
                              ByVal Pattern As String, ReturnType As HeaderInfo, _
                              Optional WholeString As Boolean = True, _
                              Optional HeaderRow As Integer = 1) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetHeaderInfo
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Return either the column letter or index or both (as Array) of specified header
      '           : pattern
      '           : (1.1) Match against whole string (^..$) unless otherwise specified
      '           : (1.2) Now supporting returning a reference to found cell
      '           : (1.3) Now supporting header search on other than row #1, which is the default
      '           : (1.4) Must escape certain chars like ?, \, [, etc
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
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           : 1.3 - 00/00/0000 - Adiv Abramson
      '           : 1.4 - 00/00/0000 - Adiv Abramson
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
      Dim strHeader As String
      Dim strCol As String
      '*********************************

      'Numerics:
      '*********************************
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
      Dim rngHeaders As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      '*********************************

      'Variants:
      '*********************************
      Dim vntDefaultReturnValue As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc


20    Select Case ReturnType
         Case HeaderInfoColumn
30          vntDefaultReturnValue = ""
40       Case HeaderInfoIndex
50          vntDefaultReturnValue = 0
60       Case HeaderInfoColumnAndIndex
70          vntDefaultReturnValue = Null
         '<NEW CODE:00/00/0000>
80       Case HeaderInfoReference
            'Nothing
90    End Select


100   If ReturnType = HeaderInfoReference Then
110      Set GetHeaderInfo = Nothing
120   Else
130      GetHeaderInfo = vntDefaultReturnValue
140   End If

      '<NEW CODE:00/00/0000> Use HeaderRow variable to support search on other rows
150   If Application.WorksheetFunction.CountA(DataWorksheet.Rows(HeaderRow)) = 0 Then
160      Exit Function
170   End If

      '<NEW CODE:00/00/0000> Must escape certain characters
180   Set objRegex = New RegExp
190   With objRegex
200      .IgnoreCase = True
210      .Global = True
220      .Pattern = "[\?\.\|\*\^\$\(\)\[\]\-]"
230      If .Test(Pattern) Then
240         Pattern = .Replace(Pattern, "\$&")
250      End If
260   End With






      '<NEW CODE:00/00/0000> Optionallly prepend "^" and append "$" to match whole string
270   If WholeString Then
280      If Left(Pattern, 1) <> "^" And Right(Pattern, 1) <> "$" Then
290         Pattern = "^" & Pattern & "$"
300      End If
310   End If

320   Set objRegex = New RegExp
330   With objRegex
340      .IgnoreCase = True
350      .Global = False
360      .Pattern = Pattern
370   End With

      '<NEW CODE:00/00/0000> Use HeaderRow variable to support search on other rows
380   Set rngHeaders = DataWorksheet.Rows(HeaderRow).SpecialCells(xlCellTypeConstants)
390   For Each rngCell In rngHeaders.Cells
400      strHeader = rngCell.Value
410      If objRegex.Test(strHeader) Then
420         strCol = GetChars(rngCell.Address, Alpha)
430         intIndex = rngCell.Column
440         Select Case ReturnType
            
               Case HeaderInfoColumn
450               GetHeaderInfo = strCol
460            Case HeaderInfoIndex
470               GetHeaderInfo = intIndex
480            Case HeaderInfoColumnAndIndex
490               GetHeaderInfo = Array(strCol, intIndex)
500            Case HeaderInfoReference
510               Set GetHeaderInfo = rngCell
520         End Select
530         Exit Function
540      End If
550   Next rngCell




560   Exit Function

ErrProc:
570      GetHeaderInfo = vntDefaultReturnValue
580      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetHeaderInfo of Module mdlUtil"
End Function
Public Function GetUniqueRangeValues(ByRef DataRange As Range) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetUniqueRangeValues
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of the unique values in specified range. Based on mdlMONY::GetUniqueFieldValues
      '           :
      '           : Calling SpecialCells(xlCellTypeVisible) seems to return invalid results if the range to
      '           : which it is applied is only a single cell. It will return reference to entire worksheet
      '           : (1.1) Fixed bug in which first item in range was excluded
      '           : (1.2) Use a dictionary instead of Excel's Advanced Filter for far better performance.
      '           : Array returned is now ZERO BASED!
      '           : (1.3) Now supporting optional skip over first cell, such as if it's a header that we don't
      '           : want to return in the array
      '           : (1.4) For a filtered range or with otherwise hidden rows/columns, there will be multiple
      '           : areas, which must be cycled through.
      '           : Removed Skip Over First option, since it was rarely used
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
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           : 1.3 - 00/00/0000 - Adiv Abramson
      '           : 1.4 - 00/00/0000 - Adiv Abramson
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

      '*********************************

      'Numerics:
      '*********************************
      Dim lngCellCount As Long
      Dim lngArea As Long
      Dim lngAreas As Long
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
      Dim rngArea As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim dictValues As Dictionary
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************
      Dim blnIncludeCell As Boolean
      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    GetUniqueRangeValues = Null

      '<NEW CODE:00/00/0000> Use dictionary instead of Advanced Filter for better performance
30    If DataRange Is Nothing Then
40       GetUniqueRangeValues = Null
50       Exit Function
60    ElseIf DataRange.Cells.Count = 1 Then
70       GetUniqueRangeValues = Array(DataRange.Cells(1, 1).Value)
80       Exit Function
90    End If


100   Set dictValues = New Dictionary
      'The following code auto dedupes
110   lngAreas = DataRange.Areas.Count
120   For lngArea = 1 To lngAreas
130      Set rngArea = DataRange.Areas(lngArea)
140      For Each rngCell In rngArea.Cells
150         dictValues.Item(rngCell.Value) = ""
160      Next rngCell
170   Next lngArea


      'Return 0 based array
180   GetUniqueRangeValues = dictValues.Keys


190   Exit Function

ErrProc:
200      GetUniqueRangeValues = Null
210      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetUniqueRangeValues of Module mdlUtil"
End Function


Public Function GetRangeAddress(ByRef objRange As Range) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GetRangeAddress
      ' Author    : Adiv Abramson
      ' Date      : 3/7/2016
      ' Purpose   : Address property cannot exceed 255 characters; construct string
      '           : of addresses for each area in range
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 3/7/2016 - Adiv Abramson
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      Dim intAreaCount As Integer
      Dim i As Integer
      Dim strAddress As String


10    On Error GoTo ErrProc

20    intAreaCount = objRange.Areas.Count

30    For i = 1 To intAreaCount
40        strAddress = IIf(strAddress = "", objRange.Areas(i).Address, _
              strAddress & "," & objRange.Areas(i).Address)
50     Next i

60    GetRangeAddress = strAddress


70    Exit Function

ErrProc:
80        GetRangeAddress = ""
90        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure GetRangeAddress of Module mdlUtil"
End Function

Public Function StringCompare(ByVal String1 As String, ByVal String2 As String, _
                             Optional ByVal PatternMatching As Variant = False) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : StringCompare
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Shortcut method for performing case insensitive string comparison
      '           : (1.1) Use Regex for pattern matching
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
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
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

20    If Not PatternMatching Then
30        If UCase(String1) = UCase(String2) Then
40           StringCompare = True
50        Else
60           StringCompare = False
70        End If
80    Else
         '<NEW CODE: 00/00/0000 Use Regex for pattern matching
90       Set objRegex = New RegExp
100      With objRegex
110         .Global = True
120         .IgnoreCase = True
130         .Pattern = String2
140         StringCompare = .Test(String1)
150      End With
160   End If 'Not PatternMatching

170   Exit Function

ErrProc:
180       StringCompare = False
190       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure StringCompare of Module mdlUtil"
End Function

Public Function IsLoaded(ByVal strFormName As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsLoaded
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Returns True if specified form is loaded into memory.
      '           : Returns False if specified form is not loaded into memory
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------



      Dim objForm As Object

10    On Error GoTo ErrProc

20    IsLoaded = False

30    For Each objForm In UserForms
40         If StringCompare(objForm.Name, strFormName) Then
50            IsLoaded = True
60            Exit Function
70         End If
80    Next objForm


90    Exit Function

ErrProc:
100       IsLoaded = False
110       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsLoaded of Module mdlUtil"
              
End Function
Public Function IsInList(ByVal vntLookup As Variant, ParamArray arValues() As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsInList
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Indicate whether vntLookup is in the list of values in arValues
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

      Dim i As Integer
      Dim intSize As Integer


         

10    On Error GoTo ErrProc


      'validate inputs
20    If Not IsArray(arValues) Then
30       IsInList = False
40       Exit Function
50    End If

60    intSize = UBound(arValues)

70    For i = 0 To intSize
80       If arValues(i) = vntLookup Then
90          IsInList = True
100         Exit Function
110      End If
120   Next i

130   IsInList = False

140   Exit Function

ErrProc:

150     MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsInList of Module mdlUtil"

End Function

Public Function IsWorksheet(ByVal SheetName As String, Optional ByVal XLFile As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsWorksheet
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Tests if specified workbook contains specified worksheet
      '           :
      '           :
      '           :
      '           :
      '           :
      '  Versions : 1.0 - 00/00/0000 - Adiv Abramson
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
      Dim ws As Worksheet
      '*********************************

      'Workbooks:
      '*********************************
      Dim wbData As Workbook
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

      'Constants:
      '*********************************

      '*********************************

10    IsWorksheet = False

20    If IsMissing(XLFile) Then
30       Set wbData = ThisWorkbook
40    Else
50       Set wbData = Application.Workbooks(XLFile)
60    End If

70    For Each ws In wbData.Worksheets
80       If StringCompare(String1:=ws.Name, String2:=SheetName) Then
90          IsWorksheet = True
100         Exit Function
110      End If
120   Next ws


130   Exit Function

ErrProc:
140       IsWorksheet = False
150       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsWorksheet of Module mdlUtil"
End Function

Public Sub Incr(ByRef vntValue, ByVal vntAmount As Variant)
      '---------------------------------------------------------------------------------------
      ' Procedure : Incr
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Increments/decrements referenced variable
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
      '---------------------------------------------------------------------------------------


10    On Error GoTo ErrProc

20    vntValue = vntValue + vntAmount

30    Exit Sub

ErrProc:

40        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure Incr of Module mdlUtil"

End Sub
Public Function IsActiveWorkbook() As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsActiveWorkbook
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Ensures there's an active workbook that's been saved to disk. If not, ranges and
      '           : other objects cannot be referenced
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
      Dim wbData As Workbook
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

      'Constants:
      '*********************************

      '*********************************


10    On Error GoTo ErrProc

      'get reference to active workbook
20    If Application.Workbooks.Count = 0 Then
30       msgAttention "No workbook is currently open in Excel."
40       IsActiveWorkbook = False
50       Exit Function
60    End If

70    Set wbData = ActiveWorkbook

80    If Len(wbData.Path) = 0 Then
90       msgAttention "Active workbook has not been saved to disk."
100      IsActiveWorkbook = False
110      Exit Function
120   End If



130   IsActiveWorkbook = True
         
140   Exit Function

ErrProc:
150       IsActiveWorkbook = False
160       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure IsActiveWorkbook of Module mdlUtil"
End Function

Public Function CountUniqueValues(ByRef rngData As Range) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : CountUniqueValues
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Determines number of unique values in range using standard formula
      '           : SUMPRODUCT(1/COUNTIF(range,range & "")), with & "" added to account for blank cells
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
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strAddress As String
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

      'Constants:
      '*********************************

      '*********************************


10    On Error GoTo ErrProc

20    strAddress = rngData.Address
30    CountUniqueValues = Evaluate("SUMPRODUCT(1/COUNTIF(" & strAddress & ", " & strAddress & " & """"))")

         
40    Exit Function

ErrProc:
50        CountUniqueValues = -1
60        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure CountUniqueValues of Module mdlUtil"
End Function


Public Sub DisplayAboutBox()
      '---------------------------------------------------------------------------------------
      ' Procedure : DisplayAboutBox
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Display message box with current version number and date
      '           : Coerced VersionDate to mm/dd/yyyy. For some reason it was displaying with time included
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

      'Constants:
      '*********************************

      '*********************************


10    On Error GoTo ErrProc

20    msgInfo "Version Number " & ThisWorkbook.CustomDocumentProperties("VersionNumber").Value & vbCrLf _
                & Format(ThisWorkbook.CustomDocumentProperties("VersionDate").Value, "mm/dd/yyyy")

         
30    Exit Sub

ErrProc:

40        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure DisplayAboutBox of Module mdlUtil"
End Sub

Public Sub FeatureNotAvailable(Optional ByVal ReplacementText As Variant)
      '---------------------------------------------------------------------------------------
      ' Procedure : FeatureNotAvailable
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Assign this macro to buttons whose function is not ready yet
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
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strMessage As String
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

      'Constants:
      '*********************************

      '*********************************


10    On Error GoTo ErrProc

20    If Not IsMissing(ReplacementText) Then
30       strMessage = ReplacementText
40    Else
50       strMessage = "Feature not available."
60    End If

70    msgInfo strMessage
         
80    Exit Sub

ErrProc:

90        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure FeatureNotAvailable of Module mdlUtil"
End Sub


Public Function GetRowCount(ByRef DataWorksheet As Worksheet, _
                Optional ByVal IsPivotTable As Boolean = False) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : GetRowCount
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Return number of VISIBLE rows for specified worksheet, excluding header row
      '           : Special handling required for a Pivot Table
      '           : Except for Pivot Tables, assume column A always has some data
      '           : (1.1) Must make fail safe so that it doesn't break the program if
      '           : it cannot determine the row count
      '           : (1.2) Now using updated Row Count tab with SUBTOTAL formulas
      '           : for non pivot table tabs, as it is easier and more accurate
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
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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

      '*********************************

      'Numerics:
      '*********************************
      Dim intFirstDataFieldIndex As Integer
      '*********************************

      'Worksheets:
      '*********************************
      Dim wsRowCount As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngColumnData As Range
      Dim rngPivotTableData As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objPivotTable As PivotTable
      '*********************************

      'Variants:
      '*********************************
      Dim vntRowCount As Variant
      Dim vntWorksheetNameRow As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************




10    On Error Resume Next

20    GetRowCount = 0

30    If Not HasData(DataWorksheet:=DataWorksheet) Then Exit Function
40    If IsPivotTable And DataWorksheet.PivotTables.Count = 0 Then Exit Function

50    If IsPivotTable Then
60       Set objPivotTable = DataWorksheet.PivotTables(1)
70       If objPivotTable.DataBodyRange.Rows.Count = 1 Then Exit Function
80       Set rngPivotTableData = objPivotTable.TableRange1
90       Set rngPivotTableData = rngPivotTableData.Offset(RowOffset:=2).Resize(RowSize:=rngPivotTableData.Rows.Count - 3)
         
100      intFirstDataFieldIndex = objPivotTable.RowFields.Count + 1
110      Set rngColumnData = rngPivotTableData.Columns(intFirstDataFieldIndex)
120      vntRowCount = Application.WorksheetFunction.Subtotal(103, rngColumnData)
130      If Not IsError(vntRowCount) Then GetRowCount = vntRowCount
140   Else
         '========================================================
         '<NEW CODE: 00/00/0000 Find worksheet name in column A
         'of hidden Row Count tab. Row count for this worksheet
         'will be in the adjacent column.
         '========================================================
150      Set wsRowCount = ThisWorkbook.Worksheets(WorksheetNames.RowCount)
160      If Not HasData(DataWorksheet:=wsRowCount) Then Exit Function
         
170      Set rngColumnData = wsRowCount.UsedRange.Columns(1)
180      vntWorksheetNameRow = Application.Match(DataWorksheet.Name, rngColumnData, 0)
190      If Not IsError(vntWorksheetNameRow) Then
200         vntRowCount = wsRowCount.Cells(vntWorksheetNameRow, 2).Value
210         GetRowCount = vntRowCount
220      End If 'Not IsError(vntWorksheetNameRow)
230   End If 'IsPivotTable

240   Exit Function

ErrProc:
250      GetRowCount = 0
         'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetRowCount of Module mdlUtil"
End Function


Public Function HasData(ByRef DataWorksheet As Worksheet, Optional ByVal HeaderRows As Integer = 1) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : HasData
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Tests UsedRange of specified worksheet to determine if it has any rows of data
      '           : Discount header row
      '           : (1.1) A tab containing just headers may have a UsedRange that consists of 2 rows. Can't
      '           : use count of rows in UsedRange to reliably determine whether a tab actually has data
      '           : (1.2) A completely empty worksheet will not have a UsedRange
      '           : (1.3) Revised logic to compare number of cells (excluding header) with number of blank cells.
      '           : Worksheet has no data if # blanks = # cells. But worksheet could contain cells with formulas
      '           : that evaluate to a ZLS. In this case they would not be counted as blank cells, since something
      '           : is actually in those cells. So skip over any columns containing formulas and count blanks
      '           : in cells in the other columns.
      '           : (1.4) Now using optional HeaderRows parameter to exclude header rows in excess of the default,
      '           : which is 1. This is useful for tabs having more than 1 row of headers.
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           : 1.3 - 00/00/0000 - Adiv Abramson
      '           : 1.4 - 00/00/0000 - Adiv Abramson
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

      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intCols As Integer
      Dim lngBlankCells As Long
      Dim lngTotalCells As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngData As Range
      Dim rngColumnData As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim vntResult As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    HasData = False

30    If Application.WorksheetFunction.CountA(DataWorksheet.Cells) = 0 Then
40       Exit Function
50    End If

60    Set rngData = DataWorksheet.UsedRange

      'Only header row is present
70    If rngData.Rows.Count = 1 Then
80       Exit Function
90    End If

      '========================================================
      'Get count of non blank cells on tab, excluding header
      'row(s).
      '========================================================
100   Set rngData = rngData.Offset(RowOffset:=HeaderRows).Resize(RowSize:=rngData.Rows.Count - HeaderRows)
110   intCols = rngData.Columns.Count

      '========================================================
      'NEW CODE: 00/00/0000 Revised method for determining
      'if worksheet has data.
      '========================================================
120   For i = 1 To intCols
130      Set rngColumnData = rngData.Columns(i)
140      vntResult = rngColumnData.HasFormula
         '========================================================
         'NEW CODE: 00/00/0000 If some cells have a formula while
         'others don't, the result will be Null.
         '========================================================
150      If IsNull(vntResult) Then
160         GoTo NextColumn
170      ElseIf vntResult Then
180         GoTo NextColumn
190     End If
        
200     lngTotalCells = rngColumnData.Cells.Count
210     lngBlankCells = Application.WorksheetFunction.CountBlank(rngColumnData)
220     If lngBlankCells <> lngTotalCells Then
230        HasData = True
240        Exit Function
250     End If
        
NextColumn:
260   Next i







270   Exit Function

ErrProc:
280      HasData = False
290      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure HasData of Module mdlUtil"
End Function


Public Sub SetRangeFormat(ByRef DataRange As Range, ByVal FormatString As String)
      '---------------------------------------------------------------------------------------
      ' Procedure : SetRangeFormat
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Convert values in single column range to text
      '           : (1.1) Handle case where range consists of only 1 cell
      '           : (1.2) Changed name of proc to SetRangeFormat and added 2nd parameter to
      '           : specify format string as text
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
      '           :
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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

      '*********************************

      'Numerics:
      '*********************************
      Dim lngSize As Long
      Dim i As Long
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
      Dim arValues As Variant
      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim vntTemp As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc


      '<NEW CODE:00/00/0000>
20    If DataRange.Cells.Count = 1 Then
30       With DataRange.Cells(1, 1)
40          vntTemp = CStr(.Value)
50          .ClearContents
60          .NumberFormat = FormatString
70          .Value = vntTemp
80       End With
90    Else
100      arValues = DataRange.Value
110      lngSize = UBound(arValues, 1)
120      For i = 1 To lngSize
130         arValues(i, 1) = CStr(arValues(i, 1))
140      Next i
150      With DataRange
160         .ClearContents
170         .NumberFormat = FormatString
180         .Value = arValues
190      End With
200   End If 'DataRange.Cells.Count = 1



210   Exit Sub

ErrProc:
        
220      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure SetRangeFormat of Module mdlUtil"
End Sub

Public Function HasCustomProperty(ByRef SourceWorkbook As Workbook, ByVal WorksheetName As String, _
                                  ByVal PropertyName As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : HasCustomProperty
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Determine if specified worksheet in specified workbook has specified property
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
      Dim i As Long
      Dim lngPropCount As Long
      '*********************************

      'Worksheets:
      '*********************************
      Dim ws As Worksheet
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

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    HasCustomProperty = False

      'Validate inputs
30    If SourceWorkbook Is Nothing Then
40       Exit Function
50    ElseIf WorksheetName = "" Then
60       Exit Function
70    ElseIf PropertyName = "" Then
80       Exit Function
90    ElseIf Not IsWorksheet(SourceWorkbook, WorksheetName) Then
100     Exit Function
110   End If 'SourceWorkbook Is Nothing

120   Set ws = SourceWorkbook.Worksheets(WorksheetName)
130   lngPropCount = ws.CustomProperties.Count

140   If lngPropCount = 0 Then
150      Exit Function
160   End If

170   For i = 1 To lngPropCount
180      If StringCompare(ws.CustomProperties.Item(i).Name, PropertyName) Then
190         HasCustomProperty = True
200         Exit Function
210      End If
220   Next i

         
230   Exit Function

ErrProc:
        
240      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure HasCustomProperty of Module mdlUtil"
End Function

Public Sub SetCustomProperty(ByRef SourceWorkbook As Workbook, ByVal WorksheetName As String, _
                                  ByVal PropertyName As String, ByVal PropertyValue As Variant)
      '---------------------------------------------------------------------------------------
      ' Procedure : SetCustomProperty
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Set specified value of specified custom property of specified worksheet in
      '           : specified workbook
      '           : If property doesn't exist, create it
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
      Dim i As Long
      Dim lngPropCount As Long
      '*********************************

      'Worksheets:
      '*********************************
      Dim ws As Worksheet
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

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc


      'Validate inputs
20    If SourceWorkbook Is Nothing Then
30       Exit Sub
40    ElseIf WorksheetName = "" Then
50       Exit Sub
60    ElseIf PropertyName = "" Then
70       Exit Sub
80    ElseIf Not IsWorksheet(SourceWorkbook, WorksheetName) Then
90      Exit Sub
100   End If 'SourceWorkbook Is Nothing

110   Set ws = SourceWorkbook.Worksheets(WorksheetName)
120   lngPropCount = ws.CustomProperties.Count

130   For i = 1 To lngPropCount
140      If StringCompare(ws.CustomProperties.Item(i).Name, PropertyName) Then
150         ws.CustomProperties.Item(i).Value = PropertyValue
160         Exit Sub
170      End If
180   Next i

      'Property doesn't exist so create it
190   ws.CustomProperties.Add Name:=PropertyName, Value:=PropertyValue


200   Exit Sub

ErrProc:
        
210      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure SetCustomProperty of Module mdlUtil"
End Sub

Public Function GetCustomProperty(ByRef SourceWorkbook As Workbook, ByVal WorksheetName As String, _
                                  ByVal PropertyName As String) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetCustomProperty
      ' Author    : Adiv
      ' Date      : 00/00/0000
      ' Purpose   : Return specified value of specified custom property of specified worksheet in
      '           : specified workbook'           :
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
      Dim i As Long
      Dim lngPropCount As Long
      '*********************************

      'Worksheets:
      '*********************************
      Dim ws As Worksheet
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

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    GetCustomProperty = ""

      'Validate inputs
30    If SourceWorkbook Is Nothing Then
40       Exit Function
50    ElseIf WorksheetName = "" Then
60       Exit Function
70    ElseIf PropertyName = "" Then
80       Exit Function
90    ElseIf Not IsWorksheet(SourceWorkbook, WorksheetName) Then
100     Exit Function
110   End If 'SourceWorkbook Is Nothing

120   Set ws = SourceWorkbook.Worksheets(WorksheetName)
130   lngPropCount = ws.CustomProperties.Count

140   For i = 1 To lngPropCount
150      If StringCompare(ws.CustomProperties.Item(i).Name, PropertyName) Then
160         GetCustomProperty = ws.CustomProperties.Item(i).Value
170         Exit Function
180      End If
190   Next i



200   Exit Function

ErrProc:
        
210      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetCustomProperty of Module mdlUtil"
End Function

Public Sub MoveColumn(ByRef DataWorksheet As Worksheet, ByVal SourceHeader As String, _
     ByVal TargetHeader As String, Optional MoveAfter As Boolean = False)
      '---------------------------------------------------------------------------------------
      ' Procedure : MoveColumn
      ' Author    : Adiv
      ' Date      : 00/00/0000
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
      ' Versions  : 1.0 - 00/00/0000 - Adiv Abramson
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



Public Function IsOpenWorkbook(ByVal FilePath As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsOpenWorkbook
      ' Author    : Adiv Abramson
      ' Date      : 8/14/2019
      ' Purpose   :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Change Log:---------------------------------------
      '           :Version: 1.0 Date: 8/14/2019
      '           :---------------------------------------
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :---------------------------------------
      '           :Version: 1.X Date: NN/NN/2019
      '           :---------------------------------------
      '           :
      '           :
      '           :---------------------------------------
      '           :Version: 1.X Date: NN/NN/2019
      '           :---------------------------------------
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
      Dim strWorkbookName As String
      '*********************************

      'Numerics:
      '*********************************

      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************
      Dim Wb As Workbook
      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objFS As FileSystemObject
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

20    IsOpenWorkbook = False

30    FilePath = Trim(FilePath)

40    If FilePath = "" Then Exit Function

50    Set objFS = New FileSystemObject

60    If Not objFS.FileExists(FilePath) Then Exit Function

70    strWorkbookName = objFS.GetFileName(FilePath)

80    For Each Wb In Application.Workbooks
90       IsOpenWorkbook = StringCompare(String1:=Wb.Name, String2:=strWorkbookName)
100      If IsOpenWorkbook Then Exit Function
110   Next Wb




         
120   Exit Function

ErrProc:
130      IsOpenWorkbook = False
140      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure IsOpenWorkbook of Module mdlUtil"


End Function

Public Function GetAutoFilterSettings(ByRef DataWorksheet As Worksheet) As Dictionary
      '---------------------------------------------------------------------------------------
      ' Procedure : GetAutoFilterSettings
      ' Author    : Adiv Abramson
      ' Date      : 8/15/2019
      ' Purpose   : Returns a dictionary whose keys are field names and whose items are
      '           : Array(<field pos>, <value>)
      '           : or Array(<field pos>, <value1>, <value2>) for "OR" conditions
      '           :
      '           : Use to programmatically get filter settings for a worksheet with AutoFilter
      '           : turned on and criteria applied to one or more fields. This in turn can be used
      '           : to set AutoFilter properties for a different worksheet with the same layout that
      '           : hasn't yet been filtered. Better than hard coding criteria, especially if the
      '           : source worksheet's AutoFilter settings may change from time to time
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
      ' Change Log:---------------------------------------
      '           :Version: 1.0 Date: 8/15/2019
      '           :---------------------------------------
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :---------------------------------------
      '           :Version: 1.X Date: NN/NN/2019
      '           :---------------------------------------
      '           :
      '           :
      '           :---------------------------------------
      '           :Version: 1.X Date: NN/NN/2019
      '           :---------------------------------------
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
      Dim strFilterField As String
      Dim strOutput As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intSize As Integer
      Dim i As Integer
      Dim intFieldIndex As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngHeaders As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arCriteria As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objAutoFilter As AutoFilter
      Dim objFilters As Filters
      Dim objFilter As Filter
      Dim dictOperators As New Dictionary
      Dim objTextStream As TextStream
      Dim objFS As New FileSystemObject
      Dim dictAutoFilterSettings As New Dictionary
      '*********************************

      'Variants
      '*********************************
      Dim vntCriteria1 As Variant
      Dim vntCriteria2 As Variant
      '*********************************

      'Booleans
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************


10    On Error GoTo ErrProc

20    Set GetAutoFilterSettings = Nothing

30    If DataWorksheet.AutoFilter Is Nothing Then Exit Function

40    Set objAutoFilter = DataWorksheet.AutoFilter
50    Set objFilters = objAutoFilter.Filters
60    If objFilters.Count = 0 Then Exit Function

70    Set rngHeaders = objAutoFilter.Range.Rows(1)

80    With dictOperators
90       .Item(xlAnd) = "AND"
100      .Item(xlFilterValues) = "VALUES"
110      .Item(xlOr) = "OR"
120   End With

130   intFieldIndex = 0
140   For Each objFilter In objFilters
150      Incr intFieldIndex, 1
160      strFilterField = rngHeaders.Cells(1, intFieldIndex).Value
         
170      If objFilter.On Then
180         strOutput = IIf(strOutput = "", "", strOutput & vbCrLf) & "Field: " & strFilterField
190         vntCriteria1 = objFilter.Criteria1
200         If IsArray(vntCriteria1) Then
210            intSize = UBound(vntCriteria1)
               'Remove leading "=" from filter values
               'Criteria array is 1-based
220            For i = 1 To intSize
230               vntCriteria1(i) = LTRUNC(vntCriteria1(i), 1)
240            Next i
               
250            For i = 1 To intSize
260               strOutput = IIf(strOutput = "", "", strOutput & vbCrLf) & vbTab & vntCriteria1(i)
270            Next i
               
280         Else
290            strOutput = IIf(strOutput = "", "", strOutput & vbCrLf) & vbTab & LTRUNC(vntCriteria1, 1)
            
300         End If 'IsArray(vntCriteria1)
             
310         dictAutoFilterSettings.Item(strFilterField) = Array(intFieldIndex, vntCriteria1)
            
            
320         If objFilter.Operator = xlAnd Or objFilter.Operator = xlOr Then
330            vntCriteria2 = LTRUNC(objFilter.Criteria2, 1)
340            strOutput = IIf(strOutput = "", "", strOutput & vbCrLf) _
                  & vbTab & dictOperators.Item(objFilter.Operator) & ": " & vntCriteria2
                  
350            dictAutoFilterSettings.Item(strFilterField) = Array(intFieldIndex, vntCriteria1, vntCriteria2)
360         End If 'objFilter.Operator = xlAnd Or objFilter.Operator = xlOr
               
370         strOutput = IIf(strOutput = "", "", strOutput & vbCrLf) & String(20, "=")
         
380      End If 'objFilter.On
         

390   Next objFilter

400   Set objTextStream = objFS.CreateTextFile( _
         Filename:=ThisWorkbook.Path & "\AutoFilter Settings For Worksheet " & DataWorksheet.Name & ".txt", _
         overwrite:=True)
         
410   objTextStream.Write strOutput
420   objTextStream.Close

430   msgAttention "Done!"


         
440   Exit Function

ErrProc:
450      Set GetAutoFilterSettings = Nothing
460      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetAutoFilterSettings of Module mdlUtil"


End Function

Public Function GetTable(ByVal TableName As String, Optional ByVal DataWorksheet As String = "Lists") As ListObject
      '---------------------------------------------------------------------------------------
      ' Procedure : GetTable
      ' Author    : Adiv Abramson
      ' Date      : 9/2/2019
      ' Purpose   : Reference specified table on Lists worksheet.
      '           : Enhance by providing optional parameter to specify
      '           : alternative source worksheet. {DONE}
      '           : (1.1) Make function more robust so it doesn't crash when provided with
      '           : an invalid table name.
      '           : (1.2) Remove dependency on m_wsLists and allow referencing tables on
      '           : other worksheets.
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
      'Versions   : 1.0 - 00/00/0000 - Adiv Abramson
      '           : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
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
      Dim wsData As Worksheet
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
      Dim objTable As ListObject
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

20    Set GetTable = Nothing

30    If Not IsWorksheet(SheetName:=DataWorksheet) Then Exit Function
40    Set wsData = ThisWorkbook.Worksheets(DataWorksheet)

50    For Each objTable In wsData.ListObjects
60       If UCase(objTable.Name) = UCase(TableName) Then
70          Set GetTable = objTable
80          Exit Function
90       End If
100   Next objTable

         
110   Exit Function

ErrProc:

120       Set GetTable = Nothing
130       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetTable of Module mdlUtil"


End Function


Public Function QueryList_Old(ByVal TableName As String, ByVal ReturnField As String, _
                          ParamArray QueryParams()) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : QueryList_Old
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of values from specified field of specified listobject
      '           : where optional query conditions are satisfied
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
      Dim strField As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intFields As Integer
      Dim intFieldIndex As Integer
      Dim intMatchingRows As Integer
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngColumnData As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arFields As Variant
      Dim arCriteria As Variant
      Dim arValues As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objList As ListObject
      Dim objListRow As ListRow
      Dim objListColumn As ListColumn
      '*********************************

      'Variants:
      '*********************************
      Dim vntCriteria As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnHasCriteria As Boolean
      Dim blnIsField As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    QueryList_Old = Null

30    If m_wsLists Is Nothing Then Exit Function

40    ReleaseListTabFilters

      '================================================================
      'Validate inputs
      '================================================================
50    If Not IsMissing(QueryParams) Then
60       intFields = UBound(QueryParams)
         'Array must have even number of elements.
         'Array is 0 based
70       If intFields Mod 2 = 0 Then
80          msgAttention "Invalid QueryParams list: Number of fields must match number of criteria."
90          Exit Function
100      End If 'intFields Mod 2 = 0
110      blnHasCriteria = True
120   Else
130      blnHasCriteria = False
140   End If 'Not IsMissing(QueryParams)

150   If blnHasCriteria Then
160      For i = 0 To intFields Step 2
170         AddToArray arFields, QueryParams(i), True
180      Next i
190      For i = 1 To intFields Step 2
200         AddToArray arCriteria, QueryParams(i), True
210      Next i
220   End If

230   Set objList = GetTable(TableName)

      '================================================================
      'Ensure ReturnField and any QueryFields actually exist in the
      'specified table
      '================================================================
240   blnIsField = False
250   For Each objListColumn In objList.ListColumns
260      blnIsField = StringCompare(String1:=ReturnField, String2:=objListColumn.Name)
270      If blnIsField Then Exit For
280   Next objListColumn
290   If Not blnIsField Then Exit Function

300   intFields = UBound(arFields)
310   For i = 0 To intFields
320      strField = arFields(i)
330      blnIsField = False
340      For Each objListColumn In objList.ListColumns
350         blnIsField = StringCompare(String1:=strField, String2:=objListColumn.Name)
360         If blnIsField Then Exit For
370      Next objListColumn
380      If Not blnIsField Then Exit Function
390   Next i





      '================================================================
      'Apply filters, if any
      '================================================================
400   For i = 0 To intFields
410      strField = arFields(i)
420      vntCriteria = arCriteria(i)
430      Set objListColumn = objList.ListColumns(strField)
440      Set rngColumnData = objListColumn.DataBodyRange
450      intFieldIndex = objListColumn.Index
460      objList.DataBodyRange.AutoFilter Field:=intFieldIndex, Criteria1:=vntCriteria
470   Next i

480   On Error Resume Next
490   intMatchingRows = objList.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
500   If Err <> 0 Then
510      Err.Clear
520      On Error GoTo ErrProc
530      ReleaseListTabFilters
540      Exit Function
550   End If 'Err <> 0

560   On Error GoTo ErrProc

570   Set rngColumnData = objList.ListColumns(ReturnField).DataBodyRange.SpecialCells(xlCellTypeVisible)
580   For Each rngCell In rngColumnData.Cells
590      AddToArray arValues, rngCell.Value, False
600   Next rngCell

610   ReleaseListTabFilters

620   QueryList_Old = arValues

630   Exit Function

ErrProc:
640    QueryList_Old = Null
650    ReleaseListTabFilters
660    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure QueryList_Old of Module mdlUtil"

End Function

Public Sub ReleaseListTabFilters()
      '---------------------------------------------------------------------------------------
      ' Procedure : ReleaseListTabFilters
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Clear filters from all tables on the Lists tab.
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

10    On Error GoTo ErrProc

20    If m_wsLists Is Nothing Then Exit Sub

30    For Each objList In m_wsLists.ListObjects
40       objList.AutoFilter.ShowAllData
50    Next objList


60    Exit Sub

ErrProc:

70     MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure ReleaseListTabFilters of Module mdlUtil"

End Sub

Public Function QueryDBRange(ByVal DataRange As Range, ByVal ReturnField As String, _
                             ByVal DistinctOnly As Boolean, ByVal Operator As QueryLogicOperator, _
                             ParamArray FilterSettings() As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : QueryDBRange
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of values from specified return field if specified
      '           : query fields satisfy corresponding criteria.
      '           : (1.1) Now critera can be prefixed with comparison operators, for greater
      '           : flexibility
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
      Dim strQueryField As String
      Dim strCompare As String
      Dim strFirstChar As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      Dim intSize As Integer
      Dim lngRow As Long
      Dim lngRows As Long
      '*********************************

      'Worksheets:
      '*********************************

      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngReturnField As Range
      Dim rngQueryField As Range
      Dim rngHeaders As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arResult As Variant
      Dim arRowResult As Variant
      Dim arQueryFieldCriteria As Variant
      Dim arQueryFieldIndexes As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      Dim objMatches As MatchCollection
      Dim objMatch As Match
      Dim objSubmatches As SubMatches
      '*********************************

      'Variants:
      '*********************************
      Dim vntCriteria As Variant
      Dim vntReturnFieldIndex As Variant
      Dim vntQueryFieldIndex As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnIsMatch As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    QueryDBRange = Null

30    If DataRange Is Nothing Then Exit Function
40    If DataRange.Cells.Count = 1 Then Exit Function

50    Set rngHeaders = DataRange.Rows(1)
60    vntReturnFieldIndex = Application.Match(ReturnField, rngHeaders, 0)
70    If IsError(vntReturnFieldIndex) Then Exit Function


      '================================================================
      'If upper bound is even criteria fields and values will not be
      'matched.
      '================================================================
80    intSize = UBound(FilterSettings)

90    If intSize Mod 2 = 0 Then Exit Function

100   For i = 0 To intSize Step 2
110      strQueryField = FilterSettings(i)
120      vntQueryFieldIndex = Application.Match(strQueryField, rngHeaders, 0)
130      If IsError(vntQueryFieldIndex) Then Exit Function
140      AddToArray arQueryFieldIndexes, vntQueryFieldIndex, False
150   Next i


160   For i = 1 To intSize Step 2
170      strQueryField = FilterSettings(i)
180      vntCriteria = FilterSettings(i)
190      AddToArray arQueryFieldCriteria, vntCriteria, False
200   Next i

210   If UBound(arQueryFieldIndexes) <> UBound(arQueryFieldCriteria) Then Exit Function

220   Set objRegex = New RegExp
230   With objRegex
240      .Global = False
250      .IgnoreCase = True
260      .Pattern = "^(\>\=?|\<\=?).+$"
270   End With



280   Set rngReturnField = DataRange.Columns(vntReturnFieldIndex)
290   arResult = Null
300   intSize = UBound(arQueryFieldIndexes)

310   With DataRange
320      lngRows = .Rows.Count
330      For lngRow = 2 To lngRows
340         arRowResult = Null
350         For i = 0 To intSize
360            vntQueryFieldIndex = arQueryFieldIndexes(i)
370            vntCriteria = arQueryFieldCriteria(i)
380            With objRegex
390               If .Test(vntCriteria) Then
400                  Set objMatches = .Execute(vntCriteria)
410                  Set objMatch = objMatches(0)
420                  Set objSubmatches = objMatch.SubMatches
430                  strCompare = objSubmatches(0)
440                  vntCriteria = Replace(vntCriteria, strCompare, "")
450                  If IsNumeric(vntCriteria) Then
460                     vntCriteria = CDbl(vntCriteria)
470                  End If
480               Else
490                  strCompare = "="
500               End If '.Test(vntCriteria)
510            End With
               
520            Select Case strCompare
                  Case "="
530                  blnIsMatch = (.Cells(lngRow, vntQueryFieldIndex).Value = vntCriteria)
540               Case ">"
550                  blnIsMatch = (.Cells(lngRow, vntQueryFieldIndex).Value > vntCriteria)
560               Case ">="
570                  blnIsMatch = (.Cells(lngRow, vntQueryFieldIndex).Value >= vntCriteria)
580               Case "<"
590                  blnIsMatch = (.Cells(lngRow, vntQueryFieldIndex).Value < vntCriteria)
600               Case "<="
610                  blnIsMatch = (.Cells(lngRow, vntQueryFieldIndex).Value <= vntCriteria)
620            End Select
               
630            AddToArray arRowResult, blnIsMatch, False
640         Next i
650         If Operator = LogicalAND Then
660            blnIsMatch = Application.WorksheetFunction.And(arRowResult)
670         Else
680            blnIsMatch = Application.WorksheetFunction.Or(arRowResult)
690         End If
            
700         If blnIsMatch Then
710            AddToArray arResult, .Cells(lngRow, vntReturnFieldIndex).Value, DistinctOnly
720         End If
730      Next lngRow
740   End With


750   QueryDBRange = arResult


760   Exit Function

ErrProc:
770      QueryDBRange = Null
780      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure QueryDBRange of Module mdlUtil"

End Function

Public Function StrEval(ByVal InputString As String) As Variant
'---------------------------------------------------------------------------------------
' Procedure : StrEval
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Evaluate embedded expressions in input string delimited by ${} and
'           : compose output string with expressions replaced with their
'           : values, cast as strings
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
Dim strOutput As String
Dim strExpression As String
'*********************************

'Numerics:
'*********************************
Dim i As Integer
Dim intExpressionCount As Integer
Dim intInputLength As Integer
Dim intExpressionLength As Integer
Dim intCurrentPosition As Integer
Dim intExpressionStart As Integer
Dim intExpressionEnd As Integer
Dim Value1 As Integer
Dim Value2 As Integer
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
Dim arExpressions() As Variant
'*********************************

'Objects:
'*********************************

'*********************************

'Variants:
'*********************************
Dim vntExpression As Variant
Dim vntExpressionValue As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************

On Error GoTo ErrProc

StrEval = InputString

'Validate input
If Len(InputString) = 0 Then Exit Function
If InStr(1, InputString, "${") = 0 Then Exit Function

'The variables that appear in the embedded expressions
'must have already been defined and set by the calling
'code

'For testing purposes, define those values here
Value1 = 10
Value2 = 110

'Collect the starting and ending positions of all
'embedded expressions
intExpressionCount = -1
intExpressionStart = InStr(1, InputString, "${")
intExpressionEnd = InStr(intExpressionStart + 1, InputString, "}")
intExpressionLength = intExpressionEnd - intExpressionStart - 2
While (intExpressionStart > 0) And (intExpressionEnd > 0)
   strExpression = Mid(InputString, intExpressionStart + 2, intExpressionLength)
   Debug.Print strExpression
   'Add expression and the positions it occupies in the
   'input string
   Stop
   vntExpressionValue = Application.Evaluate("=" & """" & strExpression & """")
   Mid(InputString, intExpressionStart, intExpressionLength) = vntExpressionValue
   
   intInputLength = Len(InputString)
   
   'Update search range to reflect changed input string, whose length
   'may have changed but the starting position of the current expression
   'will still be the same, so we can use it to find the starting
   'position of the next expression, if any
   intExpressionStart = InStr(intExpressionStart + 1, InputString, "${")
   intExpressionEnd = InStr(intExpressionStart + 1, InputString, "}")
Wend

   
'Now evaluate each expression






Exit Function

ErrProc:
   StrEval = ""
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure StrEval of Module mdlUtil"

End Function


Public Sub ProcessTimer(Optional ByVal ProcessName As Variant, _
                        Optional ByVal Report As Boolean = False)
      '---------------------------------------------------------------------------------------
      ' Procedure : ProcessTimer
      ' Author    : Adiv Abramson
      ' Date      : 5/25/2021
      ' Purpose   : Tracks the amount of time to open a supporting workbook.
      '           : For diagnostic purposes.
      '           : When Report is True, calculate and display the duration.
      '           : ProcessName can be a file path or an arbitrary descriptor.
      '           : GetFileName() will simply return the string passed to it if
      '           : it isn't an actual file path.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 5/25/2021 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strProcessName As String
      '*********************************

      'Numerics:
      '*********************************
      Dim lngDuration As Long
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
      Static objFS As FileSystemObject
      '*********************************

      'Variants:
      '*********************************
      Dim dteStart As Variant
      Dim dteEnd As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnIsFile As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    If objFS Is Nothing Then
30       Set objFS = New FileSystemObject
40    End If

50    If Report Then
60       If IsMissing(ProcessName) Then Exit Sub
70       dteEnd = Now
80       lngDuration = DateDiff("s", dteStart, dteEnd)
90       strProcessName = objFS.GetFileName(ProcessName)
100      blnIsFile = (strProcessName Like "*.xlsm")
110      If Not blnIsFile Then strProcessName = """" & strProcessName & """"
120      Debug.Print IIf(blnIsFile, "Opening", "Executing") & " " _
         & strProcessName & " took " & Format(lngDuration, "#,##0") & " seconds"
130   Else
140      dteStart = Now
150   End If


160   Exit Sub

ErrProc:

170      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessTimer of Module mdlUtil"
End Sub

Public Function GetSelectedItems(ByRef DataListBox As MSForms.ListBox)
      '---------------------------------------------------------------------------------------
      ' Procedure : GetSelectedItems
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Return an array of the items selected in the specified ListBox control.
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
      Dim i As Integer
      Dim intItems As Integer
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
      Dim arSelectedItems As Variant
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

20    GetSelectedItems = Null

30    With DataListBox
40       If .MultiSelect = fmMultiSelectSingle Then Exit Function
50       intItems = .ListCount
60       For i = 0 To intItems - 1
70          If .Selected(i) Then
80             AddToArray DataArray:=arSelectedItems, InputValue:=.List(i), KeepUnique:=False
90          End If
100      Next i
110   End With


120   Exit Function

ErrProc:
130      GetSelectedItems = Null
140      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetSelectedItems of Module " & MODULE_NAME
End Function

Public Sub GenerateProperties()
'---------------------------------------------------------------------------------------
' Procedure : GenerateProperties
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Generate Property Get procedures for specified table, for use
'           : in TableColumnNames module
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
Dim strTableName As String
Dim strFieldName As String
Dim strStrippedFieldName As String
Dim strProc As String
Dim strPropertyName As String
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
Dim objTable As ListObject
Dim objListColumn As ListColumn
Dim dictTableInfo As Dictionary
'*********************************

'Variants:
'*********************************
Dim vntTableName As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************

On Error GoTo ErrProc

InitGlobalVars


Set dictTableInfo = New Dictionary
With dictTableInfo
   .Item(TableNames.VisibleSheets) = "VisibleSheets_"
   .Item(TableNames.PivotTableInfo) = "PivotTableInfo_"
   .Item(TableNames.Subjects) = "Subjects_"
End With


For Each vntTableName In dictTableInfo.Keys
   Set objTable = GetTable(TableName:=vntTableName)
   For Each objListColumn In objTable.ListColumns
      strTableName = dictTableInfo.Item(vntTableName)
      strFieldName = objListColumn.Name
      strStrippedFieldName = Replace(strFieldName, " ", "")
      strPropertyName = PrintFormat("{$0}{$1}", strTableName, strStrippedFieldName)
      strProc = PrintFormat("'Public Property Get {$0} As String", strPropertyName) & vbCrLf
      strProc = strProc & PrintFormat("'{$0}{$1} = ""{$2}""", String(3, vbTab), _
                strPropertyName, strFieldName) & vbCrLf
      
      strProc = strProc & "'End Property"
      Debug.Print strProc
   Next objListColumn
Next vntTableName


'Output
'---------------------------
'Public Property Get VisibleSheets_WorksheetName As String
'   VisibleSheets_WorksheetName = "Worksheet Name"
'End Property
'Public Property Get VisibleSheets_OKToClean As String
'   VisibleSheets_OKToClean = "OK To Clean"
'End Property
'Public Property Get VisibleSheets_CountColumn As String
'   VisibleSheets_CountColumn = "Count Column"
'End Property
'Public Property Get VisibleSheets_IsPivotTable As String
'   VisibleSheets_IsPivotTable = "Is Pivot Table"
'End Property
'Public Property Get VisibleSheets_Calculation As String
'   VisibleSheets_Calculation = "Calculation"
'End Property





Exit Sub





ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GenerateProperties of Module " & MODULE_NAME
End Sub

Public Function IsTableField(ByVal FieldName As String, ByVal DataTableName As String, _
                       Optional ByVal DataWorksheetName As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsTableField
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Checks if specified field exists in specified table.
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
      Dim objTable As ListObject
      Dim objListColumn As ListColumn
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

20    IsTableField = False

      '========================================================
      'Validate inputs
      '========================================================
30    If Trim(FieldName) = "" Or Trim(DataTableName) = "" Then Exit Function
40    If IsMissing(DataWorksheetName) Then DataWorksheetName = "Lists"
50    If Not IsWorksheet(SheetName:=DataWorksheetName) Then Exit Function

60    Set objTable = GetTable(TableName:=DataTableName, DataWorksheet:=DataWorksheetName)
70    If objTable Is Nothing Then Exit Function

80    For Each objListColumn In objTable.ListColumns
90       If UCase(FieldName) = UCase(objListColumn.Name) Then
100         IsTableField = True
110         Exit Function
120      End If
130   Next objListColumn



140   Exit Function

ErrProc:
150      IsTableField = False
160      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure IsTableField of Module " & MODULE_NAME
End Function

Public Function GetNumberRange(ByVal StartNum As Long, ByVal EndNum As Long, _
                               Optional ByVal StepNum As Long = 1) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetNumberRange
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Return array of every StepNum value, starting at StartNum and ending at EndNum
      '           : Emulates Python's Range() function, sort of.
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
      Dim lngNum As Long
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
      Dim arResult As Variant
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

20    GetNumberRange = Null

      '========================================================
      'Validate parameters.
      '========================================================
      '========================================================
      'Can't use positive step when starting value is greater
      'ending value
      '========================================================
30    If StartNum > EndNum And StepNum > 0 Then Exit Function

      '========================================================
      'Can't use negative step when starting value is less than
      'ending value
      '========================================================
40    If StartNum < EndNum And StepNum < 0 Then Exit Function

      '========================================================
      'Step value must be non zero.
      '========================================================
50    If StepNum = 0 Then Exit Function

60    If StartNum = EndNum Then
70       GetNumberRange = Array(StartNum)
80       Exit Function
90    End If


100   For lngNum = StartNum To EndNum Step StepNum
110      AddToArray DataArray:=arResult, InputValue:=lngNum, KeepUnique:=False
120   Next lngNum


130   GetNumberRange = arResult

140   Exit Function

ErrProc:
150      GetNumberRange = Null
160      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetNumberRange of Module " & MODULE_NAME
End Function

Public Function GetMatchingCells(ByRef DataRange As Range, ByVal SearchValue As Variant) As Range
      '---------------------------------------------------------------------------------------
      ' Procedure : GetMatchingCells
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Finds all cells in specified range whose value matches SearchValue.
      '           : Returns a range object containing multiple cell references.
      '           : Must iterate through each Area of DataRange.
      '           : Search is case sensitive.
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
      Dim i As Integer
      Dim intArea As Integer
      Dim intAreas As Integer
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
      Dim rngArea As Range
      Dim rngMatchingCells As Range
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

20    Set GetMatchingCells = Nothing

      '========================================================
      'Validate inputs: DataRange must consist of two or more
      'cells and SearchValue cannot be Null or an empty string.
      '========================================================
30    If DataRange Is Nothing Then Exit Function
40    If DataRange.Cells.Count < 2 Then Exit Function
50    If IsNull(SearchValue) Then Exit Function
60    If SearchValue = "" Then Exit Function

      '========================================================
      'Get count of areas in DataRange in order to look for
      'matching cells in each one.
      '========================================================
70    intAreas = DataRange.Areas.Count
80    For intArea = 1 To intAreas
90       Set rngArea = DataRange.Areas(intArea)
100      For Each rngCell In rngArea.Cells
110         If rngCell.Value = SearchValue Then
120          If rngMatchingCells Is Nothing Then
130             Set rngMatchingCells = rngCell
140          Else
150             Set rngMatchingCells = Application.Union(rngMatchingCells, rngCell)
160          End If 'rngMatchingCells Is Nothing
170         End If
180      Next rngCell
190   Next intArea

200   Set GetMatchingCells = rngMatchingCells


210   Exit Function

ErrProc:
220      Set GetMatchingCells = Nothing
230      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetMatchingCells of Module " & MODULE_NAME
End Function

Public Sub LoadTablesIntoPowerQuery(Optional ByRef DataWorkbook As Workbook)
      '---------------------------------------------------------------------------------------
      ' Procedure : LoadTablesIntoPowerQuery
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Display a form with a listbox containing the names of all tables
      '           : in the workbook. Make listbox multiselect so user can select
      '           : which tables to load as connection only into Power Query.
      '           : Detects tables on hidden worksheets as well.
      '           : (1.1) Modified proc to support optionally loading tables from specified
      '           : external workbook or from current workbook if no external
      '           : workbook supplied.
      '           : IMPORTANT: If the referenced external workbook is on OneDrive,
      '           : the queries will be created but will have errors. PQ will complain
      '           : about needing an absolute path. So files that have a URL for
      '           : the path instead of a regular path on the file system will have
      '           : issues. But otherwise, this code seems to work.
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
      Dim strTableName As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intLoadedTables As Integer
      '*********************************

      'Worksheets:
      '*********************************
      Dim ws As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************

      '*********************************

      'Arrays:
      '*********************************
      Dim arAllTables As Variant
      Dim arSelectedTables As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objTable As ListObject
      '*********************************

      'Variants:
      '*********************************
      Dim arTable As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

      '========================================================
      'Cycle through all worksheets in current workbook and
      'collect the names of all of the tables in each, which
      'should be unique over the workbook.
      '========================================================
      '========================================================
      'NEW CODE: 00/00/0000 Now loading tables in external
      'workbook, if specified. Otherwise loading tables in
      'current workbook.
      '========================================================
20    If DataWorkbook Is Nothing Then
30       Set DataWorkbook = ThisWorkbook
40    End If

50    For Each ws In DataWorkbook.Worksheets
60       For Each objTable In ws.ListObjects
70          AddToArray DataArray:=arAllTables, InputValue:=objTable.Name, KeepUnique:=True
80       Next objTable
90    Next ws

100   If Not IsArray(arAllTables) Then
110      msgAttention "No tables found in this workbook."
120      Exit Sub
130   End If

      '========================================================
      'Load table names into custom form. If user didn't click
      'cancel, then at least one table has been selected.
      '========================================================
140   With frmList
150      With .lstOptions
160         .MultiSelect = fmMultiSelectMulti
170         .List = Application.WorksheetFunction.Transpose(arAllTables)
180      End With
190      .Show vbModal
200      If Not IsLoaded("frmList") Then Exit Sub
210      arSelectedTables = .SelectedListItems
220   End With

230   Unload frmList
      '========================================================
      'Create a connection only query for each selected table.
      'CreateWorkbookQueryFromTable() will not allow queries with
      'the same name to be loaded, although the same data can
      'be loaded under different names.
      '========================================================
240   intLoadedTables = 0
250   For Each arTable In arSelectedTables
260      If CreateWorkbookQueryFromTable(DataWorkbook:=DataWorkbook, _
                                         TableName:=arTable, _
                                         Description:="Automatically loaded") Then
270         Incr intLoadedTables, 1
280      End If
         
290   Next arTable

300   msgInfo PrintFormat2("{$count} tables have been loaded into Power Query as connection only objects.", "count", intLoadedTables)


310   Exit Sub

ErrProc:

320      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure LoadTablesIntoPowerQuery of Module " & MODULE_NAME
End Sub

Public Function GetMostFrequentValue(ByVal DataArray As Variant) As Dictionary
      '---------------------------------------------------------------------------------------
      ' Procedure : GetMostFrequentValue
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Given a 1D array of numbers or strings, return a dictionary
      '           : whose keys are the most frequently occurring values and whose
      '           : items are the corresponding counts. If an array has just one value
      '           : that occurs more often than all other values in the array, the output
      '           : dictionary will have a count of 1. If two or more array elements occur
      '           : most frequently, e.g. 2 and 45 occur 27 times each in the input array,
      '           : the dictionary returned will have two keys: 2 and 45, each of whose
      '           : items will be 27.
      '           : If each element in the input array occurs only once, return Nothing.
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
      Dim lngCount As Long
      Dim lngMaxCount As Long
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
      Dim dictStats As Dictionary
      Dim dictOutput As Dictionary
      '*********************************

      'Variants:
      '*********************************
      Dim vntItem As Variant
      Dim vntKey As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    Set GetMostFrequentValue = Nothing

30    If Not IsArray(DataArray) Then Exit Function

      '========================================================
      'An empty array could be passed to the function.
      '========================================================
40    If UBound(DataArray) = -1 Then Exit Function

50    Set dictStats = New Dictionary

      '========================================================
      'Populate dictionary with the frequencies of the array
      'elements.
      '========================================================
60    For Each vntItem In DataArray
70       If Not dictStats.Exists(vntItem) Then
80          dictStats.Item(vntItem) = 1
90       Else
100         lngCount = dictStats.Item(vntItem)
110         Incr lngCount, 1
120         dictStats.Item(vntItem) = lngCount
130      End If ' Not dictStats.Exists(vntItem)
140   Next vntItem

      '========================================================
      'Check the maximum frequency. If it's 1, there are no
      'duplicates.
      '========================================================
150   lngMaxCount = Application.WorksheetFunction.Max(dictStats.Items)
160   If lngMaxCount = 1 Then Exit Function

      '========================================================
      'Populate output dictionary with the most frequently
      'occurring values.
      '========================================================
170   Set dictOutput = New Dictionary
180   For Each vntKey In dictStats.Keys
190      If dictStats.Item(vntKey) = lngMaxCount Then
200         dictOutput.Item(vntKey) = lngMaxCount
210      End If
220   Next vntKey

230   Set GetMostFrequentValue = dictOutput


240   Exit Function

ErrProc:
250      Set GetMostFrequentValue = Nothing
260      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetMostFrequentValue of Module " & MODULE_NAME
End Function

Public Sub SortSheets(Optional ByRef DataWorkbook As Workbook, Optional ByVal SortOrder As Variant = 1)
      '---------------------------------------------------------------------------------------
      ' Procedure : SortSheets
      ' Author    : Adiv Abramson
      ' Date      : 4/15/2024
      ' Purpose   : Sort worksheets of specified workbook (default is active workbook)
      '           : in ascending (default) or descending order
      '           : Index property of worksheet is read only.
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
      ' Versions  : 1.0 - 4/15/2024 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim wsName As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intWorksheetNameIndex As Integer
      Dim intWorksheetIndex As Integer
      Dim intWorksheetNames As Integer
      Dim intWorksheets As Integer
      Dim intFirstCharCode As Integer
      '*********************************

      'Worksheets:
      '*********************************
      Dim ws As Worksheet
      Dim wsNext As Worksheet
      Dim wsTemp As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngSortedSheetNames As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arWorksheetNames As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objSheets As ListObject
      Dim objListRow As ListRow
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const TEMP_WS_NAME = ".___"
      Const SORTED_SHEET_NAMES As String = "SortedSheetNames"
      '*********************************

10    On Error GoTo ErrProc


      '========================================================
      ' Validate inputs
      '========================================================
20    If DataWorkbook Is Nothing Then
30       Set DataWorkbook = ActiveWorkbook
40    End If

50    If SortOrder < 1 Or SortOrder > 2 Then Exit Sub


      '========================================================
      ' Collect names of visible worksheets.
      ' Clear Sheets table and populate.
      '========================================================
60    Set objSheets = GetTable(TableName:="Sheets")
70    With objSheets
80       If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
        
90       For Each ws In DataWorkbook.Worksheets
100         If ws.Visible = xlSheetVisible Then
110            Set objListRow = .ListRows.Add
120            objListRow.Range.Cells(1, 1).Value = ws.Name
130         End If
140      Next ws
150   End With



      '========================================================
      ' WORKING WITH DYNAMIC ARRAYS
      ' Range("SortedSheetNames").HasSpill = True
      ' Range("SortedSheetNames").Formula2 = "=SORT(Sheets[Name],1,-1)"
      ' Range("SortedSheetNames").SpillingToRange.Address = "$X$1:$X$8"
      '
      '
      '
      ' Range("SortedSheetNames#").HasSpill = True
      ' Range("SortedSheetNames#").Formula = ERROR
      ' Range("SortedSheetNames#").Cells(1,1).Formula2 = "=SORT(Sheets[Name],1,-1)"
      ' Range("SortedSheetNames#").SpillingToRange.Address = ERROR
      ' Range("SortedSheetNames#").Cells(1,1).SpillingToRange.Address = "$X$1:$X$8"
      '
      ' Seems like using the plain name without the final "#" is
      ' easier.
      '
      '========================================================


      '========================================================
      ' To reference the parent (first) cell of a dynamic range,
      ' you can use Range("...#").
      '
      ' You can reference the spilled range with
      ' Range("...").SpillingToRange.
      '========================================================
160   Set rngSortedSheetNames = Range(SORTED_SHEET_NAMES)

      '========================================================
      ' For ascending sort order, 3rd argument of SORT() is 1.
      ' For descending sort order, 3rd argument of SORT() is -1
      ' To change the Formula of a dynamic range, use Formula2
      ' instead of Formula property.
      '========================================================
170   rngSortedSheetNames.Formula2 = IIf(SortOrder = 1, "=SORT(Sheets[Name],1,1)", "=SORT(Sheets[Name],1,-1)")


      '========================================================
      ' Create a temporary worksheet whose name is alphabetically
      ' less than any existing name.
      '========================================================
180   If IsWorksheet(SheetName:=TEMP_WS_NAME) Then
190      DataWorkbook.Worksheets(TEMP_WS_NAME).Delete
200   End If

210   Set wsTemp = DataWorkbook.Worksheets.Add
220   wsTemp.Name = TEMP_WS_NAME

      '========================================================
      ' wsNext will be the sheet after which the worksheet
      ' currently being processed will be moved.
      ' After moving the current worksheet, must update the ws
      ' variable so that wsNext can reference it in the next
      ' iteration.
      '========================================================
230   Set wsNext = Nothing

240   For Each rngCell In rngSortedSheetNames.SpillingToRange.Cells
250      wsName = rngCell.Value
260      Set ws = ThisWorkbook.Worksheets(wsName)
270      If wsNext Is Nothing Then
280         ws.Move After:=wsTemp
290         Set ws = ThisWorkbook.Worksheets(wsName)
300      Else
310         ws.Move After:=wsNext
320         Set ws = ThisWorkbook.Worksheets(wsName)
330      End If
         
340      Set wsNext = ws
         
350   Next rngCell

      '========================================================
      ' Delete the temporary worksheet.
      '========================================================
360   Application.DisplayAlerts = False
370   DataWorkbook.Worksheets(TEMP_WS_NAME).Delete
380   Application.DisplayAlerts = True



390   Exit Sub

ErrProc:
400      Application.DisplayAlerts = True
410      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure SortSheets of Module " & MODULE_NAME
End Sub
