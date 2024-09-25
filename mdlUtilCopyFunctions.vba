'---------------------------------------------------------------------------------------
' Module    : mdlUtilCopyFunctions
' Author    : Adiv Abramson
' Date      : 4/10/2024
' Purpose   : Contains some copy and paste related functions moved from mdlUtil
'           :
'           :
'           :
'           :
'           :
'           :
' Versions  : 1.0 - 4/10/2024 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Option Explicit
Option Private Module
Private Const MODULE_NAME = "mdlUtilCopyFunctions"



Public Sub CopyFromTable(ByVal SourceTable As String)
'---------------------------------------------------------------------------------------
' Procedure : CopyFromTable
' Author    : Adiv Abramson
' Date      : 4/24/2024
' Purpose   : Copies the contents of the DataBodyRange of the specifield table to the Clipboard, which can later be pasted back into the VBA Editor.
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
' Versions  : 1.0 - 4/24/2024 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

'Strings:
'*********************************
Dim strDataWorksheet As String
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
Dim objSource As ListObject
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
      
      '========================================================
      'Validate inputs
      '========================================================
20    strDataWorksheet = WorksheetNames.CodeHelper
30    Set objSource = GetTable(TableName:=SourceTable, DataWorksheet:=strDataWorksheet)
40    If objSource Is Nothing Then Exit Sub
50    If objSource.DataBodyRange Is Nothing Then Exit Sub
60    objSource.DataBodyRange.Copy


Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure CopyFromTable of Module " & MODULE_NAME
End Sub

Sub PasteInPlace()
      '---------------------------------------------------------------------------------------
      ' Procedure : PasteInPlace
      ' Author    : Adiv Abramson
      ' Date      : 4/10/2024
      ' Purpose   : Replace selection, which may contain formulas, with values and
      '           : source formatting.
      '           : Can be invoked by Ribbon XML <contextMenus> or CommandBars
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
      ' Versions  : 1.0 - 4/10/2024 - Adiv Abramson
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
      Dim rngData As Range
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


20    Set rngData = Selection
30    rngData.Copy
40    rngData.PasteSpecial xlPasteValuesAndNumberFormats
50    Application.CutCopyMode = False


60    Exit Sub

ErrProc:

70       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure PasteInPlace of Module " & MODULE_NAME
End Sub



Public Function CopyTabData2(ByRef SourceWorksheet As Worksheet, _
       ByRef TargetWorksheet As Worksheet, ByVal DeleteSourceRecords As Boolean, _
       Optional ByVal SourceHeaderRow As Integer = 1, _
       Optional ByVal TargetHeaderRow As Integer = 1, _
       Optional TargetRow As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyTabData2
      ' Author    : Adiv
      ' Date      : 06/03/2020
      ' Purpose   : A faster (I hope) version of CopyTabData(), for use when generating a
      '           : Change Report, when the new Month fields have been added.
      '           : The major change is instead of copying to the Clipboard worksheet, copy visible cells
      '           : of source range to an array and paste array at target location cell (resize to accommodate)
      '           ' the array.
      '           : Convenience proc for copying possibly filtered data from source tab to target tab
      '           : Calling proc must ensure both worksheets are in a proper state
      '           : Inspect headers of target worksheet and find matching column headers in source worksheet and then
      '           : copy data to end of existing data on target worksheet. Respect any filter in place on the source
      '           : worksheet.
      '           : Optionally exclude field names listed
      '           : Added new parameter to control deletion of source records
      '           : Fixed bug in constructing address for filtered dataset
      '           : (1.2) Use GetRangeAddress() to obtain the full address of a multi-area range, whose .Address
      '           : property would exceed Excel limit of 255 characters. Hope it works as an argument of the
      '           : Range function
      '           : (1.3) Can't replace "$A$1" from strAddress with ZLS unless there is more than 1 Area
      '           : and the # of cells in Areas(1) is more than 1
      '           : (1.4) Fixed bug in copying filtered data where source and target ranges were not of same
      '           : size
      '           : (1.5) Copy visible rows to new Clipboard worksheet and then copy all but header to
      '           : target worksheet, eliminating complex and error prone parsing of the address
      '           : (1.6) Changed method of adding to dictionary to avoid duplicate key errors
      '           : (1.7) Ensure Source header and Target header ranges are not empty
      '           : Must inspect entire first row, not just constants in first row, since header range
      '           : can have more than one area. If that's the case, code will not be able to locate
      '           : all headers since Match() works only with first area
      '           : Don't use CurrentRegion to reference data on Clipboard tab to be copied to Target
      '           : worksheet since that range may have some blank cells. Instead, construct reference
      '           : to range to be copied to Target worksheet manually
      '           : (1.8) Now including default values for Source and Target worksheet header row
      '           : Removed arExclude. Never used much in any case
      '           : (1.9) Use GetLastDataRow() for more concise code
      '           : (2.0) Override the target row normally determined with optional TargetRow
      '           : parameter. If it exists, proc will write data starting at that row of the
      '           : target worksheet, regardless of what the actual last row of data is on the
      '           : target worksheet. This feature can be useful if data is being copied from two
      '           : different source worksheets and we need for the data to align on the target
      '           : worksheet
      '           : (2.1) Testing if SpecialCells() for visible cells fails and if so, go to next header.
      '           : Testing if last data row on Clipboard is 0, such as when source column is empty
      '           : except for header. In this case last data row on Clipboard tab will be 0 since
      '           : the header copied from the source is always deleted. Seems this scenario is
      '           : occurring when the source column has only a header label.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 06/17/2018 - Adiv Abramson
      '           : 1.1 - 06/19/2018 - Adiv Abramson
      '           : 1.2 - 06/20/2018 - Adiv Abramson
      '           : 1.3 - 06/26/2018 - Adiv Abramson
      '           : 1.4 - 06/29/2018 - Adiv Abramson
      '           : 1.5 - 07/25/2018 - Adiv Abramson
      '           : 1.6 - 07/31/2018 - Adiv Abramson
      '           : 1.7 - 08/13/2018 - Adiv Abramson
      '           : 1.8 - 08/20/2019 - Adiv Abramson
      '           : 1.9 - 08/21/2019 - Adiv Abramson
      '           : 2.0 - 08/26/2019 - Adiv Abramson
      '           : 2.1 - 10/11/2019 - Adiv Abramson
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
      Dim strTargetHeader As String
      Dim strTargetHeaderCol As String
      Dim strSourceHeaderCol As String
      '*********************************

      'Numerics:
      '*********************************
      Dim lngSourceLastDataRow As Long
      Dim lngTargetRow As Long

      Dim intValues As Integer
      Dim intAreas As Integer
      Dim intArea As Integer
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
      Dim rngTargetCell As Range
      Dim rngSourceHeaders As Range
      Dim rngTargetHeaders As Range
      Dim rngColumnData As Range
      Dim rngArea As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arSingleAreaValues As Variant
      Dim arMultiAreaValues() As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim dictTargetHeaders As Variant
      '*********************************

      'Variants:
      '*********************************
      Dim vntKey As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    CopyTabData2 = False

30    Application.ScreenUpdating = False

      '<NEW CODE:08/21/2019>
40    Set rngTargetHeaders = TargetWorksheet.Rows(TargetHeaderRow)

      '<NEW CODE:08/13/2018> Ensure worksheet has column headings in row 1
50    If Application.WorksheetFunction.CountA(rngTargetHeaders) = 0 Then
60       msgAttention TargetWorksheet & " does not have any column headings in row " & TargetHeaderRow & "."
70       Exit Function
80    Else
90       Set rngTargetHeaders = TargetWorksheet.Rows(TargetHeaderRow).SpecialCells(xlCellTypeConstants)
100   End If

      '<NEW CODE:08/21/2019>
110   Set rngSourceHeaders = SourceWorksheet.Rows(SourceHeaderRow)
      '<NEW CODE:08/13/2018> Ensure worksheet has column headings in row 1
120   If Application.WorksheetFunction.CountA(rngSourceHeaders) = 0 Then
130      msgAttention SourceWorksheet & " does not have any column headings in row " & SourceHeaderRow & "."
140      Exit Function
150   Else
160      Set rngSourceHeaders = SourceWorksheet.Rows(SourceHeaderRow).SpecialCells(xlCellTypeConstants)
170   End If


      'Get list of headers on Target worksheet
180   Set dictTargetHeaders = New Dictionary
190   For Each rngCell In rngTargetHeaders
200      strTargetHeader = rngCell.Value
         '<NEW CODE:08/13/2018>
210      strTargetHeaderCol = GetHeaderInfo(TargetWorksheet, strTargetHeader, HeaderInfoColumn)
         '<NEW CODE:07/31/2018> Avoid duplicate key errors
220      dictTargetHeaders.Item(strTargetHeader) = strTargetHeaderCol
230   Next rngCell


      'For each column header in target worksheet, find matching header in selected
      'Source tab
      '<NEW CODE:08/26/2019>
240   If IsMissing(TargetRow) Then
         'Determine target row normally
250      lngTargetRow = Application.WorksheetFunction.Max(TargetHeaderRow + 1, GetLastDataRow(DataWorksheet:=TargetWorksheet) + 1)
260   Else
270      lngTargetRow = TargetRow
280   End If 'IsMissing(TargetRow)

290   lngSourceLastDataRow = GetLastDataRow(DataWorksheet:=SourceWorksheet)

300   For Each vntKey In dictTargetHeaders
310      strTargetHeader = vntKey
320      strTargetHeaderCol = dictTargetHeaders.Item(vntKey)
         '<NEW CODE:08/13/2018>
330      strSourceHeaderCol = GetHeaderInfo(SourceWorksheet, strTargetHeader, HeaderInfoColumn, HeaderRow:=SourceHeaderRow)
340      If strSourceHeaderCol <> "" Then
            
350         If lngSourceLastDataRow >= SourceHeaderRow + 1 Then
               '========================================================
               '<NEW CODE: 06/03/2020 Copy arrays of data and avoid
               'using the Clipboard worksheet altogether
               '========================================================
                
               '========================================================
               '<NEW CODE: 10/11/2019 For some situation(s) the following
               'line fails. If so must go to next header
               '========================================================
360             On Error Resume Next
370             Set rngColumnData = GetSpecialCells(SourceWorksheet.Range(strSourceHeaderCol & SourceHeaderRow + 1 _
                                    & ":" & strSourceHeaderCol & lngSourceLastDataRow), xlCellTypeVisible)
                
380             If Err <> 0 Then
390                Err.Clear
400                On Error GoTo ErrProc
410                GoTo NextHeader
420             End If 'Err <> 0
                
430             On Error GoTo ErrProc
                
440             intAreas = rngColumnData.Areas.Count
450             If intAreas = 1 Then
460                Set rngArea = rngColumnData
470                If rngArea.Cells.Count > 1 Then
480                   arSingleAreaValues = rngColumnData.Value
490                   intValues = UBound(arSingleAreaValues, 1)
500                Else
510                   arSingleAreaValues = rngColumnData.Value
520                   intValues = 1
530                End If 'rngArea.Cells.Count > 1
540                Set rngTargetCell = TargetWorksheet.Range(strTargetHeaderCol & lngTargetRow).Resize(RowSize:=intValues)
550                rngTargetCell.Value = arSingleAreaValues
560             Else
570                intValues = -1
580                For intArea = 1 To intAreas
590                   Set rngArea = rngColumnData.Areas(intArea)
600                   For Each rngCell In rngArea.Cells
610                      Incr intValues, 1
620                      ReDim Preserve arMultiAreaValues(intValues)
630                      arMultiAreaValues(intValues) = rngCell.Value
640                   Next rngCell
650                Next intArea
                   
660                Incr intValues, 1
670                Set rngTargetCell = TargetWorksheet.Range(strTargetHeaderCol & lngTargetRow).Resize(RowSize:=intValues)
680                rngTargetCell.Value = Application.WorksheetFunction.Transpose(arMultiAreaValues)
690             End If 'intAreas = 1
700          End If 'lngSourceLastDataRow >= SourceHeaderRow + 1
710      End If 'strSourceHeaderCol <> ""
NextHeader:
720   Next vntKey


      '<NEW CODE:06/19/2018> If specified, delete source tab rows
      'But don't delete the header
730   If DeleteSourceRecords Then
740      If Not rngColumnData Is Nothing Then
750         SourceWorksheet.Activate
760         lngSourceLastDataRow = SourceWorksheet.Range("A" & LAST_WORKSHEET_ROW).End(xlUp).Row
770         If lngSourceLastDataRow >= 2 Then
780            With SourceWorksheet
790               .Range("A2:A" & lngSourceLastDataRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
800            End With
810         End If 'lngSourceLastDataRow >= 2
820      End If 'Not rngColumnData Is Nothing
830   End If 'DeleteSourceRecords

840   TargetWorksheet.Activate
850   TargetWorksheet.Columns.AutoFit

860   Application.ScreenUpdating = True

870   CopyTabData2 = True


880   Exit Function

ErrProc:
890      CopyTabData2 = False
900      Application.ScreenUpdating = True
910      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure CopyTabData2 of Module mdlUtil"
End Function






Public Function CopyTabData(ByRef SourceWorksheet As Worksheet, _
       ByRef TargetWorksheet As Worksheet, ByVal DeleteSourceRecords As Boolean, ParamArray arExclude() As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyTabData
      ' Author    : Adiv
      ' Date      : 06/17/2018
      ' Purpose   : Convenience proc for copying possibly filtered data from source tab to target tab
      '           : Call proc must ensure both worksheets are in a proper state
      '           : Inspect headers of target worksheet and find matching column headers in source worksheet and then
      '           : copy data to end of existing data on target worksheet. Respect any filter in place on the source
      '           : worksheet.
      '           : Optionally exclude field names listed
      '           : Added new parameter to control deletion of source records
      '           : Fixed bug in constructing address for filtered dataset
      '           : (1.2) Use GetRangeAddress() to obtain the full address of a multi-area range, whose .Address
      '           : property would exceed Excel limit of 255 characters. Hope it works as an argument of the
      '           : Range function
      '           : (1.3) Can't replace "$A$1" from strAddress with ZLS unless there is more than 1 Area
      '           : and the # of cells in Areas(1) is more than 1
      '           : (1.4) Fixed bug in copying filtered data where source and target ranges were not of same
      '           : size
      '           : (1.5) Copy visible rows to new Clipboard worksheet and then copy all but header to
      '           : target worksheet, eliminating complex and error prone parsing of the address
      '           : (1.6) Changed method of adding to dictionary to avoid duplicate key errors
      '           : (1.7) Ensure Source header and Target header ranges are not empty
      '           : Must inspect entire first row, not just constants in first row, since header range
      '           : can have more than one area. If that's the case, code will not be able to locate
      '           : all headers since Match() works only with first area
      '           : Don't use CurrentRegion to reference data on Clipboard tab to be copied to Target
      '           : worksheet since that range may have some blank cells. Instead, construct reference
      '           : to range to be copied to Target worksheet manually
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
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 06/17/2018 - Adiv Abramson
      '           : 1.1 - 06/19/2018 - Adiv Abramson
      '           : 1.2 - 06/20/2018 - Adiv Abramson
      '           : 1.3 - 06/26/2018 - Adiv Abramson
      '           : 1.4 - 06/29/2018 - Adiv Abramson
      '           : 1.5 - 07/25/2018 - Adiv Abramson
      '           : 1.6 - 07/31/2018 - Adiv Abramson
      '           : 1.7 - 08/13/2018 - Adiv Abramson
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
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strTargetHeader As String
      Dim strTargetHeaderCol As String
      Dim strSourceHeaderCol As String
      '*********************************

      'Numerics:
      '*********************************
      Dim lngSourceLastDataRow As Long
      Dim lngTargetRow As Long
      Dim lngClipboard_LastDataRow As Long
      '*********************************

      'Worksheets:
      '*********************************
      Dim wsClipboard As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngCell As Range
      Dim rngTargetCell As Range
      Dim rngSourceHeaders As Range
      Dim rngTargetHeaders As Range
      Dim rngColumnData As Range
      Dim rngClipboardData As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arTargetHeaders As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim dictTargetHeaders As Variant
      '*********************************

      'Variants:
      '*********************************
      Dim vntKey As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    CopyTabData = False

30    Application.ScreenUpdating = False

40    Set rngTargetHeaders = TargetWorksheet.Rows(1)
      '<NEW CODE:08/13/2018> Ensure worksheet has column headings in row 1
50    If Application.WorksheetFunction.CountA(rngTargetHeaders) = 0 Then
60       msgAttention TargetWorksheet & " does not have any column headings in row 1."
70       Exit Function
80    Else
90       Set rngTargetHeaders = TargetWorksheet.Rows(1).SpecialCells(xlCellTypeConstants)
100   End If

110   Set rngSourceHeaders = SourceWorksheet.Rows(1)
      '<NEW CODE:08/13/2018> Ensure worksheet has column headings in row 1
120   If Application.WorksheetFunction.CountA(rngSourceHeaders) = 0 Then
130      msgAttention SourceWorksheet & " does not have any column headings in row 1."
140      Exit Function
150   Else
160      Set rngSourceHeaders = SourceWorksheet.Rows(1).SpecialCells(xlCellTypeConstants)
170   End If

      '<NEW CODE:07/25/2018>
180   Set wsClipboard = ThisWorkbook.Worksheets(WorksheetNames.Clipboard)
190   wsClipboard.UsedRange.ClearContents

200   Set dictTargetHeaders = New Dictionary

      'Populate dictionary with header/column letter pairs for the non excluded target headers
210   If Not IsMissing(arExclude) Then
         'Can't use ParamArray as argument in Match()
220      arTargetHeaders = arExclude
         '<NEW CODE:06/17/2018> Populate an array of just the headers to be copied
230      For Each rngCell In rngTargetHeaders
240         strTargetHeader = rngCell.Value
            '<NEW CODE:08/13/2018>
250         strTargetHeaderCol = GetHeaderInfo(TargetWorksheet, strTargetHeader, HeaderInfoColumn)
            'If this header isn't in the list of excluded headers, add it to the list of
            'target headers
260         If strTargetHeaderCol <> "" Then
               '<NEW CODE:07/31/2018> Avoid duplicate key errors
270            dictTargetHeaders.Item(strTargetHeader) = strTargetHeaderCol
280         End If
290      Next rngCell
300   Else
310      For Each rngCell In rngTargetHeaders
320         strTargetHeader = rngCell.Value
            '<NEW CODE:08/13/2018>
330         strTargetHeaderCol = GetHeaderInfo(TargetWorksheet, strTargetHeader, HeaderInfoColumn)
            '<NEW CODE:07/31/2018> Avoid duplicate key errors
340         dictTargetHeaders.Item(strTargetHeader) = strTargetHeaderCol
350      Next rngCell
360   End If ' Not IsMissing(arExclude)


      'For each column header in target worksheet, find matching header in selected
      'Source tab
370   lngTargetRow = Application.WorksheetFunction.Max(2, TargetWorksheet.Range("A" _
                     & LAST_WORKSHEET_ROW).End(xlUp).Row + 1)
380   For Each vntKey In dictTargetHeaders
390      strTargetHeader = vntKey
400      strTargetHeaderCol = dictTargetHeaders.Item(vntKey)
         '<NEW CODE:08/13/2018>
410      strSourceHeaderCol = GetHeaderInfo(SourceWorksheet, strTargetHeader, HeaderInfoColumn)
420      If strSourceHeaderCol <> "" Then
430         lngSourceLastDataRow = SourceWorksheet.Range(strSourceHeaderCol & LAST_WORKSHEET_ROW).End(xlUp).Row
440         If lngSourceLastDataRow >= 2 Then
               '<NEW CODE:07/25/2018> Copy this column to the Clipboard worksheet and from there to the target
               'worksheet
450             wsClipboard.UsedRange.ClearContents
460             Set rngColumnData = SourceWorksheet.Range(strSourceHeaderCol & 1 _
                                    & ":" & strSourceHeaderCol & lngSourceLastDataRow).SpecialCells(xlCellTypeVisible)
470             rngColumnData.Copy Destination:=wsClipboard.Range("A1")
480             Application.CutCopyMode = False
                'Delete header
490             wsClipboard.Range("A1").EntireRow.Delete
                '<NEW CODE:08/13/2018> Can't always use CurrentRegion because column may have some blank cells
500             lngClipboard_LastDataRow = wsClipboard.Range("A" & LAST_WORKSHEET_ROW).End(xlUp).Row
510             Set rngClipboardData = wsClipboard.Range("A1:A" & lngClipboard_LastDataRow)
520             Set rngTargetCell = TargetWorksheet.Range(strTargetHeaderCol & lngTargetRow)
530             rngClipboardData.Copy Destination:=rngTargetCell
540             Application.CutCopyMode = False
550          End If 'lngSourceLastDataRow >= 2
560      End If 'strSourceHeaderCol <> ""
NextHeader:
570   Next vntKey


      '<NEW CODE:06/19/2018> If specified, delete source tab rows
      'But don't delete the header
580   If DeleteSourceRecords Then
590      If Not rngColumnData Is Nothing Then
600         SourceWorksheet.Activate
610         lngSourceLastDataRow = SourceWorksheet.Range("A" & LAST_WORKSHEET_ROW).End(xlUp).Row
620         If lngSourceLastDataRow >= 2 Then
630            With SourceWorksheet
640               .Range("A2:A" & lngSourceLastDataRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
650            End With
660         End If 'lngSourceLastDataRow >= 2
670      End If 'Not rngColumnData Is Nothing
680   End If 'DeleteSourceRecords

690   TargetWorksheet.Activate
700   TargetWorksheet.Columns.AutoFit
710   Application.ScreenUpdating = True

720   CopyTabData = True


730   Exit Function

ErrProc:
740      CopyTabData = False
750      Application.ScreenUpdating = True
760      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure CopyTabData of Module mdlUtil"
End Function



Public Function CopyColumns_ExpandRanges(ByVal ColumnSpecs As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyColumns_ExpandRanges
      ' Author    : Adiv Abramson
      ' Date      : 03/11/2021
      ' Purpose   : Replace elements specifying a valid range e.g. "A:K", with individual
      '           : column names. Return the updated and expanded array. If no elements contain
      '           : a range or if an invalid range is specified, return Null. Later on
      '           : CopyColumns_ValidateColumnSpecs() will reject Null values and terminate the copying
      '           : process.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 03/11/2021 - Adiv Abramson
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
      Dim intCol As Integer
      Dim intCols As Integer
      Dim intSize As Integer
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
      Dim rngCell As Range
      Dim rngTest As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arRange As Variant
      Dim arExpanded As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      '*********************************

      'Variants:
      '*********************************
      Dim vntColumn As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const MAX_COLUMNS = 100
      '*********************************

10    On Error GoTo ErrProc

20    CopyColumns_ExpandRanges = Null

30    If Not IsArray(ColumnSpecs) Then Exit Function
40    If UBound(ColumnSpecs) > MAX_COLUMNS Then
50       CopyColumns_ExpandRanges = ColumnSpecs
60       Exit Function
70    End If 'UBound(ColumnSpecs) > MAX_COLUMNS

      '========================================================
      'Use Data tab to validate the ranges since it doesn't
      'have any merged cells in row 1.
      '========================================================
80    Set wsData = ThisWorkbook.Worksheets(WorksheetNames.Data)

      '========================================================
      'Set up regex to validate range specifications
      '========================================================
90    Set objRegex = New RegExp
100   With objRegex
110      .Global = False
120      .IgnoreCase = True
130      .Pattern = "^[A-Z]{1,3}:[A-Z]{1,3}$"
140   End With

      '========================================================
      'Expand any embedded range specifications
      '========================================================
150   For Each vntColumn In ColumnSpecs
160      If Trim(vntColumn) = "" Then
170         msgAttention "Invalid (blank) column range specification: " & vntColumn & "."
180         Exit Function
190      End If
         
200      If InStr(1, vntColumn, ":") > 0 Then
            '========================================================
            'Specified range must look like <col>:<col>
            '========================================================
210         If Not objRegex.Test(vntColumn) Then
220            msgAttention "Invalid column range specified: " & vntColumn & "."
230            Exit Function
240         End If 'Not objRegex.Test(vntColumn)
            '========================================================
            'A backwards specification ("X:A") is technically valid
            'per Excel but enforce normal order
            '========================================================
250         arRange = Split(vntColumn, ":")
260         With wsData
270            If .Range(arRange(0) & 1).Column > .Range(arRange(1) & 1).Column Then
280               msgAttention "Invalid column range specified: " & vntColumn & "."
290               Exit Function
300            End If
310         End With
            
            '========================================================
            'Expand the original column array with the column number
            'of the range specification.
            '========================================================
320         Set rngTest = wsData.Range(arRange(0) & "1:" & arRange(1) & "1")
330         For Each rngCell In rngTest.Cells
340            AddToArray DataArray:=arExpanded, InputValue:=rngCell.Column, KeepUnique:=False
350         Next rngCell
360      Else
370         AddToArray DataArray:=arExpanded, InputValue:=vntColumn, KeepUnique:=False
380      End If 'InStr(1, vntColumn, ":") > 0
390   Next vntColumn
       
400   CopyColumns_ExpandRanges = arExpanded

         
410   Exit Function

ErrProc:
420      CopyColumns_ExpandRanges = Null
430      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of  procedure CopyColumns_ExpandRanges of Module " & MODULE_NAME
End Function

Public Function CopyColumns_ValidateColumnSpecs(ByVal SourceColumns As Variant, _
                                                ByVal TargetColumns As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyColumns_ValidateColumnSpecs
      ' Author    : Adiv Abramson
      ' Date      : 03/11/2021
      ' Purpose   : Ensure specified column arrays have the same number of elements
      '           : and that they contain no more than 100 elements (arbitrary limit).
      '           : Each element must be a number, or up to 3 letters. Must detect and
      '           : ignore header row labels since CopyColumns() handles those.
      '           : Header row labels now enclosed in [].
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 03/11/2021 - Adiv Abramson
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
      Dim strCol As String
      '*********************************

      'Numerics:
      '*********************************
      Dim lngCols As Long
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
      Dim rngCell As Range
      Dim rngTest As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arColumns As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim objRegex As RegExp
      '*********************************

      'Variants:
      '*********************************
      Dim vntColumn As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnIsValidCol As Boolean
      '*********************************

      'Constants
      '*********************************
      Const MAX_COLUMNS = 100
      Const LAST_COL = "XFD"
      '*********************************

10    On Error GoTo ErrProc

20    CopyColumns_ValidateColumnSpecs = False

      '========================================================
      'Validate data type and size of column array variables.
      '========================================================
30    If Not IsArray(SourceColumns) Then msgAttention "Invalid Source Columns array.": Exit Function
40    If Not IsArray(TargetColumns) Then msgAttention "Invalid Target Columns array.": Exit Function
50    If UBound(SourceColumns) <> UBound(TargetColumns) Then
60       strMessage = "Unequal column arrays:" _
            & vbCrLf & UBound(SourceColumns) + 1 & " source columns" _
            & vbCrLf & UBound(TargetColumns) + 1 & " target columns"
70       msgAttention strMessage
80       Exit Function
90    End If

      '========================================================
      'Impose limit on the size of the column arrays.
      '========================================================
100   lngCols = UBound(SourceColumns)
110   If lngCols > MAX_COLUMNS Then
120      msgAttention "Cannot copy more than " & MAX_COLUMNS & " columns at a time."
130      Exit Function
140   End If

      '========================================================
      'Use Data tab to validate the ranges since it doesn't
      'have any merged cells in row 1.
      '========================================================
150   Set wsData = ThisWorkbook.Worksheets(WorksheetNames.Data)

      '========================================================
      'Ensure each column spec is a valid number or a series
      'of 1 to 3 letters
      '========================================================
160   Set objRegex = New RegExp
170   With objRegex
180      .Global = False
190      .IgnoreCase = True
200      .Pattern = "^([A-Z]{1,3}|[1-9]\d*)$"
210   End With

220   For Each arColumns In Array(SourceColumns, TargetColumns)
230      For Each vntColumn In arColumns
240         If Left(vntColumn, 1) = "[" And Right(vntColumn, 1) = "]" Then
250            GoTo NextColumn
260         End If
            
270         If Not objRegex.Test(vntColumn) Then
280            msgAttention "Invalid column specification: """ & vntColumn & """."
290            Exit Function
300         End If
            
310         If IsNumeric(vntColumn) Then
320            vntColumn = CLng(vntColumn)
330            If vntColumn < 1 Or vntColumn > LAST_WORKSHEET_COL Then
340               msgAttention "Invalid column specification: " & vntColumn & "."
350               Exit Function
360            End If 'IsNumeric(vntColumn)
370         Else
              '========================================================
              'Ensure column letter specification is valid. Column
              'specificatipn might be a label in the header row.
              '========================================================
380           blnIsValidCol = False
390           For Each rngCell In wsData.Rows(1).Cells
400              strCol = Split(rngCell.Address, "$")(0)
410              If strCol = UCase(vntColumn) Then
420                 blnIsValidCol = True
430                 Exit For
440              End If 'strCol = UCase(vntColumn)
450           Next rngCell
              
460           If Not blnIsValidCol Then
470              msgAttention "Invalid column specification: " & vntColumn & "."
480              Exit Function
490           End If 'Not blnIsValidCol
500        End If 'IsNumeric(vntColumn)
            
NextColumn:
510      Next vntColumn
520   Next arColumns
            

530   CopyColumns_ValidateColumnSpecs = True
         
540   Exit Function

ErrProc:
550      CopyColumns_ValidateColumnSpecs = False
560      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of  procedure CopyColumns_ValidateColumnSpecs of Module " & MODULE_NAME
End Function

Public Function CopyColumns_ValidateWorksheetInfo(ByRef DataWorksheet As Worksheet, _
                                                  ByVal HeaderRow As Long, _
                                                  Optional ByVal MustHaveData As Boolean = True) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyColumns_ValidateWorksheets
      ' Author    : Adiv Abramson
      ' Date      : 03/11/2021
      ' Purpose   : Ensure specified worksheets have data and that a valid
      '           : header row is specified.
      '           : (1.1) Now supporting validation of empty worksheets as long as
      '           : a header row is present. Using new optional parameter MustHaveData, whose
      '           : default value is True.
      '           : Now also validating the header row. But keep in mind that row may have blanks.
      '           : So as long as it isn't ALL blanks it's OK.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 03/11/2021 - Adiv Abramson
      '           : 1.1 - 04/06/2021 - Adiv Abramson
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
      Dim rngHeaders As Range
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

20    CopyColumns_ValidateWorksheetInfo = False

30    If DataWorksheet Is Nothing Then msgAttention "Invalid worksheet specified.": Exit Function
      '========================================================
      'NEW CODE: 04/06/2021 Now checking for data is optional, e.g.
      'when specified worksheet is the target, which is typically
      'empty, except for the headers.
      '========================================================
40    If MustHaveData Then
50       If Not HasData(DataWorksheet:=DataWorksheet) Then
60          msgAttention DataWorksheet.Name & " has no data."
70          Exit Function
80       End If
90    End If

100   If HeaderRow < 1 Or HeaderRow > LAST_WORKSHEET_ROW - 1 Then
110      msgAttention "Invalid Header Row:" & HeaderRow
120      Exit Function
130   End If


      '========================================================
      'NEW CODE: 04/06/2021 Validate the header row.
      '========================================================
140   Set rngHeaders = DataWorksheet.UsedRange.Rows(HeaderRow)
150   If Application.WorksheetFunction.CountA(rngHeaders) = 0 Then Exit Function


160   CopyColumns_ValidateWorksheetInfo = True

170   Exit Function

ErrProc:
180      CopyColumns_ValidateWorksheetInfo = False
190      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of  procedure CopyColumns_ValidateWorksheets of Module " & MODULE_NAME
End Function

Public Function CopyColumns(ByRef SourceSheet As Worksheet, ByRef TargetSheet As Worksheet, _
                            ByVal SourceColumns As Variant, ByVal TargetColumns As Variant, _
                            Optional ByVal SourceHeaderRow As Integer = 1, _
                            Optional ByVal TargetHeaderRow As Integer = 1) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : CopyColumns
      ' Author    : Adiv Abramson
      ' Date      : 03/11/2021
      ' Purpose   : Convenience function for copying one or more columns from one worksheet
      '           : to another. Worksheets may be in different workbooks or columns may be
      '           : copied within the same worksheet.
      '           : Columns may be identified by position (1,2,3), letter (A,Z,BW) or label
      '           : ([Posting Date]). When labels are specified the header row will be searched.
      '           : Header row is 1 by default.
      '           : For each column in SourceSheet, all cells are assumed to be of the same type.
      '           : This enables copying via arrays instead of ranges.
      '           : Assume first column of worksheet has data and determine the last row of data
      '           : from it. Assume that data in the first column is contiguous (no gaps at all).
      '           : Copying range by value will NOT copy formulas but rather the values those
      '           : formulas represent.
      '           : (1.1) Preserve horizontal alignment when copying.
      '           : (1.2) Added more validations.
      '           : (1.3) Accommodate ranges in column specifications, e.g. "A:K" or "BX:XAB".
      '           : Impose a limit of 100 columns
      '           : (1.4) Differentiate column letters from header labels by enclosing the latter in [].
      '           : (1.5) Don't rely on first column of the Source worksheet being populated. Instead, for
      '           : each Source column being copied, determine the size of the range to be copied of
      '           : populated cells.
      '           : (1.6) Now supporting filtered source rows. If target worksheet is filtered, release its
      '           : Autofilter first. If target and source worksheets are the same, collect the arrays of data
      '           : corresponding to each target column in dictMapping, then release AutoFilter, and finally,
      '           : assign the arrays to the specified target columns, which are on the same worksheet as the
      '           : source columns.
      '           : TODO: Restore AutoFilter settings on source worksheet in all cases.
      '           : (1.7) Now CopyColumns_ValidateWorksheetInfo() has optional parameter MustHaveData, which by
      '           : default is True. But set it to False when validating the target worksheet, which typically
      '           : will be empty.
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
      ' Versions  : 1.0 - 03/11/2021 - Adiv Abramson
      '           : 1.1 - 03/11/2021 - Adiv Abramson
      '           : 1.2 - 03/11/2021 - Adiv Abramson
      '           : 1.3 - 03/11/2021 - Adiv Abramson
      '           : 1.4 - 03/11/2021 - Adiv Abramson
      '           : 1.5 - 03/11/2021 - Adiv Abramson
      '           : 1.6 - 04/06/2021 - Adiv Abramson
      '           : 1.7 - 04/06/2021 - Adiv Abramson
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
      Dim strNumberFormat As String
      '*********************************

      'Numerics:
      '*********************************
      Dim intSheet As Integer
      Dim intCol As Integer
      Dim intCols As Integer
      Dim lngArea As Long
      Dim lngAreas As Long
      Dim intSourceCol As Integer
      Dim intTargetCol As Integer
      Dim lngFirstSourceDataRow As Long
      Dim lngFirstTargetDataRow As Long
      Dim lngLastDataRow As Long
      Dim lngDataRows As Long
      Dim intHeaderRow As Integer
      Dim lngHAlignment As Long
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
      Dim rngSourceHeader As Range
      Dim rngSourceHeaders As Range
      Dim rngHeaders As Range
      Dim rngTargetHeader As Range
      Dim rngTargetHeaders As Range
      Dim rngSourceColumnData As Range
      Dim rngTargetColumnData As Range
      Dim rngArea As Range
      '*********************************

      'Arrays:
      '*********************************
      Dim arColumnData As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim dictMapping As Dictionary
      '*********************************

      'Variants:
      '*********************************
      Dim vntSourceColumn As Variant
      Dim vntTargetColumn As Variant
      Dim vntColumn As Variant
      Dim vntIndex As Variant
      Dim vntMappingKey As Variant
      '*********************************

      'Booleans:
      '*********************************
      Dim blnSourceFiltered As Boolean
      Dim blnTargetFiltered As Boolean
      Dim blnTargetIsSource As Boolean
      Dim blnAllBlanks As Boolean
      '*********************************

      'Constants
      '*********************************
      Const TOTAL_SHEETS = 2
      '*********************************

10    On Error GoTo ErrProc

20    CopyColumns = False

      '========================================================
      'NEW CODE: 04/06/2021 Note if source worksheet or target
      'worksheet is filtered.
      '========================================================
30    blnSourceFiltered = (SourceSheet.AutoFilterMode)
40    blnTargetFiltered = (TargetSheet.AutoFilterMode)
50    blnTargetIsSource = (SourceSheet.Name = TargetSheet.Name)

      '========================================================
      'Validate specified worksheets
      '========================================================
60    If Not CopyColumns_ValidateWorksheetInfo(DataWorksheet:=SourceSheet, _
         HeaderRow:=SourceHeaderRow) Then Exit Function


      '========================================================
      'NEW CODE: 04/06/2021 Release AutoFilter, if any, on
      'Target sheet.
      '========================================================
70    If Not blnTargetIsSource Then
80       With TargetSheet
90          If .AutoFilterMode Then .AutoFilter.ShowAllData
100      End With
110   End If 'Not blnTargetIsSource

120   If Not CopyColumns_ValidateWorksheetInfo(DataWorksheet:=TargetSheet, _
         HeaderRow:=TargetHeaderRow, MustHaveData:=False) Then Exit Function
         
         
      '========================================================
      'NEW CODE: 03/11/2021 Accommodate embedded range specifications
      'in Source and Target columns.
      '========================================================
130    SourceColumns = CopyColumns_ExpandRanges(ColumnSpecs:=SourceColumns)
140    TargetColumns = CopyColumns_ExpandRanges(ColumnSpecs:=TargetColumns)
       
      '========================================================
      'Inspect and validate each element in (potentially expanded)
      'Source and Target column arrays.
      '========================================================
150   If Not CopyColumns_ValidateColumnSpecs(SourceColumns:=SourceColumns, _
         TargetColumns:=TargetColumns) Then Exit Function

160   intCols = UBound(SourceColumns)

      '========================================================
      'Reference header row of Source and Target sheets
      'so we can index or seach into them to locate the source
      'and target columns to copy.
      '========================================================
170   Set rngSourceHeaders = SourceSheet.UsedRange.Rows(SourceHeaderRow)
180   Set rngTargetHeaders = TargetSheet.UsedRange.Rows(TargetHeaderRow)

      '========================================================
      'Map Source columns to Target columns. Column specification
      'may be a number (index), a letter or a label.
      'If column spec is a number, determine source header cell
      'using Cells().
      'If column is a string, try to locate it in the Source
      'Sheet row.
      'Otherwise it must be a column letter, e.g. "AB".
      '========================================================
      '========================================================
      'NEW CODE: 03/11/2021 Header labels now enclosed in [].
      'This helps CopyColumns_ValidateColumnSpecs(), which will
      'won't try to validate such elements.
      '========================================================
190   Set dictMapping = New Dictionary
200   For intCol = 0 To intCols
210      vntSourceColumn = SourceColumns(intCol)
220      If Left(vntSourceColumn, 1) = "[" And Right(vntSourceColumn, 1) = "]" Then
230         vntSourceColumn = LTRUNC(vntSourceColumn, 1)
240         vntSourceColumn = RTRUNC(vntSourceColumn, 1)
250      End If
         
260      vntTargetColumn = TargetColumns(intCol)
270      If Left(vntTargetColumn, 1) = "[" And Right(vntTargetColumn, 1) = "]" Then
280         vntTargetColumn = LTRUNC(vntTargetColumn, 1)
290         vntTargetColumn = RTRUNC(vntTargetColumn, 1)
300      End If
         
310      If IsNumeric(vntSourceColumn) Then
320         Set rngSourceHeader = SourceSheet.Cells(SourceHeaderRow, vntSourceColumn)
330      Else
340         vntIndex = Application.Match(vntSourceColumn, rngSourceHeaders, 0)
350         If Not IsError(vntIndex) Then
360            Set rngSourceHeader = SourceSheet.Cells(SourceHeaderRow, vntIndex)
370         Else
380            Set rngSourceHeader = SourceSheet.Range(vntSourceColumn & SourceHeaderRow)
390         End If 'Not IsError(vntIndex
400      End If 'IsNumeric(vntSourceColumn)
410      intSourceCol = rngSourceHeader.Column
         
420      If IsNumeric(vntTargetColumn) Then
430         Set rngTargetHeader = TargetSheet.Cells(TargetHeaderRow, vntTargetColumn)
440      Else
450         vntIndex = Application.Match(vntTargetColumn, rngTargetHeaders, 0)
460         If Not IsError(vntIndex) Then
470            Set rngTargetHeader = TargetSheet.Cells(TargetHeaderRow, vntIndex)
480         Else
490            Set rngTargetHeader = TargetSheet.Range(vntTargetColumn & TargetHeaderRow)
500         End If 'Not IsError(vntIndex
510      End If 'IsNumeric(vntTargetColumn)
520      intTargetCol = rngTargetHeader.Column
         
530      dictMapping.Item(intSourceCol) = intTargetCol
540   Next intCol

      '========================================================
      'Copy specified Source columns to specified Target columns.
      '========================================================
550   lngFirstSourceDataRow = SourceHeaderRow + 1
560   lngFirstTargetDataRow = TargetHeaderRow + 1

570   For Each vntMappingKey In dictMapping.Keys
580      intSourceCol = Int(vntMappingKey)
590      intTargetCol = dictMapping.Item(vntMappingKey)
         '========================================================
         'Determine size (# of rows) of Source Sheet data columns,
         'taking into accoun the header row.
         '========================================================
         '========================================================
         'NEW CODE: 03/11/2021 Account for emplty column and for
         'non-contiguous data.
         '========================================================
600      With SourceSheet.UsedRange
610         lngLastDataRow = .Cells(LAST_WORKSHEET_ROW, intSourceCol).End(xlUp).Row
620         lngDataRows = lngLastDataRow - SourceHeaderRow
630      End With
         
640      If lngDataRows < 1 Then
650         dictMapping.Remove vntMappingKey
660         GoTo NextMappingKey
670      End If 'lngDataRows < 1
         
         '========================================================
         'NEW CODE: 04/06/2021 Source and/or Target sheet may be
         'in a filtered state. Always copy visible cells from Source,
         'which will work even if there is no active filter.
         '========================================================
680      With SourceSheet
690         Set rngSourceColumnData = .Range(.Cells(lngFirstSourceDataRow, intSourceCol), _
               .Cells(lngLastDataRow, intSourceCol))
700      End With
710      Set rngSourceColumnData = GetSpecialCells(DataRange:=rngSourceColumnData, CellType:=xlCellTypeVisible)
         
         '========================================================
         'NEW CODE: 04/06/2021 If source range is empty, skip it.
         'This is a second test in case the first test above fails
         'to detect an empty column.
         '========================================================
720      blnAllBlanks = True
730      With rngSourceColumnData
740         lngAreas = .Areas.Count
750         For lngArea = 1 To lngAreas
760            Set rngArea = .Areas(lngArea)
770            If rngArea.Cells.Count <> Application.WorksheetFunction.CountBlank(rngArea) Then
780               blnAllBlanks = False
790               Exit For
800            End If
810         Next lngArea
820      End With
         
830      If blnAllBlanks Then
840         dictMapping.Remove vntMappingKey
850         GoTo NextMappingKey
860      End If
         
         '========================================================
         'NEW CODE: 04/06/2021 The actual size of the data array
         'from the Source sheet depends on the number of **visible**
         'rows. Therefore, the Target range must be sized accordingly.
         '========================================================
870      lngDataRows = rngSourceColumnData.Cells.Count
880      Set rngTargetColumnData = TargetSheet.Cells(lngFirstTargetDataRow, _
            intTargetCol).Resize(RowSize:=lngDataRows)
         
890      With rngSourceColumnData.Cells(1, 1)
900         strNumberFormat = .NumberFormat
910         lngHAlignment = .HorizontalAlignment
920      End With
         
930      If Not blnSourceFiltered Then
            '========================================================
            'NEW CODE: 04/06/2021 Any AutoFilter on Target has already
            'been released.
            '========================================================
940         rngTargetColumnData.Value = rngSourceColumnData.Value
950         SetRangeFormat DataRange:=rngTargetColumnData, FormatString:=strNumberFormat
960         rngTargetColumnData.HorizontalAlignment = lngHAlignment
970      ElseIf blnSourceFiltered And Not blnTargetIsSource Then
980          arColumnData = GetAllRangeValues(DataRange:=rngSourceColumnData)
990          rngTargetColumnData.Value = Application.WorksheetFunction.Transpose(arColumnData)
1000         SetRangeFormat DataRange:=rngTargetColumnData, FormatString:=strNumberFormat
1010         rngTargetColumnData.HorizontalAlignment = lngHAlignment
1020     ElseIf blnSourceFiltered And blnTargetIsSource Then
            '========================================================
            'NEW CODE: 04/06/2021 Source and Target are the same and
            'in AutoFilterMode. Must store the data arrays temporarily,
            'release the AutoFilter and then assign the arrays to the
            'Target columns.
            '========================================================
1030        vntTargetColumn = dictMapping.Item(vntMappingKey)
1040        arColumnData = GetAllRangeValues(DataRange:=rngSourceColumnData)
1050        dictMapping.Item(vntMappingKey) = Array(vntTargetColumn, arColumnData, strNumberFormat, lngHAlignment)
1060     End If 'Not blnSourceFiltered
NextMappingKey:
1070  Next vntMappingKey


      '========================================================
      'NEW CODE: 04/06/2021 When copying filtered data from one
      'part of a given worksheet to another part of the same
      'worksheet, retrieve the stored data arrays. Target column
      'is first element of item corresponding to each source
      'column and the corresponding data array for that filtered
      'source column is the second element and NumberFormat is
      'the third element. Horizontal Alignment is the fourth.
      '========================================================
1080  If blnSourceFiltered And blnTargetIsSource Then
1090     SourceSheet.AutoFilter.ShowAllData
1100     For Each vntMappingKey In dictMapping.Keys
1110        intSourceCol = Int(vntMappingKey)
1120        intTargetCol = dictMapping.Item(vntMappingKey)(0)
1130        arColumnData = dictMapping.Item(vntMappingKey)(1)
1140        lngDataRows = UBound(arColumnData) + 1
1150        strNumberFormat = dictMapping.Item(vntMappingKey)(2)
1160        lngHAlignment = dictMapping.Item(vntMappingKey)(3)
            
1170        Set rngTargetColumnData = TargetSheet.Cells(lngFirstTargetDataRow, _
               intTargetCol).Resize(RowSize:=lngDataRows)
1180        rngTargetColumnData.Value = Application.WorksheetFunction.Transpose(arColumnData)
1190        SetRangeFormat DataRange:=rngTargetColumnData, FormatString:=strNumberFormat
1200        rngTargetColumnData.HorizontalAlignment = lngHAlignment
1210     Next vntMappingKey
1220  End If 'blnSourceFiltered And blnTargetIsSource

        
1230  CopyColumns = True
         
         
1240  Exit Function

ErrProc:
1250     CopyColumns = False
1260     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of  procedure CopyColumns of Module " & MODULE_NAME
End Function



Public Sub PasteFromEditor(ByVal TargetTable As String, ByVal TableField As String)
      '---------------------------------------------------------------------------------------
      ' Procedure : PasteFromEditor
      ' Author    : Adiv Abramson
      ' Date      : 4/24/2024
      ' Purpose   : Copies selected text from VBA editor that's currently in the Clipboard
      '           : into the target table's first row, in the specified field.
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
      ' Versions  : 1.0 - 4/24/2024 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strDataWorksheet As String
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
      Dim rngTarget As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objTarget As ListObject
      Dim objDataObject As DataObject
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

      '========================================================
      'Validate inputs
      '========================================================
20    strDataWorksheet = WorksheetNames.CodeHelper
30    Set objTarget = GetTable(TableName:=TargetTable, DataWorksheet:=strDataWorksheet)
40    If objTarget Is Nothing Then Exit Sub

50    If Not IsTableField(FieldName:=TableField, _
                          DataTableName:=TargetTable, _
                          DataWorksheetName:=strDataWorksheet) Then Exit Sub


      '========================================================
      'Set target to first cell in specified table field.
      '========================================================
      '========================================================
      'NEW CODE: 04/24/2024 If target table is completely
      'empty, add a new row to it
      '========================================================
60    If objTarget.DataBodyRange Is Nothing Then
70       objTarget.ListRows.Add
80    End If
90    Set rngTarget = objTarget.ListColumns(TableField).DataBodyRange.Cells(1, 1)
100   rngTarget.PasteSpecial xlPasteAll



110   Exit Sub

ErrProc:

120      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure PasteFromEditor of Module " & MODULE_NAME
End Sub
