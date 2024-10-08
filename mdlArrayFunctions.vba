'---------------------------------------------------------------------------------------
' Module    : mdlUtilArrayFunctions
' Author    : Adiv Abramson
' Date      : 00/00/2024
' Purpose   : Contains some array functions moved from mdlUtil
'           :
'           :
'           :
'           :
'           :
'           :
' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Option Explicit
Option Private Module

Private Const MODULE_NAME = "mdlUtilArrayFunctions"

Public Function TransposeArray(ByVal InputArray As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : TransposeArray
      ' Author    : Adiv Abramson
      ' Date      : 00/00/2024
      ' Purpose   : Transpose 1D or 2D arrays having > 64K elements, to circumvent
      '           : VBA limit
      '           : (1.1) Updated to handle 2D arrays as well
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
      'Versions   : 1.0 - 00/00/2024 - Adiv Abramson
      '           : 1.1 - 00/00/2024 - Adiv Abramson
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
      Dim lngRow As Long
      Dim lngRows As Long
      Dim lngColumn As Long
      Dim lngColumns As Long

      Dim intDimensions As Integer
      Dim lngInputArrayLBound As Long
      Dim lngInputArrayUBound As Long

      Dim lngTArrayLBoundD1 As Long
      Dim lngTArrayUBoundD1 As Long
      Dim lngTArrayLBoundD2 As Long
      Dim lngTArrayUBoundD2 As Long
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
      Dim arTArray() As Variant
      Dim arBounds(0 To 1) As Variant
      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants
      '*********************************
      Dim vntTemp As Variant
      '*********************************

      'Booleans
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const ROWS_INDEX = 0
      Const COLUMNS_INDEX = 1
      Const LBOUND_INDEX = 0
      Const UBOUND_INDEX = 1
      '*********************************


10    On Error GoTo ErrProc

20    TransposeArray = Null
30    If Not IsArray(InputArray) Then Exit Function

      '========================================================
      'Determine how many dimensions array has, since this
      'affects how it will be transposed (I think)
      '========================================================
40    For intDimensions = 2 To 1 Step -1
50       On Error Resume Next
60       vntTemp = LBound(InputArray, intDimensions)
70       If Err = 0 Then
80          lngInputArrayLBound = LBound(InputArray, intDimensions)
90          lngInputArrayUBound = UBound(InputArray, intDimensions)
100         Exit For
110      End If
120   Next intDimensions

      '========================================================
      'The L/UBounds of the InputArray's highest dimension
      'should now be stored in the corresponding variables
      '========================================================
130   Err.Clear
140   On Error GoTo ErrProc

      '========================================================
      'Cannot transpose array of rank 3 or higher
      '========================================================
150   If intDimensions > 2 Then Exit Function

      '========================================================
      'Transpose a 1D array with n elements into an n X 1
      '2D array.
      '========================================================
160   If intDimensions = 1 Then
         '========================================================
         'Now get the L and U bounds of the array corresponding
         'to this dimension
         '========================================================
170      lngInputArrayLBound = LBound(InputArray, intDimensions)
180      lngInputArrayUBound = UBound(InputArray, intDimensions)
         
190      ReDim arTArray(lngInputArrayLBound To lngInputArrayUBound, 0 To 0)
200      For i = lngInputArrayLBound To lngInputArrayUBound
210      arTArray(i, 0) = InputArray(i)
220   Next i

230   Else
         '========================================================
         'Transpose an n x m array (2D) into a m x n (2D)
         'First get the bounds of each dimension of the InputArray
         'so that the bounds of the transposed array can be
         'determined
         '========================================================
240      arBounds(ROWS_INDEX) = Array(LBound(InputArray, 1), UBound(InputArray, 1))
250      arBounds(COLUMNS_INDEX) = Array(LBound(InputArray, 2), UBound(InputArray, 2))
260      ReDim arTArray(arBounds(COLUMNS_INDEX)(LBOUND_INDEX) To arBounds(COLUMNS_INDEX)(UBOUND_INDEX), _
            arBounds(ROWS_INDEX)(LBOUND_INDEX) To arBounds(ROWS_INDEX)(UBOUND_INDEX))
         
270      lngRows = arBounds(COLUMNS_INDEX)(UBOUND_INDEX)
280      lngColumns = arBounds(ROWS_INDEX)(UBOUND_INDEX)
290      For lngRow = arBounds(COLUMNS_INDEX)(LBOUND_INDEX) To lngRows
300         For lngColumn = arBounds(ROWS_INDEX)(LBOUND_INDEX) To lngColumns
310            arTArray(lngRow, lngColumn) = InputArray(lngColumn, lngRow)
320         Next lngColumn
330      Next lngRow

340   End If 'intDimensions = 1

350   TransposeArray = arTArray

         
360   Exit Function

ErrProc:
          
370       TransposeArray = Null
380       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure TransposeArray of Module mdlUtil"


End Function





Public Function GetArraySlice(ByVal InputArray As Variant, ByVal StartIndex As Long, _
     ByVal StopIndex As Long, Optional ByVal DELIMITER As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : GetArraySlice
      ' Author    : Adiv
      ' Date      : 00/00/2024
      ' Purpose   : Returns the elements of the specified array from StartIndex up to and
      '           : including StopIndex. Optionally delimit the elements of the returned array
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
      'Versions   : 1.0 - 00/00/2024 - Adiv Abramson
      '           : 1.1 - 00/00/2024 - Adiv Abramson
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
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************

      '*********************************

      'Numerics:
      '*********************************
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

      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim arOutput As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants:
      '*********************************

      '*********************************



10    On Error GoTo ErrProc

20    GetArraySlice = Null

      'Validate inputs
30    If IsNull(InputArray) Then
40       Exit Function
50    ElseIf Not IsArray(InputArray) Then
60       Exit Function
70    ElseIf StartIndex < 0 Then
80       Exit Function
90    ElseIf StopIndex < 0 Then
100      Exit Function
110   ElseIf StopIndex < StartIndex Then
120      Exit Function
130   ElseIf StartIndex < LBound(InputArray) Then
140      Exit Function
150   ElseIf StartIndex > UBound(InputArray) Then
160      Exit Function
170   ElseIf StopIndex > UBound(InputArray) Then
180      StopIndex = UBound(InputArray)
190   ElseIf Not IsMissing(DELIMITER) Then
200      If Trim(DELIMITER) = "" Then
210         Exit Function
220      End If
230   End If 'IsNull(InputArray)

240   For i = StartIndex To StopIndex
250      If IsMissing(DELIMITER) Then
260         AddToArray arOutput, InputArray(i), False
270      Else
280         AddToArray arOutput, DELIMITER & CStr(InputArray(i)) & DELIMITER, False
290      End If 'IsMissing(Delimiter)
300   Next i

310   GetArraySlice = arOutput

320   Exit Function

ErrProc:
330      GetArraySlice = Null
340      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure GetArraySlice of Module mdlUtil"
End Function


Public Function GetArrayInfo(ByVal DataArray As Variant) As Dictionary
'---------------------------------------------------------------------------------------
' Procedure : GetArrayInfo
' Author    : Adiv Abramson
' Date      : 00/00/2024
' Purpose   : Generates dictionary with the following keys:
'           :
'           : CountAll, CountOfStrings, CountOfNumbers, CountOfVariants, CountOfBooleans
'           : For string elements: Shortest, Longest
'           : For numeric elements: Smallest, Largest
'           : Strings - array of string elements, if any
'           : Numbers - array of numeric elements, if any
'           : Booleans - array of Boolean elements, if any
'           : Variants - array of all other elements, if any
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
' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

'Strings:
'*********************************
Dim strDataType As String
Dim strKey As String
'*********************************

'Numerics:
'*********************************
Dim lngMaxLength As Long
Dim lngMinLength As Long
Dim dblMinValue As Double
Dim dblMaxValue As Double
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
Dim arStrings As Variant
Dim arNumbers As Variant
Dim arBooleans As Variant
Dim arVariants As Variant
Dim arLengths As Variant
Dim arArrays As Variant
Dim arData As Variant
'*********************************

'Objects:
'*********************************
Dim dictArrayInfo As Dictionary
'*********************************

'Variants:
'*********************************
Dim vntElement As Variant
Dim vntValue As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************

10    On Error GoTo ErrProc
      
20    Set GetArrayInfo = Nothing
      
      '========================================================
      'Validate DataArray
      '========================================================
30    If IsNull(DataArray) Then Exit Function
40    If IsEmpty(DataArray) Then Exit Function
50    If Not IsArray(DataArray) Then Exit Function
60    If LBound(DataArray) = -1 Then Exit Function
      
      '========================================================
      'Populate array for each data type, if any elements for
      'it exist.
      '========================================================
70    arStrings = Null
80    arNumbers = Null
90    arBooleans = Null
100   arVariants = Null
110   arArrays = Null
120   arLengths = Null
      
      
130   For Each vntElement In DataArray
140      strDataType = TypeName(vntElement)
150      Select Case strDataType
            Case "String"
170            AddToArray DataArray:=arStrings, InputValue:=vntElement, KeepUnique:=False
               
            Case "Double", "Long", "Integer"
200            AddToArray DataArray:=arNumbers, InputValue:=vntElement, KeepUnique:=False
               
            Case "Boolean"
230            AddToArray DataArray:=arBooleans, InputValue:=vntElement, KeepUnique:=False
               
            Case Else
260            AddToArray DataArray:=arVariants, InputValue:=vntElement, KeepUnique:=False
            
280      End Select
290   Next vntElement
      
300   Set dictArrayInfo = New Dictionary
      '========================================================
      'Compute stats
      '========================================================
310   If Not IsNull(arStrings) Then
320      For Each vntElement In arStrings
330         vntValue = Len(vntElement)
340         AddToArray DataArray:=arLengths, InputValue:=vntValue, KeepUnique:=True
350      Next vntElement
         
370      lngMinLength = Application.WorksheetFunction.Min(arLengths)
380      lngMaxLength = Application.WorksheetFunction.Max(arLengths)
390      dictArrayInfo.Item("Shortest") = lngMinLength
400      dictArrayInfo.Item("Longest") = lngMaxLength
410   End If 'Not IsNull(arStrings)
      
      
420   If Not IsNull(arNumbers) Then
430      dblMinValue = Application.WorksheetFunction.Min(arNumbers)
440      dblMaxValue = Application.WorksheetFunction.Max(arNumbers)
         
460      dictArrayInfo.Item("Smallest") = dblMinValue
470      dictArrayInfo.Item("Largest") = dblMaxValue
480   End If 'Not IsNull(arNumbers)
      
      '========================================================
      'Add overall stats to dictionary
      '========================================================
490   arArrays = Array(Array("Strings", arStrings), _
                       Array("Numbers", arNumbers), _
                       Array("Booleans", arBooleans), _
                       Array("Variants", arVariants))
      
      
530   With dictArrayInfo
540      For Each vntElement In arArrays
550         strDataType = vntElement(0)
560         strKey = "CountOf" & strDataType
570         arData = vntElement(1)
580         If Not IsNull(arData) Then
590            .Item(strKey) = UBound(arData) + 1
600            .Item(strDataType) = arData
610         Else
620            .Item(strKey) = 0
630            .Item(strDataType) = Null
640         End If 'Not IsNull(arData)
650      Next vntElement
660   End With
      
670   With dictArrayInfo
680      .Item("CountAll") = UBound(DataArray) + 1
690   End With
      
700   Set GetArrayInfo = dictArrayInfo



Exit Function

ErrProc:
   Set GetArrayInfo = Nothing
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure GetArrayInfo of Module " & MODULE_NAME
End Function


Public Function IsInArray(ByVal InputValues As Variant, ByVal LookupList As Variant, _
                Optional ByRef NonMatchingItems As Variant = Null, _
                Optional ByRef MatchingItems As Variant = Null, _
                         Optional ByVal Strict As Boolean = True) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsInArray
' Author    : Adiv Abramson
' Date      : 00/00/2024
' Purpose   : Indicate whether vntLookup is in the array LookupList
'           : Parameter Strict = True means return False if even 1 elememt in input array does not
'           : exist lookup array. Strict = False means return True if at least 1 element in input
'           : array exists in lookup array
'           : (1.1) Now supporting optional out params for Non-Matching and Matching input array
'           : items
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
' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
'           : 1.1 - 00/00/2024 - Adiv Abramson
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

'*********************************

'Numerics:
'*********************************
Dim i As Integer
Dim intListSize As Integer

Dim j As Integer
Dim intInputSize As Integer

Dim intMatches As Integer
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
Dim blnIsInputArray As Boolean
Dim blnMatched As Boolean
'*********************************

'Constants:
'*********************************

'*********************************


10    On Error GoTo ErrProc
      
20    IsInArray = False
      
      'validate inputs
30    If Not IsArray(LookupList) Then
40       Exit Function
50    End If
      
60    intListSize = UBound(LookupList)
      
      
      '<NEW CODE: If first parameter is an array, iterate through its items
      'and return True only if each item is in LookupList
70    blnIsInputArray = IsArray(InputValues)
      
80    If blnIsInputArray Then
90       intMatches = -1
100      intInputSize = UBound(InputValues)
110      For i = 0 To intInputSize
            '<NEW CODE:08/24/2019> Optionally capture NonMatching and Matching Items
130         blnMatched = False
140         For j = 0 To intListSize
150            If InputValues(i) = LookupList(j) Then
160               Incr intMatches, 1
170               blnMatched = True
180               Exit For
190            End If 'InputValues(j) = LookupList(i)
200         Next j
210         If Not IsMissing(MatchingItems) Then
220            If blnMatched Then AddToArray MatchingItems, InputValues(i), False
230         End If
240         If Not IsMissing(NonMatchingItems) Then
250            If Not blnMatched Then AddToArray NonMatchingItems, InputValues(i), False
260         End If
            
280      Next i
290      If Strict Then
300         IsInArray = (intMatches = intInputSize)
310      Else
320         IsInArray = (intMatches > -1)
330      End If
      
340   Else
         'Only a single value needs to be looked up
360      For i = 0 To intListSize
370         If LookupList(i) = InputValues Then
380            IsInArray = True
390            Exit Function
400         End If
410      Next i
420   End If 'blnIsInputArray

Exit Function

ErrProc:
  IsInArray = False
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsInArray of Module mdlUtil"

End Function



Public Function EquivArrays(ByRef Array1 As Variant, ByRef Array2 As Variant) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : EquivArrays
      ' Author    : Adiv
      ' Date      : 00/00/2024
      ' Purpose   : Determines if a pair of 1D arrays are equivalent, i.e. same elements exist
      '           : in both arrays.
      '           : Arrays don't have to be of the same size
      '           : Calls IsInArray() twice, switching InputValues and LookupList arguments and
      '           : returns True only if both calls to IsInArray() return True
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
      ' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
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

20    EquivArrays = False

      'Validate inputs
30    If Not IsArray(Array1) Or Not IsArray(Array2) Then
40       Exit Function
50    End If

60    EquivArrays = (IsInArray(InputValues:=Array1, LookupList:=Array2, Strict:=True) _
                     And IsInArray(InputValues:=Array2, LookupList:=Array1, Strict:=True))


70    Exit Function

ErrProc:
        
80       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure EquivArrays of Module mdlUtil"
End Function


Public Function ArrayFind(ByVal SearchValue As Variant, ByVal DataArray As Variant) As Long
      '---------------------------------------------------------------------------------------
      ' Procedure : ArrayFind
      ' Author    : Adiv Abramson
      ' Date      : 00/00/2024
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
      ' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
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


Public Sub AddToArray(ByRef DataArray As Variant, ByVal InputValue As Variant, _
                      ByVal KeepUnique As Boolean, Optional ByVal Flatten As Boolean = False)
      '---------------------------------------------------------------------------------------
      ' Procedure : AddToArray
      ' Author    : Adiv Abramson
      ' Date      : 04/01/2018
      ' Purpose   : Add value to specified array and optionally ensure unique values
      '           : (1.1) Now supporting adding 1D arrays in addition to traditional single values
      '           : (1.2) Now simply adding input arrays to DataArray and in this case not
      '           : enforcing uniqueness.
      '           : (1.3) If InputValue is itself an array, optionally "flatten" it by adding each
      '           : element in InputValue individually to the existing array. Default behavior is
      '           : to add the InputValue array as a single element in the existing array, i.e.
      '           : embedding or nesting it.
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
      '  Versions : 1.0 - 04/03/2018 - Adiv Abramson
      '           : 1.1 - 08/31/2019 - Adiv Abramson
      '           : 1.2 - 06/10/2022 - Adiv Abramson
      '           : 1.3 - 08/26/2024 - Adiv Abramson
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
      Dim lngDataArrayLBound As Long
      Dim lngDataArrayUBound As Long
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

      'Variants
      '*********************************
      Dim vntInputValueItem As Variant
      '*********************************

      'Booleans
      '*********************************
      Dim blnAddingArray As Boolean
      '*********************************

      'Constants
      '*********************************

      '*********************************



10    On Error GoTo ErrProc


      '========================================================
      'NEW CODE: 08/26/2024 Now supporting adding an array to
      'the existing array (embedded array). Will have the option
      'to flatten the new array or simply add it as an element
      'to the existing array
      '========================================================

20    blnAddingArray = IsArray(InputValue)

30    If Not blnAddingArray Then
         '=========================================================================
         'NEW CODE: 08/31/2019 "Classic" mode. Handles single input value
         '=========================================================================

40       If IsEmpty(DataArray) Or IsNull(DataArray) Then
50          ReDim DataArray(0)
60          DataArray(0) = InputValue
70          Exit Sub
80       End If
         
         'this may raise an error if DataArray had been erased
90       On Error Resume Next
100      lngDataArrayUBound = UBound(DataArray)
         
110      If Err <> 0 Then
120         Err.Clear
130         ReDim DataArray(0)
140         DataArray(0) = InputValue
150         Exit Sub
160      End If
         
         
170      On Error GoTo ErrProc
         
180      lngDataArrayLBound = LBound(DataArray)
         
190      If KeepUnique Then
200         For i = lngDataArrayLBound To lngDataArrayUBound
210             If DataArray(i) = InputValue Then
220                Exit Sub
230             End If
240         Next i
250      End If
         
260      ReDim Preserve DataArray(lngDataArrayUBound + 1)
270      DataArray(lngDataArrayUBound + 1) = InputValue

280      Exit Sub
290   End If 'Not blnAddingArray

      '=========================================================================
      'NEW CODE: 08/31/2019 Advanced mode: Handles input values that are
      'themselves 1D arrays
      '=========================================================================
300   If IsEmpty(DataArray) Then
310      DataArray = Null
320      ReDim DataArray(0)
330      DataArray(0) = InputValue
340   ElseIf IsNull(DataArray) Then
350      DataArray = Null
360      ReDim DataArray(0)
370      DataArray(0) = InputValue
380   ElseIf Not IsArray(DataArray) Then
390      DataArray = Null
400      ReDim DataArray(0)
410      DataArray(0) = InputValue
420   Else
         '========================================================
         'NEW CODE: 08/26/2024 Optionally flatten InputValue array.
         'Default is to add as a single element in the existing
         'array
         '========================================================
430      If Not Flatten Then
440         lngDataArrayUBound = UBound(DataArray)
450         ReDim Preserve DataArray(lngDataArrayUBound + 1)
460         DataArray(lngDataArrayUBound + 1) = InputValue
470      Else
480         For Each vntInputValueItem In InputValue
490            lngDataArrayUBound = UBound(DataArray)
500            ReDim Preserve DataArray(lngDataArrayUBound + 1)
510            DataArray(lngDataArrayUBound + 1) = vntInputValueItem
520         Next vntInputValueItem
530      End If 'Not Flatten
540   End If 'IsEmpty(DataArray)



550   Exit Sub

ErrProc:

560       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) _
                  & " in procedure AddToArray of Module " & MODULE_NAME

End Sub



Public Function ConvertDynamicArray(ByVal DynamicArray As Range) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : ConvertDynamicArray
      ' Author    : Adiv Abramson
      ' Date      : 00/00/2024
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
      ' Versions  : 1.0 - 00/00/2024 - Adiv Abramson
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
