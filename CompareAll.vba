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
      ' Versions  : 1.0 - 09/21/2016 - Adiv Abramson
      '           : 1.1 - 11/08/2021 - Adiv Abramson
      '           : 1.2 - 12/23/2021 - Adiv Abramson
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
      'NEW CODE: 11/08/2021 If the option MatchingValues
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
      'NEW CODE: 11/08/2021 Function returns True if number of
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





