Public Function CombineArrays(ByVal ArrayList As Variant) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : CombineArrays
      ' Author    : Adiv Abramson
      ' Date      : 09/25/2024
      ' Purpose   : Combine multiple 1D arrays.
      '           : Nested arrays are not supported.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 09/25/2024 - Adiv Abramson
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
      Dim lngSubArraySize As Long
      Dim lngCombinedArraySize As Long
      Dim lngCombinedArrayIndex As Long
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
      Dim arCombined() As Variant
      '*********************************

      'Objects:
      '*********************************

      '*********************************

      'Variants:
      '*********************************
      Dim vntSubArray As Variant
      Dim vntSubArrayElement As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************
      
      '*********************************

10    On Error GoTo ErrProc

20    CombineArrays = Null

      '========================================================
      'Validate input
      'ArrayList must contain at least 2 arrays.
      '========================================================
30    If Not IsArray(ArrayList) Then Exit Function
40    If GetUBound(ArrayList) < 1 Then Exit Function

50    lngCombinedArraySize = -1
60    For Each vntSubArray In ArrayList
70       If Not IsArray(vntSubArray) Then Exit Function
80       lngSubArraySize = GetUBound(vntSubArray)
90       If lngSubArraySize = -1 Then Exit Function
         '========================================================
         'Use the number of elements in each subarray, which is
         'the UBound + 1 and add that to the size of the
         'combined array to be created.
         '========================================================
100      lngCombinedArraySize = lngCombinedArraySize + lngSubArraySize + 1
110   Next vntSubArray

120   ReDim arCombined(lngCombinedArraySize)

      '========================================================
      'Populate array combining elements of all the sub arrays.
      '========================================================
130   lngCombinedArrayIndex = 0
140   For Each vntSubArray In ArrayList
150      For Each vntSubArrayElement In vntSubArray
160         arCombined(lngCombinedArrayIndex) = vntSubArrayElement
170         lngCombinedArrayIndex = lngCombinedArrayIndex + 1
180      Next vntSubArrayElement
190   Next vntSubArray

200   CombineArrays = arCombined

210   Exit Function

ErrProc:
220      CombineArrays = Null
230      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                & Erl & " in procedure CombineArrays of Module " & MODULE_NAME
End Function
