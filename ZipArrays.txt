Public Function ZipArrays(ByVal ArrayList As Variant, Optional ByRef UnzippedItems As Dictionary = Nothing) As Variant
      '---------------------------------------------------------------------------------------
      ' Procedure : ZipArrays
      ' Author    : Adiv Abramson
      ' Date      : 09/23/2024
      ' Purpose   : Mimic Python's Zip() function
      '           : Added powerful option to return dictionary of unzipped elements
      '           : from each array in ArrayList, if there are any. Structure of dictionary:
      '           : dictUnZipped.Item(<ArrayListIndex>) = Array(<unzipped items in ArrayList<index>).
      '           : In order to do this, ParamArray must be changed to a regular array so that
      '           : we can add an Optional parameter to the function. This has the added benefit
      '           : of not having to hard code the arrays to be zip. Now an array of unknown size
      '           : containing elements that are arrays can be processed by this function, making it
      '           : more flexible.
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 09/23/2024 - Adiv Abramson
      '           : 1.1 - 09/24/2024 - Adiv Abramson
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
      Dim intArrayIndex As Integer
      Dim intArrays As Integer
      Dim lngMaxUbound As Long
      Dim lngLBound As Long
      Dim lngUbound As Long
      Dim lngArrayIndex As Long
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
      Dim arZippedArrays As Variant
      Dim arZippedArray As Variant
      '*********************************

      'Objects:
      '*********************************
      Dim dictArraySize As Dictionary
      '*********************************

      'Variants:
      '*********************************
      Dim vntArrayElement As Variant
      Dim vntArrayListElement As Variant
      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************

      '*********************************

10    On Error GoTo ErrProc

20    ZipArrays = Null

      '========================================================
      'Validate inputs
      '========================================================
30    If Not IsArray(ArrayList) Then Exit Function
40    intArrays = UBound(ArrayList)
50    If intArrays < 1 Then Exit Function

60    Set dictArraySize = New Dictionary
70    For Each vntArrayListElement In ArrayList
80       If Not IsArray(vntArrayListElement) Then Exit Function
         '========================================================
         'Calling UBound() on an empty array raises an error.
         'Use new GetUBound() function to return -1 if item is
         'an empty array
         '========================================================
90       lngUbound = GetUBound(DataArray:=vntArrayListElement)
100      If lngUbound = -1 Then Exit Function
         '========================================================
         'Arrays may be of different lengths. The shortest array
         'governs processing. Elements beyond the length of the
         'shortest array will not be processed.
         '========================================================
110      dictArraySize.Item(lngUbound) = ""
120   Next vntArrayListElement

130   lngMaxUbound = Application.WorksheetFunction.Min(dictArraySize.Keys)

      '========================================================
      'Create "tuples" (arrays, actually) consisting of the
      'elements at the current array index of each array in
      'the list.
      '========================================================
140   For lngArrayIndex = 0 To lngMaxUbound
150      arZippedArray = Null
160      For Each vntArrayListElement In ArrayList
170         AddToArray DataArray:=arZippedArray, InputValue:=vntArrayListElement(lngArrayIndex), KeepUnique:=False
180      Next vntArrayListElement
190      AddToArray DataArray:=arZippedArrays, InputValue:=arZippedArray, KeepUnique:=False
200   Next lngArrayIndex

210   If Not IsArray(arZippedArrays) Then Exit Function

      '========================================================
      'NEW CODE: 09/24/2024 Optionally populate referenced
      'dictionary of unzipped items. Keys are index of input
      'arrays having unzipped items. Items in dictionary are
      'arrays of the unzipped items corresponding to the index
      'in the key.
      '========================================================
220   If Not UnzippedItems Is Nothing Then
230      For intArrayIndex = 0 To intArrays
240         vntArrayListElement = ArrayList(intArrayIndex)
250         lngUbound = GetUBound(vntArrayListElement)
260         If lngUbound > lngMaxUbound Then
270            lngLBound = lngMaxUbound + 1
280            UnzippedItems.Item(intArrayIndex) = GetArraySlice(InputArray:=vntArrayListElement, _
                                                                StartIndex:=lngLBound, _
                                                                StopIndex:=lngUbound)
290         End If 'lngUBound > lngMaxUbound
300      Next intArrayIndex
310   End If 'Not UnzippedItems Is Nothing


320   ZipArrays = arZippedArrays

330   Exit Function

ErrProc:
340      ZipArrays = Null
350      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                & Erl & " in procedure ZipArrays of Module " & MODULE_NAME
End Function

