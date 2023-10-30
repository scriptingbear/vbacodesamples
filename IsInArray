Public Function IsInArray(ByVal InputValues As Variant, ByVal LookupList As Variant, _
                Optional ByRef NonMatchingItems As Variant = Null, _
                Optional ByRef MatchingItems As Variant = Null, _
                Optional ByVal Strict As Boolean = True) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : IsInArray
      ' Author    : Adiv Abramson
      ' Date      : 06/27/2018
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
      ' Versions  : 1.0 - 06/27/2018 - Adiv Abramson
      '           : 1.1 - 08/24/2019 - Adiv Abramson
      '           :
      '           :
      '           :
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


      '<NEW CODE:06/27/2018> If first parameter is an array, iterate through its items
      'and return True only if each item is in LookupList
70    blnIsInputArray = IsArray(InputValues)

80    If blnIsInputArray Then
90       intMatches = -1
100      intInputSize = UBound(InputValues)
110      For i = 0 To intInputSize
            '<NEW CODE:08/24/2019> Optionally capture NonMatching and Matching Items
120         blnMatched = False
130         For j = 0 To intListSize
140            If InputValues(i) = LookupList(j) Then
150               Incr intMatches, 1
160               blnMatched = True
170               Exit For
180            End If 'InputValues(j) = LookupList(i)
190         Next j
200         If Not IsMissing(MatchingItems) Then
210            If blnMatched Then AddToArray MatchingItems, InputValues(i), False
220         End If
230         If Not IsMissing(NonMatchingItems) Then
240            If Not blnMatched Then AddToArray NonMatchingItems, InputValues(i), False
250         End If
            
260      Next i
270      If Strict Then
280         IsInArray = (intMatches = intInputSize)
290      Else
300         IsInArray = (intMatches > -1)
310      End If

320   Else
         'Only a single value needs to be looked up
330      For i = 0 To intListSize
340         If LookupList(i) = InputValues Then
350            IsInArray = True
360            Exit Function
370         End If
380      Next i
390   End If 'blnIsInputArray

400   Exit Function

ErrProc:
410     IsInArray = False
420     MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure IsInArray of Module mdlUtil"

End Function


