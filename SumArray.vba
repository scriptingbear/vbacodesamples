Public Function SumArray(ByVal DataArray As Variant) As Double
      '---------------------------------------------------------------------------------------
      ' Procedure : SumArray
      ' Author    : Adiv Abramson
      ' Date      : 08/20/2024
      ' Purpose   : Sum an array of values without looping through it.
      '           : Use JOIN() to delimit array elements with "+", then
      '           : use EVAL() to return the result of the summation.
      '           : This is for MS Excel. For MS Access, use Eval().
      '           : Be aware of the 255 character limit of Evaluate().
      '           : Using Regex to parse input into substrings compatible
      '           : with Evaluate(). Seems to work OK for arrays up to ~ 200K elements
      '           : but slows down dramatically after that.
      '           : So some looping still occurs, just not over individual array elements.
      '           :
      '           :
      '           :
      '           :
      '           :
      ' Versions  : 1.0 - 08/20/2024 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strJoinedArrayItems As String
      Dim strSubArrayValues As String
      '*********************************

      'Numerics:
      '*********************************
      Dim lngEvalExpressionLength As Long
      Dim dblTotal As Double
      Dim dblSubTotal As Double
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
      Dim objMatches As MatchCollection
      Dim objMatch As Match
      Dim objSubmatches As SubMatches
      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const MAX_EXP_LEN As Integer = 255
      Const SUBSTRING_PATTERN As String = "^[0-9\.\+]{1,255}(?=(\+|$))"
      '*********************************

10    On Error GoTo ErrProc

20    SumArray = 0

      '========================================================
      'Account for joined string in excess of 255 characters
      '========================================================
30    strJoinedArrayItems = Join(DataArray, "+")
40    lngEvalExpressionLength = Len(strJoinedArrayItems)

50    If lngEvalExpressionLength <= MAX_EXP_LEN Then
60       SumArray = Evaluate(strJoinedArrayItems)
70       Exit Function
80    End If

90    dblTotal = 0
      '========================================================
      'Must break string of joined array items into 255 character
      'substrings and accumulate the total.
      '========================================================
100   Set objRegex = New RegExp
110   With objRegex
120      .Global = False
130      .IgnoreCase = True
140      .PATTERN = SUBSTRING_PATTERN
150      While .Test(strJoinedArrayItems)
            '========================================================
            'Get the first matching substring, pass it to Evaluate(),
            'then discard and adjust strJoinedArrayItems
            '========================================================
160         Set objMatches = .Execute(strJoinedArrayItems)
170         Set objMatch = objMatches(0)
180         strSubArrayValues = objMatch.Value
190         dblSubTotal = Evaluate(strSubArrayValues)
200         dblTotal = dblTotal + dblSubTotal
210         strJoinedArrayItems = .Replace(strJoinedArrayItems, "")
220      Wend
230   End With

240   SumArray = dblTotal


250   Exit Function

ErrProc:
260      SumArray = 0
270      MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " _
                 & Erl & " in procedure SumArray of Module " & MODULE_NAME
End Function

