'---------------------------------------------------------------------------------------
' Module    : mdlSystemDialogs
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   :
'           :
'           :
'           :
' Versions  : 1.0 - 00/00/0000
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Option Explicit
Option Private Module



Public Function msgAsk(ByVal strMess As String, Optional ByVal vntCustomCaption) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : msgAsk
      ' Author    : Modified by Adiv Abramson on 00/00/0000 at 11:10
      ' Date      : 00/00/0000
      ' Purpose   : (1.1) Add optional custom caption parameter to override default caption
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '  Versions : 1.1 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------



      Dim intResponse As Integer

10    On Error GoTo ErrProc


      '00/00/0000>NEW CODE: Display alternative custom caption, if desired
20    intResponse = MsgBox(Prompt:=strMess, Buttons:=vbExclamation + vbYesNo, Title:=IIf(IsMissing(vntCustomCaption), _
                    Globals.APP_NAME, vntCustomCaption))
                    
30    If intResponse = vbNo Then
40       msgAsk = False
50    Else
60       msgAsk = True
70    End If

80    Exit Function

ErrProc:

90        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure msgAsk of Module mdlSystemDialogs"

End Function


Public Sub msgAttention(ByVal strText As String, Optional ByVal vntCustomCaption)
'---------------------------------------------------------------------------------------
' Procedure : msgAttention
' Author    : Modified by Adiv Abramson on 00/00/0000 at 11:10
' Date      : 00/00/0000
' Purpose   : (1.1) Add optional custom caption parameter to override default caption
'           :
'           :
'           :
'           :
'           :
'           :
'  Versions : 1.1 - 00/00/0000 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------


10    On Error GoTo ErrProc
      
      '00/00/0000>NEW CODE: Display alternative custom caption, if desired
      MsgBox Prompt:=strText, Buttons:=vbExclamation, Title:=IIf(IsMissing(vntCustomCaption), _
                    Globals.APP_NAME, vntCustomCaption)


Exit Sub

ErrProc:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure msgAttention of Module mdlSystemDialogs"
End Sub





Public Sub msgInfo(ByVal strText As String, Optional ByVal blnDisplay = True, Optional ByVal vntCustomCaption)
      '---------------------------------------------------------------------------------------
      ' Procedure : msgInfo
      ' Author    : Modified by Adiv Abramson on 00/00/0000 at 11:10
      ' Date      : 00/00/0000
      ' Purpose   : (1.2) Add optional parameter to suppress display of message as needed
      '           :
      '           :
      '           :
      '           :
      '           :
      '  Versions : 1.1 - 00/00/0000 - Adiv Abramson
      '           : 1.2 - 00/00/0000 - Adiv Abramson
      '           :
      '           :
      '---------------------------------------------------------------------------------------


10    On Error GoTo ErrProc


20    If blnDisplay Then
         '00/00/0000>NEW CODE: Display alternative custom caption, if desired
30       MsgBox Prompt:=strText, Buttons:=vbInformation, Title:=IIf(IsMissing(vntCustomCaption), _
                    Globals.APP_NAME, vntCustomCaption)
40    End If

50    Exit Sub

ErrProc:

60        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure msgInfo of Module mdlSystemDialogs"
End Sub






Public Sub msgStop(ByVal strText As String, Optional ByVal vntCustomCaption)
      '---------------------------------------------------------------------------------------
      ' Procedure : msgStop
      ' Author    : Modified by Adiv Abramson on 00/00/0000 at 11:10
      ' Date      : 00/00/0000
      ' Purpose   :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------


10    On Error GoTo ErrProc


      '00/00/0000>NEW CODE: Display alternative custom caption, if desired
20    MsgBox Prompt:=strText, Buttons:=vbCritical, Title:=IIf(IsMissing(vntCustomCaption), _
                Globals.APP_NAME, vntCustomCaption)

30    Exit Sub

ErrProc:

40        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure msgStop of Module mdlSystemDialogs"
End Sub


Public Function msgYesNoCancel(ByVal strText As String, Optional ByVal vntCustomCaption) As Integer
      '---------------------------------------------------------------------------------------
      ' Procedure : msgYesNoCancel
      ' Author    : Modified by Adiv Abramson on 00/00/0000 at 11:10
      ' Date      : 00/00/0000
      ' Purpose   :
      '           :
      '           :
      '           :
      '           :
      '---------------------------------------------------------------------------------------



10    On Error GoTo ErrProc


20    msgYesNoCancel = MsgBox(Prompt:=strText, Buttons:=vbYesNoCancel + vbQuestion, Title:=Globals.APP_NAME)

30    Exit Function

ErrProc:

40        MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure msgYesNoCancel of Module mdlSystemDialogs"
End Function




