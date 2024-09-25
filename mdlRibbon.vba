'---------------------------------------------------------------------------------------
' Module    : mdlRibbon
' Author    : Adiv
' Date      : 00/00/0000
' Purpose   : Contains functions related to display of custom Ribbon tabs, their controls
'           : and event procedures
'           :
'           :
'           :
'           :
'           :
'           :
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


Option Explicit
Option Private Module

Private Const MODULE_NAME = "mdlRibbon"


Public Sub Shared_chk_GetPressed(control As IRibbonControl, ByRef returnedVal)
      '---------------------------------------------------------------------------------------
      ' Procedure : Shared_chk_GetPressed
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that sets the state of checkbox controls
      '           :
      '           :
      '           :
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

20    Select Case control.Tag
         Case "tag1"
30          returnedVal = "Boolean value from global variable or other source"

40    End Select


50    Exit Sub

ErrProc:

60       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_chk_GetPressed of Module " & MODULE_NAME
End Sub



Public Sub Shared_chk_OnAction(control As IRibbonControl, pressed As Boolean)
      '---------------------------------------------------------------------------------------
      ' Procedure : Shared_chk_OnAction
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that handles click events on checkbox controls
      '           :
      '           :
      '           :
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

20    Select Case control.Tag
         Case "tag1"
30          'global variable = pressed

40    End Select


50    Exit Sub

ErrProc:

60       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_chk_OnAction of Module " & MODULE_NAME
End Sub



Public Sub Shared_ddm_GetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
'---------------------------------------------------------------------------------------
' Procedure : Shared_ddm_GetSelectedItemIndex
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Generic code that clears selected item from dropdown menu.
'           : If first item in dropdown menu is a blank (indicating no selection has been made)
'           : setting returnedVal = 0 clears the exiting selection from the dropdown menu.
'           : A global variable corresponding to this control will be set to "" and downstream
'           : code will display warning about missing value.
'           : (1.1) For dropdown menus that DO NOT have a blank entry, we want to ensure that the
'           : corresponding global variables are assigned to the value corresponding to top entry
'           : in the dropdown menu. We need to "read" that value.
'           :
'           :
'           :
'           :
'           :
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

'Constants
'*********************************

'*********************************

On Error GoTo ErrProc

returnedVal = 0


Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_ddm_GetSelectedItemIndex of Module " & MODULE_NAME
End Sub


Public Sub Shared_GetLabel(control As IRibbonControl, ByRef returnedVal)
      '---------------------------------------------------------------------------------------
      ' Procedure : Shared_GetLabel
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Since a dynamic menu doesn't display the selected item, must use a
      '           : paired label control to display it. The button in the dynamic menu that is clicked needs to
      '           : invalidate the associated label control.
      '           :
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


20    If control.ID = "controlID_1" Then
30       returnedVal = ""
40    End If

50    Exit Sub

ErrProc:

60       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_GetLabel of Module " & MODULE_NAME
End Sub


Sub dmnu_getContentTemplate(ByRef control As IRibbonControl, ByRef returnedVal)
'---------------------------------------------------------------------------------------
' Procedure : dmnu_getContentTemplate
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Populates a dynamic control on the Ribbon with an XML string generated
'           : by a <control name>MenuBuilder() function
'           :
'           :
'           :
'           :
'           :
'Versions   : 1.0 - 00/00/0000 - Adiv Abramson
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

'Variants
'*********************************

'*********************************

'Booleans
'*********************************

'*********************************

'Constants
'*********************************

'*********************************


On Error GoTo ErrProc

returnedVal = ""

'returnedVal = SomeMenuBuilder()
   
Exit Sub

ErrProc:
    returnedVal = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure dmnu_getContentTemplate of Module mdlRibbon"


End Sub
Public Function GenericeMenuBuilder() As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GenericeMenuBuilder
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that creates XML code that will populate a dynamicMenu control's
      '           : getContent() callback procedure. Copy this proc and rename it, e.g. "MyItemsMenuBuilder"
      '           :
      '           : strXML = strXML & ButtonBuilder(ButtonID:= "", ButtonLabel:= "", ImageMso:= "", CallBack:= "") & " "
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
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
      '           :
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strXML As String
      Dim strItem As String
      Dim strButtons As String
      '*********************************

      'Numerics:
      '*********************************
      Dim i As Integer
      '*********************************

      'Worksheets:
      '*********************************
      Dim wsItems As Worksheet
      '*********************************

      'Workbooks:
      '*********************************

      '*********************************

      'Ranges:
      '*********************************
      Dim rngItems As Range
      Dim rngCell As Range
      '*********************************

      'Arrays:
      '*********************************

      '*********************************

      'Objects:
      '*********************************
      Dim objItemsList As ListObject
      '*********************************

      'Variants
      '*********************************

      '*********************************

      'Booleans
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const XML_HEADER = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >"
      Const XML_FOOTER = "</menu>"
      '*********************************


10    On Error GoTo ErrProc

20    GenericeMenuBuilder = ""


      '========================================================
      'Reference table containing items for the dynamic menu
      '========================================================
30    Set wsItems = ThisWorkbook.Worksheets("Items WorksheetNames")
40    Set objItemsList = wsItems.ListObjects("Items Table")
50    Set rngItems = objItemsList.DataBodyRange


60    strXML = XML_HEADER & " "

      '====================================================================
      'Put code here to generate buttons for the dynamicMenu control, making repeated
      'calls to ButtonBuilder and adding the returned strings to the existing XML
      'strXML = strXML & ButtonBuilder(ButtonID:= "", ButtonLabel:= "", ImageMso:= "", CallBack:= "", [Tag:= ""]) & " "
      '====================================================================
70    strButtons = ""
80    i = 1

90    For Each rngCell In rngItems.Cells
100      strItem = rngCell.Value
110      strButtons = strButtons & vbCrLf _
         & ButtonBuilder(ButtonID:="btnItem" & i, _
                         ButtonLabel:=strItem, _
                         ImageMso:="DataGraphicNewItem", _
                         AddCallBack:=True, _
                         OverrideCallBack:="Shared_btn_OnAction", _
                         Tag:=strItem)
120     Incr i, 1
130   Next rngCell

140   strXML = strXML & strButtons
150   strXML = strXML & XML_FOOTER

160   GenericeMenuBuilder = strXML


170   Exit Function

ErrProc:
180        GenericeMenuBuilder = ""
190       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure GenericeMenuBuilder of Module mdlRibbon"
End Function

Public Sub cboControl_onChange(control As IRibbonControl, ByRef text)
      '---------------------------------------------------------------------------------------
      ' Procedure : cboControl_onChange
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that reads the value in the textbox portion of the combobox control
      '           : and assigns it to global variable.
      '           :
      '           :
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

20    'm_strGenericVar = text


30    Exit Sub

ErrProc:

40       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure cboControl_onChange of Module " & MODULE_NAME
End Sub


Public Sub cboControl_GetText(control As IRibbonControl, ByRef returnedVal)
'---------------------------------------------------------------------------------------
' Procedure : cboControl_GetText
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Generic code that sets or clears the string in the combobox control.
'           :
'           : NOT IN USE CURRENTLY
'           :
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

On Error GoTo ErrProc



Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure cboControl_GetText of Module " & MODULE_NAME
End Sub


Public Sub cboControl_GetItemCount(control As IRibbonControl, ByRef returnedVal)
      '---------------------------------------------------------------------------------------
      ' Procedure : cboControl_GetItemCount
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that determines number of items in combobox menu, which may come from a range or
      '           : table or global variable, for example.
      '           :
      '           :
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


20    returnedVal = "set the count of items in combobox control"

30    Exit Sub

ErrProc:

40       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure cboControl_GetItemCount of Module " & MODULE_NAME
End Sub


Public Sub cboCrontrol_GetItemID(control As IRibbonControl, itemIndex As Integer, ByRef returnedVal)
      '---------------------------------------------------------------------------------------
      ' Procedure : cboCrontrol_GetItemID
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that assigns an id for each dynamically added combobox menu item
      '           : itemIndex starts at 0!
      '           :
      '           :
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

20    returnedVal = "prefic" & itemIndex


30    Exit Sub

ErrProc:

40       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure cboCrontrol_GetItemID of Module " & MODULE_NAME
End Sub



Public Sub cboControl_GetItemLabel(control As IRibbonControl, itemIndex As Integer, ByRef returnedVal)
      '---------------------------------------------------------------------------------------
      ' Procedure : cboControl_GetItemLabel
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Generic code that assigns label to each dynamically added combobox menu item.
      '           : itemIndex starts at 0!
      '           :
      '           :
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

20    returnedVal = Range("range name").Cells(itemIndex + 1, 1).Value
    

30    Exit Sub

ErrProc:

40       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure cboControl_GetItemLabel of Module " & MODULE_NAME
End Sub
Sub Shared_txt_OnChange(control As IRibbonControl, text As String)
'---------------------------------------------------------------------------------------
' Procedure : Shared_txt_OnChange
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Capture value on edit box control in global variable.
'           : Set global variable to 0 if of type Integer, which will cause the Validate<reading type>()
'           : functions to return False, which prevents reuse of previous values.
'           : Now using new function mdlUtil::GetNumericValue(), which returns 0 if text isn't a number
'           : and returns Integer by default.
'           :
'           :
'           :
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
Dim vntValue As Variant
'*********************************

'Booleans:
'*********************************

'*********************************

'Constants
'*********************************

'*********************************

      
10    On Error GoTo ErrProc

20    Select Case control.Tag
         Case "Module Name"
40          m_strModuleName = text
         Case "Procedure Name"
60          m_strProcName = text
70    End Select



Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_txt_OnChange of Module " & MODULE_NAME
End Sub

Public Sub Shared_btn_OnAction(ByRef control As IRibbonControl)
'---------------------------------------------------------------------------------------
' Procedure : Shared_btn_OnAction
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Shared callback for processing OnAction events of Ribbon button controls.
'           : (1.1) Using CleanString() to strip initial "=" and quotation marks that are
'           : returned from Name("MyName").Value, e.g. returns ="MyTextValue" but we want
'           : the string MyTextValue, with no "=" and no delimiting quotation marks.
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
'---------------------------------------------------------------------------------------

'Strings:
'*********************************

'*********************************

'Numerics:
'*********************************

'*********************************

'Worksheets:
'*********************************
Dim wsCodeHelper As Worksheet
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

'*********************************

'Booleans
'*********************************

'*********************************

'Constants
'*********************************

'*********************************




10    On Error GoTo ErrProc
      
      
20    Select Case control.Tag
         Case "Apply Code Helper Settings"
            '========================================================
            'The OnChange callback seems to be triggered when the
            'editBox control loses focus, which occurs when the user
            'clicks the Apply button. Not sure if anything
            'needs to be done since the global variables will already
            'have been updated. Similar arguments apply for the OnAction
            'callback for the dropdown menu controls.
            'Seems like we need to cause the worksheet to recalculate
            '========================================================
130         Set wsCodeHelper = ThisWorkbook.Worksheets(WorksheetNames.CodeHelper)
140         wsCodeHelper.Calculate
         
         
         Case ""
180         
         
         Case "Paste Line Number Generator Table"
210         
         
         Case "Copy Line Number Generator Table"
250         
260         
         
         
         Case "Clear Strip Line Numbers Table"
300         
         
         
         Case "Paste Strip Line Numbers Table"
340         
         
         Case "Clear Code Only Table"
380         
         
         Case "Copy Code Only Table"
410         
420         
         
         
         Case Else
         
            '========================================================
            'In order to use the Like operator.
            'use Select Case True and Case <exp> Like <exp>
            '========================================================
         '   Select Case True
         '      Case control.ID Like "pattern"
         '         '<invoke function>
         '   End Select
      
550   End Select





   
Exit Sub

ErrProc:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_btn_OnAction of Module " & MODULE_NAME


End Sub

Sub Shared_ddm_OnAction(control As IRibbonControl, itemID As String, itemIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : Shared_ddm_OnAction
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Generic code that invokes procedure corresponding to item selected from
'           : a dropdown menu. Can use button's id property or its index to
'           : determine which procedure to invoke.
'           :
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
      
20    Select Case itemID
      
      '========================================================
      ' Invoke procedure corresponding to selected dropdown menu
      ' item.
      '========================================================
      Case ""
40       
      
      Case ""
60       
      
      Case ""
80       
      
      Case ""
100      
      
      Case ""
120      
      
      Case ""
140      
      
150   End Select



Exit Sub

ErrProc:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure Shared_ddm_OnAction of Module " & MODULE_NAME
End Sub

Sub dmnu_getContent(ByRef control As IRibbonControl, ByRef returnedVal)
'---------------------------------------------------------------------------------------
' Procedure : dmnu_getContent
' Author    : Adiv Abramson
' Date      : 9/1/2019
' Purpose   : Populate a dynamic control on the Ribbon with an XML string generated
'           : by a <control name>MenuBuilder() function
'           :
'           :
'           :
'           :
'           :
'Versions   : 1.0 - 9/1/2019 - Adiv Abramson
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

'Variants
'*********************************

'*********************************

'Booleans
'*********************************

'*********************************

'Constants
'*********************************

'*********************************


On Error GoTo ErrProc

returnedVal = ""

'returnedVal = SomeMenuBuilder()
   
Exit Sub

ErrProc:
    returnedVal = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure dmnu_getContent of Module mdlRibbon"


End Sub
Public Function DynamicMenuBuilderTemplate() As String
'---------------------------------------------------------------------------------------
' Procedure : DynamicMenuBuilderTemplate
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Creates XML code that will be provided to a dynamicMenu control's
'           : getContent() callback procedure.
'           :
'           : strXML = strXML & ButtonBuilder(ButtonID:= "", ButtonLabel:= "", ImageMso:= "", CallBack:= "") & " "
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
'           :
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
'           :
'           :
'---------------------------------------------------------------------------------------

'Strings:
'*********************************
Dim strXML As String
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

'Variants
'*********************************

'*********************************

'Booleans
'*********************************

'*********************************

'Constants
'*********************************
Const XML_HEADER = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >"
Const XML_FOOTER = "</menu>"
'*********************************


On Error GoTo ErrProc

DynamicMenuBuilderTemplate = ""

strXML = XML_HEADER & " "

'====================================================================
'Put code here to generate buttons for the dynamicMenu control, making repeated
'calls to ButtonBuilder and adding the returned strings to the existing XML
'strXML = strXML & ButtonBuilder(ButtonID:= "", ButtonLabel:= "", ImageMso:= "", CallBack:= "", [Tag:= ""]) & " "
'====================================================================

strXML = strXML & XML_FOOTER


DynamicMenuBuilderTemplate = strXML


Exit Function

ErrProc:
    DynamicMenuBuilderTemplate = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure DynamicMenuBuilder of Module mdlRibbon"
End Function



Public Function ButtonBuilder(ByVal ButtonID As String, ByVal ButtonLabel As String, ByVal ImageMso As String, _
       Optional ByVal AddCallBack As Boolean = True, Optional ByVal OverrideCallBack As String = "", _
       Optional Tag As String = "") As String
      '---------------------------------------------------------------------------------------
      ' Procedure : ButtonBuilder
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Returns the XML for a button object that will be part of a dynamic menu
      '           : control on the ribbon
      '           :
      '           :
      '           :
      '           :
      '           :
      '           :
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
      '           :
      '---------------------------------------------------------------------------------------

      'Strings:
      '*********************************
      Dim strXML As String
      Dim strCallBack As String
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

20    If AddCallBack And OverrideCallBack = "" Then
30       strCallBack = "BACKTICK" & ButtonID & "_onClick" & "BACKTICK"
40    ElseIf AddCallBack And OverrideCallBack <> "" Then
50       strCallBack = "BACKTICK" & OverrideCallBack & "BACKTICK"
60    End If

70    strXML = "<button id=BACKTICK" & ButtonID & "BACKTICK "
80    strXML = strXML & "label=BACKTICK" & ButtonLabel & "BACKTICK "
90    strXML = strXML & "imageMso=BACKTICK" & ImageMso & "BACKTICK "
100   If Tag <> "" Then
110      strXML = strXML & "tag=BACKTICK" & Tag & "BACKTICK "
120   End If
130   If strCallBack <> "" Then
140      strXML = strXML & "onAction=" & strCallBack & " />"
150   Else
160      strXML = strXML & "/>"
170   End If

180   strXML = Replace(strXML, "BACKTICK", BACKTICK)

190   ButtonBuilder = EmbedQuotes(strXML)

200   Exit Function

ErrProc:

210       ButtonBuilder = ""
220       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure ButtonBuilder of Module mdlRibbon"
End Function


Public Sub btnAbout_onClick(ByRef control As IRibbonControl)
'---------------------------------------------------------------------------------------
' Procedure : btnAbout_onClick
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Display information about app
'           : (1.1) Now storing version info as custom properties of ThisWorkbook
'           : NOTE: Must include LinkToContent argument or property will NOT be added
'           : to workbook!
'           :
'           : Checked parameter passing type and return data type where applicable
'           :
'  Versions : 1.1 - 00/00/0000 - Adiv Abramson
'           : 1.2 - 00/00/0000 - Adiv Abramson
'           :
'           :
'---------------------------------------------------------------------------------------

Dim strVersionNumber As String
Dim dteVersionDate As Date


On Error GoTo ErrProc


strVersionNumber = ThisWorkbook.CustomDocumentProperties("VersionNumber").Value
dteVersionDate = ThisWorkbook.CustomDocumentProperties("VersionDate").Value

msgInfo "Study Design Application" & vbCrLf & "Version " & strVersionNumber & vbCrLf & dteVersionDate

Exit Sub

ErrProc:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & CStr(Erl) & " in procedure btnAbout_onClick of Module Utilities"
End Sub


Public Sub Ribbon_onLoad(ribbon As IRibbonUI)
'---------------------------------------------------------------------------------------
' Procedure : Ribbon_onLoad
' Author    : Adiv
' Date      : 00/00/0000
' Purpose   :
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



On Error GoTo ErrProc


Set objRibbonUI = ribbon
'objRibbonUI.ActivateTab "<tab name>"



Exit Sub

ErrProc:
  
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in line " & Erl & " of procedure Ribbon_onLoad of Module mdlRibbon"
End Sub

Public Sub MaximizeRibbon()
      '---------------------------------------------------------------------------------------
      ' Procedure : MaximizeRibbon
      ' Author    : Adiv Abramson
      ' Date      : 00/00/0000
      ' Purpose   : Force Ribbon to expand, if collapsed. Otherwise do nothing to it.
      '           :
      '           :
      '           :
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

      '*********************************

      'Variants:
      '*********************************

      '*********************************

      'Booleans:
      '*********************************

      '*********************************

      'Constants
      '*********************************
      Const MIN_RIBBON_HEIGHT = 200
      '*********************************

10    On Error GoTo ErrProc


20    If Application.CommandBars("Ribbon").Height < MIN_RIBBON_HEIGHT Then
30       Application.CommandBars.ExecuteMso "MinimizeRibbon"
40    End If



50    Exit Sub

ErrProc:

60       MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl & " in procedure MaximizeRibbon of Module " & MODULE_NAME
End Sub
