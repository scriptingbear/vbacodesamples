'---------------------------------------------------------------------------------------
' Module    : Globals
' Author    : Adiv Abramson
' Date      : 00/00/0000
' Purpose   : Contains Property Get procedures to return global variables. More
'           : maintainable and flexible than hard coded constants.
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

Option Explicit

Private Const MODULE_NAME = "Globals"


Public Property Get APP_NAME() As String
10       APP_NAME = "<app name>"
End Property

Public Property Get DELIMITER() As String
10       DELIMITER = "^"
End Property

Public Property Get LAST_WORKSHEET_ROW() As Long
10       LAST_WORKSHEET_ROW = 1048576
End Property


Public Property Get LAST_WORKSHEET_COL() As Long
10       LAST_WORKSHEET_COL = 16384
End Property

Public Property Get COLUMN_HEADING() As Long
10       COLUMN_HEADING = 0
End Property

Public Property Get COLUMN_INDEX() As Long
10       COLUMN_INDEX = 1
End Property

Public Property Get COLUMN_RANGE() As Long
10       COLUMN_RANGE = 3
End Property

Public Property Get DATA_RANGE() As String
10       DATA_RANGE = "Data Range"
End Property

Public Property Get INCLUDE_HEADERS() As Long
10       INCLUDE_HEADERS = 0
End Property

Public Property Get EXCLUDE_HEADERS() As Long
10       EXCLUDE_HEADERS = 1
End Property

Public Property Get BACKTICK() As String
10       BACKTICK = "`"
End Property
Property Get OPTION_NONE() As String
10       OPTION_NONE = "(None)"
End Property

Public Property Get QUOTE() As String
10       QUOTE = """"
End Property

Public Property Get WORD_FILE_TYPES_FILTER(Optional AllTypes As Boolean = False) As String
10       If AllTypes Then
20          WORD_FILE_TYPES_FILTER = "Word Documents (*.docx;*.docm), *.docx;*.docm"
30       Else
40          WORD_FILE_TYPES_FILTER = "Word Documents (*.docx), *.docx"
50       End If
End Property


Public Property Get EXCEL_FILE_TYPES_FILTER(Optional AllTypes As Boolean = False) As String
10       If AllTypes Then
20          EXCEL_FILE_TYPES_FILTER = "Excel Documents (*.xlsx;*.xlsm), *.xlsx;*.xlsm"
30       Else
40          EXCEL_FILE_TYPES_FILTER = "Excel Documents (*.xlsx), *.xlsx"
50       End If
End Property

Public Property Get TAG_DELIMITER() As String
10       TAG_DELIMITER = "|"
End Property


