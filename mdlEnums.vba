'---------------------------------------------------------------------------------------
' Module    : mdlEnums
' Author    : Adiv Abramson
' Date      : 05/07/2024
' Purpose   : Contains enumerations used by various procedures and functions.
'           :
'           :
'           :
'           :
'           :
'           :
' Versions  : 1.0 - 05/07/2024 - Adiv Abramson
'           :
'           :
'           :
'           :
'           :
'           :
'---------------------------------------------------------------------------------------

Private Const MODULE_NAME = "mdlEnums"


Public Enum GetCharsType
  Alpha = 0
  Numeric = 1
End Enum

Public Enum FeatureNotImplementedType
   ndcinput = 1
   GPI_Input = 2
End Enum


'<NEW CODE:06/25/2018>
Public Enum HeaderInfo
   HeaderInfoColumn
   HeaderInfoIndex
   HeaderInfoColumnAndIndex
   HeaderInfoReference
End Enum

Public Enum QueryLogicOperator
   LogicalOR
   LogicalAND
End Enum

Public Enum TabNavigation
   MoveToFirstColumn
   MoveToLastColumn
   MoveToFirstRow
   MoveToLastRow
End Enum


Public Enum ComparisonOperator
   IsEqualTo = 0
   IsGreaterThan = 1
   IsGreaterThanOrEqualTo = 2
   IsLessThan = 3
   IsLessThanOrEqualTo = 4
   IsNotEqualTo = 5
End Enum


Public Enum CompareType
   ValueIsEqualTo
   ValueIsLessThan
   ValueIsLessThanOrEqualTo
   ValueIsGreaterThan
   ValueIsGreaterThanOrEqualTo
   ValueIsBetween_OpenLeft
   ValueIsBetween_OpenRight
   valueisbetween
   ValueIsUnset
End Enum



Public Enum DataRecordType
   VisibleSheetsRecord
   PTInfoRecord
   SubjectRecord
   UnsetRecord
End Enum
