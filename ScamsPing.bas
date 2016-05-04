Attribute VB_Name = "ScamsPing"
Option Explicit

Function isEmptyCell(r As Range) As Boolean
'
' Rempl1 Macro
'

'
    isEmptyCell = Trim(r.Value & vbNullString) = vbNullString
End Function

