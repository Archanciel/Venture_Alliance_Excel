Attribute VB_Name = "DailySiteStats"
Option Explicit

Sub retrieveData()
Attribute retrieveData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' update Macro
'

'
    Dim beforeLastRowHeadCell As Range
    Dim lastRowHeadCell As Range
    Dim nextRowHeadCell As Range
    Dim lastRowRange As Range
    Dim nextRowRange As Range
    
    Set lastRowHeadCell = Cells(getLastDataRow(Range("A1:A1")), 1)
    Set beforeLastRowHeadCell = lastRowHeadCell.Offset(-1, 0)
    
    If (sameDate(beforeLastRowHeadCell.Value, DateTime.Now)) Then
        'si une entrée existe déjà pour la date courante, cette ligne est supprimée.
        'Elle sera remplacée par l'exécution de la suite de la macro !
        Rows(beforeLastRowHeadCell.Row).Select
        Selection.Delete Shift:=xlUp
        Set lastRowHeadCell = Cells(getLastDataRow(Range("A1:A1")), 1)
    End If
    
    Set lastRowRange = Range(lastRowHeadCell, lastRowHeadCell.Offset(0, 5))
    
    lastRowRange.Select
    Selection.Copy
    
    Set nextRowHeadCell = lastRowHeadCell.Offset(1, 0)
    nextRowHeadCell.Select
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Set nextRowRange = Range(nextRowHeadCell, nextRowHeadCell.Offset(0, 5))
    
    nextRowRange.Select
    Selection.Copy
    lastRowRange.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    clearAnySelection
End Sub
Private Function getLastDataRow(colCell As Range) As Long
    Dim lastCell As Range
    Dim lastCellRow As Long
    
    Set lastCell = colCell.End(xlDown)
    getLastDataRow = lastCell.Row
End Function
Private Sub clearAnySelection()
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
End Sub
Private Function sameDate(dateStr As String, today As Date) As Boolean
    sameDate = Day(dateStr) = Day(today) And Month(dateStr) = Month(today) And Year(dateStr) = Year(today)
End Function

