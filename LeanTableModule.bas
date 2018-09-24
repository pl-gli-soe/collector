Attribute VB_Name = "LeanTableModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Public Sub generate_lean_table(ictrl As IRibbonControl)
    
    ' =====================================================================================================
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("allSource")
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
    
    On Error Resume Next
        sh.ShowAllData
    
    sh.Range("A1").AutoFilter Field:=9, Criteria1:="=*FMA*", Operator:=xlAnd
    
    Dim r As Range
    
    
    Dim leanSh As Worksheet
    Set leanSh = ThisWorkbook.Sheets.Add
    
    With leanSh
        .Cells(1, 1).Value = "DUNS"
        
        .Cells(2, 1).Value = "FAZA"
        .Cells(2, 2).Value = "PN"
        .Cells(2, 3).Value = "PCD PN"
        .Cells(2, 4).Value = "Part Name"
        .Cells(2, 5).Value = "DUNS"
        .Cells(2, 6).Value = "PICK UP DATE"
        .Cells(2, 7).Value = "MRD"
        .Cells(2, 8).Value = "Ordered Date"
        .Cells(2, 9).Value = "Ordered Qty"
        .Cells(2, 10).Value = "Confirmed Qty"
        
    End With
    
    
    lastSourceRow = calcLastSourceRow(sh.Cells(1, 1))
    
    
    ' faza
    sourceColumn = 5
    leanColumn = 1
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    
    
    
    ' pn
    sourceColumn = 11
    leanColumn = 2
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 13
    
    
    ' alter pn
    sourceColumn = 31
    leanColumn = 3
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 13
    
    ' gdps part name
    'sourceColumn = 32
    'leanColumn = 4
    'qc lastSourceRow, Sh, sourceColumn, leanSh, leanColumn
    'leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 17
    
    ' part name
    sourceColumn = 33
    leanColumn = 4
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 17
    
    
    ' DUNS
    sourceColumn = 14
    leanColumn = 5
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 11
    
    
    
    ' pu date
    sourceColumn = 23
    leanColumn = 6
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 13
    
    ' mrd
    sourceColumn = 6
    leanColumn = 7
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 12
    
    
    ' ordered date
    sourceColumn = 17
    leanColumn = 8
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 13
    
    
    ' ordered qty
    sourceColumn = 20
    leanColumn = 9
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 12
    
    
    ' conf qty
    sourceColumn = 21
    leanColumn = 10
    qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
    leanSh.Cells(1, leanColumn).EntireColumn.ColumnWidth = 12
    
    leanSh.Range("A2").AutoFilter
    
    
    'If leanSh.Range("b1").Value <> "" Then
    '    leanSh.Range("A2").AutoFilter Field:=5, Criteria1:="=*" & CStr(leanSh.Range("B1")) & "*"
    'Else
    '    leanSh.Range("A2").AutoFilter
    'End If
    
    ' =====================================================================================================
End Sub

Private Sub qc(lastSourceRow, sh, sourceColumn, leanSh, leanColumn)
    sh.Range(sh.Cells(2, sourceColumn), sh.Cells(lastSourceRow, sourceColumn)).Copy leanSh.Range(leanSh.Cells(3, leanColumn), leanSh.Cells(lastSourceRow + 1, leanColumn))
End Sub



Private Function calcLastSourceRow(r As Range)
    
    If r.Offset(1, 0).Value <> "" Then
        Set r = r.End(xlDown)
        calcLastSourceRow = r.Row
    Else
        calcLastSourceRow = 2
    End If
    
    
End Function
