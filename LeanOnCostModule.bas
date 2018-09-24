Attribute VB_Name = "LeanOnCostModule"
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

Public Sub generate_extended_lean_table(ictrl As IRibbonControl)

    
    ' =====================================================================================================
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("extended")
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
    
    On Error Resume Next
    sh.ShowAllData
    
    sh.Range("A1").AutoFilter Field:=28, Criteria1:="=*FMA*", Operator:=xlAnd
    
    lastSourceRow = calcLastSourceRow(sh.Cells(1, 1))
    
    Dim r As Range
    
    Dim leanSh As Worksheet
    Set leanSh = ThisWorkbook.Sheets.Add
    
    
    Dim refLabel As Range
    Set refLabel = leanSh.Range("A1")
    
    
    Dim cfgRef As Range
    Set cfgRef = ThisWorkbook.Sheets("config").Range("extendedStart")
    
    Do
        If cfgRef.Offset(0, 4) <> "" Then
        
            sourceColumn = cfgRef.Row - 1
            leanColumn = CLng(cfgRef.Offset(0, 4))
            qc lastSourceRow, sh, sourceColumn, leanSh, leanColumn
        End If
        
        Set cfgRef = cfgRef.Offset(1, 0)
    Loop Until Trim(cfgRef) = ""
    
    leanSh.Range("A2").AutoFilter
    
    MsgBox "lean oncost table ready!"
    
    ' =====================================================================================================
End Sub

Private Sub qc(lastSourceRow, sh, sourceColumn, leanSh, leanColumn)
    sh.Range(sh.Cells(1, sourceColumn), sh.Cells(lastSourceRow, sourceColumn)).Copy leanSh.Range(leanSh.Cells(1, leanColumn), leanSh.Cells(lastSourceRow + 1, leanColumn))
End Sub



Private Function calcLastSourceRow(r As Range)
    
    If r.Offset(1, 0).Value <> "" Then
        Set r = r.End(xlDown)
        calcLastSourceRow = r.Row
    Else
        calcLastSourceRow = 2
    End If
    
    
End Function
