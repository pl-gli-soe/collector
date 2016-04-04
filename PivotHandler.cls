VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
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



Private pivot_cache As PivotCache
Private pivot_table As PivotTable
Private source_range As Range
Private source_sh As Worksheet
Private pivot_sh As Worksheet
Private slajser_cache1 As SlicerCache
Private slajser_cache2 As SlicerCache
Private slajser_cache3 As SlicerCache
Private slajser_cache4 As SlicerCache
Private slajser_cache5 As SlicerCache




Public Sub init(sr As Range)


    Set source_range = sr
    
    ' Set source_range = sr
    Set pivot_cache = ThisWorkbook.PivotCaches.Create(xlDatabase, source_range)
    
    Set source_sh = ThisWorkbook.Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME)
    
    Set pivot_sh = ThisWorkbook.Sheets(XWiz.PIVOT_SHEET_NAME)

    
    Set pivot_table = pivot_sh.PivotTables.Add(PivotCache:=pivot_cache, TableDestination:=Range("B4"))
    
End Sub

Public Sub config_pivot()



    ' .Cells(x, 1) = "PLT"
    ' .Cells(x, 2) = "PROJ"
    ' .Cells(x, 3) = "BG"
    ' .Cells(x, 4) = "MY"
    ' .Cells(x, 5) = "FAZA"
    ' .Cells(x, 6) = "MRD"
    ' .Cells(x, 7) = "COORD"
    ' .Cells(x, 8) = "RESP"
    ' .Cells(x, 9) = "FUP"
    ' .Cells(x, 10) = "PN"
    ' .Cells(x, 11) = "DEL CONF"
    ' .Cells(x, 12) = "Comment"

    Dim pivot_field As PivotField
    
    With pivot_table
        .PivotFields("DEL CONF").Orientation = xlRowField
        .PivotFields("DEL CONF").Position = 1
        
        '.PivotFields("PROJ").Orientation = xlRowField
        '.PivotFields("PROJ").Position = 2
        
        
        
        .PivotFields("MRD").Orientation = xlColumnField
        .PivotFields("MRD").Position = 1
        
        
        Set pivot_field = .PivotFields("PROJ")
        pivot_field.Orientation = xlDataField
        pivot_field.Function = xlCount
        
        ' .PivotFields("DEL CONF").ShowDetail = False
        
        
        '.PivotFields("PROJ").Orientation = xlPageField
        '.PivotFields("PROJ").Position = 1
        '.PivotFields("COORD").Orientation = xlPageField
        '.PivotFields("COORD").Position = 2
        
        
        'With .PivotFields("PROJ")
        '    .ClearAllFilters
        '    .EnableMultiplePageItems = True
        '    .CurrentPage = "(All)"
        'End With
        
        'With .PivotFields("FAZA")
        '    .ClearAllFilters
        '    .EnableMultiplePageItems = True
        '    .CurrentPage = "(All)"
        'End With
        
        .ColumnGrand = False
        .RowGrand = False
        
        .TableStyle2 = "PivotStyleMedium15"
    End With
End Sub

Public Sub add_slicers()

    Set slajser_cache1 = Nothing
    Set slajser_cache2 = Nothing
    Set slajser_cache3 = Nothing
    Set slajser_cache4 = Nothing
    Set slajser_cache5 = Nothing
    
    ' tutaj ino troche sliska sprawa z tymi on error resume next
    ' zabopiega oczywiscie pojawieniu sie bledu jakiegokolwiek zwiazanego z slicerami
    ' i jedyne co sie wtedy stanie to brak na pivocie naszych slicerow
    ' hmmmm ale na dluzsza mete trzeba bedzie o tym jeszcze pomyslec jak ewentualne inne bledy zalatwic
    ' bo jednak jest to nieeleganckie rozwiazanie
    ' 2014-07-01

    On Error Resume Next
    Set slajser_cache1 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "COORD")
    slajser_cache1.Slicers.Add pivot_sh, , "COORD" & CStr(pivot_sh.Name), "Coordinator Slicer", 100, 400, 150, 200
        
    On Error Resume Next
    Set slajser_cache2 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "FUP")
    slajser_cache2.Slicers.Add pivot_sh, , "FUP" & CStr(pivot_sh.Name), "FUP Slicer", 140, 500, 200, 200
    
    On Error Resume Next
    Set slajser_cache3 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "PROJ")
    slajser_cache3.Slicers.Add pivot_sh, , "PROJ" & CStr(pivot_sh.Name), "Proj", 180, 600, 250, 250
    
    On Error Resume Next
    Set slajser_cache4 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "FAZA")
    slajser_cache4.Slicers.Add pivot_sh, , "FAZA" & CStr(pivot_sh.Name), "Faza", 220, 700, 300, 300
        
    On Error Resume Next
    Set slajser_cache5 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "MRDd", , xlTimeline)
    slajser_cache5.Slicers.Add pivot_sh, , "MRDd", "MRDd", 50, 350, 400, 100
    
    
End Sub

'Public Sub add_timeline_mrdd()''
'
'    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), "MRDd" _
'        , , xlTimeline).Slicers.Add ActiveSheet, , "MRDd", "MRDd", 171, 363, 262.5, 108
'    ActiveSheet.Shapes.Range(Array("MRDd")).Select
'    ActiveSheet.Shapes("MRDd").IncrementLeft 452.25
'    ActiveSheet.Shapes("MRDd").IncrementTop -104.25
'End Sub

