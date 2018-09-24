Attribute VB_Name = "SlicersModule"
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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

Public Sub add_slicers_for_resp()


    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.RESP_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then

        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PLT" _
            ).Slicers.Add ActiveSheet, , "PLT", "PLT", 126.75, 508.5, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "PROJ").Slicers.Add ActiveSheet, , "PROJ", "PROJ", 164.25, 546, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "FAZA").Slicers.Add ActiveSheet, , "FAZA", "FAZA", 201.75, 583.5, 144, 198.75
        ActiveSheet.Shapes.Range(Array("FAZA")).Select
    End If
End Sub

Public Sub add_slicers_for_ppap()

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.PPAP_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then


        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PLT") _
            .Slicers.Add ActiveSheet, , "PLT 2", "PLT", 89.25, 471, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PROJ" _
            ).Slicers.Add ActiveSheet, , "PROJ 2", "PROJ", 126.75, 508.5, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "FAZA" _
            ).Slicers.Add ActiveSheet, , "FAZA 2", "FAZA", 164.25, 546, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "MRD") _
            .Slicers.Add ActiveSheet, , "MRD", "MRD", 201.75, 583.5, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "COORD").Slicers.Add ActiveSheet, , "COORD", "COORD", 239.25, 621, 144, 198.75
        ActiveSheet.Shapes.Range(Array("COORD")).Select

        
    End If
End Sub

Sub add_slicers_for_del_conf()


    'slajsers(1) = "COORD"
    'slajsers(2) = "FUP"
    'slajsers(3) = "PLT"
    'slajsers(4) = "PROJ"
    'slajsers(5) = "FAZA"
    'slajsers(6) = "PPAP Status"

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.DEL_CONF_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then

        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PLT") _
            .Slicers.Add ActiveSheet, , "PLT 11", "PLT", 89.25, 471, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PROJ" _
            ).Slicers.Add ActiveSheet, , "PROJ 11", "PROJ", 126.75, 508.5, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "FAZA" _
            ).Slicers.Add ActiveSheet, , "FAZA 13", "FAZA", 164.25, 546, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "COORD").Slicers.Add ActiveSheet, , "COORD 10", "COORD", 201.75, 583.5, 144, _
            198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "Fst Pickup Date").Slicers.Add ActiveSheet, , "Fst Pickup Date", _
            "Fst Pickup Date", 239.25, 621, 144, 198.75
        ActiveSheet.Shapes.Range(Array("Fst Pickup Date")).Select
    End If
End Sub

Sub add_slicers_for_fup()
Attribute add_slicers_for_fup.VB_ProcData.VB_Invoke_Func = " \n14"

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.FUP_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then

        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PLT") _
            .Slicers.Add ActiveSheet, , "PLT 1", "PLT", 108, 489.75, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "PROJ" _
            ).Slicers.Add ActiveSheet, , "PROJ 1", "PROJ", 145.5, 527.25, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "FAZA" _
            ).Slicers.Add ActiveSheet, , "FAZA 1", "FAZA", 183, 564.75, 144, 198.75
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), "MRD") _
            .Slicers.Add ActiveSheet, , "MRD 1", "MRD", 220.5, 602.25, 144, 198.75
        ActiveSheet.Shapes.Range(Array("MRD 1")).Select
    End If
End Sub


Sub add_timeline_for_fup()
Attribute add_timeline_for_fup.VB_ProcData.VB_Invoke_Func = " \n14"

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.FUP_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then
    
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "MRDd", , xlTimeline).Slicers.Add ActiveSheet, , "MRDd", "MRDd", 10, 600 _
            , 300, 100
        ActiveSheet.Shapes.Range(Array("MRDd")).Select
    End If
End Sub


Sub add_timeline_for_del_conf()
Attribute add_timeline_for_del_conf.VB_ProcData.VB_Invoke_Func = " \n14"
    ThisWorkbook.Activate
    ThisWorkbook.Sheets(XWIZ.DEL_CONF_PIVOT_SHEET_NAME).Activate

    Dim ptb As PivotTable, p As PivotTable
    Set p = Nothing
    For Each ptb In ActiveSheet.PivotTables
        
        On Error Resume Next
        Set p = ptb
        Exit For
    Next ptb

    If Not p Is Nothing Then
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(CStr(p.Name)), _
            "MRDd", , xlTimeline).Slicers.Add ActiveSheet, , "MRDd 1", "MRDd", 10, _
            800, 300, 108
    
    End If
End Sub
