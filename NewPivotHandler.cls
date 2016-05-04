VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewPivotHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



Private pivot_cache As PivotCache
Private pivot_table As PivotTable
Private source_range As Range
Private source_sh As Worksheet
Private pivot_sh As Worksheet

Private sc As SlicerCache


Public Sub init(sr As Range, pivot_sh_nm As String)


    Set source_range = sr
    
    
    ' Debug.Print source_range.Rows.Count
    
    ' example
    ' Create the Cache
    'Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
    '    SourceData:=Range("Dynamic_Field_Summary"))
    
    ' Set source_range = sr
    Set pivot_cache = Nothing
    On Error Resume Next
    Set pivot_cache = _
        ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:="pivotSource!R1C1:R" & CStr(source_range.Rows.Count) & "C" & CStr(OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    If pivot_cache Is Nothing Then
        MsgBox "cos poszlonie tak z proba utworzenia pivota"
        End
    Else
    
        Set source_sh = ThisWorkbook.Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME)
        
        Set pivot_sh = ThisWorkbook.Sheets(CStr(pivot_sh_nm))
    
        
        Set pivot_table = pivot_sh.PivotTables.Add(PivotCache:=pivot_cache, TableDestination:=Range("B4"))
    End If
End Sub

Public Sub config_pivot(ww() As String, _
    wk() As String, _
    wc() As String, _
    wp() As String)



        ' ww = w wierszach
        ' wk = w kolumnach
        ' wc = podliczamy county
        ' wp = filtry calej strony
        
        ' iteratory pod pozycje
        wwi = 1
        wki = 1
        wci = 1
        wpi = 1
        
        With pivot_table
        
            On Error Resume Next
            For w = LBound(ww) To UBound(ww)
            
            
                If Trim(CStr(ww(w))) <> "" Then
                    .PivotFields(CStr(ww(w))).Orientation = xlRowField
                    .PivotFields(CStr(ww(w))).Position = wwi
                    wwi = wwi + 1
                End If
            Next w
            
            On Error Resume Next
            For k = LBound(wk) To UBound(wk)
                
                If Trim(CStr(wk(k))) <> "" Then
                    .PivotFields(CStr(wk(k))).Orientation = xlColumnField
                    .PivotFields(CStr(wk(k))).Position = wki
                    
                    wki = wki + 1
                End If
            Next k
            
            On Error Resume Next
            For c = LBound(wc) To UBound(wc)
            
                If Trim(CStr(wk(c))) <> "" Then
                    .PivotFields(CStr(wk(c))).Orientation = xlDataField
                    .PivotFields(CStr(wk(c))).Function = xlCount
                    .PivotFields(CStr(wk(c))).Position = wci
                    
                    wci = wci + 1
                End If
            Next c
            
            On Error Resume Next
            For p = LBound(wp) To UBound(wp)
            
                If CStr(wp(p)) <> "" Then
                    .PivotFields(CStr(wp(p))).Orientation = xlDataField
                    .PivotFields(CStr(wp(p))).Function = xlCount
                    .PivotFields(CStr(wp(p))).Position = wpi
                    
                    wpi = wpi + 1
                End If
            Next p
        
            
            
            .ColumnGrand = False
            .RowGrand = False
            
            .TableStyle2 = "PivotStyleMedium15"
        End With
End Sub


Public Sub add_slicers(arr() As String)

    'Range("E12").Select
    'ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable11"), _
    '    "COORD").Slicers.Add ActiveSheet, , "COORD", "COORD", 145.5, 648.75, 144, _
    '    198.75
    
    'Set slajser_cache5 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "MRDd", , xlTimeline)
    'slajser_cache5.Slicers.Add pivot_sh, , "MRDd", "MRDd", 50, 350, 400, 100
    
    'ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), "PLT") _
    '    .Slicers.Add ActiveSheet, , "PLT 1", "PLT", 140.25, 508.5, 144, 198.75
    'ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), "PROJ" _
    '    ).Slicers.Add ActiveSheet, , "PROJ 1", "PROJ", 177.75, 546, 144, 198.75
    'ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), "FAZA" _
    '    ).Slicers.Add ActiveSheet, , "FAZA 1", "FAZA", 215.25, 583.5, 144, 198.75
    'ActiveSheet.Shapes.Range(Array("FAZA 1")).Select
    
    pivot_table.Parent.Activate
    
    
    
    lvl = 1
    For x = LBound(arr) To UBound(arr)
        If Trim(arr(x)) <> "" Then
        
        
        
            ThisWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(1), CStr(arr(x))).Slicers.Add _
               ActiveSheet, , _
               CStr(arr(x)) & "_1", _
               CStr(arr(x)), _
               lvl * 20 + 100, _
               lvl * 100 + 500, _
               150, _
               200

            
               
            lvl = lvl + 1
        End If
    Next x
    

    
    
End Sub

Public Sub add_totals()
    
    With pivot_table
        .ColumnGrand = True
        .RowGrand = True
    End With
End Sub

