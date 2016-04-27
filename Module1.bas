Attribute VB_Name = "Module1"
Private Sub nowy_pivot_duzy()
Attribute nowy_pivot_duzy.VB_ProcData.VB_Invoke_Func = " \n14"
'
' nowy_pivot_duzy Macro
'

'
    Sheets("pivotSource").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1:V2789").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "pivotSource!R1C1:R2789C22", Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Sheet9!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15
    Sheets("Sheet9").Select
    Cells(3, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
End Sub
