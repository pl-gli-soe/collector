Attribute VB_Name = "PivotModule"
Public Sub del_conf_pivot(ictrl As IRibbonControl)
    new_pivot
End Sub

Public Sub new_pivot()
    
    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet, ph As PivotHandler
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWiz.PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWiz.PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME)
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWiz.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    If source_range.Count > 1 Then
    
    
        Set ph = New PivotHandler
        
        With ph
            ' -------------------------------------------------
            .init source_range
            .config_pivot
            .add_slicers
            ' -------------------------------------------------
        End With
        
        Set ph = Nothing
    End If
    
End Sub

' wykorzystanie z prio
'Public Sub create_PRIO_pivot(e As pivot_layout)
'    Dim ph As PivotHandler
'    Set ph = New PivotHandler
'
'    Application.EnableEvents = False
'
'    ph.if_flat_table_prepare_source_range
'    ph.init e
'    ph.config_pivot
'    ph.add_slicers e
'
'    Application.EnableEvents = True
'End Sub
