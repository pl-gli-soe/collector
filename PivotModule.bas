Attribute VB_Name = "PivotModule"
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
Public Sub del_conf_pivot(ictrl As IRibbonControl)
    inner_del_conf_pivot
End Sub

Public Sub pn_pivot(ictrl As IRibbonControl)
    inner_pn_pivot
End Sub


Public Sub fup_pivot(ictrl As IRibbonControl)
    inner_fup_pivot
End Sub

Public Sub ppap_pivot(ictrl As IRibbonControl)
    inner_ppap_pivot
End Sub

Public Sub resp_pivot(ictrl As IRibbonControl)
    inner_resp_pivot
End Sub


Public Sub inner_resp_pivot()


    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWIZ.RESP_PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWIZ.RESP_PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWIZ.PIVOT_SOURCE_SHEET_NAME)
        ActiveWindow.Zoom = 80
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWIZ.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    
    
    If source_range.Count > 1 Then
        
        
        Dim ww(16) As String, wk(16) As String, wc(16) As String, wp(16) As String
        Dim slajsers(16) As String
        
        clear_arr_ slajsers, 16
        
        ww(1) = "RESP"
        wk(1) = "COORD"
        wc(1) = "PN"
        
        
        
        
        Dim np As NewPivotHandler
        Set np = New NewPivotHandler
        np.init source_range, XWIZ.RESP_PIVOT_SHEET_NAME
        np.config_pivot ww, wk, wc, wp
        'np.add_slicers slajsers
        
        np.add_totals
        Set np = Nothing
        
        
        add_slicers_for_resp
    End If
End Sub

Public Sub inner_ppap_pivot()


    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWIZ.PPAP_PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWIZ.PPAP_PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWIZ.PIVOT_SOURCE_SHEET_NAME)
        ActiveWindow.Zoom = 80
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWIZ.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    
    
    If source_range.Count > 1 Then
        
        
        Dim ww(16) As String, wk(16) As String, wc(16) As String, wp(16) As String
        Dim slajsers(16) As String
        
        clear_arr_ slajsers, 16
        
        ww(1) = "PROJ"
        ww(2) = "PPAP Status"
        wk(1) = "COORD"
        wc(1) = "PN"
        
        
        
        
        Dim np As NewPivotHandler
        Set np = New NewPivotHandler
        np.init source_range, XWIZ.PPAP_PIVOT_SHEET_NAME
        np.config_pivot ww, wk, wc, wp
        ' np.add_slicers slajsers
        np.add_totals
        Set np = Nothing
        
        
        add_slicers_for_ppap
    End If
End Sub

Public Sub inner_fup_pivot()
    
    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWIZ.FUP_PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWIZ.FUP_PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWIZ.PIVOT_SOURCE_SHEET_NAME)
        ActiveWindow.Zoom = 80
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWIZ.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    
    
    If source_range.Count > 1 Then
        
        
        Dim ww(16) As String, wk(16) As String, wc(16) As String, wp(16) As String
        Dim slajsers(16) As String
        
        clear_arr_ slajsers, 16
        
        ww(1) = "PLT"
        ww(2) = "PROJ"
        ww(3) = "FAZA"
        wk(1) = "FUP"
        wc(1) = "PN"
        
        
        
        Dim np As NewPivotHandler
        Set np = New NewPivotHandler
        np.init source_range, XWIZ.FUP_PIVOT_SHEET_NAME
        np.config_pivot ww, wk, wc, wp
        ' np.add_slicers slajsers
        np.add_totals
        Set np = Nothing
        
        
        add_slicers_for_fup
        add_timeline_for_fup
    End If
    
End Sub

Private Sub clear_arr_(arr() As String, ile)

    For x = 0 To ile
        On Error Resume Next
        arr(x) = ""
    Next x
End Sub


Public Sub inner_pn_pivot()
    
    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWIZ.PN_PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWIZ.PN_PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWIZ.PIVOT_SOURCE_SHEET_NAME)
        ActiveWindow.Zoom = 80
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWIZ.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    
    
    If source_range.Count > 1 Then
        
        
        Dim ww(16) As String, wk(16) As String, wc(16) As String, wp(16) As String
        Dim slajsers(16) As String
        
        ww(1) = "PN"
        wk(1) = "MRD"
        wc(1) = "PN"
        
        slajsers(1) = "PLT"
        slajsers(2) = "PROJ"
        slajsers(3) = "FAZA"
        slajsers(4) = "BG"
        
        Dim np As NewPivotHandler
        Set np = New NewPivotHandler
        np.init source_range, XWIZ.PN_PIVOT_SHEET_NAME
        np.config_pivot ww, wk, wc, wp
        np.add_slicers slajsers
        Set np = Nothing
        
        
        
    End If
    
End Sub


Private Sub inner_del_conf_pivot()

    Dim pivotsh As Worksheet, pivotsourcesh As Worksheet
    Dim source_range As Range
    Dim p As Range
    Dim k As Range
    
    
    
    
    
    With ThisWorkbook
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(XWIZ.DEL_CONF_PIVOT_SHEET_NAME).Delete
        Application.DisplayAlerts = True
        Set pivotsh = .Sheets.Add
        pivotsh.Name = XWIZ.DEL_CONF_PIVOT_SHEET_NAME
        Set pivotsourcesh = .Sheets(XWIZ.PIVOT_SOURCE_SHEET_NAME)
        ActiveWindow.Zoom = 80
    End With
    
    Set p = pivotsourcesh.Range("A1")
    Set k = pivotsourcesh.Range("A1")
    
    Do
        Set k = k.Offset(1, 0)
    Loop Until k = ""
    
    Set source_range = pivotsourcesh.Range(p, k.Offset(-1, XWIZ.OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE))
    
    
    If source_range.Count > 1 Then
    
        
        Dim ww(16) As String, wk(16) As String, wc(16) As String, wp(16) As String
        Dim slajsers(16) As String
        
        
        ww(1) = "DEL CONF"
        wk(1) = "MRD"
        wc(1) = "PN"
        
        'slajsers(1) = "COORD"
        'slajsers(2) = "FUP"
        'slajsers(3) = "PLT"
        'slajsers(4) = "PROJ"
        'slajsers(5) = "FAZA"
        'slajsers(6) = "PPAP Status"
        
        Dim np As NewPivotHandler
        Set np = New NewPivotHandler
        np.init source_range, XWIZ.DEL_CONF_PIVOT_SHEET_NAME
        np.config_pivot ww, wk, wc, wp
        ' np.add_slicers slajsers
        Set np = Nothing
        
        
        add_slicers_for_del_conf
        add_timeline_for_del_conf
    
    End If
End Sub

