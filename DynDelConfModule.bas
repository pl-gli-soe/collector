Attribute VB_Name = "DynDelConfModule"
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


Public Sub dyn_del_conf(ictrl As IRibbonControl)
    start_dyn_del_conf
End Sub

Private Sub start_dyn_del_conf()

    Dim cfg As Worksheet, r As Range, combobox As Variant, chbox As Variant
    Set cfg = ThisWorkbook.Sheets(XWIZ.CONFIG_SHEET_NAME)
    
    With DynamicDelConfForm
    
        ' BLANK
        Set r = cfg.Range("N9")
        Set chbox = .CheckBoxBlank
        ' std ok nok
        std_ok_nok r, chbox
        
        ' ITDC
        Set r = cfg.Range("N10")
        Set chbox = .CheckBoxPOTITDC
        ' std ok nok
        std_ok_nok r, chbox
        
        ' HO
        Set r = cfg.Range("N13")
        Set chbox = .CheckBoxHO
        ' std ok nok
        std_ok_nok r, chbox
        
        
        ' EDI
        Set r = cfg.Range("N14")
        Set chbox = .CheckBoxEDI
        ' std ok nok
        std_ok_nok r, chbox
        
        ' ON STOCK
        Set r = cfg.Range("N16")
        Set chbox = .CheckBoxOS
        ' std ok nok
        std_ok_nok r, chbox
        
        ' NA
        Set r = cfg.Range("N17")
        Set chbox = .CheckBoxNA
        ' std ok nok
        std_ok_nok r, chbox
        
        
        ' undfd
        Set r = cfg.Range("N19")
        Set chbox = .CheckBoxUNDEF
        ' std ok nok
        std_ok_nok r, chbox
        
        
        
        
        ' mrd stuff
        ' ====================================
        
        
        ' MRD
        Set r = cfg.Range("N11")
        fill_combo_box_and_set_value r, .ComboBoxMRD
        
        Set r = cfg.Range("N12")
        fill_combo_box_and_set_value r, .ComboBoxMRDStaggered
        
        
        ' obsolete
        Set r = cfg.Range("N15")
        fill_combo_box_and_set_value r, .ComboBoxMRDTWO
        
        Set r = cfg.Range("N18")
        fill_combo_box_and_set_value r, .ComboBoxALTMRD
        
        
        ' obsolete
        Set r = cfg.Range("N20")
        fill_combo_box_and_set_value r, .ComboBoxTWOStaggeredMRD
        
        ' ALT TWO
        Set r = cfg.Range("N21")
        fill_combo_box_and_set_value r, .ComboBoxMRDALTTWO
        
        ' Staggered ALT TWO
        Set r = cfg.Range("N22")
        fill_combo_box_and_set_value r, .ComboBoxMRDStaggeredALTTWO
        
        ' ONCOST
        Set r = cfg.Range("N23")
        fill_combo_box_and_set_value r, .ComboBoxMRDONCOST
        
        ' Staggered ONCOST
        Set r = cfg.Range("N24")
        fill_combo_box_and_set_value r, .ComboBoxMRDStaggeredONCOST
        
        
        
        ' ====================================
        
        
        
        
        .show
        
    End With
End Sub


Private Sub fill_combo_box_and_set_value(ByRef r As Range, cbbx As Variant)
    
    cbbx.Clear
    cbbx.AddItem XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_OK
    cbbx.AddItem XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_NOK
    cbbx.AddItem XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
    
    txt = ""
    If r = XWIZ.E_DYNAMIC_CFG_FOR_DEL_CONF_NOK Then
        txt = XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_NOK
    ElseIf r = XWIZ.E_DYNAMIC_CFG_FOR_DEL_CONF_OK Then
        txt = XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_OK
    ElseIf r = XWIZ.E_DYNAMIC_CFG_FOR_DEL_CONF_CALC_WITH_MRD Then
        txt = XWIZ.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
    Else
        txt = "sth went wrong in fill_combo_box_and_set_value"
        MsgBox "sth went wrong in fill_combo_box_and_set_value"
        End
    End If
    
    cbbx.Value = txt
End Sub


Private Sub std_ok_nok(ByRef r As Range, ByRef chbox As Variant)
    
    If r = 3 Then
        chbox.Value = True
    ElseIf r = 2 Then
        chbox.Value = False
    Else
        chbox.Value = True
        r = 1
    End If
End Sub
