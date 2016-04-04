VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicDelConfForm 
   Caption         =   "Dynamic Del Conf"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   OleObjectBlob   =   "DynamicDelConfForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DynamicDelConfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Sub BtnReset_Click()
    set_default_values_on_checkboxes
    recalc_config_sheet_dyn_del_conf_table
End Sub

Private Sub BtnSubmit_Click()
    ' hide
    recalc_config_sheet_dyn_del_conf_table
End Sub

Private Sub set_default_values_on_checkboxes()
    Me.CheckBoxBlank.Value = True
    Me.CheckBoxEDI.Value = False
    Me.CheckBoxHO.Value = False
    Me.CheckBoxPOTITDC.Value = True
    'Me.CheckBoxMRD.Value = True
    'Me.CheckBoxSMRD.Value = True
    'Me.CheckBoxTWOMRD.Value = True
    Me.CheckBoxNA.Value = False
    Me.CheckBoxOS.Value = False
    'Me.CheckBoxALTMRD.Value = True
    Me.CheckBoxUNDEF.Value = True
    ' Me.CheckBoxMRDS.Value = False
    
    ' mrd stuff
    ' ====================================
    With Me
        .ComboBoxALTMRD.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_NOK
        .ComboBoxMRD.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
        .ComboBoxMRDStaggered.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
        .ComboBoxMRDTWO.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
        .ComboBoxTWOStaggeredMRD.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT
    End With
    
    ' ====================================
End Sub

Private Sub recalc_config_sheet_dyn_del_conf_table()
    
    Dim cfg As Worksheet
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    
    Dim r As Range
    
    
    With Me
        ' m9 to blank
        Set r = cfg.Range("N9")
        If .CheckBoxBlank.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        Set r = cfg.Range("N10")
        If .CheckBoxPOTITDC.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        
        ' MRDs
        'If .CheckBoxMRD.Value = True Then
        '    cfg.Range("N11") = 3
        '    cfg.Range("N12") = 3
        '    ' ...
        '    cfg.Range("N15") = 3
        'Else
        '    cfg.Range("N11") = 2
        '    cfg.Range("N12") = 2
        '    ' ...
        '    cfg.Range("N15") = 2
        'End If
        
        
        ' MRD
        'If .CheckBoxMRDS.Value = True Then
        '    cfg.Range("N11") = 1
        'ElseIf .CheckBoxMRD.Value = True Then
        '    cfg.Range("N11") = 3
        'Else
        '    cfg.Range("N11") = 2
        'End If
        
        ' stagg MRD
        'If .CheckBoxMRDS.Value = True Then
        '    cfg.Range("N12") = 1
        'ElseIf .CheckBoxSMRD.Value = True Then
        '    cfg.Range("N12") = 3
        'Else
        '    cfg.Range("N12") = 2
        'End If
        
        'If .CheckBoxMRDS.Value = True Then
        '    cfg.Range("N15") = 1
        'ElseIf .CheckBoxTWOMRD.Value = True Then
        '    cfg.Range("N15") = 3
        'Else
        '    cfg.Range("N15") = 2
        'End If
        
        
        ' ho n13
        Set r = cfg.Range("N13")
        If .CheckBoxHO.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        
        ' EDI
        Set r = cfg.Range("N14")
        If .CheckBoxEDI.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        ' ON STOCK
        Set r = cfg.Range("N16")
        If .CheckBoxOS.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        ' NA
        Set r = cfg.Range("N17")
        If .CheckBoxNA.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        
        ' alt mrd
        'Set r = cfg.Range("N18")
        'If .CheckBoxALTMRD.Value = True Then
        '    r = 1
        'Else
        '    r = 2
        'End If
        
        ' undef
        Set r = cfg.Range("N19")
        If .CheckBoxUNDEF.Value = True Then
            r = 1
        Else
            r = 2
        End If
        
        
        ' mrd stuff
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        
        
        ' mrd
        Set r = cfg.Range("N11")
        r = foo_get_value_from_combobox(.ComboBoxMRD)
        
        ' mrd stag
        Set r = cfg.Range("N12")
        r = foo_get_value_from_combobox(.ComboBoxMRDStaggered)
        
        ' mrd two
        Set r = cfg.Range("N15")
        r = foo_get_value_from_combobox(.ComboBoxMRDTWO)
        
        ' alt mrd
        Set r = cfg.Range("N18")
        r = foo_get_value_from_combobox(.ComboBoxALTMRD)
        
        ' alt mrd
        Set r = cfg.Range("N20")
        r = foo_get_value_from_combobox(.ComboBoxTWOStaggeredMRD)
        
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        
        
    End With
End Sub

Private Function foo_get_value_from_combobox(cbbx As Variant) As E_DYNAMIC_CFG_FOR_DEL_CONF


    If cbbx.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_NOK Then
    
        foo_get_value_from_combobox = XWiz.E_DYNAMIC_CFG_FOR_DEL_CONF_NOK
        
    ElseIf cbbx.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_OK Then
    
        foo_get_value_from_combobox = XWiz.E_DYNAMIC_CFG_FOR_DEL_CONF_OK
        
    ElseIf cbbx.Value = XWiz.COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT Then
    
        foo_get_value_from_combobox = XWiz.E_DYNAMIC_CFG_FOR_DEL_CONF_CALC_WITH_MRD
    Else
    
        MsgBox "to nie moze sie wydarzyc! - foo_get_value_from_combobox"
        End
    End If
End Function
