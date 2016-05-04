Attribute VB_Name = "GoToRepModule"
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

Public Sub go_to_source_pivot_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME).Activate
End Sub

Public Sub go_to_rep(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME).Activate
End Sub


Public Sub go_to_rep_fup(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.REP_FUP_SHEET_NAME).Activate
End Sub

Public Sub go_to_rep_all(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.ALL_SHEET_NAME).Activate
End Sub

Public Sub go_to_config(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Activate
End Sub

Public Sub go_to_through_selection(ictrl As IRibbonControl)
    inner_go_to_through_selection ActiveCell
End Sub


' pivots
' ===============================================
Public Sub go_to_del_conf_pivot(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.DEL_CONF_PIVOT_SHEET_NAME).Activate
End Sub

Public Sub go_to_rep_pn_pivot(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.PN_PIVOT_SHEET_NAME).Activate
End Sub



Public Sub go_to_rep_ppap_pivot(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.PPAP_PIVOT_SHEET_NAME).Activate
End Sub

Public Sub go_to_rep_fup_pivot(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.FUP_PIVOT_SHEET_NAME).Activate
End Sub

Public Sub go_to_rep_resp_pivot(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(XWiz.RESP_PIVOT_SHEET_NAME).Activate
End Sub
' ===============================================


Public Sub inner_go_to_through_selection(target As Range)

    Dim r As Range
    Set r = target.Offset(0, -1)

    's = "." & remove_special_cases(CStr(r)) & _
    '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
    '    "*" & remove_special_cases(CStr(r.Offset(0, 2))) & _
    '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"

    Set r = r.Offset(0, XWiz.E_ACTIVE - 1)
    s = ""
    On Error Resume Next
    s = Trim(CStr(r.Comment.Text))
    
    
    If s <> "" Then
    
    
        For Each Sh In ThisWorkbook.Sheets
        
            shname = Sh.Name
            If UCase(CStr(Sh.Range("C1"))) = UCase(CStr(s)) Then
                ' Debug.Print sh.Name & " " & sh.Range("B1")
                Sh.Activate
                Exit For
            End If
        Next Sh
    Else
        MsgBox "Brak Unique ID!"
    End If
End Sub
