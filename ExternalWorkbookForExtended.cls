VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalWorkbookForExtended"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

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

Implements IDataFromWizard



' wczesniej zapisac
Private wrkbk As Workbook
Private source_master As Worksheet
Private source_details As Worksheet
Private source_pickups As Worksheet
Private cfg As Worksheet

Private all_sh As Worksheet

Private beg_of_del_conf_oknok_dyn As Range

Private fn As String ' filename
Private sp As String ' sampath


Private e_fup As E_FUP_FILTER
Private e_run_type As E_RUN_REP_TYPE


Private Sub Class_Initialize()

    fn = ""
    sp = ""
End Sub

Private Sub IDataFromWizard_init(Optional mwrkbk As Workbook, Optional e As E_FUP_FILTER, Optional e2 As E_RUN_REP_TYPE)


    Debug.Print mwrkbk.FullName

    Set wrkbk = mwrkbk
    Set source_master = wrkbk.Sheets(XWIZ.MASTER_SHEET_NAME)
    Set source_details = wrkbk.Sheets(XWIZ.DETAILS_SHEET_NAME)
    Set source_pickups = wrkbk.Sheets(XWIZ.PICKUPS_SHEET_NAME)
    
    e_fup = e
    e_run_type = e2
    
    
End Sub


Private Sub IDataFromWizard_runMainLogic(r As Range, sampath As Variant, arr As Variant, sh As StatusHandler, Optional c As Collection)



    ' Debug.Print arr(0)
    
    Dim rsh As Worksheet
    Set rsh = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    Dim repR As Range
    Set repR = rsh.Range("A1")
    
    Dim detR As Range
    Debug.Print source_details.Name
    
    Set detR = source_details.Cells(1, 2)
    
    For x = 0 To 20
        
        repR.Offset(0, x).Value = detR.Offset(x, 0).Value
    Next x
    
    
    
    lr = source_master.Range("a1").End(xlDown).Row
    lc = source_master.Range("a1").End(xlToRight).Column
    
    rsh.Cells(2, 1).Value = lr
    rsh.Cells(2, 2).Value = lc
    rsh.Cells(2, 3).Value = XWIZ.G_SIDE_SIGNATURE
    
    source_master.Range(source_master.Cells(1, 1), source_master.Cells(lr, lc)).Copy
    rsh.Cells(3, 1).PasteSpecial xlPasteValues
    rsh.Activate
    rsh.Cells(3, 1).Activate
    
    
    
End Sub






























' not important subs for this class

Private Sub IDataFromWizard_mrd1ConfStatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_mrd1StatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, wypelnij_wstepnie_kolekcje_fma_resp As Boolean, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_mrd2ConfStatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_mrd2StatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_podsumowanieSlownikaNaNwymArkuszu()
    ' nop
End Sub

Private Sub IDataFromWizard_subMrd1PusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_subMrd2PusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub

Private Sub IDataFromWizard_subTotalPusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, sh As StatusHandler)
    ' nop
End Sub




















Private Sub IDataFromWizard_addComment(fcx As Variant, nok_counter As Variant, x As Variant, sampath As Variant, filename As Variant, jaki_status As E_COLUMNS_WITH_FORMULAS)
    ' nop
End Sub

Private Sub IDataFromWizard_deliveryConfirmationLogic(count_pns As Variant, sampath As Variant, r As Variant, td As Workbook, sh As StatusHandler)
    ' nop
End Sub

Private Function IDataFromWizard_fooCountPns(szukaj_pustego As Variant, filename As Variant, sampath As Variant) As Variant
    ' nop
End Function
