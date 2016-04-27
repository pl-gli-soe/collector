VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalReadOnlyWorkbook"
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



Implements IDataFromWizard


Private kolekcja_fma_resp_zeby_tylko_raz_przegladac As Collection

' tandem
Private rhaa As RecordHandlerAllApproach
Private kolekcja_obiektow_typu_RecordHandlerAllApproach As Collection

Private for_comment(1 To 1024) As String


' to na potrzeby comentarza
Private d As Dictionary
Private rh As RecordsHandler


' wczesniej zapisac
Private wrkbk As Workbook
Private source_master As Worksheet
Private source_detilas As Worksheet
Private source_pickups As Worksheet
Private cfg As Worksheet

Private all_sh As Worksheet

Private beg_of_del_conf_oknok_dyn As Range

Private fn As String ' filename
Private sp As String ' sampath


Private e_fup As E_FUP_FILTER
Private e_run_type As E_RUN_REP_TYPE


Private Sub Class_Initialize()
    Set d = New Dictionary
End Sub

Private Sub IDataFromWizard_addComment(fcx As Variant, nok_counter As Variant, x As Variant, _
    sampath As Variant, filename As Variant, jaki_status As E_COLUMNS_WITH_FORMULAS)



    Dim pn As String, pn_nm As String, duns As String
    Dim supp_nm As String, resp As String, fma_fup As String
    
    Dim del_conf As String
    Dim order_date As String
    Dim cmnts As String


    ' hash & chr(10) & row
    fcx = fcx & _
        XWiz.XWIZ_TXT_CMNT_HASH & dopelnij_spacjami(CStr(nok_counter), 5) & _
        " | " & XWiz.XWIZ_TXT_CMNT_ROW & dopelnij_spacjami(CStr(x), 5) & " | "


    If Not d.Exists(x) Then
        Set rh = Nothing
        Set rh = New RecordsHandler
            
        ' PN
        'strRng = Cells(x, XWiz.pn).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'pn = CStr(ExecuteExcel4Macro(strRef))
        pn = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.pn)
        pn = dopelnij_spacjami(pn, XWiz.G_PN_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_PREFIX & pn & " | "
        
        
        
        ' PN name
        'strRng = Cells(x, XWiz.PN_Name).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'pn_nm = CStr(ExecuteExcel4Macro(strRef))
        pn_nm = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.PN_Name)
        pn_nm = dopelnij_spacjami(pn_nm, XWiz.G_PN_NM_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_NM_PREFIX & pn_nm & " | "
        
        
        ' DUNS
        'strRng = Cells(x, XWiz.duns).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'duns = CStr(ExecuteExcel4Macro(strRef))
        duns = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.duns))
        duns = dopelnij_spacjami(Trim(duns), 12)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DUNS_PREFIX & duns & " | "
        
        ' supplier name
        'strRng = Cells(x, XWiz.Supplier_Name).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'supp_nm = CStr(ExecuteExcel4Macro(strRef))
        supp_nm = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Supplier_Name)
        supp_nm = dopelnij_spacjami(supp_nm, XWiz.G_SUPP_NM_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_SUPP_NM_PREFIX & supp_nm & " | "
        
        ' RESP
        'strRng = Cells(x, XWiz.Responsibility).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'resp = CStr(ExecuteExcel4Macro(strRef))
        resp = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility)
        resp = dopelnij_spacjami(resp, XWiz.G_RESP_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_RESP_PREFIX & resp & " | "
        
        ' FMA FUP
        'strRng = Cells(x, XWiz.fup_code).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'fma_fup = CStr(ExecuteExcel4Macro(strRef))
        fma_fup = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code)
        fma_fup = dopelnij_spacjami(fma_fup, XWiz.G_FUPCODE_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_FUP_PREFIX & fma_fup & " | "
        
        
        del_conf = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Delivery_confirmation)
        del_conf = dopelnij_spacjami(del_conf, XWiz.G_DEL_CONF_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DEL_CONF_PREFIX & del_conf & " | "
        
        order_date = CStr(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_Ordered_Date))
        order_date = dopelnij_spacjami(order_date, XWiz.G_DATES_CW_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX & CStr(order_date) & " | "
        
        
        cmnts = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Comments)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_CMNTS_PREFIX & cmnts & " | "
        
        
        rh.pn = pn
        rh.pn_nm = pn_nm
        rh.duns = duns
        rh.supp_nm = supp_nm
        rh.resp = resp
        rh.fup = fma_fup
        rh.nok_counter = CStr(nok_counter)
        rh.Delivery_confirmation = del_conf
        rh.MRD1_Ordered_Date = order_date
        rh.Comments = cmnts
        
        
        If jaki_status = E_F_MRD1_ORDERED_STATUS Then
            rh.MRD1_Ordered_STATUS = False
        ElseIf jaki_status = E_F_MRD1_CONFIRMED_STATUS Then
            rh.MRD1_confirmed_qty_Status = False
        ElseIf jaki_status = E_F_MRD1_PUS_STATUS Then
            rh.MRD1_PUS_STATUS = False
        ElseIf jaki_status = E_F_MRD2_ORDERED_STATUS Then
            rh.MRD2_Ordered_STATUS = False
        ElseIf jaki_status = E_F_MRD2_CONFIRMED_STATUS Then
            rh.MRD2_confirmed_qty_Status = False
        ElseIf jaki_status = E_F_MRD2_PUS_STATUS Then
            rh.MRD2_PUS_STATUS = False
        ElseIf jaki_status = E_F_TOTAL_PUS_STATUS Then
            rh.TOTAL_PUS_STATUS = False
        ElseIf jaki_status = E_F_Delivery_confirmation Then
            rh.Delivery_confirmation_status = False
        End If
    
        d.Add x, rh
        
    
        
    Else
    
        Set rh = Nothing
        Set rh = d(x)
        
        With rh
            If jaki_status = E_F_MRD1_ORDERED_STATUS Then
                .MRD1_Ordered_STATUS = False
            ElseIf jaki_status = E_F_MRD1_CONFIRMED_STATUS Then
                .MRD1_confirmed_qty_Status = False
            ElseIf jaki_status = E_F_MRD1_PUS_STATUS Then
                .MRD1_PUS_STATUS = False
            ElseIf jaki_status = E_F_MRD2_ORDERED_STATUS Then
                .MRD2_Ordered_STATUS = False
            ElseIf jaki_status = E_F_MRD2_CONFIRMED_STATUS Then
                .MRD2_confirmed_qty_Status = False
            ElseIf jaki_status = E_F_MRD2_PUS_STATUS Then
                .MRD2_PUS_STATUS = False
            ElseIf jaki_status = E_F_TOTAL_PUS_STATUS Then
                .TOTAL_PUS_STATUS = False
            ElseIf jaki_status = E_F_Delivery_confirmation Then
                .Delivery_confirmation_status = False
            End If
        End With
        
        rh_pn = dopelnij_spacjami(rh.pn, XWiz.G_PN_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_PREFIX & rh_pn & " | "
        
        rh_pn_nm = dopelnij_spacjami(rh.pn_nm, XWiz.G_PN_NM_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_NM_PREFIX & rh_pn_nm & " | "
        
        rh_duns = dopelnij_spacjami(Trim(rh.duns), XWiz.G_DUNS_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DUNS_PREFIX & rh_duns & " | "
        
        
        rh_supp_nm = dopelnij_spacjami(rh.supp_nm, XWiz.G_SUPP_NM_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_SUPP_NM_PREFIX & rh_supp_nm & " | "
        
        rh_resp = dopelnij_spacjami(rh.resp, XWiz.G_RESP_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_RESP_PREFIX & rh_resp & " | "
        
        rh_fup = dopelnij_spacjami(rh.fup, XWiz.G_FUPCODE_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_FUP_PREFIX & rh_fup & " | "
        
        rh_del_conf = dopelnij_spacjami(rh.Delivery_confirmation, XWiz.G_DEL_CONF_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DEL_CONF_PREFIX & rh_del_conf & " | "
        
        rh_MRD1_Ordered_Date = dopelnij_spacjami(rh.MRD1_Ordered_Date, XWiz.G_DATES_CW_LEN)
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX & rh_MRD1_Ordered_Date & " | "
        
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_CMNTS_PREFIX & rh.Comments & " | "
    End If
    
    fcx = fcx & XWiz.XWIZ_TXT_CMNT_LINIA & Chr(10)
    
    
    ' Debug.Print fcx
End Sub


Private Sub IDataFromWizard_deliveryConfirmationLogic(count_pns As Variant, sampath As Variant, r As Variant, td As Workbook, ByRef Sh As StatusHandler)
    
    Dim count_noks_in_del_conf_status_with_fma_star_resp As Range
    Set count_noks_in_del_conf_status_with_fma_star_resp = r.Offset(0, 9)
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD1_confirmed_qty_dot__Status).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Delivery_confirmation)
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
        
        
            
            If special_treatment_for_delivery_confirmation(Result, td) Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20 = 0) Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_Delivery_confirmation
            End If
        End If
        
        ' sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count))

    Next x
    
    With count_noks_in_del_conf_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            Else
                Exit For
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
End Sub

Private Function special_treatment_for_delivery_confirmation(r, td As Workbook) As Boolean

    ' true is NOK - default way
    special_treatment_for_delivery_confirmation = True
    Dim mrd_range As Range
    ' tutaj pobieramy date MRD prosto z Wizarda z arkusza DETAILS
    Set mrd_range = td.Sheets(XWiz.DETAILS_SHEET_NAME).Range("mrd")
    
    
    ' Set mrd_range = source_details.Range("mrd")
    ' MsgBox (source_details Is Nothing)
    
    
    Dim mrd_jako_liczba_porzadkowa As Long
    mrd_bez_txtu = ""
    On Error Resume Next
    mrd_bez_txtu = Replace(Replace(CStr(mrd_range.Value), "Y", ""), "CW", "")
    
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ' cos jest nie tak z mrd
    If Len(mrd_bez_txtu) <> 6 Then
    
        If Len(mrd_bez_txtu) = 5 Then
            ' cw jest mniejsze od 10 zatem moze brakowac zera
            ' quick hack!
            mrd_bez_txtu = Left(mrd_bez_txtu, 4) & "0" & Right(mrd_bez_txtu, 1)
        Else
            MsgBox "special_treatment_for_delivery_confirmation - MRD w ogole bez std!"
            End
        End If
    End If
    
    If mrd_bez_txtu <> "" Then
    
        mrd_jako_liczba_porzadkowa = CLng(mrd_bez_txtu)
    Else
        MsgBox "cos poszlo nie tak z proba podmiany wartosci MRD w subie special_treatment_for_delivery_confirmation"
        special_treatment_for_delivery_confirmation = True
        Exit Function
    End If
    

    Dim tmp As String
    
    
    ' od razu przypisujemy wartosc
    tmp = r
    
    
    ' przeszkoda - usunieta
    'If Trim(r) = "" Then
    '    special_treatment_for_delivery_confirmation = True
    '    Exit Function
    'End If
    
    
    ' On Error Resume Next
    ' Debug.Print CStr(r)
    If Len(CStr(r)) > 0 Then
        
        ' zatem to jest juz zbedne jesli wartosc zostala przypisana
        ' tmp = r
        
        
        
        ' tutaj temat dynamicznosci OK / NOK'ow
        ' teraz dopiero widac jak bardzo statyczny jest ten kod
        ' ale zgodnie z zalozeniami nie musialoby mi to przeszkadzac
        ' w koncu jesli bym wszystko z gory tak przewidywal kodu bym nigdy nie skonczyl1
        ' please use E_DYNAMIC_CFG_FOR_DEL_CONF
        If CStr(dynamic_delivery_confirmation_get_content_look_for_oks()) Like "*" & tmp & "*" Then
        
                
                special_treatment_for_delivery_confirmation = False
                Exit Function
                
        
        ' powizane z mrd -> normal, staggered, two
        ElseIf tmp Like "*Y*CW*" Then
        
        
            ' Debug.Print tmp
        
        
            If check_if_mrds_calc_is_taken_into_consideration(tmp) = 3 Then
        
        
                tmp = Trim(Replace(Replace(Replace(Replace(Replace(tmp, "Staggered ", ""), "TWO/", ""), "CW", ""), "Y", ""), "ALT ", ""))
                Dim licza_z_tmp As Long
                liczba_z_tmp = CLng(tmp)
                
                If liczba_z_tmp <= mrd_jako_liczba_porzadkowa Then
                    special_treatment_for_delivery_confirmation = False
                    Exit Function
                Else
                    Exit Function
                End If
            ElseIf check_if_mrds_calc_is_taken_into_consideration(tmp) = 2 Then
                special_treatment_for_delivery_confirmation = False
                Exit Function
            ElseIf check_if_mrds_calc_is_taken_into_consideration(tmp) = 1 Then
                special_treatment_for_delivery_confirmation = True
                Exit Function
            End If
            
            
        ' cos co napewno nie jest zwiazane z MRD
        ' ani z niczym innym
        Else
            If ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("N19") = 1 Then
                special_treatment_for_delivery_confirmation = True
            Else
            
                If check_if_its_really_undefined(tmp) Then
                    special_treatment_for_delivery_confirmation = False
                End If
            End If
            Exit Function
        End If
        
    ElseIf tmp = "" And blank_checked_on_form() Then ' to je chyba pod blank w takim razie :) - zatem potrzebne dodatkowa logika
        special_treatment_for_delivery_confirmation = True
    ElseIf Not blank_checked_on_form() Then
        special_treatment_for_delivery_confirmation = False
    End If
End Function


Private Function blank_checked_on_form() As Boolean
    blank_checked_on_form = False
    
    
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    
    If cfg.Range("N9").Value = XWiz.E_DYNAMIC_CFG_FOR_DEL_CONF_NOK Then
        blank_checked_on_form = True
    Else
        blank_checked_on_form = False
    End If
End Function


Private Function check_if_its_really_undefined(tmp) As Boolean
    check_if_its_really_undefined = True
    
    
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    ' lecimy pomijajac blanks - on ma osobna logike
    Set beg_of_del_conf_oknok_dyn = cfg.Range("M9")
    
    Dim b As Range
    Set b = beg_of_del_conf_oknok_dyn
    
    
    Do
        If CStr(b) = CStr(tmp) Then
            check_if_its_really_undefined = False
            Exit Function
        End If
        Set b = b.Offset(1, 0)
    Loop Until Trim(b) = ""
    
    check_if_its_really_undefined = True
    
End Function

Private Function dynamic_delivery_confirmation_get_content_look_for_oks() As String
    dynamic_delivery_confirmation_get_content_look_for_oks = ""
    
    ' Private cfg As Worksheet

    ' Private beg_of_del_conf_oknok_dyn As Range
    
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    ' lecimy pomijajac blanks - on ma osobna logike
    Set beg_of_del_conf_oknok_dyn = cfg.Range("M9")
    
    Dim b As Range
    Set b = beg_of_del_conf_oknok_dyn
    
    
    Do
        If CStr(b.Offset(0, 1)) = "2" Then
            dynamic_delivery_confirmation_get_content_look_for_oks = _
                dynamic_delivery_confirmation_get_content_look_for_oks & ";" & b
        End If
        Set b = b.Offset(1, 0)
    Loop Until Trim(b) = ""
    
End Function

Private Function check_if_mrds_calc_is_taken_into_consideration(tmp As String) As Integer
    check_if_mrds_calc_is_taken_into_consideration = -1
    
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    
    
    'Dim sum As Integer
    'sum = Int(cfg.Range("n11") + cfg.Range("n12") + cfg.Range("n15"))
    '
    ' w prosty sposob sumujemy wartosci i sprawdzamy czy mamy 3 x 3 = 9 :)
    'If sum = 9 Then
    '    check_if_mrds_calc_is_taken_into_consideration = True
    'Else
    '    check_if_mrds_calc_is_taken_into_consideration = False
    'End If
    
    ' powyzsza logika miala sens wtedy gdy bralismy wszystko na jednego bata
    ' teraz musismy to podzielic
    
    
    ' tmp = Trim(Replace(Replace(Replace(Replace(tmp, "Staggered ", ""), "TWO/", ""), "CW", ""), "Y", ""))
    
    'life hack - spradzam od razu two i staggered co by wczesniej nie zlapalo
    If tmp Like "*TWO*Staggered*" Then
    
        check_if_mrds_calc_is_taken_into_consideration = Int(cfg.Range("n20"))
    
    ' staggered
    ElseIf tmp Like "*Staggered*" Then
    
    
        check_if_mrds_calc_is_taken_into_consideration = Int(cfg.Range("n12"))
    
    'two
    ElseIf tmp Like "*TWO*" Then
    
        check_if_mrds_calc_is_taken_into_consideration = Int(cfg.Range("n15"))
        
        
    'alt
    ElseIf tmp Like "*ALT*" Then
    
        check_if_mrds_calc_is_taken_into_consideration = Int(cfg.Range("n18"))
    
    ' czyste mrd
    Else
    
        check_if_mrds_calc_is_taken_into_consideration = Int(cfg.Range("n11"))
    End If
End Function


Private Function sprawdz_true_czy_false_na_podstawie(ir As Range) As Boolean
    
    If Int(ir) = 3 Then
        sprawdz_true_czy_false_na_podstawie = True
    Else
        sprawdz_true_czy_false_na_podstawie = False
    End If
    

End Function


Private Function IDataFromWizard_fooCountPns(szukaj_pustego As Variant, filename As Variant, sampath As Variant) As Variant


    Dim r As Range
    Set r = source_master.Cells(1, XWiz.pn)
    
    Do
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    IDataFromWizard_fooCountPns = r.Row - 1
End Function

Private Sub IDataFromWizard_init(Optional mwrkbk As Workbook, Optional e As E_FUP_FILTER, Optional e2 As E_RUN_REP_TYPE)

    Set wrkbk = mwrkbk
    Set source_master = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME)
    Set source_details = wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME)
    Set source_pickups = wrkbk.Sheets(XWiz.PICKUPS_SHEET_NAME)
    
    e_fup = e
    e_run_type = e2
    
End Sub

Private Sub IDataFromWizard_mrd1ConfStatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)
    
    
    Dim count_noks_in_mrd1_c_status_with_fma_star_resp As Range
    Set count_noks_in_mrd1_c_status_with_fma_star_resp = r.Offset(0, 3)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD1_confirmed_qty_dot__Status).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_confirmed_qty_dot__Status)
        
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20 = 0) Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD1_CONFIRMED_STATUS
            End If
        End If
        
        Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count))

    Next x
    
    With count_noks_in_mrd1_c_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            Else
                Exit For
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
End Sub

Private Sub IDataFromWizard_mrd1StatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, wypelnij_wstepnie_kolekcje_fma_resp As Boolean, Sh As StatusHandler)


    fn = CStr(filename)
    sp = CStr(sampath)

    If wypelnij_wstepnie_kolekcje_fma_resp Then
        ' wszystko ok
        Set kolekcja_fma_resp_zeby_tylko_raz_przegladac = Nothing
        Set kolekcja_fma_resp_zeby_tylko_raz_przegladac = New Collection
        
        Set d = Nothing
        Set d = New Dictionary
    Else
        MsgBox "cos poszlo nie tak z subie mrd1_status_logic"
    End If
    
    Dim count_noks_in_mrd1_o_status_with_fma_star_resp As Range
    Set count_noks_in_mrd1_o_status_with_fma_star_resp = r.Offset(0, 2)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    For x = 2 To count_pns
    
        If e_run_type = RUN_STD Then
            really_important_iteration_inside_mrd1_status _
                x, kolekcja_fma_resp_zeby_tylko_raz_przegladac, nok_counter, step, sampath, filename, Sh, count_pns
        Else
        
            If e_run_type = RUN_DUNS Then
                
                really_important_iteration_inside_mrd1_status_only_dunses _
                    x, kolekcja_fma_resp_zeby_tylko_raz_przegladac, nok_counter, step, sampath, filename, Sh, count_pns
                
            ElseIf e_run_type = RUN_PN Then
            
                really_important_iteration_inside_mrd1_status_only_pns _
                    x, kolekcja_fma_resp_zeby_tylko_raz_przegladac, nok_counter, step, sampath, filename, Sh, count_pns
                    
            End If
        End If
        
    Next x
    
    With count_noks_in_mrd1_o_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
    
    
    ' ile pn'ow pod fma resp
    ' count_noks_in_mrd1_o_status_with_fma_star_resp
    r.Offset(0, 1) = kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count
End Sub


Private Sub really_important_iteration_inside_mrd1_status_only_dunses(x, _
    kolekcja_fma_resp_zeby_tylko_raz_przegladac As Collection, _
    ByRef nok_counter As Variant, _
    ByRef step As Variant, _
    sampath As Variant, _
    filename As Variant, _
    ByRef Sh As StatusHandler, _
    ByRef count_pns As Variant)
    
    
    'strRng = Cells(x, XWiz.Responsibility).Address(1, 1, xlR1C1)
    'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
    'Result = CStr(ExecuteExcel4Macro(strRef))
    Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility)
    
    
    If Result Like XWiz.FMA_WITH_STARS Then
    
        If e_fup = E_FUP_FILTER_YES Then
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code)
        ElseIf e_fup = E_FUP_FILTER_NO Then
            Result = XWiz.XWIZ_FLAGA
        End If
        
        If (Result = XWiz.XWIZ_FLAGA) Or (UCase(Result) = UCase(ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code"))) Then
        
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.duns)
            ' Debug.Print UCase(ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("DUNSes"))
            txt = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("DUNSes")
            If InStr(1, CStr(txt), Result) > 0 Then
    
        
                kolekcja_fma_resp_zeby_tylko_raz_przegladac.Add x
            
                'strRng = Cells(x, XWiz.MRD1_Ordered_STATUS).Address(1, 1, xlR1C1)
                'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
                'Result = CStr(ExecuteExcel4Macro(strRef))
                Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_Ordered_STATUS)
                
                
                On Error Resume Next
                If Not Application.WorksheetFunction.IsNA(Result) Then
                
                    If Result = XWiz.G_NOK Then
                        nok_counter = nok_counter + 1
                        If (nok_counter Mod 20) = 0 Then
                            step = step + 1
                        End If
                        
                        
                        IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD1_ORDERED_STATUS
        
                    End If
                End If
            End If
        End If
        
    End If
    
    Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(count_pns))
    
End Sub


Private Sub really_important_iteration_inside_mrd1_status_only_pns(x, _
    kolekcja_fma_resp_zeby_tylko_raz_przegladac As Collection, _
    ByRef nok_counter As Variant, _
    ByRef step As Variant, _
    sampath As Variant, _
    filename As Variant, _
    ByRef Sh As StatusHandler, _
    ByRef count_pns As Variant)
    
    'strRng = Cells(x, XWiz.Responsibility).Address(1, 1, xlR1C1)
    'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
    'Result = CStr(ExecuteExcel4Macro(strRef))
    Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility)
    
    
    If Result Like XWiz.FMA_WITH_STARS Then
    
        If e_fup = E_FUP_FILTER_YES Then
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code)
        ElseIf e_fup = E_FUP_FILTER_NO Then
            Result = XWiz.XWIZ_FLAGA
        End If
        
        If (Result = XWiz.XWIZ_FLAGA) Or (UCase(Result) = UCase(ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code"))) Then
        
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.pn)
            
            If InStr(1, ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("pns"), Result) > 0 Then
    
        
                kolekcja_fma_resp_zeby_tylko_raz_przegladac.Add x
            
                'strRng = Cells(x, XWiz.MRD1_Ordered_STATUS).Address(1, 1, xlR1C1)
                'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
                'Result = CStr(ExecuteExcel4Macro(strRef))
                Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_Ordered_STATUS)
                
                
                On Error Resume Next
                If Not Application.WorksheetFunction.IsNA(Result) Then
                
                    If Result = XWiz.G_NOK Then
                        nok_counter = nok_counter + 1
                        If (nok_counter Mod 20) = 0 Then
                            step = step + 1
                        End If
                        
                        
                        IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD1_ORDERED_STATUS
        
                    End If
                End If
            End If
        End If
        
    End If
    
    Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(count_pns))
    
End Sub

Private Sub really_important_iteration_inside_mrd1_status(x, _
    kolekcja_fma_resp_zeby_tylko_raz_przegladac As Collection, _
    ByRef nok_counter As Variant, _
    ByRef step As Variant, _
    sampath As Variant, _
    filename As Variant, _
    ByRef Sh As StatusHandler, _
    ByRef count_pns As Variant)


    'strRng = Cells(x, XWiz.Responsibility).Address(1, 1, xlR1C1)
    'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
    'Result = CStr(ExecuteExcel4Macro(strRef))
    Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility)
    
    
    If Result Like XWiz.FMA_WITH_STARS Then
    
        If e_fup = E_FUP_FILTER_YES Then
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code)
        ElseIf e_fup = E_FUP_FILTER_NO Then
            Result = XWiz.XWIZ_FLAGA
        End If
        
        If (Result = XWiz.XWIZ_FLAGA) Or (UCase(Result) = UCase(ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code"))) Then
    
        
            kolekcja_fma_resp_zeby_tylko_raz_przegladac.Add x
        
            'strRng = Cells(x, XWiz.MRD1_Ordered_STATUS).Address(1, 1, xlR1C1)
            'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
            'Result = CStr(ExecuteExcel4Macro(strRef))
            Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_Ordered_STATUS)
            
            
            On Error Resume Next
            If Not Application.WorksheetFunction.IsNA(Result) Then
            
                If Result = XWiz.G_NOK Then
                    nok_counter = nok_counter + 1
                    If (nok_counter Mod 20) = 0 Then
                        step = step + 1
                    End If
                    
                    
                    IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD1_ORDERED_STATUS
    
                End If
            End If
        
        End If
        
    End If
    
    Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(count_pns))
End Sub

Private Sub IDataFromWizard_mrd2ConfStatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)


    Dim count_noks_in_mrd2_c_status_with_fma_star_resp As Range
    Set count_noks_in_mrd2_c_status_with_fma_star_resp = r.Offset(0, 6)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD2_confirmed_qty_dot__Status).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD2_confirmed_qty_dot__Status)
        
        
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
        
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20 = 0) Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD2_CONFIRMED_STATUS
            End If
        End If
        
        Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count))

    Next x
    
    With count_noks_in_mrd2_c_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            Else
                Exit For
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With

    
End Sub

Private Sub IDataFromWizard_mrd2StatusLogic(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)
    
    
    Dim count_noks_in_mrd2_o_status_with_fma_star_resp As Range
    Set count_noks_in_mrd2_o_status_with_fma_star_resp = r.Offset(0, 5)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD2_Ordered_STATUS).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD2_Ordered_STATUS)
        
        
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
            
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20 = 0) Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD2_ORDERED_STATUS
            End If
        End If
        
        Sh.mini_progress_increase (ILE_PODZIALOW_W_LECIMY_TUTAJ * CLng(kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count))

    Next x
    
    With count_noks_in_mrd2_o_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            Else
                Exit For
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
End Sub



Private Sub IDataFromWizard_podsumowanieSlownikaNaNwymArkuszu()


    Dim Sh As Worksheet ', nie trzeba jest juz takie pole rh As RecordsHandler
    Set Sh = ThisWorkbook.Sheets.Add
    
    Sh.Range("A1") = sp
    Sh.Range("B1") = fn
    cmnt_txt = create_cmnt_txt(sp, fn)
    Sh.Range("B1").addComment cmnt_txt
    Sh.Range("B1").Comment.Shape.TextFrame.AutoSize = True
    
    'kopia z create_cmnt_txt
    'y = LBound(lbls)
    '
    'For x = LBound(arr) To UBound(arr)
    '    create_cmnt_txt = create_cmnt_txt & CStr(lbls(y)) & ": " & CStr(arr(x)) & Chr(10)
    '    y = y + 1
    'Next x
    ' Split(fn, ".")
    Sh.Range("C1") = Split(fn, ".")(UBound(Split(fn, ".")))
    With Sh.Range("C1")
        .Font.Size = 1
    End With
    
    assign_sheet_name Sh, fn, 1
    
    
    Dim e As Range
    
    Set e = Sh.Range("A2")
    With e
        .Value = "#"
        .Offset(0, 1) = "ROW"
        .Offset(0, 2) = "PN"
        .Offset(0, 3) = "PN NM"
        .Offset(0, 4) = "DUNS"
        .Offset(0, 5) = "SUPP NM"
        .Offset(0, 6) = "RESP"
        .Offset(0, 7) = "FUP Code"
        
        
        .Offset(0, 8) = "MRD1 O."
        .Offset(0, 9) = "MRD1 Conf O."
        .Offset(0, 10) = "MRD1 PUS STATUS"
        
        .Offset(0, 11) = "MRD2 O."
        .Offset(0, 12) = "MRD2 Conf O."
        .Offset(0, 13) = "MRD2 PUS STATUS"
        
        .Offset(0, 14) = "TOTAL PUS STATUS"
        .Offset(0, 15) = "Delivery Confirmation"
        .Offset(0, 16) = "Delivery Confirmation STATUS"
        ' to ponoc ma byc kopia danych z mrd1 ordered date
        .Offset(0, 17) = "Order date"
        .Offset(0, 18) = "Comments"
    End With
    
    Set e = Sh.Range("A3")
    
    For Each Key In d.Keys
        
            ' =====================================
            Set rh = d(Key)
            
            e = rh.nok_counter
            e.Offset(0, 1) = Key
            e.Offset(0, 2) = rh.pn
            e.Offset(0, 3) = rh.pn_nm
            e.Offset(0, 4) = rh.duns
            e.Offset(0, 5) = rh.supp_nm
            e.Offset(0, 6) = rh.resp
            e.Offset(0, 7) = rh.fup
            
            ' statusy now
            ' MRD1
            If rh.MRD1_Ordered_STATUS Then
                e.Offset(0, 8) = XWiz.G_OK
            Else
                e.Offset(0, 8) = XWiz.G_NOK
            End If
            
            If rh.MRD1_confirmed_qty_Status Then
                e.Offset(0, 9) = XWiz.G_OK
            Else
                e.Offset(0, 9) = XWiz.G_NOK
            End If
            
            If rh.MRD1_PUS_STATUS Then
                e.Offset(0, 10) = XWiz.G_OK
            Else
                e.Offset(0, 10) = XWiz.G_NOK
            End If
            
            ' MRD2
            If rh.MRD2_Ordered_STATUS Then
                e.Offset(0, 11) = XWiz.G_OK
            Else
                e.Offset(0, 11) = XWiz.G_NOK
            End If
            
            If rh.MRD2_confirmed_qty_Status Then
                e.Offset(0, 12) = XWiz.G_OK
            Else
                e.Offset(0, 12) = XWiz.G_NOK
            End If
            
            If rh.MRD2_PUS_STATUS Then
                e.Offset(0, 13) = XWiz.G_OK
            Else
                e.Offset(0, 13) = XWiz.G_NOK
            End If
            
            If rh.TOTAL_PUS_STATUS Then
                e.Offset(0, 14) = XWiz.G_OK
            Else
                e.Offset(0, 14) = XWiz.G_NOK
            End If
            
            e.Offset(0, 15) = rh.Delivery_confirmation
            
            
            If rh.Delivery_confirmation_status Then
                e.Offset(0, 16) = XWiz.G_OK
            Else
                e.Offset(0, 16) = XWiz.G_NOK
            End If
            
            e.Offset(0, 17) = rh.MRD1_Ordered_Date
            e.Offset(0, 18) = rh.Comments
            ' =====================================
            
            Set e = e.Offset(1, 0)
            x = x + 1
    Next Key
End Sub

Private Function create_cmnt_txt(sp, fn) As String


    Dim lbls(0 To 64) As String
    lbls(1) = "PLT"
    lbls(2) = "PROJ"
    lbls(3) = "BG"
    lbls(4) = "MY"
    lbls(5) = "PHASE"
    lbls(6) = "BOM"
    lbls(7) = "PUS Date"
    lbls(8) = "PPAP gate"
    lbls(9) = "MRD"
    lbls(10) = "BUILD start"
    lbls(11) = "BUILD end"
    lbls(13) = "Coord"
    lbls(14) = "active"
    lbls(15) = "capacity check"
    
    ' new in details:
    lbls(16) = "MRD date"
    lbls(17) = "MRD reg routes"
    lbls(18) = "platform"
    lbls(19) = "trans account #"
    lbls(20) = "uID"
    
    
    create_cmnt_txt = CStr(sp) & Chr(10) & Chr(10)
    
    arr = Split(fn, ".")
    
    
    y = LBound(lbls)
    
    For x = LBound(arr) To UBound(arr)
        create_cmnt_txt = create_cmnt_txt & CStr(lbls(y)) & ": " & CStr(arr(x)) & Chr(10)
        y = y + 1
    Next x
End Function
Private Sub assign_sheet_name(ByRef Sh As Worksheet, nm, proba)

    If proba = 1 Then
        nm = remove_special_cases(nm)
    End If
    
    nm = shorten_plt_proj_nm_phase(nm)
    
    If Len(nm) > XWiz.MAX_SHEET_NAME_LEN Then nm = Left(nm, XWiz.MAX_SHEET_NAME_LEN)
    
    For Each s In Sh.Parent.Sheets
        If s.Name = nm & proba Then
            assign_sheet_name Sh, nm, proba + 1
            Exit Sub
        End If
    Next s

    Sh.Name = nm & proba
End Sub

Private Function shorten_plt_proj_nm_phase(nm) As String


    a = Split(nm, ".")
    
    
    
    's = "." & remove_special_cases(CStr(r)) & _
    '        "*" & Left(remove_special_cases(CStr(r.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
    '        "*" & remove_special_cases(CStr(r.Offset(0, 2))) & _
    '        "*" & Left(remove_special_cases(CStr(r.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"
    
    
    nm = "."
    ' od ktorego jest plt - zaczynamy od gory co by zrobic trick w warunku if z and
    y = UBound(a)
    For x = LBound(a) To UBound(a)
    
        If Len(Trim(CStr(a(x)))) > 0 Then
        
            If x < y Then
                y = x
            End If
            
            If x = y + 1 Then
                nm = nm & Left(a(x), XWiz.G_CUT_PROJECT) & "."
            ElseIf x = y + 4 Then
                nm = nm & Left(a(x), XWiz.G_CUT_PHAZE) & "."
            Else
                nm = nm & a(x) & "."
            End If
        
        End If
    Next x
    
    
    shorten_plt_proj_nm_phase = CStr(nm)
    
End Function

Private Sub IDataFromWizard_runMainLogic(r As Range, sampath, arr, Sh As StatusHandler, ByRef c As Collection)

   
    If e_run_type < NEW_RUN_ALL Then
    
        jesli_run_type_jest_mniejszy_niz_new_run_all r, sampath, arr, Sh
    
    ElseIf e_run_type = NEW_RUN_ALL Then
        ' tutaj robimy ujecie calosciowe - zupelenie niezalezna logika
        
        Set kolekcja_obiektow_typu_RecordHandlerAllApproach = c
        
        jesli_run_type_to_new_run_all r, sampath, arr, Sh
        
        Set r = r.Parent.Cells(r.Row, 1)
    End If
End Sub

Private Sub jesli_run_type_to_new_run_all(r As Range, sampath, arr, Sh As StatusHandler)

    ' Set all_sh = ThisWorkbook.Sheets(XWiz.ALL_SHEET_NAME)
    
    For x = LBound(arr) To UBound(arr) - 1
        sampath = sampath & arr(x) & "\"
    Next x
    
    filename = arr(UBound(arr))
    
    ' 1 progres
    Sh.mini_progress_increase POD_MINI_PROGRES_DLA_REP_ALL
    
 
    For x = XWiz.plt To XWiz.E_UNIQUE_ID
        Result = wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x))
        
        If Trim(Result) = "" And x = XWiz.plt Then
            r.Clear
            r.ClearComments
            Set r = r.Offset(-1, 0)
            wrkbk.Close False
            Exit Sub
        End If
        
        If IsDate(Result) Then
            r.Offset(0, x - 1) = Format(Result, "yyyy-mm-dd")
        Else
            r.Offset(0, x - 1) = Result
        End If
    Next x
    
    ' 2 progres
    Sh.mini_progress_increase POD_MINI_PROGRES_DLA_REP_ALL
    
    dalej = XWiz.E_UNIQUE_ID + 1
    Set r = r.Offset(0, dalej - 1)
    ' r.Offset(0, x - 1)
    ' count pn's
    count_pns = 0
    ' argumenty w przypadku wywolania obiektu typu ExternalReadOnlyWokrbook nie ma sensu
    ' zatem zostawiwam wszystko puste
    count_pns = IDataFromWizard_fooCountPns(0, "", "")
    
    r = count_pns - 1
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    
    ' 3 progres
    Sh.mini_progress_increase POD_MINI_PROGRES_DLA_REP_ALL
    
    
    For i = 2 To count_pns
        
        iteration_in_run_type_all sampath, filename, i, nok_counter, step, r.Parent.Cells(r.Row, 1)
        
    Next i
    
    ' 4 progres po iteracji
    Sh.mini_progress_increase POD_MINI_PROGRES_DLA_REP_ALL
    
    
    With r.Offset(0, 1)
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
    
    ' 5 progres
    Sh.mini_progress_increase POD_MINI_PROGRES_DLA_REP_ALL
        
    ' Debug.Print kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count
    ' sh.progress_increase
    
End Sub

Private Sub iteration_in_run_type_all(sampath, filename, x, nok_counter, step, rr As Range)

    Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility)
    
    
    If Result Like XWiz.FMA_WITH_STARS Then


    
    
        'strRng = Cells(x, XWiz.MRD1_Ordered_STATUS).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Delivery_confirmation))
        
        nok_counter = nok_counter + 1
        If (nok_counter Mod 20) = 0 Then
            step = step + 1
        End If
        
        Set rhaa = New RecordHandlerAllApproach
        ' tutaj assign section
        ' ----------------------
        ' plt = 1
        ' PROJECT
        ' BIW_GA ' BIW or GA
        ' my
        ' PHAZE
        ' bom
        ' PICKUP_DATE
        ' PPAP_GATE
        ' mrd
        ' BUILD_START
        ' BUILD_END
        ' koordynator
        ' E_ACTIVE
        ' CAPACITY_CHECK
        ' E_MRD_DATE
        ' E_MRD_REG_ROUTES
        ' E_PLATFORM
        ' E_TRANSPORTATION_ACCOUNT_NUMBER
        ' E_UNIQUE_ID
        
        On Error GoTo ErrHandler:
        With rhaa
            .plt = rr
            .proj = rr.Offset(0, XWiz.PROJECT - 1)
            .bg = rr.Offset(0, XWiz.BIW_GA - 1)
            .my = rr.Offset(0, XWiz.my - 1)
            .faza = rr.Offset(0, XWiz.PHAZE - 1)
            
            .mrd = Trim(rr.Offset(0, XWiz.mrd - 1))
            .mrdd = Trim(rr.Offset(0, XWiz.E_MRD_DATE - 1))
            
            .koordynator = Trim(rr.Offset(0, XWiz.koordynator - 1))
            
            .unique_id = Trim(rr.Offset(0, XWiz.E_UNIQUE_ID - 1))
            
            .del_conf = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Delivery_confirmation))
            
            .duns = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.duns))
            .pn = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.pn))
            .resp = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Responsibility))
            .fup = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code))
            
            .cmnts = Trim(wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.Comments))
        End With
        
        On Error GoTo ErrHandler:
        With wrkbk.Sheets(XWiz.MASTER_SHEET_NAME)
                
            rhaa.supp_nm = Trim(.Cells(x, XWiz.Supplier_Name))
            rhaa.total_qty = Trim(.Cells(x, XWiz.total_qty))
            
            ' mrd1
            rhaa.ordered_date = Trim(.Cells(x, XWiz.MRD1_Ordered_Date))
            rhaa.ordered_qty = Trim(.Cells(x, XWiz.MRD1_Ordered_QTY))
            rhaa.confirmed_qty = Trim(.Cells(x, XWiz.MRD1_confirmed_qty))
            
            rhaa.ppap_status = Trim(.Cells(x, XWiz.ppap_status))
            rhaa.ppap_gate = Trim(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Cells(x, XWiz.ppap_gate))
            
            rhaa.fst_pickup_date = Trim(.Cells(x, XWiz.First_Confirmed_PUS_Date))
            
        End With
        
        
        IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_Delivery_confirmation
        
ErrHandler:
        ' ----------------------
        kolekcja_obiektow_typu_RecordHandlerAllApproach.Add rhaa
        
        
        
    End If
End Sub


Private Sub jesli_run_type_jest_mniejszy_niz_new_run_all(r As Range, sampath, arr, Sh As StatusHandler)

    For x = LBound(arr) To UBound(arr) - 1
        sampath = sampath & arr(x) & "\"
    Next x
    
    filename = arr(UBound(arr))
    
    Dim strRng As String, strRef As String, Result As String
    For x = XWiz.plt To XWiz.E_UNIQUE_ID
        'strRng = Range("B" & CStr(x)).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.DETAILS_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        
        Result = wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x))
        filename2 = filename2 & "." & Result
        
        
        
        If x = XWiz.E_UNIQUE_ID Then
            ctxt = ""
            
            ctxt = _
                CStr(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x)))
            
            With r.Offset(0, XWiz.E_ACTIVE - 1)
                .ClearComments
                .addComment CStr(ctxt)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
        End If
        
        
        If x = XWiz.E_TRANSPORTATION_ACCOUNT_NUMBER Then
        
            ctxt = ""
            
            ctxt = ctxt & "MRD Date: " & _
                CStr(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x - 3))) & Chr(10)
            ctxt = ctxt & "MRD reg routes: " & _
                CStr(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x - 2)))
            
            With r.Offset(0, XWiz.mrd - 1)
                .ClearComments
                .addComment CStr(ctxt)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            
            ctxt = "Transportation account number: " & CStr(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x)))
            ctxt = ctxt & Chr(10) & "Platform: " & _
                CStr(wrkbk.Sheets(XWiz.DETAILS_SHEET_NAME).Range("B" & CStr(x - 1)))
            r.Offset(0, XWiz.PHAZE - 1).ClearComments
            r.Offset(0, XWiz.PHAZE - 1).addComment CStr(ctxt)
            
            With r.Offset(0, XWiz.PHAZE - 1)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            
            
        End If
        
        
        
        If x = XWiz.PROJECT Then
            ' podreczna pamiec w cfg dla calych stringow dla project
            ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("T" & CStr(r.Row)) = Result
            r.Offset(0, x - 1).Value = Left(CStr(Result), XWiz.G_CUT_PROJECT)
        ElseIf x = XWiz.PHAZE Then
            ' podreczna pamiec w cfg dla calych stringow dla phase
            ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("U" & CStr(r.Row)) = Result
            r.Offset(0, x - 1).Value = Left(CStr(Result), XWiz.G_CUT_PHAZE)
        ElseIf x = XWiz.PICKUP_DATE Or x = XWiz.ppap_gate Then
            On Error Resume Next
            r.Offset(0, x - 1).Value = Format(CStr(Result), "yyyy-mm-dd")
        Else
            r.Offset(0, x - 1).Value = CStr(Result)
        End If
        
        
    Next x
    
    
    Set r = r.Offset(0, x - 1 - XWiz.DODATKOWE_POLA_OD_DETAILS) ' teraz juz 5
    
    Sh.mini_progress_increase ILE_PODZIALOW_W_LECIMY_TUTAJ
    
    ' count pn's
    count_pns = 0
    ' argumenty w przypadku wywolania obiektu typu ExternalReadOnlyWokrbook nie ma sensu
    ' zatem zostawiwam wszystko puste
    count_pns = IDataFromWizard_fooCountPns(0, "", "")
    
    r = count_pns - 1
    
    Sh.mini_progress_increase ILE_PODZIALOW_W_LECIMY_TUTAJ
    
    ' mrd1 ordered status
    ' tutaj jedyna zmiana filename2 poniewaz ma wplyw na logike dzialania
    IDataFromWizard_mrd1StatusLogic count_pns, filename2, sampath, r, True, Sh
    IDataFromWizard_mrd1ConfStatusLogic count_pns, filename, sampath, r, Sh
    IDataFromWizard_subMrd1PusStatus count_pns, filename, sampath, r, Sh
    
    IDataFromWizard_mrd2StatusLogic count_pns, filename, sampath, r, Sh
    IDataFromWizard_mrd2ConfStatusLogic count_pns, filename, sampath, r, Sh
    IDataFromWizard_subMrd2PusStatus count_pns, filename, sampath, r, Sh
    
    IDataFromWizard_subTotalPusStatus count_pns, filename, sampath, r, Sh
    
    IDataFromWizard_deliveryConfirmationLogic count_pns, sampath, r, wrkbk, Sh
    
    IDataFromWizard_podsumowanieSlownikaNaNwymArkuszu
    

End Sub

Private Sub IDataFromWizard_subMrd1PusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)



    Dim count_noks_in_pus_status_with_fma_star_resp As Range
    Set count_noks_in_pus_status_with_fma_star_resp = r.Offset(0, 4)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD1_Total_PUS_STATUS).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD1_Total_PUS_STATUS)
        
        
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
            
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20) = 0 Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD1_PUS_STATUS
            End If
        End If
        Sh.mini_progress_increase ILE_PODZIALOW_W_LECIMY_TUTAJ * kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count

    Next x
    
    
    With count_noks_in_pus_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            Else
                Exit For
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
            
    End With
End Sub

Private Sub IDataFromWizard_subMrd2PusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)



    Dim count_noks_in_pus_status_with_fma_star_resp As Range
    Set count_noks_in_pus_status_with_fma_star_resp = r.Offset(0, 7)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.MRD2_Total_PUS_STATUS).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.MRD2_Total_PUS_STATUS)
        
        
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20) = 0 Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_MRD2_PUS_STATUS
            End If
        End If
        
        Sh.mini_progress_increase ILE_PODZIALOW_W_LECIMY_TUTAJ * kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count

    Next x
    
    
    With count_noks_in_pus_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End With
End Sub

Private Sub IDataFromWizard_subTotalPusStatus(count_pns As Variant, filename As Variant, sampath As Variant, r As Variant, Sh As StatusHandler)



    Dim count_noks_in_pus_status_with_fma_star_resp As Range
    Set count_noks_in_pus_status_with_fma_star_resp = r.Offset(0, 8)
    
    
    nok_counter = 0
    step = 1
    For q = 1 To 1024
        for_comment(q) = ""
    Next q
    ' od razu lecimy przez wybrane ino wiersze ktore zostaly wylapane
    For Each x In kolekcja_fma_resp_zeby_tylko_raz_przegladac
    
    
        
        'strRng = Cells(x, XWiz.total_pus_status).Address(1, 1, xlR1C1)
        'strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
        'Result = CStr(ExecuteExcel4Macro(strRef))
        Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.TOTAL_PUS_STATUS)
        On Error Resume Next
        If Not Application.WorksheetFunction.IsNA(Result) Then
            If Result = XWiz.G_NOK Then
                nok_counter = nok_counter + 1
                If (nok_counter Mod 20) = 0 Then
                    step = step + 1
                End If
                
                IDataFromWizard_addComment for_comment(step), nok_counter, x, sampath, filename, E_F_TOTAL_PUS_STATUS
            End If
        End If
        
        Sh.mini_progress_increase ILE_PODZIALOW_W_LECIMY_TUTAJ * kolekcja_fma_resp_zeby_tylko_raz_przegladac.Count

    Next x
    
    
    With count_noks_in_pus_status_with_fma_star_resp
        .Value = nok_counter
        .addComment for_comment(1)
        For q = 2 To 1024
        
            If for_comment(q) <> "" Then
                .Comment.Text .Comment.Text & for_comment(q)
            End If
        Next q

        '.Comment.Shape.Width = G_CMNT_WIDTH
        '.Comment.Shape.Height = (nok_counter + 1) * XWiz.G_CMNT_HEIGHT
        With .Comment.Shape
            .Top = 10
            .Left = 10
        End With
        
        
        With .Comment.Shape.TextFrame
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
        
    End With
End Sub
