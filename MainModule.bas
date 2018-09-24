Attribute VB_Name = "MainModule"
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


Public Sub static_ribbon_button(ictrl As IRibbonControl)
    
    StaticFrm.show
End Sub


Public Sub run_static_fup()


    WybierzFUPCode.show

    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_STATIC, E_FUP_FILTER_YES, RUN_STD
End Sub

Public Sub run_dynamic_fup(ictrl As IRibbonControl)

    WybierzFUPCode.show

    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_DYNAMIC, E_FUP_FILTER_YES, RUN_STD
End Sub

Public Sub run_static_fup_on_pn()


    WybierzFUPCode.show
    PN_DUNS_Frm.PNOptionButton.Value = True
    PN_DUNS_Frm.DUNSOptionButton.Value = False
    PN_DUNS_Frm.show

    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_STATIC, E_FUP_FILTER_YES, RUN_PN
End Sub

Public Sub run_dynamic_fup_on_pn(ictrl As IRibbonControl)

    WybierzFUPCode.show
    PN_DUNS_Frm.PNOptionButton.Value = True
    PN_DUNS_Frm.DUNSOptionButton.Value = False
    PN_DUNS_Frm.show

    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_DYNAMIC, E_FUP_FILTER_YES, RUN_PN
End Sub

Public Sub run_static_fup_on_duns()


    WybierzFUPCode.show
    PN_DUNS_Frm.PNOptionButton.Value = False
    PN_DUNS_Frm.DUNSOptionButton.Value = True
    PN_DUNS_Frm.show

    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_STATIC, E_FUP_FILTER_YES, RUN_DUNS
End Sub

Public Sub run_dynamic_fup_on_duns(ictrl As IRibbonControl)

    WybierzFUPCode.show
    PN_DUNS_Frm.PNOptionButton.Value = False
    PN_DUNS_Frm.DUNSOptionButton.Value = True
    PN_DUNS_Frm.show
    ' przypisanie odbedzie sie podczas klikniecia
    'ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code") = wybrany_fup_code_z_forma
    
    inner_start E_DYNAMIC, E_FUP_FILTER_YES, RUN_DUNS
End Sub

Public Sub run_static()
    
    inner_start E_STATIC, E_FUP_FILTER_NO, RUN_STD
End Sub


Public Sub run_all(ictrl As IRibbonControl)
    
    
    inner_start E_DYNAMIC, E_FUP_FILTER_NO, NEW_RUN_ALL
End Sub

Public Sub run_dynamic(ictrl As IRibbonControl)
    
    
    inner_start E_DYNAMIC, E_FUP_FILTER_NO, RUN_STD
End Sub

Public Sub run_static_on_pn()
    PN_DUNS_Frm.PNOptionButton.Value = True
    PN_DUNS_Frm.DUNSOptionButton.Value = False
    PN_DUNS_Frm.show
    inner_start E_STATIC, E_FUP_FILTER_NO, RUN_PN
End Sub

Public Sub run_dynamic_on_pn(ictrl As IRibbonControl)
    PN_DUNS_Frm.PNOptionButton.Value = True
    PN_DUNS_Frm.DUNSOptionButton.Value = False
    PN_DUNS_Frm.show
    inner_start E_DYNAMIC, E_FUP_FILTER_NO, RUN_PN
End Sub

Public Sub run_static_on_duns()
    PN_DUNS_Frm.PNOptionButton.Value = False
    PN_DUNS_Frm.DUNSOptionButton.Value = True
    PN_DUNS_Frm.show
    inner_start E_STATIC, E_FUP_FILTER_NO, RUN_DUNS
End Sub

Public Sub run_dynamic_on_duns(ictrl As IRibbonControl)
    PN_DUNS_Frm.PNOptionButton.Value = False
    PN_DUNS_Frm.DUNSOptionButton.Value = True
    PN_DUNS_Frm.show
    inner_start E_DYNAMIC, E_FUP_FILTER_NO, RUN_DUNS
End Sub

Public Sub przesun_dane_do_rep_fup(ictrl As IRibbonControl)
    Application.ScreenUpdating = False
    wyjmij_dane_repa_po_fupie
    Application.ScreenUpdating = True
    
    MsgBox "ready!"
End Sub

Public Sub przesun_dane_do_rep(ictrl As IRibbonControl)
    Application.ScreenUpdating = False
    wyjmij_dane_repa
    Application.ScreenUpdating = True
    
    MsgBox "ready!"
End Sub


Public Sub wyjmij_dane_repa_po_fupie()



    If sprawdz_czy_jestes_na_sideowym_arkuszu() Then


        WybierzFUPCode.show
        
        Dim wh As WrkHandler
        Set wh = New WrkHandler
        
        With wh
            .setEFup E_FUP_FILTER_YES
            .setSideWrksh ThisWorkbook.ActiveSheet
            .wyczysc_czesciowo_arkusz_rep_fup
            .extractDataFromRep
        End With
        
        Dim art As AddRedToNoks
        Set art = New AddRedToNoks
        art.prepare_range_and_colour_noks_red ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
        art.colour_blue_this_week_on_bom_pus_date_mrd_and_build ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
    End If
End Sub

Public Sub wyjmij_dane_repa()


    If sprawdz_czy_jestes_na_sideowym_arkuszu() Then
        ' WybierzFUPCode.show
        
        Dim wh As WrkHandler
        Set wh = New WrkHandler
        
        With wh
            .setEFup E_FUP_FILTER_NO
            .setSideWrksh ThisWorkbook.ActiveSheet
            .wyczysc_czesciowo_arkusz_rep
            .extractDataFromRep
        End With
        
        Dim art As AddRedToNoks
        Set art = New AddRedToNoks
        art.prepare_range_and_colour_noks_red ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
        art.colour_blue_this_week_on_bom_pus_date_mrd_and_build ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
    End If
End Sub

Private Function sprawdz_czy_jestes_na_sideowym_arkuszu() As Boolean
    sprawdz_czy_jestes_na_sideowym_arkuszu = False
    
    Dim r As Range
    Set r = ThisWorkbook.ActiveSheet.Cells(2, 1)
    
    If r.Value = Trim(CStr(XWIZ_TXT_CMNT_HASH)) Then
        If r.Offset(0, 1).Value = "ROW" Then
            If r.Offset(0, 2).Value = "PN" Then
            
                sprawdz_czy_jestes_na_sideowym_arkuszu = True
            End If
        End If
    End If
End Function




' START JEST TUTAJ wlasciwie dla kazdej funkcjonalnosci run
Public Sub inner_start(e As E_CREAT_COLLECTION_TYPE, e_fup As E_FUP_FILTER, e_run_type As E_RUN_REP_TYPE)


    If e_run_type = NEW_RUN_EXTENDED Then
    
        ' go back to extended module - this is first time that i do it in this way
        innerRunExtended
    Else
    
        
        unhide_all_rows_and_all_columns ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
        unhide_all_rows_and_all_columns ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
    
    
    
        ' a oto najwazniejsza deklaracja
        ' wh to zmienna typu WrkHandler
        ' glowna klasa tej aplikacja sterujaca flow informacyjnym
        Dim wh As WrkHandler
        Set wh = New WrkHandler
        
        With wh
        
            ' to jest w sumie obsoletowa funkcja
            ' potrzebna byla gdy jeszcze probowalem linkowac arkusze side'owe
            ' z lista na rep lub rep fup
            ' ale za racji tego ze nie pamietam czy nie bylo jakis akcji dodatkowych
            ' side'owych w tej procedurze nie chce ryzykowac i usuwac jej
            ' z racji tego ze kod tutaj pozostawia wiele do zyczenia
            .wyczysc_cfg_sheet_i_jej_tmp_list_na_name_i_phase
            
            
            ' jest OK
            ' .przejrzyjListe_TEST
            CzekajForm.show vbModeless
            Application.DisplayStatusBar = True
            
            ' te status bary nie dzialaja do konca tak jakbym tego oczekiwal....
            Application.StatusBar = "odnajduje pliki typu Wizard, ktore sa zgodne ze wzorcem..."
            ' stworz kolekcje uruchomi sie podczas zmiany work_pathu ponizej
            
            
            ' tutaj filtr pod rep fupa
            .setEFup e_fup
            
            ' a tutaj rodzaj raportu ze wzgledu na to czy robimy po pn po duns, czy w ogole dziwnie
            ' po wszystkim
            .setRunType e_run_type
            
            
            ' no i tutaj glowna jazda
            ' sama literka "e" chowa logike ktora ustawia w jaki sposob dalej bedziemy sie poruszac
            ' po dysku X
            If e = E_DYNAMIC Then
                ' std - na vpn cholernie dlugo to trwa
                .stworz_kolekcje XWIZ.XWIZ_PATH_FOR_SEARCHING
            ElseIf e = E_STATIC Then
                
                ' no i nasz problematyczny kawalek kodu, ktory dokad wprowadzilem
                ' run all ma problemy z prawidlowym uruchomieniem
                .stworz_kolekcje_na_podstawie_statycznych_pathow_z_arkusza_config
            End If
            CzekajForm.hide
            
            ' zgodnie z logika tego kodu
            ' kolekcje sa juz uzupelnione i mozemy leciec z koksem
            ' rozwiazanie jest to na tyle indywidualne ze dziwne ze kod sie zaafektowal
            ' searching for bug!
            ' investigation regarding static run!
            '
            With XWIZ.BedzieNieBedzieForm
                .connect_with_wrk_handler wh
                .wypelnij_work_path
                .wypelnij_listboxy
                .show
            End With
            
            
            
        End With
    End If

    
End Sub


Public Sub zmien_format_na(a As String, s As String)
    Dim arkusz As Worksheet
    Set arkusz = ThisWorkbook.Sheets(CStr(a))
    With arkusz
        .Columns("O:P").NumberFormat = CStr(s)
    End With
End Sub
