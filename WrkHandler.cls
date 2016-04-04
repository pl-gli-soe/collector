VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrkHandler"
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



' master & details ref
Private masterSourceSheetName As String
Private detailsSourceSheetName As String

Private parentPath As String
Private parentPathBackup As String

' klucz nazwy Quartera
Private prefixName As String
Private midfixName As String
Private postfixName As String

Private objFSO As Object
Private objFolder As Object
Private objFile As Object
Private i As Integer


Private e_fup As E_FUP_FILTER
Private e_run_type As E_RUN_REP_TYPE


Private fileCollection As Collection
' rep
Private rep As Worksheet
Private rep_fup As Worksheet
Private rep_all As Worksheet
Private r As Range


Private q As Workbook
Private qs As Worksheet


Private mySh As StatusHandler

' Private d As Dictionary

Private ktory_side_ma_zostac_zaktualizowany As Worksheet
Private linijka_do_zaktualizowania As Range


Private idfw As IDataFromWizard


Private coll As Collection
Private rh As RecordsHandler


' tutaj tandem
Private rhaa As RecordHandlerAllApproach
Private kolekcja_obiektow_typu_RecordHandlerAllApproach As Collection


Public Function get_e_run_type() As E_RUN_REP_TYPE
    get_e_run_type = e_run_type
End Function

Public Function get_e_fup() As E_FUP_FILTER
    get_e_fup = e_fup
End Function


Public Sub przejdz_po_kolei_przez_kolekcje_nazw_i_pobierz_dane(ByRef Sh As StatusHandler)
    ' for visibiity on macro progress
    
    If fileCollection.Count > 0 Then
        
        ' Application.Visible = False
        If e_run_type < NEW_RUN_ALL Then
            If e_fup = E_FUP_FILTER_NO Then
            
                rep.Range("A1").Value = XWiz.DETAILS_SHEET_NAME
                Set r = rep.Range("A1").Offset(2, 0)
            ElseIf e_fup = E_FUP_FILTER_YES Then
            
                rep_fup.Range("A1").Value = XWiz.DETAILS_SHEET_NAME
                Set r = rep_fup.Range("A1").Offset(2, 0)
            End If
            
        ' rep all
        ElseIf e_run_type = NEW_RUN_ALL Then
            
                Set r = rep_all.Range("A2")
        End If
        
        
        Set kolekcja_obiektow_typu_RecordHandlerAllApproach = New Collection
        
        For Each s In fileCollection
        
        
            Set r = r.Parent.Cells(r.Row, 1)
            On Error Resume Next
            r.addComment s
            With r
                '.Comment.Shape.Height = 20
                '.Comment.Shape.Width = 400
                
                With .Comment.Shape.TextFrame
                    .Characters.Font.Name = "Consolas"
                    .Characters.Font.Size = 10
                    .AutoSize = True
                End With
            End With
            
            
           
            Me.lecimy_tutaj_jawne_otwieranie_pliku CStr(s), r, Sh, kolekcja_obiektow_typu_RecordHandlerAllApproach
            Set r = r.Offset(1, 0)
            ' sh.progress_increase
        Next s
        
        
        ' przygotowanie jednego duzego outputu
        If e_run_type = NEW_RUN_ALL Then
            prepare_onepager_for_all_files kolekcja_obiektow_typu_RecordHandlerAllApproach
        Else
            
        End If
        
        ' Application.Visible = True
    Else
        MsgBox "nie ma czego przegladac"
    End If
End Sub

Private Sub prepare_onepager_for_all_files(c As Collection)

    If c.Count > 0 Then
    
        ' rhaa jako element kolekcji
        ' lecimy iteracja
        Dim allDelConfs As Worksheet
        Set allDelConfs = ThisWorkbook.Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME)
        
        clear_allDelConfs allDelConfs
         
        'labelki
        x = 1
        With allDelConfs
            .Cells(x, 1) = "PLT"
            .Cells(x, 2) = "PROJ"
            .Cells(x, 3) = "BG"
            .Cells(x, 4) = "MY"
            .Cells(x, 5) = "FAZA"
            .Cells(x, 6) = "MRD"
            .Cells(x, 7) = "MRDd"
            .Cells(x, 8) = "COORD"
            .Cells(x, 9) = "RESP"
            .Cells(x, 10) = "FUP"
            .Cells(x, 11) = "PN"
            .Cells(x, 12) = "DEL CONF"
            .Cells(x, 13) = "Comment"
            
        End With
        
        x = 2
        For Each rhaa In c
            
            ' ============================
            With allDelConfs
                .Cells(x, 1) = rhaa.plt
                .Cells(x, 2) = rhaa.proj
                .Cells(x, 3) = rhaa.bg
                .Cells(x, 4) = rhaa.my
                .Cells(x, 5) = rhaa.faza
                .Cells(x, 6) = rhaa.mrd
                .Cells(x, 7) = rhaa.convert_mrd_to_date(Date)
                .Cells(x, 8) = rhaa.koordynator
                .Cells(x, 9) = rhaa.resp
                .Cells(x, 10) = rhaa.fup
                .Cells(x, 11) = rhaa.pn
                .Cells(x, 12) = rhaa.del_conf
                .Cells(x, 13) = rhaa.cmnts
                
            End With
            ' ============================
            
            x = x + 1
        Next rhaa
    Else
        MsgBox "cos poszlo nie tak poniewaz kolekcja_obiektow_typu_RecordHandlerAllApproach jest pusta"
    End If

End Sub

Private Sub clear_allDelConfs(adc_sh As Worksheet)
    
    Dim ir1 As Range, ir2 As Range
    Set ir1 = adc_sh.Range("A1")
    
    If adc_sh.Range("b1") <> "" And adc_sh.Range("a2") <> "" Then
        Set ir2 = adc_sh.Range("A1").End(xlDown).End(xlToRight)
    Else
        Set ir2 = ir1
    End If
    
    adc_sh.Range(ir1, ir2).Clear
End Sub


Public Sub oddaj_cale_nazwy_dla_project_i_faz()

    Dim cfg As Worksheet, wsh As Worksheet, cr As Range, rr As Range
    Set cfg = ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    
    
    If e_fup = E_FUP_FILTER_NO Then
        Set wsh = rep
    Else
        Set wsh = rep_fup
        
    End If
    
    Set cr = cfg.Range("T3")
    Set rr = wsh.Range("B3")
    
    
    Do
        rr = cr
        rr.Offset(0, 3) = cr.Offset(0, 1)
        
        Set rr = rr.Offset(1, 0)
        Set cr = cr.Offset(1, 0)
    Loop Until Trim(CStr(cr)) = ""
End Sub


Public Function getColl() As Collection

    Set getColl = Nothing
    
    On Error Resume Next
    If Not coll Is Nothing Then
        Set getColl = coll
    End If
End Function


Public Sub setSideWrksh(ish As Worksheet)
    Set ktory_side_ma_zostac_zaktualizowany = ish
End Sub


Public Sub setEFup(e As E_FUP_FILTER)
    e_fup = e
End Sub

Public Sub setRunType(e As E_RUN_REP_TYPE)
    e_run_type = e
End Sub

Public Function countCollection() As Long
    countCollection = CLng(fileCollection.Count)
End Function

Public Function getCollection() As Collection
    Set getCollection = fileCollection
End Function

Public Sub refreshCollection(c As Collection)
    Set fileCollection = Nothing
    Set fileCollection = c
End Sub


Public Sub setMasterShNm(s As String)
    masterSourceSheetName = s
End Sub

Public Function setDetailsShNm(s As String)
    detailsSourceSheetName = s
End Function


Private Sub Class_Initialize()
    
    parentPath = CStr(XWiz.XWIZ_PATH_FOR_SEARCHING)
    Me.setDetailsShNm CStr(XWiz.DETAILS_SHEET_NAME)
    Me.setMasterShNm CStr(XWiz.MASTER_SHEET_NAME)
    
    
    
    ' init name
    prefixName = CStr(XWiz.XWIZ_FILE_PREFIX)
    midfixName = CStr(XWiz.XWIZ_FILE_MIDFIX)
    postfixName = CStr(XWiz.XWIZ_FLE_POSTFIX_VERSION)
    
    ' Q_Iza_postfix_catcher.xlsm
    
    Set fileCollection = New Collection
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Set rep = ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME)
    Set rep_fup = ThisWorkbook.Sheets(XWiz.REP_FUP_SHEET_NAME)
    Set rep_all = ThisWorkbook.Sheets(XWiz.ALL_SHEET_NAME)
    
End Sub

Private Sub Class_Terminate()
    Set fileCollection = Nothing
End Sub



' typowo testowy sub
Public Sub przejrzyjListe_TEST()
    rekurencyjne_przegladanie_katalogow CStr(parentPath), CStr(prefixName) & "*" & CStr(midfixName) & "*" & CStr(postfixName) & "*"
    If fileCollection.Count > 0 Then
        
        Dim i As Variant
        For Each i In fileCollection
            Debug.Print CStr(i)
        Next i
    Else
        MsgBox "collection of quarters is empty!"
    End If
End Sub


Public Function stworz_kolekcje(my_path As String) As Boolean


    Set fileCollection = New Collection

    Set objFolder = Nothing
    On Error Resume Next
    Set objFolder = objFSO.GetFolder(CStr(my_path))

    If Not objFolder Is Nothing Then
        rekurencyjne_przegladanie_katalogow CStr(my_path), CStr(prefixName) & "*" & CStr(midfixName) & "*" & CStr(postfixName) & "*"
    End If
End Function

Public Function stworz_kolekcje_na_podstawie_statycznych_pathow_z_arkusza_config()

    Dim csh As Worksheet
    Set csh = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    
    Dim r As Range
    Set r = csh.Range("a2")
    
    Set fileCollection = New Collection
    Set objFolder = Nothing
    
    Do
    
        my_path = CStr(r.Offset(0, 1))
        
        On Error Resume Next
        Set objFolder = objFSO.GetFolder(CStr(my_path))
    
        If Not objFolder Is Nothing Then
            przejrzenie_jednego_katalogu CStr(my_path), CStr(prefixName) & "*" & CStr(midfixName) & "*" & CStr(postfixName) & "*"
        End If
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function


Private Sub przejrzenie_jednego_katalogu(s As String, pattern As String)


     Dim objtmpfolder As Object
    'Get the folder object
    Set objFolder = objFSO.GetFolder(CStr(s))

    For Each objFile In objFolder.Files
        '    'print file name
        ' to jest do stestowania na dzien nastepny - zoabczymy czy ten read only dziala tak jak nalezy.
        Set q = Nothing
        'Set q = Workbooks.Open(CStr(objFile.Name), , True)
        'fileCollection.Add q
        
        ' Debug.Print CStr(objFile.Path)
        ' Debug.Print CStr(objFile.Name)
        ' ten warunek w przyszlosci ma zostac zmieniony
        If CStr(objFile.Name) Like CStr(pattern) Then
            ' Debug.Print CStr(objFile.Path)
            
            jest_juz_taki_i_go_nie_dodamy = False
            
            If fileCollection.Count > 0 Then
                For Each ic In fileCollection
                    If CStr(ic) = CStr(objFile.Path) Then
                        jest_juz_taki_i_go_nie_dodamy = True
                    End If
                Next ic
            End If
            
            If jest_juz_taki_i_go_nie_dodamy = False Then
                fileCollection.Add CStr(objFile.Path)
            End If
            
        End If
        '    'print file path
        '    Cells(i + 1, 2) = objFile.Path
        '    i = i + 1
    Next objFile
End Sub

Private Sub rekurencyjne_przegladanie_katalogow(s As String, pattern As String)


    Dim objtmpfolder As Object
    


    'Get the folder object
    Set objFolder = objFSO.GetFolder(CStr(s))
    
    
    

    For Each objFile In objFolder.Files
        '    'print file name
        ' to jest do stestowania na dzien nastepny - zoabczymy czy ten read only dziala tak jak nalezy.
        Set q = Nothing
        'Set q = Workbooks.Open(CStr(objFile.Name), , True)
        'fileCollection.Add q
        
        ' Debug.Print CStr(objFile.Path)
        ' ten warunek w przyszlosci ma zostac zmieniony
        If CStr(objFile.Name) Like CStr(pattern) Then
            ' Debug.Print CStr(objFile.Path)
            
            jest_juz_taki_i_go_nie_dodamy = False
            
            If fileCollection.Count > 0 Then
                For Each ic In fileCollection
                    If CStr(ic) = CStr(objFile.Path) Then
                        jest_juz_taki_i_go_nie_dodamy = True
                    End If
                Next ic
            End If
            
            If jest_juz_taki_i_go_nie_dodamy = False Then
                fileCollection.Add CStr(objFile.Path)
            End If
            
        End If
        '    'print file path
        '    Cells(i + 1, 2) = objFile.Path
        '    i = i + 1
    Next objFile
    
    If objFolder.SubFolders.Count > 0 Then

        For Each objtmpfolder In objFolder.SubFolders
            rekurencyjne_przegladanie_katalogow CStr(s) & Application.PathSeparator & CStr(objtmpfolder.shortName), pattern
        Next objtmpfolder
    End If

End Sub

Public Sub wyczysc_arkusz_rep()

    unhide_all_rows_and_all_columns rep
    
    rep.Range("A3:ZZ" & XWiz.DWA_DO_16).Interior.Color = RGB(255, 255, 255)
    rep.Range("A3:ZZ" & XWiz.DWA_DO_16).Clear
End Sub

Public Sub wyczysc_arkusz_rep_fup()
    
    unhide_all_rows_and_all_columns rep_fup
    
    rep_fup.Range("A3:ZZ" & XWiz.DWA_DO_16).Interior.Color = RGB(255, 255, 255)
    rep_fup.Range("A3:ZZ" & XWiz.DWA_DO_16).Clear
End Sub

Public Sub wyczysc_arkusz_rep_all()
    
    unhide_all_rows_and_all_columns rep_all
    
    With rep_all
        .Range("A2:ZZ" & XWiz.DWA_DO_16).Interior.Color = RGB(255, 255, 255)
        .Range("A2:ZZ" & XWiz.DWA_DO_16).Clear
    End With
End Sub


Public Sub wyczysc_arkusze_rep()

    wyczysc_arkusz_rep
    wyczysc_arkusz_rep_fup
    ' wyczysc_arkusz_rep_all
End Sub

Public Sub wyczysc_czesciowo_arkusz_rep()

    unhide_all_rows_and_all_columns rep
    
    ' to zajmuje calosciowo czego bardzo nie chce
    ' mam usunac tylko linijke laczace sie z side'em
    'rep.Range("P3:Z" & XWiz.DWA_DO_16).Interior.Color = RGB(255, 255, 255)
    'rep.Range("P3:Z" & XWiz.DWA_DO_16).Clear
    
    
    Dim r As Range, s As String
    Set r = rep.Range("A2")
    
    Do
        Set r = r.Offset(1, 0)
        
        If Trim(r) = "" Then
            Set linijka_do_zaktualizowania = Nothing
            Exit Sub
        End If
        
        's = "." & remove_special_cases(CStr(r)) & _
        '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
        '    "*" & remove_special_cases(CStr(r.Offset(0, 2))) & _
        '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"
        
        uid = ""
        On Error Resume Next
        uid = Trim(r.Offset(0, XWiz.E_ACTIVE - 1).Comment.Text)
        
        uid_z_side = ""
        uid_z_side = Trim(CStr(ktory_side_ma_zostac_zaktualizowany.Range("C1")))
        
    Loop Until uid_z_side = uid
    
    With rep.Range("P" & r.Row & ":AZ" & r.Row)
        .Interior.Color = RGB(255, 255, 255)
        .Clear
    End With
    
    Set linijka_do_zaktualizowania = r
End Sub


Public Sub wyczysc_czesciowo_arkusz_rep_fup()

    unhide_all_rows_and_all_columns rep_fup
    
    ' to zajmuje calosciowo czego bardzo nie chce
    ' mam usunac tylko linijke laczace sie z side'em
    'rep.Range("P3:Z" & XWiz.DWA_DO_16).Interior.Color = RGB(255, 255, 255)
    'rep.Range("P3:Z" & XWiz.DWA_DO_16).Clear
    
    
    Dim r As Range, s As String
    Set r = rep.Range("A2")
    
    Do
        Set r = r.Offset(1, 0)
        
        If Trim(r) = "" Then
            Set linijka_do_zaktualizowania = Nothing
            Exit Sub
        End If
        
        's = "." & remove_special_cases(CStr(r)) & _
        '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
        '    "*" & remove_special_cases(CStr(r.Offset(0, 2))) & _
        '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"
        
        
        uid = ""
        On Error Resume Next
        uid = Trim(r.Offset(0, XWiz.E_ACTIVE - 1).Comment.Text)
        
        uid_z_side = ""
        uid_z_side = Trim(CStr(ktory_side_ma_zostac_zaktualizowany.Range("C1")))
        
    Loop Until uid_z_side = uid
    
    
    
    With rep_fup.Range("P" & r.Row & ":AZ" & r.Row)
        .Interior.Color = RGB(255, 255, 255)
        .Clear
    End With
    
    Set linijka_do_zaktualizowania = r
End Sub







Public Sub TEST_wez_pierwszy_element_kolekcji_i_pobierz_duzy_zasieg(ByRef Sh As StatusHandler)
    
    Me.test_dla_ExecuteExcel4Macro CStr(fileCollection.Item(1)), rep.Range("d1")
End Sub

Public Sub wyczysc_cfg_sheet_i_jej_tmp_list_na_name_i_phase()

    With ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME)
    
        .Range("T3:U64000").Clear
    End With
    
End Sub

Public Sub lecimy_tutaj_jawne_otwieranie_pliku(s As String, ByRef r As Range, ByRef Sh As StatusHandler, ByRef kol As Collection)

    ' work template
    'strPath = "C:\Test\"
    'strFile = "Book1.xls"
    'strSheet = "Sheet1"
    'strRng = Range("A1").Address(1, 1, xlR1C1)
    '
    'strRef = "'" & strPath & "[" & strFile & "]" & strSheet & "'!" & strRng
    '
    'Result = ExecuteExcel4Macro(strRef)
    
    ' 1. zgodnie z zalozeniami zrodla powyzej
    ' w pierwszej kolejnosci musimy rozdzielic nazwe od pathu
    Dim sampath As String, filename As String
    Dim filename2 As String
    
    sampath = ""
    filename = ""
    filename2 = ""
    
    
    
    ' DLA TESTU CALA SCIEZKA
    ' ID 123456789
    If G_TEST_NA_DYSKU_LOKALNYM Then
        s = Replace(s, XWiz.XWIZ_PATH_FOR_SEARCHING, "")
    End If
    
    
    
    arr = Split(s, Application.PathSeparator) ' tutaj sprytnie uz
    
    Dim wrkbk As Workbook
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Set wrkbk = Workbooks.Open(CStr(s), False, True)
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Do
        
        DoEvents
    Loop Until Not wrkbk Is Nothing
    
    
    ' new one 2015-11-05
    ' ===================================================================
    ' ===================================================================
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Activate
    Set fh.r = ActiveSheet.Range("A1")
    
    fh.ustaw_swiezy_filtr
    Unhide_All_Columns ActiveSheet
    Unhide_All_Rows ActiveSheet
    ' ===================================================================
    ' ===================================================================
    
    Set idfw = New ExternalReadOnlyWorkbook
    
    'wrkbk.Parent.Visible = False
    
    
    idfw.init wrkbk, e_fup, e_run_type
    idfw.runMainLogic r, sampath, arr, Sh, kol
    
    ' MsgBox kol.Count

    
    'wrkbk.Parent.Visible = True
    wrkbk.Close False
    
End Sub





' niestety jest to slepa uliczka - zwraca zero!
Public Sub test_dla_ExecuteExcel4Macro(s As String, ByRef r As Range)
    Dim sampath As String, filename As String
    sampath = ""
    filename = ""
    arr = Split(s, "\") ' tutaj sprytnie uz
    
    For x = LBound(arr) To UBound(arr) - 1
        sampath = sampath & arr(x) & "\"
    Next x
    
    filename = arr(UBound(arr))
    
    
    strRng = Range("A1:Z10000").Address(, , xlR1C1)
    strRef = "'" & sampath & "[" & filename & "]" & XWiz.MASTER_SHEET_NAME & "'!" & strRng
    Result = ExecuteExcel4Macro(strRef)
    
    r = Result
End Sub



Public Sub zRepDoRepFup(rep As Worksheet, rep_fup As Worksheet)


    ' nieco inne wejscie na sub'a
    ' dodatkowe dwa arguemnty w sumie mozna by dac jako stale
    ' ale niech bedzie nieco urozmaicenia w kodzie
    
    ' tutaj globalnie przerzucamy cale dane z perspektywy jednego fupa
    ' w sumie implementacja bedzie bardzo podobna do extract data from rep
    ' nawet sie zastanawiam nad ukompononcecniem go
    
    Dim r As Range
    
    ' skopiuj common
    ' Workbooks("Name.xlsx").Worksheets(1).Range(Cells(6, 8), Cells(6, 9)).Copy _
        ' Destination:=ThisWorkbook.Worksheets(2).Range(Cells(6, 8), Cells(6, 9))
    If Trim(rep.Range("a3")) = "" Then
        ost_wiersz = 3
        MsgBox "nie ma czego kopiowac!"
        Exit Sub
    Else
        ost_wiersz = rep.Range("a2").End(xlDown).Row
    End If
    
    rep.Range(rep.Cells(3, 1), rep.Cells(ost_wiersz, 15)).Copy _
        Destination:=rep_fup.Range(rep_fup.Cells(3, 1), rep_fup.Cells(ost_wiersz, 15))
        
    Set mySh = New StatusHandler
    mySh.init_statusbar ost_wiersz - 2
    mySh.show
    
    
    For x = 3 To Int(ost_wiersz)
        ' lecimy przez calego repa i kopiujemy bazujac na nim z sideowych arkuszy
        Set r = rep_fup.Range("A" & CStr(x)).Offset(0, XWiz.CAPACITY_CHECK + 2)
        
        podlicz_za_pomoca_dodatkowych_arkuszy_obok_repa_co_lezy_w_odpowiedzialnosci_fupa r, rep, rep_fup, True
        ' selektywne_przekopiowaywanie_danych_z_komorek_z_komentarzem r, rep, rep_fup
        Set r = r.Offset(1, 0)
        
        mySh.progress_increase
    Next x
    
    mySh.hide
    Set mySh = Nothing
End Sub

' wyjatkowa procedura pracujaca totalnie niezaleznie od powyzszych
Public Sub extractDataFromRep()



    Dim rep_fup As Worksheet, rep As Worksheet
    Set rep_fup = ThisWorkbook.Sheets(XWiz.REP_FUP_SHEET_NAME)
    Set rep = ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME)
    
    ' jako jedyny zewnetrzny range
    ' linijka_do_zaktualizowania
    
    
    Dim r As Range, ldz As Range
    Set ldz = linijka_do_zaktualizowania
    
    
    If Not ldz Is Nothing Then
    
    
        If e_fup = E_FUP_FILTER_YES Then
            ' Result = wrkbk.Sheets(XWiz.MASTER_SHEET_NAME).Cells(x, XWiz.fup_code)
            ' -----------------------------------------------------------------------------
            
            
            ' ten assign to nic innego jak przypisanie pierwszego pustego pola fma resp
            Set r = rep_fup.Range("A" & CStr(ldz.Row)).Offset(0, XWiz.CAPACITY_CHECK + 2)
            ' Set r = rep_fup.Range("A3").Offset(0, XWiz.CAPACITY_CHECK + 2)
            
            
            
            Set mySh = New StatusHandler
            If Trim(r.Offset(1, 0)) <> "" Then
                mySh.init_statusbar Int(r.End(xlDown).Row - r.Row)
            Else
                mySh.init_statusbar Int(1)
            End If
            mySh.show
            podlicz_za_pomoca_dodatkowych_arkuszy_obok_repa_co_lezy_w_odpowiedzialnosci_fupa r, rep, rep_fup
            
            mySh.progress_increase
            
            mySh.hide
            Set mySh = Nothing
            
            ' -----------------------------------------------------------------------------
            
            
        ElseIf e_fup = E_FUP_FILTER_NO Then
            ' Result = XWiz.XWIZ_FLAGA
            ' ten assign to nic innego jak przypisanie pierwszego pustego pola fma resp
            Set r = rep.Range("A" & CStr(ldz.Row)).Offset(0, XWiz.CAPACITY_CHECK + 2)
            
            Set mySh = New StatusHandler
            mySh.init_statusbar Int(1)
            mySh.show
            
            
            podlicz_za_pomoca_dodatkowych_arkuszy_obok_repa_co_lezy_w_odpowiedzialnosci_fupa r, rep, rep_fup
            
            
            mySh.hide
            Set mySh = Nothing
        End If
    
    End If
    
End Sub

Private Function znajdz_arkusz_sideowy_dla_danej_iteracji(ByRef r As Range, _
        rep As Worksheet, _
            rep_fup As Worksheet) As Worksheet
            
            
            Set znajdz_arkusz_sideowy_dla_danej_iteracji = Nothing
            
            's = "." & remove_special_cases(CStr(r)) & _
            '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
            '    "*" & remove_special_cases(CStr(r.Offset(0, 2))) & _
            '    "*" & Left(remove_special_cases(CStr(r.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"
            
            uid = ""
            On Error Resume Next
            uid = Trim(r.Offset(0, XWiz.E_ACTIVE - 1))
            
            If uid <> "" Then
            
                Dim csh As Worksheet
                For Each csh In ThisWorkbook.Sheets
                    
                    
                    ' ten warunek pomija MY poniewaz jest posrany :)
                    ' dodatkowo wszystko ostro okraszone gwiazdkami gdyby napatoczyl nonstd
                    ' wiec powinno hulnac na dopasowaniu nazwy PLT PROJ BIW czy GA oraz Faza
                    ' w sumie xwiz.max len dla nazwy arkusza to zgodnie z 10 lutym 2016 28 znakow
                    
                    
                    
                    If CStr(uid) = Trim(CStr(csh.Range("C1"))) Then
                            
                            Set znajdz_arkusz_sideowy_dla_danej_iteracji = csh
                            Exit Function
                        
                    End If
                Next csh
            End If

End Function

Private Sub _
    podlicz_za_pomoca_dodatkowych_arkuszy_obok_repa_co_lezy_w_odpowiedzialnosci_fupa(ByRef r As Range, _
        rep As Worksheet, _
            rep_fup As Worksheet, _
            Optional z_rep_do_rep As Boolean)
            
            Dim wrk_sh As Worksheet
            Set wrk_sh = Nothing
            If e_fup = E_FUP_FILTER_YES Then
                ' ===========================

                If z_rep_do_rep Then
                    Set wrk_sh = znajdz_arkusz_sideowy_dla_danej_iteracji(r.Offset(0, XWiz.plt - r.Column), rep, rep_fup)
                Else
                    Set wrk_sh = ktory_side_ma_zostac_zaktualizowany
                End If

                
                If Not wrk_sh Is Nothing Then
                    ' --------------------------------------------
                    ' zlapalem odpowiedni arkusz - podlicze teraz resp
                    ' zrob kolekcje full - w srodku tego suba od razu wyrzucamy dane z kolekcji
                    podlicz_resp_i_noki wrk_sh, r.Offset(0, -1)
                    wyrzuc_dane_na_podstawie_wczesniej_stworzonej_kolekcji coll, rh, r
                    
                    
                    
                    ' --------------------------------------------
                End If
                
                ' ===========================
            ElseIf e_fup = E_FUP_FILTER_NO Then
                ' no filter
                Set wrk_sh = ktory_side_ma_zostac_zaktualizowany
                
                If Not wrk_sh Is Nothing Then
                    ' --------------------------------------------
                    ' zlapalem odpowiedni arkusz - podlicze teraz resp
                    ' zrob kolekcje full - w srodku tego suba od razu wyrzucamy dane z kolekcji
                    podlicz_resp_i_noki wrk_sh, r.Offset(0, -1)
                    wyrzuc_dane_na_podstawie_wczesniej_stworzonej_kolekcji coll, rh, r
                    
                    
                    
                    ' --------------------------------------------
                End If
            End If

End Sub

Private Sub podlicz_resp_i_noki(wrk_sh As Worksheet, ro As Range)
    
    podlicz_resp = 0
    loop_to_fill_rh wrk_sh, podlicz_resp, UCase(ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("fup_code"))
    
    ro.Value = podlicz_resp
    
    
End Sub

Public Sub loop_to_fill_rh(ByRef wrk_sh As Worksheet, _
    ByRef podlicz_resp As Variant, fup_code As String, _
    Optional del_conf_string As String, Optional pn_string As String)
    
    Set coll = New Collection
    
    ' Debug.Print ro.Address & " " & ro.Parent.Name
    Dim rp As Range
    Set rp = wrk_sh.Cells(3, 1)
    podlicz_resp = 0
    Do
        If rp.Offset(0, XWiz.E_SIDE_TB_RESP - 1) Like XWiz.FMA_WITH_STARS Then
        
            If (UCase(rp.Offset(0, XWiz.E_CMNT_FUP - 1)) = CStr(fup_code)) Or _
                e_fup = E_FUP_FILTER_NO Or _
                ((fup_code = XWiz.STR_BLANK) And (Trim(rp.Offset(0, XWiz.E_CMNT_FUP - 1)) = "")) Then
                
                
                    If (del_conf_string = "") Or _
                        (del_conf_string = Trim(rp.Offset(0, XWiz.E_SIDE_TB_Delivery_confirmation - 1))) Or _
                        (del_conf_string = XWiz.STR_BLANK And Trim(rp.Offset(0, XWiz.E_SIDE_TB_Delivery_confirmation - 1)) = "") Then
                        
                        
                        If (pn_string = "") Or _
                            (pn_string = Trim(rp.Offset(0, XWiz.E_SIDE_TB_PN - 1))) Or _
                            (pn_string = XWiz.STR_BLANK And Trim(rp.Offset(0, XWiz.E_SIDE_TB_PN - 1)) = "") Then

                            podlicz_resp = podlicz_resp + 1
                            
                            Set rh = Nothing
                            Set rh = New RecordsHandler
                            ' nie moglem uzyc bezposrednio comments
                            ' poniewaz enum na komentarzu nie pokrywa sie na ostatnich 3 kolumnach - korzystam chwilowo poniewaz
                            ' zgadza sie do delivery confirmation
                            rh.pn = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_PN)
                            rh.pn_nm = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_PN_NM)
                            rh.duns = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_DUNS)
                            rh.supp_nm = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_SUPP_NM)
                            rh.resp = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_RESP)
                            rh.fup = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_FUP_CODE)
                            
                            rh.Comments = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_Comments)
                            rh.Delivery_confirmation = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_Delivery_confirmation)
                            ' w excelu ok nok - natomiast pole w klasie to boolean
                            ' rh.Delivery_confirmation_status = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_Delivery_Confirmation_STATUS)
                            
                            rh.MRD1_Ordered_Date = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_Ordered_Date)
                            
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_Delivery_Confirmation_STATUS) = XWiz.G_OK Then
                                rh.Delivery_confirmation_status = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_Delivery_Confirmation_STATUS) = XWiz.G_NOK Then
                                rh.Delivery_confirmation_status = False
                            End If
                            
                            
                            'rh.duns = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_DUNS)
                            'rh.fup = wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_FUP_CODE)
                            
                            
                            ' MRD1 Status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_STATUS) = XWiz.G_OK Then
                                rh.MRD1_Ordered_STATUS = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_STATUS) = XWiz.G_NOK Then
                                rh.MRD1_Ordered_STATUS = False
                            End If
                            
                            ' MRD1 qty conf Status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_CONF_QTY) = XWiz.G_OK Then
                                rh.MRD1_confirmed_qty_Status = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_CONF_QTY) = XWiz.G_NOK Then
                                rh.MRD1_confirmed_qty_Status = False
                            End If
                            
                            ' MRD1 PUS status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_PUS_STATUS) = XWiz.G_OK Then
                                rh.MRD1_PUS_STATUS = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD1_PUS_STATUS) = XWiz.G_NOK Then
                                rh.MRD1_PUS_STATUS = False
                            End If
                            
                            
                            ' MRD2 Status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_STATUS) = XWiz.G_OK Then
                                rh.MRD2_Ordered_STATUS = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_STATUS) = XWiz.G_NOK Then
                                rh.MRD2_Ordered_STATUS = False
                            End If
                            
                            ' MRD2 qty conf Status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_CONF_QTY) = XWiz.G_OK Then
                                rh.MRD2_confirmed_qty_Status = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_CONF_QTY) = XWiz.G_NOK Then
                                rh.MRD2_confirmed_qty_Status = False
                            End If
                            
                            ' MRD2 PUS status
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_PUS_STATUS) = XWiz.G_OK Then
                                rh.MRD2_PUS_STATUS = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_MRD2_PUS_STATUS) = XWiz.G_NOK Then
                                rh.MRD2_PUS_STATUS = False
                            End If
                            
                            
                            ' TOTAL PUS STATUS
                            ' --------------------------------------------------------------------------
                            If wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_TOTAL_PUS_STATUS) = XWiz.G_OK Then
                                rh.TOTAL_PUS_STATUS = True
                            ElseIf wrk_sh.Cells(rp.Row, XWiz.E_SIDE_TB_TOTAL_PUS_STATUS) = XWiz.G_NOK Then
                                rh.TOTAL_PUS_STATUS = False
                            End If
                            
                            coll.Add rh
                        
                        End If
                    End If
                    
            End If
        End If
        Set rp = rp.Offset(1, 0)
    Loop Until Trim(rp) = ""
End Sub

Private Sub wyrzuc_dane_na_podstawie_wczesniej_stworzonej_kolekcji(ByRef c As Collection, rh As RecordsHandler, rr As Range)

    
        ' teraz jak juz mamy kolekcje mozemy sobie wyplenic arkusz rep_fup
        ' rr - mrd1 ordered status
        '
        '
        '
        ' przejdzmy teraz przez kolekcje i uzupelnijmy reszte arkusza
        '
        ' ===================================================================================
        
        For x = 0 To 7
            rr.Offset(0, x) = 0
            On Error Resume Next
            rr.Offset(0, x).addComment ""
        Next
        
        
        For Each rh In c
            If rh.MRD1_Ordered_STATUS = False Then
                rr = rr + 1
                jeden_item_komentarza rr, rh
            End If
            
            If rh.MRD1_confirmed_qty_Status = False Then
                rr.Offset(0, 1) = rr.Offset(0, 1) + 1
                jeden_item_komentarza rr.Offset(0, 1), rh
            End If
            
            If rh.MRD1_PUS_STATUS = False Then
                rr.Offset(0, 2) = rr.Offset(0, 2) + 1
                jeden_item_komentarza rr.Offset(0, 2), rh
            End If
            
            If rh.MRD2_Ordered_STATUS = False Then
                rr.Offset(0, 3) = rr.Offset(0, 3) + 1
                jeden_item_komentarza rr.Offset(0, 3), rh
            End If
            
            
            If rh.MRD2_confirmed_qty_Status = False Then
                rr.Offset(0, 4) = rr.Offset(0, 4) + 1
                jeden_item_komentarza rr.Offset(0, 4), rh
            End If
            
            
            If rh.MRD2_PUS_STATUS = False Then
                rr.Offset(0, 5) = rr.Offset(0, 5) + 1
                jeden_item_komentarza rr.Offset(0, 5), rh
            End If
            
            If rh.TOTAL_PUS_STATUS = False Then
                rr.Offset(0, 6) = rr.Offset(0, 6) + 1
                jeden_item_komentarza rr.Offset(0, 6), rh
            End If
            
            If rh.Delivery_confirmation_status = False Then
                rr.Offset(0, 7) = rr.Offset(0, 7) + 1
                jeden_item_komentarza rr.Offset(0, 7), rh
            End If
            
        Next rh
        
        
        ' ===================================================================================
End Sub



' zaklada ze koment juz istnieje
Public Sub jeden_item_komentarza(ByRef r As Range, ByRef rh As RecordsHandler)


    If Not r.Comment Is Nothing Then
    
    
        rh_pn = dopelnij_spacjami(rh.pn, XWiz.G_PN_LEN)
        rh_pn_nm = dopelnij_spacjami(rh.pn_nm, XWiz.G_PN_NM_LEN)
        rh_duns = dopelnij_spacjami(Trim(rh.duns), XWiz.G_DUNS_LEN)
        rh_supp_nm = dopelnij_spacjami(rh.supp_nm, XWiz.G_SUPP_NM_LEN)
        rh_resp = dopelnij_spacjami(rh.resp, XWiz.G_RESP_LEN)
        rh_fup = dopelnij_spacjami(rh.fup, XWiz.G_FUPCODE_LEN)
        rh_del_conf = dopelnij_spacjami(rh.Delivery_confirmation, XWiz.G_DEL_CONF_LEN)
        rh_ordered_date = dopelnij_spacjami(rh.MRD1_Ordered_Date, XWiz.G_DATES_CW_LEN)
        
        
    
        fcx = ""
        fcx = fcx & _
            XWiz.XWIZ_TXT_CMNT_HASH & " | " & XWiz.XWIZ_TXT_CMNT_ROW & dopelnij_spacjami(CStr(x), 5) & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_PREFIX & rh_pn & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_PN_NM_PREFIX & rh_pn_nm & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DUNS_PREFIX & rh_duns & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_SUPP_NM_PREFIX & rh_supp_nm & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_RESP_PREFIX & rh_resp & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_FUP_PREFIX & rh_fup & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_DEL_CONF_PREFIX & rh_del_conf & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX & rh_ordered_date & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_CMNTS_PREFIX & rh.Comments & " | "
        fcx = fcx & XWiz.XWIZ_TXT_CMNT_LINIA & Chr(10)
        
        r.Comment.Text r.Comment.Text & fcx
        
        With r.Comment.Shape.TextFrame
            
            .Characters.Font.Name = "Consolas"
            .Characters.Font.Size = 10
            .AutoSize = True
        End With
    End If
    

End Sub

Private Sub pracuj_na_tych_danych(ByRef r As Range, rep As Worksheet, rep_fup As Worksheet, e As E_REP_STATUS_COLUMNS)


    Dim rrepfup As Range
    Set rrepfup = rep_fup.Cells(r.Row, r.Column)

    If Not r.Comment Is Nothing Then
        
        If Trim(r.Comment.Text) = "" Then
            ' komentarz jest pusty zatem nie ma czego sprawdzac i wklejamy to samo czyli
            ' zero plus pusty komentarz
            rrepfup.addComment ""
            rrepfup = "0"
            
            
            
        Else
        End If
    Else
        MsgBox "nie ma komentarza dla komorki statusu - nie moze tak byc!"
    End If

End Sub
