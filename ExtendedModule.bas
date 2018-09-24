Attribute VB_Name = "ExtendedModule"
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


Public Sub complete_data_from_all_sides()

    prepareLabelsForExtendedSheet
    loopThroughSideSheets
    zmienFormatowanieDlaWybranychKolumn
    
    
    usunSidy
    
    MsgBox "merge ready!"
End Sub

Public Sub innerRunAfterClickOnBedzieNieBedzieForm(ByRef p_wh As WrkHandler, ByRef c As Collection)
    
    Set c = Nothing
    Set c = New Collection
    

    For x = 0 To BedzieNieBedzieForm.ListBoxRep.ListCount - 1
        c.Add CStr(BedzieNieBedzieForm.TextBoxWorkPath.Value) & CStr(BedzieNieBedzieForm.ListBoxRep.List(x))
    Next x

    p_wh.refreshCollection c
    
    If p_wh.countCollection > 0 Then
            
        Dim sh As StatusHandler
        Set sh = New StatusHandler
        sh.init_statusbar p_wh.countCollection

        Application.StatusBar = "uruchamiono glowna logike"
        
        sh.show
        p_wh.przejdz_po_kolei_przez_kolekcje_nazw_i_pobierz_dane sh
        
        
        
         'WIELKI FINAL!
         ' ==============================================================
         ' ==============================================================
         ' ==============================================================
        complete_data_from_all_sides
        ' ==============================================================
        ' ==============================================================
        ' ==============================================================
        
        
        
        sh.hide
        MsgBox "Gotowe! " & CStr(Now)
    Else
        
        MsgBox "kolekcja byla pusta!"
    End If
End Sub

Public Sub innerRunExtended()
    
    Dim wh As WrkHandler
    Set wh = New WrkHandler
    
    With wh
        
        

        CzekajForm.show vbModeless
        Application.DisplayStatusBar = True

        Application.StatusBar = "odnajduje pliki typu Wizard, ktore sa zgodne ze wzorcem..."


        ' tutaj szybka konfiguracja bez ponownego sprawdzania
        .setEFup E_FUP_FILTER_NO
        .setRunType NEW_RUN_EXTENDED
        .stworz_kolekcje XWIZ.XWIZ_PATH_FOR_SEARCHING
        
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
End Sub



' ==============================================================
Public Sub run_extended(ictrl As IRibbonControl)
    inner_start E_DYNAMIC, E_FUP_FILTER_NO, NEW_RUN_EXTENDED
End Sub
' ==============================================================



Private Function sprawdzPoprawnoscArkuszaExtended() As Boolean
    sprawdzPoprawnoscArkuszaExtended = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(XWIZ.EXTENDED_SHEET_NAME)
End Function







Private Sub prepareLabelsForExtendedSheet()


    Dim cfgSh As Worksheet
    Set cfgSh = ThisWorkbook.Sheets("config")
    
    Dim extSh As Worksheet
    Set extSh = ThisWorkbook.Sheets(XWIZ.EXTENDED_SHEET_NAME)
    
    
    Dim extR As Range
    Set extR = extSh.Range("A1")
    
    Dim rr As Range
    Set rr = cfgSh.Range("extendedStart")
    Do
        If UCase(rr.Offset(0, 1).Value) <> "" Then
            extR.Value = rr.Value
            Set extR = extR.Offset(0, 1)
        End If
    
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
End Sub


Private Sub loopThroughSideSheets()
    
    Dim wrksh As Worksheet
    For Each wrksh In ThisWorkbook.Sheets
    
    
        If wrksh.Cells(2, 3).Value = XWIZ.G_SIDE_SIGNATURE Then
            
            ' start copy data
            ' -------------------------------------------------------------------------
            
            copyThisSideSheetIntoExtended wrksh
            
            ' -------------------------------------------------------------------------
        End If
    Next wrksh
End Sub


Private Sub copyThisSideSheetIntoExtended(wrksh As Worksheet)


    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim extSh As Worksheet, labelExtRange As Range
    Set extSh = ThisWorkbook.Sheets(XWIZ.EXTENDED_SHEET_NAME)
    Set labelExtRange = extSh.Range("A1")
    
    
    
    
    Dim copySource As Range, copyDestination As Range, col As Long
    
    col = 1
    firstEmptyRow = getFirstEmptyRow(extSh)
    
    Set copySource = findLabelAndPrepareRangeReadyToCopy(labelExtRange, wrksh, CLng(col))
    ' this template not work for data from details
    ' Set copyDestination = extSh.Cells(firstEmptyRow, col)
    Set copyDestination = extSh.Range(extSh.Cells(firstEmptyRow, col), extSh.Cells(firstEmptyRow + CLng(wrksh.Cells(2, 1).Value) - 2, 19))
    
    copySource.Copy
    copyDestination.PasteSpecial xlPasteValues
    
    col = 20
    Set labelExtRange = labelExtRange.Offset(0, 19)
    Do
    

        Set copySource = findLabelAndPrepareRangeReadyToCopy(labelExtRange, wrksh, CLng(col))
        Set copyDestination = extSh.Cells(firstEmptyRow, CLng(col))
        
        If Not copySource Is Nothing Then
            copySource.Copy
            copyDestination.PasteSpecial xlPasteValues
        End If
        
        col = col + 1
        Set labelExtRange = labelExtRange.Offset(0, 1)
        
    Loop Until Trim(labelExtRange.Value) = ""
        
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub


Private Function findLabelAndPrepareRangeReadyToCopy(labelExtRange, wrksh As Worksheet, col As Long) As Range
    Set findLabelAndPrepareRangeReadyToCopy = Nothing
    
    
    
    If col >= 1 And col <= 19 Then
        ' info from details
        ' -------------------------------------------------
        
        Dim pasek As Range
        Set pasek = wrksh.Range("A1:S1")
        Set findLabelAndPrepareRangeReadyToCopy = pasek
        
        ' -------------------------------------------------
    Else
        ' rest of the world with std and non-std
        Dim srcLbl As Range, srcTmp As Range
        Set srcTmp = Nothing
        Set srcLbl = wrksh.Cells(3, 1)
        
        Do
            If srcLbl.Value = labelExtRange Then
                Set srcTmp = wrksh.Range(srcLbl.Offset(1, 0), wrksh.Cells(CLng(wrksh.Cells(2, 1).Value) + 4, srcLbl.Column))
                Set findLabelAndPrepareRangeReadyToCopy = srcTmp
                Exit Do
                
            End If
        
            Set srcLbl = srcLbl.Offset(0, 1)
        Loop Until Trim(srcLbl) = ""
        
    End If
    
    
    
End Function


Private Function getFirstEmptyRow(wrksh As Worksheet) As Long
    getFirstEmptyRow = CLng(2)
    
    Dim rr As Range
    Set rr = wrksh.Cells(1, 1)
    Do
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
    
    getFirstEmptyRow = CLng(rr.Row)
End Function




Private Sub zmienFormatowanieDlaWybranychKolumn()
    
    Dim ext As Worksheet, er As Range, cfgSh As Worksheet, cfgR As Range
    Set ext = ThisWorkbook.Sheets(XWIZ.EXTENDED_SHEET_NAME)
    Set cfgSh = ThisWorkbook.Sheets("config")
    Set cfgR = cfgSh.Range("extendedStart")
    Set er = ext.Range("A1")
    
    
    Do
        Set cfgR = cfgSh.Range("extendedStart")
        Do
            If cfgR = er Then
                If cfgR.Offset(0, 3) = "date" Then
                    er.EntireColumn.NumberFormat = "yyyy-mm-dd"
                    Exit Do
                End If
            End If
            Set cfgR = cfgR.Offset(1, 0)
        Loop Until Trim(cfgR) = ""
        Set er = er.Offset(0, 1)
    Loop Until Trim(er) = ""
    
End Sub




Public Sub usunSidy()
    
    Application.DisplayAlerts = False
    
    Dim sidesh As Worksheet
    For Each sidesh In ThisWorkbook.Sheets
        If sidesh.Cells(2, 3).Value = "extendedSide" Then
            sidesh.Delete
        End If
    Next sidesh
    
    Application.DisplayAlerts = True
End Sub
