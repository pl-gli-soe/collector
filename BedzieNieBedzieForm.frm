VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BedzieNieBedzieForm 
   Caption         =   "Co chcesz miec w raporcie (double click to move)"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17685
   OleObjectBlob   =   "BedzieNieBedzieForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BedzieNieBedzieForm"
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


Private p_wh As WrkHandler
Private c As Collection


Public Sub connect_with_wrk_handler(ByRef wh As WrkHandler)
    Set p_wh = wh
End Sub


Public Sub wypelnij_work_path()
    Me.TextBoxWorkPath.Value = CStr(XWIZ.XWIZ_PATH_FOR_SEARCHING)
End Sub

Public Sub wypelnij_listboxy()

    Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    
    With p_wh
    
        If .getCollection.Count > 0 Then
    
            For Each s In .getCollection
                
                
                tmp = Replace(s, Me.TextBoxWorkPath.Value, "")
                
                If tmp Like "*" & XWIZ.XWIZ_FLE_POSTFIX_VERSION & "*" Then
                    Me.ListBoxRep.AddItem tmp
                ElseIf tmp Like "*" & XWIZ.XWIZ_FLE_OLD_POSTFIX_VERSION & "*" Then
                    ' Me.ListBoxRep.AddItem "* " & tmp
                    Me.ListBoxSource.AddItem tmp
                End If
                
            Next s
            
            
            ' kolorwanie osobne nie bedzie dzialac poprawnie
            'For x = 0 To Me.ListBoxRep.ListCount - 1
            '    If Left(Me.ListBoxRep.List(x), 1) = "*" Then
            '        Me.ListBoxRep.List(x).ForeColor = RGB(255, 0, 0)
            '    End If
            'Next x
        End If
    End With
    
End Sub

Private Sub BtnAll_Click()
    
    Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            
            Me.ListBoxRep.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            
        Next s
    End With
End Sub

Private Sub BtnCopyToConfig_Click()
    If ListBoxRep.ListCount > 0 Then
    
    
        Dim c As Worksheet
        Set c = ThisWorkbook.Sheets(XWIZ.CONFIG_SHEET_NAME)
        
        Dim r As Range
        Set r = c.Range("B2:B256")
        r.Clear
    
        
    
        For x = 0 To Me.ListBoxRep.ListCount - 1
        
            init_path = Me.TextBoxWorkPath
            
            ' w srodku iteracji teraz lecimy i przerzucamy do arkusza config
            
            arr = Split(CStr(Me.ListBoxRep.List(x)), "\")
            
            For y = LBound(arr) To UBound(arr) - 1
                init_path = init_path & CStr(arr(y))
            Next y
            init_path = init_path & "\"
            ' Debug.Print init_path
            
            put_this_one_into_config c, r, init_path, x
        Next x
    End If
End Sub


Private Sub put_this_one_into_config(c As Worksheet, r As Range, ip, x)
    
    Set r = c.Range("B" & CStr(x + 2))
    r = CStr(ip)
End Sub

Private Sub BtnFiltruj_Click()
    Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            If Replace(s, Me.TextBoxWorkPath.Value, "") Like "*" & Me.TextBoxPattern.Value & "*" Then
                Me.ListBoxRep.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            Else
                Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            End If
        Next s
    End With
End Sub


Private Sub BtnFiltruj2Add_Click()
    
    ' Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            If Replace(s, Me.TextBoxWorkPath.Value, "") Like "*" & Me.TextBoxPattern2Add.Value & "*" Then
                Me.ListBoxRep.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            Else
                Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            End If
        Next s
    End With
    
End Sub

Private Sub BtnFiltruj3Usun_Click()
    
    'Me.ListBoxRep.Clear
    'Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            If Not (Replace(s, Me.TextBoxWorkPath.Value, "") Like "*" & Me.TextBoxPattern3Usun.Value & "*") Then
                ' Me.ListBoxRep.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
                For x = 0 To Me.ListBoxRep.ListCount
                    On Error Resume Next
                    If (Me.ListBoxRep.List(x) Like "*" & Me.TextBoxPattern3Usun.Value & "*") Then
                        Me.ListBoxRep.RemoveItem x
                        Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
                    End If
                Next x
                
            Else
                Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            End If
        Next s
    End With
End Sub

Private Sub BtnFiltruj4Zostaw_Click()

    ' Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            If (Replace(s, Me.TextBoxWorkPath.Value, "") Like "*" & Me.TextBoxPattern3Usun.Value & "*") Then
                ' Me.ListBoxRep.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
                For x = 0 To Me.ListBoxRep.ListCount
                    On Error Resume Next
                    If Not (Me.ListBoxRep.List(x) Like "*" & Me.TextBoxPattern3Usun.Value & "*") Then
                        Me.ListBoxRep.RemoveItem x
                        Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
                    End If
                Next x
            Else
                Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
            End If
        Next s
    End With
End Sub

Private Sub BtnNothing_Click()
    
    Me.ListBoxRep.Clear
    Me.ListBoxSource.Clear
    
    With p_wh
        For Each s In .getCollection
            
            Me.ListBoxSource.AddItem Replace(s, Me.TextBoxWorkPath.Value, "")
        Next s
        
    End With
End Sub

Private Sub BtnRun_Click()
    hide
    
    
    ' wyczysc arkusze
    With p_wh
    
        If .get_e_run_type = NEW_RUN_EXTENDED Then
            .wyczysc_arkusz_extended ' plus jest wyczyszczenie side arkuszy
        ElseIf .get_e_run_type = NEW_RUN_ALL Then
            .wyczysc_arkusz_rep_all
            .wyczysc_arkusz_pivot_source
        ElseIf .get_e_run_type < NEW_RUN_ALL Then
        
            If .get_e_fup = E_FUP_FILTER_NO Then
                .wyczysc_arkusz_rep
            ElseIf .get_e_fup = E_FUP_FILTER_YES Then
                .wyczysc_arkusz_rep_fup
            End If
        End If
        
        
        If .get_e_run_type = NEW_RUN_EXTENDED Then
            ' kolejny sub inner znajdujacy sie bezposrednio w module extended
            innerRunAfterClickOnBedzieNieBedzieForm p_wh, c
        Else
    
            Set c = Nothing
            Set c = New Collection
            
        
            For x = 0 To Me.ListBoxRep.ListCount - 1
                c.Add CStr(Me.TextBoxWorkPath.Value) & CStr(Me.ListBoxRep.List(x))
            Next x
    
            .refreshCollection c
        
        
        
            If .countCollection > 0 Then
            
                Dim sh As StatusHandler
                Set sh = New StatusHandler
                sh.init_statusbar .countCollection
                    
                
                ' Application.StatusBar = "czyszcze arkusz raportujacy"
                
                
                Application.StatusBar = "uruchamiono glowna logike"
                
                sh.show
                .przejdz_po_kolei_przez_kolekcje_nazw_i_pobierz_dane sh
                
                
                If .get_e_run_type < NEW_RUN_ALL Then
                
                    ' .wyczysc_arkusze_rep
                
                    Dim art As AddRedToNoks
                    Set art = New AddRedToNoks
                    With art
                        .prepare_range_and_colour_noks_red ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
                        .colour_blue_this_week_on_bom_pus_date_mrd_and_build ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
                        
                        
                        .prepare_range_and_colour_noks_red ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
                        .colour_blue_this_week_on_bom_pus_date_mrd_and_build ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
                        
                    End With
                    
                    
                    
                    
                    zmien_format_na CStr(XWIZ.REP_SHEET_NAME), "0"
                    zmien_format_na CStr(XWIZ.REP_FUP_SHEET_NAME), "0"
                    
                    
                    
                    .oddaj_cale_nazwy_dla_project_i_faz
                End If
                
                sh.hide
                MsgBox "Gotowe! " & CStr(Now)
            Else
                
                MsgBox "kolekcja byla pusta!"
            End If
        End If
    End With
End Sub

Private Sub ListBoxRep_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    With Me.ListBoxRep

        For x = 0 To .ListCount - 1
            If .Selected(x) Then
                With Me.ListBoxSource
                    .AddItem Me.ListBoxRep.List(x)
                End With
                .RemoveItem x
                Exit For
            End If
        Next x
    End With
    
    Me.Repaint

End Sub

Private Sub ListBoxSource_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.ListBoxSource

        For x = 0 To .ListCount - 1
            If .Selected(x) Then
                With Me.ListBoxRep
                    .AddItem Me.ListBoxSource.List(x)
                End With
                .RemoveItem x
                Exit For
            End If
        Next x
    End With
    
    Me.Repaint
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   p_wh.refreshCollection Nothing
   End
End Sub
