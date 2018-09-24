Attribute VB_Name = "DblClickOnRepModule"
Public Sub innerDblClickGoToFromRepAndRepFup(ByRef Target As Range)


    Dim wh As WrkHandler
    If Not Target.Comment Is Nothing Then
    
        If Target.Column > XWIZ.CAPACITY_CHECK Then
        
            Dim dcrh As DoubleClickReportHandler
            Set dcrh = New DoubleClickReportHandler
        
            With Target.Comment
            
                dcrh.nazwa_kolumny Target.Parent.Cells(2, Target.Column)
                dcrh.projekt Target.Parent.Cells(Target.Row, 2)
                dcrh.lokalizacja_pliku Target.Parent.Cells(Target.Row, 1).Comment.Text
            
                raw_txt = .Text
                raw_txt = dcrh.remove_all_prefixes(raw_txt)
                
                dcrh.create_array_and_put_in_it_data_from_ raw_txt
            End With
        
        ' chcemy przejsc do dysku x i tego konkretnego pliku
        ElseIf Target.Column = 1 Then
        
            link = Target.Comment.Text
            
            Workbooks.Open filename:=link, ReadOnly:=False
        End If
    Else
    
        ' to znaczy ze klikamy w druga kolumne i chcemy przejsc do arkusza konkretnego
        If (Not Target.Offset(0, -1).Comment Is Nothing) And (Target.Column = XWIZ.PROJECT) Then
        
            inner_go_to_through_selection Target
        Else
        
            If (Not Target.Offset(0, -2).Comment Is Nothing) And Target.Column = XWIZ.BIW_GA Then
                pth = makePath(Target.Parent.Cells(Target.Row, 1).Comment.Text)
                If pth <> "" Then open_project_folder CStr(pth)
            End If
            
            
            ' BUILD PLAN Opener
            ' ---------------------------------------
            If Target.Column = XWIZ.build_start Then
                
                
                Set wh = New WrkHandler
                pth = makePath(Target.Parent.Cells(Target.Row, 1).Comment.Text)
                
                If pth <> "" Then
                    CzekajForm.show vbModeless
                    wh.znajdz_build_plan pth
                    wh.otowrzWszystkieBuildPlany
                    CzekajForm.hide
                End If
                Set wh = Nothing
            End If
            
            If Target.Column = XWIZ.build_end Then
            
                
                Set wh = New WrkHandler
                pth = makePath(Target.Parent.Cells(Target.Row, 1).Comment.Text)
                
                If pth <> "" Then
                    CzekajForm.show vbModeless
                    wh.znajdz_build_plan pth
                    wh.otowrzWszystkieBuildPlany
                    CzekajForm.hide
                End If
                Set wh = Nothing
            End If
            ' ---------------------------------------
        End If
        
    End If
    
    Set dcrh = Nothing

End Sub


Private Function makePath(str)
    
    arr = Split(str, Application.PathSeparator)
    
    tmp = ""
    
    If LBound(arr) < UBound(arr) Then
        For x = LBound(arr) To UBound(arr) - 1
            
            If x < (UBound(arr) - 1) Then
                tmp = tmp & arr(x) & Application.PathSeparator
            Else
                tmp = tmp & arr(x)
            End If
            
        Next x
    End If
    
    
    makePath = CStr(tmp)
End Function

