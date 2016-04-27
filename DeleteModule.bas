Attribute VB_Name = "DeleteModule"
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

Public Sub clear_all_old_data(ictrl As IRibbonControl)
    inner_clear_all_old_data
End Sub

Public Sub clear_data_for_pivot_logic(ictrl As IRibbonControl)
    clear_sheet_named_all
    clear_pivot_sheets
End Sub

Public Sub clear_all_sides_and_custom_sheets(ictrl As IRibbonControl)
    remove_side_sheets
End Sub

Public Sub clear_rep_only(ictrl As IRibbonControl)
    inner_clear_rep_type_sheet CStr(XWiz.REP_SHEET_NAME)
End Sub

Public Sub clear_reps(ictrl As IRibbonControl)
    inner_clear_rep_type_sheet CStr(XWiz.REP_SHEET_NAME)
    inner_clear_rep_type_sheet CStr(XWiz.REP_FUP_SHEET_NAME)
End Sub

Public Sub clear_rep_fup_only(ictrl As IRibbonControl)
    inner_clear_rep_type_sheet CStr(XWiz.REP_FUP_SHEET_NAME)
End Sub




Public Sub clear_pivot_sheets()


    answer = MsgBox("Czy jestes pewien akcji pozbycia sie arkuszy pivotowych!?", vbYesNo, "!")
    
    If answer = vbYes Then

        With ThisWorkbook.Sheets(XWiz.PIVOT_SOURCE_SHEET_NAME)
        
            On Error Resume Next
            .ShowAllData
            
            Set r = .Range("a1")
            Set r = .Range(r, r.Offset(100000, 1000))
            
            r.ClearComments
            r.Clear
        
        End With
        
        With ThisWorkbook.Sheets(XWiz.PN_PIVOT_SHEET_NAME)
            
            .ShowAllData
        
            Set r = .Range("a1")
            Set r = .Range(r, r.Offset(100000, 1000))
            
            r.ClearComments
            r.Clear
        
        End With
    
    Else
        MsgBox "arkusze pivotowe nie zostaly usuniete"
    End If
End Sub

Public Sub clear_sheet_named_all()


    
    answer = MsgBox("Czy jestes pewien akcji wyczyszczenia arkusza all!?", vbYesNo, "!")
    
    If answer = vbYes Then
    
    
        With ThisWorkbook.Sheets(XWiz.ALL_SHEET_NAME)
            
            On Error Resume Next
            .ShowAllData
        
            Set r = .Range("a2")
            Set r = .Range(r, r.Offset(100000, 1000))
            
            r.ClearComments
            r.Clear
        
        End With
    Else
        MsgBox "arkusz all nie zostal wyczyszczony"
    End If
End Sub




Private Sub remove_side_sheets()



    answer = MsgBox("Czy jestes pewien akcji usuniecia arkuszy sideowych!?", vbYesNo, "!")
    
    If answer = vbYes Then

        With ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME)
            For Each Sh In .Parent.Sheets
                If Sh.Name <> XWiz.REP_SHEET_NAME And _
                Sh.Name <> XWiz.CONFIG_SHEET_NAME And _
                Sh.Name <> XWiz.REP_FUP_SHEET_NAME And _
                Sh.Name <> XWiz.PIVOT_SHEET_NAME And _
                Sh.Name <> XWiz.PIVOT_SOURCE_SHEET_NAME And _
                Sh.Name <> XWiz.PN_PIVOT_SHEET_NAME And _
                Sh.Name <> XWiz.ALL_SHEET_NAME Then
                    Sh.Delete
                End If
            Next Sh
        End With
    Else
        MsgBox "arkusze sideowe nie zostana usuniete"
    End If
End Sub

Private Sub inner_clear_rep_type_sheet(worksheet_name As String)

    answer = MsgBox("Czy jestes pewien akcji czyszczenia arkusza " & CStr(worksheet_name) & "!?", vbYesNo, "!")
    
    If answer = vbYes Then
    
        Application.DisplayAlerts = False
    
        Dim r As Range
        
        With ThisWorkbook.Sheets(worksheet_name)
        
            On Error Resume Next
            .ShowAllData
        
            Set r = .Range("a3")
            Set r = .Range(r, r.Offset(100000, 1000))
            
            r.ClearComments
            r.Clear
            
            Set r = .Range("y2")
            Set r = .Range(r, r.Offset(0, 1000))
            r.ClearComments
            r.Clear
        End With
    
    
        Application.DisplayAlerts = True
    Else
        MsgBox "arkusza " & CStr(worksheet_name) & " nie zostanie wyczyszczony"
    End If
End Sub

Private Sub inner_clear_all_old_data()



    answer = MsgBox("Czy jestes pewien!?", vbYesNo, "!")
    
    If answer = vbYes Then
    
        Application.DisplayAlerts = False
        
        inner_clear_rep_type_sheet CStr(XWiz.REP_SHEET_NAME)
            
            
        remove_side_sheets
        
        
        inner_clear_rep_type_sheet CStr(XWiz.REP_FUP_SHEET_NAME)
        
        clear_sheet_named_all
        
        
        clear_pivot_sheets
        
        
        Application.DisplayAlerts = True
    Else
        MsgBox "nic nie zostanie usuniete!"
    End If
End Sub
