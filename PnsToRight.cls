VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PnsToRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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



Implements IToRight

Private md As MyDictionary
Private wh As WrkHandler
Private rep As Worksheet
Private sh As Worksheet
Private r As Range
Private content_ref As Range
Private cr As Range
Private ost_wiersz As Long

Private pn_range As Range


Private pns As Collection


Private Sub Class_Initialize()
    Set md = New MyDictionary
    Set wh = New WrkHandler
    Set rep = ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
    Set pns = New Collection
End Sub

Private Sub Class_Terminate()
    Set md = Nothing
    Set wh = Nothing
    Set pns = Nothing
End Sub

Private Sub IToRight_goThroughSideSheetsAndFillDicWithFupsNames()
    ' side'y latwo odnalezc poniewaz na poczatku swojej nazwy maja kropke
    ' jednak zeby miec wieksza pewnosc dodatkowo dam if na contentcie arkuszy
    
    Set fups = New Collection
    
    For Each sh In ThisWorkbook.Sheets
        If Left(sh.Name, 1) = "." Then
            ' pierwsza opcja sprawdzona ze istnieje szansa ze to sideowy arkusz
            ' druga instancja - na duzo procent jesli to przejdzie to raczej mamy do
            'czynienia z arkuszem side'owym
            If Trim(sh.Range("A2")) = "#" And Trim(sh.Range("B2")) = "ROW" Then
            
                ' skoro tak to teraz po kolei dodajemy fupa jesli jest po co
                If Trim(sh.Range("A3")) <> "" Then
                    ' ost_wiersz = sh.Range("A2").End(xlDown).Row
                    
                    Set r = sh.Range("A2")
                    Do
                        Set r = r.Offset(1, 0)
                    Loop Until Trim(CStr(r)) = ""
                    ost_wiersz = r.Row
                    
                Else
                    ost_wiersz = 3
                End If
                
                For x = 3 To ost_wiersz
                    
                    proba_dodania_key_pn_narazie sh.Cells(Int(x), Int(XWIZ.E_SIDE_TB_PN)), pns
                Next x
            End If
        End If
    Next sh
End Sub

Private Sub proba_dodania_key_pn_narazie(rr As Range, ByRef c As Collection)


    Set md = pn_exists(CStr(rr), c)
    
    Dim txt_for_nm As String
    If Len(Trim(CStr(rr))) > 0 Then
        txt_for_nm = Trim(CStr(rr))
    Else
        txt_for_nm = XWIZ.STR_BLANK
    End If
    
    
    If md Is Nothing Then
        Set md = Nothing
        Set md = New MyDictionary
        md.nm = txt_for_nm
        ' Debug.Print md.coll.Count
        md.coll.Add rr
        c.Add md
        
    Else
        
        ' Debug.Print md.coll.Count
        md.coll.Add rr
        

    End If
End Sub


Private Function pn_exists(s As String, ByRef c As Collection) As MyDictionary
    
    Set pn_exists = Nothing
    
    For Each md In c
        If Trim(s) = "" And UCase(CStr(Trim(md.nm))) = XWIZ.STR_BLANK Then
        
            Set pn_exists = md
            Exit Function
            
        ElseIf (UCase(Trim(md.nm)) = UCase(Trim(s))) And Len(Trim(s)) > 0 Then
            Set pn_exists = md
            Exit Function
        End If
    Next md
End Function

Private Sub IToRight_kolorujLabelki()

    Set r = rep.Range("A2").End(xlToRight)
    
    Do
        With r
            .Font.Bold = True
            .Interior.Color = RGB(170, 215, 160)
            .Orientation = xlUpward
        End With
        Set r = r.Offset(0, -1)
    Loop Until (Trim(r) = XWIZ.MGMT_CMNTS)
    
    r.Interior.Color = RGB(160, 180, 180)
End Sub

Private Sub IToRight_putDataInLabelsAndSpreadOutValuesFromCollections()
    
    Set r = rep.Range("A2").End(xlToRight).Offset(0, 1)
    
    If pns.Count > 0 Then
        For Each md In pns
        
            ' tutaj mamy perspektywe kolumny
            r = CStr(md.nm)
            
            
            Set content_ref = r.Offset(1, 0)
            Set cr = content_ref
            
            ' teraz dobrze by bylo zajac sie kolejnymi projektami
            
            Dim tmp As Range
            Set tmp = rep.Cells(cr.Row, 1)
            Do
                
                wypelnij_odpowiednio_ta_komorke cr, tmp, md.coll, E_SIDE_TB_Delivery_Confirmation_STATUS
                
                Set cr = cr.Offset(1, 0)
                Set tmp = tmp.Offset(1, 0)
            Loop Until Trim(tmp) = ""
            
            Set r = r.Offset(0, 1)
        Next md
    End If
    
End Sub

Private Sub wypelnij_odpowiednio_ta_komorke(ByRef cr As Range, mtmp As Range, coll As Collection, e_side As E_COLLECTOR_SIDE_TABLE)
    
    Dim vr As Range, rh As RecordsHandler
    ' 1. mtmp to pierwsza kolumna zawierajaca plt code - offset zlapiemy tez nazwe projektu
    ' 2. vr to cell z side arkusza ktora jestem elementem z kolekcji coll
    ' 3. coll kolekcja jako wartosc dict (zbior celli sideowych)
    
    With cr
        .Clear
        .ClearComments
        .Value = 0
    
        For Each vr In coll
        
            
            
        
            's = "." & remove_special_cases(CStr(mtmp)) & _
            '"*" & Left(remove_special_cases(CStr(mtmp.Offset(0, 1))), XWiz.G_CUT_PROJECT) & _
            '"*" & remove_special_cases(CStr(mtmp.Offset(0, 2))) & _
            '"*" & Left(remove_special_cases(CStr(mtmp.Offset(0, 4))), XWiz.G_CUT_PHAZE) & "*"
            'vpn = CStr(vr.Parent.Name)
            
            uid = ""
            uid_z_side = ""
            
            On Error Resume Next
            uid_z_side = Trim(CStr(vr.Parent.Range("C1")))
            On Error Resume Next
            uid = CStr(Trim(mtmp.Offset(0, XWIZ.E_ACTIVE - 1).Comment.Text))
            
            If uid <> "" Then
            
                ' Debug.Assert Not (vpn Like "*RU")
                If uid = uid_z_side Then
                
                    'If CStr(vr.Parent.Cells(vr.Row, e_side).Value) = CStr(XWiz.G_NOK) Then
                        meh = 0
                        wh.setEFup E_FUP_FILTER_NO
                        wh.loop_to_fill_rh vr.Parent, meh, XWIZ.STR_BLANK, XWIZ.STR_BLANK, CStr(cr.Parent.Cells(2, cr.Column))
                        ' .Value = meh
                        .Value = 0
                        
                        If wh.getColl.Count > 0 Then
                            cr.addComment ""
                            For Each rh In wh.getColl
                                If Not rh.Delivery_confirmation_status Then
                                    wh.jeden_item_komentarza cr, rh
                                    .Value = .Value + 1
                                End If
                            Next rh
                            Exit For
                        End If
                        
                        
                    'End If
                End If
            End If
            
            
        Next vr
    End With
End Sub


