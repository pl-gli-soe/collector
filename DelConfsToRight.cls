VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DelConfsToRight"
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



Implements IToRight

Private md As MyDictionary
Private wh As WrkHandler
Private rep As Worksheet
Private Sh As Worksheet
Private r As Range
Private content_ref As Range
Private cr As Range
Private ost_wiersz As Long

Private del_conf_range As Range


Private del_confs_coll As Collection


Private Sub Class_Initialize()
    Set md = New MyDictionary
    Set wh = New WrkHandler
    Set rep = ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME)
    Set del_confs_coll = New Collection
End Sub

Private Sub Class_Terminate()
    Set md = Nothing
    Set wh = Nothing
    Set del_confs_coll = Nothing
End Sub

Private Sub IToRight_goThroughSideSheetsAndFillDicWithFupsNames()
    ' side'y latwo odnalezc poniewaz na poczatku swojej nazwy maja kropke
    ' jednak zeby miec wieksza pewnosc dodatkowo dam if na contentcie arkuszy
    
    Set del_confs_coll = New Collection
    
    For Each Sh In ThisWorkbook.Sheets
        If Left(Sh.Name, 1) = "." Then
            ' pierwsza opcja sprawdzona ze istnieje szansa ze to sideowy arkusz
            ' druga instancja - na duzo procent jesli to przejdzie to raczej mamy do
            'czynienia z arkuszem side'owym
            If Trim(Sh.Range("A2")) = "#" And Trim(Sh.Range("B2")) = "ROW" Then
            
                ' skoro tak to teraz po kolei dodajemy fupa jesli jest po co
                If Trim(Sh.Range("A3")) <> "" Then
                    ' ost_wiersz = sh.Range("A2").End(xlDown).Row
                    
                    Set r = Sh.Range("A2")
                    Do
                        Set r = r.Offset(1, 0)
                    Loop Until Trim(CStr(r)) = ""
                    ost_wiersz = r.Row
                    
                Else
                    ost_wiersz = 3
                End If
                
                For x = 3 To ost_wiersz
                    
                    proba_dodania_del_conf_narazie Sh.Cells(Int(x), Int(XWiz.E_SIDE_TB_Delivery_confirmation)), del_confs_coll
                Next x
            End If
        End If
    Next Sh
End Sub

Private Sub proba_dodania_del_conf_narazie(rr As Range, ByRef c As Collection)


    Set md = del_conf_exists(CStr(rr), c)
    
    Dim txt_for_nm As String
    If Len(Trim(CStr(rr))) > 0 Then
        txt_for_nm = Trim(CStr(rr))
    Else
        txt_for_nm = XWiz.STR_BLANK
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


Private Function del_conf_exists(s As String, ByRef c As Collection) As MyDictionary
    
    Set del_conf_exists = Nothing
    
    For Each md In c
        If Trim(s) = "" And UCase(CStr(Trim(md.nm))) = XWiz.STR_BLANK Then
        
            Set del_conf_exists = md
            Exit Function
            
        ElseIf (UCase(Trim(md.nm)) = UCase(Trim(s))) And Len(Trim(s)) > 0 Then
            Set del_conf_exists = md
            Exit Function
        End If
    Next md
End Function

Private Sub IToRight_kolorujLabelki()

    Set r = rep.Range("A2").End(xlToRight)
    
    Do
        With r
            .Font.Bold = True
            .Interior.Color = RGB(100, 200, 245)
            .Orientation = xlUpward
        End With
        Set r = r.Offset(0, -1)
    Loop Until (Trim(r) = XWiz.MGMT_CMNTS)
    
    r.Interior.Color = RGB(160, 180, 180)
End Sub

Private Sub IToRight_putDataInLabelsAndSpreadOutValuesFromCollections()
    
    Set r = rep.Range("A2").End(xlToRight).Offset(0, 1)
    
    If del_confs_coll.Count > 0 Then
        For Each md In del_confs_coll
        
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
        
            uid = ""
            uid_z_side = ""
            
            On Error Resume Next
            uid_z_side = Trim(CStr(vr.Parent.Range("C1")))
            On Error Resume Next
            uid = CStr(Trim(mtmp.Offset(0, XWiz.E_ACTIVE - 1).Comment.Text))
            
            If uid <> "" Then
            
                ' Debug.Assert Not (vpn Like "*RU")
                If vpn Like CStr(s) Then
                
                    'If CStr(vr.Parent.Cells(vr.Row, e_side).Value) = CStr(XWiz.G_NOK) Then
                        meh = 0
                        wh.setEFup E_FUP_FILTER_NO
                        wh.loop_to_fill_rh vr.Parent, meh, XWiz.STR_BLANK, CStr(cr.Parent.Cells(2, cr.Column))
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


