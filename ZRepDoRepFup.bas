Attribute VB_Name = "ZRepDoRepFup"
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



Public Sub z_rep_do_rep_fup(ictrl As IRibbonControl)
    inner_z_rep_do_rep_fup
    MsgBox "ready!"
End Sub


Private Sub inner_z_rep_do_rep_fup()



    WybierzFUPCode.show
    
    Dim rep As Worksheet
    Dim rep_fup As Worksheet
    Dim wh As WrkHandler
    Set wh = New WrkHandler
    
    
    
    
    Set rep = ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
    Set rep_fup = ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
    
    With wh
        .setEFup E_FUP_FILTER_YES
        .setSideWrksh Nothing
        .wyczysc_arkusz_rep_fup
        .ZRepDoRepFup rep, rep_fup
    End With
    
    Set wh = Nothing
    
    
    Dim art As AddRedToNoks
    Set art = New AddRedToNoks
    art.prepare_range_and_colour_noks_red rep_fup
    art.colour_blue_this_week_on_bom_pus_date_mrd_and_build rep_fup
End Sub
