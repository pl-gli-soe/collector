Attribute VB_Name = "ExtendRepToRightModule"
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


' ExtendRepToRightModule
' by milewski

' ten modul zawiera glowna procedure uruchomieniowa
' zwiazana z stworzeniem dodatkowych kolumn w arkuszu rep
' pod komentarze dla supervisora oraz kazdego fupa z osobna



Public Sub rozlej_fupy_na_prawo(ictrl As IRibbonControl)
    inner_move_data_to_right_in_rep_sheet New FupsToRight
    MsgBox "ready!'"
    
End Sub

Public Sub rozlej_del_confy_na_prawo(ictrl As IRibbonControl)
    inner_move_data_to_right_in_rep_sheet New DelConfsToRight
    MsgBox "ready!'"
    
End Sub

Public Sub rozlej_pny_na_prawo(ictrl As IRibbonControl)
    inner_move_data_to_right_in_rep_sheet New PnsToRight
    MsgBox "ready!'"
    
End Sub



Private Sub inner_move_data_to_right_in_rep_sheet(ftrh As IToRight)
    
    Dim rep As Worksheet
    
    Set rep = ThisWorkbook.Sheets(XWiz.REP_SHEET_NAME)
    
    
    With ftrh
        ' .clear_field_for_data rep
        
        
        Dim r As Range
        Set r = rep.Range("A2").End(xlToRight).Offset(0, 1)
        r = XWiz.MGMT_CMNTS
        
        
        .goThroughSideSheetsAndFillDicWithFupsNames
        .putDataInLabelsAndSpreadOutValuesFromCollections
        .kolorujLabelki
    End With
    
    
    
    
    Set ftrh = Nothing
End Sub
