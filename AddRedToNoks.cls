VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddRedToNoks"
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



' ir.NumberFormat = "0_ ;[Red]-0 "
' gdzie ir as range

' ten modul jest odpowiedzialny za odpowiednie traktowanie danych - tj - czerwony font
' tam gdzie mamy liczby dodatnie


' jednak niech Cie nazwa nie zmyli - tutaj tez bawimy sie kolorem niebieskim

Private r As Range
Private rep As Worksheet
Private rep_fup As Worksheet


Public Function getR() As Range
    
    Set getR = r
End Function


Public Sub setR(ir As Range)
    Set r = ir
End Sub

Public Sub prepare_range_and_colour_noks_red(arkusz_raportu As Worksheet)
    
    Set r = arkusz_raportu.Range(XWIZ.ADRES_POCZATKU_NOKOW_W_REP)
    
    r.Offset(0, -1).NumberFormat = "0"
    r.Offset(0, -2).NumberFormat = "0"
    
    
    If Trim(r.Offset(1, 0)) <> "" Then
        Set r = arkusz_raportu.Range(r, r.End(xlDown).End(xlToRight))
    Else
        Set r = arkusz_raportu.Range(r, r.End(xlToRight))
    End If
    
    r.NumberFormat = "[Red];[Black];[Black];[Black]"
    
End Sub

Public Sub colour_blue_this_week_on_bom_pus_date_mrd_and_build(arkusz_raportu As Worksheet)
    Set r = arkusz_raportu.Range(XWIZ.ADRES_POCZATKU_DAT_DETAILS)
    
    
    If Trim(r.Offset(1, 0)) <> "" Then
        Set r = arkusz_raportu.Range(r, r.End(xlDown).Offset(0, 5))
    Else
        Set r = arkusz_raportu.Range(r, r.Offset(0, 5))
    End If
    
    r.Interior.Color = RGB(255, 255, 255)
    r.Font.Color = RGB(0, 0, 0)
    go_trough_this_range_and_colour_blue r
End Sub

Private Sub go_trough_this_range_and_colour_blue(rr As Range)
    
    Dim ir As Range
    For Each ir In rr
        
        
        If IsDate(ir) Then
            If isThisDateIsInPastWeek(ir) Then
                ir.Interior.Color = RGB(51, 102, 255)
                ir.Font.Color = RGB(255, 255, 255)
                
            ElseIf isThisDateIsToday(ir) Then
                ir.Interior.Color = RGB(200, 200, 255)
                ir.Font.Color = RGB(0, 0, 0)
            End If
        ElseIf isMyCustomYearAndCW(ir) Then
            
            If isStringIsPastMyYearAndCW(ir) Then
                ir.Interior.Color = RGB(51, 102, 255)
                ir.Font.Color = RGB(255, 255, 255)
            ElseIf ir = currentStringForMyYearAndCW() Then
                ir.Interior.Color = RGB(200, 200, 255)
                ir.Font.Color = RGB(0, 0, 0)
            End If
        End If
    Next ir
End Sub

Private Function isThisDateIsInPastWeek(d) As Boolean
    isThisDateIsInPastWeek = False
    
    If CDate(d) < Date Then
        isThisDateIsInPastWeek = True
    End If
End Function

Private Function isThisDateIsInCurrentWeek(d) As Boolean
    isThisDateIsInCurrentWeek = False
    
    If Application.WorksheetFunction.IsoWeekNum(CDate(ir)) = Application.WorksheetFunction.IsoWeekNum(Date) Then
        isThisDateIsInCurrentWeek = True
    End If
End Function

Private Function isThisDateIsToday(d) As Boolean
    isThisDateIsToday = False
    
    If CDate(Format(Date, "yyyy-mm-dd")) = CDate(Format(d, "yyyy-mm-dd")) Then
        isThisDateIsToday = True
    End If
End Function


Private Function isMyCustomYearAndCW(ycw) As Boolean
    isMyCustomYearAndCW = False
    
    ' 1st check
    If Left(CStr(ycw), 1) = "Y" Then
        ' 2nd check
        If Len(ycw) = 9 Or Len(ycw) = 8 Then
            
            '3rd check
            If IsNumeric(Mid(ycw, 2, 3)) Then
                
                ' 4th check
                If InStr(1, ycw, "CW") = 6 Then
                    
                    ' 4 czeki - raczej super dziwnie bedzie jesli wyjdzie jakis babol
                    isMyCustomYearAndCW = True
                End If
            End If
        End If
    End If
End Function

Private Function currentStringForMyYearAndCW() As String
    
    ' tmp = "YxxxxCWqq"
    
    y = Year(Date)
    cw = Application.WorksheetFunction.IsoWeekNum(Date)
    
    If Len(cw) = 1 Then
        tmp = "Y" & CStr(y) & "CW0" & CStr(cw)
    Else
        tmp = "Y" & CStr(y) & "CW" & CStr(cw)
    End If
    
    currentStringForMyYearAndCW = CStr(tmp)
End Function

Private Function isStringIsPastMyYearAndCW(ycw) As Boolean

    isStringIsPastMyYearAndCW = False

    y = Year(Date)
    cw = Application.WorksheetFunction.IsoWeekNum(Date)
    zero = ""
    If Len(cw) = 1 Then
        zero = "0"
    End If
    Dim curr_long As Long, ycw_long As Long
    curr_long = CLng(CStr(y) & zero & CStr(cw))
    
    
    t = Replace(ycw, "Y", "")
    arr = Split(t, "CW")
    
    If Len(arr(UBound(arr))) = 1 Then
        arr(UBound(arr)) = "0" & arr(UBound(arr))
    End If
    
    ycw_long = CLng(arr(LBound(arr)) & arr(UBound(arr)))
    
    
    ' nie ma problemu ze matematycznie sprawdzamy roznice nawet jesli brakuje danych
    ' miedzy xx54, a xx 99 bo i tak jest poprawny order i flow
    ' wiec niech tak zostanie - nigdy to sie nie zepsuje
    ' raczej zepsuje dopiero kolo 9999 roku - wtedy zaczne sie martwic
    If CLng(ycw_long) < CLng(curr_long) Then
        isStringIsPastMyYearAndCW = True
    End If
End Function


Private Sub Class_Initialize()
    
    Set rep = ThisWorkbook.Sheets(XWIZ.REP_SHEET_NAME)
    Set rep_fup = ThisWorkbook.Sheets(XWIZ.REP_FUP_SHEET_NAME)
End Sub

Private Sub Class_Terminate()
    Set r = Nothing
    Set rep = Nothing
    Set rep_fup = Nothing
End Sub
