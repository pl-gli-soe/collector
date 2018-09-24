VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DoubleClickReportHandler"
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



Private sh As Worksheet

Private uid As String

Public Sub set_uid(muid As String)
    uid = CStr(muid)
End Sub

Public Function get_uid() As String
    get_uid = CStr(uid)
End Function

Public Function get_sh() As Worksheet
    Set get_sh = sh
End Function

Private Sub Class_Initialize()
    Set sh = ThisWorkbook.Sheets.Add
End Sub


Public Function remove_all_prefixes(txt) As String
    remove_all_prefixes = ""
    
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_HASH, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_ROW, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_PN_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_PN_NM_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_DUNS_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_SUPP_NM_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_RESP_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_FUP_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_DEL_CONF_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX, "")
    txt = Replace(txt, XWIZ.XWIZ_TXT_CMNT_CMNTS_PREFIX, "")
    
    ' txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_LINIA, "")
    
    
    remove_all_prefixes = txt
End Function


Public Sub create_array_and_put_in_it_data_from_(txt)

    txt = Replace(txt, Chr(10), " | ")
    
    arr = Split(txt, " | ")
    
    prepare_labels
    
    Dim r As Range
    Set r = sh.Range("A3")
    
    i = 1
    For x = LBound(arr) To UBound(arr)
    
    
        r.Offset(0, i - 1).Value = arr(x)
        
        i = i + 1
        
        If i = XWIZ.E_CMNT_LINIA Then
            i = 1
            x = x + 1
            Set r = r.Offset(1, 0)
        End If
        
    Next x
End Sub

Private Sub prepare_labels()
    
    Dim r As Range
    Set r = sh.Range("A2")
    
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_HASH, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_ROW, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_PN_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_PN_NM_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_DUNS_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_SUPP_NM_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_RESP_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_FUP_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_DEL_CONF_PREFIX, "")
    'txt = Replace(txt, XWiz.XWIZ_TXT_CMNT_CMNTS_PREFIX, "")
    
    r.Value = XWIZ_TXT_CMNT_HASH
    r.Offset(0, 1) = XWIZ_TXT_CMNT_ROW
    r.Offset(0, 2) = XWIZ_TXT_CMNT_PN_PREFIX
    r.Offset(0, 3) = XWIZ_TXT_CMNT_PN_NM_PREFIX
    r.Offset(0, 4) = XWIZ_TXT_CMNT_DUNS_PREFIX
    r.Offset(0, 5) = XWIZ_TXT_CMNT_SUPP_NM_PREFIX
    r.Offset(0, 6) = XWIZ_TXT_CMNT_RESP_PREFIX
    r.Offset(0, 7) = XWIZ_TXT_CMNT_FUP_PREFIX
    
    r.Offset(0, 8) = XWIZ_TXT_CMNT_DEL_CONF_PREFIX
    r.Offset(0, 9) = XWIZ.XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX
    r.Offset(0, 10) = XWIZ_TXT_CMNT_CMNTS_PREFIX
End Sub

'dcrh.nazwa_kolumny Target.Parent.Cells(2, Target.Column)
'dcrh.projekt Target.Parent.Cells(Target.Row, 2)
'dcrh.lokalizacja_pliku Target.Parent.Cells(Target.Row, 1).Comment.Text


Public Sub nazwa_kolumny(r)
    sh.Range("A1") = r
End Sub


Public Sub projekt(r)
    sh.Range("B1") = r
End Sub

Public Sub lokalizacja_pliku(r)
    sh.Range("C1") = r
End Sub
