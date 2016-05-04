VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordHandlerAllApproach"
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

Public plt As String
Public proj As String
Public bg As String
Public my As String
Public faza As String
Public koordynator As String


Public pn As String
Public duns As String
Public resp As String
Public fup As String
Public del_conf As String

Public cmnts As String

Public supp_nm As String
Public total_qty As String
Public ordered_date As String
Public ppap_status As String
Public ppap_gate As String
Public ordered_qty As String
Public confirmed_qty As String
Public fst_pickup_date As String

Public pickup_date_from_details As String
Public country_code As String


Public mrd As String
Public mrdd As String
Public unique_id As String


Public Function convert_mrd_to_date(moja_data As Date) As Date

    If IsDate(mrdd) Then
        convert_mrd_to_date = CDate(mrdd)
    Else

        tmp_txt = Replace(Replace(mrd, "Y", ""), "CW", "")
        
        
        y = Left(tmp_txt, 4)
        cw = Right(tmp_txt, 2)
        
        
        If CLng(Year(moja_data)) = CLng(y) Then
            
            If CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(moja_data))) = CLng(cw) Then
                
                '
                ''
                ' nop
                ''
                '
                
            ElseIf CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(moja_data))) < CLng(cw) Then
                moja_data = convert_mrd_to_date(moja_data + 1)
                
            ElseIf CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(moja_data))) > CLng(cw) Then
                moja_data = convert_mrd_to_date(moja_data - 1)
            End If
            
            convert_mrd_to_date = moja_data
        ElseIf CLng(Year(moja_data)) < CLng(y) Then
        
            moja_data = convert_mrd_to_date(moja_data + 10)
        
        ElseIf CLng(Year(moja_data)) > CLng(y) Then
        
            moja_data = convert_mrd_to_date(moja_data - 10)
            
        End If
    End If
  
End Function
