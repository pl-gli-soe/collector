VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordsHandler"
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



'Global Const XWIZ_TXT_CMNT_PN_PREFIX = "PN: "
'Global Const XWIZ_TXT_CMNT_PN_NM_PREFIX = "PN NM: "
'Global Const XWIZ_TXT_CMNT_DUNS_PREFIX = "DUNS: "
'Global Const XWIZ_TXT_CMNT_SUPP_NM_PREFIX = "SUPP NM: "
'Global Const XWIZ_TXT_CMNT_RESP_PREFIX = "Resp: "
'Global Const XWIZ_TXT_CMNT_FUP_PREFIX = "FMA FUP: "



Public pn As String
Public pn_nm As String
Public duns As String
Public supp_nm As String
Public resp As String
Public fup As String
Public nok_counter As String

Public MRD1_Ordered_STATUS As Boolean
Public MRD1_confirmed_qty_Status As Boolean
Public MRD1_PUS_STATUS As Boolean
Public MRD2_Ordered_STATUS As Boolean
Public MRD2_confirmed_qty_Status As Boolean
Public MRD2_PUS_STATUS As Boolean
Public TOTAL_PUS_STATUS As Boolean

Public Delivery_confirmation As String
Public Delivery_confirmation_status As Boolean
Public MRD1_Ordered_Date As String
Public Comments As String

Public uniq_id As Long



Private Sub Class_Initialize()
    With Me
        .MRD1_confirmed_qty_Status = True
        .MRD1_Ordered_STATUS = True
        .MRD1_PUS_STATUS = True
        
        .MRD2_confirmed_qty_Status = True
        .MRD2_Ordered_STATUS = True
        .MRD2_PUS_STATUS = True
        
        .TOTAL_PUS_STATUS = True
        
        .Delivery_confirmation_status = True
        
        .Delivery_confirmation = ""
        .MRD1_Ordered_Date = ""
        .Comments = ""
    End With
End Sub

