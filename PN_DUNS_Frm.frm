VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PN_DUNS_Frm 
   Caption         =   "Search"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   OleObjectBlob   =   "PN_DUNS_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PN_DUNS_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Sub BtnSubmit_Click()
    hide
    
    If Me.PNOptionButton.Value = True Then
        
        ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("pns").Value = Me.TextBox1.Value
    ElseIf Me.DUNSOptionButton.Value = True Then
    
        ThisWorkbook.Sheets(XWiz.CONFIG_SHEET_NAME).Range("DUNSes").Value = Me.TextBox1.Value
    Else
        MsgBox "option button na pn duns frm - nigdy nie powinien sie pojawic ten msgbox"
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    End
End Sub
