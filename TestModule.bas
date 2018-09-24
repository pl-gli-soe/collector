Attribute VB_Name = "TestModule"
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


Public Sub test()
    Dim wh As WrkHandler
    Set wh = New WrkHandler
    
    With wh
        
        
        ' jest OK
        ' .przejrzyjListe_TEST
        ' CzekajForm.show vbModeless
        Application.DisplayStatusBar = True
        Application.StatusBar = "odnajduje pliki typu Wizard, ktore sa zgodne ze wzorcem..."
        .stworz_kolekcje XWIZ.XWIZ_PATH_FOR_SEARCHING
        ' CzekajForm.hide
        
        
        Dim sh As StatusHandler
        Set sh = New StatusHandler
        sh.init_statusbar .countCollection
        
        
        Application.StatusBar = "czyszcze arkusz raportujacy"
        .wyczysc_arkusz_rep
        
        Application.StatusBar = "uruchamiono glowna logike"
        
        sh.show
        .przejdz_po_kolei_przez_kolekcje_nazw_i_pobierz_dane sh
        sh.hide
    End With
    
    
    MsgBox "Gotowe! " & CStr(Now)
End Sub
