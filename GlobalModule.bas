Attribute VB_Name = "GlobalModule"
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


Global Const COMBOBOX_SOURCE_DYN_DEL_CONF_OK = "always OK"
Global Const COMBOBOX_SOURCE_DYN_DEL_CONF_NOK = "always NOK"
Global Const COMBOBOX_SOURCE_DYN_DEL_CONF_CALC_IT = "calc it"

Global Const MAX_SHEET_NAME_LEN = 28
Global Const G_KOLUMNA_PO_DELIVERY_CONFIRMATION_STATUS_W_ARKUSZU_REP = 25

Global Const REP_SHEET_NAME = "rep"
Global Const REP_FUP_SHEET_NAME = "rep_fup"
Global Const ALL_SHEET_NAME = "all"
Global Const PIVOT_SOURCE_SHEET_NAME = "pivotSource"
Global Const PIVOT_SHEET_NAME = "PIVOT"
Global Const PN_PIVOT_SHEET_NAME = "PN_PIVOT"
Global Const MASTER_SHEET_NAME = "MASTER"
Global Const CONFIG_SHEET_NAME = "config"
Global Const REGISTER_SHEET_NAME = "register"
Global Const DETAILS_SHEET_NAME = "DETAILS"
Global Const ORDERS_SHEET_NAME = "ORDERS"
Global Const PICKUPS_SHEET_NAME = "PICKUPS"
Global Const DCS_SHEET_NAME = "delivery_confirmation_special"
Global Const COMMA = ","

Global Const DCS = "delivery_confirmation_special"
' this is pointer for details sheet during working on init wizard on proj def
Global Const POINTER = "POINTER"

Global Const FMA_STR = "FMA"
Global Const FMA_WITH_STARS = "*FMA*"
Global Const ENTER_STR = "ENTER"

Global Const MGMT_CMNTS = "MGMT Cmnts"

Global Const ADRES_POCZATKU_NOKOW_W_REP = "Q3"
Global Const ADRES_POCZATKU_DAT_DETAILS = "F3"

Global Const TBD = "tbd"
Global Const GHOST = "GHOST"
Global Const DCS_STR = "delivery_confirmation_special"
Global Const STR_BLANK = "BLANK"

Global Const MRD_KLUCZ_DO_PODMIANY = "{MRD}"

Global Const SELECTION_LIMIT = 256
' 2^14
Global Const TOP_EDIT_LIMIT = 16384
' Global Const TOP_EDIT_LIMIT = 50
Global Const ASCII_0 = 48
Global Const ASCII_9 = 57
Global Const ASCII_ENTER = 13

Global Const ALL_ORDERED_QTY = "ALL Ordered Qty"

Global Const G_PASS = "1985-07-10"

Global Const G_HOW_MANY_ROWS_WILL_BE_DELETED = 524288 ' 2^19 polowa capacity akursza excela
Global Const POLOWA_CAPACITY_ARKUSZA = 524288 ' 2^19 polowa capacity akursza excela
Global Const CAPACITY_ARKUSZA = 1048576

Global Const DWA_DO_16 = 65536 ' 2^10 polowa capacity akursza excela

Global Const SIX = 6

Global Const G_STEP_BETWEEN_PARALELL_USERS = 50000

Global Const G_OK = "OK"
Global Const G_NOK = "NOK"

Global Const G_CMNT_WIDTH = 650
Global Const G_CMNT_HEIGHT = 40


' najwazniejsza zmienna stala globalna dla poczatku projekt XWiz
' =====================================================================
Global Const XWIZ_PATH_FOR_SEARCHING = "X:\PLGLI-Exchange\SoE\FMA\"
Global Const REPO_PATH = "C:\WORKSPACE\macros\Wizard\Collector\repo\"
Global Const G_TEST_NA_DYSKU_LOKALNYM As Boolean = False

Global Const XWIZ_FILE_PREFIX = "M"
Global Const XWIZ_FILE_MIDFIX = "wizard"
Global Const XWIZ_FLE_POSTFIX_VERSION = "3.9"


Global Const XWIZ_TXT_CMNT_LINIA = "-----------"
Global Const XWIZ_TXT_CMNT_PN_PREFIX = "PN: "
Global Const XWIZ_TXT_CMNT_PN_NM_PREFIX = "PN NM: "
Global Const XWIZ_TXT_CMNT_DUNS_PREFIX = "DUNS: "
Global Const XWIZ_TXT_CMNT_SUPP_NM_PREFIX = "SUPP NM: "
Global Const XWIZ_TXT_CMNT_RESP_PREFIX = "Resp: "
Global Const XWIZ_TXT_CMNT_FUP_PREFIX = "FMA FUP: "
Global Const XWIZ_TXT_CMNT_DEL_CONF_PREFIX = "DEL CONF: "
Global Const XWIZ_TXT_CMNT_MRD1_Ordered_Date_PREFIX = "MRD1 Ordered Date: "
Global Const XWIZ_TXT_CMNT_CMNTS_PREFIX = "Comments: "

Global Const XWIZ_TXT_CMNT_HASH = "# "
Global Const XWIZ_TXT_CMNT_ROW = "row: "


Global Const ILE_PODZIALOW_W_LECIMY_TUTAJ = 9
Global Const DODATKOWE_POLA_OD_DETAILS = 5
Global Const POD_MINI_PROGRES_DLA_REP_ALL = 5

Global Const OSTATNIA_KOLUMNA_DLA_PIVOT_SOURCE = 11


Global Const XWIZ_FLAGA = "wgkiweb2o9238hf32oufn3292n2n9fh2fg293fh2923fh324ghgeoiguhasoghd"



' dlugosci stringow w komentarzach
Global Const G_ROW_LEN = 4
Global Const G_PN_LEN = 9
Global Const G_PN_NM_LEN = 10
Global Const G_DUNS_LEN = 10
Global Const G_SUPP_NM_LEN = 15
Global Const G_RESP_LEN = 10
Global Const G_FUPCODE_LEN = 2
Global Const G_DATES_CW_LEN = 12
Global Const G_DEL_CONF_LEN = 20

Global Const G_CUT_PROJECT = 9
Global Const G_CUT_PHAZE = 6




' =====================================================================

Public Function fnDateFromWeek(iYear As Integer, iWeek As Integer, iWeekDday As Integer)
    ' get the date from a certain day in a certain week in a certain year
      fnDateFromWeek = _
      DateSerial(iYear, 1, ((iWeek - 1) * 7) + iWeekDday - Weekday(DateSerial(iYear, 1, 1)) + 1)
End Function


' global sub
Public Sub nowy_schemat_offsetu_w_arkuszu_pickups(ByRef i As Range)


    Set i = i.Offset(1, 0)
    If Trim(i) = "" Then
        Set i = i.End(xlDown)
    End If
End Sub



Public Sub Unhide_All_Rows(ByRef Sh As Worksheet)
    On Error Resume Next
     'in case the sheet is protected
    Sh.Cells.EntireRow.Hidden = False
End Sub
 
Public Sub Unhide_All_Columns(ByRef Sh As Worksheet)
    On Error Resume Next
     'in case the sheet is protected
    Sh.Cells.EntireColumn.Hidden = False
End Sub

Public Function dopelnij_spacjami(ms As String, ile_znakow As Integer) As String
    dopelnij_spacjami = ""
    
    If Len(ms) < ile_znakow Then
        
        
        ile_spacji = Int(ile_znakow - Len(ms))
        
        Select Case ile_spacji
            Case 1
                spacje = " "
            Case 2
                spacje = "  "
            Case 3
                spacje = "   "
            Case 4
                spacje = "    "
            Case 5
                spacje = "     "
            Case 6
                spacje = "      "
            Case 7
                spacje = "       "
            Case 8
                spacje = "        "
            Case 9
                spacje = "         "
            Case 10
                spacje = "          "
            Case 11
                spacje = "           "
            Case 12
                spacje = "            "
            Case 13
                spacje = "             "
            Case Else
                spacje = Application.WorksheetFunction.Rept(" ", CDbl(ile_spacji))
        End Select
        
        dopelnij_spacjami = CStr(ms) & spacje
    ElseIf Len(ms) > ile_znakow Then
        dopelnij_spacjami = Left(ms, ile_znakow)
    ElseIf Len(ms) = ile_znakow Then
        dopelnij_spacjami = ms
    Else
        MsgBox "ten msgbox nie moze sie pojawic w dopelnij_spacjami"
    End If
End Function

Public Sub unhide_all_rows_and_all_columns(ish As Worksheet)

    On Error Resume Next
    ish.ShowAllData

    On Error Resume Next
    ish.Cells.EntireRow.Hidden = False
    
    On Error Resume Next
    ish.Cells.EntireColumn.Hidden = False
End Sub


Public Function remove_special_cases(nm)

    nm = Replace(nm, ".xlsm", "")
    nm = Replace(nm, "/", "")
    nm = Replace(nm, "\", "")
    nm = Replace(nm, ",", "")
    nm = Replace(nm, ";", "")
    nm = Replace(nm, "&", "")
    nm = Replace(nm, "*", "")
    nm = Replace(nm, "%", "")
    nm = Replace(nm, "#", "")
    nm = Replace(nm, "@", "")
    nm = Replace(nm, "!", "")
    nm = Replace(nm, "+", "")
    nm = Replace(nm, "=", "")
    nm = Replace(nm, "-", "")
    nm = Replace(nm, "_", "")
    nm = Replace(nm, " ", "")
    nm = Replace(nm, "M_", "")
    
    
    remove_special_cases = nm
End Function
