Attribute VB_Name = "EnumModule"
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


' zbior enumow przenosze do osobnego modulu

' wraz z niedziela!
' 2015-10-11
' jest duplikat na enumach
' wszystkie kolumny + kolumny ktore posiadaja formule
' wskazuja wartosciowo dokladnie to samo
' jednak celem dla formul bylo uszczuplenie listy wyboru, gdy skupiamy sie na formula
' handling

Public Enum E_DYNAMIC_CFG_FOR_DEL_CONF
    E_DYNAMIC_CFG_FOR_DEL_CONF_NOK = 1
    E_DYNAMIC_CFG_FOR_DEL_CONF_OK = 2
    E_DYNAMIC_CFG_FOR_DEL_CONF_CALC_WITH_MRD = 3
End Enum


Public Enum E_RUN_REP_TYPE
    RUN_STD
    RUN_DUNS
    RUN_PN
    NEW_RUN_ALL
End Enum

Public Enum E_COLLECTOR_SIDE_TABLE
    E_SIDE_TB_HASH = 1
    E_SIDE_TB_ROW
    E_SIDE_TB_PN
    E_SIDE_TB_PN_NM
    E_SIDE_TB_DUNS
    E_SIDE_TB_SUPP_NM
    E_SIDE_TB_RESP
    E_SIDE_TB_FUP_CODE
    E_SIDE_TB_MRD1_STATUS
    E_SIDE_TB_MRD1_CONF_QTY
    E_SIDE_TB_MRD1_PUS_STATUS
    E_SIDE_TB_MRD2_STATUS
    E_SIDE_TB_MRD2_CONF_QTY
    E_SIDE_TB_MRD2_PUS_STATUS
    E_SIDE_TB_TOTAL_PUS_STATUS
    E_SIDE_TB_Delivery_confirmation
    E_SIDE_TB_Delivery_Confirmation_STATUS
    E_SIDE_TB_MRD1_Ordered_Date
    E_SIDE_TB_Comments
End Enum


Public Enum E_REP_STATUS_COLUMNS
    E_REP_MRD1_ORDERED_STATUS = 17
    E_REP_MRD1_CONF_QTY
    E_REP_MRD1_PUS_STATUS
    E_MRD2_ORDERED_STATUS
    E_MRD2_CONF_QTY
    E_MRD2_PUS_STATUS
    E_TOTAL_PUS_STATUS
    E_TOTAL_DEL_CONF_STATUS
End Enum


Public Enum E_FUP_FILTER
    
    E_FUP_FILTER_NO
    E_FUP_FILTER_YES
    
End Enum


Public Enum E_CREAT_COLLECTION_TYPE
    E_DYNAMIC = 1
    E_STATIC = 2
End Enum

Public Enum E_CMNT_ORDER
    E_CMNT_HASH = 1
    E_CMNT_ROW
    E_CMNT_PN
    E_CMNT_PN_NM
    E_CMNT_DUNS
    E_CMNT_SUPP_NM
    E_CMNT_RESP
    E_CMNT_FUP
    E_CMNT_DEL_CONF
    E_CMNT_MRD1_ORDERED_DATE
    E_CMNT_CMNTS
    E_CMNT_LINIA
End Enum

Public Enum E_ADD_EDIT_PUSES
    E_ADD = 1
    E_EDIT = 2
End Enum

Public Enum E_COLUMNS_WITH_FORMULAS
    E_F_TOTAL = 16
    E_F_MRD1_ORDERED_STATUS = 20
    E_F_MRD1_CONFIRMED_STATUS = 22
    E_F_MRD1_PUS_STATUS = 23
    E_F_MRD2_ORDERED_STATUS = 26
    E_F_MRD2_CONFIRMED_STATUS = 28
    E_F_MRD2_PUS_STATUS = 29
    E_F_TOTAL_PUS = 33
    E_F_TOTAL_PUS_STATUS = 34
    E_F_Delivery_confirmation = 35
End Enum



Public Enum E_ADD_DATA
    E_DOPISZ
    E_NADPISZ
End Enum

Public Enum E_DATE_OR_CW
    E_DC_DATE
    E_DC_CW
End Enum

Public Enum E_DETAILS_WIZARD_ORDER
    PIERWSZY
    SRODEK
    OSTATNI
End Enum



Public Enum E_JAKI_FORM
    cfg = 1
    WIZARD_COMBOBOX
    WIZARD_DATEPICKER
    WIZARD_TOGGLE
    WIZARD_TXTBOX
End Enum


Public Enum E_ORDERS
    O_INDX = 1
    O_PN
    O_DUNS
    O_FUP_code
    O_Pick_up_date
    O_Delivery_Date
    O_Pick_up_Qty
    O_PUS_Number
End Enum


Public Enum E_NEW_PROJECT_ITEM
    plt = 1
    PROJECT
    BIW_GA ' BIW or GA
    my
    PHAZE
    bom
    PICKUP_DATE
    PPAP_GATE
    mrd
    BUILD_START
    BUILD_END
    koordynator
    E_ACTIVE
    CAPACITY_CHECK
    E_MRD_DATE
    E_MRD_REG_ROUTES
    E_PLATFORM
    E_TRANSPORTATION_ACCOUNT_NUMBER
    E_UNIQUE_ID
End Enum


Public Enum E_MASTER_MANDATORY_COLUMNS
    pn = 1
    Alternative_PN
    PN_Name
    GPDS_PN_Name
    duns
    Supplier_Name
    Country_code
    MGO_code
    Responsibility
    fup_code
    SQ
    PPAP_Status
    SQ_Comments
    MRD1_QTY
    MRD2_QTY
    Total_QTY
    ADD_to_T_slash_D
    MRD1_Ordered_Date
    MRD1_Ordered_QTY
    MRD1_Ordered_STATUS
    MRD1_confirmed_qty
    MRD1_confirmed_qty_dot__Status
    MRD1_Total_PUS_STATUS
    MRD2_Ordered_date
    MRD2_Ordered_QTY
    MRD2_Ordered_STATUS
    MRD2_confirmed_qty
    MRD2_confirmed_qty_dot__Status
    MRD2_Total_PUS_STATUS
    Delivery_confirmation
    First_Confirmed_PUS_Date
    Delivery_reconfirmation
    Total_PUS_QTY
    TOTAL_PUS_STATUS
    Comments
    Bottleneck
    Future_Osea
    DRE
    EDI_Received
    BLANK1 ' tu cos innego
    BLANK2 ' tutaj oncost confirmation
    BLANK3
    BLANK4
End Enum


