Attribute VB_Name = "Module36"
Option Explicit

Sub ExecuteReturns()
Attribute ExecuteReturns.VB_ProcData.VB_Invoke_Func = "q\n14"
    
    'Sets the spreadsheet to stop updating every change
    Application.ScreenUpdating = False
    
    'Defines variables
    Dim ISP As String
    Dim MTYPE As String
    Dim CN As String
    Dim CID As String
    Dim SN As String
    Dim Staff As String
    Dim PartSN1 As String
    Dim PartSN2 As String
    Dim PartSN3 As String
    Dim AusPostCheck As Boolean
    Dim Check As Boolean
    Dim EquipFound As Boolean
    Dim LastRow As String
    Dim R As Variant
    Dim C As Variant
    
    'Test Timing
    'Dim startTime As Double
    'startTime = Timer
    'Test Timing

    'Selects the last row in column A
    LastRow = Range("A1048576").End(xlUp).Row()
    Cells(LastRow, "A").Select
    
    'Sets variables for selected row and column
    R = ActiveCell.Row
    C = ActiveCell.Column
    
    'Sets customer's info variables from column A
    CN = Cells(R - 2, C).Value
    CID = Cells(R - 1, C).Value
    SN = Cells(R, C).Value
    
    'Sets current staff's name
    Staff = Cells(3, 11).Value
    
    'Gets the first letter of the serial number
    PartSN1 = Left(SN, 1)
    'Gets the first two letters of the serial number
    PartSN2 = Left(SN, 2)
    'Gets the first three letters of the serial number
    PartSN3 = Left(SN, 3)
           
    'Finds the model of the modem/router and sets modem/router type
    If (PartSN3 = "984") Then
        MTYPE = "C1200"
        EquipFound = True
        
    ElseIf (PartSN2 = "00" Or PartSN2 = "CC" Or PartSN2 = "1C" Or PartSN2 = "50" _
    Or PartSN2 = "34" Or PartSN2 = "7C" Or PartSN2 = "74" _
    Or PartSN2 = "98" Or PartSN2 = "C4" Or PartSN2 = "B0" _
    Or PartSN2 = "D8" Or PartSN2 = "3C" Or PartSN2 = "60" _
    Or PartSN2 = "E8" Or PartSN2 = "E4" Or PartSN2 = "90" _
    Or PartSN2 = "28" Or PartSN2 = "40" Or PartSN2 = "C0" _
    Or PartSN2 = "68" Or PartSN2 = "00" Or PartSN2 = "10") Then
    
        MTYPE = "VR1600v"
        EquipFound = True
        
    ElseIf (PartSN2 = "32") Then
        MTYPE = "VX420"
        EquipFound = True
    ElseIf (PartSN2 = "Z2") Then
        MTYPE = "Dongle"
        EquipFound = True
    ElseIf (PartSN2 = "FU") Then
        MTYPE = "Dongle E5576"
        EquipFound = True
    ElseIf (PartSN2 = "39") Then
        MTYPE = "NL1902"
        EquipFound = True
    ElseIf (PartSN2 = "CP") Then
        MTYPE = "TG789"
        EquipFound = True
    ElseIf (PartSN2 = "89") Then
        MTYPE = "SIM"
        EquipFound = True
    ElseIf (PartSN2 = "BS") Then
        MTYPE = "CG2200"
        EquipFound = True
    ElseIf (PartSN2 = "19" Or PartSN2 = "LB") Then
        MTYPE = "NCD"
        EquipFound = True
    ElseIf (PartSN3 = "210" Or PartSN3 = "95B") Then
        MTYPE = "NTU"
        EquipFound = True
    ElseIf (PartSN2 = "12" Or PartSN2 = "78" Or PartSN2 = "A4" Or PartSN2 = "13" _
    Or PartSN2 = "14" Or PartSN2 = "A6" Or PartSN2 = "A5") Then
        MTYPE = "NTD"
        EquipFound = True
    ElseIf (PartSN2 = "33" Or PartSN2 = "28") Then
        MTYPE = "EPC3940L"
        EquipFound = True
    ElseIf (PartSN2 = "73") Then
        MTYPE = "H626T"
        EquipFound = True
    ElseIf (PartSN2 = "93") Then
        MTYPE = "M616T"
        EquipFound = True
    ElseIf (PartSN2 = "84" Or PartSN2 = "83" Or PartSN2 = "82") Then
        MTYPE = "M605T"
        EquipFound = True
    ElseIf (PartSN1 = "H" Or PartSN1 = "E" Or PartSN1 = "K" Or PartSN1 = "J") Then
        MTYPE = "Fritz!Box7490"
        EquipFound = True
    ElseIf (PartSN2 = "A1") Then
        MTYPE = "BoBLite"
        EquipFound = True
    ElseIf (PartSN3 = "160") Then
        MTYPE = "NF12"
        EquipFound = True
    ElseIf (PartSN3 = "000") Then
        MTYPE = "A220"
        EquipFound = True
    ElseIf (PartSN2 = "J3") Then
        MTYPE = "HG659"
        EquipFound = True
    ElseIf (PartSN2 = "E7") Then
        MTYPE = "HG658"
        EquipFound = True
    ElseIf (PartSN2 = "R6") Then
        MTYPE = "HG630"
        EquipFound = True
    ElseIf (PartSN2 = "21") Then
        MTYPE = "HG532d"
        EquipFound = True

    End If

    'Checks if CID has a .
    AusPostCheck = InStr(CID, ".")
    
    'Checks if everything is all set and ready to be executed
    If (EquipFound = True And Len(CID) < 11 And AusPostCheck = False) Then
    
        'Sets the specific variables into the correct column and row of cells
        Range(Cells(R - 2, C), Cells(R, C)).ClearContents
        Cells(R - 2, C).Value = SN
        Cells(R - 2, C + 1).Value = MTYPE
        Cells(R - 2, C + 2).Value = CID
        Cells(R - 2, C + 7).Value = CN
        Cells(R - 2, C + 10).Value = Staff
        
        'Gets the ISP variable from the ISP cell
        ISP = Cells(R - 2, C + 9).Value
    
    
    
        'Checks variables and opens the necessary webpages
        If (CN = "wic" Or MTYPE = "Dongle") Then
            
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=log_note.cgi&cust_id=" + CID)
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=wh_track.cgi&cust_id=" + CID + "&type=admin")
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=user_query.cgi&cust_id=" + CID + "&username=&domain=&name_f=&name_s=&name_b=&addr_1=&addr_2=&addr_c=&addr_s=&addr_pc=&phone_h=&phone_w=&phone_m=&phone_f=&email=&bill_id=&owner=&phone_s=&msn=&sim_num=&phone_did=&prod_inst=&nbn_avc=&phone_id=&ident=&part_sn=&cc_num=&status=all")
            Cells(R - 2, C + 5).Value = "Equipment Returned"
            
        ElseIf (CID = "" And MTYPE <> "VX420" And MTYPE <> "NL1902") Then
            
            Cells(R - 2, C).Copy
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url http://blade.tpg.com.au/cgi-bin/ias.cgi?scr=user_query.cgi")
            
            If Left(CN, 3) = "ASH" Or Left(CN, 3) = "ADD" Or Left(CN, 3) = "PPA" _
            Or Left(CN, 3) = "M7Z" Or Left(CN, 3) = "AQQ" Or Left(CN, 3) = "MSK" _
            Or Left(CN, 3) = "GET" Or Left(CN, 3) = "SQU" Or Left(CN, 3) = "20F" Then
                Cells(R - 2, C + 5).Value = "Returned RTS"
            
            End If
            
        ElseIf (ISP = "TPG" And MTYPE <> "VX420" And MTYPE <> "NL1902") Then
        
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=log_note.cgi&cust_id=" + CID)
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=wh_track.cgi&cust_id=" + CID + "&type=admin")
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://blade.tpg.com.au/cgi-bin/ias.cgi?scr=user_query.cgi&cust_id=" + CID + "&username=&domain=&name_f=&name_s=&name_b=&addr_1=&addr_2=&addr_c=&addr_s=&addr_pc=&phone_h=&phone_w=&phone_m=&phone_f=&email=&bill_id=&owner=&phone_s=&msn=&sim_num=&phone_did=&prod_inst=&nbn_avc=&phone_id=&ident=&part_sn=&cc_num=&status=all")
            
            If Left(CN, 3) = "ASH" Or Left(CN, 3) = "ADD" Or Left(CN, 3) = "PPA" _
            Or Left(CN, 3) = "M7Z" Or Left(CN, 3) = "AQQ" Or Left(CN, 3) = "MSK" _
            Or Left(CN, 3) = "GET" Or Left(CN, 3) = "SQU" Or Left(CN, 3) = "20F" Then
                Cells(R - 2, C + 5).Value = "Returned RTS"
            ElseIf Left(CN, 3) = "7KG" Then
                Cells(R - 2, C + 5).Value = "Original Returned for AR"
            ElseIf Left(CN, 3) = "W0M" Then
                Cells(R - 2, C + 5).Value = "Free Router Returned"
            End If
            
        Else
            Cells(R - 2, C + 2).Copy
            Shell ("C:\Program Files (x86)\Mozilla Firefox\firefox.exe -url https://onewh.it.tpgtelecom.com.au:8200/OneWh/wh/orders/order_query.html")
            
            If Left(CN, 3) = "ASH" Or Left(CN, 3) = "ADD" Or Left(CN, 3) = "PPA" _
            Or Left(CN, 3) = "AQQ" Or Left(CN, 3) = "MSK" Or Left(CN, 3) = "GET" _
            Or Left(CN, 3) = "SQU" Or Left(CN, 3) = "20F" Then
                Cells(R - 2, C + 5).Value = "Returned RTS"
            Else
                Cells(R - 2, C + 5).Value = "Equipment returned via " + CN
            End If
            
        End If
   
    'Test Timing
    'MsgBox "Total time was: " & (Timer - startTime)
    'Test Timing
    Else
        MsgBox ("ERROR, Double Check Inputs")
        End If
    'Sets the spreadsheet to start updating every change
    Application.ScreenUpdating = True
    
End Sub

