Attribute VB_Name = "PushData_PushData"
' This script pushes the data to SAP
' Creator: Gleb Batov

Sub PushData_GetPushData()

'On Error GoTo Catch

    If Not IsObject(sapApplication) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set sapApplication = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = sapApplication.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject sapApplication, "on"
    End If
    
    Dim repeat As Long
    Dim counter, totalOrders, orderQuantity, nextInOrder, orderOffset, previousOrderQuantity As Integer
    Dim orderNumber As String
    
    totalOrders = Range("H2").Value
    
    currentQuantity = Range("H3").Value
    
    serialNumber = Range("B2").Value
    assetTag = Range("C2").Value
    UserName = Range("D2").Value
    
    counter = 0
    nextInOrderCounter = 0  'next item in the order
    orderOffset = 0
    
    'Hit F3(back) button for 5 times to make sure that SAP is on the right page
    For repeat = 1 To 5
    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
    Next repeat
    
    'go to zint to order page/sort columns
    Do While counter < totalOrders
        Range("A2").Select
        orderNumber = ActiveCell.Offset(counter, 0).Value
        session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[1]/btn[9]").press
        session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber
        session.findById("wnd[0]").sendVKey 0
        
        '
        'remove after done
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_SERNR/txtWA_SERNR-SERNR[1,0]").SetFocus 'chose first item in order
        'session.findById("wnd[0]").sendVKey 2   'F2 chose
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").reorderTable "0 3 5 6 7 1 2 4 8 9 10 11 12 13"
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(1).Width = 5
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(2).Width = 7
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(3).Width = 8
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(4).Width = 8
        ''session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(5).Width = 11
        'session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(6).Width = 8
        'return back and come back again for correct field change
        'session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
        'remove till here
        '
        
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_SERNR/txtWA_SERNR-SERNR[1,0]").SetFocus 'choose first item in order
        session.findById("wnd[0]").sendVKey 2   'F2 chose
                    
        Range("E2").Select
        orderQuantity = ActiveCell.Offset(orderQuantity, 0).Value
                    
            'fill all data for a specific item
            Do While nextInOrderCounter < orderQuantity
            
                Range("B2").Select
                serialNumber = ActiveCell.Offset(orderOffset, 0).Value
                Range("C2").Select
                assetTag = ActiveCell.Offset(orderOffset, 0).Value
                Range("D2").Select
                UserName = ActiveCell.Offset(orderOffset, 0).Value

                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-BOX[1,0]").Text = "1"
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-BOX[1,1]").Text = "1"
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-BOX[1,2]").Text = "1"
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-BOX[1,3]").Text = "1"
                On Error Resume Next
                
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-SERNR[5,0]").Text = serialNumber
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-SERNR[5,1]").Text = serialNumber
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-SERNR[5,2]").Text = serialNumber
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-SERNR[5,3]").Text = serialNumber
                On Error Resume Next
                
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-ASSETTAG[6,0]").Text = assetTag
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-ASSETTAG[6,1]").Text = assetTag
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-ASSETTAG[6,2]").Text = assetTag
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/ctxtWA_COMP_SERIAL-ASSETTAG[6,3]").Text = assetTag
                On Error Resume Next
                
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-ADDLDATA[2,0]").Text = UserName
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-ADDLDATA[2,1]").Text = UserName
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-ADDLDATA[2,2]").Text = UserName
                On Error Resume Next
                session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-ADDLDATA[2,3]").Text = UserName
                On Error Resume Next
                
                'session.findById("wnd[0]").sendVKey 9   'F9 Save
                session.findById("wnd[0]").sendVKey 8   'F8 Next
                nextInOrderCounter = nextInOrderCounter + 1
                orderOffset = orderOffset + 1
                Loop
                    counter = counter + orderQuantity
                    orderOffset = counter
                    'previousOrderQuantity = orderQuantity
                    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                    nextInOrderCounter = 0
                    Range("E2").Select
                    orderQuantity = ActiveCell.Offset(orderQuantity, 0).Value
                    
                    
        Loop
    Sheets("PushData").Select
    Range("A2").Select
    MsgBox ("Inserted!")
Exit Sub
    
Catch:
619

MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End Sub

