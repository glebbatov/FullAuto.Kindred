Attribute VB_Name = "Data_ClearColumns"
' This script clear cells that filled (Data sheet)
' Creator: Gleb Batov

Sub Data_GetClearColumns()

Dim Answer, MyNote As String

    'Message
    MyNote = "Clear Colomns?"
    
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "")
    If Answer = vbYes Then
        'Code for Yes button Press
        '2,4 = D2;
        'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")
        ActiveSheet.Range(Cells(2, 5), Cells(999, 15)).Select
        Selection.ClearContents
    Else
        'Code for No button Press
    End If
    Range("E2").Select

End Sub


