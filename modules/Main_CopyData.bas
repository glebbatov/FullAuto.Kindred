Attribute VB_Name = "Main_CopyData"
' Script pulls PO# cells from "Data" sheet to "Main" sheet,
' removes empty cell between orders and removes all duplicate orders
' Creator: Gleb Batov

Sub Main_GetCopyData()

    Dim Rng As Range
    Dim OutRng As Range
    Dim InputRng As Range
    Dim xTitle As String

    Sheets("Main").Select
    totalOrders = Range("H2").Value
    
    'If totalOrders > 0 Then
    'Answer = MsgBox("Replace current Data?", vbQuestion + vbYesNo, "")
    'If Answer = vbYes Then
    
    'copy "Data sheet" columns without blank cells
    'PO copy
    Sheets("Data").Select
    Range("E2").Select
    Application.CutCopyMode = False
        Range("E2").Select
        On Error Resume Next
            xTitle = Application.ActiveWindow.RangeSelection.Address
                Set InputRng = Range("E2:E999")
                    Set InputRng = Application.Intersect(InputRng, Application.ActiveSheet.UsedRange)
                        For Each Rng In InputRng
                            If Not Rng.Value = "" Then
                                If OutRng Is Nothing Then
                                    Set OutRng = Rng
                                Else
                                    Set OutRng = Application.Union(OutRng, Rng)
                                End If
                            End If
                        Next
                    If Not (OutRng Is Nothing) Then
                OutRng.Select
            End If
    Selection.Copy
    Sheets("Main").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    'delete all duplicate cells in range (A2:A999)
    Range("A2:A999").Cells.RemoveDuplicates Columns:=Array(1), Header:=xlNo
        
    'make all sheets cell selection looks nice again
    Sheets("Data").Select
    Application.CutCopyMode = False
    Range("E2").Select
    Sheets("Main").Select
    Range("A2").Select
End Sub


