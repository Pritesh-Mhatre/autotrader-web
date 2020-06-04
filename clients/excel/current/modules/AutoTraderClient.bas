Attribute VB_Name = "AutoTraderClient"
Option Explicit

Sub PlaceOrders()
    
    If Application.Run("isAutoTraderClientMonitoring") = False Then
        Exit Sub
    End If
    
    Dim wOrders As Worksheet: Set wOrders = Sheets("orders")
    Dim rAccounts As Range: Set rAccounts = Sheets("accounts").Range("A:A")
    Dim rOrders As Range: Set rOrders = wOrders.Range("A:M")
    Dim temp As String
    Dim OrderRow As Long, AccountRow As Long

        
    For OrderRow = 2 To rOrders.Rows.Count
        temp = rOrders.Cells(RowIndex:=OrderRow, columnIndex:="A").Value
        
        If (Trim(temp) = "") Then
            Exit For
        End If
        
        For AccountRow = 2 To rAccounts.Rows.Count
            temp = rAccounts.Cells(RowIndex:=AccountRow, columnIndex:="A").Value
            
            If (Trim(temp) = "") Then
                Exit For
            End If
            
            Application.Run "PlaceOrder", wOrders.Cells(OrderRow, 1), temp, wOrders.Cells(OrderRow, 2), _
                wOrders.Cells(OrderRow, 3), wOrders.Cells(OrderRow, 4), wOrders.Cells(OrderRow, 5), _
                wOrders.Cells(OrderRow, 6), wOrders.Cells(OrderRow, 7), wOrders.Cells(OrderRow, 8), _
                wOrders.Cells(OrderRow, 9), wOrders.Cells(OrderRow, 13), "DAY", _
                0, "", wOrders.Cells(OrderRow, 10), wOrders.Cells(OrderRow, 11), _
                wOrders.Cells(OrderRow, 12), "", -1
                
        Next
        
    Next

End Sub

Sub PlaceOrderManual()

    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    PlaceOrders

End Sub
