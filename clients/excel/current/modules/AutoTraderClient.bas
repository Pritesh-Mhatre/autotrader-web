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
            If (OrderRow = 2) Then
                MsgBox "Please enter orders in <orders> sheet.", vbCritical, "No orders found"
                Exit Sub
            End If
            
            Exit For
        End If
        
        For AccountRow = 2 To rAccounts.Rows.Count
            temp = rAccounts.Cells(RowIndex:=AccountRow, columnIndex:="A").Value
            
            If (Trim(temp) = "") Then
                If (AccountRow = 2) Then
                    MsgBox "Please add accounts in <accounts> sheet.", vbCritical, "Accounts missing"
                    Exit Sub
                End If
                
                Exit For
            End If
            
            Application.Run "PlaceOrder", wOrders.Cells(OrderRow, 1), Trim(temp), Trim(wOrders.Cells(OrderRow, 2)), _
                Trim(wOrders.Cells(OrderRow, 3)), wOrders.Cells(OrderRow, 4), wOrders.Cells(OrderRow, 5), _
                wOrders.Cells(OrderRow, 6), wOrders.Cells(OrderRow, 7), wOrders.Cells(OrderRow, 8), _
                wOrders.Cells(OrderRow, 9), wOrders.Cells(OrderRow, 13), "DAY", _
                0, "", wOrders.Cells(OrderRow, 10), wOrders.Cells(OrderRow, 11), _
                wOrders.Cells(OrderRow, 12), "", -1
                
        Next
        
    Next

    MsgBox "Orders sent to AutoTrader Desktop Client", vbInformation, "Success"

End Sub

Sub PlaceOrdersManual()

    If MsgBox("Are you sure? Make sure AutoTrader Desktop Client is monitoring.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    PlaceOrders

End Sub


Sub PlaceOrdersOnTime()

    Dim RemainingTime As Double, Deadline As Double

    Deadline = Worksheets("timer").Range("B1")
    RemainingTime = Deadline - (Now - Date)
    
    If RemainingTime > (-1 / 86400) Then
        Worksheets("timer").Range("B2").Value = Format(RemainingTime, "h:mm:ss")
        Application.OnTime Now + 1 / 86400, "PlaceOrdersOnTime"
    Else
        PlaceOrders
    End If
    
End Sub

Sub StartTimer()

    Dim RemainingTime As Double, Deadline As Double

    Deadline = Worksheets("timer").Range("B1")
    RemainingTime = Deadline - (Now - Date)
    
    If RemainingTime <= (-1 / 86400) Then
        MsgBox "Time has already expired, please correct time."
        Exit Sub
    End If
    
    PlaceOrdersOnTime
    
End Sub

