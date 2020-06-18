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
            
			Dim Variety As String: Variety = Trim(wOrders.Cells(OrderRow, 1))
			Dim PseudoAccount As String: PseudoAccount = Trim(temp)
			Dim Exchange As String: Exchange = Trim(wOrders.Cells(OrderRow, 2))
			Dim Symbol As String: Symbol = Trim(wOrders.Cells(OrderRow, 3))
			Dim TradeType As String: TradeType = Trim(wOrders.Cells(OrderRow, 4))
			Dim OrderType As String: OrderType = Trim(wOrders.Cells(OrderRow, 6))
			Dim ProductType As String: ProductType = Trim(wOrders.Cells(OrderRow, 5))
			Dim Quantity As Integer: Quantity = wOrders.Cells(OrderRow, 7)
			Dim Price As Double: Price = wOrders.Cells(OrderRow, 8)
			Dim TriggerPrice As Double: TriggerPrice = wOrders.Cells(OrderRow, 9)
			Dim Target As Double: Target = wOrders.Cells(OrderRow, 10)
			Dim Stoploss As Double: Stoploss = wOrders.Cells(OrderRow, 11)
			Dim TrailingStoploss As Double: TrailingStoploss = wOrders.Cells(OrderRow, 12)
			Dim DisclosedQuantity As Integer: DisclosedQuantity = 0
			Dim Validity As String: Validity = "DAY"
			Dim Amo As Boolean: Amo = wOrders.Cells(OrderRow, 13)
			Dim StrategyId As Integer: StrategyId = -1
			Dim Comments As String: Comments = ""
			
            Application.Run "PlaceOrderAdvanced", Variety, PseudoAccount, Exchange, _
				Symbol, TradeType, OrderType, ProductType, Quantity, Price, _
				TriggerPrice, Target, Stoploss, TrailingStoploss, DisclosedQuantity, _
				Validity, Amo, StrategyId, Comments
                
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

